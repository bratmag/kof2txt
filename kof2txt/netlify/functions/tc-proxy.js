// netlify/functions/tc-proxy.js

const ALLOWED_HOSTS = [
  "app.connect.trimble.com",
  "app21.connect.trimble.com",
  "app31.connect.trimble.com",
  "app32.connect.trimble.com",
  "app.eu.connect.trimble.com",
  "app.asia.connect.trimble.com",
  "amazonaws.com"
];

function isAllowedUrl(url) {
  try {
    const host = new URL(url).hostname;
    return ALLOWED_HOSTS.some((h) => host === h || host.endsWith("." + h));
  } catch {
    return false;
  }
}

function json(statusCode, body) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": "*"
    },
    body: JSON.stringify(body, null, 2)
  };
}

function getCoreBaseUrl(projectLocation) {
  const loc = String(projectLocation || "").toLowerCase();
  if (loc === "europe") return "https://app21.connect.trimble.com/tc/api/2.0";
  if (loc === "asia")   return "https://app31.connect.trimble.com/tc/api/2.0";
  return "https://app.connect.trimble.com/tc/api/2.0";
}

async function fetchAny(url, options = {}) {
  const response = await fetch(url, options);
  const contentType = response.headers.get("content-type") || "";
  const text = await response.text();
  let parsed = null;
  try { parsed = JSON.parse(text); } catch { /* ikke JSON */ }
  return { ok: response.ok, status: response.status, contentType, text, json: parsed };
}

// ─── downloadKofFile ─────────────────────────────────────────────────────────
async function handleDownloadKofFile(token, fileId, projectLocation) {
  const base = getCoreBaseUrl(projectLocation);
  const auth = { Authorization: `Bearer ${token}` };
  const diagnostics = { fileId, projectLocation, base, steps: [] };

  // Steg 1: Hent metadata
  let versionId = fileId;
  try {
    const metaUrl = `${base}/files/${encodeURIComponent(fileId)}`;
    const meta = await fetchAny(metaUrl, { method: "GET", headers: { ...auth, Accept: "application/json" } });
    diagnostics.steps.push({ step: "metadata", url: metaUrl, status: meta.status, ok: meta.ok, preview: meta.text.slice(0, 400) });
    if (meta.ok && meta.json?.versionId) versionId = meta.json.versionId;
  } catch (err) {
    diagnostics.steps.push({ step: "metadata", error: String(err) });
  }

  diagnostics.versionId = versionId;

  // Steg 2: Prøv ulike nedlastingsstrategier
  const candidates = [
    // Strategi A: application/octet-stream direkte på fil-endepunktet
    {
      url: `${base}/files/${encodeURIComponent(versionId)}`,
      headers: { ...auth, Accept: "application/octet-stream" }
    },
    // Strategi B: ?download=true med octet-stream
    {
      url: `${base}/files/${encodeURIComponent(versionId)}?download=true`,
      headers: { ...auth, Accept: "application/octet-stream" }
    },
    // Strategi C: /content med versionId
    {
      url: `${base}/files/${encodeURIComponent(versionId)}/content`,
      headers: { ...auth, Accept: "application/octet-stream" }
    },
    // Strategi D: blobstore (gir pre-signed URL)
    {
      url: `${base}/files/${encodeURIComponent(versionId)}/blobstore`,
      headers: { ...auth, Accept: "application/json" }
    },
    // Strategi E: fileId (ikke versionId) med octet-stream
    {
      url: `${base}/files/${encodeURIComponent(fileId)}`,
      headers: { ...auth, Accept: "application/octet-stream" }
    }
  ];

  for (const candidate of candidates) {
    try {
      const res = await fetchAny(candidate.url, { method: "GET", headers: candidate.headers });
      const isJson = res.contentType.includes("application/json");

      diagnostics.steps.push({
        step: "download-candidate",
        url: candidate.url,
        acceptHeader: candidate.headers.Accept,
        status: res.status,
        ok: res.ok,
        contentType: res.contentType,
        preview: res.text.slice(0, 300)
      });

      if (!res.ok) continue;

      // Scenario A: JSON med pre-signed URL
      const presignedUrl = res.json?.url || res.json?.downloadUrl || res.json?.href || null;
      if (isJson && presignedUrl) {
        diagnostics.steps.push({ step: "presigned-found", host: new URL(presignedUrl).hostname });
        try {
          // Hent fra pre-signed URL UTEN auth-header
          const fileRes = await fetchAny(presignedUrl, { method: "GET" });
          diagnostics.steps.push({ step: "presigned-fetch", status: fileRes.status, ok: fileRes.ok, bytes: fileRes.text.length });
          if (fileRes.ok) {
            return json(200, { ok: true, fileId, versionId, source: candidate.url, via: "presigned", content: fileRes.text, diagnostics });
          }
        } catch (err) {
          diagnostics.steps.push({ step: "presigned-fetch", error: String(err) });
        }
      }

      // Scenario B: Direkte filinnhold (ikke tom JSON-metadata)
      if (!isJson && res.text.length > 10) {
        return json(200, { ok: true, fileId, versionId, source: candidate.url, via: "direct", content: res.text, diagnostics });
      }

      // Scenario C: JSON, men ikke metadata (f.eks. KOF innpakket i JSON?)
      if (isJson && res.json && !res.json.id && !res.json.type) {
        return json(200, { ok: true, fileId, versionId, source: candidate.url, via: "json-wrapped", content: res.text, diagnostics });
      }
    } catch (err) {
      diagnostics.steps.push({ step: "download-candidate", url: candidate.url, error: String(err) });
    }
  }

  return json(502, {
    ok: false,
    error: "Ingen nedlastingsstrategier fungerte.",
    fileId,
    versionId,
    diagnostics
  });
}

// ─── Handler ─────────────────────────────────────────────────────────────────
exports.handler = async function (event) {
  if (event.httpMethod === "OPTIONS") {
    return {
      statusCode: 204,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type"
      },
      body: ""
    };
  }

  if (event.httpMethod !== "POST") return { statusCode: 405, body: "Use POST" };

  let data;
  try {
    data = JSON.parse(event.body || "{}");
  } catch {
    return json(400, { ok: false, error: "Invalid JSON body" });
  }

  const { action, token, url, method, fileId, projectLocation } = data;

  if (!token) return json(400, { ok: false, error: "Missing token" });

  // Action: downloadKofFile
  if (action === "downloadKofFile") {
    if (!fileId) return json(400, { ok: false, error: "Missing fileId" });
    return handleDownloadKofFile(token, fileId, projectLocation);
  }

  // Action: proxy (enkel gjennomgang, bakoverkompatibel)
  if (!url) return json(400, { ok: false, error: "Missing url" });
  if (!isAllowedUrl(url)) return json(403, { ok: false, error: `URL ikke tillatt: ${url}` });

  try {
    const response = await fetch(url, {
      method: method || "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json"
      }
    });
    const text = await response.text();
    return {
      statusCode: response.status,
      headers: {
        "Content-Type": response.headers.get("content-type") || "text/plain",
        "Access-Control-Allow-Origin": "*"
      },
      body: text
    };
  } catch (err) {
    return json(500, { ok: false, error: "Proxy error: " + String(err) });
  }
};
