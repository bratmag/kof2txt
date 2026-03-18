// netlify/functions/tc-proxy.js
//
// Støtter tre actions:
//   1. { action: "proxy", url, token, method }
//      → Enkel gjennomgangsproxy med Bearer token
//
//   2. { action: "downloadKofFile", fileId, token, projectLocation }
//      → Fullstendig nedlastingsflyt:
//         a) Hent metadata → versionId
//         b) Prøv blobstore/download/content → pre-signed URL
//         c) Hent filinnhold fra pre-signed URL (uten Auth-header)
//         d) Returner { ok, content, diagnostics }
//
//   3. (ingen action / gammel kode) { url, token, method }
//      → Bakoverkompatibel enkel proxy

const TRIMBLE_HOSTS = [
  "app.connect.trimble.com",
  "app.eu.connect.trimble.com",
  "app.asia.connect.trimble.com",
  "api.connect.trimble.com"
];

function isTrimbleUrl(url) {
  try {
    const host = new URL(url).hostname;
    return TRIMBLE_HOSTS.some((h) => host === h || host.endsWith("." + h));
  } catch {
    return false;
  }
}

function isAllowedUrl(url) {
  // Tillat Trimble-domener og AWS S3 pre-signed URLer
  try {
    const host = new URL(url).hostname;
    return (
      isTrimbleUrl(url) ||
      host.endsWith(".amazonaws.com") ||
      host.endsWith(".s3.amazonaws.com")
    );
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
  if (loc === "europe" || loc.includes("eu")) {
    return "https://app.eu.connect.trimble.com/tc/api/2.0";
  }
  if (loc === "asia") {
    return "https://app.asia.connect.trimble.com/tc/api/2.0";
  }
  return "https://app.connect.trimble.com/tc/api/2.0";
}

async function fetchJsonOrText(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  let parsed = null;
  try { parsed = JSON.parse(text); } catch { /* ikke JSON */ }

  return {
    ok: response.ok,
    status: response.status,
    contentType: response.headers.get("content-type") || "",
    text,
    json: parsed
  };
}

// ─── Action: downloadKofFile ──────────────────────────────────────────────────
async function handleDownloadKofFile(token, fileId, projectLocation) {
  const base = getCoreBaseUrl(projectLocation);
  const authHeaders = { Authorization: `Bearer ${token}` };
  const diagnostics = { fileId, projectLocation, base, steps: [] };

  // Steg 1: Hent metadata for å få versionId
  let versionId = fileId;
  try {
    const metaUrl = `${base}/files/${encodeURIComponent(fileId)}`;
    const meta = await fetchJsonOrText(metaUrl, {
      method: "GET",
      headers: authHeaders
    });

    diagnostics.steps.push({
      step: "metadata",
      url: metaUrl,
      status: meta.status,
      ok: meta.ok,
      preview: meta.text.slice(0, 400)
    });

    if (meta.ok && meta.json?.versionId) {
      versionId = meta.json.versionId;
    }
  } catch (err) {
    diagnostics.steps.push({ step: "metadata", error: String(err) });
  }

  diagnostics.versionId = versionId;

  // Steg 2: Prøv ulike endepunkter for å hente pre-signed URL eller direkte innhold
  const candidates = [
    `${base}/files/${encodeURIComponent(versionId)}/blobstore`,
    `${base}/files/${encodeURIComponent(versionId)}/download`,
    `${base}/files/${encodeURIComponent(versionId)}/blob`,
    `${base}/files/${encodeURIComponent(versionId)}/content`
  ];

  for (const url of candidates) {
    try {
      const res = await fetchJsonOrText(url, {
        method: "GET",
        headers: authHeaders
      });

      const isJson = res.contentType.includes("application/json");

      diagnostics.steps.push({
        step: "download-candidate",
        url,
        status: res.status,
        ok: res.ok,
        contentType: res.contentType,
        preview: res.text.slice(0, 300)
      });

      if (!res.ok) continue;

      // Scenario A: JSON-respons med pre-signed URL
      const presignedUrl =
        res.json?.url ||
        res.json?.downloadUrl ||
        res.json?.href ||
        res.json?.link ||
        null;

      if (isJson && presignedUrl) {
        diagnostics.steps.push({
          step: "presigned-url-found",
          host: (() => { try { return new URL(presignedUrl).host; } catch { return "?"; } })()
        });

        // Steg 3: Hent innhold fra pre-signed URL (UTEN Authorization-header)
        try {
          const fileRes = await fetchJsonOrText(presignedUrl, {
            method: "GET"
          });

          diagnostics.steps.push({
            step: "presigned-fetch",
            status: fileRes.status,
            ok: fileRes.ok,
            contentType: fileRes.contentType,
            byteLength: fileRes.text.length
          });

          if (fileRes.ok) {
            return json(200, {
              ok: true,
              fileId,
              versionId,
              source: url,
              content: fileRes.text,
              diagnostics
            });
          }
        } catch (err) {
          diagnostics.steps.push({ step: "presigned-fetch", error: String(err) });
        }
      }

      // Scenario B: Direkte filinnhold (ikke JSON, ikke tom)
      if (!isJson && res.text.length > 10) {
        return json(200, {
          ok: true,
          fileId,
          versionId,
          source: url,
          content: res.text,
          diagnostics
        });
      }
    } catch (err) {
      diagnostics.steps.push({
        step: "download-candidate",
        url,
        error: String(err)
      });
    }
  }

  return json(502, {
    ok: false,
    error: "Ingen nedlastingsendepunkt fungerte. Se diagnostics for detaljer.",
    fileId,
    versionId,
    diagnostics
  });
}

// ─── Handler ──────────────────────────────────────────────────────────────────
exports.handler = async function (event) {
  // CORS preflight
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

  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Use POST" };
  }

  let data;
  try {
    data = JSON.parse(event.body || "{}");
  } catch {
    return json(400, { ok: false, error: "Invalid JSON body" });
  }

  const { action, token, url, method, fileId, projectLocation } = data;

  if (!token) {
    return json(400, { ok: false, error: "Missing token" });
  }

  // ── Action: downloadKofFile ──────────────────────────────────────────────
  if (action === "downloadKofFile") {
    if (!fileId) return json(400, { ok: false, error: "Missing fileId" });
    return handleDownloadKofFile(token, fileId, projectLocation);
  }

  // ── Action: proxy (eller ingen action = bakoverkompatibel) ───────────────
  if (!url) {
    return json(400, { ok: false, error: "Missing url" });
  }

  if (!isAllowedUrl(url)) {
    return json(403, { ok: false, error: `URL ikke tillatt: ${url}` });
  }

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
