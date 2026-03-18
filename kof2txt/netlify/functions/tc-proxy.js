// netlify/functions/tc-proxy.js

function isAllowedHost(url) {
  try {
    const host = new URL(url).hostname;
    return (
      host.endsWith(".connect.trimble.com") ||
      host.endsWith(".amazonaws.com")
    );
  } catch {
    return false;
  }
}

function json(statusCode, body) {
  return {
    statusCode,
    headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" },
    body: JSON.stringify(body, null, 2)
  };
}

function getCoreBaseUrl(loc) {
  const l = String(loc || "").toLowerCase();
  if (l === "europe") return "https://app21.connect.trimble.com/tc/api/2.0";
  if (l === "asia")   return "https://app31.connect.trimble.com/tc/api/2.0";
  return "https://app.connect.trimble.com/tc/api/2.0";
}

function getWopiBaseUrl(loc) {
  const l = String(loc || "").toLowerCase();
  if (l === "europe") return "https://wopi-api-eu.connect.trimble.com/v1";
  if (l === "asia")   return "https://wopi-api-ap.connect.trimble.com/v1";
  return "https://wopi-api.connect.trimble.com/v1";
}

async function fetchAny(url, options = {}) {
  const response = await fetch(url, options);
  const contentType = response.headers.get("content-type") || "";
  const text = await response.text();
  let parsed = null;
  try { parsed = JSON.parse(text); } catch {}
  return { ok: response.ok, status: response.status, contentType, text, json: parsed };
}

async function handleDownloadKofFile(token, fileId, projectLocation) {
  const coreBase = getCoreBaseUrl(projectLocation);
  const wopiBase = getWopiBaseUrl(projectLocation);
  const auth = { Authorization: `Bearer ${token}` };
  const diagnostics = { fileId, projectLocation, coreBase, wopiBase, steps: [] };

  // Steg 1: Metadata for å hente versionId
  let versionId = fileId;
  try {
    const metaUrl = `${coreBase}/files/${encodeURIComponent(fileId)}`;
    const meta = await fetchAny(metaUrl, { method: "GET", headers: { ...auth, Accept: "application/json" } });
    diagnostics.steps.push({ step: "metadata", url: metaUrl, status: meta.status, ok: meta.ok, preview: meta.text.slice(0, 300) });
    if (meta.ok && meta.json?.versionId) versionId = meta.json.versionId;
  } catch (err) {
    diagnostics.steps.push({ step: "metadata", error: String(err) });
  }
  diagnostics.versionId = versionId;

  // Steg 2: WOPI /files/{fileId}/contents  ← Trimble sin aktive nedlastings-API
  const wopiCandidates = [
    `${wopiBase}/files/${encodeURIComponent(fileId)}/contents`,
    `${wopiBase}/files/${encodeURIComponent(versionId)}/contents`,
    // Noen versjoner av WOPI bruker ?access_token i stedet for header
    `${wopiBase}/files/${encodeURIComponent(fileId)}/contents?access_token=${encodeURIComponent(token)}`
  ];

  for (const url of wopiCandidates) {
    try {
      // WOPI: send token i header, IKKE i URL for de to første
      const headers = url.includes("access_token=")
        ? {} // token allerede i URL
        : { ...auth };

      const res = await fetchAny(url, { method: "GET", headers });
      const isJson = res.contentType.includes("application/json");

      diagnostics.steps.push({
        step: "wopi-candidate",
        url: url.replace(token, "TOKEN"),
        status: res.status,
        ok: res.ok,
        contentType: res.contentType,
        preview: res.text.slice(0, 300)
      });

      if (res.ok && !isJson && res.text.length > 10) {
        return json(200, { ok: true, fileId, versionId, source: url.replace(token, "TOKEN"), via: "wopi", content: res.text, diagnostics });
      }
      // JSON-respons med innhold (uvanlig men mulig)
      if (res.ok && isJson && res.json && !res.json.errorcode) {
        return json(200, { ok: true, fileId, versionId, source: "wopi-json", via: "wopi-json", content: res.text, diagnostics });
      }
    } catch (err) {
      diagnostics.steps.push({ step: "wopi-candidate", url: url.replace(token, "TOKEN"), error: String(err) });
    }
  }

  // Steg 3: Fallback — prøv Core API med redirect-følging (noen ganger gir det en 302 til S3)
  try {
    const redirectUrl = `${coreBase}/files/${encodeURIComponent(versionId)}?download=true`;
    const res = await fetch(redirectUrl, {
      method: "GET",
      headers: { ...auth, Accept: "*/*" },
      redirect: "follow"
    });
    const contentType = res.headers.get("content-type") || "";
    const text = await res.text();
    const isJson = contentType.includes("application/json");

    diagnostics.steps.push({
      step: "core-redirect-follow",
      url: redirectUrl,
      status: res.status,
      ok: res.ok,
      contentType,
      preview: text.slice(0, 300)
    });

    if (res.ok && !isJson && text.length > 10) {
      return json(200, { ok: true, fileId, versionId, source: redirectUrl, via: "core-redirect", content: text, diagnostics });
    }
  } catch (err) {
    diagnostics.steps.push({ step: "core-redirect-follow", error: String(err) });
  }

  return json(502, {
    ok: false,
    error: "Alle nedlastingsstrategier feilet. Se diagnostics.",
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
  try { data = JSON.parse(event.body || "{}"); }
  catch { return json(400, { ok: false, error: "Invalid JSON" }); }

  const { action, token, url, method, fileId, projectLocation } = data;
  if (!token) return json(400, { ok: false, error: "Missing token" });

  if (action === "downloadKofFile") {
    if (!fileId) return json(400, { ok: false, error: "Missing fileId" });
    return handleDownloadKofFile(token, fileId, projectLocation);
  }

  // Enkel proxy (bakoverkompatibel)
  if (!url) return json(400, { ok: false, error: "Missing url" });
  if (!isAllowedHost(url)) return json(403, { ok: false, error: `Ikke tillatt: ${url}` });

  try {
    const response = await fetch(url, {
      method: method || "GET",
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json" }
    });
    const text = await response.text();
    return {
      statusCode: response.status,
      headers: { "Content-Type": response.headers.get("content-type") || "text/plain", "Access-Control-Allow-Origin": "*" },
      body: text
    };
  } catch (err) {
    return json(500, { ok: false, error: "Proxy error: " + String(err) });
  }
};
