// netlify/functions/tc-proxy.js

function isAllowedHost(url) {
  try {
    const host = new URL(url).hostname;
    return host.endsWith(".connect.trimble.com") || host.endsWith(".amazonaws.com");
  } catch { return false; }
}

function json(statusCode, body) {
  return {
    statusCode,
    headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" },
    body: JSON.stringify(body, null, 2)
  };
}

function getUrls(loc) {
  const l = String(loc || "").toLowerCase();
  if (l === "europe") return {
    core: "https://app21.connect.trimble.com/tc/api/2.0",
    wopi: "https://wopi-api-eu.connect.trimble.com/v1",
    projects: "https://projects-api-eu.connect.trimble.com/v1"
  };
  if (l === "asia") return {
    core: "https://app31.connect.trimble.com/tc/api/2.0",
    wopi: "https://wopi-api-ap.connect.trimble.com/v1",
    projects: "https://projects-api-ap.connect.trimble.com/v1"
  };
  return {
    core: "https://app.connect.trimble.com/tc/api/2.0",
    wopi: "https://wopi-api.connect.trimble.com/v1",
    projects: "https://projects-api.connect.trimble.com/v1"
  };
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
  const urls = getUrls(projectLocation);
  const auth = { Authorization: `Bearer ${token}` };
  const diagnostics = { fileId, projectLocation, urls, steps: [] };

  // Steg 1: Metadata
  let versionId = fileId;
  try {
    const metaUrl = `${urls.core}/files/${encodeURIComponent(fileId)}`;
    const meta = await fetchAny(metaUrl, { method: "GET", headers: { ...auth, Accept: "application/json" } });
    diagnostics.steps.push({ step: "metadata", url: metaUrl, status: meta.status, ok: meta.ok, preview: meta.text.slice(0, 300) });
    if (meta.ok && meta.json?.versionId) versionId = meta.json.versionId;
  } catch (err) {
    diagnostics.steps.push({ step: "metadata", error: String(err) });
  }
  diagnostics.versionId = versionId;

  // Alle kandidater å prøve
  const candidates = [
    // projects-api — sannsynlig aktiv nedlastings-API
    { label: "projects-api/download",   url: `${urls.projects}/files/${fileId}/download`,           headers: { ...auth } },
    { label: "projects-api/transfer",   url: `${urls.projects}/files/${fileId}/transfer`,           headers: { ...auth } },
    { label: "projects-api/content",    url: `${urls.projects}/files/${fileId}/content`,            headers: { ...auth } },
    { label: "projects-api/blob",       url: `${urls.projects}/files/${fileId}/blob`,               headers: { ...auth } },
    { label: "projects-api/versions",   url: `${urls.projects}/files/${fileId}/versions/${versionId}/download`, headers: { ...auth } },
    // WOPI
    { label: "wopi/contents",           url: `${urls.wopi}/files/${fileId}/contents`,               headers: { ...auth } },
    // Core med redirect
    { label: "core-redirect",           url: `${urls.core}/files/${fileId}?download=true`,          headers: { ...auth, Accept: "*/*" }, redirect: "follow" },
  ];

  for (const c of candidates) {
    try {
      const fetchOptions = { method: "GET", headers: c.headers };
      if (c.redirect) fetchOptions.redirect = c.redirect;
      const res = await fetchAny(c.url, fetchOptions);
      const isJson = res.contentType.includes("application/json");

      diagnostics.steps.push({
        step: c.label,
        url: c.url,
        status: res.status,
        ok: res.ok,
        contentType: res.contentType,
        preview: res.text.slice(0, 400)
      });

      if (!res.ok) continue;

      // Pre-signed URL i JSON-respons
      const presignedUrl = res.json?.url || res.json?.downloadUrl || res.json?.href || res.json?.transferUrl || null;
      if (isJson && presignedUrl) {
        diagnostics.steps.push({ step: "presigned-found", url: presignedUrl.slice(0, 80) });
        const fileRes = await fetchAny(presignedUrl, { method: "GET" });
        diagnostics.steps.push({ step: "presigned-fetch", status: fileRes.status, ok: fileRes.ok, bytes: fileRes.text.length });
        if (fileRes.ok) {
          return json(200, { ok: true, fileId, versionId, source: c.label, via: "presigned", content: fileRes.text, diagnostics });
        }
        continue;
      }

      // Direkte filinnhold
      if (!isJson && res.text.length > 10) {
        return json(200, { ok: true, fileId, versionId, source: c.label, via: "direct", content: res.text, diagnostics });
      }
    } catch (err) {
      diagnostics.steps.push({ step: c.label, url: c.url, error: String(err) });
    }
  }

  return json(502, { ok: false, error: "Alle strategier feilet.", fileId, versionId, diagnostics });
}

exports.handler = async function (event) {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "POST, OPTIONS", "Access-Control-Allow-Headers": "Content-Type" }, body: "" };
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

  if (!url) return json(400, { ok: false, error: "Missing url" });
  if (!isAllowedHost(url)) return json(403, { ok: false, error: `Ikke tillatt: ${url}` });

  try {
    const response = await fetch(url, { method: method || "GET", headers: { Authorization: `Bearer ${token}`, Accept: "application/json" } });
    const text = await response.text();
    return { statusCode: response.status, headers: { "Content-Type": response.headers.get("content-type") || "text/plain", "Access-Control-Allow-Origin": "*" }, body: text };
  } catch (err) {
    return json(500, { ok: false, error: "Proxy error: " + String(err) });
  }
};
