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

// ─── probe: utforsk hva projects-api og wopi faktisk tilbyr ──────────────────
async function handleProbe(token, projectId, fileId, projectLocation) {
  const urls = getUrls(projectLocation);
  const auth = { Authorization: `Bearer ${token}` };
  const results = [];

  const probeUrls = [
    // projects-api rot
    `${urls.projects}/`,
    `${urls.projects}/openapi`,
    `${urls.projects}/v3/api-docs`,
    `${urls.projects}/swagger-ui.html`,
    // projects-api med prosjekt
    `${urls.projects}/projects`,
    `${urls.projects}/projects/${projectId}`,
    `${urls.projects}/projects/${projectId}/files`,
    `${urls.projects}/projects/${projectId}/folders`,
    // wopi rot
    `${urls.wopi}/`,
    `${urls.wopi}/files/${fileId}`,
    `${urls.wopi}/projects/${projectId}/files/${fileId}`,
  ];

  for (const url of probeUrls) {
    try {
      const r = await fetchAny(url, { method: "GET", headers: { ...auth, Accept: "application/json" } });
      results.push({ url, status: r.status, contentType: r.contentType, preview: r.text.slice(0, 300) });
    } catch (err) {
      results.push({ url, error: String(err) });
    }
  }

  return json(200, { ok: true, results });
}

// ─── downloadKofFile ─────────────────────────────────────────────────────────
async function handleDownloadKofFile(token, fileId, projectLocation, projectId) {
  const urls = getUrls(projectLocation);
  const auth = { Authorization: `Bearer ${token}` };
  const diagnostics = { fileId, projectId, projectLocation, urls, steps: [] };

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

  const pid = projectId || "";
  const candidates = [
    { label: "projects-api/proj/files/download",    url: `${urls.projects}/projects/${pid}/files/${fileId}/download` },
    { label: "projects-api/proj/files/content",     url: `${urls.projects}/projects/${pid}/files/${fileId}/content` },
    { label: "projects-api/proj/files/transfer",    url: `${urls.projects}/projects/${pid}/files/${fileId}/transfer` },
    { label: "projects-api/proj/versions/download", url: `${urls.projects}/projects/${pid}/files/${fileId}/versions/${versionId}/download` },
    { label: "wopi/proj/files/contents",            url: `${urls.wopi}/projects/${pid}/files/${fileId}/contents` },
    { label: "core-no-accept",                      url: `${urls.core}/files/${fileId}`, noAccept: true },
  ];

  for (const c of candidates) {
    try {
      const headers = c.noAccept ? { Authorization: `Bearer ${token}` } : { ...auth, Accept: "application/json" };
      const res = await fetchAny(c.url, { method: "GET", headers, redirect: "follow" });
      const isJson = res.contentType.includes("application/json") || res.contentType.includes("problem+json");

      diagnostics.steps.push({ step: c.label, url: c.url, status: res.status, ok: res.ok, contentType: res.contentType, preview: res.text.slice(0, 300) });

      if (!res.ok) continue;

      const presignedUrl = res.json?.url || res.json?.downloadUrl || res.json?.href || res.json?.transferUrl || res.json?.fileUrl || null;
      if (isJson && presignedUrl) {
        const fileRes = await fetchAny(presignedUrl, { method: "GET" });
        diagnostics.steps.push({ step: "presigned-fetch", status: fileRes.status, ok: fileRes.ok, bytes: fileRes.text.length });
        if (fileRes.ok) return json(200, { ok: true, fileId, versionId, source: c.label, via: "presigned", content: fileRes.text, diagnostics });
        continue;
      }

      if (!isJson && res.text.length > 10) {
        return json(200, { ok: true, fileId, versionId, source: c.label, via: "direct", content: res.text, diagnostics });
      }
    } catch (err) {
      diagnostics.steps.push({ step: c.label, url: c.url, error: String(err) });
    }
  }

  return json(502, { ok: false, error: "Alle strategier feilet.", fileId, versionId, diagnostics });
}

// ─── Handler ─────────────────────────────────────────────────────────────────
exports.handler = async function (event) {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Methods": "POST, OPTIONS", "Access-Control-Allow-Headers": "Content-Type" }, body: "" };
  }
  if (event.httpMethod !== "POST") return { statusCode: 405, body: "Use POST" };

  let data;
  try { data = JSON.parse(event.body || "{}"); }
  catch { return json(400, { ok: false, error: "Invalid JSON" }); }

  const { action, token, url, method, fileId, projectLocation, projectId } = data;
  if (!token) return json(400, { ok: false, error: "Missing token" });

  if (action === "probe") {
    return handleProbe(token, projectId, fileId, projectLocation);
  }

  if (action === "downloadKofFile") {
    if (!fileId) return json(400, { ok: false, error: "Missing fileId" });
    return handleDownloadKofFile(token, fileId, projectLocation, projectId);
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
