exports.handler = async function handler(event) {
  try {
    if (event.httpMethod !== "POST") {
      return jsonResponse(405, {
        ok: false,
        error: "Method not allowed"
      });
    }

    const body = safeJsonParse(event.body) || {};
const { action } = body;

if (action === "downloadKofFile") {
  return await handleDownloadKofFile(body);
}

if (action === "probeCore") {
  return await handleProbeCore(body);
}

if (action === "listProjectKofFiles") {
  return await handleListProjectKofFiles(body);
}

if (action === "uploadConvertedTxt") {
  return await handleUploadConvertedTxt(body);
}

    return jsonResponse(400, {
      ok: false,
      error: `Unknown action: ${String(action)}`
    });
  } catch (err) {
    console.error("tc-proxy fatal:", err);

    return jsonResponse(500, {
      ok: false,
      error: err?.message || String(err)
    });
  }
};

function jsonResponse(statusCode, data) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store"
    },
    body: JSON.stringify(data, null, 2)
  };
}

function safeJsonParse(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function shortText(text, max = 1200) {
  if (typeof text !== "string") return text;
  return text.length > max ? text.slice(0, max) + "..." : text;
}

function safeHost(url) {
  try {
    return new URL(url).host;
  } catch {
    return null;
  }
}

function getCoreBaseUrl(projectLocation) {
  const loc = String(projectLocation || "").toLowerCase();

  if (loc === "northamerica" || loc === "us" || !loc) {
    return "https://app.connect.trimble.com/tc/api/2.0";
  }

  if (loc === "europe") {
    return "https://app.eu.connect.trimble.com/tc/api/2.0";
  }

  if (loc === "asia") {
    return "https://app.asia.connect.trimble.com/tc/api/2.0";
  }

  return "https://app.connect.trimble.com/tc/api/2.0";
}

async function fetchRaw(url, options = {}, timeoutMs = 30000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(url, {
      ...options,
      signal: controller.signal
    });

    const contentType = res.headers.get("content-type") || "";
    const text = await res.text();
    const json = safeJsonParse(text);

    return {
      url,
      ok: res.ok,
      status: res.status,
      statusText: res.statusText,
      contentType,
      text,
      json,
      headers: Object.fromEntries(res.headers.entries())
    };
  } finally {
    clearTimeout(timer);
  }
}

async function fetchWithBearer(url, token, options = {}, timeoutMs = 30000) {
  const headers = {
    ...(options.headers || {}),
    Authorization: `Bearer ${token}`
  };

  return fetchRaw(
    url,
    {
      ...options,
      headers
    },
    timeoutMs
  );
}

async function fetchJsonWithBearer(url, token, extraHeaders = {}) {
  return fetchWithBearer(url, token, {
    method: "GET",
    headers: {
      ...extraHeaders
    }
  });
}

async function fetchTextNoAuth(url) {
  return fetchRaw(
    url,
    {
      method: "GET"
    },
    60000
  );
}

function extractPossibleUrl(payload) {
  if (!payload || typeof payload !== "object") return null;

  return (
    payload.uploadUrl ||
    payload.url ||
    payload.href ||
    payload.link ||
    payload.signedUrl ||
    payload.presignedUrl ||
    payload.data?.uploadUrl ||
    payload.data?.url ||
    payload.details?.uploadUrl ||
    payload.details?.url ||
    payload.result?.uploadUrl ||
    payload.result?.url ||
    null
  );
}

function normalizeUploadTarget(fileName) {
  const name = String(fileName || "").trim() || "output.txt";
  return name.toLowerCase().endsWith(".txt") ? name : `${name}.txt`;
}

async function getFileMetadata({ token, projectLocation, fileId }) {
  const base = getCoreBaseUrl(projectLocation);
  const url = `${base}/files/${encodeURIComponent(fileId)}`;

  const res = await fetchJsonWithBearer(url, token);

  return {
    ok: res.ok,
    url,
    status: res.status,
    contentType: res.contentType,
    preview: shortText(res.text, 1000),
    data: res.json
  };
}

async function getFileVersions({ token, projectLocation, fileId }) {
  const base = getCoreBaseUrl(projectLocation);
  const url = `${base}/files/${encodeURIComponent(fileId)}/versions?tokenThumburl=false`;

  const res = await fetchJsonWithBearer(url, token);

  return {
    ok: res.ok,
    url,
    status: res.status,
    contentType: res.contentType,
    preview: shortText(res.text, 1000),
    data: res.json
  };
}

async function tryCoreCandidates({ token, projectLocation, fileId, versionId }) {
  const base = getCoreBaseUrl(projectLocation);

  const candidates = [
    {
      name: "fs-downloadurl",
      url: `${base}/files/fs/${encodeURIComponent(fileId)}/downloadurl?versionId=${encodeURIComponent(versionId)}`,
      bearer: true
    },
    {
      name: "fs-downloadurl-versionId-path",
      url: `${base}/files/fs/${encodeURIComponent(versionId)}/downloadurl?versionId=${encodeURIComponent(versionId)}`,
      bearer: true
    },
    {
      name: "blobstore-versionId",
      url: `${base}/files/${encodeURIComponent(versionId)}/blobstore`,
      bearer: true
    },
    {
      name: "download-versionId",
      url: `${base}/files/${encodeURIComponent(versionId)}/download`,
      bearer: true
    },
    {
      name: "content-versionId",
      url: `${base}/files/${encodeURIComponent(versionId)}/content`,
      bearer: true
    },
    {
      name: "download-fileId",
      url: `${base}/files/${encodeURIComponent(fileId)}/download`,
      bearer: true
    },
    {
      name: "content-fileId",
      url: `${base}/files/${encodeURIComponent(fileId)}/content`,
      bearer: true
    },
    {
      name: "file-download-true",
      url: `${base}/files/${encodeURIComponent(fileId)}?download=true`,
      bearer: true
    },
    {
      name: "file-content-true",
      url: `${base}/files/${encodeURIComponent(fileId)}?content=true`,
      bearer: true
    },
    {
      name: "file-plain",
      url: `${base}/files/${encodeURIComponent(fileId)}`,
      bearer: true
    }
  ];

  const diagnostics = [];

  for (const candidate of candidates) {
    try {
      const res = candidate.bearer
        ? await fetchJsonWithBearer(candidate.url, token)
        : await fetchRaw(candidate.url);

      const signedUrl = extractPossibleUrl(res.json);

      const looksLikeText =
        typeof res.text === "string" &&
        res.text.length > 0 &&
        !res.contentType.includes("application/json") &&
        !res.contentType.includes("text/html");

      diagnostics.push({
        name: candidate.name,
        url: candidate.url,
        status: res.status,
        ok: res.ok,
        contentType: res.contentType,
        foundSignedUrl: !!signedUrl,
        looksLikeText,
        preview: shortText(res.text, 600)
      });

      if (signedUrl) {
        const fileRes = await fetchTextNoAuth(signedUrl);

        return {
          ok: fileRes.ok,
          source: candidate.name,
          mode: "signedUrl",
          signedUrlHost: safeHost(signedUrl),
          signedUrl,
          diagnostics,
          fileFetch: {
            status: fileRes.status,
            ok: fileRes.ok,
            contentType: fileRes.contentType,
            preview: shortText(fileRes.text, 600)
          },
          text: fileRes.ok ? fileRes.text : null,
          contentType: fileRes.contentType
        };
      }

      if (res.ok && looksLikeText) {
        return {
          ok: true,
          source: candidate.name,
          mode: "directText",
          diagnostics,
          text: res.text,
          contentType: res.contentType
        };
      }
    } catch (err) {
      diagnostics.push({
        name: candidate.name,
        url: candidate.url,
        ok: false,
        error: err?.message || String(err)
      });
    }
  }

  return {
    ok: false,
    error: "Fant ikke fungerende Core download-kandidat.",
    diagnostics
  };
}

async function handleDownloadKofFile(body) {
  const { token, projectId, projectLocation, fileId, fileName } = body;

  if (!token || !fileId) {
    return jsonResponse(400, {
      ok: false,
      error: "Mangler token eller fileId"
    });
  }

  const metadata = await getFileMetadata({
    token,
    projectLocation,
    fileId
  });

  if (!metadata.ok || !metadata.data) {
    return jsonResponse(200, {
      ok: false,
      step: "metadata",
      project: {
        id: projectId,
        location: projectLocation
      },
      file: {
        id: fileId,
        name: fileName
      },
      metadata
    });
  }

  const versions = await getFileVersions({
    token,
    projectLocation,
    fileId
  });

  let versionId = metadata.data.versionId || metadata.data.id;

  if (versions.ok && Array.isArray(versions.data) && versions.data.length > 0) {
    versionId =
      versions.data[0]?.versionId ||
      versions.data[0]?.id ||
      versions.data[0]?.version?.id ||
      versionId;
  }

  const download = await tryCoreCandidates({
    token,
    projectLocation,
    fileId,
    versionId
  });

  if (!download.ok) {
    return jsonResponse(200, {
      ok: false,
      step: "download",
      project: {
        id: projectId,
        location: projectLocation
      },
      file: {
        id: metadata.data.id,
        versionId,
        name: metadata.data.name || fileName,
        parentId: metadata.data.parentId || null
      },
      metadata: {
        status: metadata.status,
        preview: metadata.preview
      },
      versions: {
        ok: versions.ok,
        status: versions.status,
        preview: versions.preview
      },
      download
    });
  }

  return jsonResponse(200, {
    ok: true,
    project: {
      id: projectId,
      location: projectLocation
    },
    file: {
      id: metadata.data.id,
      versionId,
      name: metadata.data.name || fileName,
      parentId: metadata.data.parentId || null
    },
    source: {
      candidate: download.source,
      mode: download.mode,
      signedUrlHost: download.signedUrlHost || null
    },
    contentType: download.contentType,
    text: download.text,
    diagnostics: {
      metadata: {
        status: metadata.status,
        preview: metadata.preview
      },
      versions: {
        status: versions.status,
        preview: versions.preview
      },
      downloadDiagnostics: download.diagnostics
    }
  });
}

async function handleProbeCore(body) {
  const { token, projectId, projectLocation, fileId, fileName } = body;

  if (!token || !fileId) {
    return jsonResponse(400, {
      ok: false,
      error: "Mangler token eller fileId"
    });
  }

  const metadata = await getFileMetadata({
    token,
    projectLocation,
    fileId
  });

  const versions = await getFileVersions({
    token,
    projectLocation,
    fileId
  });

  let versionId = metadata.data?.versionId || metadata.data?.id || fileId;

  if (versions.ok && Array.isArray(versions.data) && versions.data.length > 0) {
    versionId =
      versions.data[0]?.versionId ||
      versions.data[0]?.id ||
      versions.data[0]?.version?.id ||
      versionId;
  }

  const probe = await tryCoreCandidates({
    token,
    projectLocation,
    fileId,
    versionId
  });

  return jsonResponse(200, {
    ok: true,
    probe: "core",
    project: {
      id: projectId,
      location: projectLocation
    },
    file: {
      id: fileId,
      name: fileName,
      versionId
    },
    metadata,
    versions,
    probeResult: probe
  });
}

async function handleListProjectKofFiles(body) {
  const { token, projectId, projectLocation } = body;

  if (!token || !projectId) {
    return jsonResponse(400, {
      ok: false,
      error: "Mangler token eller projectId"
    });
  }

  const listResult = await tryListProjectFilesCandidates({
    token,
    projectId,
    projectLocation
  });

  return jsonResponse(200, listResult);
}

async function handleUploadConvertedTxt(body) {
  const {
    token,
    projectId,
    projectLocation,
    fileName,
    content,
    parentId = null
  } = body;

  if (!token || !projectId || !fileName || typeof content !== "string") {
    return jsonResponse(400, {
      ok: false,
      error: "Mangler token, projectId, fileName eller content"
    });
  }

  if (!parentId) {
    return jsonResponse(400, {
      ok: false,
      error: "Mangler parentId for upload"
    });
  }

  const normalizedFileName = normalizeUploadTarget(fileName);
  const base = getCoreBaseUrl(projectLocation);

  // Dette er den viktigste endringen:
  // direkte upload til /files?parentId=...
  const uploadUrl = `${base}/files?parentId=${encodeURIComponent(parentId)}`;

  const diagnostics = [];

  try {
    // Variant 1: multipart/form-data med "file"
    const form = new FormData();
    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    form.append("file", blob, normalizedFileName);

    const multipartRes = await fetch(uploadUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`
      },
      body: form
    });

    const multipartText = await multipartRes.text();
    const multipartJson = safeJsonParse(multipartText);

    diagnostics.push({
      step: "multipartUpload",
      url: uploadUrl,
      ok: multipartRes.ok,
      status: multipartRes.status,
      preview: shortText(multipartText, 700)
    });

    if (multipartRes.ok) {
      return jsonResponse(200, {
        ok: true,
        action: "uploadConvertedTxt",
        project: {
          id: projectId,
          location: projectLocation
        },
        file: {
          name: normalizedFileName,
          parentId
        },
        uploadResult: multipartJson || multipartText,
        diagnostics
      });
    }

    // Variant 2: rå tekstbody
    const rawRes = await fetch(uploadUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "text/plain;charset=utf-8",
        "X-File-Name": normalizedFileName
      },
      body: content
    });

    const rawText = await rawRes.text();
    const rawJson = safeJsonParse(rawText);

    diagnostics.push({
      step: "rawUpload",
      url: uploadUrl,
      ok: rawRes.ok,
      status: rawRes.status,
      preview: shortText(rawText, 700)
    });

    if (rawRes.ok) {
      return jsonResponse(200, {
        ok: true,
        action: "uploadConvertedTxt",
        project: {
          id: projectId,
          location: projectLocation
        },
        file: {
          name: normalizedFileName,
          parentId
        },
        uploadResult: rawJson || rawText,
        diagnostics
      });
    }

    return jsonResponse(200, {
      ok: false,
      action: "uploadConvertedTxt",
      error: "Direkte upload feilet.",
      project: {
        id: projectId,
        location: projectLocation
      },
      file: {
        name: normalizedFileName,
        parentId
      },
      diagnostics
    });
  } catch (err) {
    return jsonResponse(500, {
      ok: false,
      action: "uploadConvertedTxt",
      error: err?.message || String(err),
      project: {
        id: projectId,
        location: projectLocation
      },
      file: {
        name: normalizedFileName,
        parentId
      },
      diagnostics
    });
  }
}

  const diagnostics = [];

  for (const candidate of createCandidates) {
    try {
      const bodyPayload = Object.fromEntries(
        Object.entries(candidate.body).filter(([, v]) => v !== undefined && v !== null && v !== "")
      );

      const createRes = await fetchWithBearer(
        candidate.url,
        token,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify(bodyPayload)
        },
        60000
      );

      const uploadUrl = extractPossibleUrl(createRes.json);
      const createdFileId =
        createRes.json?.id ||
        createRes.json?.fileId ||
        createRes.json?.details?.id ||
        createRes.json?.data?.id ||
        null;

      diagnostics.push({
        step: "create",
        candidate: candidate.name,
        url: candidate.url,
        ok: createRes.ok,
        status: createRes.status,
        foundUploadUrl: !!uploadUrl,
        createdFileId,
        preview: shortText(createRes.text, 700)
      });

      if (!createRes.ok || !uploadUrl) {
        continue;
      }

      const uploadAttempts = [
        {
          name: "put-text-plain",
          options: {
            method: "PUT",
            headers: {
              "Content-Type": "text/plain;charset=utf-8"
            },
            body: content
          }
        },
        {
          name: "put-octet-stream",
          options: {
            method: "PUT",
            headers: {
              "Content-Type": "application/octet-stream"
            },
            body: content
          }
        },
        {
          name: "post-text-plain",
          options: {
            method: "POST",
            headers: {
              "Content-Type": "text/plain;charset=utf-8"
            },
            body: content
          }
        }
      ];

      for (const uploadAttempt of uploadAttempts) {
        const uploadRes = await fetchRaw(uploadUrl, uploadAttempt.options, 60000);

        diagnostics.push({
          step: "upload",
          candidate: candidate.name,
          attempt: uploadAttempt.name,
          uploadUrlHost: safeHost(uploadUrl),
          ok: uploadRes.ok,
          status: uploadRes.status,
          preview: shortText(uploadRes.text, 700)
        });

        if (uploadRes.ok) {
          return jsonResponse(200, {
            ok: true,
            action: "uploadConvertedTxt",
            project: {
              id: projectId,
              location: projectLocation
            },
            file: {
              id: createdFileId,
              name: normalizedFileName,
              parentId
            },
            upload: {
              createCandidate: candidate.name,
              uploadAttempt: uploadAttempt.name,
              uploadUrlHost: safeHost(uploadUrl)
            },
            diagnostics
          });
        }
      }
    } catch (err) {
      diagnostics.push({
        step: "exception",
        ok: false,
        candidate: candidate.name,
        error: err?.message || String(err)
      });
    }
  }

  return jsonResponse(200, {
    ok: false,
    action: "uploadConvertedTxt",
    error: "Fant ingen fungerende upload-kandidat.",
    project: {
      id: projectId,
      location: projectLocation
    },
    file: {
      name: normalizedFileName,
      parentId
    },
    diagnostics
  });
}

async function tryListProjectFilesCandidates({ token, projectId, projectLocation }) {
  const base = getCoreBaseUrl(projectLocation);

  const candidates = [
    {
      name: "projects-files-recursive",
      url: `${base}/projects/${encodeURIComponent(projectId)}/files?recursive=true`
    },
    {
      name: "projects-files",
      url: `${base}/projects/${encodeURIComponent(projectId)}/files`
    },
    {
      name: "projects-files-includePath",
      url: `${base}/projects/${encodeURIComponent(projectId)}/files?includeFolderPath=true`
    },
    {
      name: "projects-search-kof",
      url: `${base}/projects/${encodeURIComponent(projectId)}/search?query=.kof&type=file`
    },
    {
      name: "search-kof",
      url: `${base}/search?projectId=${encodeURIComponent(projectId)}&query=.kof&type=file`
    },
    {
      name: "projects-folders-root",
      url: `${base}/projects/${encodeURIComponent(projectId)}/folders`
    }
  ];

  const diagnostics = [];

  for (const candidate of candidates) {
    try {
      const res = await fetchJsonWithBearer(candidate.url, token);

      diagnostics.push({
        name: candidate.name,
        url: candidate.url,
        ok: res.ok,
        status: res.status,
        preview: shortText(res.text, 800)
      });

      if (!res.ok || !res.json) {
        continue;
      }

      const files = normalizeFilesFromAnyResponse(res.json)
        .filter((f) => f && f.id && isKofName(f.name))
        .sort((a, b) => String(a.name).localeCompare(String(b.name), undefined, { sensitivity: "base" }));

      if (files.length) {
        return {
          ok: true,
          action: "listProjectKofFiles",
          project: {
            id: projectId,
            location: projectLocation
          },
          source: candidate.name,
          candidatesTried: diagnostics.length,
          files,
          diagnostics
        };
      }
    } catch (err) {
      diagnostics.push({
        name: candidate.name,
        url: candidate.url,
        ok: false,
        error: err?.message || String(err)
      });
    }
  }

  return {
    ok: false,
    action: "listProjectKofFiles",
    error: "Fant ingen fungerende kandidat for fillisting, eller ingen .kof-filer ble funnet.",
    project: {
      id: projectId,
      location: projectLocation
    },
    candidatesTried: diagnostics.length,
    diagnostics
  };
}

function isKofName(name) {
  return /\.kof$/i.test(String(name || ""));
}

function normalizePathValue(pathValue) {
  if (!pathValue) return "";

  if (typeof pathValue === "string") {
    return pathValue;
  }

  if (Array.isArray(pathValue)) {
    return pathValue
      .map((item) => {
        if (!item) return "";
        if (typeof item === "string") return item;
        if (typeof item === "object") return item.name || item.title || item.id || "";
        return "";
      })
      .filter(Boolean)
      .join("/");
  }

  if (typeof pathValue === "object") {
    return pathValue.name || pathValue.title || pathValue.id || "";
  }

  return String(pathValue);
}

function normalizeFilesFromAnyResponse(payload) {
  const out = [];
  const seen = new Set();

  walkAny(payload, [], out, seen);

  return out;
}

function walkAny(node, pathParts, out, seen) {
  if (node == null) return;

  if (Array.isArray(node)) {
    for (const item of node) {
      walkAny(item, pathParts, out, seen);
    }
    return;
  }

  if (typeof node !== "object") {
    return;
  }

  const maybeName =
    node.name ||
    node.fileName ||
    node.filename ||
    node.title ||
    null;

  const maybeId =
    node.id ||
    node.fileId ||
    node.versionId ||
    null;

  const maybeParentId =
    node.parentId ||
    node.parent?.id ||
    node.details?.parentId ||
    null;

  const maybeVersionId =
    node.versionId ||
    node.details?.versionId ||
    null;

  const maybePath =
    node.path ||
    node.folderPath ||
    node.fullPath ||
    node.location ||
    node.details?.path ||
    null;

  const details = node.details && typeof node.details === "object" ? node.details : null;

  const effectiveName =
    maybeName ||
    details?.name ||
    details?.fileName ||
    null;

  const effectiveId =
    maybeId ||
    details?.id ||
    details?.fileId ||
    null;

  const effectiveParentId =
    maybeParentId ||
    details?.parentId ||
    null;

  const effectiveVersionId =
    maybeVersionId ||
    details?.versionId ||
    null;

  const effectivePath =
    maybePath ||
    details?.path ||
    null;

  const childPath = effectiveName ? [...pathParts, effectiveName] : pathParts;

  if (effectiveId && effectiveName) {
    const normalized = {
      id: String(effectiveId),
      name: String(effectiveName),
      versionId: effectiveVersionId ? String(effectiveVersionId) : null,
      parentId: effectiveParentId ? String(effectiveParentId) : null,
      path: effectivePath ? normalizePathValue(effectivePath) : buildPath(pathParts)
    };

    const key = `${normalized.id}|${normalized.name}|${normalized.path}`;
    if (!seen.has(key)) {
      seen.add(key);
      out.push(normalized);
    }
  }

  for (const [key, value] of Object.entries(node)) {
    if (
      key === "parent" ||
      key === "parents" ||
      key === "_links" ||
      key === "links" ||
      key === "permissions"
    ) {
      continue;
    }

    if (Array.isArray(value) || (value && typeof value === "object")) {
      walkAny(value, childPath, out, seen);
    }
  }
}

function buildPath(parts) {
  const p = (parts || []).filter(Boolean).map((x) => String(x).trim()).filter(Boolean);
  return p.length ? p.join("/") : "";
}
