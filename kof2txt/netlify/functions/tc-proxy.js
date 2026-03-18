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

  // Denne versjonen er laget for North America-test.
  // Europe/Asia kan legges til senere hvis dere vil.
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
      json
    };
  } finally {
    clearTimeout(timer);
  }
}

async function fetchJsonWithBearer(url, token, extraHeaders = {}) {
  return fetchRaw(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      ...extraHeaders
    }
  });
}

async function fetchTextNoAuth(url) {
  return fetchRaw(url, {
    method: "GET"
  }, 60000);
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

async function tryCoreCandidates({ token, projectLocation, fileId, versionId }) {
  const base = getCoreBaseUrl(projectLocation);

  const candidates = [
    {
  name: "fs-downloadurl",
  url: `${base}/files/fs/${encodeURIComponent(fileId)}/downloadurl?versionId=${encodeURIComponent(versionId)}`,
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

      const signedUrl =
        res.json?.url ||
        res.json?.downloadUrl ||
        res.json?.href ||
        res.json?.link ||
        null;

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

  const versionId = metadata.data.versionId || metadata.data.id;

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
        name: metadata.data.name || fileName
      },
      metadata: {
        status: metadata.status,
        preview: metadata.preview
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
      name: metadata.data.name || fileName
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

  const versionId = metadata.data?.versionId || metadata.data?.id || fileId;

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
    probeResult: probe
  });
}
