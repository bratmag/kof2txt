function json(statusCode, body) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body, null, 2)
  };
}

function normalizeLocation(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function pickRegionByProjectLocation(regions, projectLocation) {
  const wanted = normalizeLocation(projectLocation);

  if (!Array.isArray(regions) || regions.length === 0) {
    return null;
  }

  const exact =
    regions.find((r) => normalizeLocation(r.name) === wanted) ||
    regions.find((r) => normalizeLocation(r.location) === wanted) ||
    regions.find((r) => normalizeLocation(r.region) === wanted) ||
    regions.find((r) => normalizeLocation(r.id) === wanted);

  if (exact) return exact;

  if (wanted.includes("europe") || wanted === "eu") {
    return (
      regions.find((r) => normalizeLocation(r.name).includes("europe")) ||
      regions.find((r) => normalizeLocation(r.location).includes("europe")) ||
      null
    );
  }

  if (wanted.includes("asia")) {
    return (
      regions.find((r) => normalizeLocation(r.name).includes("asia")) ||
      regions.find((r) => normalizeLocation(r.location).includes("asia")) ||
      null
    );
  }

  if (
    wanted.includes("us") ||
    wanted.includes("america") ||
    wanted.includes("north")
  ) {
    return (
      regions.find((r) => normalizeLocation(r.name).includes("north")) ||
      regions.find((r) => normalizeLocation(r.name).includes("america")) ||
      regions.find((r) => normalizeLocation(r.location).includes("america")) ||
      null
    );
  }

  return null;
}

function collectCandidateOrigins(region) {
  const values = [
    region?.origin,
    region?.api,
    region?.apiOrigin,
    region?.core,
    region?.coreOrigin,
    region?.baseUrl,
    region?.baseURL,
    region?.url
  ].filter(Boolean);

  const deduped = [];
  for (const v of values) {
    if (!deduped.includes(v)) deduped.push(v);
  }
  return deduped;
}

function buildListCandidates(origin, projectId) {
  const base = String(origin || "").replace(/\/+$/, "");

  return [
    `${base}/projects/${projectId}/files`,
    `${base}/project/${projectId}/files`,
    `${base}/connect/projects/${projectId}/files`,
    `${base}/connect/project/${projectId}/files`,
    `${base}/v1/projects/${projectId}/files`,
    `${base}/v1/project/${projectId}/files`,
    `${base}/files?projectId=${encodeURIComponent(projectId)}`,
    `${base}/connect/files?projectId=${encodeURIComponent(projectId)}`
  ];
}

function isObject(value) {
  return value && typeof value === "object" && !Array.isArray(value);
}

function flattenPossibleFileList(payload) {
  if (Array.isArray(payload)) return payload;
  if (!isObject(payload)) return [];

  const keysToTry = [
    "items",
    "files",
    "data",
    "results",
    "entries",
    "documents",
    "children"
  ];

  for (const key of keysToTry) {
    if (Array.isArray(payload[key])) return payload[key];
  }

  for (const key of keysToTry) {
    if (isObject(payload[key])) {
      for (const nestedKey of keysToTry) {
        if (Array.isArray(payload[key][nestedKey])) {
          return payload[key][nestedKey];
        }
      }
    }
  }

  return [];
}

function mapFile(raw) {
  const name =
    raw?.name ||
    raw?.fileName ||
    raw?.filename ||
    raw?.title ||
    raw?.displayName ||
    null;

  const path =
    raw?.path ||
    raw?.fullPath ||
    raw?.location ||
    raw?.parentPath ||
    null;

  return {
    id: raw?.id || raw?.fileId || raw?.identifier || null,
    name,
    path,
    size: raw?.size ?? raw?.fileSize ?? null,
    raw
  };
}

function isKofFile(file) {
  const candidates = [file?.name, file?.path].filter(Boolean);
  return candidates.some((v) => String(v).toLowerCase().endsWith(".kof"));
}

async function fetchJsonOrText(url, options) {
  const response = await fetch(url, options);
  const text = await response.text();

  let parsed = null;
  try {
    parsed = JSON.parse(text);
  } catch {
    parsed = null;
  }

  return {
    status: response.status,
    ok: response.ok,
    headers: Object.fromEntries(response.headers.entries()),
    text,
    json: parsed
  };
}

async function discoverRegions() {
  const masterUrl = "https://api.connect.trimble.com/regions";
  return fetchJsonOrText(masterUrl, {
    method: "GET",
    headers: {
      Accept: "application/json"
    }
  });
}

// ─── NY FUNKSJON: Last ned KOF-fil via Trimble Core API ───────────────────────
//
// Trimble sin nedlastingsflyt:
//   1. GET /tc/api/2.0/files/{fileId}          → metadata, inneholder versionId
//   2. GET /tc/api/2.0/files/{versionId}/blobstore  → { url: "<pre-signed S3 URL>" }
//   3. GET <pre-signed URL>                     → råtekst (uten Authorization-header)
//
// Faller tilbake på alternative endepunkter hvis blobstore ikke svarer.
// ──────────────────────────────────────────────────────────────────────────────
async function handleDownloadKofFile(token, fileId, projectLocation) {
  // Velg riktig base-URL basert på region
  const loc = normalizeLocation(projectLocation);
  const isEurope =
    loc.includes("europe") || loc === "eu" || loc.includes("eu");

  const baseUrl = isEurope
    ? "https://app.eu.connect.trimble.com/tc/api/2.0"
    : "https://app.connect.trimble.com/tc/api/2.0";

  const authHeaders = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json"
  };

  const diagnostics = { fileId, projectLocation, baseUrl, steps: [] };

  // Steg 1: Hent metadata for å få versionId
  let versionId = fileId; // fallback: bruk fileId direkte
  try {
    const metaRes = await fetchJsonOrText(
      `${baseUrl}/files/${fileId}`,
      { method: "GET", headers: authHeaders }
    );

    diagnostics.steps.push({
      step: "metadata",
      url: `${baseUrl}/files/${fileId}`,
      status: metaRes.status,
      ok: metaRes.ok,
      preview: metaRes.text.slice(0, 300)
    });

    if (metaRes.ok && metaRes.json?.versionId) {
      versionId = metaRes.json.versionId;
    }
  } catch (err) {
    diagnostics.steps.push({ step: "metadata", error: String(err) });
  }

  diagnostics.versionId = versionId;

  // Steg 2: Prøv å hente pre-signed URL via blobstore
  const downloadCandidates = [
    `${baseUrl}/files/${versionId}/blobstore`,
    `${baseUrl}/files/${versionId}/download`,
    `${baseUrl}/files/${versionId}/blob`,
    `${baseUrl}/files/${versionId}/content`
  ];

  for (const url of downloadCandidates) {
    try {
      const res = await fetchJsonOrText(url, {
        method: "GET",
        headers: authHeaders
      });

      const ct = res.headers["content-type"] || "";
      const isJson = ct.includes("application/json");

      diagnostics.steps.push({
        step: "download-attempt",
        url,
        status: res.status,
        contentType: ct,
        ok: res.ok,
        preview: res.text.slice(0, 300)
      });

      if (!res.ok) continue;

      // Scenario A: Svaret er JSON med en pre-signed URL
      if (isJson && res.json?.url) {
        const presignedUrl = res.json.url;
        diagnostics.steps.push({
          step: "presigned-fetch",
          url: presignedUrl.slice(0, 80) + "..."
        });

        // Hent fra pre-signed URL *uten* Authorization-header (viktig!)
        const fileRes = await fetch(presignedUrl, {
          method: "GET",
          headers: { Accept: "*/*" }
        });

        const fileText = await fileRes.text();

        diagnostics.steps.push({
          step: "presigned-result",
          status: fileRes.status,
          ok: fileRes.ok,
          byteLength: fileText.length
        });

        if (fileRes.ok) {
          return json(200, {
            ok: true,
            fileId,
            versionId,
            source: url,
            content: fileText,
            diagnostics
          });
        }
      }

      // Scenario B: Svaret er direkte filinnhold (tekst, ikke JSON)
      if (!isJson && res.text.length > 0) {
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
        step: "download-attempt",
        url,
        error: String(err)
      });
    }
  }

  // Ingenting fungerte
  return json(502, {
    ok: false,
    error: "Fant ikke fungerende nedlastings-endepunkt",
    fileId,
    versionId,
    diagnostics
  });
}
// ──────────────────────────────────────────────────────────────────────────────

exports.handler = async function (event) {
  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      body: "Use POST"
    };
  }

  try {
    const data = JSON.parse(event.body || "{}");
    const { action, token, projectId, projectLocation, fileId } = data || {};

    if (!token) {
      return json(400, { ok: false, error: "Missing token" });
    }

    if (!action) {
      return json(400, { ok: false, error: "Missing action" });
    }

    // ── Action: Last ned KOF-fil ──────────────────────────────────────────────
    if (action === "downloadKofFile") {
      if (!fileId) {
        return json(400, { ok: false, error: "Missing fileId" });
      }
      return handleDownloadKofFile(token, fileId, projectLocation);
    }

    // ── Action: List KOF-filer ────────────────────────────────────────────────
    if (action !== "listKofFiles") {
      return json(400, { ok: false, error: `Unknown action: ${action}` });
    }

    if (!projectId) {
      return json(400, { ok: false, error: "Missing projectId" });
    }

    const regionDiscovery = await discoverRegions();

    if (!regionDiscovery.ok) {
      return json(502, {
        ok: false,
        error: "Region discovery failed",
        diagnostics: {
          status: regionDiscovery.status,
          preview: regionDiscovery.text.slice(0, 1000)
        }
      });
    }

    const regionPayload = regionDiscovery.json;
    const regions = Array.isArray(regionPayload)
      ? regionPayload
      : Array.isArray(regionPayload?.items)
      ? regionPayload.items
      : Array.isArray(regionPayload?.regions)
      ? regionPayload.regions
      : [];

    const pickedRegion = pickRegionByProjectLocation(regions, projectLocation);

    if (!pickedRegion) {
      return json(502, {
        ok: false,
        error: "Could not match project region from /regions result",
        diagnostics: {
          projectLocation,
          regionCount: regions.length,
          sampleRegions: regions.slice(0, 5)
        }
      });
    }

    const origins = collectCandidateOrigins(pickedRegion);

    if (origins.length === 0) {
      return json(502, {
        ok: false,
        error: "Matched region has no usable API origin",
        diagnostics: {
          projectLocation,
          pickedRegion
        }
      });
    }

    const diagnostics = {
      projectId,
      projectLocation,
      pickedRegion,
      tried: []
    };

    for (const origin of origins) {
      const candidates = buildListCandidates(origin, projectId);

      for (const url of candidates) {
        try {
          const result = await fetchJsonOrText(url, {
            method: "GET",
            headers: {
              Authorization: `Bearer ${token}`,
              Accept: "application/json"
            }
          });

          const preview =
            result.json && typeof result.json === "object"
              ? result.json
              : result.text.slice(0, 500);

          diagnostics.tried.push({
            url,
            status: result.status,
            ok: result.ok,
            preview
          });

          if (!result.ok) {
            continue;
          }

          const list = flattenPossibleFileList(result.json ?? result.text);
          const mapped = list.map(mapFile);
          const kofFiles = mapped.filter(isKofFile);

          return json(200, {
            ok: true,
            projectId,
            projectLocation,
            usedUrl: url,
            matchedRegion: pickedRegion,
            totalFilesSeen: mapped.length,
            kofFiles,
            diagnostics
          });
        } catch (err) {
          diagnostics.tried.push({
            url,
            ok: false,
            fetchError: String(err?.message || err)
          });
        }
      }
    }

    return json(502, {
      ok: false,
      error: "Could not find a working file-list endpoint yet",
      diagnostics
    });
  } catch (err) {
    return json(500, {
      ok: false,
      error: String(err),
      message: err?.message || null,
      cause: err?.cause ? String(err.cause) : null,
      stack: err?.stack || null
    });
  }
};
