/* global TrimbleConnectWorkspace */

const statusEl = document.getElementById("status");
const outputEl = document.getElementById("output");

function setStatus(text) {
  if (statusEl) statusEl.textContent = text;
  console.log("[STATUS]", text);
}

function setOutput(data) {
  const text =
    typeof data === "string" ? data : JSON.stringify(data, null, 2);

  if (outputEl) {
    outputEl.textContent = text;
  }

  console.log("[OUTPUT]");
  console.log(text);
}

async function getAccessToken(API) {
  const result = await API.extension.requestPermission("accesstoken");

  if (typeof result === "string" && result.trim()) {
    return result.trim();
  }

  if (
    result &&
    typeof result === "object" &&
    typeof result.data === "string" &&
    result.data.trim()
  ) {
    return result.data.trim();
  }

  return null;
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

async function fetchJson(url, token) {
  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json"
    }
  });

  const text = await res.text();

  let json = null;
  try {
    json = JSON.parse(text);
  } catch {
    json = null;
  }

  return {
    status: res.status,
    ok: res.ok,
    text,
    json
  };
}

function buildFileListCandidates(projectId) {
  const base = "https://app.connect.trimble.com/tc/api/2.0";

  return [
    `${base}/projects/${projectId}/files`,
    `${base}/project/${projectId}/files`,
    `${base}/files?projectId=${encodeURIComponent(projectId)}`,
    `${base}/projects/${projectId}/folders/root/files`,
    `${base}/projects/${projectId}/documents`,
    `${base}/projects/${projectId}/items`
  ];
}

async function validateToken(accessToken) {
  const url = "https://app.connect.trimble.com/tc/api/2.0/users/me";
  return fetchJson(url, accessToken);
}

async function listKofFilesDirect(project, accessToken) {
  const diagnostics = {
    tokenValidation: null,
    projectId: project.id,
    projectLocation: project.location,
    tried: []
  };

  setStatus("Validerer token...");

  const me = await validateToken(accessToken);
  diagnostics.tokenValidation = {
    url: "https://app.connect.trimble.com/tc/api/2.0/users/me",
    status: me.status,
    ok: me.ok,
    preview: me.json || me.text.slice(0, 300)
  };

  if (!me.ok) {
    return {
      ok: false,
      error: "Token-validering feilet",
      diagnostics
    };
  }

  setStatus("Henter filliste...");

  const urls = buildFileListCandidates(project.id);

  for (const url of urls) {
    try {
      const result = await fetchJson(url, accessToken);

      diagnostics.tried.push({
        url,
        status: result.status,
        ok: result.ok,
        preview: result.json || result.text.slice(0, 300)
      });

      if (!result.ok) {
        continue;
      }

      const list = flattenPossibleFileList(result.json);
      const files = list.map(mapFile);
      const kofFiles = files.filter(isKofFile);

      return {
        ok: true,
        usedUrl: url,
        totalFilesSeen: files.length,
        kofFiles,
        diagnostics
      };
    } catch (err) {
      diagnostics.tried.push({
        url,
        ok: false,
        error: String(err?.message || err)
      });
    }
  }

  return {
    ok: false,
    error: "Fant ikke fungerende file-list endpoint ennå.",
    diagnostics
  };
}

(async function main() {
  try {
    setStatus("Kobler til Trimble Connect...");

    const API = await TrimbleConnectWorkspace.connect(
      window.parent,
      (event, args) => {
        console.log("WS EVENT:", event, args?.data || args);
      },
      30000
    );

    if (!API?.project?.getProject) {
      throw new Error("API.project.getProject finnes ikke.");
    }

    const project = await API.project.getProject();

    console.log("API keys:", Object.keys(API || {}));
    console.log("Project:", project);

    if (!project?.id) {
      throw new Error("Fant ikke prosjekt-ID.");
    }

    if (!API?.ui?.setMenu) {
      throw new Error("API.ui.setMenu finnes ikke.");
    }

    await API.ui.setMenu({
      title: "KOF2TXT",
      command: "kof2txt_main",
      subMenus: [
        {
          title: "Finn KOF-filer",
          command: "kof2txt_list"
        }
      ]
    });

    if (API.ui.setActiveMenuItem) {
      await API.ui.setActiveMenuItem("kof2txt_list");
    }

    if (API.extension?.setStatusMessage) {
      await API.extension.setStatusMessage("KOF2TXT klar");
    }

    setStatus("Ber om access token...");
    const accessToken = await getAccessToken(API);

    if (!accessToken) {
      setStatus("Fikk ikke access token");
      setOutput({
        ok: false,
        step: "token",
        project
      });
      return;
    }

    console.log("Access token mottatt:", true);

    const result = await listKofFilesDirect(project, accessToken);

    if (!result.ok) {
      setStatus("Fant ikke .kof-filer");
      setOutput({
        ok: false,
        step: "listKofFilesDirect",
        project,
        result
      });
      return;
    }

    setStatus(
      result.kofFiles.length > 0
        ? `Fant ${result.kofFiles.length} .kof-fil(er)`
        : "Fant ingen .kof-filer"
    );

    setOutput({
      ok: true,
      step: "listKofFilesDirect",
      project,
      usedUrl: result.usedUrl,
      totalFilesSeen: result.totalFilesSeen,
      kofFiles: result.kofFiles.map((f) => ({
        id: f.id,
        name: f.name,
        path: f.path,
        size: f.size
      })),
      diagnostics: result.diagnostics
    });
  } catch (err) {
    console.error("Fatal error:", err);
    setStatus("Feil");
    setOutput({
      ok: false,
      error: String(err?.message || err),
      stack: err?.stack || null
    });
  }
})();
