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

async function fetchText(url, token) {
  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "*/*"
    }
  });

  const text = await res.text();

  return {
    status: res.status,
    ok: res.ok,
    text
  };
}

function isObject(value) {
  return value && typeof value === "object" && !Array.isArray(value);
}

function flattenPossibleList(payload) {
  if (Array.isArray(payload)) return payload;
  if (!isObject(payload)) return [];

  const keysToTry = [
    "items",
    "files",
    "folders",
    "entries",
    "children",
    "documents",
    "results",
    "data"
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

function mapEntry(raw) {
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

  const type =
    raw?.type ||
    raw?.entryType ||
    raw?.kind ||
    raw?.objectType ||
    null;

  return {
    id: raw?.id || raw?.fileId || raw?.folderId || raw?.identifier || null,
    versionId: raw?.versionId || raw?.id || null,
    name,
    path,
    type,
    size: raw?.size ?? raw?.fileSize ?? null,
    raw
  };
}

function isKofFile(entry) {
  const candidates = [entry?.name].filter(Boolean);
  return candidates.some((v) => String(v).toLowerCase().endsWith(".kof"));
}

async function validateAndFindProject(project, accessToken) {
  const diagnostics = {
    currentProjectFromWorkspace: project,
    tokenValidation: null,
    projectsCall: null
  };

  setStatus("Validerer token...");

  const meUrl = "https://app.connect.trimble.com/tc/api/2.0/users/me";
  const me = await fetchJson(meUrl, accessToken);

  diagnostics.tokenValidation = {
    url: meUrl,
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

  setStatus("Henter prosjektliste...");

  const projectsUrl =
    "https://app.connect.trimble.com/tc/api/2.0/projects?fullyLoaded=false";

  const projectsRes = await fetchJson(projectsUrl, accessToken);

  diagnostics.projectsCall = {
    url: projectsUrl,
    status: projectsRes.status,
    ok: projectsRes.ok,
    preview: projectsRes.json || projectsRes.text.slice(0, 500)
  };

  if (!projectsRes.ok) {
    return {
      ok: false,
      error: "Prosjektliste-kall feilet",
      diagnostics
    };
  }

  const raw = projectsRes.json;
  const projects = Array.isArray(raw)
    ? raw
    : Array.isArray(raw?.items)
    ? raw.items
    : Array.isArray(raw?.data)
    ? raw.data
    : Array.isArray(raw?.projects)
    ? raw.projects
    : [];

  const matchedProject =
    projects.find((p) => p.id === project.id) ||
    projects.find((p) => p.projectId === project.id) ||
    null;

  return {
    ok: true,
    projectsCount: projects.length,
    matchedProject,
    diagnostics
  };
}

async function listRootItems(projectId, rootId, accessToken) {
  const url = `https://app.connect.trimble.com/tc/api/2.0/folders/${rootId}/items`;
  const result = await fetchJson(url, accessToken);

  if (!result.ok) {
    return {
      ok: false,
      url,
      status: result.status,
      preview: result.json || result.text.slice(0, 400)
    };
  }

  const entries = flattenPossibleList(result.json).map(mapEntry);

  return {
    ok: true,
    url,
    entries
  };
}

async function downloadKofFile(fileId, accessToken) {
  const candidates = [
    `https://app.connect.trimble.com/tc/api/2.0/files/${fileId}`,
    `https://app.connect.trimble.com/tc/api/2.0/files/${fileId}/content`,
    `https://app.connect.trimble.com/tc/api/2.0/files/${fileId}/download`,
    `https://app.connect.trimble.com/tc/api/2.0/data/${fileId}`
  ];

  const diagnostics = [];

  for (const url of candidates) {
    try {
      const result = await fetchText(url, accessToken);

      diagnostics.push({
        url,
        status: result.status,
        ok: result.ok,
        preview: result.text.slice(0, 200)
      });

      if (!result.ok) {
        continue;
      }

      if (result.text && result.text.trim().length > 0) {
        return {
          ok: true,
          usedUrl: url,
          text: result.text,
          diagnostics
        };
      }
    } catch (err) {
      diagnostics.push({
        url,
        ok: false,
        error: String(err?.message || err)
      });
    }
  }

  return {
    ok: false,
    error: "Fant ikke fungerende nedlastings-endepunkt ennå.",
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
          title: "Last KOF-fil",
          command: "kof2txt_download"
        }
      ]
    });

    if (API.ui.setActiveMenuItem) {
      await API.ui.setActiveMenuItem("kof2txt_download");
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

    const projectCheck = await validateAndFindProject(project, accessToken);

    if (!projectCheck.ok) {
      setStatus("Core API-test feilet");
      setOutput({
        ok: false,
        step: "coreApiTest",
        result: projectCheck
      });
      return;
    }

    if (!projectCheck.matchedProject) {
      setStatus("Fant ikke prosjektet i Core API");
      setOutput({
        ok: false,
        step: "coreApiMatch",
        result: projectCheck
      });
      return;
    }

    const rootId = projectCheck.matchedProject.rootId;

    if (!rootId) {
      setStatus("Fant ikke rootId");
      setOutput({
        ok: false,
        step: "rootId",
        matchedProject: projectCheck.matchedProject
      });
      return;
    }

    setStatus("Lister filer i root...");
    const rootList = await listRootItems(project.id, rootId, accessToken);

    if (!rootList.ok) {
      setStatus("Klarte ikke liste root-filer");
      setOutput({
        ok: false,
        step: "listRootItems",
        rootList
      });
      return;
    }

    const kofFiles = rootList.entries.filter(isKofFile);

    if (kofFiles.length === 0) {
      setStatus("Fant ingen .kof-filer");
      setOutput({
        ok: false,
        step: "findKof",
        matchedProject: projectCheck.matchedProject,
        rootUrl: rootList.url,
        entries: rootList.entries.map((f) => ({
          id: f.id,
          name: f.name,
          type: f.type,
          size: f.size
        }))
      });
      return;
    }

    const firstKof = kofFiles[0];

    setStatus(`Laster ned ${firstKof.name} ...`);
    const download = await downloadKofFile(firstKof.id, accessToken);

    if (!download.ok) {
      setStatus("Klarte ikke laste ned KOF-fil");
      setOutput({
        ok: false,
        step: "downloadKof",
        file: {
          id: firstKof.id,
          name: firstKof.name
        },
        download
      });
      return;
    }

    setStatus("KOF-fil lastet ned");

    setOutput({
      ok: true,
      step: "downloadKof",
      matchedProject: projectCheck.matchedProject,
      rootUrl: rootList.url,
      file: {
        id: firstKof.id,
        name: firstKof.name,
        size: firstKof.size
      },
      downloadUrl: download.usedUrl,
      preview: download.text.slice(0, 2000),
      diagnostics: download.diagnostics
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
