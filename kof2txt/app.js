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
    name,
    path,
    type,
    size: raw?.size ?? raw?.fileSize ?? null,
    raw
  };
}

function isKofFile(entry) {
  const candidates = [entry?.name, entry?.path].filter(Boolean);
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

function buildRootCandidates(projectId, rootId) {
  const base = "https://app.connect.trimble.com/tc/api/2.0";

  return [
    `${base}/projects/${projectId}/folders/${rootId}/contents`,
    `${base}/projects/${projectId}/folders/${rootId}/items`,
    `${base}/projects/${projectId}/folders/${rootId}/children`,
    `${base}/projects/${projectId}/folders/${rootId}`,
    `${base}/folders/${rootId}/contents`,
    `${base}/folders/${rootId}/items`,
    `${base}/folders/${rootId}/children`,
    `${base}/folders/${rootId}`,
    `${base}/projects/${projectId}/root`,
    `${base}/projects/${projectId}`
  ];
}

async function listFromRoot(project, matchedProject, accessToken) {
  const rootId = matchedProject?.rootId;

  if (!rootId) {
    return {
      ok: false,
      error: "Matched project mangler rootId"
    };
  }

  const diagnostics = {
    rootId,
    tried: []
  };

  setStatus("Tester root-folder endepunkter...");

  const urls = buildRootCandidates(project.id, rootId);

  for (const url of urls) {
    try {
      const result = await fetchJson(url, accessToken);

      diagnostics.tried.push({
        url,
        status: result.status,
        ok: result.ok,
        preview: result.json || result.text.slice(0, 400)
      });

      if (!result.ok) {
        continue;
      }

      const list = flattenPossibleList(result.json);
      const entries = list.map(mapEntry);
      const kofFiles = entries.filter(isKofFile);

      return {
        ok: true,
        usedUrl: url,
        rootId,
        entryCount: entries.length,
        entries,
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
    error: "Fant ikke fungerende root-endepunkt ennå.",
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

    const rootTest = await listFromRoot(
      project,
      projectCheck.matchedProject,
      accessToken
    );

    if (!rootTest.ok) {
      setStatus("Fant ikke fungerende root-endepunkt");
      setOutput({
        ok: false,
        step: "rootEndpointTest",
        matchedProject: projectCheck.matchedProject,
        result: rootTest,
        diagnostics: projectCheck.diagnostics
      });
      return;
    }

    setStatus(
      rootTest.kofFiles.length > 0
        ? `Fant ${rootTest.kofFiles.length} .kof-fil(er)`
        : `Root-endepunkt virker, men fant ingen .kof-filer`
    );

    setOutput({
      ok: true,
      step: "rootEndpointTest",
      matchedProject: projectCheck.matchedProject,
      usedUrl: rootTest.usedUrl,
      rootId: rootTest.rootId,
      entryCount: rootTest.entryCount,
      kofFiles: rootTest.kofFiles.map((f) => ({
        id: f.id,
        name: f.name,
        path: f.path,
        type: f.type,
        size: f.size
      })),
      sampleEntries: rootTest.entries.slice(0, 20).map((f) => ({
        id: f.id,
        name: f.name,
        path: f.path,
        type: f.type,
        size: f.size
      })),
      diagnostics: rootTest.diagnostics
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
