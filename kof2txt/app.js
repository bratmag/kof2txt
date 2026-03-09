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

async function validateAndListProjects(project, accessToken) {
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
          title: "Test Core API",
          command: "kof2txt_coretest"
        }
      ]
    });

    if (API.ui.setActiveMenuItem) {
      await API.ui.setActiveMenuItem("kof2txt_coretest");
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

    const result = await validateAndListProjects(project, accessToken);

    if (!result.ok) {
      setStatus("Core API-test feilet");
      setOutput({
        ok: false,
        step: "coreApiTest",
        result
      });
      return;
    }

    setStatus("Core API-test OK");
    setOutput({
      ok: true,
      step: "coreApiTest",
      projectsCount: result.projectsCount,
      matchedProject: result.matchedProject,
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
