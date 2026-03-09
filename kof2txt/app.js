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

function normalizeCommand(cmd) {
  if (!cmd) return "";
  if (typeof cmd === "string") return cmd;
  if (typeof cmd === "object") {
    return cmd.command || cmd.id || cmd.title || "";
  }
  return "";
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

async function callProxy(payload) {
  const res = await fetch("/.netlify/functions/tc-proxy", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  const text = await res.text();

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    parsed = text;
  }

  return {
    status: res.status,
    ok: res.ok,
    body: parsed
  };
}

function extractKofSummary(proxyBody) {
  const kofFiles = Array.isArray(proxyBody?.kofFiles) ? proxyBody.kofFiles : [];

  return {
    count: kofFiles.length,
    files: kofFiles.map((f) => ({
      id: f.id || null,
      name: f.name || null,
      path: f.path || f.fullPath || null,
      size: f.size ?? null
    }))
  };
}

async function listKofFiles(project, accessToken) {
  setStatus("Henter .kof-filer via proxy...");

  const result = await callProxy({
    action: "listKofFiles",
    token: accessToken,
    projectId: project.id,
    projectLocation: project.location
  });

  return result;
}

(async function main() {
  try {
    setStatus("Kobler til Trimble Connect...");

    let latestCommand = "";

    const API = await TrimbleConnectWorkspace.connect(
      window.parent,
      (event, args) => {
        console.log("WS EVENT:", event, args?.data || args);

        if (event === "extension.command") {
          latestCommand = normalizeCommand(args?.data);
          console.log("Menykommando:", latestCommand);
        }
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
        message: "Trimble Connect returnerte ikke access token.",
        project
      });
      return;
    }

    console.log("Access token mottatt:", true);

    const proxyResult = await listKofFiles(project, accessToken);

    if (!proxyResult.ok) {
      setStatus(`Proxy-feil (${proxyResult.status})`);
      setOutput({
        ok: false,
        step: "listKofFiles",
        project,
        proxyStatus: proxyResult.status,
        proxyResponse: proxyResult.body
      });
      return;
    }

    const summary = extractKofSummary(proxyResult.body);

    setStatus(
      summary.count > 0
        ? `Fant ${summary.count} .kof-fil(er)`
        : "Fant ingen .kof-filer"
    );

    setOutput({
      ok: true,
      step: "listKofFiles",
      project,
      summary,
      diagnostics: proxyResult.body?.diagnostics || null
    });

    // Dersom bruker trykker meny på nytt senere, kan vi bruke dette senere.
    console.log("Siste kommando:", latestCommand);
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

