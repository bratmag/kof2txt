/* global TrimbleConnectWorkspace */

const statusEl = document.getElementById("status");
const outputEl = document.getElementById("output");

function setStatus(text) {
  if (statusEl) statusEl.textContent = text;
}

function setOutput(obj) {
  if (!outputEl) return;
  outputEl.textContent =
    typeof obj === "string" ? obj : JSON.stringify(obj, null, 2);
}

(async function main() {
  try {
    setStatus("Kobler til Trimble Connect...");

    const API = await TrimbleConnectWorkspace.connect(
      window.parent,
      (event, args) => {
        console.log("WS EVENT:", event, args);
      },
      30000
    );

    console.log("API object:", API);
    console.log("API keys:", Object.keys(API || {}));
    console.log("API.ui:", API?.ui);
    console.log("API.project:", API?.project);
    console.log("API.extension:", API?.extension);

    // Vis dette i UI også
    setOutput({
      connected: true,
      apiKeys: Object.keys(API || {}),
      hasUi: !!API?.ui,
      hasProject: !!API?.project,
      hasExtension: !!API?.extension
    });

    // Test prosjekt først, før ui.setMenu
    let project = null;
    if (API?.project?.getProject) {
      project = await API.project.getProject();
      console.log("Project:", project);
    }

    if (!API?.ui?.setMenu) {
      setStatus("Koblet, men API.ui.setMenu finnes ikke");
      setOutput({
        connected: true,
        project,
        apiKeys: Object.keys(API || {}),
        hasUi: !!API?.ui,
        uiKeys: API?.ui ? Object.keys(API.ui) : []
      });
      return;
    }

    setStatus("Koblet. Setter meny...");

    await API.ui.setMenu({
      title: "KOF DEBUG",
      command: "kof_debug_main",
      subMenus: [
        {
          title: "Åpne KOF DEBUG",
          command: "kof_debug_open"
        }
      ]
    });

    await API.ui.setActiveMenuItem("kof_debug_open");

    setStatus("Meny satt OK");
    setOutput({
      message: "KOF DEBUG er lastet",
      project
    });
  } catch (err) {
    console.error(err);
    setStatus("Feil");
    setOutput({
      error: String(err),
      stack: err?.stack || null
    });
  }
})();
