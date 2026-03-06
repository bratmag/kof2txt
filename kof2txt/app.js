/* global TrimbleConnectWorkspace */

const statusEl = document.getElementById("status");
const outputEl = document.getElementById("output");

function setStatus(text) {
  if (statusEl) statusEl.textContent = text;
}

function setOutput(obj) {
  if (outputEl) {
    if (typeof obj === "string") {
      outputEl.textContent = obj;
    } else {
      outputEl.textContent = JSON.stringify(obj, null, 2);
    }
  }
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

    const project = await API.project.getProject();

    setStatus("Meny satt OK");
    setOutput({
      message: "KOF DEBUG er lastet",
      project
    });

  } catch (err) {
    console.error(err);
    setStatus("Feil");
    setOutput(String(err));
  }
})();
