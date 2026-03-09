/* global TrimbleConnectWorkspace */

const statusEl = document.getElementById("status");
const outputEl = document.getElementById("output");

function setStatus(text) {
  if (statusEl) statusEl.textContent = text;
  console.log("[STATUS]", text);
}

function setOutput(data) {
  if (outputEl) {
    outputEl.textContent =
      typeof data === "string" ? data : JSON.stringify(data, null, 2);
  }
  console.log("[OUTPUT]", data);
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

async function testProxy(projectId, accessToken, projectLocation) {
  const res = await fetch("/.netlify/functions/tc-proxy", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      action: "debugProjectAccess",
      projectId,
      projectLocation,
      token: accessToken
    })
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

    const project = await API.project.getProject();

    console.log("API keys:", Object.keys(API || {}));
    console.log("Project:", project);

    await API.ui.setMenu({
      title: "KOF2TXT",
      command: "kof2txt_main",
      subMenus: [
        {
          title: "Konverter KOF",
          command: "kof2txt_convert"
        }
      ]
    });

    if (API.ui.setActiveMenuItem) {
      await API.ui.setActiveMenuItem("kof2txt_convert");
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

    setStatus("Tester proxy...");

    const proxyResult = await testProxy(
      project.id,
      accessToken,
      project.location
    );

    if (!proxyResult.ok) {
      setStatus(`Proxy-feil (${proxyResult.status})`);
      setOutput({
        ok: false,
        step: "proxy",
        project,
        proxyStatus: proxyResult.status,
        proxyResponse: proxyResult.body
      });
      return;
    }

    setStatus("Proxy-test OK");
    setOutput({
      ok: true,
      step: "proxy-test-ok",
      project,
      proxyStatus: proxyResult.status,
      proxyResponse: proxyResult.body
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
