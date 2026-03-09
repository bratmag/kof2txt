/* global TrimbleConnectWorkspace */

const statusEl = document.getElementById("status");
const outputEl = document.getElementById("output");

function setStatus(text) {
  if (statusEl) {
    statusEl.textContent = text;
  }
  console.log("[STATUS]", text);
}

function setOutput(data) {
  if (outputEl) {
    outputEl.textContent =
      typeof data === "string" ? data : JSON.stringify(data, null, 2);
  }
  console.log("[OUTPUT]", data);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function getAccessToken(API) {
  let accessToken = null;

  const tokenPromise = new Promise((resolve) => {
    const timeout = setTimeout(() => {
      resolve(null);
    }, 5000);

    // Midlertidig event-hook via connect-callbacken under
    window.__kof2txtResolveToken = (token) => {
      clearTimeout(timeout);
      resolve(token || null);
    };
  });

  try {
    const permissionResult = await API.extension.requestPermission("accesstoken");
    console.log("Permission result:", permissionResult);

    if (typeof permissionResult === "string" && permissionResult.trim()) {
      accessToken = permissionResult.trim();
      return accessToken;
    }

    if (
      permissionResult &&
      typeof permissionResult === "object" &&
      typeof permissionResult.data === "string" &&
      permissionResult.data.trim()
    ) {
      accessToken = permissionResult.data.trim();
      return accessToken;
    }

    // Vent litt på event dersom requestPermission returnerer pending el.l.
    accessToken = await tokenPromise;
    return accessToken;
  } finally {
    delete window.__kof2txtResolveToken;
  }
}

async function testProxy(projectId, accessToken) {
  const url = `https://api.connect.trimble.com/connect/project/${projectId}`;

  const res = await fetch("/.netlify/functions/tc-proxy", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      url,
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
        console.log("WS EVENT:", event, args);

        if (event === "extension.accessToken") {
          const token = args?.data;
          console.log("Access token event mottatt:", !!token);

          if (window.__kof2txtResolveToken) {
            window.__kof2txtResolveToken(token);
          }
        }

        if (event === "extension.command") {
          console.log("Menykommando:", args?.data);
        }
      },
      30000
    );

    console.log("API object:", API);
    console.log("API keys:", Object.keys(API || {}));
    console.log("API.ui:", API?.ui);
    console.log("API.project:", API?.project);
    console.log("API.extension:", API?.extension);

    if (!API) {
      throw new Error("Fikk ikke API-objekt fra Trimble Connect.");
    }

    if (!API.project?.getProject) {
      throw new Error("API.project.getProject finnes ikke.");
    }

    const project = await API.project.getProject();
    console.log("Project:", project);

    if (!project?.id) {
      throw new Error("Fant ikke gyldig prosjekt-ID.");
    }

    if (!API.ui?.setMenu) {
      throw new Error("API.ui.setMenu finnes ikke.");
    }

    setStatus("Setter meny...");

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
        message: "Trimble Connect ga ikke access token.",
        project
      });
      return;
    }

    console.log("Access token mottatt:", true);

    setStatus("Tester proxy...");

    const proxyResult = await testProxy(project.id, accessToken);

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

    setStatus("Alt virker så langt");
    setOutput({
      ok: true,
      step: "proxy-test-ok",
      message: "Extension, token og proxy fungerer.",
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
