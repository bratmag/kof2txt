(() => {
  "use strict";

  const CONFIG = {
    DEBUG: true,
    CONNECT_TIMEOUT_MS: 30000,
    AUTO_PROCESS_ON_FILE_SELECTED: true,
    PROXY_URL: "/.netlify/functions/tc-proxy"
  };

  const els = {
    status:
      document.getElementById("status") ||
      document.getElementById("statusText") ||
      null,
    output:
      document.getElementById("output") ||
      document.getElementById("result") ||
      null,
    runBtn:
      document.getElementById("runBtn") ||
      document.getElementById("startBtn") ||
      null
  };

  const state = {
    api: null,
    accessToken: null,
    project: null,
    selectedFile: null,
    lastResult: null
  };

  function log(...args) {
    console.log(...args);
  }

  function debug(...args) {
    if (CONFIG.DEBUG) console.log(...args);
  }

  function setStatus(message) {
    log(`[STATUS] ${message}`);
    if (els.status) els.status.textContent = message;

    if (state.api?.extension?.setStatusMessage) {
      state.api.extension.setStatusMessage(message).catch(() => {});
    }
  }

  function setOutput(data) {
    log("[OUTPUT]");
    log(data);
    if (els.output) {
      els.output.textContent =
        typeof data === "string" ? data : JSON.stringify(data, null, 2);
    }
  }

  function shortText(text, len = 1200) {
    if (typeof text !== "string") return text;
    return text.length > len ? text.slice(0, len) + "..." : text;
  }

  function isKofFile(file) {
    return /\.kof$/i.test(String(file?.name || ""));
  }

  function safeJsonParse(text) {
    try {
      return JSON.parse(text);
    } catch {
      return null;
    }
  }

  function withTimeout(promise, ms, label) {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        reject(new Error(`${label} timed out after ${ms} ms`));
      }, ms);

      promise
        .then((value) => {
          clearTimeout(timer);
          resolve(value);
        })
        .catch((err) => {
          clearTimeout(timer);
          reject(err);
        });
    });
  }

  function triggerDownload(filename, text) {
    const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function fileLikeFromEventArg(arg) {
    if (!arg) return null;

    if (arg.id && arg.name) return arg;
    if (arg.data?.id && arg.data?.name) return arg.data;
    if (arg.file?.id && arg.file?.name) return arg.file;

    if (Array.isArray(arg.files) && arg.files.length) return arg.files[0];
    if (Array.isArray(arg.data?.files) && arg.data.files.length) return arg.data.files[0];

    return null;
  }

  async function connectWorkspace() {
    setStatus("Kobler til Trimble Connect...");

    if (!window.TrimbleConnectWorkspace?.connect) {
      throw new Error(
        "TrimbleConnectWorkspace ikke funnet. Sjekk at Workspace API-scriptet er lastet i index.html."
      );
    }

    const api = await TrimbleConnectWorkspace.connect(
      window.parent,
      onWorkspaceEvent,
      CONFIG.CONNECT_TIMEOUT_MS
    );

    state.api = api;
    debug("API keys:", Object.keys(api || {}));

    if (api?.ui?.setMenu) {
      try {
        await api.ui.setMenu({
          title: "KOF2TXT",
          icon: "",
          command: "kof2txt"
        });
      } catch (err) {
        debug("setMenu ignorert:", err?.message || err);
      }
    }

    return api;
  }

  async function requestAccessToken() {
    setStatus("Ber om access token...");

    if (!state.api?.extension?.requestPermission) {
      throw new Error("extension.requestPermission finnes ikke.");
    }

    const token = await state.api.extension.requestPermission("accesstoken");

    if (!token || typeof token !== "string" || token === "pending" || token === "denied") {
      throw new Error(`Fikk ikke gyldig access token. Svar: ${String(token)}`);
    }

    state.accessToken = token;
    debug("Access token mottatt:", true);
    return token;
  }

  async function getProject() {
    setStatus("Henter prosjektinfo...");

    if (!state.api?.project?.getProject) {
      throw new Error("project.getProject finnes ikke.");
    }

    const project = await state.api.project.getProject();

    if (!project?.id) {
      throw new Error("Fant ikke aktivt prosjekt.");
    }

    state.project = project;
    debug("Project:", project);
    return project;
  }

  async function callProxy(action, payload) {
    const res = await withTimeout(
      fetch(CONFIG.PROXY_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          action,
          ...payload
        })
      }),
      60000,
      `Proxy ${action}`
    );

    const text = await res.text();
    const json = safeJsonParse(text);

    return {
      ok: res.ok,
      status: res.status,
      text,
      json
    };
  }

  function convertKofToTxt(kofText) {
    // Foreløpig råtekst.
    // Sett inn faktisk KOF->TXT-logikk her når nedlasting fungerer stabilt.
    return kofText;
  }

  async function processSelectedFile() {
    try {
      if (!state.api) {
        throw new Error("Ikke koblet til Workspace API.");
      }

      if (!state.accessToken) {
        await requestAccessToken();
      }

      if (!state.project) {
        await getProject();
      }

      if (!state.selectedFile) {
        setStatus("Velg en .kof-fil i Trimble Connect først.");
        setOutput({
          ok: false,
          step: "noSelectedFile",
          message: "Ingen fil valgt. Marker eller åpne en .kof-fil i Trimble Connect."
        });
        return;
      }

      if (!isKofFile(state.selectedFile)) {
        setStatus("Valgt fil er ikke en .kof-fil");
        setOutput({
          ok: false,
          step: "wrongFileType",
          file: state.selectedFile
        });
        return;
      }

      setStatus(`Laster ned ${state.selectedFile.name} ...`);

      const proxyRes = await callProxy("downloadKofFile", {
        token: state.accessToken,
        projectId: state.project.id,
        projectLocation: state.project.location,
        fileId: state.selectedFile.id,
        fileName: state.selectedFile.name
      });

      if (!proxyRes.ok || !proxyRes.json) {
        setStatus("Proxy-feil");
        setOutput({
          ok: false,
          step: "proxyHttp",
          status: proxyRes.status,
          preview: shortText(proxyRes.text, 1500)
        });
        return;
      }

      const result = proxyRes.json;
      state.lastResult = result;

      if (!result.ok) {
        setStatus("Klarte ikke laste ned KOF-fil");
        setOutput(result);
        return;
      }

      const txt = convertKofToTxt(result.text || "");
      const outName = String(result.file?.name || "output.kof").replace(/\.kof$/i, ".txt");

      setStatus("KOF-fil lastet ned");
      setOutput({
        ok: true,
        project: result.project,
        file: result.file,
        source: result.source,
        contentType: result.contentType,
        preview: shortText(result.text || "", 1500)
      });

      triggerDownload(outName, txt);
    } catch (err) {
      console.error(err);
      setStatus("Feil");
      setOutput({
        ok: false,
        error: err?.message || String(err)
      });
    }
  }

  async function runCoreProbe() {
    try {
      if (!state.accessToken) await requestAccessToken();
      if (!state.project) await getProject();

      if (!state.selectedFile?.id) {
        setOutput({
          ok: false,
          step: "probeNoFile",
          message: "Velg en .kof-fil først, så kan probe kjøre mot valgt fil."
        });
        return;
      }

      setStatus("Kjører Core probe...");

      const proxyRes = await callProxy("probeCore", {
        token: state.accessToken,
        projectId: state.project.id,
        projectLocation: state.project.location,
        fileId: state.selectedFile.id,
        fileName: state.selectedFile.name
      });

      setStatus("Core probe ferdig");
      setOutput(proxyRes.json || proxyRes.text);
      return proxyRes.json || proxyRes.text;
    } catch (err) {
      console.error(err);
      setStatus("Feil i probe");
      setOutput({
        ok: false,
        error: err?.message || String(err)
      });
    }
  }

  function onWorkspaceEvent(event, args) {
    debug("[TC EVENT]", event, args);

    if (event === "extension.fileSelected" || event === "extension.fileViewClicked") {
      const file = fileLikeFromEventArg(args);

      if (!file) {
        debug("Fikk file-event, men klarte ikke lese filobjektet.");
        return;
      }

      state.selectedFile = file;

      setStatus(`Valgt fil: ${file.name}`);
      setOutput({
        ok: true,
        message: "Fil valgt i Trimble Connect",
        file: {
          id: file.id,
          name: file.name,
          type: file.type,
          versionId: file.versionId
        }
      });

      if (CONFIG.AUTO_PROCESS_ON_FILE_SELECTED && isKofFile(file)) {
        processSelectedFile().catch((err) => {
          console.error(err);
          setStatus("Feil");
          setOutput({
            ok: false,
            error: err?.message || String(err)
          });
        });
      }
    }
  }

  async function init() {
    try {
      setStatus("Starter...");
      await connectWorkspace();
      await requestAccessToken();
      await getProject();

      setStatus("Klar. Velg en .kof-fil i Trimble Connect.");
      setOutput({
        ok: true,
        project: {
          id: state.project.id,
          name: state.project.name,
          location: state.project.location
        },
        message: "Extension er klar. Velg eller åpne en .kof-fil i Trimble Connect."
      });

      window.kof2txt = {
        state,
        processSelectedFile,
        runCoreProbe
      };
    } catch (err) {
      console.error(err);
      setStatus("Feil");
      setOutput({
        ok: false,
        error: err?.message || String(err)
      });
    }
  }

  if (els.runBtn) {
    els.runBtn.addEventListener("click", () => {
      processSelectedFile();
    });
  }

  window.addEventListener("load", init);
})();
