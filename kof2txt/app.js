/* app.js - KOF2TXT Trimble Connect Extension */

(() => {
  "use strict";

  const CONFIG = {
    DEBUG: true,
    AUTO_RUN: true,
    ROOT_SEARCH_RECURSIVE: false,
    CONNECT_TIMEOUT_MS: 30000,
    TOKEN_WAIT_MS: 30000,
    FETCH_TIMEOUT_MS: 30000,
    PRESIGNED_FETCH_TIMEOUT_MS: 60000,
    PROXY_URL: "/.netlify/functions/tc-proxy"
  };

  // =========================
  // DOM helpers
  // =========================
  const $status =
    document.getElementById("status") ||
    document.getElementById("statusText") ||
    document.getElementById("log") ||
    null;

  const $output =
    document.getElementById("output") ||
    document.getElementById("result") ||
    document.getElementById("json") ||
    null;

  const $runBtn =
    document.getElementById("runBtn") ||
    document.getElementById("startBtn") ||
    document.getElementById("btnRun") ||
    null;

  // =========================
  // State
  // =========================
  const state = {
    api: null,
    accessToken: null,
    tokenWaiters: [],
    project: null
  };

  // =========================
  // Logging / UI
  // =========================
  function log(...args) { console.log(...args); }
  function debug(...args) { if (CONFIG.DEBUG) console.log(...args); }

  function setStatus(msg) {
    console.log(`[STATUS] ${msg}`);
    if ($status) $status.textContent = msg;
    if (state.api?.extension?.setStatusMessage) {
      state.api.extension.setStatusMessage(msg).catch(() => {});
    }
  }

  function setOutput(obj) {
    console.log("[OUTPUT]");
    console.log(obj);
    if ($output) {
      $output.textContent =
        typeof obj === "string" ? obj : JSON.stringify(obj, null, 2);
    }
  }

  function shortText(text, max = 300) {
    if (typeof text !== "string") return text;
    return text.length > max ? text.slice(0, max) + "..." : text;
  }

  function safeJsonParse(text) {
    try { return JSON.parse(text); } catch { return null; }
  }

  function delay(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function withTimeout(promise, ms, label = "Operation") {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        reject(new Error(`${label} timed out after ${ms} ms`));
      }, ms);
      promise
        .then((value) => { clearTimeout(timer); resolve(value); })
        .catch((err) => { clearTimeout(timer); reject(err); });
    });
  }

  // =========================
  // Workspace API
  // =========================
  async function connectWorkspace() {
    setStatus("Kobler til Trimble Connect...");

    if (!window.TrimbleConnectWorkspace?.connect) {
      throw new Error(
        "TrimbleConnectWorkspace ikke funnet. Sjekk at index.js fra Workspace API er lastet."
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
        await api.ui.setMenu({ title: "KOF2TXT", icon: "", command: "kof2txt_main" });
      } catch (err) {
        debug("setMenu ble ignorert:", err?.message || err);
      }
    }

    return api;
  }

  function onWorkspaceEvent(event, args) {
    debug("[TC EVENT]", event, args);
    if (event === "extension.accessToken") {
      const token = args?.data;
      if (typeof token === "string" && token && token !== "pending" && token !== "denied") {
        state.accessToken = token;
        const waiters = [...state.tokenWaiters];
        state.tokenWaiters.length = 0;
        waiters.forEach((resolve) => resolve(token));
      }
    }
  }

  async function requestAccessToken() {
    setStatus("Ber om access token...");

    if (!state.api?.extension?.requestPermission) {
      throw new Error("extension.requestPermission finnes ikke på API-et.");
    }

    const result = await state.api.extension.requestPermission("accesstoken");
    debug("requestPermission result:", result);

    if (typeof result === "string" && result && result !== "pending" && result !== "denied") {
      state.accessToken = result;
      return result;
    }
    if (result === "denied") {
      throw new Error("Tilgang til access token ble avslått i Trimble Connect.");
    }
    if (result === "pending") {
      const token = await withTimeout(
        new Promise((resolve) => { state.tokenWaiters.push(resolve); }),
        CONFIG.TOKEN_WAIT_MS,
        "Venter på extension.accessToken"
      );
      return token;
    }
    if (state.accessToken) return state.accessToken;

    throw new Error(`Uventet svar fra requestPermission: ${String(result)}`);
  }

  async function getProject() {
    setStatus("Henter prosjektinfo...");

    let project = null;
    if (state.api?.project?.getProject) {
      project = await state.api.project.getProject();
    } else if (state.api?.project?.getCurrentProject) {
      project = await state.api.project.getCurrentProject();
    }

    if (!project?.id) throw new Error("Fant ikke aktivt prosjekt.");

    state.project = project;
    debug("Project:", project);
    return project;
  }

  // =========================
  // Region / base URL
  // =========================
  function getCoreBaseUrl(projectLocation) {
    const loc = String(projectLocation || "").toLowerCase();
    if (loc === "europe") return "https://app.eu.connect.trimble.com/tc/api/2.0";
    if (loc === "asia") return "https://app.asia.connect.trimble.com/tc/api/2.0";
    return "https://app.connect.trimble.com/tc/api/2.0";
  }

  // =========================
  // Fetch helpers
  // =========================
  async function fetchRaw(url, options = {}, timeoutMs = CONFIG.FETCH_TIMEOUT_MS) {
    const res = await withTimeout(fetch(url, options), timeoutMs, `Fetch ${url}`);
    const contentType = res.headers.get("content-type") || "";
    const text = await res.text();
    const json = safeJsonParse(text);
    return { url, ok: res.ok, status: res.status, contentType, text, json };
  }

  async function fetchJsonWithBearer(url, token, timeoutMs = CONFIG.FETCH_TIMEOUT_MS) {
    return fetchRaw(url, { method: "GET", headers: { Authorization: `Bearer ${token}` } }, timeoutMs);
  }

  async function postJson(url, body, timeoutMs = CONFIG.FETCH_TIMEOUT_MS) {
    return fetchRaw(
      url,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body)
      },
      timeoutMs
    );
  }

  // =========================
  // File listing
  // =========================
  async function listRootFiles(accessToken, project) {
    setStatus("Lister filer i root...");

    const base = getCoreBaseUrl(project.location);
    const candidates = [
      `${base}/files?projectId=${encodeURIComponent(project.id)}&parentId=root`,
      `${base}/files?parentId=root&projectId=${encodeURIComponent(project.id)}`,
      `${base}/projects/${encodeURIComponent(project.id)}/files?parentId=root`,
      `${base}/projects/${encodeURIComponent(project.id)}/files`,
      `${base}/files?projectId=${encodeURIComponent(project.id)}`
    ];

    const diagnostics = [];

    for (const url of candidates) {
      try {
        const r = await fetchJsonWithBearer(url, accessToken);
        diagnostics.push({ url, ok: r.ok, status: r.status, preview: shortText(r.text, 250) });

        if (!r.ok) continue;

        const items = extractArrayFromPayload(r.json);
        if (Array.isArray(items) && items.length) {
          return { ok: true, url, files: items, diagnostics };
        }
      } catch (err) {
        diagnostics.push({ url, ok: false, error: err?.message || String(err) });
      }
    }

    return { ok: false, error: "Fant ingen fungerende root-listing endpoint.", diagnostics };
  }

  function extractArrayFromPayload(payload) {
    if (!payload) return [];
    if (Array.isArray(payload)) return payload;
    for (const key of ["items", "data", "files", "results", "children"]) {
      if (Array.isArray(payload[key])) return payload[key];
    }
    return [];
  }

  function findFirstKof(files) {
    if (!Array.isArray(files)) return null;
    return (
      files.find((f) => /\.kof$/i.test(String(f?.name || ""))) ||
      files.find((f) => String(f?.name || "").toLowerCase().includes(".kof")) ||
      null
    );
  }

  // =========================
  // Metadata
  // =========================
  async function getFileMetadata(accessToken, project, fileId) {
    const base = getCoreBaseUrl(project.location);
    const url = `${base}/files/${encodeURIComponent(fileId)}`;
    const r = await fetchJsonWithBearer(url, accessToken);
    return {
      ok: r.ok,
      status: r.status,
      url,
      contentType: r.contentType,
      preview: shortText(r.text, 300),
      data: r.json
    };
  }

  // =========================
  // Download via proxy
  // =========================
  async function downloadKofFile(accessToken, project, fileId) {
    // Hent metadata for å vise navn/versionId i output
    const meta = await getFileMetadata(accessToken, project, fileId);

    if (!meta.ok || !meta.data) {
      return { ok: false, step: "metadata", metadata: meta };
    }

    setStatus(`Laster ned ${meta.data.name || fileId} via proxy...`);

    // Alt nedlastingsarbeid skjer i Netlify-funksjonen (ingen CORS)
    const proxyRes = await postJson(
      CONFIG.PROXY_URL,
      {
        action: "downloadKofFile",
        token: accessToken,
        fileId,
        projectLocation: project.location
      },
      CONFIG.PRESIGNED_FETCH_TIMEOUT_MS
    );

    const result = proxyRes.json;

    if (!proxyRes.ok || !result?.ok) {
      return {
        ok: false,
        step: "downloadKof",
        metadata: { id: meta.data?.id, versionId: meta.data?.versionId, name: meta.data?.name },
        proxyStatus: proxyRes.status,
        proxyResult: result
      };
    }

    return {
      ok: true,
      step: "done",
      file: {
        id: meta.data?.id,
        versionId: meta.data?.versionId,
        name: meta.data?.name
      },
      source: result.source,
      contentType: null,
      text: result.content,
      diagnostics: result.diagnostics
    };
  }

  function safeHost(url) {
    try { return new URL(url).host; } catch { return null; }
  }

  // =========================
  // KOF -> TXT
  // =========================
  function convertKofToTxt(kofText) {
    // Placeholder: returner råtekst.
    // Bytt ut med faktisk KOF-parser her når nedlastingen fungerer.
    return kofText;
  }

  function downloadTextFile(filename, text) {
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

  function outputSuccess(result) {
    const txt = convertKofToTxt(result.text);

    setOutput({
      ok: true,
      project: { id: state.project?.id, location: state.project?.location, name: state.project?.name },
      file: result.file,
      source: result.source,
      preview: shortText(result.text, 1000)
    });

    const baseName = String(result.file?.name || "output.kof").replace(/\.kof$/i, "");
    downloadTextFile(`${baseName}.txt`, txt);
  }

  function outputError(errObj) {
    setOutput(errObj);
  }

  // =========================
  // Main
  // =========================
  async function main() {
    try {
      setStatus("Starter...");
      await connectWorkspace();

      setStatus("Validerer token...");
      const accessToken = await requestAccessToken();
      debug("Access token mottatt:", !!accessToken);

      const project = await getProject();

      const list = await listRootFiles(accessToken, project);
      if (!list.ok) {
        throw { ok: false, step: "listRootFiles", project, list };
      }

      const kofFile = findFirstKof(list.files);
      if (!kofFile) {
        throw {
          ok: false,
          step: "findKof",
          project,
          rootListUrl: list.url,
          rootFileCount: list.files.length,
          filesPreview: list.files.slice(0, 20).map((f) => ({
            id: f.id, name: f.name, type: f.type, versionId: f.versionId
          }))
        };
      }

      setStatus(`Laster ned ${kofFile.name} ...`);
      const result = await downloadKofFile(accessToken, project, kofFile.id);

      if (!result.ok) {
        setStatus("Klarte ikke laste ned KOF-fil");
        outputError({
          ok: false,
          step: "downloadKof",
          project: { id: project.id, location: project.location, name: project.name },
          file: { id: kofFile.id, name: kofFile.name, versionId: kofFile.versionId },
          download: result
        });
        return;
      }

      setStatus("KOF-fil lastet ned ✓");
      outputSuccess(result);
    } catch (err) {
      console.error(err);
      const payload = err && typeof err === "object"
        ? err
        : { ok: false, error: String(err) };
      setStatus("Feil");
      outputError(payload);
    }
  }

  // =========================
  // Init
  // =========================
  if ($runBtn) {
    $runBtn.addEventListener("click", () => { main(); });
  }

  if (CONFIG.AUTO_RUN) {
    window.addEventListener("load", () => { main(); });
  }
})();
