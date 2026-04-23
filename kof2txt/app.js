(() => {
  "use strict";

  const CONFIG = {
    DEBUG: true,
    CONNECT_TIMEOUT_MS: 30000,
    TOKEN_WAIT_MS: 30000,
    PROXY_URL: "/.netlify/functions/tc-proxy",

    // Bytt denne hvis du har en annen stabil ikon-URL
    MENU_ICON_URL: "https://kof2txt.netlify.app/icon-192.png",

    MENU_MAIN_COMMAND: "KOF2TXT_MAIN",
    MENU_OPEN_COMMAND: "KOF2TXT_OPEN"
  };

  const state = {
    api: null,
    accessToken: null,
    project: null,
    selectedFile: null,
    lastResult: null,
    fileList: [],
    tokenWaiters: [],
    isEmbedded: false,
    lastCommand: null
  };

  let ui = {};

  function log(...args) {
    console.log(...args);
  }

  function debug(...args) {
    if (CONFIG.DEBUG) console.log(...args);
  }

  function setStatus(message) {
    log(`[STATUS] ${message}`);
    if (ui.status) ui.status.textContent = message;

    if (state.api?.extension?.setStatusMessage) {
      state.api.extension.setStatusMessage(message).catch(() => {});
    }
  }

  function setOutput(data) {
    log("[OUTPUT]");
    log(data);

    if (ui.output) {
      ui.output.textContent =
        typeof data === "string" ? data : JSON.stringify(data, null, 2);
    }
  }

  function shortText(text, len = 1500) {
    if (typeof text !== "string") return text;
    return text.length > len ? text.slice(0, len) + "..." : text;
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

  function isKofFileName(name) {
    return /\.kof$/i.test(String(name || ""));
  }

  function resolveTokenWaiters(token) {
    const waiters = [...state.tokenWaiters];
    state.tokenWaiters = [];
    for (const resolve of waiters) {
      try {
        resolve(token);
      } catch {}
    }
  }

  function waitForToken(ms = CONFIG.TOKEN_WAIT_MS) {
    if (state.accessToken) return Promise.resolve(state.accessToken);

    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        state.tokenWaiters = state.tokenWaiters.filter((fn) => fn !== wrappedResolve);
        reject(new Error(`Ventet for lenge på access token (${ms} ms)`));
      }, ms);

      function wrappedResolve(token) {
        clearTimeout(timer);
        resolve(token);
      }

      state.tokenWaiters.push(wrappedResolve);
    });
  }

  function buildUi() {
    const app = document.getElementById("app");
    if (!app) {
      throw new Error("Fant ikke #app i index.html");
    }

    app.innerHTML = "";

    const root = document.createElement("div");

    const titleCard = document.createElement("div");
    titleCard.className = "card";

    const title = document.createElement("h2");
    title.textContent = "KOF2TXT";
    title.style.margin = "0 0 8px 0";

    const intro = document.createElement("div");
    intro.className = "muted";
    intro.textContent = "Hent .kof-filer fra prosjektet og konverter valgt fil eller alle filer til .txt.";

    titleCard.appendChild(title);
    titleCard.appendChild(intro);

    const projectCard = document.createElement("div");
    projectCard.className = "card";

    const projectLabel = document.createElement("div");
    projectLabel.style.fontWeight = "bold";
    projectLabel.style.marginBottom = "4px";
    projectLabel.textContent = "Prosjekt";

    const projectValue = document.createElement("div");
    projectValue.textContent = "-";

    projectCard.appendChild(projectLabel);
    projectCard.appendChild(projectValue);

    const actionCard = document.createElement("div");
    actionCard.className = "card";

    const btnRow = document.createElement("div");
    btnRow.className = "btn-row";

    const refreshBtn = document.createElement("button");
    refreshBtn.textContent = "Oppdater liste";

    const convertSelectedBtn = document.createElement("button");
    convertSelectedBtn.textContent = "Konverter valgt";

    const convertAllBtn = document.createElement("button");
    convertAllBtn.textContent = "Konverter alle";


    btnRow.appendChild(refreshBtn);
    btnRow.appendChild(convertSelectedBtn);
    btnRow.appendChild(convertAllBtn);

    const selectedInfo = document.createElement("div");
    selectedInfo.className = "muted";
    selectedInfo.style.marginTop = "8px";
    selectedInfo.textContent = "Ingen fil valgt.";

    const listWrap = document.createElement("div");
    listWrap.style.marginTop = "12px";

    const listLabel = document.createElement("div");
    listLabel.style.fontWeight = "bold";
    listLabel.style.marginBottom = "6px";
    listLabel.textContent = "KOF-filer i prosjektet";

    const fileList = document.createElement("div");
    fileList.id = "fileList";

    listWrap.appendChild(listLabel);
    listWrap.appendChild(fileList);

    actionCard.appendChild(btnRow);
    actionCard.appendChild(selectedInfo);
    actionCard.appendChild(listWrap);

    const statusBox = document.createElement("div");
    statusBox.id = "statusBox";
    statusBox.textContent = "Starter...";

    const output = document.createElement("pre");
    output.id = "output";
    output.textContent = "";

    root.appendChild(titleCard);
    root.appendChild(projectCard);
    root.appendChild(actionCard);
    root.appendChild(statusBox);
    root.appendChild(output);

    app.appendChild(root);

    ui = {
      root,
      projectValue,
      refreshBtn,
      convertSelectedBtn,
      convertAllBtn,
      selectedInfo,
      fileList,
      status: statusBox,
      output
    };
  }

  function renderFileList() {
    if (!ui.fileList) return;

    ui.fileList.innerHTML = "";

    if (!state.fileList.length) {
      const empty = document.createElement("div");
      empty.className = "muted";
      empty.textContent = "Ingen .kof-filer funnet ennå.";
      ui.fileList.appendChild(empty);
      return;
    }

    const wrap = document.createElement("div");
    wrap.style.display = "grid";
    wrap.style.gap = "8px";

    for (const file of state.fileList) {
      const row = document.createElement("label");
      row.style.display = "block";
      row.style.padding = "8px";
      row.style.border = "1px solid #ddd";
      row.style.borderRadius = "8px";
      row.style.cursor = "pointer";

      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = "kofFile";
      radio.value = file.id;
      radio.style.marginRight = "8px";
      radio.checked = state.selectedFile?.id === file.id;

      radio.addEventListener("change", () => {
        state.selectedFile = file;
        updateSelectedInfo();
      });

      const name = document.createElement("strong");
      name.textContent = file.name || "(uten navn)";

      const meta = document.createElement("div");
      meta.className = "muted";
      meta.style.marginTop = "4px";
      meta.textContent = `ID: ${file.id}${file.path ? ` | Sti: ${file.path}` : ""}`;

      row.appendChild(radio);
      row.appendChild(name);
      row.appendChild(meta);
      wrap.appendChild(row);
    }

    ui.fileList.appendChild(wrap);
  }

  function updateSelectedInfo() {
    if (!ui.selectedInfo) return;

    if (!state.selectedFile) {
      ui.selectedInfo.textContent = "Ingen fil valgt.";
      return;
    }

    ui.selectedInfo.textContent =
      `Valgt fil: ${state.selectedFile.name} | ID: ${state.selectedFile.id}`;
  }

  async function connectWorkspace() {
    setStatus("Kobler til Trimble Connect...");

    if (!window.TrimbleConnectWorkspace?.connect) {
      throw new Error(
        "TrimbleConnectWorkspace ikke funnet. Sjekk at Workspace API-scriptet er lastet."
      );
    }

    const api = await TrimbleConnectWorkspace.connect(
      window.parent,
      onWorkspaceEvent,
      CONFIG.CONNECT_TIMEOUT_MS
    );

    state.api = api;
    state.isEmbedded = window.parent && window.parent !== window;
    debug("API keys:", Object.keys(api || {}));

    return api;
  }

  async function ensureMenu() {
  if (!state.api?.ui?.setMenu) {
    debug("ui.setMenu finnes ikke.");
    return false;
  }

  const mainMenuObject = {
    title: "KOF2TXT",
    icon: `${window.location.origin}/icon.png`,
    command: CONFIG.MENU_MAIN_COMMAND,
    subMenus: [
      {
        title: "Konverter KOF",
        command: CONFIG.MENU_OPEN_COMMAND
      }
    ]
  };

  await state.api.ui.setMenu(mainMenuObject);
  await state.api.ui.setActiveMenuItem(CONFIG.MENU_OPEN_COMMAND).catch(() => {});
  debug("Meny satt via ui.setMenu");
  return true;
  }

  async function requestAccessToken() {
    if (state.accessToken) return state.accessToken;

    setStatus("Ber om access token...");

    if (!state.api?.extension?.requestPermission) {
      throw new Error("extension.requestPermission finnes ikke.");
    }

    const result = await state.api.extension.requestPermission("accesstoken");
    debug("requestPermission svar:", result);

    if (typeof result === "string" && result && result !== "pending" && result !== "denied") {
      state.accessToken = result;
      resolveTokenWaiters(result);
      debug("Access token mottatt direkte.");
      return result;
    }

    if (result === "denied") {
      throw new Error("Tilgang til access token ble avslått.");
    }

    if (result === "pending" || !result) {
      debug("Token er pending. Venter på extension.accessToken-event...");
      const token = await waitForToken(CONFIG.TOKEN_WAIT_MS);
      state.accessToken = token;
      return token;
    }

    throw new Error(`Uventet svar fra requestPermission: ${String(result)}`);
  }

  async function getProject() {
    if (state.project) return state.project;

    setStatus("Henter prosjektinfo...");

    const getProjectFn =
      state.api?.project?.getCurrentProject ||
      state.api?.project?.getProject;

    if (!getProjectFn) {
      throw new Error("Fant verken project.getCurrentProject eller project.getProject.");
    }

    const project = await getProjectFn.call(state.api.project);

    if (!project?.id) {
      throw new Error("Fant ikke aktivt prosjekt.");
    }

    state.project = project;
    debug("Project:", project);

    if (ui.projectValue) {
      ui.projectValue.textContent =
        `${project.name || "-"} (${project.id}) | region: ${project.location || "-"}`;
    }

    return project;
  }

  async function ensureReady() {
    if (!state.api) {
      throw new Error("Ikke koblet til Workspace API.");
    }

    if (!state.accessToken) {
      await requestAccessToken();
    }

    if (!state.project) {
      await getProject();
    }
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
    const points = parseKofPoints(kofText);

    if (!points.length) {
      return [
        "Punktnavn,Nord,Øst,Høyde",
        "# Fant ingen punkter i KOF-fila",
        "# Første 1000 tegn fra fila:",
        ...String(kofText || "").slice(0, 1000).split(/\r?\n/)
      ].join("\n");
    }

    const lines = ["Punktnavn,Nord,Øst,Høyde"];

    for (const p of points) {
      lines.push([
        csvEscape(p.name || ""),
        formatNumberForTxt(p.north),
        formatNumberForTxt(p.east),
        formatNumberForTxt(p.height)
      ].join(","));
    }

    return lines.join("\n");
  }

  function parseKofPoints(kofText) {
    const text = String(kofText || "");
    const lines = text.split(/\r?\n/);

    const points = [];
    let current = {};

    for (const rawLine of lines) {
      const line = rawLine.trim();
      if (!line) continue;

      if (/^OBJ/i.test(line) || /^PUNKT/i.test(line) || /^POINT/i.test(line) || /^BEGIN/i.test(line)) {
        if (isCompletePoint(current)) points.push(normalizePoint(current));
        current = {};
        continue;
      }

      if (/^END/i.test(line) || /^SLUTT/i.test(line)) {
        if (isCompletePoint(current)) points.push(normalizePoint(current));
        current = {};
        continue;
      }

      const kv = line.match(/^([^=:]+)\s*[:=]\s*(.+)$/);
      if (kv) {
        const key = normalizeKey(kv[1]);
        const value = kv[2].trim();

        if (!current.name && isNameKey(key)) current.name = cleanValue(value);
        if (current.north == null && isNorthKey(key)) current.north = parseNumber(value);
        if (current.east == null && isEastKey(key)) current.east = parseNumber(value);
        if (current.height == null && isHeightKey(key)) current.height = parseNumber(value);
        continue;
      }

      const free = tryParseFreePointLine(line);
      if (free) {
        if (isCompletePoint(current)) points.push(normalizePoint(current));
        current = free;
      }
    }

    if (isCompletePoint(current)) points.push(normalizePoint(current));

    return dedupePoints(points);
  }

  function tryParseFreePointLine(line) {
    const s = String(line || "").trim();

    let m = s.match(
      /^05\s+([^\s]+)\s+(-?\d+(?:[.,]\d+)?)\s+(-?\d+(?:[.,]\d+)?)\s+(-?\d+(?:[.,]\d+)?)\s*$/
    );

    if (m) {
      return {
        name: m[1],
        north: parseNumber(m[2]),
        east: parseNumber(m[3]),
        height: parseNumber(m[4])
      };
    }

    m = s.match(
      /^([^\s,;]+)[\s,;]+(-?\d+(?:[.,]\d+)?)[\s,;]+(-?\d+(?:[.,]\d+)?)[\s,;]+(-?\d+(?:[.,]\d+)?)\s*$/
    );

    if (m) {
      return {
        name: m[1],
        north: parseNumber(m[2]),
        east: parseNumber(m[3]),
        height: parseNumber(m[4])
      };
    }

    return null;
  }

  function isCompletePoint(p) {
    return !!p && p.name && p.north != null && p.east != null;
  }

  function normalizePoint(p) {
    return {
      name: String(p.name || "").trim(),
      north: p.north != null ? Number(p.north) : null,
      east: p.east != null ? Number(p.east) : null,
      height: p.height != null ? Number(p.height) : null
    };
  }

  function dedupePoints(points) {
    const seen = new Set();
    const out = [];

    for (const p of points) {
      const key = `${p.name}|${p.north}|${p.east}|${p.height}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(p);
    }

    return out;
  }

  function normalizeKey(key) {
    return String(key || "")
      .trim()
      .toLowerCase()
      .replace(/[æøå]/g, (c) => ({ "æ": "ae", "ø": "o", "å": "a" }[c]))
      .replace(/[^a-z0-9]/g, "");
  }

  function cleanValue(value) {
    return String(value || "").trim().replace(/^"|"$/g, "");
  }

  function parseNumber(value) {
    if (value == null) return null;

    const s = String(value)
      .trim()
      .replace(/\s+/g, "")
      .replace(",", ".");

    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  function formatNumberForTxt(n) {
    if (n == null || !Number.isFinite(n)) return "";
    return String(n);
  }

  function csvEscape(value) {
    const s = String(value ?? "");
    if (/[",;\n]/.test(s)) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  }

  function isNameKey(key) {
    return ["punktnavn", "punktnummer", "punktnr", "punktid", "punkt", "navn", "name", "id", "label"].includes(key);
  }

  function isNorthKey(key) {
    return ["n", "nord", "north", "northing", "y"].includes(key);
  }

  function isEastKey(key) {
    return ["e", "ost", "east", "easting", "x"].includes(key);
  }

  function isHeightKey(key) {
    return ["h", "z", "hoyde", "height", "elev", "elevation", "kote"].includes(key);
  }

  async function refreshKofList() {
    try {
      await ensureReady();
      setStatus("Henter .kof-filer fra prosjektet ...");

      const proxyRes = await callProxy("listProjectKofFiles", {
        token: state.accessToken,
        projectId: state.project.id,
        projectLocation: state.project.location
      });

      if (!proxyRes.ok || !proxyRes.json) {
        setStatus("Proxy-feil ved listing");
        setOutput({
          ok: false,
          step: "listProxyHttp",
          status: proxyRes.status,
          preview: shortText(proxyRes.text, 1500)
        });
        return;
      }

      const result = proxyRes.json;

      if (!result.ok) {
        setStatus("Klarte ikke hente filliste");
        setOutput(result);
        return;
      }

      state.fileList = Array.isArray(result.files) ? result.files : [];

      if (!state.fileList.length) {
        state.selectedFile = null;
      } else if (!state.selectedFile || !state.fileList.some((f) => f.id === state.selectedFile.id)) {
        state.selectedFile = state.fileList[0];
      }

      renderFileList();
      updateSelectedInfo();

      setStatus(`Fant ${state.fileList.length} .kof-fil(er)`);
      setOutput({
        ok: true,
        action: "listProjectKofFiles",
        project: result.project,
        fileCount: state.fileList.length,
        candidatesTried: result.candidatesTried,
        diagnostics: result.diagnostics
      });
    } catch (err) {
      console.error(err);
      setStatus("Feil ved henting av filliste");
      setOutput({
        ok: false,
        error: err?.message || String(err)
      });
    }
  }

  async function downloadAndConvertFile(file) {
    const proxyRes = await callProxy("downloadKofFile", {
      token: state.accessToken,
      projectId: state.project.id,
      projectLocation: state.project.location,
      fileId: file.id,
      fileName: file.name
    });

    if (!proxyRes.ok || !proxyRes.json) {
      throw new Error(`Proxy-feil (${proxyRes.status})`);
    }

    const result = proxyRes.json;
    if (!result.ok) {
      throw new Error(result.error || result.step || "Klarte ikke laste ned KOF-fil");
    }

    const txt = convertKofToTxt(result.text || "");
    const outName = String(result.file?.name || file.name || "output.kof").replace(/\.kof$/i, ".txt");

    return {
      outName,
      txt,
      result
    };
  }

  async function processSelectedFile() {
    try {
      await ensureReady();

      const file = state.selectedFile;
      if (!file?.id) {
        setStatus("Ingen fil valgt");
        setOutput({
          ok: false,
          step: "noSelectedFile",
          message: "Velg en .kof-fil i listen først."
        });
        return;
      }

      setStatus(`Laster ned ${file.name} ...`);

      const converted = await downloadAndConvertFile(file);
      state.lastResult = converted.result;

      setStatus(`Klar: ${converted.outName}`);
      setOutput({
        ok: true,
        project: converted.result.project,
        file: converted.result.file,
        source: converted.result.source,
        contentType: converted.result.contentType,
        preview: shortText(converted.result.text || "", 1500)
      });

      triggerDownload(converted.outName, converted.txt);
    } catch (err) {
      console.error(err);
      setStatus("Feil");
      setOutput({
        ok: false,
        error: err?.message || String(err)
      });
    }
  }

  async function processAllFiles() {
    try {
      await ensureReady();

      if (!state.fileList.length) {
        setStatus("Ingen .kof-filer i listen");
        setOutput({
          ok: false,
          step: "noFiles",
          message: "Trykk Oppdater liste først."
        });
        return;
      }

      const summary = [];
      let count = 0;

      for (const file of state.fileList) {
        setStatus(`Konverterer (${count + 1}/${state.fileList.length}) ${file.name} ...`);

        try {
          const converted = await downloadAndConvertFile(file);
          triggerDownload(converted.outName, converted.txt);

          summary.push({
            ok: true,
            file: file.name,
            outName: converted.outName
          });
        } catch (err) {
          summary.push({
            ok: false,
            file: file.name,
            error: err?.message || String(err)
          });
        }

        count += 1;
      }

      const okCount = summary.filter((x) => x.ok).length;
      const failCount = summary.length - okCount;

      setStatus(`Ferdig. OK: ${okCount}, Feil: ${failCount}`);
      setOutput({
        ok: failCount === 0,
        action: "convertAll",
        total: summary.length,
        okCount,
        failCount,
        files: summary
      });
    } catch (err) {
      console.error(err);
      setStatus("Feil i Konverter alle");
      setOutput({
        ok: false,
        error: err?.message || String(err)
      });
    }
  }


  function onWorkspaceEvent(event, args) {
    debug("[TC EVENT]", event, args);

    if (event === "extension.accessToken") {
      const token = args?.data;
      if (typeof token === "string" && token && token !== "pending" && token !== "denied") {
        state.accessToken = token;
        resolveTokenWaiters(token);
        debug("Access token mottatt via event.");
      }
      return;
    }

    if (event === "extension.command") {
      state.lastCommand = args?.data || null;
      debug("extension.command:", state.lastCommand);

      if (state.lastCommand === CONFIG.MENU_OPEN_COMMAND) {
        setStatus("KOF2TXT åpnet fra meny.");
      }

      return;
    }
  }

  function wireUi() {
    ui.refreshBtn.addEventListener("click", () => {
      refreshKofList();
    });

    ui.convertSelectedBtn.addEventListener("click", () => {
      processSelectedFile();
    });

    ui.convertAllBtn.addEventListener("click", () => {
      processAllFiles();
    });

  }

  async function init() {
    try {
      buildUi();
      wireUi();

      setStatus("Starter...");
      await connectWorkspace();
      await ensureMenu();

      setStatus("Klar. Trykk Oppdater liste.");
      setOutput({
        ok: true,
        embedded: state.isEmbedded,
        message: "Extension lastet. Meny registrert. Token og prosjekt hentes når du trykker Oppdater liste eller kjører en handling."
      });

      window.kof2txt = {
  state,
  refreshKofList,
  processSelectedFile,
  processAllFiles,
  ensureMenu,
  async reconnect() {
    state.api = null;
    state.accessToken = null;
    state.project = null;
    await connectWorkspace();
    await ensureMenu();
    return true;
  },
  async prime() {
    await ensureReady();
    return {
      accessToken: !!state.accessToken,
      project: state.project
    };
  }
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

  window.addEventListener("load", init);
})();
