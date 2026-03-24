(() => {
  "use strict";

  const CONFIG = {
    DEBUG: true,
    CONNECT_TIMEOUT_MS: 30000,
    PROXY_URL: "/.netlify/functions/tc-proxy",
    DEFAULT_TEST_FILE_ID: "RZPc08vH2VU",
    DEFAULT_TEST_FILE_NAME: "Eiendomspunkter kof.kof"
  };

  const state = {
    api: null,
    accessToken: null,
    project: null,
    selectedFile: null,
    lastResult: null
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

  function ensureKofFileName(name) {
    const n = String(name || "").trim();
    if (!n) return "output.kof";
    return isKofFileName(n) ? n : `${n}.kof`;
  }

  function fileFromInputs() {
    const id = String(ui.fileIdInput?.value || "").trim();
    const name = ensureKofFileName(ui.fileNameInput?.value || "");

    if (!id) return null;

    return { id, name };
  }

  function setInputsFromFile(file) {
    if (!file) return;
    if (ui.fileIdInput) ui.fileIdInput.value = file.id || "";
    if (ui.fileNameInput) ui.fileNameInput.value = file.name || "";
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
    intro.textContent = "Skriv inn File ID manuelt og trykk Konverter KOF.";

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

    const formCard = document.createElement("div");
    formCard.className = "card";

    const form = document.createElement("div");
    form.className = "row";

    const fileIdWrap = document.createElement("div");
    const fileIdLabel = document.createElement("label");
    fileIdLabel.textContent = "File ID";
    const fileIdInput = document.createElement("input");
    fileIdInput.type = "text";
    fileIdInput.placeholder = "F.eks. RZPc08vH2VU";
    fileIdWrap.appendChild(fileIdLabel);
    fileIdWrap.appendChild(fileIdInput);

    const fileNameWrap = document.createElement("div");
    const fileNameLabel = document.createElement("label");
    fileNameLabel.textContent = "Filnavn";
    const fileNameInput = document.createElement("input");
    fileNameInput.type = "text";
    fileNameInput.placeholder = "F.eks. Eiendomspunkter kof.kof";
    fileNameWrap.appendChild(fileNameLabel);
    fileNameWrap.appendChild(fileNameInput);

    form.appendChild(fileIdWrap);
    form.appendChild(fileNameWrap);

    const btnRow = document.createElement("div");
    btnRow.className = "btn-row";

    const convertBtn = document.createElement("button");
    convertBtn.textContent = "Konverter KOF";

    const testBtn = document.createElement("button");
    testBtn.textContent = "Bruk testfil";

    const probeBtn = document.createElement("button");
    probeBtn.textContent = "Core probe";

    btnRow.appendChild(convertBtn);
    btnRow.appendChild(testBtn);
    btnRow.appendChild(probeBtn);

    formCard.appendChild(form);
    formCard.appendChild(btnRow);

    const statusBox = document.createElement("div");
    statusBox.id = "statusBox";
    statusBox.textContent = "Starter...";

    const output = document.createElement("pre");
    output.id = "output";
    output.textContent = "";

    root.appendChild(titleCard);
    root.appendChild(projectCard);
    root.appendChild(formCard);
    root.appendChild(statusBox);
    root.appendChild(output);

    app.appendChild(root);

    ui = {
      root,
      projectValue,
      fileIdInput,
      fileNameInput,
      convertBtn,
      testBtn,
      probeBtn,
      status: statusBox,
      output
    };
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

    if (ui.projectValue) {
      ui.projectValue.textContent = `${project.name || "-"} (${project.id}) | region: ${project.location || "-"}`;
    }

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
    const m = line.match(
      /^([^\s,;]+)[\s,;]+(-?\d+(?:[.,]\d+)?)[\s,;]+(-?\d+(?:[.,]\d+)?)[\s,;]+(-?\d+(?:[.,]\d+)?)$/
    );

    if (!m) return null;

    return {
      name: m[1],
      north: parseNumber(m[2]),
      east: parseNumber(m[3]),
      height: parseNumber(m[4])
    };
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

  async function processSelectedFile() {
    try {
      if (!state.api) throw new Error("Ikke koblet til Workspace API.");
      if (!state.accessToken) await requestAccessToken();
      if (!state.project) await getProject();

      const file = fileFromInputs();
      if (!file) {
        setStatus("Mangler File ID");
        setOutput({
          ok: false,
          step: "noSelectedFile",
          message: "Skriv inn File ID først."
        });
        return;
      }

      state.selectedFile = file;

      setStatus(`Laster ned ${file.name} ...`);

      const proxyRes = await callProxy("downloadKofFile", {
        token: state.accessToken,
        projectId: state.project.id,
        projectLocation: state.project.location,
        fileId: file.id,
        fileName: file.name
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

      const file = fileFromInputs();
      if (!file?.id) {
        setOutput({
          ok: false,
          step: "probeNoFile",
          message: "Skriv inn File ID først."
        });
        return;
      }

      state.selectedFile = file;
      setStatus("Kjører Core probe...");

      const proxyRes = await callProxy("probeCore", {
        token: state.accessToken,
        projectId: state.project.id,
        projectLocation: state.project.location,
        fileId: file.id,
        fileName: file.name
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

  function useTestFile() {
    const testFile = {
      id: CONFIG.DEFAULT_TEST_FILE_ID,
      name: CONFIG.DEFAULT_TEST_FILE_NAME
    };

    state.selectedFile = testFile;
    setInputsFromFile(testFile);

    setStatus(`Testfil satt: ${testFile.name}`);
    setOutput({
      ok: true,
      message: "Testfil satt i inputfeltene",
      file: testFile
    });
  }

  function onWorkspaceEvent(event, args) {
    debug("[TC EVENT]", event, args);
  }

  function wireUi() {
    ui.convertBtn.addEventListener("click", () => {
      processSelectedFile();
    });

    ui.testBtn.addEventListener("click", () => {
      useTestFile();
    });

    ui.probeBtn.addEventListener("click", () => {
      runCoreProbe();
    });
  }

  async function init() {
    try {
      buildUi();
      wireUi();

      setStatus("Starter...");
      await connectWorkspace();
      await requestAccessToken();
      await getProject();

      setStatus("Klar. Skriv inn File ID og trykk Konverter KOF.");
      setOutput({
        ok: true,
        project: {
          id: state.project.id,
          name: state.project.name,
          location: state.project.location
        },
        message: "Extension er klar. Bruk inputfeltene for å konvertere KOF."
      });

      window.kof2txt = {
        state,
        processSelectedFile,
        runCoreProbe,
        useTestFile,
        setFile(fileId, fileName) {
          const file = {
            id: String(fileId || "").trim(),
            name: ensureKofFileName(fileName || "")
          };
          state.selectedFile = file;
          setInputsFromFile(file);
          return file;
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
