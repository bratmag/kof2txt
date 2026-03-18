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

    // Ny blokk / nytt objekt
    if (
      /^OBJ/i.test(line) ||
      /^PUNKT/i.test(line) ||
      /^POINT/i.test(line) ||
      /^BEGIN/i.test(line)
    ) {
      if (isCompletePoint(current)) {
        points.push(normalizePoint(current));
      }
      current = {};
      continue;
    }

    // Slutt på blokk
    if (
      /^END/i.test(line) ||
      /^SLUTT/i.test(line)
    ) {
      if (isCompletePoint(current)) {
        points.push(normalizePoint(current));
      }
      current = {};
      continue;
    }

    // Nøkkel=verdi eller nøkkel: verdi
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

    // Frie linjer: prøv å finne "navn + 3 tall"
    // Eks: P1 6643210.123 245678.456 123.789
    const free = tryParseFreePointLine(line);
    if (free) {
      if (isCompletePoint(current)) {
        points.push(normalizePoint(current));
      }
      current = free;
      continue;
    }
  }

  if (isCompletePoint(current)) {
    points.push(normalizePoint(current));
  }

  // Fjern duplikater på navn+nord+øst+høyde
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
  return [
    "punktnavn",
    "punktnummer",
    "punktnr",
    "punktid",
    "punkt",
    "navn",
    "name",
    "id",
    "label"
  ].includes(key);
}

function isNorthKey(key) {
  return [
    "n",
    "nord",
    "north",
    "northing",
    "y"
  ].includes(key);
}

function isEastKey(key) {
  return [
    "e",
    "ost",
    "east",
    "easting",
    "x"
  ].includes(key);
}

function isHeightKey(key) {
  return [
    "h",
    "z",
    "hoyde",
    "height",
    "elev",
    "elevation",
    "kote"
  ].includes(key);
}
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
