/* global TrimbleConnectWorkspace */

let API = null;
let ACCESS_TOKEN = null;
let PROJECT_INFO = null;

const els = {
  fileInput: document.getElementById("fileInput"),
  btnConvert: document.getElementById("btnConvert"),
  btnConvertProject: document.getElementById("btnConvertProject"),
  status: document.getElementById("status"),
  preview: document.getElementById("preview"),
};

function setStatus(msg, type = "muted") {
  if (!els.status) return;
  els.status.className = type;
  els.status.textContent = msg;
}

function showPreview(text) {
  if (!els.preview) return;
  els.preview.style.display = "block";
  els.preview.textContent = text;
}

function appendPreview(text) {
  if (!els.preview) return;
  els.preview.style.display = "block";
  els.preview.textContent = `${els.preview.textContent || ""}\n\n${text}`.trim();
}

function safeJson(x) {
  try {
    return JSON.stringify(x, null, 2);
  } catch {
    return String(x);
  }
}

function isJwtLike(s) {
  return typeof s === "string" && s.startsWith("eyJ") && s.length > 100;
}

function downloadText(filename, text) {
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// --------------------
// KOF -> TXT
// Output: punktId,nord,øst,høyde
// --------------------
function parseKofLine(line) {
  const cleaned = (line || "").trim();
  if (!cleaned) return null;

  const parts = cleaned.split(/\s+/).filter(Boolean);
  if (parts.length < 5) return null;
  if (parts[0] !== "05") return null;

  return {
    pointId: parts[1],
    north: parts[2],
    east: parts[3],
    height: parts[4],
  };
}

function kofToCsv4(kofText) {
  const out = [];
  for (const line of (kofText || "").split(/\r?\n/)) {
    const r = parseKofLine(line);
    if (!r) continue;
    out.push(`${r.pointId},${r.north},${r.east},${r.height}`);
  }
  return out.join("\r\n") + "\r\n";
}

// --------------------
// TC region -> API base
// --------------------
function getTcApiBase(location) {
  const loc = (location || "").toLowerCase();
  if (loc.includes("europe") || loc === "eu") {
    return "https://app21.connect.trimble.com/tc/api/2.0";
  }
  if (loc.includes("asia") || loc === "ap" || loc === "apac") {
    return "https://app31.connect.trimble.com/tc/api/2.0";
  }
  return "https://app.connect.trimble.com/tc/api/2.0";
}

// --------------------
// Proxy calls via Netlify Function
// --------------------
async function proxyFetchText(url, { method = "GET", headers = {}, body = undefined } = {}) {
  const res = await fetch("/.netlify/functions/tc-proxy", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      url,
      method,
      token: ACCESS_TOKEN,
      headers,
      body,
    }),
  });

  const text = await res.text();

  if (!res.ok) {
    throw new Error(`Proxy HTTP ${res.status}\nURL: ${url}\nBody:\n${text.slice(0, 2000)}`);
  }

  return text;
}

async function proxyFetchJson(url, opts) {
  const text = await proxyFetchText(url, opts);
  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

// --------------------
// Workspace / token
// --------------------
async function tryConnectWorkspace() {
  API = await TrimbleConnectWorkspace.connect(
    window.parent,
    async (event, args) => {
      if (event === "extension.accessToken") {
        const tok = args?.data?.accessToken || args?.data || args?.accessToken;
        if (isJwtLike(tok)) {
          ACCESS_TOKEN = tok;
          appendPreview(`✅ ACCESS TOKEN (event) mottatt.\nLengde: ${tok.length}\nStarter med: ${tok.slice(0, 20)}...`);
          setStatus("Access token mottatt ✅", "ok");
        }
      }

      if (event === "extension.command") {
        appendPreview("extension.command:\n" + safeJson(args));
      }
    },
    30000
  );

  // 👇 DETTE ER DET VIKTIGE SOM GJØR AT DEN SYNS I VENSTRESIDEN
  await API.ui.setMenu({
    title: "KOF → TXT",
    command: "kof2txt_main",
    subMenus: [
      {
        title: "Konverter .kof → .txt",
        command: "kof2txt_convert"
      }
    ]
  });

  await API.ui.setActiveMenuItem("kof2txt_convert");

  PROJECT_INFO = await API.project.getProject();
  appendPreview("projectInfo:\n" + safeJson(PROJECT_INFO));
  setStatus("Koblet til Trimble Connect.", "ok");
}

async function requestAccessToken() {
  if (ACCESS_TOKEN) return ACCESS_TOKEN;

  setStatus("Ber om access token...", "muted");
  appendPreview("Sender requestPermission('accesstoken') ...");

  const ret = await API.extension.requestPermission("accesstoken");

  const maybeTok =
    ret?.accessToken ||
    ret?.data?.accessToken ||
    ret?.data ||
    ret;

  if (isJwtLike(maybeTok)) {
    ACCESS_TOKEN = maybeTok;
    appendPreview(`✅ ACCESS TOKEN (requestPermission()) mottatt.\nLengde: ${ACCESS_TOKEN.length}\nStarter med: ${ACCESS_TOKEN.slice(0, 20)}...`);
    return ACCESS_TOKEN;
  }

  appendPreview("Venter på extension.accessToken-event...");

  const tok = await new Promise((resolve, reject) => {
    const t0 = Date.now();
    const timer = setInterval(() => {
      if (ACCESS_TOKEN) {
        clearInterval(timer);
        resolve(ACCESS_TOKEN);
        return;
      }
      if (Date.now() - t0 > 15000) {
        clearInterval(timer);
        reject(new Error("Token event kom ikke innen 15 sek"));
      }
    }, 250);
  });

  return tok;
}

// --------------------
// Local convert
// --------------------
async function handleConvertClick() {
  const files = els.fileInput?.files;
  if (!files || files.length === 0) {
    setStatus("Velg minst én .kof-fil først.", "err");
    return;
  }

  setStatus(`Konverterer ${files.length} fil(er) lokalt...`, "muted");

  for (const f of files) {
    const text = await f.text();
    const csv = kofToCsv4(text);
    const baseName = f.name.replace(/\.[^.]+$/, "");
    downloadText(`${baseName}.txt`, csv);
    if (files.length === 1) {
      showPreview(csv.slice(0, 2500));
    }
  }

  setStatus("Ferdig (lokal konvertering).", "ok");
}

// --------------------
// Project listing helpers
// --------------------
async function getProjectRootId(apiBase, projectId) {
  const p = await proxyFetchJson(`${apiBase}/projects/${encodeURIComponent(projectId)}`);
  const rootId = p?.rootId || p?.data?.rootId || p?.rootFolderId;

  if (!rootId) {
    throw new Error("Fant ikke rootId i prosjekt-respons:\n" + safeJson(p));
  }

  return rootId;
}

async function listChildren(apiBase, parentId) {
  const r = await proxyFetchJson(`${apiBase}/files?parentId=${encodeURIComponent(parentId)}`);

  if (Array.isArray(r)) return r;
  if (Array.isArray(r?.items)) return r.items;
  if (Array.isArray(r?.data)) return r.data;

  appendPreview("Uventet respons fra listChildren:\n" + safeJson(r));
  return [];
}

function isFolder(item) {
  return item?.type === "folder" || item?.isFolder === true || item?.itemType === "Folder";
}

function isKofFile(item) {
  const name = (item?.name || item?.fileName || "").toLowerCase();
  return name.endsWith(".kof");
}

async function listKofFilesRecursive(apiBase, folderId, acc = []) {
  const children = await listChildren(apiBase, folderId);

  for (const it of children) {
    if (isFolder(it)) {
      const childId = it?.id || it?.folderId;
      if (childId) {
        await listKofFilesRecursive(apiBase, childId, acc);
      }
    } else if (isKofFile(it)) {
      acc.push(it);
    }
  }

  return acc;
}

async function getDownloadUrl(apiBase, fileItem) {
  if (fileItem?.downloadUrl) return fileItem.downloadUrl;

  const fileId = fileItem?.id || fileItem?.fileId;
  if (!fileId) {
    throw new Error("Fant ikke fileId på filobjekt:\n" + safeJson(fileItem));
  }

  const meta = await proxyFetchJson(`${apiBase}/files/${encodeURIComponent(fileId)}`);
  const url = meta?.downloadUrl || meta?.data?.downloadUrl;

  if (!url) {
    throw new Error("Fant ikke downloadUrl i metadata:\n" + safeJson(meta));
  }

  return url;
}

// --------------------
// Convert all .kof in project
// --------------------
async function handleConvertProjectClick() {
  try {
    if (!API || !PROJECT_INFO?.id) {
      setStatus("Ikke koblet til Connect / mangler prosjektinfo.", "err");
      return;
    }

    await requestAccessToken();

    const apiBase = getTcApiBase(PROJECT_INFO.location);
    appendPreview(`API base: ${apiBase}`);

    setStatus("Henter root-folder...", "muted");
    const rootId = await getProjectRootId(apiBase, PROJECT_INFO.id);
    appendPreview(`rootId: ${rootId}`);

    setStatus("Lister .kof-filer i prosjektet...", "muted");
    const kofFiles = await listKofFilesRecursive(apiBase, rootId);
    appendPreview(`Fant ${kofFiles.length} .kof-fil(er).`);

    if (kofFiles.length === 0) {
      setStatus("Fant ingen .kof-filer i prosjektet.", "muted");
      return;
    }

    let done = 0;

    for (const f of kofFiles) {
      done += 1;

      const name = f?.name || f?.fileName || `file_${done}.kof`;
      appendPreview(`\n---\n🔽 ${name} (${done}/${kofFiles.length})`);

      const dlUrl = await getDownloadUrl(apiBase, f);
      const kofText = await proxyFetchText(dlUrl, { method: "GET" });
      const txt = kofToCsv4(kofText);

      const baseName = name.replace(/\.[^.]+$/, "");
      const outName = `${baseName}.txt`;

      downloadText(outName, txt);

      if (done === 1) {
        showPreview(txt.slice(0, 2500));
      }

      setStatus(`Konvertert ${done}/${kofFiles.length} ...`, "muted");
    }

    setStatus(`Ferdig! Konverterte ${done} .kof-fil(er) og lastet ned .txt lokalt.`, "ok");
    appendPreview("\n✅ Neste steg er å laste opp .txt tilbake til prosjektet.");

  } catch (e) {
    console.error(e);
    appendPreview("\n❌ Feil:\n" + String(e));
    setStatus("Feil under prosjekt-konvertering (se preview).", "err");
  }
}

// --------------------
// Boot
// --------------------
(async function main() {
  if (els.btnConvert) {
    els.btnConvert.addEventListener("click", handleConvertClick);
  }

  if (els.btnConvertProject) {
    els.btnConvertProject.addEventListener("click", handleConvertProjectClick);
  }

  try {
    await tryConnectWorkspace();
  } catch (e) {
    setStatus("Åpne denne siden via Trimble Connect (ikke direkte i nettleser).", "err");
    appendPreview(String(e));
  }
})();
