// Gets the coach workbooks out of an anonymous OneDrive / SharePoint "anyone with the link"
// folder share. The share link is a secret: it is never printed. Every log line goes
// through redact(). The Actions log of a public repo is public, so logs also never contain
// coach names, client names or figures.
import { unzipSync } from "fflate";

const UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36";
const SHEET_RE = /\.(xlsx|xlsm|xls)$/i;

export function makeLogger(shareUrl) {
  const secrets = new Set();
  try {
    const u = new URL(shareUrl);
    u.pathname.split("/").filter((p) => p.length >= 12).forEach((p) => secrets.add(p));
    u.searchParams.forEach((v) => { if (v.length >= 6) secrets.add(v); });
    secrets.add(shareUrl);
  } catch {}
  const enc = Buffer.from(shareUrl).toString("base64").replace(/=+$/, "");
  secrets.add(enc); secrets.add(enc.replace(/\//g, "_").replace(/\+/g, "-"));
  const redact = (s) => {
    let out = String(s);
    for (const x of [...secrets].sort((a, b) => b.length - a.length)) out = out.split(x).join("<redacted>");
    return out.replace(/(sig|tempauth|token|access_token)=[^&\s"']+/gi, "$1=<redacted>");
  };
  return { redact, log: (...a) => console.log(redact(a.join(" "))) };
}

function shareParts(shareUrl) {
  const u = new URL(shareUrl);
  const m = u.pathname.match(/^\/:[a-z]:\/[a-z]\/(personal\/[^/]+|sites\/[^/]+|teams\/[^/]+)\/([^/?#]+)/i);
  return { u, host: u.host, site: m ? "/" + m[1] : null, token: m ? m[2] : null };
}

async function get(url, opts, L, label) {
  const res = await fetch(url, { redirect: "follow", ...opts, headers: { "User-Agent": UA, ...(opts?.headers || {}) } });
  L.log(`  ${label}: HTTP ${res.status} ${res.headers.get("content-type") || ""}`);
  return res;
}

// ---- Strategy "api": SharePoint/OneDrive shares API, anonymous with the share token ----
async function viaApi(shareUrl, L) {
  const { host } = shareParts(shareUrl);
  const enc = "u!" + Buffer.from(shareUrl).toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  const base = `https://${host}/_api/v2.0/shares/${enc}/driveItem`;
  const files = [];
  async function walk(path, depth) {
    const res = await get(`${base}${path}/children?$top=200`, { headers: { Accept: "application/json" } }, L, "list" + (depth ? " (subfolder)" : ""));
    if (!res.ok) { const t = (await res.text()).slice(0, 200); throw new Error(`listing failed (${res.status}) ${t}`); }
    const j = await res.json();
    for (const it of j.value || []) {
      if (it.folder && depth < 2) { try { await walk(`/items/${it.id}`, depth + 1); } catch (e) { L.log("  subfolder skipped:", e.message); } }
      else if (it.file && SHEET_RE.test(it.name)) {
        const url = it["@content.downloadUrl"] || it["@microsoft.graph.downloadUrl"] || `${base}/items/${it.id}/content`;
        files.push({ name: it.name, url });
      }
    }
  }
  await walk("", 0);
  const out = [];
  for (const f of files) {
    const res = await get(f.url, { headers: { Accept: "*/*" } }, L, "download");
    if (!res.ok) throw new Error(`download failed (${res.status})`);
    out.push({ name: f.name, buf: Buffer.from(await res.arrayBuffer()) });
  }
  return out;
}

// ---- Strategy "zip": download.aspx?share=<token> returns the whole folder as a zip ----
async function viaZip(shareUrl, L) {
  const { host, site, token } = shareParts(shareUrl);
  if (!site || !token) throw new Error("share link has an unexpected shape");
  const url = `https://${host}${site}/_layouts/15/download.aspx?share=${token}`;
  const res = await get(url, { headers: { Accept: "*/*" } }, L, "zip");
  if (!res.ok) throw new Error(`zip download failed (${res.status})`);
  const buf = Buffer.from(await res.arrayBuffer());
  if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) throw new Error(`response was not a zip (${buf.length} bytes)`);
  const entries = unzipSync(new Uint8Array(buf));
  const out = [];
  for (const [path, data] of Object.entries(entries)) {
    const name = path.split("/").pop();
    if (SHEET_RE.test(name) && !name.startsWith("~$")) out.push({ name, buf: Buffer.from(data) });
  }
  if (!out.length) {
    // The share may itself be a single workbook, so the zip is the xlsx.
    throw new Error("zip contained no spreadsheets");
  }
  return out;
}

// ---- Strategy "rest": follow the link (collect cookies), then use SharePoint REST ----
async function viaRest(shareUrl, L) {
  const { host } = shareParts(shareUrl);
  const jar = new Map();
  const cookie = () => [...jar].map(([k, v]) => `${k}=${v}`).join("; ");
  const eat = (res) => (res.headers.getSetCookie?.() || []).forEach((c) => { const [kv] = c.split(";"); const i = kv.indexOf("="); if (i > 0) jar.set(kv.slice(0, i).trim(), kv.slice(i + 1)); });
  let url = shareUrl, finalUrl = shareUrl;
  for (let i = 0; i < 10; i++) {
    const res = await fetch(url, { redirect: "manual", headers: { "User-Agent": UA, Cookie: cookie(), Accept: "text/html" } });
    eat(res);
    const loc = res.headers.get("location");
    L.log(`  follow ${i}: HTTP ${res.status}${loc ? " -> redirect" : ""}`);
    if (res.status >= 300 && res.status < 400 && loc) { url = new URL(loc, url).toString(); finalUrl = url; continue; }
    break;
  }
  const id = new URL(finalUrl).searchParams.get("id");
  if (!id) throw new Error("could not find the folder path in the redirected address");
  const siteMatch = id.match(/^(\/(?:personal|sites|teams)\/[^/]+)/);
  if (!siteMatch) throw new Error("folder path had an unexpected shape");
  const apiBase = `https://${host}${siteMatch[1]}/_api/web`;
  const H = { "User-Agent": UA, Cookie: cookie(), Accept: "application/json;odata=nometadata" };
  const esc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const res = await get(`${apiBase}/GetFolderByServerRelativePath(decodedurl='${esc(id)}')/Files?$select=Name,ServerRelativeUrl`, { headers: H }, L, "list");
  if (!res.ok) throw new Error(`listing failed (${res.status})`);
  const j = await res.json();
  const out = [];
  for (const f of j.value || []) {
    if (!SHEET_RE.test(f.Name)) continue;
    const r = await get(`${apiBase}/GetFileByServerRelativePath(decodedurl='${esc(f.ServerRelativeUrl)}')/$value`, { headers: { ...H, Accept: "*/*" } }, L, "download");
    if (!r.ok) throw new Error(`download failed (${r.status})`);
    out.push({ name: f.Name, buf: Buffer.from(await r.arrayBuffer()) });
  }
  return out;
}

const STRATEGIES = { api: viaApi, zip: viaZip, rest: viaRest };

export async function fetchWorkbooks(shareUrl, order, L) {
  const tried = [];
  for (const name of order) {
    const fn = STRATEGIES[name]; if (!fn) continue;
    L.log(`Fetch strategy "${name}"`);
    try {
      const files = await fn(shareUrl, L);
      if (files.length) { L.log(`  -> ${files.length} workbook(s)`); return { strategy: name, files }; }
      tried.push(`${name}: no workbooks found`); L.log("  no workbooks found");
    } catch (e) {
      tried.push(`${name}: ${L.redact(e.message)}`); L.log("  failed:", e.message);
    }
  }
  const err = new Error("All fetch strategies failed.\n" + tried.map((t) => " - " + t).join("\n"));
  err.tried = tried;
  throw err;
}
