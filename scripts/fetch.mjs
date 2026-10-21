// Gets the coach workbooks out of an anonymous "anyone with the link" OneDrive folder share.
// Two families of link exist and they need different APIs:
//   consumer OneDrive (personal Microsoft account) -> api.onedrive.com, "Vroom"
//   business SharePoint/OneDrive                   -> <tenant>.sharepoint.com/_api
// Both are tried, consumer first.
//
// The share link is a secret. It is never printed: every log line goes through redact(), and
// describeLink() prints only the SHAPE of the link - host, known path words, parameter names -
// never a token, an id or a file path. The Actions log of a public repository is public, so
// logs also never carry coach names, client names or figures.
import { unzipSync } from "fflate";

const UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36";
const SHEET_RE = /\.(xlsx|xlsm|xls)$/i;

export function makeLogger(shareUrl) {
  const secrets = new Set();
  const add = (v) => { if (v && String(v).length >= 6) secrets.add(String(v)); };
  try {
    const u = new URL(shareUrl);
    u.pathname.split("/").filter((p) => p.length >= 10).forEach(add);
    u.searchParams.forEach(add);
    [...u.searchParams.values()].forEach((v) => { try { add(decodeURIComponent(v)); } catch {} });
  } catch {}
  add(shareUrl);
  const b64 = Buffer.from(shareUrl).toString("base64").replace(/=+$/, "");
  add(b64); add(b64.replace(/\//g, "_").replace(/\+/g, "-"));
  const redact = (s) => {
    let out = String(s);
    for (const x of [...secrets].sort((a, b) => b.length - a.length)) out = out.split(x).join("<redacted>");
    return out
      .replace(/(sig|tempauth|token|access_token|authkey|resid|cid|share|at|e)=[^&\s"'<]+/gi, "$1=<redacted>")
      .replace(/https?:\/\/[^\s"'<]*\/(download|content)[^\s"'<]*/gi, "<download url redacted>");
  };
  return { redact, log: (...a) => console.log(redact(a.join(" "))) };
}

// Prints what kind of link this is without revealing it. Only words from a fixed list are
// printed literally; everything else becomes <n chars>.
const SAFE_WORDS = new Set(["personal", "sites", "teams", "_layouts", "15", "onedrive.aspx", "redir.aspx",
  "download.aspx", "guestaccess.aspx", "Documents", "Shared Documents", "f", "g", "r", "s", "u", "w", "b", "t", "v", "y", "p"]);
export function describeLink(shareUrl, L) {
  let u;
  try { u = new URL(shareUrl); } catch { L.log("SHARE_URL is not a valid web address - check the secret."); return {}; }
  const seg = (s) => (SAFE_WORDS.has(s) || /^:[a-z]:$/.test(s) ? s : `<${s.length} chars>`);
  const shape = u.pathname.split("/").filter(Boolean).map(seg).join("/");
  const params = [...u.searchParams.keys()];
  L.log(`Link shape: host=${u.host} path=/${shape} params=[${params.join(", ")}]`);
  const idRaw = u.searchParams.get("id");
  if (idRaw) {
    let id = idRaw; try { id = decodeURIComponent(idRaw); } catch {}
    L.log(`  id looks like: /${id.split("/").filter(Boolean).map(seg).join("/")}`);
  }
  return { u, host: u.host, consumerish: /1drv\.ms|onedrive\.live\.com|storage\.live\.com/i.test(u.host) };
}

async function get(url, opts, L, label) {
  const res = await fetch(url, { redirect: "follow", ...opts, headers: { "User-Agent": UA, ...(opts?.headers || {}) } });
  L.log(`  ${label}: HTTP ${res.status} ${(res.headers.get("content-type") || "").split(";")[0]}`);
  return res;
}
function encodeShare(url) {
  return "u!" + Buffer.from(url).toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
}
// Follows redirects by hand so a short 1drv.ms link can be resolved to its real address,
// collecting cookies on the way (some anonymous links need the session cookie to download).
async function resolveLink(shareUrl, L) {
  const jar = new Map();
  const cookie = () => [...jar].map(([k, v]) => `${k}=${v}`).join("; ");
  const eat = (res) => (res.headers.getSetCookie?.() || []).forEach((c) => {
    const [kv] = c.split(";"); const i = kv.indexOf("=");
    if (i > 0) jar.set(kv.slice(0, i).trim(), kv.slice(i + 1));
  });
  let url = shareUrl, hops = 0;
  for (; hops < 10; hops++) {
    const res = await fetch(url, { redirect: "manual", headers: { "User-Agent": UA, Cookie: cookie(), Accept: "text/html,*/*" } });
    eat(res);
    const loc = res.headers.get("location");
    if (res.status >= 300 && res.status < 400 && loc) { url = new URL(loc, url).toString(); continue; }
    L.log(`  resolved after ${hops} redirect(s): HTTP ${res.status}`);
    break;
  }
  return { finalUrl: url, cookie: cookie() };
}

function pickFiles(children) {
  return (children || []).filter((it) => it.file && SHEET_RE.test(it.name) && !it.name.startsWith("~$"))
    .map((it) => ({ name: it.name, url: it["@content.downloadUrl"] || it["@microsoft.graph.downloadUrl"] || null, id: it.id }));
}
async function download(files, L, headers) {
  const out = [];
  for (const f of files) {
    if (!f.url) throw new Error(`no download address for one of the files`);
    const res = await get(f.url, { headers: { Accept: "*/*", ...(headers || {}) } }, L, "download");
    if (!res.ok) throw new Error(`download failed (${res.status})`);
    out.push({ name: f.name, buf: Buffer.from(await res.arrayBuffer()) });
  }
  return out;
}

/* ---- "od": consumer OneDrive shares API. The documented anonymous route for a personal
   account's "anyone with the link" share: the link itself, base64url-encoded, is the key. ---- */
async function viaOdShares(shareUrl, L) {
  const key = encodeShare(shareUrl);
  const bases = [
    `https://api.onedrive.com/v1.0/shares/${key}`,
    `https://my.microsoftpersonalcontent.com/_api/v2.0/shares/${key}`,
  ];
  let lastErr = "no endpoint answered";
  for (const base of bases) {
    for (const path of ["/root?expand=children(expand=thumbnails)", "/root?expand=children", "/driveItem?expand=children"]) {
      const res = await get(base + path, { headers: { Accept: "application/json" } }, L, "list");
      if (!res.ok) { lastErr = `listing failed (${res.status})`; if (res.status === 404 || res.status === 400) continue; else { lastErr += " " + (await res.text()).slice(0, 160); continue; } }
      const j = await res.json();
      const children = j.children || j.value || (j.folder ? [] : null);
      const files = pickFiles(children);
      if (files.length) return download(files, L);
      lastErr = "the folder answered but held no spreadsheets";
    }
  }
  throw new Error(lastErr);
}

/* ---- "odkey": consumer OneDrive by item, using cid/resid/authkey out of the link. Older
   share links carry these directly and they work anonymously. ---- */
async function viaOdAuthKey(shareUrl, L) {
  const { finalUrl } = await resolveLink(shareUrl, L);
  const q = new URL(finalUrl).searchParams;
  const authkey = q.get("authkey") || q.get("AuthKey");
  const resid = q.get("resid") || q.get("id");
  const cid = q.get("cid") || (resid && resid.includes("!") ? resid.split("!")[0] : null);
  if (!authkey || !resid || !cid) throw new Error(`link has no cid/resid/authkey (found: ${["cid", "resid", "authkey"].filter((k) => q.get(k)).join(",") || "none"})`);
  const base = `https://api.onedrive.com/v1.0/drives/${encodeURIComponent(cid)}/items/${encodeURIComponent(resid)}`;
  const res = await get(`${base}?authkey=${encodeURIComponent(authkey)}&expand=children`, { headers: { Accept: "application/json" } }, L, "list");
  if (!res.ok) throw new Error(`listing failed (${res.status})`);
  const j = await res.json();
  const files = pickFiles(j.children);
  if (!files.length) throw new Error("the folder answered but held no spreadsheets");
  return download(files, L);
}

/* ---- "spshares": business SharePoint / OneDrive-for-Business shares API on the link's own host ---- */
async function viaSpShares(shareUrl, L) {
  const { host } = new URL(shareUrl);
  const base = `https://${host}/_api/v2.0/shares/${encodeShare(shareUrl)}/driveItem`;
  const files = [];
  async function walk(path, depth) {
    const res = await get(`${base}${path}/children?$top=200`, { headers: { Accept: "application/json" } }, L, depth ? "list (subfolder)" : "list");
    if (!res.ok) throw new Error(`listing failed (${res.status}) ${(await res.text()).slice(0, 160)}`);
    const j = await res.json();
    for (const it of j.value || []) {
      if (it.folder && depth < 2) { try { await walk(`/items/${it.id}`, depth + 1); } catch (e) { L.log("  subfolder skipped:", e.message); } }
    }
    files.push(...pickFiles(j.value));
  }
  await walk("", 0);
  if (!files.length) throw new Error("the folder answered but held no spreadsheets");
  return download(files, L);
}

/* ---- "zip": SharePoint's download.aspx returns a whole folder as one zip ---- */
async function viaZip(shareUrl, L) {
  const u = new URL(shareUrl);
  const m = u.pathname.match(/^\/:[a-z]:\/[a-z]\/((?:personal|sites|teams)\/[^/]+)\/([^/?#]+)/i);
  const candidates = [];
  if (m) candidates.push(`https://${u.host}/${m[1]}/_layouts/15/download.aspx?share=${m[2]}`);
  candidates.push(shareUrl + (shareUrl.includes("?") ? "&" : "?") + "download=1");
  let lastErr = "no candidate address";
  for (const url of candidates) {
    const res = await get(url, { headers: { Accept: "*/*" } }, L, "zip");
    if (!res.ok) { lastErr = `download failed (${res.status})`; continue; }
    const buf = Buffer.from(await res.arrayBuffer());
    if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) { lastErr = `response was not a zip (${buf.length} bytes)`; continue; }
    const out = [];
    for (const [p, data] of Object.entries(unzipSync(new Uint8Array(buf)))) {
      const name = p.split("/").pop();
      if (SHEET_RE.test(name) && !name.startsWith("~$")) out.push({ name, buf: Buffer.from(data) });
    }
    if (out.length) return out;
    lastErr = "the zip held no spreadsheets";
  }
  throw new Error(lastErr);
}

/* ---- "rest": follow the link, then read the folder through SharePoint's REST API ---- */
async function viaRest(shareUrl, L) {
  const { finalUrl, cookie } = await resolveLink(shareUrl, L);
  const fu = new URL(finalUrl);
  let id = fu.searchParams.get("id") || fu.searchParams.get("RootFolder");
  if (!id) throw new Error("the opened page carries no folder path");
  try { id = decodeURIComponent(id); } catch {}
  const siteMatch = id.match(/^(\/(?:personal|sites|teams)\/[^/]+)/);
  if (!siteMatch) throw new Error(`the folder path is not a SharePoint path (starts with "/${(id.split("/").filter(Boolean)[0] || "").slice(0, 2)}…")`);
  const apiBase = `https://${fu.host}${siteMatch[1]}/_api/web`;
  const H = { Cookie: cookie, Accept: "application/json;odata=nometadata" };
  const esc = (s) => encodeURIComponent(s.replace(/'/g, "''"));
  const res = await get(`${apiBase}/GetFolderByServerRelativePath(decodedurl='${esc(id)}')/Files?$select=Name,ServerRelativeUrl`, { headers: H }, L, "list");
  if (!res.ok) throw new Error(`listing failed (${res.status})`);
  const j = await res.json();
  const out = [];
  for (const f of j.value || []) {
    if (!SHEET_RE.test(f.Name) || f.Name.startsWith("~$")) continue;
    const r = await get(`${apiBase}/GetFileByServerRelativePath(decodedurl='${esc(f.ServerRelativeUrl)}')/$value`, { headers: { ...H, Accept: "*/*" } }, L, "download");
    if (!r.ok) throw new Error(`download failed (${r.status})`);
    out.push({ name: f.Name, buf: Buffer.from(await r.arrayBuffer()) });
  }
  if (!out.length) throw new Error("the folder answered but held no spreadsheets");
  return out;
}

const STRATEGIES = { od: viaOdShares, odkey: viaOdAuthKey, spshares: viaSpShares, zip: viaZip, rest: viaRest };

export async function fetchWorkbooks(shareUrl, order, L) {
  describeLink(shareUrl, L);
  const tried = [];
  for (const name of order) {
    const fn = STRATEGIES[name];
    if (!fn) { L.log(`Unknown fetch strategy "${name}" - skipped`); continue; }
    L.log(`Fetch strategy "${name}"`);
    try {
      const files = await fn(shareUrl, L);
      if (files.length) { L.log(`  -> ${files.length} workbook(s)`); return { strategy: name, files }; }
      tried.push(`${name}: no workbooks found`);
    } catch (e) {
      tried.push(`${name}: ${L.redact(e.message)}`);
      L.log("  failed:", e.message);
    }
  }
  const err = new Error("All fetch strategies failed.\n" + tried.map((t) => " - " + t).join("\n"));
  err.tried = tried;
  throw err;
}
