// Last-resort fetch: open the share link in a real (headless) Chromium, exactly as a person
// would, and let OneDrive's own page do the authenticating. The page asks its API for the
// folder listing; we read that answer as it goes past, then download each workbook through
// the same browser session so its cookies apply.
//
// Loaded only when the plain HTTP strategies have failed, so a normal run never pays for it.
// Nothing here prints the link, a file path or a download address.
const UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36";
const SHEET_RE = /\.(xlsx|xlsm|xls)$/i;
const DL_FIELDS = ["@content.downloadUrl", "@microsoft.graph.downloadUrl", "downloadUrl", "@content.downloadURL"];
// The older pages call a file "name"; the newer ones (personal OneDrive after its 2026 move to
// Microsoft's newer storage) list rows the SharePoint way, as "FileLeafRef", with no download
// address at all - only the item's API address (".spItemUrl") or its path ("FileRef").
const NAME_FIELDS = ["name", "Name", "FileLeafRef", "FileName", "fileName", "LinkFilename"];
const nameOf = (n) => { for (const k of NAME_FIELDS) if (typeof n[k] === "string" && n[k]) return n[k]; return null; };
const isFolderNode = (n) => (n.folder && typeof n.folder === "object") || n.FSObjType === "1" || n.FSObjType === 1;
const isFileNode = (n) => !isFolderNode(n);

// Walks any JSON the page received and keeps every spreadsheet it mentions, with whatever the
// answer offers for getting at it later. src is the request that brought the answer: its
// address and the sign-in headers the page sent with it, reused for the download.
function collect(node, ctx, depth = 0) {
  if (!node || depth > 10) return;
  if (Array.isArray(node)) { for (const n of node) collect(n, ctx, depth + 1); return; }
  if (typeof node !== "object") return;
  const name = nameOf(node);
  if (name && SHEET_RE.test(name) && !name.startsWith("~$") && isFileNode(node)) {
    const url = DL_FIELDS.map((f) => node[f]).find((v) => typeof v === "string" && /^https?:/.test(v)) || null;
    const entry = { url, id: typeof node.id === "string" ? node.id : null,
                    driveId: node.parentReference && node.parentReference.driveId || null,
                    spItemUrl: typeof node[".spItemUrl"] === "string" ? node[".spItemUrl"] : null,
                    fileRef: typeof node.FileRef === "string" ? node.FileRef : null,
                    uniqueId: typeof node.UniqueId === "string" ? node.UniqueId.replace(/[{}]/g, "") : null,
                    folder: ctx.folderName, src: ctx.src };
    const prev = ctx.found.get(name);
    // The first sighting wins, so a main-folder file is never replaced by a subfolder copy.
    // A later sighting can still add a way of downloading the same file.
    if (!prev) ctx.found.set(name, entry);
    else if (prev.folder === entry.folder) { for (const k of Object.keys(entry)) if (prev[k] == null && entry[k] != null) prev[k] = entry[k]; }
    else ctx.dupes.add(name);
  }
  // Subfolders - "Inactive" and the like - are followed, because coaches who have left can
  // still hold months nobody has reported anywhere else. Old pages move between folders by
  // item id, new ones by path.
  if (name && isFolderNode(node) && !ctx.folders.has(name)) {
    const target = typeof node.FileRef === "string" ? node.FileRef : (typeof node.id === "string" ? node.id : null);
    if (target && ctx.folderName == null) ctx.folders.set(name, target);
    else if (target && ctx.folderName != null) ctx.folders.set(`${ctx.folderName}/${name}`, target);
  }
  for (const k of Object.keys(node)) collect(node[k], ctx, depth + 1);
}

// Every way there is of asking for one listed file, best first.
function downloadCandidates(e, pageUrl) {
  const out = [];
  if (e.url) out.push({ how: "address", url: e.url, auth: false });
  if (e.spItemUrl) out.push({ how: "item", url: e.spItemUrl.split("?")[0] + "/content", auth: true });
  const api = e.src && e.src.url && e.src.url.match(/^(https:\/\/[^/]+(?:\/_api)?\/v\d\.\d)\//);
  if (api && e.driveId && e.id) out.push({ how: "drive", url: `${api[1]}/drives/${e.driveId}/items/${e.id}/content`, auth: true });
  const ref = e.fileRef && e.fileRef.match(/^(\/(?:personal|sites|teams)\/[^/]+)\//);
  const origins = [...new Set([e.src && e.src.url, pageUrl].filter(Boolean).map((u) => { try { return new URL(u).origin; } catch { return null; } }).filter(Boolean))];
  if (ref) for (const o of origins) {
    out.push({ how: "path", url: `${o}${ref[1]}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(e.fileRef)}`, auth: false });
    if (e.uniqueId) out.push({ how: "uid", url: `${o}${ref[1]}/_layouts/15/download.aspx?UniqueId=${e.uniqueId}`, auth: false });
  }
  return out;
}

// Describes an address without revealing it: the host, path words that are plainly words, and
// the parameter names. Ids, tokens, names and paths become <n>.
function shapeOf(u) {
  try {
    const x = new URL(u);
    const seg = (p) => (/^(_api|_layouts|_vti_bin|v\d+(\.\d+)?|web|lists|items|children|drives?|root|shares|driveItem|content|RenderListDataAsStream|GetList|GetFolderByServerRelativePath|files|folders|me|personal|Documents|15|[A-Za-z]{2,24}\.aspx)$/i.test(p) ? p : `<${p.length}>`);
    return `${x.host}/${x.pathname.split("/").filter(Boolean).map(seg).join("/")}${[...x.searchParams.keys()].length ? "?" + [...x.searchParams.keys()].join(",") : ""}`;
  } catch { return "<unreadable address>"; }
}
function keysOf(j) {
  const top = Object.keys(j || {}).slice(0, 10).join(",");
  let rows = j && (j.value || j.children || (j.ListData && j.ListData.Row) || j.Row || j.items);
  const first = Array.isArray(rows) && rows[0] && typeof rows[0] === "object" ? Object.keys(rows[0]).filter((k) => !/^\d+$/.test(k)).slice(0, 25).join(",") : "";
  return `keys=[${top}]${first ? ` first row=[${first}]` : ""}${Array.isArray(rows) ? ` rows=${rows.length}` : ""}`;
}
// Answers from Microsoft's own hosts only. OD_TEST_HOST_RE exists for the local stand-in test.
const API_HOST_RE = process.env.OD_TEST_HOST_RE ? new RegExp(process.env.OD_TEST_HOST_RE)
  : /^https:\/\/[^/]*(live\.com|svc\.ms|microsoftpersonalcontent\.com|sharepoint\.com|onedrive\.com|office\.com|office\.net)\//i;
// Only these request headers are carried over to a download: the page's sign-in, nothing else.
const AUTH_HEADERS = /^(authorization|prefer|x-requestdigest|x-forms_based_auth_accepted|application|scenario|x-clientservice-clienttag)$/i;

export async function viaBrowser(shareUrl, L, opts = {}) {
  let chromium;
  try { ({ chromium } = await import("playwright")); }
  catch { throw new Error("the browser component isn't installed in this job"); }

  const browser = await chromium.launch({
    executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH || undefined,   // set only when a specific build must be used
    args: ["--no-sandbox", "--disable-dev-shm-usage"],
  });
  try {
    const bctx = await browser.newContext({ userAgent: UA, viewport: { width: 1400, height: 1000 }, locale: "en-GB" });
    const page = await bctx.newPage();
    const state = { found: new Map(), folders: new Map(), dupes: new Set(), folderName: null, src: null };
    const shapes = [];   // what each data answer looked like - logged only if nothing is found
    let answers = 0;
    const pending = [];

    page.on("response", (res) => {
      pending.push((async () => {
        try {
          if (!/json/i.test(res.headers()["content-type"] || "")) return;
          if (!API_HOST_RE.test(res.url())) return;
          const j = await res.json().catch(() => null);
          if (!j || typeof j !== "object") return;
          answers++;
          let hdrs = {};
          try { const all = await res.request().allHeaders(); for (const k of Object.keys(all)) if (AUTH_HEADERS.test(k)) hdrs[k] = all[k]; } catch {}
          const folderName = state.folderName;
          const before = state.found.size;
          collect(j, { ...state, folderName, src: { url: res.url(), headers: hdrs } });
          if (shapes.length < 14) shapes.push(`${res.request().method()} ${shapeOf(res.url())} ${keysOf(j)}`);
          if (state.found.size > before) L.log(`  listed ${state.found.size} spreadsheet(s)${folderName ? " (in a subfolder)" : ""}`);
        } catch {}
      })());
    });
    const settle = async () => { await Promise.allSettled(pending.splice(0)); };

    L.log("  opening the folder in a headless browser");
    await page.goto(shareUrl, { waitUntil: "domcontentloaded", timeout: 90000 });
    try { await page.waitForLoadState("networkidle", { timeout: 30000 }); } catch {}
    // Nudge the list so lazily-loaded rows are requested too.
    for (let i = 0; i < 25 && !state.found.size; i++) {
      await page.keyboard.press("End").catch(() => {});
      await page.mouse.wheel(0, 1200).catch(() => {});
      await page.waitForTimeout(1000);
      await settle();
    }
    for (let i = 0; i < 6; i++) { await page.mouse.wheel(0, 2000).catch(() => {}); await page.waitForTimeout(700); }
    await settle();
    // Where the link landed: subfolders are opened from here, by changing its "id".
    const landed = page.url();

    // Every subfolder, and the subfolders inside them, to a fixed depth - unless only the
    // folder itself is wanted.
    const MAX_FOLDERS = opts.mainOnly ? 0 : Number(process.env.MAX_SUBFOLDERS || 40), MAX_DEPTH = 3;
    const done = new Set();
    let opened = 0;
    for (let depth = 1; depth <= MAX_DEPTH; depth++) {
      const queue = [...state.folders.entries()].filter(([n]) => !done.has(n));
      if (!queue.length) break;
      L.log(`  ${queue.length} subfolder(s) to scan at level ${depth}`);
      for (const [fname, target] of queue) {
        done.add(fname);
        if (++opened > MAX_FOLDERS) { L.log(`  stopping after ${MAX_FOLDERS} subfolders`); break; }
        try {
          state.folderName = fname;
          const before = state.found.size;
          const u = new URL(landed);
          u.searchParams.set("id", target);
          await page.goto(u.toString(), { waitUntil: "domcontentloaded", timeout: 60000 });
          try { await page.waitForLoadState("networkidle", { timeout: 25000 }); } catch {}
          for (let i = 0; i < 10 && state.found.size === before; i++) {
            await page.mouse.wheel(0, 1200).catch(() => {});
            await page.waitForTimeout(800);
            await settle();
          }
          await settle();
          L.log(`  a subfolder: ${state.found.size - before} new spreadsheet(s)`);
        } catch (e) {
          L.log(`  a subfolder could not be opened and was skipped: ${e.message.slice(0, 60)}`);
        } finally {
          state.folderName = null;
        }
      }
      if (opened > MAX_FOLDERS) break;
    }

    const describe = () => { L.log(`  what the page received (${answers} data answer(s)):`); shapes.forEach((x) => L.log("    " + x)); };
    if (!state.found.size) {
      const title = (await page.title().catch(() => "")) || "";
      const signin = await page.locator("input[type=password], #i0116, [data-task=signin]").count().catch(() => 0);
      describe();
      throw new Error(signin
        ? "the page asked for a sign-in, so the folder is not shared with everyone"
        : `the page opened but listed no spreadsheets (${answers} data answers seen, page title "${title.slice(0, 40)}")`);
    }

    if (opts.only) for (const k of [...state.found.keys()]) if (!opts.only.test(k)) state.found.delete(k);
    if (!state.found.size) throw new Error("the folder opened but holds no file with the expected name");
    const out = [], worked = {}, failedHow = {};
    for (const [name, e] of state.found) {
      let buf = null;
      for (const c of downloadCandidates(e, landed)) {
        try {
          const headers = { Accept: "*/*", ...(c.auth && e.src ? e.src.headers : {}) };
          const res = await bctx.request.get(c.url, { headers, timeout: 120000 });
          if (!res.ok()) { failedHow[c.how] = (failedHow[c.how] || "") + res.status() + " "; continue; }
          const b = Buffer.from(await res.body());
          if (b.length > 4 && b[0] === 0x50 && b[1] === 0x4b) { buf = b; worked[c.how] = (worked[c.how] || 0) + 1; break; }
          failedHow[c.how] = (failedHow[c.how] || "") + "not-a-workbook ";
        } catch { failedHow[c.how] = (failedHow[c.how] || "") + "error "; }
      }
      if (!buf) continue;
      out.push({ name, folder: e.folder || null, buf });
    }
    if (state.dupes.size) L.log(`  ${state.dupes.size} subfolder file(s) share a name with one in the main folder and were left out`);
    L.log(`  downloaded ${out.length} of ${state.found.size} file(s)` +
          (Object.keys(worked).length ? ` - by ${Object.entries(worked).map(([k, v]) => `${k} ${v}`).join(", ")}` : ""));
    if (out.length < state.found.size) {
      L.log(`  download answers that did not work: ${Object.entries(failedHow).map(([k, v]) => `${k}: ${v.trim()}`).join(" | ") || "none tried"}`);
      if (!out.length) describe();
    }
    if (!out.length) throw new Error("files were listed but none could be downloaded");
    return out;
  } finally {
    await browser.close().catch(() => {});
  }
}

// A link to ONE shared workbook rather than a folder. OneDrive opens such a link in Excel
// for the web, so the file has to be asked for explicitly. Three ways are tried: the
// download address, a download triggered in the page, and any spreadsheet the page itself
// fetches while loading.
export async function viaBrowserFile(shareUrl, L) {
  let chromium;
  try { ({ chromium } = await import("playwright")); }
  catch { throw new Error("the browser component isn't installed in this job"); }

  const isXlsx = (buf) => buf && buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b;   // a zip header
  const browser = await chromium.launch({
    executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH || undefined,
    args: ["--no-sandbox", "--disable-dev-shm-usage"],
  });
  try {
    const ctx = await browser.newContext({ userAgent: UA, viewport: { width: 1400, height: 1000 }, locale: "en-GB", acceptDownloads: true });
    const page = await ctx.newPage();
    let sniffed = null;
    page.on("response", async (res) => {
      try {
        if (sniffed) return;
        const ct = res.headers()["content-type"] || "";
        if (!/spreadsheet|octet-stream|excel/i.test(ct)) return;
        const b = Buffer.from(await res.body());
        if (isXlsx(b)) sniffed = b;
      } catch {}
    });

    L.log("  opening the workbook in a headless browser");
    await page.goto(shareUrl, { waitUntil: "domcontentloaded", timeout: 90000 });
    try { await page.waitForLoadState("networkidle", { timeout: 30000 }); } catch {}

    const sep = shareUrl.includes("?") ? "&" : "?";
    const dl = shareUrl + sep + "download=1";

    // 1. ask for the file with the page's own session
    try {
      const res = await ctx.request.get(dl, { headers: { Accept: "*/*" }, timeout: 45000 });
      L.log(`  download address: HTTP ${res.status()}`);
      if (res.ok()) {
        const b = Buffer.from(await res.body());
        if (isXlsx(b)) return b;
        L.log(`  that returned a page, not a workbook (${b.length} bytes)`);
      }
    } catch (e) { L.log("  download address failed: " + e.message.slice(0, 60)); }

    // 2. let the browser handle it as a download
    try {
      const [download] = await Promise.all([
        page.waitForEvent("download", { timeout: 25000 }),
        page.goto(dl, { timeout: 25000 }).catch(() => {}),
      ]);
      const path = await download.path();
      if (path) {
        const fs = await import("node:fs");
        const b = fs.readFileSync(path);
        L.log(`  the browser downloaded ${b.length} bytes`);
        if (isXlsx(b)) return b;
      }
    } catch { L.log("  no download was offered by the page"); }

    // 3. anything spreadsheet-shaped the page fetched on its own
    if (sniffed) { L.log(`  took the workbook from the page's own request (${sniffed.length} bytes)`); return sniffed; }

    const signin = await page.locator("input[type=password], #i0116").count().catch(() => 0);
    throw new Error(signin ? "the link asked for a sign-in" : "the workbook could not be downloaded from this link");
  } finally {
    await browser.close().catch(() => {});
  }
}
