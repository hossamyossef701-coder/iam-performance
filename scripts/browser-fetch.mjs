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

// Walks any JSON the page received and picks out spreadsheet entries that carry a download
// address. The shape of these answers changes between OneDrive's own versions, so this looks
// for the fields rather than assuming a layout.
function collect(node, out, folders, folderName, dupes, depth = 0) {
  if (!node || depth > 8) return;
  if (Array.isArray(node)) { for (const n of node) collect(n, out, folders, folderName, dupes, depth + 1); return; }
  if (typeof node !== "object") return;
  const name = typeof node.name === "string" ? node.name : (typeof node.Name === "string" ? node.Name : null);
  if (name && SHEET_RE.test(name) && !name.startsWith("~$")) {
    const url = DL_FIELDS.map((f) => node[f]).find((v) => typeof v === "string" && /^https?:/.test(v));
    // The first sighting wins, so a file in the main folder is never replaced by one of the
    // same name inside a subfolder.
    if (url) { if (out.has(name)) dupes.add(name); else out.set(name, { url, folder: folderName }); }
  }
  // Subfolders - "Inactive" and the like - are followed, because coaches who have left can
  // still hold months nobody has reported anywhere else.
  if (folders && name && node.folder && typeof node.folder === "object" && node.id && !folders.has(name)) {
    folders.set(name, node.id);
  }
  for (const k of Object.keys(node)) collect(node[k], out, folders, folderName, dupes, depth + 1);
}

export async function viaBrowser(shareUrl, L) {
  let chromium;
  try { ({ chromium } = await import("playwright")); }
  catch { throw new Error("the browser component isn't installed in this job"); }

  const browser = await chromium.launch({
    executablePath: process.env.PLAYWRIGHT_CHROMIUM_PATH || undefined,   // set only when a specific build must be used
    args: ["--no-sandbox", "--disable-dev-shm-usage"],
  });
  try {
    const ctx = await browser.newContext({ userAgent: UA, viewport: { width: 1400, height: 1000 }, locale: "en-GB" });
    const page = await ctx.newPage();
    const found = new Map();
    const folders = new Map();
    const dupeNames = new Set();   // same file name seen again inside a subfolder
    let scanning = null;          // the subfolder currently open, or null for the main folder
    let answers = 0;

    page.on("response", async (res) => {
      try {
        if (!/json/i.test(res.headers()["content-type"] || "")) return;
        if (!/children|driveItem|items|listing|RenderListData|GetFolder|files/i.test(res.url())) return;
        const j = await res.json().catch(() => null);
        if (!j) return;
        answers++;
        const before = found.size;
        collect(j, found, folders, scanning, dupeNames);
        if (found.size > before) L.log(`  listed ${found.size} spreadsheet(s)${scanning ? " (in a subfolder)" : ""}`);
      } catch {}
    });

    L.log("  opening the folder in a headless browser");
    await page.goto(shareUrl, { waitUntil: "domcontentloaded", timeout: 90000 });
    try { await page.waitForLoadState("networkidle", { timeout: 30000 }); } catch {}

    // Nudge the list so lazily-loaded rows are requested too.
    for (let i = 0; i < 25 && !found.size; i++) {
      await page.keyboard.press("End").catch(() => {});
      await page.mouse.wheel(0, 1200).catch(() => {});
      await page.waitForTimeout(1000);
    }
    for (let i = 0; i < 6; i++) { await page.mouse.wheel(0, 2000).catch(() => {}); await page.waitForTimeout(700); }

    // Now the subfolders, one at a time. Navigating by the folder's own id is how the site
    // itself moves between folders; if that fails the folder is skipped rather than the run.
    // Every subfolder, and the subfolders inside them, to a fixed depth. A coach's file
    // filed two levels down ("Inactive/2022") was invisible before this.
    const MAX_FOLDERS = Number(process.env.MAX_SUBFOLDERS || 40), MAX_DEPTH = 3;
    const done = new Set();
    let opened = 0;
    for (let depth = 1; depth <= MAX_DEPTH; depth++) {
      const queue = [...folders.entries()].filter(([n]) => !done.has(n));
      if (!queue.length) break;
      L.log(`  ${queue.length} subfolder(s) to scan at level ${depth}`);
      for (const [fname, fid] of queue) {
        done.add(fname);
        if (++opened > MAX_FOLDERS) { L.log(`  stopping after ${MAX_FOLDERS} subfolders`); break; }
        try {
          scanning = fname;
          const before = found.size;
          const u = new URL(shareUrl);
          u.searchParams.set("id", fid);
          await page.goto(u.toString(), { waitUntil: "domcontentloaded", timeout: 60000 });
          try { await page.waitForLoadState("networkidle", { timeout: 25000 }); } catch {}
          for (let i = 0; i < 10 && found.size === before; i++) {
            await page.mouse.wheel(0, 1200).catch(() => {});
            await page.waitForTimeout(800);
          }
          L.log(`  "${fname}": ${found.size - before} new spreadsheet(s)`);
        } catch (e) {
          L.log(`  "${fname}" could not be opened and was skipped: ${e.message.slice(0, 60)}`);
        } finally {
          scanning = null;
        }
      }
      if (opened > MAX_FOLDERS) break;
    }

    if (!found.size) {
      const title = (await page.title().catch(() => "")) || "";
      const signin = await page.locator("input[type=password], #i0116, [data-task=signin]").count().catch(() => 0);
      throw new Error(signin
        ? "the page asked for a sign-in, so the folder is not shared with everyone"
        : `the page opened but listed no spreadsheets (${answers} data answers seen, page title "${title.slice(0, 40)}")`);
    }

    const out = [];
    for (const [name, info] of found) {
      const res = await ctx.request.get(info.url, { headers: { Accept: "*/*" }, timeout: 120000 });
      if (!res.ok()) { L.log(`  download failed (${res.status()}) - file skipped`); continue; }
      out.push({ name, folder: info.folder || null, buf: Buffer.from(await res.body()) });
    }
    if (dupeNames.size) L.log(`  ${dupeNames.size} subfolder file(s) share a name with one in the main folder and were left out`);
    L.log(`  downloaded ${out.length} of ${found.size} file(s)`);
    if (!out.length) throw new Error("none of the listed files could be downloaded");
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
