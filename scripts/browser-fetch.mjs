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
        collect(j, found, scanning ? null : folders, scanning, dupeNames);
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
    const subfolders = [...folders.entries()].slice(0, 8);
    if (subfolders.length) L.log(`  ${subfolders.length} subfolder(s) to scan`);
    for (const [fname, fid] of subfolders) {
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
        L.log(`  subfolder scanned: ${found.size - before} new spreadsheet(s)`);
      } catch (e) {
        L.log(`  a subfolder could not be opened and was skipped: ${e.message.slice(0, 60)}`);
      } finally {
        scanning = null;
      }
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
