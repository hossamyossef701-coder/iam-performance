// Pulls every coach workbook from the shared OneDrive folder, parses the monthly sheets,
// encrypts the result and writes live.enc.json (+ history.enc.json when HISTORY_JSON is set).
// Env: SHARE_URL, DATA_PASSPHRASE (required) · HISTORY_JSON
//      LIVE_FROM (default "1900-01": every month the files hold) · COACH_FILE_PATTERN (default "performance sheet") · FETCH_STRATEGIES (default "od,odkey,spshares,zip,rest,browser") · OUT_DIR (default ".") · MIN_AGE_HOURS (default 12)
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import XLSX from "xlsx";
import { fetchWorkbooks, makeLogger } from "./fetch.mjs";
import { parseCoachSheet, probeSheet, sheetNameToPeriod, coachFromFilename } from "./parse.mjs";
import { encryptJSON, decryptJSON } from "./crypto.mjs";

const env = process.env;
const SHARE_URL = (env.SHARE_URL || "").trim();
const PASS = (env.DATA_PASSPHRASE || "").trim();
const OUT = env.OUT_DIR || ".";
// Every month tab in every workbook, as far back as each file goes. Set a LIVE_FROM
// secret ("2026-08") to read only from a given month instead.
const LIVE_FROM = env.LIVE_FROM || "1900-01";
const MIN_AGE_H = Number(env.MIN_AGE_HOURS || 12);
const ORDER = (env.FETCH_STRATEGIES || "od,odkey,spshares,zip,rest,browser").split(",").map((s) => s.trim());

if (!SHARE_URL || !PASS) { console.error("SHARE_URL and DATA_PASSPHRASE secrets must both be set."); process.exit(2); }
const L = makeLogger(SHARE_URL);
console.log("refresh build 2026-10-22e (subfolders + duplicates)");  // proves which version the job is running

function dubaiPeriod() {
  const p = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Dubai", year: "numeric", month: "2-digit" }).format(new Date());
  return p.slice(0, 7);
}

async function main() {
  // 1. History file (one-off): encrypt the private history JSON held in a secret.
  if (env.HISTORY_JSON && env.HISTORY_JSON.trim()) {
    let h; try { h = JSON.parse(env.HISTORY_JSON); } catch { console.error("HISTORY_JSON secret is not valid JSON."); process.exit(2); }
    const hp = path.join(OUT, "history.enc.json");
    let same = false;
    if (fs.existsSync(hp)) { // re-encrypting identical content would change the file (random IV) and cause a pointless commit
      try { same = JSON.stringify(await decryptJSON(JSON.parse(fs.readFileSync(hp, "utf8")), PASS)) === JSON.stringify(h); } catch {}
    }
    if (same) L.log("history.enc.json already up to date");
    else { fs.writeFileSync(hp, JSON.stringify(await encryptJSON(h, PASS))); L.log("history.enc.json written"); }
  } else if (!fs.existsSync(path.join(OUT, "history.enc.json"))) {
    console.warn("Warning: no HISTORY_JSON secret and no history.enc.json yet - the dashboard will have no history.");
  }

  // 2. Fetch + parse.
  const { strategy, files } = await fetchWorkbooks(SHARE_URL, ORDER, L);
  const now = dubaiPeriod();
  let probesLeft = Number(env.PROBE_LIMIT || 20);   // describes the first few unreadable sheets
  const recentFrom = (function () {              // two months back, for probe reporting only
    const [y, m] = now.split("-").map(Number);
    const d = new Date(Date.UTC(y, m - 3, 1));
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
  })();
  const periods = {}, report = [];
  // Only the per-coach performance workbooks. Anything else in the folder (consolidated
  // finance files and so on) is listed as skipped rather than reported as broken.
  const COACH_FILE = new RegExp(env.COACH_FILE_PATTERN || "performance sheet", "i");
  const skipped = files.filter((f) => !COACH_FILE.test(f.name));
  if (skipped.length) L.log(`Skipped ${skipped.length} file(s) that are not coach performance sheets.`);
  // Files in the main folder are read first, so that where the same coach also has a file in
  // a subfolder ("Inactive"), the main one wins for any month they both hold. The subfolder
  // copy still fills in months the main file does not cover.
  const coachFiles = files.filter((x) => COACH_FILE.test(x.name))
    .sort((a, b) => (a.folder ? 1 : 0) - (b.folder ? 1 : 0));
  // A coach with a file in the main folder is read from there and nowhere else: their
  // subfolder copy is ignored outright, not merged. Only coaches who exist solely in a
  // subfolder - the ones who have left - are read from it.
  const mainCoaches = new Set(coachFiles.filter((f) => !f.folder).map((f) => coachFromFilename(f.name)));
  for (const f of coachFiles) {
    const coach = coachFromFilename(f.name);
    if (f.folder && mainCoaches.has(coach)) {
      report.push({ file: f.name, coach, folder: f.folder, months: [], empty: [], errors: [], warnings: [], superseded: [], ignoredDuplicate: true });
      continue;
    }
    const entry = { file: f.name, coach, folder: f.folder || null, months: [], empty: [], errors: [], warnings: [], superseded: [] };
    report.push(entry);
    let wb;
    try { wb = XLSX.read(f.buf, { type: "buffer", cellDates: true }); }
    catch { entry.errors.push("could not read this file as Excel"); continue; }
    const targets = wb.SheetNames.map((n) => ({ n, p: sheetNameToPeriod(n) })).filter((x) => x.p && x.p >= LIVE_FROM && x.p <= now);
    if (!targets.length) entry.errors.push(`no month sheet from ${LIVE_FROM} onward`);
    for (const t of targets) {
      const arr = XLSX.utils.sheet_to_json(wb.Sheets[t.n], { header: 1, raw: true, defval: null });
      const r = parseCoachSheet(arr);
      if (r.error) {
        entry.errors.push(`${t.p}: ${r.error}`);
        // Describe the layout so it can be taught to the parser. Column headers and rows
        // labelled like totals only - a client row can never appear here.
        // Only recent months are described; an old tab that will never be revisited is noise.
        if (probesLeft > 0 && t.p >= recentFrom) {
          probesLeft--;
          const pr = probeSheet(arr);
          L.log(`  layout of ${coach} / ${t.n} (${pr.rows} rows):`);
          if (pr.headerRowFull) L.log("    headers " + pr.headerRowFull);
          else pr.headerRows.forEach((h) => L.log("    " + h));
          L.log("    labelled rows: " + (pr.labelled.join(" ; ") || "(none)"));
        }
        continue;
      }
      r.warnings.forEach((w) => entry.warnings.push(`${t.p}: ${w}`));
      if (!(r.sessions || r.revenue || r.cancelled)) { entry.empty.push(t.p); continue; } // month sheet exists but nothing entered yet
      const bucket = (periods[t.p] = periods[t.p] || {});
      if (bucket[coach]) { entry.superseded.push(t.p); continue; }   // already taken from the main folder
      bucket[coach] = { s: r.sessions, r: r.revenue, x: r.cancelled };
      entry.months.push(t.p);
    }
  }
  const fromSub = report.filter((r) => r.folder && r.months.length);
  if (fromSub.length) L.log(`${fromSub.length} file(s) in subfolders contributed months: ${fromSub.map((r) => r.coach + " (" + r.months.length + ")").join(", ")}`);
  const dupes = report.filter((r) => r.ignoredDuplicate);
  if (dupes.length) L.log(`${dupes.length} subfolder file(s) ignored because that coach already has a file in the main folder: ${dupes.map((r) => r.coach).join(", ")}`);
  const okCount = report.filter((r) => r.months.length).length;
  L.log(`Parsed ${files.length} workbook(s): ${okCount} with data, ${report.filter((r) => r.errors.length).length} with problems, ${Object.keys(periods).length} month(s).`);
  if (!okCount) { console.error("No coach data could be parsed from any workbook - not overwriting live data."); process.exit(1); }

  // 3. Skip the write when nothing changed and the last write is recent (avoids commit noise).
  const livePath = path.join(OUT, "live.enc.json");
  const digest = crypto.createHash("sha256").update(JSON.stringify({ periods, report })).digest("hex");
  if (fs.existsSync(livePath)) {
    try {
      const prev = await decryptJSON(JSON.parse(fs.readFileSync(livePath, "utf8")), PASS);
      const ageH = (Date.now() - Date.parse(prev.generated)) / 36e5;
      if (prev.digest === digest && ageH < MIN_AGE_H) { L.log(`No change (last write ${ageH.toFixed(1)}h ago) - live.enc.json left as is.`); return; }
    } catch { L.log("Previous live.enc.json could not be read (new passphrase?) - rewriting."); }
  }
  const live = { v: 1, generated: new Date().toISOString(), digest, strategy, liveFrom: LIVE_FROM, periods, files: report };
  fs.writeFileSync(livePath, JSON.stringify(await encryptJSON(live, PASS)));
  L.log("live.enc.json written");
}

main().catch((e) => { console.error(L.redact(e.stack || e.message)); process.exit(1); });
