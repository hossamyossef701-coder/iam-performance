// Pulls every coach workbook from the shared OneDrive folder, parses the monthly sheets,
// encrypts the result and writes live.enc.json (+ history.enc.json when HISTORY_JSON is set).
// Env: SHARE_URL, DATA_PASSPHRASE (required) · HISTORY_JSON
//      ASSESS_SHARE_URL (the assessment tracker) · LIVE_FROM (default "1900-01": every month the files hold) · COACH_FILE_PATTERN (default "performance sheet") · FETCH_STRATEGIES (default "od,odkey,spshares,zip,rest,browser") · OUT_DIR (default ".") · MIN_AGE_HOURS (default 12)
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import XLSX from "xlsx";
import { fetchWorkbooks, fetchOneWorkbook, makeLogger } from "./fetch.mjs";
import { parseCoachSheet, probeSheet, sheetNameToPeriod, coachFromFilename } from "./parse.mjs";
import { parseAssessments } from "./assess.mjs";
import { encryptJSON, decryptJSON } from "./crypto.mjs";

const env = process.env;
const SHARE_URL = (env.SHARE_URL || "").trim();
const PASS = (env.DATA_PASSPHRASE || "").trim();
const OUT = env.OUT_DIR || ".";
// Every month tab in every workbook, as far back as each file goes. Set a LIVE_FROM
// secret ("2026-08") to read only from a given month instead.
const LIVE_FROM = env.LIVE_FROM || "1900-01";
const MIN_AGE_H = Number(env.MIN_AGE_HOURS || 12);
const ASSESS_URL = (env.ASSESS_SHARE_URL || "").trim();
const ORDER = (env.FETCH_STRATEGIES || "od,odkey,spshares,zip,rest,browser").split(",").map((s) => s.trim());

if (!SHARE_URL || !PASS) { console.error("SHARE_URL and DATA_PASSPHRASE secrets must both be set."); process.exit(2); }
const L = makeLogger(SHARE_URL);
console.log("refresh build 2026-10-27c (day row found anywhere)");  // proves which version the job is running

// One typed character of difference, so "Shayla" finds "Shyla".
function near(a, b) {
  if (a === b) return true;
  if (Math.abs(a.length - b.length) > 1) return false;
  let i = 0, j = 0, slips = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) { i++; j++; continue; }
    if (++slips > 1) return false;
    if (a.length > b.length) i++; else if (a.length < b.length) j++; else { i++; j++; }
  }
  return slips + (a.length - i) + (b.length - j) <= 1;
}
// Today's date in Dubai: the cutoff for the month still running.
function dubaiDay() {
  return Number(new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Dubai", day: "numeric" }).format(new Date()));
}
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
  const today = dubaiDay();
  let probesLeft = Number(env.PROBE_LIMIT || 24);   // describes the first few unreadable sheets
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
  // Named, not just counted: a coach's file renamed out of the pattern would otherwise
  // drop them from the dashboard with nothing to show for it.
  if (skipped.length) {
    L.log(`Skipped ${skipped.length} file(s) whose name does not contain "${COACH_FILE.source}":`);
    skipped.forEach((f) => L.log(`  - ${f.name}${f.folder ? " (in " + f.folder + ")" : ""}`));
  }
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
      // For the month still running, also count only the days that have happened, because
      // sessions are booked ahead and the sheet's totals include them.
      const r = parseCoachSheet(arr, t.p === now ? { cutoffDay: today } : null);
      if (r.error) {
        entry.errors.push(`${t.p}: ${r.error}`);
        // Describe the layout so it can be taught to the parser. Column headers and rows
        // labelled like totals only - a client row can never appear here.
        // Only recent months are described; an old tab that will never be revisited is noise.
        if (probesLeft > 0) {
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
      // ms/mr/mx: the same month counted only to day md. Present for the current month only.
      if (r.mtd) { bucket[coach].ms = r.mtd.sessions; bucket[coach].mr = r.mtd.revenue; bucket[coach].mx = r.mtd.cancelled; bucket[coach].md = r.mtd.day; }
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

  // 3. The assessment tracker, when a link for it is configured. A failure here never
  //    stops the run: the coach figures are the point, assessments are an addition.
  let assess = null, assessReport = null;
  if (!ASSESS_URL) L.log("No ASSESS_SHARE_URL secret is set, so the assessment tracker was not read.");
  if (ASSESS_URL) {
    try {
      L.log("Assessment tracker");
      const AL = makeLogger(ASSESS_URL);
      const buf = await fetchOneWorkbook(ASSESS_URL, AL);
      const wb = XLSX.read(buf, { type: "buffer", cellDates: true });
      const merged = {}; const warnings = []; let counted = 0, rows = 0, sheetsRead = 0;
      for (const name of wb.SheetNames) {
        const arr = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: null });
        const r = parseAssessments(arr);
        if (r.error) { L.log(`  sheet "${name}": ${r.error}`); continue; }
        sheetsRead++; counted += r.counted; rows += r.rows;
        r.warnings.forEach((w) => warnings.push(`${name}: ${w}`));
        for (const p of Object.keys(r.periods)) {
          const bucket = (merged[p] = merged[p] || {});
          for (const c of Object.keys(r.periods[p])) {
            const cell = (bucket[c] = bucket[c] || { a: 0, c: 0 });
            cell.a += r.periods[p][c].a; cell.c += r.periods[p][c].c;
          }
        }
      }
      if (!sheetsRead) throw new Error("no sheet in the tracker had Coach and Month columns");
      // Match the tracker's spelling of a coach to the one the performance files use
      // ("Shayla" and "Shyla" are the same person).
      const known = [...new Set(Object.values(periods).flatMap((o) => Object.keys(o)))];
      const mapped = {}, unmatched = new Set(), renamed = new Set();
      const sq = (x) => x.toLowerCase().replace(/[^a-z]/g, "");
      for (const p of Object.keys(merged)) {
        const bucket = (mapped[p] = mapped[p] || {});
        for (const c of Object.keys(merged[p])) {
          let target = known.find((k) => sq(k) === sq(c));
          if (!target) target = known.find((k) => near(sq(k), sq(c)));
          if (target && target !== c) renamed.add(`${c} -> ${target}`);
          if (!target) unmatched.add(c);
          const key = target || c;
          const cell = (bucket[key] = bucket[key] || { a: 0, c: 0 });
          cell.a += merged[p][c].a; cell.c += merged[p][c].c;
        }
      }
      if (renamed.size) L.log(`  coach names matched to the performance files: ${[...renamed].join(", ")}`);
      if (unmatched.size) L.log(`  coach names with no performance file: ${[...unmatched].join(", ")}`);
      assess = mapped;
      assessReport = { rows, counted, months: Object.keys(mapped).sort(), warnings, unmatched: [...unmatched] };
      L.log(`  ${counted} completed assessment(s) across ${Object.keys(mapped).length} month(s)`);
      warnings.forEach((w) => L.log("  " + w));
    } catch (e) {
      L.log("  the tracker could not be read: " + L.redact(e.message).split("\n")[0]);
      assessReport = { error: L.redact(e.message).split("\n")[0] };
    }
  }

  // 4. Skip the write when nothing changed and the last write is recent (avoids commit noise).
  const livePath = path.join(OUT, "live.enc.json");
  const digest = crypto.createHash("sha256").update(JSON.stringify({ periods, report, assess, skipped: skipped.map((f) => f.name) })).digest("hex");
  if (fs.existsSync(livePath)) {
    try {
      const prev = await decryptJSON(JSON.parse(fs.readFileSync(livePath, "utf8")), PASS);
      const ageH = (Date.now() - Date.parse(prev.generated)) / 36e5;
      if (prev.digest === digest && ageH < MIN_AGE_H) { L.log(`No change (last write ${ageH.toFixed(1)}h ago) - live.enc.json left as is.`); return; }
    } catch { L.log("Previous live.enc.json could not be read (new passphrase?) - rewriting."); }
  }
  const live = { v: 1, generated: new Date().toISOString(), digest, strategy, liveFrom: LIVE_FROM, cutoffDay: today, cutoffPeriod: now,
                 periods, files: report, assess, assessReport,
                 skipped: skipped.map((f) => ({ name: f.name, folder: f.folder || null })) };
  fs.writeFileSync(livePath, JSON.stringify(await encryptJSON(live, PASS)));
  L.log("live.enc.json written");
}

main().catch((e) => { console.error(L.redact(e.stack || e.message)); process.exit(1); });
