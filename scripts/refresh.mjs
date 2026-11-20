// Pulls every coach workbook from the shared OneDrive folder, parses the monthly sheets,
// encrypts the result and writes live.enc.json (+ history.enc.json when HISTORY_JSON is set).
// Env: SHARE_URL, DATA_PASSPHRASE (required) · HISTORY_JSON
//      FINANCE_PASSPHRASE / CEO_PASSPHRASE / HEADCOACH_PASSPHRASE (optional: the Finance and Overview pages' own files) ·
//      ASSESS_SHARE_URL (the assessment tracker) · CLIENTS_SHARE_URL + CLIENTS_PASSPHRASE (the Clients page) · LIVE_FROM (default "1900-01": every month the files hold) · COACH_FILE_PATTERN (default "performance sheet") · FETCH_STRATEGIES (default "badger,od,odkey,spshares,zip,rest,browser") · OUT_DIR (default ".") · MIN_AGE_HOURS (default 12)
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import XLSX from "xlsx";
import { fetchWorkbooks, fetchOneWorkbook, makeLogger } from "./fetch.mjs";
import { parseCoachSheet, probeSheet, sheetNameToPeriod, coachFromFilename } from "./parse.mjs";
import { parseAssessments } from "./assess.mjs";
import { encryptJSON, decryptJSON } from "./crypto.mjs";
import { buildClientsModel } from "./clients-model.mjs";
import { buildOverview, convWithoutPackage } from "./join.mjs";
import { colourFamily } from "./parse.mjs";
import { guardProjections } from "./projection.mjs";

const COLOUR_NAME = { r: "red (last session)", g: "green (first session)", y: "yellow (charged)", o: "orange (cover)", b: "blue (assessment)", p: "purple (rewarded)" };
const colourLog = {};
// One character per cell, in the same grid as sheet_to_json(header: 1) gives the values: the
// colour family of a solid fill, or "" for none. "?rrggbb" for a fill that fits no family.
function fillGrid(ws) {
  if (!ws || !ws["!ref"]) return [];
  const R = XLSX.utils.decode_range(ws["!ref"]), out = [];
  for (const addr of Object.keys(ws)) {
    if (addr[0] === "!") continue;
    const cell = ws[addr], f = cell && cell.s;
    if (!f || f.patternType !== "solid" || !f.fgColor || !f.fgColor.rgb) continue;
    const a = XLSX.utils.decode_cell(addr), fam = colourFamily(f.fgColor.rgb);
    if (!fam) continue;
    (out[a.r - R.s.r] || (out[a.r - R.s.r] = []))[a.c - R.s.c] = fam === "?" ? "?" + String(f.fgColor.rgb).slice(-6).toUpperCase() : fam;
  }
  return out;
}

const env = process.env;
const SHARE_URL = (env.SHARE_URL || "").trim();
const PASS = (env.DATA_PASSPHRASE || "").trim();
// Changing the main passphrase: put the old one in OLD_DATA_PASSPHRASE for one read. The files
// that are not rebuilt from the sheets each time (history, the projection lock, your projection
// changes) are opened with it and saved again under the new one. Delete the secret afterwards.
const OLD_PASS = (env.OLD_DATA_PASSPHRASE || "").trim();
const OUT = env.OUT_DIR || ".";
// Every month tab in every workbook, as far back as each file goes. Set a LIVE_FROM
// secret ("2026-08") to read only from a given month instead.
const LIVE_FROM = env.LIVE_FROM || "1900-01";
const MIN_AGE_H = Number(env.MIN_AGE_HOURS || 12);
const ASSESS_URL = (env.ASSESS_SHARE_URL || "").trim();
const ORDER = (env.FETCH_STRATEGIES || "badger,od,odkey,spshares,zip,rest,browser").split(",").map((s) => s.trim());

if (!SHARE_URL || !PASS) { console.error("SHARE_URL and DATA_PASSPHRASE secrets must both be set."); process.exit(2); }
const L = makeLogger(SHARE_URL);
console.log("refresh build 2026-11-19a (projection lock, floor and overrides)");  // proves which version the job is running

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

/* The Clients page: "I AM - Clients.xlsx", read on its own link and locked with its own
   passphrase, so the Performance passphrase never opens it and the other way round. It is read
   first and on its own: a problem with it is reported and skipped, never allowed to stop the
   coach figures. Only the five typed tabs are passed on; the file's views are never read. */
// What the Overview and Coach review pages need from the Clients file, kept for later in the run.
let clientsInner = null, nameLinks = [];
// The "Assessments" tab of the Clients file, when it has one: the tracker now lives there.
let assessTab = null;
async function readClients() {
  const url = (env.CLIENTS_SHARE_URL || "").trim(), pass = (env.CLIENTS_PASSPHRASE || "").trim();
  if (!url || !pass) {
    const missing = [!url && "CLIENTS_SHARE_URL", !pass && "CLIENTS_PASSPHRASE"].filter(Boolean).join(" and ");
    if (url || pass) console.log(`::warning title=Clients page not updated::The ${missing} secret is missing (check its name is spelled exactly like that), so the Clients file was not read.`);
    else console.log("No CLIENTS_SHARE_URL / CLIENTS_PASSPHRASE secret, so the Clients page was not updated.");
    return;
  }
  if (pass === PASS) console.log("Note: the Clients passphrase is the same as the Performance one, so either opens both pages.");
  try {
    const LC = makeLogger(url);
    console.log("Reading the Clients file");
    // The link can be to the file itself or to the folder that holds it. From a folder, the file
    // is picked by name ("I AM - Clients", or CLIENTS_FILE_PATTERN), so the old lists and
    // anything else kept beside it are ignored.
    const NAME = new RegExp(env.CLIENTS_FILE_PATTERN || "^I AM\\s*-\\s*Clients\\b[^/]*\\.xlsx$", "i");
    let buf = null;
    const fromFolder = async () => {
      const { files } = await fetchWorkbooks(url, ORDER, LC, { only: NAME, mainOnly: true });
      const hits = files.filter((f) => NAME.test(f.name));
      const best = hits.find((f) => !f.folder && /^I AM\s*-\s*Clients\.xlsx$/i.test(f.name)) || hits.find((f) => !f.folder) || hits[0];
      if (!best) throw new Error("the folder holds no file named like \"I AM - Clients.xlsx\"");
      if (hits.length > 1) console.log(`  ${hits.length} files match the name; using the one in the folder itself`);
      return best.buf;
    };
    if (/\/f\//.test(url)) {
      console.log("  the link is to a folder: looking for the Clients file in it");
      buf = await fromFolder();
    } else {
      try { buf = await fetchOneWorkbook(url, LC, "the Clients file"); }
      catch (e) { console.log("  not readable as a single file - trying it as a folder"); buf = await fromFolder(); }
    }
    const wb = XLSX.read(buf, { type: "buffer" });
    const sheets = {};
    for (const n of ["Clients", "Packages", "Coach Changes", "Lists", "Settings"]) {
      if (wb.Sheets[n]) sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null });
    }
    // The optional "Name links" tab: coach-sheet names confirmed by hand (any spelling of the tab name).
    const linkTab = wb.SheetNames.find((n) => /^name\s*links?$/i.test(n.trim()));
    nameLinks = linkTab ? XLSX.utils.sheet_to_json(wb.Sheets[linkTab], { header: 1, raw: true, defval: null }) : [];
    const aTab = wb.SheetNames.find((n) => /^assessments?(\s*tracker)?$/i.test(n.trim()));
    assessTab = aTab ? XLSX.utils.sheet_to_json(wb.Sheets[aTab], { header: 1, raw: true, defval: null }) : null;
    const model = buildClientsModel(sheets, { internal: true });
    clientsInner = model._in; delete model._in;
    clientsModel = model;
    // Written now, so the Clients page is updated even if the coach sheets fail later in the run.
    // Assessments come from the tracker, read further on: until then, the last ones are kept.
    await writeClients(undefined);
  } catch (e) {
    const msg = e.plain ? e.plain.replace(/\bSHARE_URL/g, "CLIENTS_SHARE_URL").replace(/coach folder/g, "Clients folder")
                        : `The Clients file could not be read: ${String(e.message || e).split("\n")[0].slice(0, 160)}`;
    console.log(`::warning title=Clients page not updated::${msg}`);
    console.log(L.redact(String(e.stack || e.message)));
  }
}

// Writes clients.enc.json. assess: completed and converted assessments per month and coach, for the
// Coaches table (undefined = keep what the file already has). The same figures as last time are
// not written again (a new lock changes the file and makes a pointless commit) unless they are old.
let clientsModel = null;
async function writeClients(assess, rt) {
  if (!clientsModel) return;
  const pass = (env.CLIENTS_PASSPHRASE || "").trim();
  const out = path.join(OUT, "clients.enc.json");
  let prev = null;
  if (fs.existsSync(out)) { try { prev = await decryptJSON(JSON.parse(fs.readFileSync(out, "utf8")), pass); } catch {} }
  const model = { ...clientsModel };
  delete model.generated;
  model.assess = assess !== undefined ? assess : (prev && prev.assess) || null;
  // rt: renewals per coach and month, on time or after the last paid session (from the coach sheets).
  model.rt = rt !== undefined ? rt : (prev && prev.rt) || null;
  if (prev) {
    const { generated: g, ...rest } = prev;
    if (JSON.stringify(rest) === JSON.stringify(model) && Date.now() - Date.parse(g) < MIN_AGE_H * 3600e3) {
      console.log("Clients: same figures as the last read, not rewritten."); return;
    }
  }
  model.generated = new Date().toISOString();
  fs.writeFileSync(out, JSON.stringify(await encryptJSON(model, pass)));
  console.log(`Clients: ${model.kpi.total} clients, ${model.kpi.active} active - clients.enc.json written (no names, no IDs)`);
}

// [sessions per day, revenue per day] for one coach-month, from the client rows' day marks.
function dayTotals(cl) {
  const w = (ch) => (ch === "1" || ch === "x" ? 1 : ch === "h" ? 0.5 : ch >= "2" && ch <= "9" ? +ch : 0);
  const s = [], r = [];
  for (const row of cl) { const m = row[2] || ""; for (let i = 0; i < m.length; i++) { const x = w(m[i]); if (!x) continue; s[i] = (s[i] || 0) + x; r[i] = (r[i] || 0) + x * (row[1] || 0); } }
  const n = Math.max(s.length, r.length);
  for (let i = 0; i < n; i++) { s[i] = s[i] || 0; r[i] = Math.round((r[i] || 0) * 100) / 100; }
  return [s, r];
}

async function rekey() {
  if (!OLD_PASS || OLD_PASS === PASS) return;
  for (const name of ["history.enc.json", "proj-lock.enc.json", "proj-override.enc.json"]) {
    const fp = path.join(OUT, name); if (!fs.existsSync(fp)) continue;
    const e = JSON.parse(fs.readFileSync(fp, "utf8"));
    try { await decryptJSON(e, PASS); continue; } catch {}          // already under the new passphrase
    try { const v = await decryptJSON(e, OLD_PASS); fs.writeFileSync(fp, JSON.stringify(await encryptJSON(v, PASS))); L.log(`${name} moved to the new passphrase`); }
    catch { console.log(`::warning title=Passphrase change::${name} opens with neither the new nor the old passphrase - left as is.`); }
  }
}

async function main() {
  await rekey();
  await readClients();
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
  let fetched;
  try { fetched = await fetchWorkbooks(SHARE_URL, ORDER, L); }
  catch (e) { annotate(e.plain || "The coach folder could not be read - send the run log."); throw e; }
  const { strategy, files } = fetched;
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
  // Day-by-day client rows are kept for the month running and the twelve before it.
  const CL_FROM = (() => { const [y, m] = now.split("-").map(Number); const d = new Date(Date.UTC(y, m - 1 - 12, 1)); return d.toISOString().slice(0, 7); })();
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
    try { wb = XLSX.read(f.buf, { type: "buffer", cellDates: true, cellStyles: true }); }
    catch { entry.errors.push("could not read this file as Excel"); continue; }
    const targets = wb.SheetNames.map((n) => ({ n, p: sheetNameToPeriod(n) })).filter((x) => x.p && x.p >= LIVE_FROM && x.p <= now);
    if (!targets.length) entry.errors.push(`no month sheet from ${LIVE_FROM} onward`);
    for (const t of targets) {
      // One tab that trips the reader is reported against that coach and month; it must never
      // stop the read, or every other coach's figures go unsaved with it.
      let arr, r;
      try {
        arr = XLSX.utils.sheet_to_json(wb.Sheets[t.n], { header: 1, raw: true, defval: null });
        // For the month still running, also count only the days that have happened, because
        // sessions are booked ahead and the sheet's totals include them.
        // The fill colour of every cell, as a family (red = last session of a package, green = first,
        // yellow = a charged cancellation, orange = cover...): the same grid as the values.
        const fills = fillGrid(wb.Sheets[t.n]);
        r = parseCoachSheet(arr, t.p === now ? { cutoffDay: today, fills } : { fills });
      } catch (e) {
        entry.errors.push(`${t.p}: this month's tab could not be read (${String(e && e.message || e).slice(0, 80)})`);
        continue;
      }
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
      // A projection on its own counts as content: on the 1st of the month it is the only thing there.
      const projected = r.projection && (r.projection.s || r.projection.r);
      if (!(r.sessions || r.revenue || r.cancelled || projected)) { entry.empty.push(t.p); continue; } // month sheet exists but nothing entered yet
      const bucket = (periods[t.p] = periods[t.p] || {});
      if (bucket[coach]) { entry.superseded.push(t.p); continue; }   // already taken from the main folder
      bucket[coach] = { s: r.sessions, r: r.revenue, x: r.cancelled };
      // ms/mr/mx: the same month counted only to day md. Present for the current month only.
      if (r.mtd) { bucket[coach].ms = r.mtd.sessions; bucket[coach].mr = r.mtd.revenue; bucket[coach].mx = r.mtd.cancelled; bucket[coach].md = r.mtd.day; }
      // ps/pr: the month's projected sessions and revenue. Only on sheets that carry the columns;
      // a 0 means the columns are there but nothing has been typed in them yet.
      if (r.projection) { bucket[coach].ps = r.projection.s; bucket[coach].pr = r.projection.r; bucket[coach].pm = r.projection.missing; }
      // cl: each client's row day by day, for the last 13 months only. Names as the coach typed
      // them; this file is encrypted with the Performance passphrase like everything in it.
      if (r.people && r.people.length && t.p >= CL_FROM) bucket[coach].cl = r.people.map((x) => [x.n, x.p, x.m, x.k]);
      // How many day cells carried each colour: printed below, so a coach's own shades can be checked.
      if (r.colours && t.p >= CL_FROM) { const cc = colourLog[coach] || (colourLog[coach] = {}); for (const [k, v] of Object.entries(r.colours)) {
        // "other" is itself a tally by fill colour; the rest are counts.
        if (k === "other" && v && typeof v === "object") { const o = cc.other || (cc.other = {}); for (const [rgb, n] of Object.entries(v)) o[rgb] = (o[rgb] || 0) + n; }
        else cc[k] = (cc[k] || 0) + v;
      } }
      entry.months.push(t.p);
    }
  }
  // Colours read from the client rows' day cells, last 13 months, per coach. "other" lists the
  // fills that fit no family, so a shade a coach uses for a session mark can be added.
  for (const [coach, cc] of Object.entries(colourLog)) {
    const { other, ...fam } = cc;
    L.log(`Colours on ${coach}'s sheets: ` + Object.entries(fam).map(([k, v]) => `${COLOUR_NAME[k] || k} ${v}`).join(", "));
  }
  {
    const odd = {}; for (const cc of Object.values(colourLog)) for (const [rgb, n] of Object.entries(cc.other || {})) odd[rgb] = (odd[rgb] || 0) + n;
    const top = Object.entries(odd).sort((a, b) => b[1] - a[1]).slice(0, 12);
    if (top.length) L.log("Fills that fit no colour family: " + top.map(([k, v]) => `#${k} x${v}`).join(", "));
  }
  // Warnings explain a figure that looks odd, so they belong in the log as well as the page.
  report.filter((r) => r.warnings.length).forEach((r) => {
    L.log(`Notes on ${r.coach}:`);
    r.warnings.slice(0, 6).forEach((w) => L.log("  " + w));
    if (r.warnings.length > 6) L.log(`  ...and ${r.warnings.length - 6} more`);
  });
  const fromSub = report.filter((r) => r.folder && r.months.length);
  if (fromSub.length) L.log(`${fromSub.length} file(s) in subfolders contributed months: ${fromSub.map((r) => r.coach + " (" + r.months.length + ")").join(", ")}`);
  const dupes = report.filter((r) => r.ignoredDuplicate);
  if (dupes.length) L.log(`${dupes.length} subfolder file(s) ignored because that coach already has a file in the main folder: ${dupes.map((r) => r.coach).join(", ")}`);
  const okCount = report.filter((r) => r.months.length).length;
  L.log(`Parsed ${files.length} workbook(s): ${okCount} with data, ${report.filter((r) => r.errors.length).length} with problems, ${Object.keys(periods).length} month(s).`);
  if (!okCount) { console.error("No coach data could be parsed from any workbook - not overwriting live data."); process.exit(1); }

  // 2b. The projection that counts: locked on the 3rd, never below the coach's floor, or as the
  //     Fitness Director set it (scripts/projection.mjs). The lock is kept in proj-lock.enc.json;
  //     overrides come from the Team Performance page in proj-override.enc.json.
  {
    const readEnc = async (name) => { try { return await decryptJSON(JSON.parse(fs.readFileSync(path.join(OUT, name), "utf8")), PASS); } catch { return {}; } };
    const lock = await readEnc("proj-lock.enc.json"), overrides = await readEnc("proj-override.enc.json");
    const todayIso = now + "-" + String(today).padStart(2, "0");
    const g = guardProjections(periods, { now, today, todayIso, lock, overrides });
    if (g.lockChanged) { fs.writeFileSync(path.join(OUT, "proj-lock.enc.json"), JSON.stringify(await encryptJSON(g.lock, PASS))); L.log("proj-lock.enc.json written (projections locked)"); }
    L.log(`Projection floor: an honest month falls below ${Math.round(g.dip.s.q * 100)}% of the coach's 3-month average 1 time in 10 (${g.dip.s.n} coach-months of history; revenue ${Math.round(g.dip.r.q * 100)}%)`);
    for (const n of g.notes) L.log("  " + n);
  }

  // 3. The assessment tracker, when a link for it is configured. A failure here never
  //    stops the run: the coach figures are the point, assessments are an addition.
  let assess = null, assessReport = null, assessRows = [];
  // The Assessments tab of the Clients file comes first; the separate tracker file is read only
  // when the Clients file has no such tab (or none of its rows could be used).
  const fromTab = assessTab && assessTab.length ? parseAssessments(assessTab) : null;
  const useTab = !!(fromTab && !fromTab.error && fromTab.rows > 0);
  if (assessTab && !useTab) L.log("The Clients file's Assessments tab " + (fromTab && fromTab.error ? "could not be used: " + fromTab.error : "is empty") + (ASSESS_URL ? " - reading the separate tracker instead." : "."));
  if (!useTab && !ASSESS_URL) L.log("No Assessments tab in the Clients file and no ASSESS_SHARE_URL secret, so assessments were not read.");
  if (useTab || ASSESS_URL) {
    try {
      let sheetsIn;
      if (useTab) { L.log("Assessments: the Clients file's Assessments tab"); sheetsIn = [["Assessments", assessTab]]; }
      else {
        L.log("Assessment tracker (separate file)");
        const AL = makeLogger(ASSESS_URL);
        const buf = await fetchOneWorkbook(ASSESS_URL, AL);
        const wb = XLSX.read(buf, { type: "buffer", cellDates: true });
        sheetsIn = wb.SheetNames.map((name) => [name, XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: null })]);
      }
      const merged = {}; const warnings = []; let counted = 0, rows = 0, sheetsRead = 0;
      for (const [name, arr] of sheetsIn) {
        const r = parseAssessments(arr);
        if (r.error) { L.log(`  sheet "${name}": ${r.error}`); continue; }
        sheetsRead++; counted += r.counted; rows += r.rows; assessRows.push(...r.list);
        r.warnings.forEach((w) => warnings.push(`${name}: ${w}`));
        for (const p of Object.keys(r.periods)) {
          const bucket = (merged[p] = merged[p] || {});
          for (const c of Object.keys(r.periods[p])) {
            const cell = (bucket[c] = bucket[c] || { a: 0, c: 0 });
            cell.a += r.periods[p][c].a; cell.c += r.periods[p][c].c;
          }
        }
      }
      if (!sheetsRead) throw new Error("no sheet of assessments had Coach and Month columns");
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
      { const to = (c) => known.find((k) => sq(k) === sq(c)) || known.find((k) => near(sq(k), sq(c))) || c; assessRows = assessRows.map((x) => ({ ...x, coach: to(x.coach) })); }
      if (renamed.size) L.log(`  coach names matched to the performance files: ${[...renamed].join(", ")}`);
      if (unmatched.size) L.log(`  coach names with no performance file: ${[...unmatched].join(", ")}`);
      assess = mapped;
      assessReport = { rows, counted, months: Object.keys(mapped).sort(), warnings, unmatched: [...unmatched] };
      L.log(`  ${counted} completed assessment(s) across ${Object.keys(mapped).length} month(s)`);
      warnings.forEach((w) => L.log("  " + w));
    } catch (e) {
      L.log("  the assessments could not be read: " + L.redact(e.message).split("\n")[0]);
      assessReport = { error: L.redact(e.message).split("\n")[0] };
    }
  }

  // The Clients page's Coaches table shows each coach's assessments converted: pass them on.
  if (assess) { try { await writeClients(assess); } catch (e) { console.log("Clients: assessments not added - " + String(e.message || e).split("\n")[0]); } }

  // 3b. Overview and Coach review: the coach sheets joined to the Clients file. Locked with the
  //     Performance passphrase. A problem here is reported and never stops the coach figures.
  if (clientsInner) {
    try {
      const revenueOf = (p, c) => {
        const x = periods[p] && periods[p][c]; if (!x) return null;
        // The month running: what has happened so far, not the bookings ahead.
        return p === now && x.mr != null ? { r: x.mr, s: x.ms } : { r: x.r || 0, s: x.s || 0 };
      };
      const ov = buildOverview({ inner: clientsInner, periods, links: nameLinks, now, cutoffDay: today, revenueOf });
      // Converted on the Assessments tab but no package in the Clients file around that month:
      // the conversion counts, the new client does not, until the package is logged.
      ov.convNoPk = convWithoutPackage(assessRows, clientsInner, now);
      if (ov.convNoPk.length) L.log(`Assessments converted with no package logged in the Clients file (last 3 months): ${ov.convNoPk.length} - ${Object.entries(ov.convNoPk.reduce((o, x) => { o[x.coach] = (o[x.coach] || 0) + 1; return o; }, {})).map(([c, n]) => c + " " + n).join(", ")}`);
      const ovPath = path.join(OUT, "overview.enc.json");
      let same = false;
      if (fs.existsSync(ovPath)) {
        try {
          const prev = await decryptJSON(JSON.parse(fs.readFileSync(ovPath, "utf8")), PASS);
          const { generated: g, ...rest } = prev;
          same = JSON.stringify(rest) === JSON.stringify(ov) && Date.now() - Date.parse(g) < MIN_AGE_H * 3600e3;
        } catch {}
      }
      // The Head coach page: the client lists alone, under its own passphrase.
      const hcPass = (env.HEADCOACH_PASSPHRASE || "").trim(), hcPath = path.join(OUT, "roster.enc.json");
      if (!hcPass) L.log("No HEADCOACH_PASSPHRASE secret, so roster.enc.json was not written.");
      else if (hcPass === PASS) console.log("::warning title=Same passphrase::roster.enc.json would use the Performance passphrase - give it its own.");
      else {
        const hcRoster = { ...ov.roster }; delete hcRoster.Hossam;   // the Fitness Director's own clients stay off it
        let hcSame = false;
        try { hcSame = JSON.stringify((await decryptJSON(JSON.parse(fs.readFileSync(hcPath, "utf8")), hcPass)).roster) === JSON.stringify(hcRoster); } catch {}
        if (!hcSame) { fs.writeFileSync(hcPath, JSON.stringify(await encryptJSON({ generated: new Date().toISOString(), roster: hcRoster }, hcPass))); L.log("roster.enc.json written"); }
      }
      if (same) L.log("Overview: same figures as the last read, not rewritten.");
      else {
        ov.generated = new Date().toISOString();
        fs.writeFileSync(ovPath, JSON.stringify(await encryptJSON(ov, PASS)));
        const n = ov.names;
        L.log(`Overview: ${n.auto + n.byHand} coach-sheet names linked (${n.byHand} by the Name links tab), ${n.unclear.length} to confirm, ${n.notInFile.length} not in the file - overview.enc.json written`);
      }
      for (const b of ov.names.bad) L.log("  Name links: " + b);
      { const lk = ov.leak, t = {}; for (const c in lk.m) for (const p in lk.m[c]) for (const f in lk.m[c][p]) { t[f] = t[f] || [0, 0]; t[f][0] += lk.m[c][p][f][0]; t[f][1] += lk.m[c][p][f][1]; }
        const g = (f) => t[f] || [0, 0];
        L.log(`Money on the table (everything followed, ${lk.packages} packages past expiry): ${g("xf")[0]} extension fees not charged, ${g("xc")[0]} charged; ${g("ax")[0]} sessions after expiry with no fee; ${g("ul")[0]} sessions never paid; ${g("fl")[0]} packages used up late, ${g("fo")[0]} on time`); }
      const rc = ov.kpiM.rtCover;
      L.log(`Renewals on time: ${rc.measured} of ${rc.all} renewals followed on the coach sheets, ${rc.red} ended by a red mark; ${rc.misfit} red mark(s) too far from the package size to use`);
      // The Clients page's Coaches table shows renewals on time too (no names: coach and month counts).
      try { await writeClients(undefined, { by: ov.kpiM.rt, cover: rc }); } catch (e) { console.log("Clients: renewals on time not added - " + String(e.message || e).split("\n")[0]); }
    } catch (e) {
      console.log(`::warning title=Overview not updated::${String(e.message || e).split("\n")[0].slice(0, 160)}`);
      console.log(L.redact(String(e.stack || e.message)));
    }
  } else L.log("The Clients file was not read, so the Overview and Coach review pages were not updated.");

  // 3b. The Finance and Overview pages, each with its own passphrase and only what it shows:
  //     coach totals by month (sessions, revenue, cancelled; Overview also the projection). No
  //     client names, no file report. Without the secret, that file is simply not written.
  let histPlain = null;
  try { histPlain = env.HISTORY_JSON && env.HISTORY_JSON.trim() ? JSON.parse(env.HISTORY_JSON) : await decryptJSON(JSON.parse(fs.readFileSync(path.join(OUT, "history.enc.json"), "utf8")), PASS); } catch {}
  const AUD = [
    { file: "finance.enc.json", pass: (env.FINANCE_PASSPHRASE || "").trim(), keys: ["s", "r", "x", "ms", "mr", "mx", "md"] },
    { file: "ceo.enc.json", pass: (env.CEO_PASSPHRASE || "").trim(), keys: ["s", "r", "x", "ms", "mr", "mx", "md", "ps", "pr", "pm", "ps0", "pr0", "pg"] },
  ];
  const writeAudience = async (onlyMissing) => {
    for (const a of AUD) {
      if (!a.pass) { L.log(`No ${a.file === "finance.enc.json" ? "FINANCE" : "CEO"}_PASSPHRASE secret, so ${a.file} was not written.`); continue; }
      if (a.pass === PASS) { console.log(`::warning title=Same passphrase::${a.file} would use the Performance passphrase - give it its own, so it opens that page only.`); continue; }
      const fp = path.join(OUT, a.file);
      if (onlyMissing && fs.existsSync(fp)) continue;
      const per = {};
      for (const [p, cs] of Object.entries(periods)) for (const [c, v] of Object.entries(cs)) {
        const o = {}; for (const k of a.keys) if (v[k] !== undefined) o[k] = v[k];
        // Sessions and revenue per day of the month (no client names), so the comparisons "to the
        // same day" read the same as on Team Performance.
        if (v.cl) o.dd = dayTotals(v.cl);
        (per[p] || (per[p] = {}))[c] = o;
      }
      const H = histPlain ? { HIST: histPlain.HIST || {}, SEED: histPlain.SEED || {} } : { HIST: {}, SEED: {} };
      // Overview also carries the assessment counts per coach and month (no names), for its coach table.
      const extra = a.file === "ceo.enc.json" && assess ? { assess } : {};
      fs.writeFileSync(fp, JSON.stringify(await encryptJSON({ hist: H, live: { v: 1, generated: new Date().toISOString(), cutoffDay: today, cutoffPeriod: now, periods: per, ...extra } }, a.pass)));
      L.log(`${a.file} written`);
    }
  };

  // 4. Skip the write when nothing changed and the last write is recent (avoids commit noise).
  //    A read somebody asked for by hand is different: they pressed the button to find out
  //    where things stand now, so the file is always rewritten and the timestamp always moves.
  //    Without this, "Refresh now" would finish having changed nothing the dashboard can see,
  //    which is indistinguishable from the read having failed.
  const FORCED = env.FORCE_WRITE === "1" || env.GITHUB_EVENT_NAME === "workflow_dispatch";
  const livePath = path.join(OUT, "live.enc.json");
  const digest = crypto.createHash("sha256").update(JSON.stringify({ periods, report, assess, skipped: skipped.map((f) => f.name) })).digest("hex");
  if (fs.existsSync(livePath)) {
    try {
      const prev = await decryptJSON(JSON.parse(fs.readFileSync(livePath, "utf8")), PASS);
      const ageH = (Date.now() - Date.parse(prev.generated)) / 36e5;
      if (prev.digest === digest) {
        if (FORCED) L.log(`The figures are unchanged since the last read ${ageH.toFixed(1)}h ago, but this read was asked for by hand - rewriting so the dashboard can see it ran.`);
        else if (ageH < MIN_AGE_H) { L.log(`No change (last write ${ageH.toFixed(1)}h ago) - live.enc.json left as is.`); await writeAudience(true); return; }
      }
    } catch { L.log("Previous live.enc.json could not be read (new passphrase?) - rewriting."); }
  }
  const live = { v: 1, generated: new Date().toISOString(), digest, strategy, liveFrom: LIVE_FROM, cutoffDay: today, cutoffPeriod: now,
                 periods, files: report, assess, assessReport,
                 skipped: skipped.map((f) => ({ name: f.name, folder: f.folder || null })) };
  fs.writeFileSync(livePath, JSON.stringify(await encryptJSON(live, PASS)));
  L.log("live.enc.json written");
  await writeAudience(false);
}

// A one-line reason pinned to the run, which the dashboard reads and shows. GitHub's own
// "annotation" format; the text never carries the link, a path or a name.
let annotated = false;
function annotate(msg) {
  annotated = true;
  console.log(`::error title=Read failed::${String(msg).replace(/%/g, "%25").replace(/\r?\n/g, " ")}`);
}
main().catch((e) => {
  if (!annotated) annotate("The read stopped on an unexpected error - send the run log.");
  console.error(L.redact(e.stack || e.message)); process.exit(1);
});
