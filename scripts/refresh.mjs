// Pulls every coach workbook from the shared OneDrive folder, parses the monthly sheets,
// encrypts the result and writes live.enc.json (+ history.enc.json when HISTORY_JSON is set).
// Env: SHARE_URL, DATA_PASSPHRASE (required) · HISTORY_JSON · LIVE_FROM (default 2026-08)
//      FETCH_STRATEGIES (default "od,odkey,spshares,zip,rest") · OUT_DIR (default ".") · MIN_AGE_HOURS (default 12)
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import XLSX from "xlsx";
import { fetchWorkbooks, makeLogger } from "./fetch.mjs";
import { parseCoachSheet, sheetNameToPeriod, coachFromFilename } from "./parse.mjs";
import { encryptJSON, decryptJSON } from "./crypto.mjs";

const env = process.env;
const SHARE_URL = (env.SHARE_URL || "").trim();
const PASS = (env.DATA_PASSPHRASE || "").trim();
const OUT = env.OUT_DIR || ".";
const LIVE_FROM = env.LIVE_FROM || "2026-08";
const MIN_AGE_H = Number(env.MIN_AGE_HOURS || 12);
const ORDER = (env.FETCH_STRATEGIES || "od,odkey,spshares,zip,rest").split(",").map((s) => s.trim());

if (!SHARE_URL || !PASS) { console.error("SHARE_URL and DATA_PASSPHRASE secrets must both be set."); process.exit(2); }
const L = makeLogger(SHARE_URL);

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
  const periods = {}, report = [];
  for (const f of files) {
    const coach = coachFromFilename(f.name);
    const entry = { file: f.name, coach, months: [], empty: [], errors: [], warnings: [] };
    report.push(entry);
    let wb;
    try { wb = XLSX.read(f.buf, { type: "buffer", cellDates: true }); }
    catch { entry.errors.push("could not read this file as Excel"); continue; }
    const targets = wb.SheetNames.map((n) => ({ n, p: sheetNameToPeriod(n) })).filter((x) => x.p && x.p >= LIVE_FROM && x.p <= now);
    if (!targets.length) entry.errors.push(`no month sheet from ${LIVE_FROM} onward`);
    for (const t of targets) {
      const arr = XLSX.utils.sheet_to_json(wb.Sheets[t.n], { header: 1, raw: true, defval: null });
      const r = parseCoachSheet(arr);
      if (r.error) { entry.errors.push(`${t.p}: ${r.error}`); continue; }
      r.warnings.forEach((w) => entry.warnings.push(`${t.p}: ${w}`));
      if (!(r.sessions || r.revenue || r.cancelled)) { entry.empty.push(t.p); continue; } // month sheet exists but nothing entered yet
      (periods[t.p] = periods[t.p] || {})[coach] = { s: r.sessions, r: r.revenue, x: r.cancelled };
      entry.months.push(t.p);
    }
  }
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
