// Parser for the other department's Assessment Tracker: one row per assessment, with the
// coach, the month, whether the assessment happened, and whether it converted.
//
// Only assessments marked done are counted. A conversion is recognised from a fixed list of
// values (a tick, "Online Training", and so on); anything unfamiliar counts as NOT converted
// and is reported, so a new way of writing it can never quietly inflate the rate.
const norm = (v) => (v == null ? "" : String(v).replace(/\s+/g, " ").trim());
const low = (v) => norm(v).toLowerCase();
const squash = (v) => low(v).replace(/[^a-z0-9]/g, "");

const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

// "Jul-26", "Jul 26", "July 2026", "2026-07", or a real date cell -> "2026-07".
export function monthToPeriod(v) {
  if (v instanceof Date && !isNaN(v)) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}`;
  }
  const t = low(v);
  if (!t) return null;
  let m = t.match(/^(\d{4})-(\d{1,2})$/);
  if (m) return `${m[1]}-${String(+m[2]).padStart(2, "0")}`;
  m = t.match(/^([a-z]{3,9})[\s\-\/.]+(\d{2}|\d{4})$/);
  if (m) {
    const mo = MONTHS[m[1].slice(0, 3)];
    if (!mo) return null;
    const yr = m[2].length === 2 ? 2000 + Number(m[2]) : Number(m[2]);
    return `${yr}-${String(mo).padStart(2, "0")}`;
  }
  return null;
}

// Values in the "Converted" column. Anything outside both lists is treated as not converted
// and reported back, rather than guessed at.
const CONVERTED = new Set(["✓", "✔", "☑", "√", "yes", "y", "converted", "onlinetraining", "online", "ot", "pt", "true", "1"]);
const NOT_CONVERTED = new Set(["×", "x", "✗", "✘", "-", "–", "—", "no", "n", "false", "0", "notconverted", "na", "n/a", "pending", ""]);
function isConverted(cell) {
  const raw = norm(cell);
  if (CONVERTED.has(raw)) return { converted: true };
  const key = squash(raw);
  if (CONVERTED.has(key)) return { converted: true };
  if (NOT_CONVERTED.has(raw) || NOT_CONVERTED.has(key)) return { converted: false };
  // "Online training - 1 month", "online pkg" and the like still read as a conversion.
  if (/online|converted|signed|joined/.test(key)) return { converted: true };
  return { converted: false, unknown: raw };
}

const HEADER_KEYS = {
  coach: ["coach", "coachname", "trainer", "ptcoach"],
  month: ["month", "assessmentmonth", "date"],
  converted: ["converted", "conversion", "convert"],
  status: ["assessmentstatus", "status"],
  client: ["clientsname", "clientname", "clients", "client", "name"],
};
function findHeader(arr) {
  for (let r = 0; r < Math.min(arr.length, 10); r++) {
    const row = arr[r]; if (!row) continue;
    const cols = {};
    for (let c = 0; c < row.length; c++) {
      const k = squash(row[c]); if (!k) continue;
      for (const field of Object.keys(HEADER_KEYS)) {
        if (cols[field] === undefined && HEADER_KEYS[field].includes(k)) cols[field] = c;
      }
    }
    // A row is the header only if it carries the two columns everything depends on.
    if (cols.coach !== undefined && cols.month !== undefined) return { row: r, cols };
  }
  return null;
}

// Returns { periods: { "2026-07": { Shayla: {a,c} } }, rows, counted, warnings }
export function parseAssessments(arr) {
  const head = findHeader(arr);
  if (!head) return { error: "no Coach and Month columns were found — the tracker layout may have changed." };
  const { cols } = head;
  if (cols.converted === undefined) return { error: "no Converted column was found — the tracker layout may have changed." };

  const periods = {};
  const unknownConv = new Map(), unknownStatus = new Map(), badMonths = new Map();
  let rows = 0, counted = 0, skippedStatus = 0, skippedMonth = 0;

  for (let r = head.row + 1; r < arr.length; r++) {
    const row = arr[r]; if (!row) continue;
    const coach = norm(row[cols.coach]);
    const monthCell = row[cols.month];
    if (!coach && !norm(monthCell)) continue;      // a blank line, not a record
    rows++;

    // Only completed assessments count. Scheduled, rescheduled and no-response rows are the
    // other department's pipeline, not work done.
    if (cols.status !== undefined) {
      const st = squash(row[cols.status]);
      if (st !== "done") {
        if (st) unknownStatus.set(st, (unknownStatus.get(st) || 0) + 1);
        skippedStatus++;
        continue;
      }
    }
    const period = monthToPeriod(monthCell);
    if (!period) { skippedMonth++; const k = norm(monthCell); if (k) badMonths.set(k, (badMonths.get(k) || 0) + 1); continue; }
    if (!coach) continue;

    const conv = isConverted(row[cols.converted]);
    if (conv.unknown) unknownConv.set(conv.unknown, (unknownConv.get(conv.unknown) || 0) + 1);

    const bucket = (periods[period] = periods[period] || {});
    const cell = (bucket[coach] = bucket[coach] || { a: 0, c: 0 });
    cell.a++;
    if (conv.converted) cell.c++;
    counted++;
  }

  const warnings = [];
  const top = (m) => [...m.entries()].sort((a, b) => b[1] - a[1]).slice(0, 6).map(([k, n]) => `"${k}" x${n}`).join(", ");
  if (unknownConv.size) warnings.push(`values in Converted that were not recognised and so counted as not converted: ${top(unknownConv)}`);
  if (badMonths.size) warnings.push(`rows skipped because the Month could not be read: ${top(badMonths)}`);
  return { periods, rows, counted, skippedStatus, skippedMonth, statuses: [...unknownStatus.keys()], warnings };
}
