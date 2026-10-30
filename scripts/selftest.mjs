// Runs before every refresh. These are not anybody's real sheets: each case is a LAYOUT
// FAMILY the parser must handle, written from scratch. If a change to the parser breaks one,
// the job stops here rather than publishing figures nobody has checked.
import { parseCoachSheet, sheetNameToPeriod } from "./parse.mjs";
import { parseAssessments, monthToPeriod } from "./assess.mjs";

const W = 48;
const row = (pairs) => { const r = Array(W).fill(null); for (const [i, v] of pairs) r[i] = v; return r; };
const days = (at, n = 31, from = 1) => Array.from({ length: n }, (_, i) => [at + i, from + i]);
const marks = (at, list) => list.map((d) => [at + d - 1, 1]);

let failures = 0;
function check(name, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) { failures++; console.error(`FAIL ${name}\n  expected ${JSON.stringify(want)}\n  got      ${JSON.stringify(got)}`); }
  else console.log(`pass  ${name}`);
}
const totals = (r) => (r.error ? { error: r.error } : { s: r.sessions, r: Math.round(r.revenue), x: r.cancelled });

/* 1. A totals row that is labelled, with PPS/TA headers and a cancelled column. */
{
  const sheet = [
    row([[2, "Mon"]]),
    row([[0, "Commitment %"], [1, "Client Name"], ...days(2), [34, "PPS"], [36, "TA"], [40, "Cancelled Sessions"]]),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200], [40, 1], ...marks(2, [1, 3, 5, 7])]),
    row([[1, "Ben"], [34, 250], [35, 2], [36, 500], [40, 0], ...marks(2, [2, 4])]),
    row([[1, "Group Training"], [35, 9]]),
    row([[1, "Conducted Sessions / Day"], [35, 6], [36, 1700]]),
  ];
  check("labelled totals row", totals(parseCoachSheet(sheet)), { s: 6, r: 1700, x: 1 });
}

/* 2. No totals row at all: the client rows are the total. */
{
  const sheet = [
    row([[0, "Commitment %"], [1, "Clients"], ...days(2), [34, "PPS"], [36, "TA"]]),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200]]),
    row([[1, "Ben"], [34, 250], [35, 2], [36, 500]]),
    row([[1, "Group Classes"], [35, 9]]),
    row([[1, "First Session"]]),
  ];
  check("no totals row", totals(parseCoachSheet(sheet)), { s: 6, r: 1700, x: null });
}

/* 3. Price headed differently and the money column unheaded, two columns along. */
{
  const sheet = [
    row([[0, "CLIENTS"], ...days(1), [33, "Price PS"]]),
    row([[0, "Ann"], [33, 250], [34, 4], [35, 1000], ...marks(1, [2, 4])]),
    row([[0, "GROUP SESSIONS"], [34, 0], [35, 0]]),
  ];
  check("unheaded money column", totals(parseCoachSheet(sheet)), { s: 4, r: 1000, x: null });
}

/* 4. The same, but the money column holds something else: it must refuse, not guess. */
{
  const sheet = [
    row([[0, "CLIENTS"], ...days(1), [33, "Price PS"]]),
    row([[0, "Ann"], [33, 250], [34, 4], [35, 99999], ...marks(1, [2])]),
  ];
  const r = parseCoachSheet(sheet);
  check("refuses an unproven money column", !!r.error, true);
}

/* 5. Typos and stray spaces in every label. */
{
  const sheet = [
    row([[0, "Commitment %"], [1, "Clinet  Name"], ...days(2), [34, "PPS "], [36, "TA"], [40, "Cancelled Session"]]),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200], [40, 2]]),
    row([[1, "Conducted  Sesssions / Day"], [35, 4], [36, 1200]]),
  ];
  check("typos tolerated", totals(parseCoachSheet(sheet)), { s: 4, r: 1200, x: 2 });
}

/* 6. Day numbers on the row above the header, weekdays above those. */
{
  const sheet = [
    row(days(2).map(([c]) => [c, "Mon"])),
    row(days(2)),
    row([[0, "Commitment %"], [1, "Client Name"], [34, "PPS"], [36, "TA"]]),
    row([[1, "Ann"], [34, 300], [35, 3], [36, 900], ...marks(2, [1, 2, 20])]),
    row([[1, "Group Training"], [35, 0]]),
  ];
  const r = parseCoachSheet(sheet, { cutoffDay: 10 });
  check("day row above the header", { cols: r.dayCols, mtd: r.mtd && r.mtd.sessions }, { cols: 31, mtd: 2 });
}

/* 7. Day headers that are real dates rather than numbers. */
{
  const sheet = [
    row(Array.from({ length: 30 }, (_, i) => [i + 2, new Date(2026, 9, i + 1)])),
    row([[0, "Commitment %"], [1, "Client Name"], [34, "PPS"], [36, "TA"]]),
    row([[1, "Ann"], [34, 300], [35, 2], [36, 600], ...marks(2, [1, 15])]),
    row([[1, "Group Training"], [35, 0]]),
  ];
  const r = parseCoachSheet(sheet, { cutoffDay: 10 });
  check("dates as day headers", { cols: r.dayCols, mtd: r.mtd && r.mtd.sessions }, { cols: 30, mtd: 1 });
}

/* 8. Month to date: bookings ahead are excluded, zeros are cancellations. */
{
  const sheet = [
    row([[0, "Commitment %"], [1, "Client Name"], ...days(2), [34, "PPS"], [36, "TA"], [40, "Cancelled Sessions"]]),
    // day 1 held, day 2 cancelled, days 5 and 25 held, days 28 and 30 booked ahead
    row([[1, "Ann"], [34, 300], [35, 5], [36, 1500], [40, 2], [2, 1], [3, 0], [6, 1], [26, 1], [29, 1], [31, 1]]),
    row([[1, "Group Training"], [35, 0]]),
  ];
  const r = parseCoachSheet(sheet, { cutoffDay: 27 });
  check("month to date excludes bookings ahead",
    { s: r.mtd.sessions, r: r.mtd.revenue, x: r.mtd.cancelled, whole: r.sessions },
    { s: 3, r: 900, x: 1, whole: 5 });
}

/* 9. A sheet with no calendar at all: no month-to-date figure, and it says so. */
{
  const sheet = [
    row([[0, "Commitment %"], [1, "Client Name"], [34, "PPS"], [36, "TA"]]),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200]]),
  ];
  const r = parseCoachSheet(sheet, { cutoffDay: 27 });
  check("no calendar leaves no to-date figure", { mtd: r.mtd, warned: r.warnings.some((w) => /day numbers/.test(w)) }, { mtd: null, warned: true });
}

/* 9b. The totals row carries no label, and a salary block sits below the clients. This is the
   April 2026 shape: the client block has no end, so the parser used to walk on into the block
   beneath and read its TOT as a thirteenth client - revenue 63,402 against the sheet's 37,839.
   The rows below the clients must contribute nothing. */
{
  const sheet = [
    row([[0, "Committment %"], [1, "Clients"], ...days(2, 30), [34, "PPS"], [36, "TA"]]),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200], ...marks(2, [1, 3, 5, 7])]),
    row([[1, "Ben"], [34, 250], [35, 2], [36, 500], ...marks(2, [2, 4])]),
    row([[0, 0.846], ...days(2, 30).map(([i]) => [i, 1]), [35, 6], [36, 1700]]),  // totals row, no label
    row([[1, "First Session"], [34, "BS"], [36, 12000]]),
    row([[1, "Cover"], [34, "CM"], [36, 13244]]),
    row([[1, "Rewarded Session"], [34, "TOT"], [36, 25563]]),    // the row that used to slip in
  ];
  check("salary block below the clients is not revenue",
        totals(parseCoachSheet(sheet)), { s: 6, r: 1700, x: null });
}

/* 10. Month tab names. */
check("sheet names", ["Oct 2026", "Sept 2026", "October 2026", " October 2022", "Assessment Report - 2026"].map(sheetNameToPeriod),
  ["2026-10", "2026-09", "2026-10", "2022-10", null]);

/* 11. The assessment tracker. */
{
  const H = ["Client's Name", "Assessment Status", "Month", "Converted", "Sessions Billed", "Coach", "Survey"];
  const r = parseAssessments([H,
    ["A", "Done", "Jul-26", "✓", 6, "Shayla", "Done"],
    ["B", "Done", "Jul-26", "×", "×", "Shayla", "Done"],
    ["C", "Done", "Aug-26", "Online Training", "1Month", "Sue", "Done"],
    ["D", "Scheduled", "Aug-26", "-", "-", "Sue", "-"],
  ]);
  check("assessment tracker", r.periods, { "2026-07": { Shayla: { a: 2, c: 1 } }, "2026-08": { Sue: { a: 1, c: 1 } } });
  check("tracker month forms", ["Jul-26", "July 2026", "2026-07"].map(monthToPeriod), ["2026-07", "2026-07", "2026-07"]);
}

if (failures) { console.error(`\n${failures} self-test(s) failed - refusing to publish figures.`); process.exit(1); }
console.log("\nAll layout self-tests passed.");
