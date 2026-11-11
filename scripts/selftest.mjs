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

/* 8b. Online 1+1 sessions are marked 0.5 - two of them for the price of one in person. Half a
   session counts as half, earns half the rate, and is never treated as an unreadable cell. */
{
  const sheet = [
    row([[0, "Commitment %"], [1, "Client Name"], ...days(2), [34, "PPS"], [36, "TA"]]),
    // two in person, three online halves, one cancelled - 3.5 sessions at 300
    row([[1, "Ann"], [34, 300], [35, 3.5], [36, 1050], [2, 1], [3, 0.5], [4, 0.5], [5, 1], [6, 0.5], [7, 0]]),
  ];
  const r = parseCoachSheet(sheet, { cutoffDay: 27 });
  check("half sessions count as half",
        { s: r.mtd.sessions, r: r.mtd.revenue, x: r.mtd.cancelled },
        { s: 3.5, r: 1050, x: 1 });
  check("...and are not called unreadable",
        r.warnings.some((w) => /other than 0 or 1/.test(w)), false);
}

/* 9c. The totals row is there but its money and session cells are zero, while the client rows
   below hold real figures - Ivan's sheets for sixteen months. A zero there is an empty cell,
   not a total: believing it reported his revenue as nothing. */
{
  const sheet = [
    row([[0, "Commitment %"], [1, "Clients"], ...days(2), [34, "PPS"], [36, "TA"]]),
    row([[1, "Ann"], [34, 250], [35, 8], [36, 2000], ...marks(2, [1, 3, 5, 7, 9, 11, 13, 15])]),
    row([[1, "Ben"], [34, 250], [35, 5], [36, 1250], ...marks(2, [2, 4, 6, 8, 10])]),
    row([[1, "Conducted Sessions / Day"], [35, 0], [36, 0]]),
  ];
  const r = parseCoachSheet(sheet);
  check("zero totals row falls back to the client rows", totals(r), { s: 13, r: 3250, x: null });
  check("...and says so", r.warnings.some((w) => /client rows were added up instead/.test(w)), true);
}

/* 9d. A coach who genuinely did nothing stays at zero - the fallback must not invent figures. */
{
  const sheet = [
    row([[0, "Commitment %"], [1, "Clients"], ...days(2), [34, "PPS"], [36, "TA"]]),
    row([[1, "Ann"], [34, 250], [35, 0], [36, 0]]),
    row([[1, "Conducted Sessions / Day"], [35, 0], [36, 0]]),
  ];
  check("an empty month stays empty", totals(parseCoachSheet(sheet)), { s: 0, r: 0, x: null });
}

/* 9e. The Nov 2026 layout: Projected Sessions and Projection (AED) after TA, Rewards after those,
   the totals row summing both, and the progress % on the row beneath it. The projection is the
   client rows; the % row and the Group Training row add nothing; the rest of the sheet reads as
   before. */
const PROJ_HEAD = [[0, "Commitment %"], [1, "Client Name"], ...days(2, 30), [34, "PPS"], [36, "TA"],
                   [38, "Projected Sessions"], [39, "Projection\n(AED)"], [40, "Rewards (AED)"], [41, "Rwd Session"]];
{
  const sheet = [
    row([[1, "Sun"]]),
    row(PROJ_HEAD),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200], [38, 12], [39, 3600], [40, 60], ...marks(2, [1, 3, 5, 7])]),
    row([[1, "Ben"], [34, 250], [35, 2], [36, 500], [38, 8], [39, 2000], [40, 48], ...marks(2, [2, 4])]),
    row([[1, "Cy"], [34, 280], [35, 0], [36, 0]]),                       // no projection typed: still a client
    row([[1, "Group Training"], [35, 0], [38, 4]]),                       // must not join the projection
    row([[1, "Conducted Sessions / Day"], [35, 6], [36, 1700], [38, 24], [39, 5600]]),
    row([[39, 0.3036]]),                                                  // the progress % row
  ];
  const r = parseCoachSheet(sheet, { cutoffDay: 10 });
  check("projection read from the client rows", { ...totals(r), p: r.projection },
        { s: 6, r: 1700, x: null, p: { s: 20, r: 5600, clients: 2, priced: false } });
  check("...and the Group Training figure is reported, not added",
        r.warnings.some((w) => /projected sessions: sheet total 24, client rows add to 20/.test(w)), true);
}

/* 9f. Projected sessions without the money column: priced at each client's own rate. */
{
  const sheet = [
    row([[0, "Commitment %"], [1, "Clients"], ...days(2, 30), [34, "PPS"], [36, "TA"], [38, "Projected Sessions"]]),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200], [38, 10]]),
    row([[1, "Ben"], [34, 250], [35, 2], [36, 500], [38, 6]]),
  ];
  check("projection priced when the money column is missing", parseCoachSheet(sheet).projection,
        { s: 16, r: 4500, clients: 2, priced: true });
}

/* 9g. A projection in AED that is not projected sessions x price is flagged. */
{
  const sheet = [
    row(PROJ_HEAD),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200], [38, 12], [39, 9999]]),
  ];
  check("an unreconciled projection is flagged",
        parseCoachSheet(sheet).warnings.some((w) => /projection does not reconcile/.test(w)), true);
}

/* 9h. Every earlier layout: no projection columns, no projection, no new warnings. */
{
  const sheet = [
    row([[0, "Commitment %"], [1, "Client Name"], ...days(2), [34, "PPS"], [36, "TA"], [40, "Cancelled Sessions"]]),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200], [40, 1], ...marks(2, [1, 3, 5, 7])]),
    row([[1, "Conducted Sessions / Day"], [35, 4], [36, 1200]]),
  ];
  const r = parseCoachSheet(sheet);
  check("older sheets carry no projection", { p: r.projection, w: r.warnings.length }, { p: null, w: 0 });
}

/* 9i. The columns are there but nothing is typed yet: a projection of zero, not an error. */
{
  const sheet = [
    row(PROJ_HEAD),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200], ...marks(2, [1, 3, 5, 7])]),
  ];
  check("blank projection columns read as zero", parseCoachSheet(sheet).projection,
        { s: 0, r: 0, clients: 0, priced: false });
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
