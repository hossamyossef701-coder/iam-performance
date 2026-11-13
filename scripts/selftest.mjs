// Runs before every refresh. These are not anybody's real sheets: each case is a LAYOUT
// FAMILY the parser must handle, written from scratch. If a change to the parser breaks one,
// the job stops here rather than publishing figures nobody has checked.
import { buildClientsModel } from "./clients-model.mjs";
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
        { s: 6, r: 1700, x: null, p: { s: 20, r: 5600, clients: 2, active: 2, missing: 0, priced: false } });
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
        { s: 16, r: 4500, clients: 2, active: 2, missing: 0, priced: true });
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

/* 9g2. Projected sessions typed but the Projection (AED) formulas missing: priced instead of 0. */
{
  const sheet = [
    row(PROJ_HEAD),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200], [38, 10]]),
    row([[1, "Ben"], [34, 250], [35, 2], [36, 500], [38, 6]]),
  ];
  const r = parseCoachSheet(sheet);
  check("an empty Projection (AED) column falls back to sessions x price",
        { r: r.projection.r, priced: r.projection.priced, flagged: r.warnings.some((w) => /column is empty/.test(w)) },
        { r: 4500, priced: true, flagged: true });
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
        { s: 0, r: 0, clients: 0, active: 1, missing: 1, priced: false });
}

/* 9j. Partly filled: three clients have sessions, one of them has nothing projected. */
{
  const sheet = [
    row(PROJ_HEAD),
    row([[1, "Ann"], [34, 300], [35, 4], [36, 1200], [38, 12], [39, 3600]]),
    row([[1, "Ben"], [34, 250], [35, 2], [36, 500], [38, 8], [39, 2000]]),
    row([[1, "New client"], [34, 280], [35, 1], [36, 280]]),
    row([[1, "Old client"], [34, 280], [35, 0], [36, 0]]),            // no sessions: not counted either way
  ];
  const p = parseCoachSheet(sheet).projection;
  check("a client with sessions but no projection is counted", { active: p.active, missing: p.missing }, { active: 3, missing: 1 });
}

/* 9k. The Clients page: every status and renewal decision worked out from the typed tabs alone.
   One made-up client per rule. Day numbers are Excel's (D(y, m, d) below). */
{
  const D = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  const settings = [["Settings"], [], ["Grace period (days)", 90], ["Go-live date", D(2026, 11, 4)], ["Carry-over period (days)", 120],
    ["Every purchase logged since", D(2026, 1, 1)], ["Expiring soon window (days)", 10], [],
    ["Sessions", "Valid for (days)", "Applies to purchases from"], [12, 49, D(2020, 1, 1)], [24, 84, D(2020, 1, 1)], [12, 60, D(2027, 1, 1)]];
  const clients = [[], ["Client ID", "Full name", "Coach", "Carried-over status", "End date (if they told us)"],
    ["C0001", "Running", "Sue"], ["C0002", "Ran out in September", "Sue"], ["C0003", "Ran out in June", "Ann"],
    ["C0004", "Extended", "Ann"], ["C0005", "Told us", "Ann", null, D(2026, 10, 15)], ["C0006", "Moved over", "Sue", "Active"],
    ["C0007", "Family member", "Sue"], ["C0008", "Family payer", "Sue"], ["C0009", "Renewed then left", "Ann"], ["C0010", "Ends soon", "Ann"]];
  const packages = [[], ["Date", "Client ID", "Size", "Unit", "Coach", "Extended until", "Shared with (Client IDs)"],
    [D(2026, 11, 1), "C0001", 12, "Sessions", "Sue"],
    [D(2026, 8, 1), "C0002", 12, "Sessions", "Sue"],
    [D(2026, 5, 1), "C0003", 12, "Sessions", "Ann"],
    [D(2026, 9, 1), "C0004", 12, "Sessions", "Ann", D(2026, 12, 1)],
    [D(2026, 9, 1), "C0005", 24, "Sessions", "Ann"],
    [D(2026, 11, 5), "C0008", 24, "Sessions", "Sue", null, "C0007"],
    [D(2026, 3, 1), "C0009", 12, "Sessions", "Ann"],
    [D(2026, 4, 15), "C0009", 12, "Sessions", "Ann"],
    [D(2026, 9, 25), "C0010", 12, "Sessions", "Ann"]];
  const sheets = { Settings: settings, Clients: clients, Packages: packages };
  const m = buildClientsModel(sheets, { today: D(2026, 11, 11), debug: true });
  const count = (o) => m.debug.decisions.filter((x) => x[1] === o).length;
  check("clients: who is active (running, extended, moved over, family, payer, ends soon)", m.kpi.active, 6);
  check("clients: renewal due", m.debug.due, ["C0002"]);
  check("clients: ending in 30 days is cover from a package, never the move from the old lists", m.debug.soon, ["C0004", "C0010"]);
  check("clients: renewal decisions", { kept: count("kept"), lost: count("lost"), pending: count("pending") }, { kept: 1, lost: 3, pending: 1 });
  const { debug, ...page } = m;
  check("clients: no names and no Client IDs in what the page receives", /Running|Ran out in|Family|C00\d\d/.test(JSON.stringify(page)), false);
  check("clients: every month adds up (start + new + came back - ran out = end)",
        m.months.every((x) => x.start + x.nw + x.back - x.out === x.end), true);
  // A new rule for 12 sessions from 1 Jan 2027: a package bought after it runs 60 days, one before it 49.
  packages.push([D(2027, 1, 2), "C0001", 12, "Sessions", "Sue"], [D(2026, 12, 30), "C0002", 12, "Sessions", "Sue"]);
  const later = buildClientsModel(sheets, { today: D(2027, 2, 25), debug: true });
  check("clients: a dated rule applies only to purchases from its date", { due: later.debug.due, soon: later.debug.soon },
        // C0002 (bought before: 49 days) has run out; C0001 (bought after: 60 days) is about to.
        // The extension, the family package and C0010 have all run out by then too.
        { due: ["C0002", "C0004", "C0007", "C0008"], soon: ["C0001"] });
  // Amounts are without VAT: a price typed as 1,050 counts as 1,000.
  packages.push([D(2027, 2, 20), "C0009", 12, "Sessions", "Ann", null, null, 1050]);
  packages[1].push("Price (AED incl. VAT)");
  const vat = buildClientsModel(sheets, { today: D(2027, 2, 25) });
  check("clients: package sales are without VAT", vat.kpi.salesMonth, 1000);
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
