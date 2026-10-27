// Parser for one coach's monthly sheet.
//
// The coaches' files are not identical: one calls the name column "Client Name" and another
// "Clients"; one labels the totals row "Conducted Sessions / Day" and another just
// "Conducted Sessions"; the PPS/TA block sits in a different column in each file, and the
// cancelled total lands on the Group Training row in some files and on an unnamed row in
// others. So nothing is located by a fixed row or column, and every label is matched on its
// meaning rather than its exact spelling.
const MONTH_IDX = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

const norm = (v) => (v == null ? "" : String(v).replace(/\s+/g, " ").trim().toLowerCase());
const num = (v) => {
  if (typeof v === "number") return isFinite(v) ? v : null;
  const n = parseFloat(String(v == null ? "" : v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? null : n;
};

// "Oct 2026", "October 2026", "Sept 2026" -> "2026-10". Anything else -> null.
export function sheetNameToPeriod(name) {
  const m = norm(name).match(/^([a-z]{3,9})\.?\s+(\d{4})$/);
  if (!m) return null;
  const mo = MONTH_IDX[m[1].slice(0, 3)];
  return mo ? `${m[2]}-${String(mo).padStart(2, "0")}` : null;
}

// Coach name = the text before the first "_" or "-" ("Faisal - I AM Performance sheet" -> "Faisal").
export function coachFromFilename(filename) {
  const base = String(filename).replace(/\.(xlsx|xlsm|xls|csv)$/i, "");
  const m = base.match(/^([^_\-]+)/);
  return (m ? m[1] : base).trim();
}

// When a sheet will not parse, describe its STRUCTURE so the layout can be fixed without
// anyone having to send the file.
//
// A client's name must never appear here, so a cell is reported only when EVERY word in it
// belongs to the structural vocabulary below. "Conducted Sessions / Day" passes; a person's
// name cannot, because their name is not in the list.
const SAFE_WORDS = new Set(("commitment committment client clients name names customer customers pps ta rewards reward rwd " +
  "session sessions total totals cancelled canceled cancel conducted conduct revenue wasted covered cover group training " +
  "class classes day days per coach if not for charged first last assessment assessments brought forwarded rewarding " +
  "notes note weekly schedule data report salary lost bs cm bn ded extra tot min max sum rate income pt i am performance " +
  "sheet consolidated vs hours time type status month year and the of to from with in on at by is are no yes new old " +
  "mon tue wed thu fri sat sun monday tuesday wednesday thursday friday saturday sunday " +
  "jan feb mar apr may jun jul aug sep sept oct nov dec").split(" "));
function safeLabel(t) {
  if (!t || t.length > 46) return null;
  const words = t.toLowerCase().split(/[^a-z%]+/).filter((w) => w && w !== "%");
  if (!words.length) return null;                       // pure numbers carry no layout information
  if (!words.every((w) => SAFE_WORDS.has(w))) return null;
  return t;
}
export function probeSheet(arr) {
  const headerRows = [];
  for (let r = 0; r < Math.min(arr.length, 6); r++) {
    const row = arr[r]; if (!row) continue;
    const cells = [];
    for (let c = 0; c < row.length; c++) {
      const t = row[c] == null ? "" : String(row[c]).replace(/\s+/g, " ").trim();
      const safe = safeLabel(t);
      if (safe) cells.push(c + ":" + safe);
    }
    if (cells.length) headerRows.push(`row ${r}: ${cells.slice(0, 16).join(" | ")}`);
  }
  const labelled = [];
  for (let r = 0; r < arr.length; r++) {
    const row = arr[r]; if (!row) continue;
    for (let c = 0; c < Math.min(row.length, 4); c++) {
      const t = row[c] == null ? "" : String(row[c]).replace(/\s+/g, " ").trim();
      const safe = safeLabel(t);
      if (safe && /(session|total|cancel|conduct|group)/i.test(safe)) { labelled.push(`r${r}c${c}: ${safe}`); break; }
    }
  }
  // The row carrying the client-name header holds COLUMN headings, never client rows, so it
  // is reported in full apart from the name column itself — that is how an unexpected set of
  // columns gets diagnosed.
  let headerRowFull = null;
  for (let r = 0; r < Math.min(arr.length, 12) && !headerRowFull; r++) {
    const row = arr[r]; if (!row) continue;
    let nameCol = -1;
    for (let c = 0; c < Math.min(row.length, 8); c++) {
      if (NAME_HEADERS.includes(norm(row[c]))) { nameCol = c; break; }
    }
    if (nameCol < 0) continue;
    const cells = [];
    for (let c = 0; c < row.length; c++) {
      if (c === nameCol) continue;
      const t = row[c] == null ? "" : String(row[c]).replace(/\s+/g, " ").trim();
      if (t) cells.push(c + ":" + t.slice(0, 38));
    }
    headerRowFull = `row ${r} (name col ${nameCol}): ${cells.join(" | ")}`;
  }
  return { rows: arr.length, headerRows, headerRowFull, labelled: labelled.slice(0, 30) };
}

const NAME_HEADERS = ["client name", "clients", "client", "client names", "name", "names", "customer", "customers"];
const NOT_A_CLIENT = /(conduct|cancell?ed|^total|sessions? ?\/ ?day|first session|last session|charged session|penultimate|rewarding|brought forward|assessment|^cover$|^notes?$|^legend$|salary|basic|petrol|salik)/;
const GROUP_ROW = /^group\s*(training|class|classes|session|sessions)/;

// Spelling drifts: a stray space, a plural, a typo. Labels are compared with the spaces and
// punctuation stripped, and one typed character of difference is forgiven.
const squash = (s) => norm(s).replace(/[^a-z0-9]/g, "");
function within1(a, b) {
  if (a === b) return true;
  if (Math.abs(a.length - b.length) > 1) return false;
  // Two letters typed the wrong way round ("Clinet Name") counts as one slip.
  if (a.length === b.length) {
    for (let k = 0; k < a.length - 1; k++) {
      if (a[k] !== b[k]) {
        return a[k] === b[k + 1] && a[k + 1] === b[k] && a.slice(k + 2) === b.slice(k + 2);
      }
    }
  }
  let i = 0, j = 0, slips = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) { i++; j++; continue; }
    if (++slips > 1) return false;
    if (a.length > b.length) i++;
    else if (a.length < b.length) j++;
    else { i++; j++; }
  }
  return slips + (a.length - i) + (b.length - j) <= 1;
}
const matchesAny = (text, options) => { const t = squash(text); return !!t && options.some((o) => within1(t, o)); };

// Coaches book sessions ahead, so the sheet's own monthly totals include days that have not
// happened yet. For the month still running, the figures are rebuilt by counting only the
// day columns up to and including today. opts.cutoffDay says which day that is.
export function parseCoachSheet(arr, opts) {
  // 1. The header row is the one carrying the client-name column; its position and spelling
  //    both vary, so search the opening rows across the opening columns.
  const NAME_KEYS = NAME_HEADERS.map(squash);
  let headerRowIdx = -1, nameCol = -1;
  outer:
  for (let r = 0; r < Math.min(arr.length, 12); r++) {
    const row = arr[r]; if (!row) continue;
    for (let c = 0; c < Math.min(row.length, 8); c++) {
      if (matchesAny(row[c], NAME_KEYS)) { headerRowIdx = r; nameCol = c; break outer; }
    }
  }
  if (headerRowIdx < 0) return { error: "no client-name column was found in the opening rows — the sheet layout may have changed." };

  const headers = arr[headerRowIdx];
  const findCol = (test) => {
    for (let i = 0; i < headers.length; i++) if (test(norm(headers[i]))) return i;
    return -1;
  };
  // The per-session price column is "PPS" in most files and "VALUE" in at least one.
  const ppsCol = findCol((h) => matchesAny(h, ["pps", "value", "rate", "price", "ppc", "sessionvalue"]));
  // The money column is "TA" in most files; in some it carries no header at all, in which
  // case it sits two columns right of the price (price, session count, money).
  const taCol = findCol((h) => matchesAny(h, ["ta", "totalamount", "amount", "revenue", "total"]));
  // "Cancelled Sessions" is sometimes wrapped across two lines inside the cell.
  const cancelledCol = findCol((h) => matchesAny(h, ["cancelledsessions", "canceledsessions", "cancelled", "cancelledsession"]));
  if (ppsCol < 0) return { error: "no per-session price column (PPS / VALUE / Rate) was found — the sheet layout may have changed." };
  const sessionsCol = ppsCol + 1;            // the session count sits right after the price
  // The columns holding the days of the month, taken from the header row's 1..31.
  const dayCols = [];
  for (let c = 0; c < headers.length; c++) {
    if (c === ppsCol || c === sessionsCol) continue;
    const d = num(headers[c]);
    if (d != null && Number.isInteger(d) && d >= 1 && d <= 31) dayCols.push({ c, d });
  }
  const moneyCol = taCol >= 0 ? taCol : ppsCol + 2;   // and the money right after that
  const moneyGuessed = taCol < 0;

  // 2. The totals row: "Conducted Sessions", with or without "/ Day", in or near the name column.
  let conductedRow = -1;
  for (let r = headerRowIdx + 1; r < arr.length; r++) {
    const row = arr[r]; if (!row) continue;
    for (let c = 0; c <= nameCol + 2 && c < row.length; c++) {
      const t = norm(row[c]);
      if (squash(t).startsWith("conducted")) { conductedRow = r; break; }
    }
    if (conductedRow >= 0) break;
  }
  // Some sheets carry no totals row at all - every August file, and a few coaches in every
  // month. There the client rows ARE the total, so they are added up rather than failing.
  const clientEnd = conductedRow >= 0 ? conductedRow : arr.length;

  // 3. Cancelled sessions. The sheets disagree about where the total goes - an unnamed row
  //    in one file, the Group Training row in another, and in some files no total at all -
  //    so the per-client column is added up instead, and the sheet's own figure is only used
  //    to cross-check it below.
  let sheetCancelled = null;
  if (cancelledCol >= 0) {
    for (let r = headerRowIdx + 1; r < clientEnd; r++) {
      const v = arr[r] ? num(arr[r][cancelledCol]) : null;
      if (v != null) sheetCancelled = v;
    }
  }

  // 4. Add up the individual client rows: the cross-check when the sheet has its own total,
  //    and the figure itself when it has not. Group training is counted separately by the
  //    sheet, so it ends the client block.
  let sumS = 0, sumR = 0, sumC = 0, clients = 0;
  for (let r = headerRowIdx + 1; r < clientEnd; r++) {
    const row = arr[r]; if (!row) continue;
    const nm = norm(row[nameCol]);
    if (GROUP_ROW.test(nm)) break;
    if (!nm) continue;                       // a blank line inside the client block is not the end of it
    // A summary or legend line is not a client. Without this guard a totals row whose label
    // the parser did not recognise would be added on top of the clients it already totals.
    if (NOT_A_CLIENT.test(nm)) continue;
    clients++;
    const s = num(row[sessionsCol]); if (s != null) sumS += s;
    const rv = num(row[moneyCol]); if (rv != null) sumR += rv;
    const cc = cancelledCol >= 0 ? num(row[cancelledCol]) : null; if (cc != null) sumC += cc;
  }

  // When the money column had no header and its position was inferred, prove it: on a
  // client row, price x sessions should come out at the money figure. Older sheets are laid
  // out more simply than today's, and a wrong column would otherwise pass silently.
  if (moneyGuessed) {
    let checked = 0, agreed = 0;
    for (let r = headerRowIdx + 1; r < clientEnd && checked < 12; r++) {
      const row = arr[r]; if (!row) continue;
      const nm = norm(row[nameCol]);
      if (GROUP_ROW.test(nm)) break;
      if (!nm || NOT_A_CLIENT.test(nm)) continue;
      const price = num(row[ppsCol]), sess = num(row[sessionsCol]), money = num(row[moneyCol]);
      if (price == null || !sess || money == null) continue;
      checked++;
      if (Math.abs(price * sess - money) <= Math.max(5, money * 0.1)) agreed++;
    }
    // No verifiable row either: refuse rather than report a revenue of zero for a month
    // that plainly has sessions in it.
    const unproven = checked === 0 ? sumS > 0 : agreed < Math.ceil(checked * 0.6);
    if (unproven) {
      return { error: "the revenue column could not be identified with confidence — this sheet is laid out differently." };
    }
  }

  const cancelled = (cancelledCol >= 0) ? sumC : null;
  const derived = conductedRow < 0;
  const sessions = derived ? sumS : num(arr[conductedRow][sessionsCol]);
  const revenue = derived ? Math.round(sumR * 100) / 100 : num(arr[conductedRow][moneyCol]);
  const rate = (revenue != null && sessions) ? Math.round((revenue / sessions) * 100) / 100 : null;

  const warnings = [];
  if (!derived && sessions != null && Math.abs(sumS - sessions) > 0.5) warnings.push(`sessions: sheet total ${sessions}, client rows add to ${sumS}`);
  if (!derived && revenue != null && Math.abs(sumR - revenue) > 5) warnings.push(`revenue: sheet total ${revenue}, client rows add to ${Math.round(sumR)}`);
  // A rate far outside anything the business charges means a column was misread.
  if (rate != null && (rate < 20 || rate > 5000)) warnings.push(`rate per session works out at ${Math.round(rate)} — check this sheet's columns`);
  if (sheetCancelled != null && cancelled != null && Math.abs(sheetCancelled - cancelled) > 0.5)
    warnings.push(`cancelled: client rows add to ${cancelled}, the sheet's own total row says ${sheetCancelled}`);

  // Month to date, counted from the day columns rather than taken from the totals.
  const cutoff = opts && opts.cutoffDay ? Number(opts.cutoffDay) : null;
  let mtd = null;
  if (cutoff) {
    if (!dayCols.length) {
      warnings.push("no day columns were found, so a month-to-date figure could not be worked out for this sheet");
    } else {
      const upTo = dayCols.filter((x) => x.d <= cutoff);
      let mS = 0, mR = 0, mX = 0, marks = 0, odd = 0;
      for (let r = headerRowIdx + 1; r < clientEnd; r++) {
        const row = arr[r]; if (!row) continue;
        const nm = norm(row[nameCol]);
        if (GROUP_ROW.test(nm)) break;
        if (!nm || NOT_A_CLIENT.test(nm)) continue;
        const price = num(row[ppsCol]);
        for (const { c } of upTo) {
          const v = num(row[c]);
          if (v == null) continue;             // an empty day is simply no session
          marks++;
          // Matching the sheet's own rule: a 1 is a session held, an exact 0 is a session
          // cancelled. Anything else in a day cell is left out of both counts.
          if (v >= 1) { mS += v; if (price != null) mR += v * price; }
          else if (v === 0) mX += 1;
          else odd++;
        }
      }
      // The sheet's own per-day tally, where it has one, is the check on the count above.
      let tallied = null;
      if (conductedRow >= 0) {
        tallied = 0;
        for (const { c } of upTo) { const v = num(arr[conductedRow][c]); if (v != null) tallied += v; }
      }
      // The client rows are the basis: they are the same cells the revenue and the
      // cancellations come from, so the three always agree with each other. Where the
      // sheet's own day row disagrees, that is reported rather than silently preferred.
      if (tallied != null && Math.abs(tallied - mS) > 0.5) {
        warnings.push(`month to date: the client rows add to ${mS} sessions, the sheet's day row to ${tallied}`);
      }
      mtd = {
        day: cutoff,
        sessions: mS,
        dayRow: tallied,
        revenue: Math.round(mR * 100) / 100,
        cancelled: mX,
        marks,
      };
      if (!marks) warnings.push("month to date: no day cells were filled in up to today");
      if (odd) warnings.push(`month to date: ${odd} day cell(s) held something other than 0 or 1 and were not counted`);
    }
  }

  return { sessions, revenue, cancelled, rate, clients, derived, dayCols: dayCols.length, mtd, warnings };
}
