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

const NAME_HEADERS = ["client name", "clients", "client", "client names", "name"];
const GROUP_ROW = /^group\s*(training|class|classes)/;

export function parseCoachSheet(arr) {
  // 1. The header row is the one carrying the client-name column; its position and spelling
  //    both vary, so search the opening rows across the opening columns.
  let headerRowIdx = -1, nameCol = -1;
  outer:
  for (let r = 0; r < Math.min(arr.length, 12); r++) {
    const row = arr[r]; if (!row) continue;
    for (let c = 0; c < Math.min(row.length, 8); c++) {
      if (NAME_HEADERS.includes(norm(row[c]))) { headerRowIdx = r; nameCol = c; break outer; }
    }
  }
  if (headerRowIdx < 0) return { error: "no client-name column was found in the opening rows — the sheet layout may have changed." };

  const headers = arr[headerRowIdx];
  const findCol = (test) => {
    for (let i = 0; i < headers.length; i++) if (test(norm(headers[i]))) return i;
    return -1;
  };
  const ppsCol = findCol((h) => h === "pps");
  const taCol = findCol((h) => h === "ta");
  // "Cancelled Sessions" is sometimes wrapped across two lines in the cell.
  const cancelledCol = findCol((h) => /^cancell?ed sessions$/.test(h));
  if (ppsCol < 0 || taCol < 0) return { error: "the PPS or TA column was not found — the sheet layout may have changed." };
  const sessionsCol = ppsCol + 1;   // the session count sits right after PPS, under a blank header

  // 2. The totals row: "Conducted Sessions", with or without "/ Day", in or near the name column.
  let conductedRow = -1;
  for (let r = headerRowIdx + 1; r < arr.length; r++) {
    const row = arr[r]; if (!row) continue;
    for (let c = 0; c <= nameCol + 2 && c < row.length; c++) {
      const t = norm(row[c]);
      if (t.startsWith("conducted session")) { conductedRow = r; break; }
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
    clients++;
    const s = num(row[sessionsCol]); if (s != null) sumS += s;
    const rv = num(row[taCol]); if (rv != null) sumR += rv;
    const cc = cancelledCol >= 0 ? num(row[cancelledCol]) : null; if (cc != null) sumC += cc;
  }

  const cancelled = (cancelledCol >= 0) ? sumC : null;
  const derived = conductedRow < 0;
  const sessions = derived ? sumS : num(arr[conductedRow][sessionsCol]);
  const revenue = derived ? Math.round(sumR * 100) / 100 : num(arr[conductedRow][taCol]);
  const rate = (revenue != null && sessions) ? Math.round((revenue / sessions) * 100) / 100 : null;

  const warnings = [];
  if (!derived && sessions != null && Math.abs(sumS - sessions) > 0.5) warnings.push(`sessions: sheet total ${sessions}, client rows add to ${sumS}`);
  if (!derived && revenue != null && Math.abs(sumR - revenue) > 5) warnings.push(`revenue: sheet total ${revenue}, client rows add to ${Math.round(sumR)}`);
  if (sheetCancelled != null && cancelled != null && Math.abs(sheetCancelled - cancelled) > 0.5)
    warnings.push(`cancelled: client rows add to ${cancelled}, the sheet's own total row says ${sheetCancelled}`);

  return { sessions, revenue, cancelled, rate, clients, derived, warnings };
}
