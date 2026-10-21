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
  if (conductedRow < 0) return { error: "no “Conducted Sessions” row was found — the sheet layout may have changed." };

  const sessions = num(arr[conductedRow][sessionsCol]);
  const revenue = num(arr[conductedRow][taCol]);

  // 3. Cancelled sessions: the running total is the last figure in that column above the
  //    conducted row - which is the unnamed totals row in some files and the Group Training
  //    row in others.
  let cancelled = null;
  if (cancelledCol >= 0) {
    for (let r = headerRowIdx + 1; r < conductedRow; r++) {
      const v = arr[r] ? num(arr[r][cancelledCol]) : null;
      if (v != null) cancelled = v;
    }
  }

  const rate = (revenue != null && sessions) ? Math.round((revenue / sessions) * 100) / 100 : null;

  // 4. Cross-check: re-add the individual client rows and compare with the sheet's own total.
  //    Group training is counted separately by the sheet, so it ends the client block.
  let sumS = 0, sumR = 0, sumC = 0, clients = 0;
  for (let r = headerRowIdx + 1; r < conductedRow; r++) {
    const row = arr[r]; if (!row) continue;
    const nm = norm(row[nameCol]);
    if (GROUP_ROW.test(nm)) break;
    if (!nm) continue;                       // a blank line inside the client block is not the end of it
    clients++;
    const s = num(row[sessionsCol]); if (s != null) sumS += s;
    const rv = num(row[taCol]); if (rv != null) sumR += rv;
    const cc = cancelledCol >= 0 ? num(row[cancelledCol]) : null; if (cc != null) sumC += cc;
  }

  const warnings = [];
  if (sessions != null && Math.abs(sumS - sessions) > 0.5) warnings.push(`sessions: sheet total ${sessions}, client rows add to ${sumS}`);
  if (revenue != null && Math.abs(sumR - revenue) > 5) warnings.push(`revenue: sheet total ${revenue}, client rows add to ${Math.round(sumR)}`);
  if (cancelled != null && Math.abs(sumC - cancelled) > 0.5) warnings.push(`cancelled: sheet total ${cancelled}, client rows add to ${sumC}`);

  return { sessions, revenue, cancelled, rate, clients, warnings };
}
