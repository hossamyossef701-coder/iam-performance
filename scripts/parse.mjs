// Parser for one coach's monthly sheet. Ported 1:1 from the dashboard's verified
// odParseCoachSheet: every value is located by header text / row label, never by a fixed
// row or column index, so a different client count or month length does not break it.
const MONTH_IDX = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 };

// "Oct 2026", "October 2026", "Sept 2026" -> "2026-10". Anything else -> null.
export function sheetNameToPeriod(name) {
  const m = String(name).trim().toLowerCase().match(/^([a-z]{3,9})\.?\s+(\d{4})$/);
  if (!m) return null;
  const mo = MONTH_IDX[m[1].slice(0, 3)];
  if (!mo) return null;
  return m[2] + "-" + String(mo).padStart(2, "0");
}

// Coach name = text before the first "_" or "-" in the file name ("Faisal - I AM ..." -> "Faisal").
export function coachFromFilename(filename) {
  const base = String(filename).replace(/\.(xlsx|xls|csv)$/i, "");
  const m = base.match(/^([^_\-]+)/);
  return (m ? m[1] : base).trim();
}

export function parseCoachSheet(arr) {
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(arr.length, 5); i++) {
    if (arr[i] && arr[i][1] != null && String(arr[i][1]).trim().toLowerCase() === "client name") { headerRowIdx = i; break; }
  }
  if (headerRowIdx < 0) return { error: "Could not find the “Client Name” header row — sheet layout may have changed." };

  const headers = arr[headerRowIdx];
  const findCol = (label) => {
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] != null && String(headers[i]).trim().toLowerCase() === label.toLowerCase()) return i;
    }
    return -1;
  };
  const ppsCol = findCol("pps"), taCol = findCol("ta"), cancelledCol = findCol("cancelled sessions");
  if (ppsCol < 0 || taCol < 0 || cancelledCol < 0) return { error: "Expected columns (PPS / TA / Cancelled Sessions) were not found — sheet layout may have changed." };
  const sessionsCol = ppsCol + 1; // the sessions column sits right after PPS and has a blank header

  let totalsRow = null, conductedRow = null;
  for (let r = headerRowIdx + 1; r < arr.length; r++) {
    const row = arr[r]; if (!row) continue;
    const label = row[1] != null ? String(row[1]).trim().toLowerCase() : "";
    const nameBlank = !(row[1] != null && String(row[1]).trim());
    if (label === "conducted sessions / day") { conductedRow = row; break; }
    if (nameBlank && row[cancelledCol] != null && row[cancelledCol] !== "") totalsRow = row;
  }
  if (!conductedRow) return { error: "Could not find the “Conducted Sessions / Day” row — sheet layout may have changed." };

  const num = (v) => { const n = parseFloat(v); return isNaN(n) ? null : n; };
  const sessions = num(conductedRow[sessionsCol]);
  const revenue = num(conductedRow[taCol]);
  const cancelled = totalsRow ? num(totalsRow[cancelledCol]) : null;
  const rate = (revenue != null && sessions) ? Math.round((revenue / sessions) * 100) / 100 : null;

  // Cross-check: re-sum every named client row and compare with the sheet's own total.
  let sS = 0, sR = 0, sC = 0;
  for (let c = headerRowIdx + 1; c < arr.length; c++) {
    const row = arr[c]; if (!row) continue;
    const nm = row[1] != null ? String(row[1]).trim() : "";
    if (!nm) break;
    const s = num(row[sessionsCol]); if (s != null) sS += s;
    const rv = num(row[taCol]); if (rv != null) sR += rv;
    const cc = num(row[cancelledCol]); if (cc != null) sC += cc;
  }
  const warnings = [];
  if (sessions != null && Math.abs(sS - sessions) > 0.5) warnings.push(`sessions: sheet total ${sessions}, named clients sum to ${sS} (check for an unnamed row)`);
  if (revenue != null && Math.abs(sR - revenue) > 5) warnings.push(`revenue: sheet total ${revenue}, named clients sum to ${sR}`);
  if (cancelled != null && Math.abs(sC - cancelled) > 0.5) warnings.push(`cancelled: sheet total ${cancelled}, named clients sum to ${sC}`);
  return { sessions, revenue, cancelled, rate, warnings };
}
