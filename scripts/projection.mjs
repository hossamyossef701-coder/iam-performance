// The projection that counts, month by month, for every coach.
//
// Coaches type a projection per client on their sheet each month. Three things keep it honest:
//  1. The lock: the projection read on the 3rd of the month (or the first read after it) is the one
//     that counts. Later edits are shown, not used.
//  2. The floor: the coach's own average over the 3 months before, less the usual dip (how far an
//     honest month falls below that average 1 month in 10, measured on every coach-month on the
//     sheets), adjusted for the season (the same month in past years against the 3 months before
//     it, so a Ramadan or summer month gets a lower floor by itself). The higher of the locked projection and
//     the floor counts. No projection by the 3rd: the floor counts.
//  3. Overrides, set by the Fitness Director on the Team Performance page after talking to a coach:
//     "set" (use these figures) or "nofloor" (use the projection as given, no floor this month).
// Sessions delivered are never touched: every session on a sheet counts, new clients and covers
// included (a cover counts for the coach who delivered it).
//
// The counted figures replace ps / pr on each coach-month, so every page that reads the projection
// uses them. The sheet's own figures stay as ps0 / pr0, and pg says how the count was reached.

export const PROJ_FROM = "2026-11";        // the first month coaches set projections
export const LOCK_DAY = 3;

const prevP = (p, n) => { let y = +p.slice(0, 4), m = +p.slice(5, 7) - n; while (m < 1) { m += 12; y--; } while (m > 12) { m -= 12; y++; } return y + "-" + String(m).padStart(2, "0"); };
const key = (n) => String(n || "").toLowerCase().replace(/[^a-z0-9؀-ۿ]+/g, " ").trim();
const wOf = (ch) => (ch === "1" || ch === "x" ? 1 : ch === "h" ? 0.5 : ch >= "2" && ch <= "9" ? +ch : 0);

// The usual dip: across every coach-month on the sheets (the coach steady for the 3 months before),
// the month against the coach's own 3-month average, with the season taken out. The 10th
// percentile is how far an honest month can fall: below it only 1 month in 10.
function dipFromHistory(periods, field) {
  const rs = [];
  for (const p of Object.keys(periods).sort()) for (const [c, v] of Object.entries(periods[p])) {
    const b = [1, 2, 3].map((i) => (periods[prevP(p, i)] || {})[c] || {});
    if (b.some((x) => !((x.s || 0) >= 20)) || !(v[field] > 0) || b.some((x) => !(x[field] > 0))) continue;
    rs.push(v[field] / (b.reduce((a, x) => a + x[field], 0) / 3) / seasonOf(periods, p));
  }
  if (rs.length < 20) return { q: 0.65, n: rs.length };
  rs.sort((a, b) => a - b);
  return { q: Math.min(0.9, Math.max(0.4, rs[Math.floor(0.1 * (rs.length - 1))])), n: rs.length };
}

// The season: the same month in past years against the 3 months before it, team-wide.
function seasonOf(periods, p) {
  const k0 = "_season_" + p; if (periods[k0] !== undefined) return periods[k0];
  const team = (q) => Object.values(periods[q] || {}).reduce((t, v) => t + (v.s || 0), 0);
  const rs = [];
  for (let k = 1; k <= 4; k++) {
    const q = (+p.slice(0, 4) - k) + p.slice(4), b = [1, 2, 3].map((i) => team(prevP(q, i))), t = team(q);
    if (t > 0 && b.every((x) => x > 0)) rs.push(t / (b.reduce((a, c) => a + c, 0) / 3));
  }
  let m = 1;
  if (rs.length) { rs.sort((a, b) => a - b); m = rs.length % 2 ? rs[(rs.length - 1) / 2] : (rs[rs.length / 2 - 1] + rs[rs.length / 2]) / 2; }
  m = Math.min(1.2, Math.max(0.6, m));
  Object.defineProperty(periods, k0, { value: m, enumerable: false, configurable: true });
  return m;
}

// The floor for one coach-month: their own average over the 3 months before, times the season,
// times the usual dip. A coach without 3 steady months before (new, back from leave) has none.
export function floorOf(periods, p, c, dip) {
  const b = [1, 2, 3].map((i) => (periods[prevP(p, i)] || {})[c] || {});
  if (b.some((x) => !((x.s || 0) >= 20))) return { fl: null, flr: null };
  const avgS = b.reduce((a, x) => a + x.s, 0) / 3, avgR = b.reduce((a, x) => a + (x.r || 0), 0) / 3, se = seasonOf(periods, p);
  return { usual: Math.round(avgS), usualR: Math.round(avgR), se: Math.round(se * 100) / 100, a: Math.round((1 - dip.s.q) * 100),
    fl: Math.round(avgS * se * dip.s.q), flr: avgR ? Math.round(avgR * se * dip.r.q) : null };
}

// Each client's projection on the sheet now: { client key: AED }, from the cl rows (index 4).
function clientProj(v) {
  const out = {};
  for (const row of v.cl || []) { const a = row[4]; if (a > 0) { const k = key(row[0]); out[k] = Math.round(((out[k] || 0) + a) * 100) / 100; } }
  return out;
}
// The projection checked against the clients it was made for: the clients projected when it
// locked, and what those clients brought in (new clients after the lock are left out - they
// count as new clients, not against the projection). own: their projection (or the figure the
// Fitness Director set); got: those clients' revenue; n / of: clients found on the sheet / projected.
function projectedClients(v, lk, ov) {
  const pc = (lk && lk.pc) || clientProj(v), keys = Object.keys(pc);
  let own = keys.reduce((a, k) => a + pc[k], 0);
  if (ov && ov.mode === "set" && ov.pr != null) own = ov.pr;
  if (!(own > 0)) return null;
  let got = 0; const seen = {};
  for (const row of v.cl || []) {
    const k = key(row[0]); if (!(k in pc)) continue;
    seen[k] = 1;
    const r = row[5] != null ? row[5] : (row[1] || 0) * [...(row[2] || "")].reduce((a, ch) => a + wOf(ch), 0);
    got += r || 0;
  }
  const out = { own: Math.round(own), got: Math.round(got), n: Object.keys(seen).length, of: keys.length, est: !(lk && lk.pc) || !!(lk && lk.pcLate), ng: 0, nn: 0 };
  Object.defineProperty(out, "pcKeys", { value: pc, enumerable: false, configurable: true });   // for addNewBusiness, never written
  return out;
}

// New and returning clients (from the Clients file, joined in scripts/join.mjs: fresh[month][coach]
// = sheet name keys) added to each coach-month's check: ng = their revenue, nn = how many. Only
// clients not projected on the 3rd. A client counts once even if two rows carry the name.
export function addNewBusiness(periods, fresh, nameKey) {
  for (const p of Object.keys(fresh || {})) for (const c of Object.keys(fresh[p])) {
    const v = periods[p] && periods[p][c]; if (!v || !v.pk) continue;
    const want = new Set(fresh[p][c]), pc = v.pk.pcKeys || {};
    let ng = 0, nn = 0; const seen = new Set();
    for (const row of v.cl || []) {
      const nk = nameKey(row[0]); if (!want.has(nk) || pc[key(row[0])]) continue;
      const r = row[5] != null ? row[5] : (row[1] || 0) * [...(row[2] || "")].reduce((a, ch) => a + wOf(ch), 0);
      ng += r || 0; if (!seen.has(nk)) { seen.add(nk); nn++; }
    }
    v.pk.ng = Math.round(ng); v.pk.nn = nn;
  }
  for (const p of Object.keys(periods)) for (const v of Object.values(periods[p])) if (v.pk) delete v.pk.pcKeys;
}

// lock: { "2026-11": { coach: { ps, pr, at, late, pc: { client: AED } } } }; overrides: { "2026-11": { coach: { mode, ps, pr, why, at } } }
export function guardProjections(periods, { now, today, todayIso, lock = {}, overrides = {}, from = PROJ_FROM }) {
  const dip = { s: dipFromHistory(periods, "s"), r: dipFromHistory(periods, "r") }, notes = [];
  let lockChanged = false;
  for (const p of Object.keys(periods).sort()) {
    if (p < from || p > now) continue;
    const L = lock[p] || (lock[p] = {}), O = overrides[p] || {};
    for (const [c, v] of Object.entries(periods[p])) {
      const sheet = { ps: v.ps ?? null, pr: v.pr ?? null };
      v.ps0 = sheet.ps; v.pr0 = sheet.pr;
      // 1. The lock: from the 3rd of the month running, or straight away for a month already past.
      if (!L[c] && (p < now || today >= LOCK_DAY)) { L[c] = { ps: sheet.ps, pr: sheet.pr, at: todayIso, late: p < now || today > LOCK_DAY, pc: clientProj(v) }; lockChanged = true; }
      // Locks made before the clients were kept (Nov 2026): the clients projected on the sheet when
      // first seen, marked as taken late.
      else if (L[c] && !L[c].pc && v.cl) { L[c].pc = clientProj(v); L[c].pcLate = true; lockChanged = true; }
      const lk = L[c] || null, ov = O[c] || null;
      const base = ov && ov.mode === "set" ? { ps: ov.ps ?? null, pr: ov.pr ?? null } : lk ? { ps: lk.ps, pr: lk.pr } : sheet;
      // 2. The floor.
      const f = floorOf(periods, p, c, dip);
      const set = (base.ps || 0) > 0 || (base.pr || 0) > 0;
      let ps = base.ps, pr = base.pr, st;
      if (ov && ov.mode === "nofloor") st = "override";
      else if (ov && ov.mode === "set") st = "override";
      else if (!set) { st = f.fl ? (lk ? "unset" : "pending") : "none"; if (f.fl) { ps = f.fl; pr = f.flr; } }
      else if (f.fl == null) st = "unchecked";
      else {
        const lowS = (base.ps || 0) > 0 && base.ps < f.fl, lowR = (base.pr || 0) > 0 && base.pr < f.flr;
        if (lowS) ps = f.fl; if (lowR) pr = f.flr;
        st = lowS || lowR ? "below" : "ok";
      }
      if (ps != null) v.ps = ps; else delete v.ps;
      if (pr != null) v.pr = pr; else delete v.pr;
      const pk = projectedClients(v, lk, ov); if (pk) v.pk = pk; else delete v.pk;
      v.pg = { st, fl: f.fl, flr: f.flr, a: f.a ?? null, se: f.se ?? null, usual: f.usual ?? null,
        lk: lk ? { ps: lk.ps, pr: lk.pr, at: lk.at, late: !!lk.late } : null,
        ov: ov ? { mode: ov.mode, why: ov.why || "", at: ov.at || "" } : null,
        changed: lk && !(ov && ov.mode === "set") && ((sheet.ps ?? null) !== (lk.ps ?? null) || (sheet.pr ?? null) !== (lk.pr ?? null)) ? { ps: sheet.ps, pr: sheet.pr } : null };
      if (st === "below") notes.push(`${p} ${c}: projected ${base.ps ?? "-"} sessions, floor ${f.fl} (usual ${f.usual}, season x${f.se}, less ${f.a}%) - the floor counts`);
      if (st === "unset") notes.push(`${p} ${c}: no projection by the ${LOCK_DAY}rd - the floor (${f.fl}) counts`);
    }
  }
  return { lock, lockChanged, dip, notes };
}
