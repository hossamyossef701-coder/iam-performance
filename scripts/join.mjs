/* The coach sheets joined to the Clients file, for the Overview and Coach review pages.

   A coach sheet has a client's name as the coach typed it ("Lisa R (24)") and the days they
   trained. The Clients file has the client's ID, packages and dates. Joining them means
   finding, for each name on each coach's sheet, the client it belongs to:

     1. A row on the Clients file's "Name links" tab (Coach | Name on coach sheet | Client ID)
        decides it outright - that is how a name is confirmed once and remembered. "none"
        there means "not a client in the file" (a trial, a family member, a typo).
     2. Otherwise the name is matched by its words: every word typed on the sheet has to fit
        a word of the client's full name (a single letter fits a surname's first letter,
        one slip of the keyboard is allowed), and the client must have had a package with
        that coach. Only a single clear winner is linked; anything closer is listed for a
        person to confirm.

   Nothing here is written anywhere but the encrypted overview file. */

const DAY_MS = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);
const iso = (n) => new Date(EPOCH + n * DAY_MS).toISOString().slice(0, 10);
const serial = (s) => Math.round((Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) - EPOCH) / DAY_MS);
const daysIn = (p) => new Date(Date.UTC(+p.slice(0, 4), +p.slice(5, 7), 0)).getUTCDate();
const wOf = (ch) => (ch === "1" || ch === "x" ? 1 : ch === "h" ? 0.5 : ch >= "2" && ch <= "9" ? +ch : 0);
const TITLES = new Set(["mr", "mrs", "ms", "miss", "dr", "sir", "madam", "mme"]);

// The words of a name, as the matching sees them.
export function words(n) {
  return String(n || "").toLowerCase()
    .replace(/\([^)]*\)/g, " ")                       // package notes: "(24)", "(online)"
    .normalize("NFD").replace(/[̀-ͯ]/g, "")  // é -> e
    .replace(/[^a-z؀-ۿ ]+/g, " ")
    .split(/\s+/).filter((w) => w && !TITLES.has(w))
    .map((w) => (/^m[ou]h?[a]?m+[ae]d$/.test(w) ? "mohammed" : w));   // Mohamed, Muhammad, Mohammad
}
export const nameKey = (n) => words(n).join(" ");

function near(a, b) {                                  // at most one character different
  if (a === b) return true;
  if (Math.abs(a.length - b.length) > 1 || Math.min(a.length, b.length) < 4) return false;
  let i = 0, j = 0, slips = 0;
  while (i < a.length && j < b.length) {
    if (a[i] === b[j]) { i++; j++; continue; }
    if (++slips > 1) return false;
    if (a.length > b.length) i++; else if (a.length < b.length) j++; else { i++; j++; }
  }
  return slips + (a.length - i) + (b.length - j) <= 1;
}
// How well the words typed on a sheet fit a client's full name: f is the share of words that
// fit at all (0 to 1), ex the share that fit exactly (an initial or a slip does not).
function fit(sheetW, fullW) {
  const none = { f: 0, ex: 0 };
  if (!sheetW.length || !fullW.length) return none;
  // The first word typed is the name people go by: it has to be there.
  if (!fullW.some((f) => f === sheetW[0] || near(f, sheetW[0]))) return none;
  const used = new Set();
  let hit = 0, exact = 0;
  sheetW.forEach((w, k) => {
    let j = fullW.findIndex((f, i) => !used.has(i) && f === w);
    if (j >= 0) { used.add(j); hit++; exact++; return; }
    j = fullW.findIndex((f, i) => !used.has(i) && (near(f, w) || (w.length === 1 && k > 0 && f[0] === w) || (w.length >= 3 && k > 0 && f.startsWith(w))));
    if (j >= 0) { used.add(j); hit++; }
  });
  return { f: hit / sheetW.length, ex: exact / sheetW.length };
}

/* inner: from buildClientsModel(..., { internal: true })._in
   periods: the coach figures, with cl (each client's day-by-day row) on recent months
   links: the "Name links" tab as rows of cells
   now: "YYYY-MM" running in Dubai; cutoffDay: today's day of the month */
export function buildOverview({ inner, periods, links = [], now, cutoffDay, revenueOf, debug = false }) {
  const today = inner.today;
  const clients = new Map(inner.clients.map((c) => [c.id, c]));
  const fullW = new Map(inner.clients.map((c) => [c.id, words(c.name)]));
  const sq = (s) => String(s || "").toLowerCase().replace(/[^a-z]/g, "");

  /* 1. Every name on every coach's sheet, with the days it trained ---- */
  const sheetNames = new Map();                       // "coach|key" -> {coach, name, key, months:Set, days:Map(serial->w), reds:Set(serial)}
  for (const p of Object.keys(periods).sort()) {
    for (const coach of Object.keys(periods[p])) {
      const cl = periods[p][coach].cl; if (!cl) continue;
      const last = p === now ? cutoffDay : daysIn(p);
      for (const [name, , m, col] of cl) {
        const key = nameKey(name); if (!key) continue;
        const k = coach + "|" + key;
        const o = sheetNames.get(k) || { coach, name, key, months: new Set(), days: new Map(), reds: new Set() };
        o.name = name; o.months.add(p);
        for (let i = 0; i < Math.min(last, m.length); i++) {
          const w = wOf(m[i]); if (!w) continue;
          const d = serial(`${p}-${String(i + 1).padStart(2, "0")}`);
          o.days.set(d, (o.days.get(d) || 0) + w);
          // A session in red: the coach's mark for the last session of the package.
          if (col && col[i] === "r") o.reds.add(d);
          // A session in orange: a cover - this coach trained another coach's client.
          if (col && col[i] === "o") o.cover = (o.cover || 0) + 1;
        }
        sheetNames.set(k, o);
      }
    }
  }

  /* 2. Confirmed links from the Clients file ---- */
  const confirmed = new Map();                         // "sqcoach|key" -> id | null
  const bad = [];
  let head = -1;
  for (let r = 0; r < Math.min(links.length, 10); r++) {
    const hs = (links[r] || []).map((h) => String(h || "").toLowerCase());
    if (hs.some((h) => h.includes("coach")) && hs.some((h) => h.includes("client id"))) { head = r; break; }
  }
  if (head >= 0) {
    const hs = links[head].map((h) => String(h || "").toLowerCase());
    const ci = hs.findIndex((h) => h.includes("coach")), ni = hs.findIndex((h) => h.includes("name")), ii = hs.findIndex((h) => h.includes("client id"));
    for (let r = head + 1; r < links.length; r++) {
      const row = links[r] || [];
      const coach = String(row[ci] || "").trim(), name = String(row[ni] || "").trim(), id = String(row[ii] || "").trim().toUpperCase();
      // A row with no Client ID yet is a note to come back to: it decides nothing.
      if (!coach || !name || !id) continue;
      if (id !== "NONE" && !clients.has(id)) { bad.push(`${coach}: "${name}" points to ${id}, which is not on the Clients tab`); continue; }
      confirmed.set(sq(coach) + "|" + nameKey(name), id !== "NONE" ? id : null);
    }
  }

  /* 3. Match each name ---- */
  // The coach sheets and the Clients file may spell a coach differently ("Hossam Y" / "Hossam",
  // "Shayla" / "Shyla"): same letters, or the same first name, counts as the same coach.
  const first = (n) => sq(String(n || "").trim().split(/\s+/)[0]);
  const sameCoach = (a, b) => !!a && !!b && (sq(a) === sq(b) || first(a) === first(b) || (first(a).length >= 4 && near(first(a), first(b))));
  const coachHad = (c, coach) => sameCoach(c.coach, coach) || c.pk.some((p) => p.coaches.some((x) => sameCoach(x, coach)));
  const linkOf = new Map();                            // "coach|key" -> id
  const unclear = [], notInFile = [], dbg = [];
  let auto = 0, byHand = 0, ignored = 0;
  for (const [k, o] of sheetNames) {
    const ck = sq(o.coach) + "|" + o.key;
    if (confirmed.has(ck)) { const id = confirmed.get(ck); if (id) { linkOf.set(k, id); byHand++; } else ignored++; continue; }
    const sw = words(o.name);
    const scored = [];
    for (const c of inner.clients) {
      const { f, ex } = fit(sw, fullW.get(c.id)); if (!f) continue;
      const exact = fullW.get(c.id).join(" ") === sw.join(" ");
      scored.push({ id: c.id, f, ex, exact, had: coachHad(c, o.coach) });
    }
    scored.sort((a, b) => (b.exact - a.exact) || (b.had - a.had) || (b.f - a.f) || (b.ex - a.ex));
    // Linked only when one client clearly wins: the same full name, or the only client of this
    // coach whose name every typed word fits - exactly, or failing that, at all.
    const one = (list) => (list.length === 1 ? list[0] : null);
    // A client covered for another coach never had a package with this one. Such a name is
    // still linked when it can only be one client in the whole file: the full name as the file
    // has it, or - for a row with sessions marked as cover (orange) - the only client whose name
    // every typed word fits exactly.
    const pickOne = one(scored.filter((s) => s.exact && s.had))
      || one(scored.filter((s) => s.ex === 1 && s.had))
      || (scored.some((s) => s.ex === 1 && s.had) ? null : one(scored.filter((s) => s.f === 1 && s.had)))
      || (scored.some((s) => s.had) ? null : (sw.length >= 2 ? one(scored.filter((s) => s.exact)) : null) || (o.cover ? one(scored.filter((s) => s.ex === 1)) : null));
    const lastMonth = [...o.months].sort().pop();
    if (pickOne) { linkOf.set(k, pickOne.id); auto++; if (debug) dbg.push({ coach: o.coach, name: o.name, last: lastMonth, id: pickOne.id, how: pickOne.had ? (pickOne.exact ? "exact" : pickOne.ex === 1 ? "words" : "close") : "cover", others: scored.filter((x) => x !== pickOne).slice(0, 2).map((x) => x.id) }); continue; }
    if (scored.length) {
      unclear.push({ coach: o.coach, name: o.name, last: lastMonth,
        sug: scored.slice(0, 3).map((s) => ({ id: s.id, name: clients.get(s.id).name, had: s.had, f: s.f, ex: s.ex, exact: s.exact, why: s.had ? "had a package with " + o.coach : "never had a package with " + o.coach })) });
    } else notInFile.push({ coach: o.coach, name: o.name, last: lastMonth });
  }

  /* 4. Each linked client's days, all coaches together ---- */
  const daysOf = new Map();                            // id -> Map(serial -> w)
  const redsOf = new Map();                            // id -> Set(serial): last sessions marked in red
  const namesOf = new Map();                           // id -> [{coach,name}]
  for (const [k, id] of linkOf) {
    const o = sheetNames.get(k);
    const d = daysOf.get(id) || new Map();
    for (const [day, w] of o.days) d.set(day, (d.get(day) || 0) + w);
    daysOf.set(id, d);
    const rs = redsOf.get(id) || new Set(); for (const day of o.reds) rs.add(day); redsOf.set(id, rs);
    (namesOf.get(id) || namesOf.set(id, []).get(id)).push({ coach: o.coach, name: o.name, months: o.months });
  }
  const sumDays = (d, from, to) => { let s = 0; if (d) for (const [day, w] of d) if (day >= from && day <= to) s += w; return s; };

  // Clients still on this month's sheets whom the file has as ended.
  const endedOnSheet = [];
  for (const [k, id] of linkOf) {
    const o = sheetNames.get(k), c = clients.get(id);
    if (o.months.has(now) && c.status === "Ended") endedOnSheet.push({ coach: o.coach, name: o.name, id, until: c.P != null ? iso(c.P) : null });
  }

  /* 5. Packages in use, per client ----
     Sessions are counted against packages in order, not by date: a package takes sessions until
     its size is used up, and every session after that belongs to the next package, whenever it
     is paid for. So a client who finishes 24 sessions on 31 Oct, trains on 1, 3 and 5 Nov and
     renews on 6 Nov has 3 of the new package used on the day it is bought - nothing has to look
     back for them. Before the renewal is logged, those 3 show as sessions not yet paid for. And
     a client who renews early uses up the old package first. */
  const clMonths = [...new Set([...sheetNames.values()].flatMap((o) => [...o.months]))].sort();
  const W0 = clMonths.length ? serial(clMonths[0] + "-01") : today;   // the first day the sheets cover
  /* A package ends at its last paid session. The coach marks it in red; where a red mark falls
     within NEAR sessions of the package's size it is taken as the end (the count can be off by a
     missed or extra mark), otherwise the count decides - so a sheet with no colours works as before.
     Each package records the day it ended (last) and whether a red mark said so (red). */
  const NEAR = 3;
  function ledger(c, d, T = today, fromDay = T - 30, reds = redsOf.get(c.id)) {
    const pk = c.pk.filter((p) => p.unit !== "Months" && p.size && p.date <= T);
    if (!pk.length || c.pk.some((p) => p.unit === "Months" && p.date <= T && p.end >= T)) return null;
    // Start at the first package bought inside the days the sheets cover (or the latest one).
    let i = pk.findIndex((p) => p.date >= W0); if (i < 0) i = pk.length - 1;
    const prev = pk[i - 1];
    // Sessions before that package count toward it from two weeks before it was bought at the
    // earliest (a client who trained before paying), and not while the package before it was
    // still valid. An older package is not taken to end on its expiry date: packages are often
    // used past it, so its sessions must not be pushed onto the next one.
    const from = Math.max(W0, pk[i].date - 14, prev ? prev.end + 1 : 0);
    const ses = [...(d || new Map())].filter(([day]) => day >= from && day <= T).sort((a, b) => a[0] - b[0]);
    const pkgs = pk.slice(i).map((p) => ({ date: p.date, end: p.end, size: p.size, coach: p.coach, type: p.type, used: 0, last: null, red: false, closed: false,
      px: p.px, rule: p.rule ?? null, fee: p.fee ?? null, sh: !!p.sh, ax: 0, axL: [] }));
    const isRed = (day) => !!(reds && reds.has(day));
    let k = 0, over = 0, overRecent = 0, late = 0, lateRecent = 0, misfit = 0;
    const overL = [];                                   // [day, sessions] trained beyond every paid session
    const close = (q, day, red) => { q.closed = true; q.last = day; q.red = red; };
    for (let si = 0; si < ses.length; si++) {
      const [day, w0] = ses[si];
      let w = w0;
      while (w > 0) {
        while (k < pkgs.length && pkgs[k].closed) k++;
        if (k >= pkgs.length) { over += w; overL.push([day, w]); if (day >= fromDay) overRecent += w; break; }
        const q = pkgs[k];
        // Up to the size; past it only while waiting for a red mark that is on its way.
        const room = q.used < q.size ? q.size - q.used : w;
        const take = Math.min(w, room);
        q.used += take; w -= take;
        if (day > q.end) { late += take; if (day >= fromDay) lateRecent += take; }
        // After the expiry date the policy gives (an extension does not move this one).
        if (q.rule != null && day > q.rule) { q.ax += take; q.axL.push([day, take]); }
        if (isRed(day) && w <= 0) {
          if (q.size - q.used <= NEAR) close(q, day, true); else misfit++;   // a red mark far from the count is not trusted
        } else if (q.used >= q.size) {
          // Size reached: end here, unless the coach's red mark comes within the next NEAR sessions.
          let ahead = false, extra = 0;
          for (let sj = si + 1; sj < ses.length && extra < NEAR; sj++) { extra += ses[sj][1]; if (isRed(ses[sj][0])) { ahead = true; break; } }
          if (!ahead || q.used - q.size >= NEAR) close(q, day, isRed(day));
        }
      }
    }
    while (k < pkgs.length && pkgs[k].closed) k++;
    const cur = pkgs[Math.min(k, pkgs.length - 1)], last = pkgs[pkgs.length - 1];
    const left = pkgs.reduce((t, p) => t + Math.max(0, p.size - p.used) * (p.closed ? 0 : 1), 0);
    return { cur, last, left, over, overL, overRecent, late, lateRecent, pkgs, ses, misfit };
  }
  const inUse = [];
  for (const c of inner.clients) {
    if (!["Active", "Renewal due", "Ended"].includes(c.status)) continue;
    const d = daysOf.get(c.id);
    const curPk = c.pk.filter((p) => p.date <= today).pop();
    const coach = (curPk && curPk.coach) || c.coach;
    const sheet = (namesOf.get(c.id) || []).find((x) => x.coach === coach) || (namesOf.get(c.id) || [])[0];
    const row = { id: c.id, name: c.name, coach, sheetName: sheet ? sheet.name : null, linked: !!d, status: c.status };
    if (!d) { if (c.status === "Active") inUse.push(Object.assign(row, { end: curPk ? iso(curPk.end) : null, state: "unlinked" })); continue; }
    const L = ledger(c, d);
    if (!L) { if (c.status === "Active" && curPk) inUse.push(Object.assign(row, { unit: curPk.unit, end: iso(curPk.end), state: "ok" })); continue; }
    const pace = sumDays(d, today - 27, today) / 4;
    const weeksLeft = Math.max(0, (L.last.end - today) / 7);
    Object.assign(row, { size: L.cur.size, used: L.cur.used, left: L.left, end: iso(L.cur.end), lastEnd: iso(L.last.end), pace: +pace.toFixed(2), weeksLeft: +weeksLeft.toFixed(1), unpaid: L.overRecent, late: L.lateRecent });
    if (L.overRecent > 0) row.state = "unpaid";                 // trained beyond every session paid for
    else if (L.lateRecent > 0) row.state = "expired";           // using sessions after the package ended
    else if (c.status !== "Active") continue;
    else row.state = L.left <= 2 ? "nearly" : L.left - pace * weeksLeft > 1 ? "short" : "ok";
    inUse.push(row);
  }

  // Up for renewal right now: cover from a package ending in the next 30 days, or run out within
  // the grace period with nothing new. Valued at the client's last package.
  const lastPx = (c) => { const l = c.pk.filter((p) => p.date <= today).pop(); return (l && l.px) || 0; };
  const soon = inner.clients.filter((c) => c.status === "Active" && c.P != null && c.P - today >= 0 && c.P - today <= 30 && c.pk.some((p) => p.end === c.P));
  const due = inner.clients.filter((c) => c.status === "Renewal due");
  const up = { n: soon.length + due.length, due: due.length, aed: Math.round([...soon, ...due].reduce((t, c) => t + lastPx(c), 0)), coach: {} };
  for (const c of [...soon, ...due]) { const k = c.coach || "(no coach)", o = up.coach[k] || (up.coach[k] = { n: 0, due: 0, aed: 0 }); o.n++; if (c.status === "Renewal due") o.due++; o.aed += lastPx(c); }
  for (const k in up.coach) up.coach[k].aed = Math.round(up.coach[k].aed);

  /* 6. Package sales against revenue, month by month ---- */
  const from = iso(inner.loggedFrom).slice(0, 7);
  const months = [];
  for (const p of Object.keys(periods).filter((q) => q >= from && q <= now).sort()) {
    const sold = {}, soldS = {};
    let tot = 0, totS = 0;
    for (const c of inner.clients) for (const pk of c.pk) {
      if (iso(pk.date).slice(0, 7) !== p || pk.px == null) continue;
      // A shared package is listed under each client sharing it: count it once.
      if (pk._seen === p) continue; pk._seen = p;
      const k = pk.coach || "(no coach)";
      sold[k] = (sold[k] || 0) + pk.px; tot += pk.px;
      if (pk.unit !== "Months" && pk.size) { soldS[k] = (soldS[k] || 0) + pk.size; totS += pk.size; }
    }
    const rev = {}; let totR = 0, totD = 0; const del = {};
    for (const coach of Object.keys(periods[p])) {
      const x = revenueOf(p, coach); if (!x) continue;
      rev[coach] = x.r; totR += x.r; del[coach] = x.s; totD += x.s;
    }
    months.push({ p, partial: p === now, sales: Math.round(tot), rev: Math.round(totR), soldS: totS, delS: Math.round(totD * 10) / 10,
      coach: Object.fromEntries([...new Set([...Object.keys(sold), ...Object.keys(rev)])].map((k) => [k, { s: Math.round(sold[k] || 0), r: Math.round(rev[k] || 0), ss: soldS[k] || 0, ds: del[k] || 0 }])) });
  }
  for (const c of inner.clients) for (const pk of c.pk) delete pk._seen;

  /* 7. Per coach: how long clients stay, and what keeps them ---- */
  const coaches = [...new Set(Object.values(periods).flatMap((o) => Object.keys(o)))];
  const year = today - 365;
  const dec = inner.decisions.filter((x) => x.outcome !== "pending" && x.E >= year && x.E <= today);
  // How long a client stays in their first year, by the coach of their first package: the area
  // under the survival curve to 12 months (clients still with us count as long as watched).
  const stayOf = (list) => {
    if (!list.length) return null;
    const t = [...new Set(list.map((r) => r.t))].sort((a, b) => a - b);
    let S = 1, prev = 0, area = 0;
    for (const ti of t) {
      const capped = Math.min(ti, 365);
      area += S * (capped - prev); prev = capped;
      if (ti >= 365) break;
      const n = list.filter((r) => r.t >= ti).length, e = list.filter((r) => r.t === ti && r.ev).length;
      if (n && e) S *= 1 - e / n;
    }
    if (prev < 365) area += S * (Math.min(365, Math.max(...list.map((r) => r.t))) - prev);
    return { months: area / 30.44, n: list.length };
  };
  const firstYear = [];
  for (const c of inner.clients) {
    const first = c.pk[0];
    if (!first || c.joined == null || c.joined < inner.loggedFrom || first.type === "Renewal") continue;
    const lost = inner.decisions.filter((x) => x.id === c.id && x.outcome === "lost").sort((a, b) => a.E - b.E)[0];
    firstYear.push({ coach: first.coach, t: Math.max(1, (lost ? lost.E : today) - c.joined), ev: lost ? 1 : 0 });
  }
  const perCoach = {};
  for (const coach of coaches) {
    const mine = dec.filter((x) => x.coach === coach);
    const last = new Map(); for (const x of mine) { const o = last.get(x.id); if (!o || o.E <= x.E) last.set(x.id, x); }
    const o = { kept: 0, moved: 0, lost: 0 }; for (const x of last.values()) o[x.outcome]++;
    const ren = mine.filter((x) => x.outcome !== "lost");
    const sizes = {}; for (const c of inner.clients) for (const p of c.pk) if (p.coach === coach && p.date >= year && p.unit !== "Months" && p.size) sizes[p.size] = (sizes[p.size] || 0) + 1;
    const common = Object.keys(sizes).sort((a, b) => sizes[b] - sizes[a])[0];
    perCoach[coach] = { ...o, early: ren.filter((x) => x.at <= x.E).length, renewals: ren.length,
      stay: stayOf(firstYear.filter((r) => r.coach === coach)), common: common ? +common : null };
  }
  const team = { stay: stayOf(firstYear) };

  /* 8. Each coach's clients: how often they train against what they signed up for ---- */
  const clientsOfCoach = {};
  const eightWeeks = today - 55;
  for (const c of inner.clients) {
    if (c.status !== "Active") continue;
    const cur = c.pk.filter((p) => p.date <= today).pop(); if (!cur) continue;
    const coach = cur.coach || c.coach; if (!coach) continue;
    const d = daysOf.get(c.id);
    // Agreed: what the client signed up for, when the file has it; otherwise the pace the
    // package itself needs (its sessions over its weeks).
    const pkPace = cur.unit !== "Months" && cur.size && cur.end > cur.date ? cur.size / ((cur.end - cur.date) / 7) : null;
    const agreed = c.agreed != null ? c.agreed : pkPace;
    const since = Math.max(eightWeeks, cur.date, c.joined ?? cur.date);
    const weeks = Math.max(1, (today - since + 1) / 7);
    (clientsOfCoach[coach] = clientsOfCoach[coach] || []).push({ id: c.id, name: c.name, linked: !!d,
      actual: d ? +(sumDays(d, since, today) / weeks).toFixed(2) : null, agreed: agreed != null ? +agreed.toFixed(2) : null, signed: c.agreed != null });
  }

  /* 9. Month by month: clients who trained in the month beyond every session paid for by its
     end (a renewal logged later in the same month clears it). Put under the coach whose sheet
     has most of that client's sessions in the month. */
  const upM = {};
  const byCoachDays = new Map();                       // id -> [{coach, days}]
  for (const [k, id] of linkOf) { const o = sheetNames.get(k); (byCoachDays.get(id) || byCoachDays.set(id, []).get(id)).push({ coach: o.coach, days: o.days }); }
  for (const p of clMonths.filter((q) => q <= now)) {
    const s0 = serial(p + "-01"), T = Math.min(serial(p + "-" + String(daysIn(p)).padStart(2, "0")), today);
    for (const [id, d] of daysOf) {
      const c = clients.get(id); if (!c || !sumDays(d, s0, T)) continue;
      const L = ledger(c, d, T, s0); if (!L || !L.overRecent) continue;
      let best = null, bw = 0;
      for (const x of byCoachDays.get(id) || []) { const w = sumDays(x.days, s0, T); if (w > bw) { bw = w; best = x.coach; } }
      if (!best) continue;
      const b = upM[best] || (upM[best] = {}); b[p] = (b[p] || 0) + 1;
    }
  }

  /* 10. Renewals: on time or after the last paid session ----
     A renewal is on time when the client trained no session after the last paid session of the
     package before it and before the day the renewal was bought. Filed in the month the renewal
     was bought, under the coach of the package that ran out. Only renewals the coach sheets can
     follow (linked clients, packages inside the sheets' months) are counted; n/N below shows how
     many that was. */
  const rt = {}, rtCover = { measured: 0, all: 0, red: 0, misfit: 0 };
  for (const c of inner.clients) {
    const own = c.pk.filter((p) => p.date <= today);
    for (const p of own) if (p.type === "Renewal" && iso(p.date) >= iso(inner.loggedFrom)) rtCover.all++;
    const d = daysOf.get(c.id); if (!d) continue;
    const L = ledger(c, d); if (!L) continue;
    rtCover.misfit += L.misfit;
    for (let q = 1; q < L.pkgs.length; q++) {
      const A = L.pkgs[q - 1], Bp = L.pkgs[q];
      if (Bp.type !== "Renewal" || Bp.date < inner.loggedFrom) continue;
      // Not used up when the renewal was bought: on time. Used up: any session between its last
      // session and the purchase day means the client trained unpaid.
      const lateS = A.last == null || A.last >= Bp.date ? 0 : L.ses.filter(([day]) => day > A.last && day < Bp.date).reduce((t, x) => t + x[1], 0);
      const coach = A.coach || c.coach || "(no coach)", m = iso(Bp.date).slice(0, 7);
      const o = (rt[coach] || (rt[coach] = {}))[m] || (rt[coach][m] = { n: 0, ok: 0, red: 0 });
      o.n++; if (!lateS) o.ok++; if (A.red) { o.red++; rtCover.red++; }
      rtCover.measured++;
    }
  }

  /* 11. Money left on the table ----
     Per coach, month by month: [count, AED without VAT]. Only clients whose names are linked and
     packages bought inside the months the coach sheets cover can be followed.
       xf  a package used after its expiry date with no extension fee logged: one fee each, in the
           month of the first session after expiry (the policy: the sessions expire, or the client
           pays the fee). An extension typed without a fee counts here too.
       xc  the same with a fee logged: charged.
       ax  the sessions trained after the expiry date on those packages without a fee, by the month
           trained, at the package's rate per session: what the policy would have turned into a fee
           or a new package ("up to").
       ul  sessions trained after every paid session by clients who then ended without buying
           again: never paid. (When a client renews, those sessions come off the new package.)
       fl  packages used up after their expiry date: count, AED = price x (days taken / days
           allowed - 1): the renewals that would have come in the same time at the policy's pace.
           fd = days past expiry added up. Month of the last session.
       fo  packages used up by their expiry date.
       ro  packages used up by their expiry date: [renewed, decided], by the month of expiry, once
           the grace period after that date has passed.
       rl  packages not used up by their expiry date: [renewed, decided].
       lc  clients lost: their cover ran out with no renewal in the grace period (or they told us
           they were stopping), have not bought since and are not training on a sheet: count and
           their last package's price, month the cover ran out. lb = came back later; ls = the file
           says lost but a sheet shows them training in the last 31 days.
     now.tn: sessions trained beyond every paid session by clients still with us (they come off the
     next package if the client renews); now.due: clients in the grace period, last package's price. */
  const VAT = 1 + (inner.vat ?? 5) / 100, FEE = (inner.extFee ?? 200) / VAT;
  const lk = {}, mOf = (day) => iso(day).slice(0, 7);
  const put = (coach, p, f, a, b = 0) => { const o = (lk[coach] || (lk[coach] = {}))[p] || (lk[coach][p] = {}); const v = o[f] || (o[f] = [0, 0]); v[0] += a; v[1] += b; };
  const leakNow = { tn: {}, due: {} };
  const dbgLeak = {};
  const ulList = [];
  let lkPk = 0;
  for (const c of inner.clients) {
    const d = daysOf.get(c.id); if (!d) continue;
    const L = ledger(c, d); if (!L) continue;
    const coachOf = (q) => q.coach || c.coach || "(no coach)";
    L.pkgs.forEach((q, i) => {
      if (q.sh || q.rule == null || q.rule >= today) return;
      lkPk++;
      const rate = q.px && q.size ? q.px / q.size : 0, k = coachOf(q), charged = q.fee > 0;
      if (q.ax > 0) {
        put(k, mOf(q.axL[0][0]), charged ? "xc" : "xf", 1, charged ? q.fee / VAT : FEE);
        if (!charged) for (const [day, w] of q.axL) put(k, mOf(day), "ax", w, w * rate);
      }
      // A package bought before the sheets start was partly used where the sheets cannot see, so
      // when it was used up is not known: it is left out of "used up late" and the renewal split.
      const seen = q.date >= W0;
      const usedUp = seen && q.closed && q.last != null;
      if (usedUp && q.last > q.rule) {
        const allowed = q.rule - q.date;
        put(k, mOf(q.last), "fl", 1, q.px && allowed > 0 ? q.px * ((q.last - q.date) / allowed - 1) : 0);
        put(k, mOf(q.last), "fd", q.last - q.rule);
      } else if (usedUp) put(k, mOf(q.last), "fo", 1);
      // Renewed or not, by whether the package was used up by its expiry date.
      const next = L.pkgs[i + 1];
      const renewed = next ? next.type !== "Returning" : c.status === "Ended" ? false : null;
      // Only once the grace period after the expiry date has passed: before that, renewals are in
      // and losses are not, which would make every recent month look like 100%.
      if (seen && renewed != null && q.rule + (inner.grace ?? 90) <= today) put(k, mOf(q.rule), usedUp && q.last <= q.rule ? "ro" : "rl", renewed ? 1 : 0, 1);
    });
    if (L.over > 0) {
      const q = L.pkgs[L.pkgs.length - 1], rate = q.px && q.size ? q.px / q.size : 0, k = coachOf(q);
      const lastS = L.ses.length ? L.ses[L.ses.length - 1][0] : null;
      // Sessions beyond every package in the file. Counted as never paid only for a client the file
      // has as ended who has not trained for a month; a client still training is more likely a
      // purchase not logged yet, so they go with "not paid yet". Either way the client is listed,
      // because the file only knows the purchases typed into it.
      const gone = c.status === "Ended" && lastS != null && lastS < today - 31;
      if (debug) (dbgLeak.all || (dbgLeak.all = [])).push({ name: c.name, id: c.id, status: c.status, over: L.over,
        pk: c.pk.map((p) => `${iso(p.date)} ${p.unit === "Months" ? p.size + "mo" : p.size} ${p.coach || ""} exp ${p.rule != null ? iso(p.rule) : "-"}${p.sh ? " shared" : ""}`),
        ledger: L.pkgs.map((p) => `${iso(p.date)} ${p.size} used ${p.used} last ${p.last != null ? iso(p.last) : "-"}`), from: L.ses.length ? iso(L.ses[0][0]) : null,
        months: Object.entries(L.ses.reduce((o, [day, w]) => { const m = iso(day).slice(0, 7); o[m] = (o[m] || 0) + w; return o; }, {})).map(([m, n]) => m.slice(2) + ":" + n).join(" "),
        sheets: (namesOf.get(c.id) || []).map((x) => x.coach + ": " + x.name),
        both: (() => { const cnt = new Map(); for (const [kk, id2] of linkOf) if (id2 === c.id) for (const [day] of sheetNames.get(kk).days) cnt.set(day, (cnt.get(day) || 0) + 1); return [...cnt].filter(([, n]) => n > 1).length; })() });
      ulList.push({ coach: k, name: c.name, n: L.over, aed: Math.round(L.over * rate), last: lastS != null ? iso(lastS) : null, pk: q.size + " on " + iso(q.date), used: q.last != null ? iso(q.last) : null, gone, status: c.status });
      if (gone) { for (const [day, w] of L.overL) put(k, mOf(day), "ul", w, w * rate);
        if (debug) (dbgLeak.ul || (dbgLeak.ul = [])).push({ id: c.id, name: c.name, coach: k, over: L.over, days: L.overL.map(([day, w]) => iso(day) + (w > 1 ? "x" + w : "")),
          pk: L.pkgs.map((q) => `${iso(q.date)} ${q.size}${q.coach ? " " + q.coach : ""} used ${q.used}${q.last ? " last " + iso(q.last) : ""}`),
          allPk: c.pk.map((q) => `${iso(q.date)} ${q.unit === "Months" ? q.size + "mo" : q.size} ${q.coach || ""}${q.sh ? " (shared)" : ""}`),
          sheets: (namesOf.get(c.id) || []).map((x) => x.coach + ": " + x.name), first: iso(L.ses[0] ? L.ses[0][0] : 0), told: c.told != null ? iso(c.told) : null }); }
      else { const o = leakNow.tn[k] || (leakNow.tn[k] = [0, 0]); o[0] += L.over; o[1] += L.over * rate; }
    }
  }
  // A client lost: the cover ran out with no renewal in the grace period. Not when they bought
  // again later (they came back: "lb"), and not when the coach sheets show them training in the
  // last 31 days (the file says ended, the sheet says otherwise: a purchase not logged, or
  // sessions on an old package past its expiry).
  for (const x of inner.decisions) {
    if (x.outcome !== "lost" || x.E > today) continue;
    const c = clients.get(x.id), k = x.coach || "(no coach)";
    if (c && c.pk.some((p) => p.date > x.E && !p.sh)) { put(k, mOf(x.E), "lb", 1, x.px || 0); continue; }
    const d = daysOf.get(x.id);
    if (d && [...d.keys()].some((day) => day > today - 31 && day <= today)) { put(k, mOf(x.E), "ls", 1, x.px || 0); continue; }
    put(k, mOf(x.E), "lc", 1, x.px || 0);
  }
  for (const c of inner.clients) if (c.status === "Renewal due") {
    const own = c.pk.filter((p) => p.date <= today), l = own.filter((p) => !p.sh).pop() || own.pop();
    const k = c.coach || (l && l.coach) || "(no coach)", o = leakNow.due[k] || (leakNow.due[k] = [0, 0]);
    o[0]++; o[1] += (l && l.px) || 0;
  }
  for (const k in lk) for (const p in lk[k]) for (const f in lk[k][p]) lk[k][p][f][1] = Math.round(lk[k][p][f][1]);
  for (const g of ["tn", "due"]) for (const k in leakNow[g]) leakNow[g][k][1] = Math.round(leakNow[g][k][1]);
  const leak = { fee: Math.round(FEE), feeIncl: inner.extFee ?? 200, from: clMonths[0] || null, m: lk, now: leakNow, packages: lkPk,
    // Each client behind "never paid" and "not paid yet", for checking against the purchases.
    list: ulList.sort((a, b) => b.n - a.n) };

  return {
    v: 1, today: iso(today), now, grace: inner.grace, loggedFrom: iso(inner.loggedFrom),
    // Each linked sheet name (coach|name words) -> the client's name in the Clients file, so the
    // Coach review page can show one name per client.
    ...(debug ? { _links: dbg, _leak: dbgLeak } : {}),
    sheetLink: Object.fromEntries([...linkOf].map(([k, id]) => [k, clients.get(id) ? clients.get(id).name : null])),
    names: { auto, byHand, ignored, unclear: unclear.sort((a, b) => (b.last || "").localeCompare(a.last || "") || a.coach.localeCompare(b.coach)), notInFile: notInFile.sort((a, b) => (b.last || "").localeCompare(a.last || "")), endedOnSheet, bad, hasTab: head >= 0 },
    inUse, up, months, perCoach, team, clientsOfCoach, leak,
    // Month by month, per coach, for the Coach review page's KPIs.
    kpiM: { ...(inner.monthly || { from: null, by: {} }), up: upM, rt, rtCover },
    // Active clients per coach right now, by the coach on the Clients tab (a handoff updates it).
    activeByCoach: inner.clients.filter((c) => c.status === "Active").reduce((o, c) => {
      const k = c.coach || ((c.pk.filter((p) => p.date <= today).pop() || {}).coach) || "(no coach)"; o[k] = (o[k] || 0) + 1; return o; }, {}),
  };
}
