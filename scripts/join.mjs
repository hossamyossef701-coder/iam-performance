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
export function buildOverview({ inner, periods, links = [], now, cutoffDay, revenueOf }) {
  const today = inner.today;
  const clients = new Map(inner.clients.map((c) => [c.id, c]));
  const fullW = new Map(inner.clients.map((c) => [c.id, words(c.name)]));
  const sq = (s) => String(s || "").toLowerCase().replace(/[^a-z]/g, "");

  /* 1. Every name on every coach's sheet, with the days it trained ---- */
  const sheetNames = new Map();                       // "coach|key" -> {coach, name, key, months:Set, days:Map(serial->w)}
  for (const p of Object.keys(periods).sort()) {
    for (const coach of Object.keys(periods[p])) {
      const cl = periods[p][coach].cl; if (!cl) continue;
      const last = p === now ? cutoffDay : daysIn(p);
      for (const [name, , m] of cl) {
        const key = nameKey(name); if (!key) continue;
        const k = coach + "|" + key;
        const o = sheetNames.get(k) || { coach, name, key, months: new Set(), days: new Map() };
        o.name = name; o.months.add(p);
        for (let i = 0; i < Math.min(last, m.length); i++) {
          const w = wOf(m[i]); if (!w) continue;
          const d = serial(`${p}-${String(i + 1).padStart(2, "0")}`);
          o.days.set(d, (o.days.get(d) || 0) + w);
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
      if (!coach || !name) continue;
      if (id && id !== "NONE" && !clients.has(id)) { bad.push(`${coach}: "${name}" points to ${id}, which is not on the Clients tab`); continue; }
      confirmed.set(sq(coach) + "|" + nameKey(name), id && id !== "NONE" ? id : null);
    }
  }

  /* 3. Match each name ---- */
  const coachHad = (c, coach) => c.coach === coach || c.pk.some((p) => p.coaches.includes(coach));
  const linkOf = new Map();                            // "coach|key" -> id
  const unclear = [], notInFile = [];
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
    const pickOne = one(scored.filter((s) => s.exact && s.had))
      || one(scored.filter((s) => s.ex === 1 && s.had))
      || (scored.some((s) => s.ex === 1 && s.had) ? null : one(scored.filter((s) => s.f === 1 && s.had)));
    const lastMonth = [...o.months].sort().pop();
    if (pickOne) { linkOf.set(k, pickOne.id); auto++; continue; }
    if (scored.length) {
      unclear.push({ coach: o.coach, name: o.name, last: lastMonth,
        sug: scored.slice(0, 3).map((s) => ({ id: s.id, name: clients.get(s.id).name, why: s.had ? "had a package with " + o.coach : "never had a package with " + o.coach })) });
    } else notInFile.push({ coach: o.coach, name: o.name, last: lastMonth });
  }

  /* 4. Each linked client's days, all coaches together ---- */
  const daysOf = new Map();                            // id -> Map(serial -> w)
  const namesOf = new Map();                           // id -> [{coach,name}]
  for (const [k, id] of linkOf) {
    const o = sheetNames.get(k);
    const d = daysOf.get(id) || new Map();
    for (const [day, w] of o.days) d.set(day, (d.get(day) || 0) + w);
    daysOf.set(id, d);
    (namesOf.get(id) || namesOf.set(id, []).get(id)).push({ coach: o.coach, name: o.name, months: o.months });
  }
  const sumDays = (d, from, to) => { let s = 0; if (d) for (const [day, w] of d) if (day >= from && day <= to) s += w; return s; };

  // Clients still on this month's sheets whom the file has as ended.
  const endedOnSheet = [];
  for (const [k, id] of linkOf) {
    const o = sheetNames.get(k), c = clients.get(id);
    if (o.months.has(now) && c.status === "Ended") endedOnSheet.push({ coach: o.coach, name: o.name, id, until: c.P != null ? iso(c.P) : null });
  }

  /* 5. Packages in use, per client ---- */
  const inUse = [];
  for (const c of inner.clients) {
    if (!["Active", "Renewal due", "Ended"].includes(c.status)) continue;
    const d = daysOf.get(c.id);
    const cur = c.pk.filter((p) => p.date <= today).pop();
    const coach = (cur && cur.coach) || c.coach;
    const sheet = (namesOf.get(c.id) || []).find((x) => x.coach === coach) || (namesOf.get(c.id) || [])[0];
    const row = { id: c.id, name: c.name, coach, sheetName: sheet ? sheet.name : null, linked: !!d, status: c.status };
    if (c.status === "Active" && cur && cur.unit !== "Months" && cur.size) {
      const used = sumDays(d, cur.date, today), left = cur.size - used;
      const pace = sumDays(d, today - 27, today) / 4;
      const weeksLeft = Math.max(0, (cur.end - today) / 7);
      Object.assign(row, { size: cur.size, start: iso(cur.date), end: iso(cur.end), used, left, pace: +pace.toFixed(2), weeksLeft: +weeksLeft.toFixed(1),
        state: !d ? "unlinked" : left <= 2 ? "nearly" : left - pace * weeksLeft > 1 ? "short" : "ok" });
      inUse.push(row);
    } else if (c.status === "Active" && cur) {
      inUse.push(Object.assign(row, { unit: cur.unit, end: iso(cur.end), state: !d ? "unlinked" : "ok" }));
    } else if (c.P != null && d) {
      // Ran out, but trained after the package ended (in the last 31 days).
      const after = sumDays(d, Math.max(c.P + 1, today - 30), today);
      if (after > 0) inUse.push(Object.assign(row, { end: iso(c.P), after, state: "expired" }));
    }
  }

  // Up for renewal right now: cover from a package ending in the next 30 days, or run out within
  // the grace period with nothing new. Valued at the client's last package.
  const lastPx = (c) => { const l = c.pk.filter((p) => p.date <= today).pop(); return (l && l.px) || 0; };
  const soon = inner.clients.filter((c) => c.status === "Active" && c.P != null && c.P - today >= 0 && c.P - today <= 30 && c.pk.some((p) => p.end === c.P));
  const due = inner.clients.filter((c) => c.status === "Renewal due");
  const up = { n: soon.length + due.length, due: due.length, aed: Math.round([...soon, ...due].reduce((t, c) => t + lastPx(c), 0)) };

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

  return {
    v: 1, today: iso(today), now, grace: inner.grace, loggedFrom: iso(inner.loggedFrom),
    names: { auto, byHand, ignored, unclear: unclear.sort((a, b) => (b.last || "").localeCompare(a.last || "") || a.coach.localeCompare(b.coach)), notInFile: notInFile.sort((a, b) => (b.last || "").localeCompare(a.last || "")), endedOnSheet, bad, hasTab: head >= 0 },
    inUse, up, months, perCoach, team, clientsOfCoach,
  };
}
