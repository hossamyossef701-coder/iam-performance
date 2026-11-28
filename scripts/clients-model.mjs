/* The Clients page's figures, worked out from "I AM - Clients.xlsx".

   Only what Vanessa types is read: the Clients, Packages, Coach Changes, Lists and Settings
   tabs. Every status, end date and package type is then worked out here with the same rules
   the file's own formulas use - so the page is right even on a day nobody has opened the file
   (Excel only refreshes TODAY() when someone does), and does not depend on the file's views.

   Columns are found by their headings, not their positions, so a column added or moved in the
   file does not break the read. Mobile numbers and emails are never read at all.

   Dates are handled as Excel day numbers throughout (1 = 1 Jan 1900), so there is no time zone
   to get wrong: "today" is the date in Dubai. */

const DAY_MS = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);
const toSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS);
// A renewal is on time when bought before the package expires or within this many days after.
export const ON_TIME_DAYS = 7;
export const iso = (n) => (n == null ? null : new Date(EPOCH + n * DAY_MS).toISOString().slice(0, 10));
const monthOf = (n) => iso(n).slice(0, 7);
function edate(n, months) {                        // Excel's EDATE: same day, n months later
  const d = new Date(EPOCH + n * DAY_MS);
  const y = d.getUTCFullYear(), m = d.getUTCMonth() + months, day = d.getUTCDate();
  const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return Math.round((Date.UTC(y, m, Math.min(day, last)) - EPOCH) / DAY_MS);
}
export function dubaiToday(now = new Date()) {
  const p = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Dubai", year: "numeric", month: "2-digit", day: "2-digit" }).format(now);
  const [y, m, d] = p.split("-").map(Number);
  return toSerial(y, m, d);
}

// A cell as a day number: Excel dates arrive as numbers; typed text dates are accepted too.
function dayNum(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number" && isFinite(v)) return Math.floor(v);
  if (v instanceof Date && !isNaN(v)) return toSerial(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
  const t = String(v).trim();
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return toSerial(+m[1], +m[2], +m[3]);
  const MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  m = t.match(/^(\d{1,2})[\s-]([A-Za-z]{3})[A-Za-z]*[\s-](\d{4})$/);
  if (m && MON[m[2].toLowerCase()]) return toSerial(+m[3], MON[m[2].toLowerCase()], +m[1]);
  return null;
}
const num = (v) => (v == null || v === "" ? null : (typeof v === "number" ? v : (isFinite(+String(v).replace(/,/g, "")) ? +String(v).replace(/,/g, "") : null)));
const txt = (v) => (v == null ? "" : String(v).trim());
const key = (h) => txt(h).replace(/★/g, "").replace(/\s+/g, " ").trim().toLowerCase();

// A tab as rows of {heading: value}, starting below the row that holds the given headings.
function table(rows, must) {
  const want = must.map((m) => m.toLowerCase());
  for (let r = 0; r < Math.min(rows.length, 12); r++) {
    const hs = (rows[r] || []).map(key);
    if (want.every((w) => hs.includes(w))) {
      const cols = {};
      hs.forEach((h, i) => { if (h && cols[h] == null) cols[h] = i; });
      const out = [];
      for (let i = r + 1; i < rows.length; i++) {
        const row = rows[i] || [];
        const o = {};
        for (const h of Object.keys(cols)) o[h] = row[cols[h]];
        out.push(o);
      }
      return { cols, rows: out };
    }
  }
  throw new Error(`no heading row with ${must.join(", ")}`);
}
const pick = (o, ...names) => { for (const n of names) if (o[n] !== undefined) return o[n]; return undefined; };

export function buildClientsModel(sheets, opts = {}) {
  const today = opts.today ?? dubaiToday();
  const withNames = !!opts.names;
  const notes = [];

  /* ---- Settings: the six labelled values, then the dated package rules ---- */
  const S = sheets.Settings || [];
  const label = (re) => { for (const r of S) if (r && re.test(txt(r[0]))) return r[1]; return undefined; };
  const grace = num(label(/^grace period/i)) ?? 90;
  const goLive = dayNum(label(/^go-live date/i)) ?? today;
  const carryDays = num(label(/^carry-over period/i)) ?? 120;
  const loggedFrom = dayNum(label(/^every purchase logged since/i)) ?? toSerial(2026, 1, 1);
  const reminder = num(label(/^expiring soon window/i)) ?? 10;
  // The fee for extending a package past its expiry date (AED incl. VAT, like package prices).
  // A Settings row labelled "Extension fee" overrides it.
  const extFee = num(label(/^extension fee/i)) ?? 200;
  // Every amount on the page is without VAT, like the coach sheets' rates. Package prices are
  // typed with VAT, so they are divided by (1 + VAT). A "VAT %" row on Settings overrides 5.
  // The Settings row may hold a percent (5), a percent cell (0.05) or a multiplier (1.05).
  const vatRaw = num(label(/^vat/i));
  const vat = vatRaw == null ? 5 : vatRaw < 1 ? vatRaw * 100 : vatRaw < 2 ? (vatRaw - 1) * 100 : vatRaw;
  const exVat = (v) => (v == null ? null : v / (1 + vat / 100));
  const rules = [];
  let ruleHead = S.findIndex((r) => r && /^sessions$/i.test(txt(r[0])));
  if (ruleHead >= 0) for (let i = ruleHead + 1; i < S.length; i++) {
    const r = S[i] || [];
    const size = num(r[0]), days = num(r[1]), from = dayNum(r[2]);
    if (size != null && days != null) rules.push({ size, days, from: from ?? 0 });
  }
  if (!rules.length) notes.push("The Settings tab has no package-length table; every package is taken as 7 days.");

  // The rule in force on the purchase date: the largest listed size up to this one, and for
  // that size the row with the latest start date on or before the purchase.
  function validDays(size, date) {
    const live = rules.filter((x) => x.size <= size && x.from <= date);
    if (!live.length) return 7;
    const s = Math.max(...live.map((x) => x.size));
    const forSize = live.filter((x) => x.size === s);
    const from = Math.max(...forSize.map((x) => x.from));
    return Math.max(...forSize.filter((x) => x.from === from).map((x) => x.days));
  }

  /* ---- Lists: coaches and whether they are current ---- */
  const coachStatus = {};
  try {
    const L = table(sheets.Lists || [], ["Coach", "Coach status"]);
    for (const r of L.rows) { const c = txt(r["coach"]); if (c) coachStatus[c] = txt(r["coach status"]) || "Current"; }
  } catch { notes.push("The Lists tab has no Coach column; coaches are taken from the clients and packages."); }

  /* ---- Clients ---- */
  const CT = table(sheets.Clients || [], ["Client ID", "Full name"]);
  const clients = new Map();
  for (const r of CT.rows) {
    const id = txt(r["client id"]).toUpperCase();
    if (!/^C\d+$/.test(id)) continue;
    clients.set(id, {
      id, name: txt(r["full name"]),
      gender: txt(r["gender"]), nat: txt(r["nationality"]), lang: txt(r["language"]),
      coach: txt(r["coach"]), service: txt(r["service"]),
      dob: dayNum(r["date of birth"]), ageOld: num(r["age (old list)"]),
      legacy: txt(r["carried-over status"]),
      told: dayNum(pick(r, "end date (if they told us)", "end date")),
      reason: txt(r["end reason"]),
      // Optional column: the sessions a week the client signed up for, when the file has it.
      agreed: num(pick(r, "agreed sessions per week", "sessions per week (agreed)", "agreed sessions / week")),
      fileStatus: txt(r["status"]),                 // Excel's own answer, kept only to cross-check
      pk: [], sh: [],
    });
  }

  /* ---- Packages ---- */
  const PT = table(sheets.Packages || [], ["Date", "Client ID", "Size"]);
  const packages = [];
  let unknownIds = 0, extNoFee = 0, offRule = new Set();
  for (const r of PT.rows) {
    const date = dayNum(r["date"]);
    const id = txt(r["client id"]).toUpperCase();
    if (date == null || !id) continue;
    const size = num(r["size"]), unit = txt(r["unit"]) || "Sessions";
    const ext = dayNum(r["extended until"]);
    const p = {
      date, id, size, unit, ext,
      price: num(pick(r, "price (aed incl. vat)", "price")),
      px: exVat(num(pick(r, "price (aed incl. vat)", "price"))),
      coach: txt(r["coach"]), coach2: txt(r["2nd coach"]), coach3: txt(r["3rd coach"]),
      src: txt(pick(r, "source (how they found us)", "source")), by: txt(r["brought in by"]),
      shared: txt(pick(r, "shared with (client ids)", "shared with")).toUpperCase().match(/C\d+/g) || [],
      recorded: txt(pick(r, "hlp: type in the old files")),
      fee: num(r["extension fee (aed)"]),
    };
    if (size == null && ext == null) continue;
    if (unit === "Months") p.end = ext ?? edate(date, size);
    else {
      p.days = validDays(size || 0, date);
      if (size != null && !rules.some((x) => x.size === size && x.from <= date)) offRule.add(size);
      p.end = ext ?? date + p.days;
    }
    if (ext != null && p.fee == null) extNoFee++;
    packages.push(p);
    const c = clients.get(id);
    if (!c) { unknownIds++; continue; }
    c.pk.push(p);
    for (const s of p.shared) { const o = clients.get(s); if (o && s !== id) o.sh.push(p); }
  }
  packages.sort((a, b) => a.date - b.date);
  for (const c of clients.values()) { c.pk.sort((a, b) => a.date - b.date); c.sh.sort((a, b) => a.date - b.date); }

  /* ---- Coach changes ---- */
  const changes = [];
  try {
    const CC = table(sheets["Coach Changes"] || [], ["Date", "Client ID", "From coach", "To coach"]);
    for (const r of CC.rows) {
      const d = dayNum(r["date"]), id = txt(r["client id"]).toUpperCase();
      if (d != null && id) changes.push({ date: d, id, from: txt(r["from coach"]), to: txt(r["to coach"]), why: txt(r["why"]) });
    }
  } catch { /* the tab is optional */ }

  /* ---- Status of one client on a given day, from purchases made up to that day ----
     The same rules as the Clients tab's Covered until / Status columns. */
  function statusOn(c, d) {
    const all = c.pk.concat(c.sh).filter((p) => p.date <= d);
    const X = all.length ? Math.max(...all.map((p) => p.end)) : 0;
    const AK = all.length ? Math.max(...all.map((p) => p.date)) : 0;
    const told = c.told != null && c.told <= d ? c.told : null;
    let carry = 0;
    if (!(told != null || X >= goLive || AK >= loggedFrom)) {
      if (c.legacy === "Active") carry = goLive + carryDays;
      else if (c.legacy === "Renewal due") carry = goLive;
    }
    const P = Math.max(X, carry) || null;
    const stop = told != null && told >= AK;
    let status;
    if (stop) status = "Ended";
    else if (P == null) status = c.legacy || "No package";
    else if (P >= d) status = "Active";
    else if (d - P <= grace) status = "Renewal due";
    else status = "Ended";
    return { status, P, X, AK, stop, fromPackages: X >= (P || 0) && X > 0, last: all.length ? all.reduce((a, b) => (b.date >= a.date ? b : a)) : null };
  }

  /* ---- Each purchase's type: New / Renewal / Returning - the Packages tab's Type column ---- */
  for (const c of clients.values()) {
    for (const p of c.pk) {
      const before = c.pk.filter((q) => q.date < p.date);
      const Z = before.length ? Math.max(...before.map((q) => q.end)) : 0;
      const AA = before.length ? Math.max(...before.map((q) => q.date)) : 0;
      const toldBefore = c.told != null && c.told < p.date && c.told >= AA;
      if (toldBefore) p.type = "Returning";
      else if (!Z) {
        if (p.recorded) p.type = p.recorded;
        else if (p.date < goLive) p.type = "New";
        else p.type = c.legacy === "Ended" ? "Returning" : (c.legacy ? "Renewal" : "New");
      } else p.type = p.date - Z > grace ? "Returning" : "Renewal";
    }
  }

  /* ---- Every client today ---- */
  const now = [];
  let agree = 0, compared = 0;
  for (const c of clients.values()) {
    const s = statusOn(c, today);
    const all = c.pk.concat(c.sh);
    c.now = s;
    c.joined = all.length ? Math.min(...all.map((p) => p.date)) : null;
    c.endedOn = s.status !== "Ended" ? null : (s.stop ? c.told : (s.P != null ? s.P + grace : null));
    if (["Active", "Renewal due", "Ended", "No package"].includes(c.fileStatus)) { compared++; if (c.fileStatus === s.status) agree++; }
    now.push(c);
  }

  /* ---- Renewal decisions ----
     Each time a client's cover runs out, they either buy again within the grace period
     (kept, or moved if the new package is with another coach) or they do not (lost). A
     renewal bought early counts as kept too. Each decision belongs to the coach on the
     package that was running out. Only cover that ran out after every purchase was being
     logged (Settings: "Every purchase logged since") is judged - before that, renewals were
     not recorded and would look like losses. */
  const decisions = [];
  for (const c of clients.values()) {
    const all = c.pk.concat(c.sh).sort((a, b) => a.date - b.date);
    if (!all.length) continue;
    let E = all[0].end, owner = all[0];
    for (let i = 1; i <= all.length; i++) {
      const next = all[i];
      const toldStop = c.told != null && c.told <= E + grace && (!next || c.told <= next.date) && c.told >= owner.date;
      if (next && next.date <= E + grace && !toldStop) {
        if (E >= loggedFrom) decisions.push({ id: c.id, coach: owner.coach, E, outcome: next.coach && owner.coach && next.coach !== owner.coach ? "moved" : "kept", at: next.date, to: next.coach, joined: c.joined,
          size: owner.size, unit: owner.unit, nsize: next.size, nunit: next.unit });
        if (next.end >= E) { E = next.end; owner = next; }
        continue;
      }
      // Cover ran out with no renewal in time (or the client told us they were stopping).
      // Cover still running is no decision yet; cover that ran out inside the grace period is
      // pending - the client is Renewal due.
      if (!toldStop && E >= today) { if (!next) break; E = next.end; owner = next; continue; }
      if (E >= loggedFrom || toldStop) {
        const resolved = toldStop ? c.told : E + grace;
        if (resolved <= today) decisions.push({ id: c.id, coach: owner.coach, E: toldStop ? Math.min(E, c.told) : E, outcome: "lost", at: resolved, joined: c.joined, told: toldStop, size: owner.size, unit: owner.unit, px: owner.px });
        else decisions.push({ id: c.id, coach: owner.coach, E, outcome: "pending", at: null, joined: c.joined, size: owner.size, unit: owner.unit });
      }
      if (!next) break;
      E = next.end; owner = next;
    }
  }

  /* ---- From here on: the figures the page shows. Client IDs are replaced by a running
     number and names are never sent - the page shows patterns, not who to call. ---- */
  const idx = new Map([...clients.keys()].map((id, i) => [id, i]));
  const yearStart = toSerial(+iso(today).slice(0, 4), 1, 1);
  const own = (p) => clients.has(p.id);

  /* ---- Months: who was active at the start and end, and what changed in between ----
     Active at the start (the day before the 1st) + joined - left = active at the end, so
     every month adds up. "New" joined with their first package that month; everyone else
     who joined came back (a renewal after a gap, a returning client). */
  const firstMonth = Math.max(loggedFrom, toSerial(2025, 1, 1));
  // Who was active on a day, and with which coach (the coach of the package covering it).
  // Today the coach on the Clients tab counts (a handoff changes it before any new package);
  // on earlier days, the coach of the package that covered the day.
  const coachAt = (c, st, d) => (d >= today ? c.coach || (st.last && st.last.coach) : (st.last && st.last.coach) || c.coach) || "(no coach)";
  const activeOn = (d) => { const s = new Map(); for (const c of clients.values()) { const st = statusOn(c, d); if (st.status === "Active") s.set(c.id, coachAt(c, st, d)); } return s; };
  const blank = () => ({ start: 0, nw: 0, back: 0, out: 0, outBack: 0, outLeft: 0, outWait: 0, movIn: 0, movOut: 0, movFrom: {}, movTo: {}, end: 0,
    pk: { ren: 0, nw: 0, ret: 0 }, aed: { ren: 0, nw: 0, ret: 0 }, ses: { ren: 0, nw: 0, ret: 0 } });
  const tKey = { Renewal: "ren", New: "nw", Returning: "ret" };
  const months = [];
  {
    const d0 = new Date(EPOCH + firstMonth * DAY_MS);
    let y = d0.getUTCFullYear(), m = d0.getUTCMonth() + 1;
    let A0 = activeOn(toSerial(y, m, 1) - 1);
    for (;;) {
      const start = toSerial(y, m, 1);
      if (start > today) break;
      const endM = toSerial(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1) - 1;
      const at = Math.min(endM, today);
      const A1 = activeOn(at);
      // The same figures for each coach, so a page can leave a coach out and still add up.
      // A client who changed coach inside the month moves out of one and into the other.
      const by = {};
      const get = (k) => by[k] || (by[k] = blank());
      for (const [, k] of A0) get(k).start++;
      for (const [, k] of A1) get(k).end++;
      for (const [id, k] of A1) {
        if (A0.has(id)) {
          const was = A0.get(id);
          if (was !== k) { const a = get(k), b = get(was); a.movIn++; b.movOut++; a.movFrom[was] = (a.movFrom[was] || 0) + 1; b.movTo[k] = (b.movTo[k] || 0) + 1; }
          continue;
        }
        const c = clients.get(id);
        if (c.joined != null && c.joined >= start && c.joined <= endM) get(k).nw++; else get(k).back++;
      }
      // Everyone whose package ended this month: did they buy again within the grace period
      // (came back later), not (left), or is it too early to say (still deciding)?
      for (const [id, k] of A0) if (!A1.has(id)) {
        const o = get(k); o.out++;
        const c = clients.get(id);
        const again = c.pk.concat(c.sh).some((p) => p.date >= start && p.date <= Math.min(today, endM + grace));
        if (again) o.outBack++; else if (endM + grace > today) o.outWait++; else o.outLeft++;
      }
      const inMonth = packages.filter((p) => p.date >= start && p.date <= endM && own(p));
      const coach = {};
      for (const p of inMonth) {
        const k = p.coach || "(no coach)", t = tKey[p.type]; if (!t) continue;
        const o = get(k); o.pk[t]++; o.aed[t] += p.px || 0; if (p.unit !== "Months") o.ses[t] += p.size || 0;
        const c2 = coach[k] || (coach[k] = { s: 0, n: 0 }); c2.s += p.px || 0; if (p.type === "New") c2.n++;
      }
      for (const k in coach) coach[k].s = Math.round(coach[k].s);
      for (const k in by) for (const t of ["ren", "nw", "ret"]) by[k].aed[t] = Math.round(by[k].aed[t]);
      // The team: every coach added up (moves between coaches cancel out).
      const tot = blank();
      for (const k in by) {
        for (const f of ["start", "nw", "back", "out", "outBack", "outLeft", "outWait", "end"]) tot[f] += by[k][f];
        for (const g of ["pk", "aed", "ses"]) for (const t of ["ren", "nw", "ret"]) tot[g][t] += by[k][g][t];
      }
      delete tot.movIn; delete tot.movOut; delete tot.movFrom; delete tot.movTo;
      months.push({ p: `${y}-${String(m).padStart(2, "0")}`, partial: at < endM, ...tot, coach, by });
      if (at < endM) break;
      A0 = A1; m++; if (m > 12) { m = 1; y++; }
    }
  }

  /* ---- Up for renewal, right now: packages ending in the next 30 days, plus clients whose
     package ran out inside the grace period with nothing new. Cover from the move off the old
     lists is not a package, so it is not counted as ending. Valued at the client's last
     package, on the assumption they would buy the same again. ---- */
  const upWindow = 30;
  const lastPx = (c) => (c.now.last && c.now.last.px) || 0;
  const upSoon = [...clients.values()].filter((c) => c.now.status === "Active" && c.now.fromPackages && c.now.P - today >= 0 && c.now.P - today <= upWindow);
  const upDue = [...clients.values()].filter((c) => c.now.status === "Renewal due");
  const up = { days: upWindow, soon: upSoon.length, due: upDue.length, aed: Math.round([...upSoon, ...upDue].reduce((a, c) => a + lastPx(c), 0)), coach: {} };
  for (const c of [...upSoon, ...upDue]) {
    const k = c.coach || "(no coach)"; const o = up.coach[k] || (up.coach[k] = { n: 0, aed: 0, due: 0 });
    o.n++; o.aed += lastPx(c); if (c.now.status === "Renewal due") o.due++;
  }
  for (const k in up.coach) up.coach[k].aed = Math.round(up.coach[k].aed);

  /* ---- Renewal decisions, without names, for the page to count over any window ---- */
  const dec = decisions.map((x) => ({
    k: idx.get(x.id), c: x.coach || "(no coach)", E: iso(x.E), at: x.at == null ? null : iso(x.at), o: x.outcome,
    sz: x.unit === "Months" ? null : x.size, nsz: x.nunit === "Months" ? null : (x.nsize ?? null), early: x.at != null && x.outcome !== "lost" ? x.at <= x.E + ON_TIME_DAYS : null,
  }));

  /* ---- New clients since every purchase was logged: who they are, who coached them first,
     whether they bought a second package, and how long they stayed. One row per client. ---- */
  const ageOf = (c) => {
    const a = c.dob != null ? Math.floor((today - c.dob) / 365.25) : c.ageOld;
    if (a == null) return null;
    return a < 18 ? "Under 18" : a < 30 ? "18-29" : a < 40 ? "30-39" : a < 50 ? "40-49" : a < 60 ? "50-59" : "60+";
  };
  const decByClient = new Map();
  for (const x of decisions) { if (!decByClient.has(x.id)) decByClient.set(x.id, []); decByClient.get(x.id).push(x); }
  const cohort = [];
  for (const c of clients.values()) {
    const first = c.pk[0];
    if (!first || c.joined == null || c.joined < loggedFrom || first.type === "Renewal") continue;
    const ds = (decByClient.get(c.id) || []).sort((a, b) => a.E - b.E);
    const d1 = ds[0];
    const lost = ds.find((x) => x.outcome === "lost");
    const pend = ds.find((x) => x.outcome === "pending");
    // How long they stayed: to the end of the cover they did not renew, or still going.
    let stay, ev;
    if (lost) { stay = lost.E - c.joined; ev = 1; }
    else if (pend) { stay = pend.E - c.joined; ev = 0; }
    else { stay = today - c.joined; ev = 0; }
    cohort.push({
      k: idx.get(c.id), c: first.coach || "(no coach)", d: iso(c.joined), t: first.type,
      age: ageOf(c), g: ({ F: "Female", M: "Male" })[c.gender] || c.gender || null, svc: c.service || null,
      pkg: first.unit === "Months" ? `${first.size} mo` : (first.size != null ? `${first.size}` : null),
      nat: c.nat || null, lang: c.lang || null, src: first.src || null,
      second: !d1 || d1.outcome === "pending" ? null : (d1.outcome === "lost" ? 0 : 1),
      stay: Math.max(1, stay), ev,
      aed: Math.round(c.pk.reduce((a, p) => a + (p.px || 0), 0)),
    });
  }

  /* ---- Who each client is, by running number, so the page can split any count by type ---- */
  const cl = [...clients.values()].map((c) => ({
    age: ageOf(c), g: ({ F: "Female", M: "Male" })[c.gender] || c.gender || null, svc: c.service || null,
    nat: c.nat || null, lang: c.lang || null, src: (c.pk[0] && c.pk[0].src) || null,
  }));

  /* ---- Coaches ---- */
  const coachNames = new Set([...Object.keys(coachStatus), ...[...clients.values()].map((c) => c.coach).filter(Boolean)]);
  const coaches = [];
  for (const name of coachNames) {
    const mine = [...clients.values()].filter((c) => c.coach === name);
    const theirs = packages.filter((p) => p.coach === name && own(p));
    const row = {
      c: name, status: coachStatus[name] || "Current",
      active: mine.filter((c) => c.now.status === "Active").length,
      due: mine.filter((c) => c.now.status === "Renewal due").length,
      since: theirs.length ? iso(theirs[0].date) : null,           // their first package with us
    };
    if (row.status === "Former" && !row.active && !row.due) continue;
    coaches.push(row);
  }
  coaches.sort((a, b) => (a.status === "Former") - (b.status === "Former") || b.active - a.active || a.c.localeCompare(b.c));

  /* ---- Who the active clients are: today, and at the start of each period the page offers ---- */
  const activeNow = [...clients.values()].filter((c) => c.now.status === "Active");
  const tally = (list, f) => { const m = {}; for (const x of list) { const k = f(x) || "(not recorded)"; m[k] = (m[k] || 0) + 1; } return m; };
  const mixFor = (list) => ({
    service: tally(list, (c) => c.service),
    gender: tally(list, (c) => ({ F: "Female", M: "Male" })[c.gender] || c.gender),
    age: tally(list, ageOf),
    lang: tally(list, (c) => c.lang),
  });
  const ty = +iso(today).slice(0, 4), tm = +iso(today).slice(5, 7), td = +iso(today).slice(8, 10);
  const back = (n) => toSerial(ty, tm - n, td);       // Date.UTC rolls the month over the year
  // Also split by coach, so a page can leave a coach out.
  const coachOn = (c, d) => coachAt(c, statusOn(c, d), d);
  const byCoach = (list, d) => { const g = {}; for (const c of list) { const k = coachOn(c, d); (g[k] = g[k] || []).push(c); } const o = {}; for (const k in g) o[k] = mixFor(g[k]); return o; };
  const mix = { now: mixFor(activeNow), nowBy: byCoach(activeNow, today), at: {} };
  for (const [k, d] of [["3", back(3)], ["6", back(6)], ["12", back(12)], ["y", toSerial(ty, 1, 1)]]) {
    const list = [...clients.values()].filter((c) => statusOn(c, d).status === "Active");
    mix.at[k] = { d: iso(d), n: list.length, ...mixFor(list), by: byCoach(list, d) };
  }

  /* ---- Existing clients kept: of the clients active at the start of a period, how many are
     still active today, per coach (the coach they were with at the start). Unlike the renewal
     figures this needs no package to have run out, so it covers every client, long-standing
     ones included. A period that would start before purchases were fully logged starts there. ---- */
  const withUs = (d) => { const s = new Map(); for (const c of clients.values()) { const st = statusOn(c, d);
    if (st.status === "Active" || st.status === "Renewal due") s.set(c.id, { coach: coachAt(c, st, d), due: st.status === "Renewal due" }); } return s; };
  const kept = {};
  for (const [k, d0] of [["3", back(3)], ["6", back(6)], ["12", back(12)], ["y", toSerial(ty, 1, 1)]]) {
    // A client inside the grace period (Renewal due) is still a client on both dates: only one whose
    // grace period ran out, or who said they are stopping, counts as gone. wait = of those still
    // here, how many are inside the grace period today, so the figure can show what is not settled.
    const d = Math.max(d0, loggedFrom), then = withUs(d), nowA = withUs(today), by = {};
    for (const [id, coach] of then) {
      const o = by[coach.coach] || (by[coach.coach] = { start: 0, still: 0, same: 0, wait: 0 });
      o.start++; if (nowA.has(id)) { o.still++; if (nowA.get(id).coach === coach.coach) o.same++; if (nowA.get(id).due) o.wait++; }
    }
    kept[k] = { d: iso(d), by };
  }

  /* ---- Who referred a client: "Brought in by" names a coach (first name is enough), or a client,
     whose coach on the day of the purchase gets the referral. ---- */
  const refCoaches = new Set([...Object.keys(coachStatus), ...packages.map((x) => x.coach).filter(Boolean)]);
  const coachFirst = new Map(); for (const n of refCoaches) { const f = n.toLowerCase().split(/\s+/)[0]; if (!coachFirst.has(f)) coachFirst.set(f, n); }
  const clientByName = new Map(); for (const c of clients.values()) if (c.name) clientByName.set(c.name.toLowerCase().replace(/\s+/g, " ").trim(), c);
  const referrer = (pk) => {
    const b = (pk.by || "").toLowerCase().replace(/\s+/g, " ").trim(); if (!b) return null;
    for (const n of refCoaches) if (n.toLowerCase() === b) return n;
    const c = clientByName.get(b); if (c) return coachOn(c, pk.date) || null;
    return coachFirst.get(b.split(" ")[0]) || null;
  };

  /* ---- Month by month, per coach: the client figures the Coach review page tracks ----
     act = active at the month's end (today for the month running). ks/kk = clients with the coach
     the day before the month started / of them, still with us at its end (the grace period counts
     as still with us). nw = new clients. rn/re = packages that ran out in the month and were
     renewed / of them, renewed on or before the day they ran out. For the refresh job only. */
  const monthly = { from: iso(loggedFrom).slice(0, 7), by: {} };
  {
    let y = +iso(loggedFrom).slice(0, 4), m = +iso(loggedFrom).slice(5, 7);
    for (;;) {
      const s0 = toSerial(y, m, 1); if (s0 > today) break;
      const e = Math.min(toSerial(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1) - 1, today);
      const p = iso(s0).slice(0, 7);
      const cell = (c) => { const b = monthly.by[c || "(no coach)"] || (monthly.by[c || "(no coach)"] = {}); return b[p] || (b[p] = { act: 0, ks: 0, kk: 0, nw: 0, rn: 0, re: 0, ref: 0 }); };
      for (const [, k] of activeOn(e)) cell(k).act++;
      const W1 = withUs(e);
      for (const [id, x] of withUs(s0 - 1)) { const o = cell(x.coach); o.ks++; if (W1.has(id)) o.kk++; }
      // Counted as the Clients page counts them: each New package in the month.
      // nd: the day of the month of each, so a month running can be set against the same day before.
      for (const pk of packages) if (pk.type === "New" && pk.date >= s0 && pk.date <= e && own(pk)) { const o = cell(pk.coach); o.nw++; (o.nd || (o.nd = [])).push(pk.date - s0 + 1); }
      for (const x of decisions) if (x.E >= s0 && x.E <= e && (x.outcome === "kept" || x.outcome === "moved")) { const o = cell(x.coach); o.rn++; if (x.at <= x.E + ON_TIME_DAYS) o.re++; }
      // ref = referrals: New or Returning packages whose source is a referral, credited to the coach
      // who brought the client in (referrer(): a coach named in "Brought in by", or the coach of the
      // client named there), whoever the client then trains with.
      for (const pk of packages) if ((pk.type === "New" || pk.type === "Returning") && pk.date >= s0 && pk.date <= e && own(pk) && /referral/i.test(pk.src || "")) {
        const who = referrer(pk); if (who) cell(who).ref++;
      }
      if (m === 12) { y++; m = 1; } else m++;
    }
    /* Kept over a stretch of months, each client counted once: of the clients with the coach the
       day before the first month began, how many are still clients at the end of the last.
       span["2026-07|2026-09"].Hossam = [with the coach at the start, still with us at the end]. */
    const ps = [], starts = {}, ends = {};
    { let yy = +iso(loggedFrom).slice(0, 4), mm = +iso(loggedFrom).slice(5, 7);
      for (;;) { const s0 = toSerial(yy, mm, 1); if (s0 > today) break;
        const p = iso(s0).slice(0, 7); ps.push(p);
        starts[p] = withUs(s0 - 1); ends[p] = withUs(Math.min(toSerial(mm === 12 ? yy + 1 : yy, mm === 12 ? 1 : mm + 1, 1) - 1, today));
        if (mm === 12) { yy++; mm = 1; } else mm++; } }
    monthly.span = {};
    for (let i = 0; i < ps.length; i++) for (let j = i + 1; j < ps.length; j++) {
      const o = {}, W1 = ends[ps[j]];
      for (const [id, x] of starts[ps[i]]) { const c = x.coach || "(no coach)"; const v = o[c] || (o[c] = [0, 0]); v[0]++; if (W1.has(id)) v[1]++; }
      monthly.span[ps[i] + "|" + ps[j]] = o;
    }
  }

  /* ---- Headline numbers ---- */
  const monthStart = toSerial(+iso(today).slice(0, 4), +iso(today).slice(5, 7), 1);
  const thisMonth = packages.filter((p) => p.date >= monthStart && p.date <= today && own(p));
  const kpi = {
    active: activeNow.length,
    active30: [...clients.values()].filter((c) => statusOn(c, today - 30).status === "Active").length,
    due: upDue.length, ending: upSoon.length,
    ended: [...clients.values()].filter((c) => c.now.status === "Ended").length,
    noPkg: [...clients.values()].filter((c) => c.now.status === "No package").length,
    total: clients.size,
    newYear: new Set(packages.filter((p) => p.date >= yearStart && p.type === "New" && own(p)).map((p) => p.id)).size,
    salesMonth: Math.round(thisMonth.reduce((a, p) => a + (p.px || 0), 0)),
    nwMonth: new Set(thisMonth.filter((p) => p.type === "New").map((p) => p.id)).size,
    renMonth: thisMonth.filter((p) => p.type === "Renewal").length,
    carry: activeNow.filter((c) => !c.now.fromPackages).length,
    // The same two counts per coach, so a page can leave a coach out.
    activeBy: Object.fromEntries([...activeOn(today)].reduce((m, [, k]) => m.set(k, (m.get(k) || 0) + 1), new Map())),
    active30By: Object.fromEntries([...activeOn(today - 30)].reduce((m, [, k]) => m.set(k, (m.get(k) || 0) + 1), new Map())),
  };

  /* ---- Things in the file worth a look ---- */
  const checks = [];
  if (unknownIds) checks.push([`${unknownIds} package row(s) use a Client ID that is not on the Clients tab`, "fix"]);
  const noCoach = activeNow.filter((c) => !c.coach).length;
  if (noCoach) checks.push([`${noCoach} active client(s) have no coach on the Clients tab`, "fix"]);
  if (extNoFee) checks.push([`${extNoFee} extension(s) have no fee typed - 500, or 0 if waived`, "fix"]);
  const noPrice = packages.filter((p) => p.date >= loggedFrom && own(p) && p.px == null).length;
  if (noPrice) checks.push([`${noPrice} package(s) since ${iso(loggedFrom)} have no price, so they count as 0 in package sales`, "fix"]);
  const newYear = cohort.filter((x) => x.d >= iso(yearStart) && x.t === "New");
  const noSrc = newYear.filter((x) => !x.src).length;
  if (noSrc) checks.push([`How they found us is blank for ${noSrc} of ${newYear.length} new clients this year, so value by source is thin`, "info"]);
  if (offRule.size) checks.push([`Package size(s) ${[...offRule].sort((a, b) => a - b).join(", ")} have no rule of their own on Settings, so they take the next smaller size's length`, "info"]);
  if (kpi.carry) checks.push([`${kpi.carry} active client(s) are covered only by the move from the old lists, until ${iso(goLive + carryDays)}; their next renewal needs logging`, "info"]);
  if (compared) checks.push([`Statuses agree with the file's own for ${agree} of ${compared} clients` + (agree < compared ? " - the rest differ only because the file was last recalculated on an earlier day" : ""), agree === compared ? "ok" : "info"]);
  notes.forEach((n) => checks.push([n, "fix"]));

  const out = {
    v: 2, today: iso(today),
    settings: { grace, reminder, vat, goLive: iso(goLive), carryUntil: iso(goLive + carryDays), loggedFrom: iso(loggedFrom), newCoachMonths: num(label(/^new coach priority/i)) ?? 3 },
    kpi, months, up, coaches, dec, cohort, cl, mix, kept, checks,
  };
  // For the refresh job only, never written to the Clients page: what the coach sheets are
  // joined to for the Overview and Coach review pages (which are locked with the Performance
  // passphrase, not this one).
  if (opts.internal) out._in = {
    today, grace, loggedFrom, vat, extFee,
    clients: [...clients.values()].map((c) => ({
      id: c.id, name: c.name, coach: c.coach, status: c.now.status, P: c.now.P, joined: c.joined, told: c.told, agreed: c.agreed,
      pk: c.pk.concat(c.sh).sort((a, b) => a.date - b.date).map((p) => ({ date: p.date, end: p.end, size: p.size, unit: p.unit, px: p.px, coach: p.coach, coaches: [p.coach, p.coach2, p.coach3].filter(Boolean), type: p.type,
        // For the money left on the table: the expiry date the policy gives (before any
        // extension), the extension and its fee as logged, and whether the package is another
        // client's shared with this one (counted once, under its owner).
        rule: p.days != null ? p.date + p.days : null, ext: p.ext, fee: p.fee, sh: p.id !== c.id })),
    })),
    decisions, monthly,
  };
  // For the self-test only: which clients landed where. Never written to the page's data.
  if (opts.debug) out.debug = {
    due: upDue.map((c) => c.id).sort(), soon: upSoon.map((c) => c.id).sort(),
    decisions: decisions.map((x) => [x.id, x.outcome]),
  };
  return out;
}
