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
  // Every amount on the page is without VAT, like the coach sheets' rates. Package prices are
  // typed with VAT, so they are divided by (1 + VAT). A "VAT %" row on Settings overrides 5.
  const vat = num(label(/^vat/i)) ?? 5;
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
        if (resolved <= today) decisions.push({ id: c.id, coach: owner.coach, E: toldStop ? Math.min(E, c.told) : E, outcome: "lost", at: resolved, joined: c.joined, told: toldStop, size: owner.size, unit: owner.unit });
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
  const activeOn = (d) => { const s = new Set(); for (const c of clients.values()) if (statusOn(c, d).status === "Active") s.add(c.id); return s; };
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
      let nw = 0, back = 0, out = 0;
      for (const id of A1) if (!A0.has(id)) { const c = clients.get(id); if (c.joined != null && c.joined >= start && c.joined <= endM) nw++; else back++; }
      // Everyone whose package ended this month: did they buy again within the grace period
      // (came back later), not (left), or is it too early to say (still deciding)?
      let outBack = 0, outLeft = 0, outWait = 0;
      for (const id of A0) if (!A1.has(id)) {
        out++;
        const c = clients.get(id);
        const again = c.pk.concat(c.sh).some((p) => p.date >= start && p.date <= Math.min(today, endM + grace));
        if (again) outBack++; else if (endM + grace > today) outWait++; else outLeft++;
      }
      const inMonth = packages.filter((p) => p.date >= start && p.date <= endM && own(p));
      const by = (t) => inMonth.filter((p) => p.type === t);
      const sum = (l) => Math.round(l.reduce((a, p) => a + (p.px || 0), 0));
      const coach = {};
      for (const p of inMonth) {
        const k = p.coach || "(no coach)"; const o = coach[k] || (coach[k] = { s: 0, n: 0 });
        o.s += p.px || 0; if (p.type === "New") o.n++;
      }
      for (const k in coach) coach[k].s = Math.round(coach[k].s);
      months.push({
        p: `${y}-${String(m).padStart(2, "0")}`, partial: at < endM,
        start: A0.size, nw, back, out, outBack, outLeft, outWait, end: A1.size,
        pk: { ren: by("Renewal").length, nw: by("New").length, ret: by("Returning").length },
        aed: { ren: sum(by("Renewal")), nw: sum(by("New")), ret: sum(by("Returning")) },
        // Sessions bought, so a fall in renewals can be told apart from clients moving to bigger packages.
        ses: Object.fromEntries([["ren", "Renewal"], ["nw", "New"], ["ret", "Returning"]].map(([k, t]) => [k, by(t).filter((p) => p.unit !== "Months").reduce((a, p) => a + (p.size || 0), 0)])),
        coach,
      });
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
    sz: x.unit === "Months" ? null : x.size, nsz: x.nunit === "Months" ? null : (x.nsize ?? null), early: x.at != null && x.outcome !== "lost" ? x.at <= x.E : null,
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
  const mix = { now: mixFor(activeNow), at: {} };
  for (const [k, d] of [["3", back(3)], ["6", back(6)], ["12", back(12)], ["y", toSerial(ty, 1, 1)]]) {
    mix.at[k] = { d: iso(d), n: 0, ...mixFor([...clients.values()].filter((c) => statusOn(c, d).status === "Active")) };
    mix.at[k].n = Object.values(mix.at[k].service).reduce((a, b) => a + b, 0);
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
    kpi, months, up, coaches, dec, cohort, cl, mix, checks,
  };
  // For the refresh job only, never written to the Clients page: what the coach sheets are
  // joined to for the Overview and Coach review pages (which are locked with the Performance
  // passphrase, not this one).
  if (opts.internal) out._in = {
    today, grace, loggedFrom, vat,
    clients: [...clients.values()].map((c) => ({
      id: c.id, name: c.name, coach: c.coach, status: c.now.status, P: c.now.P, joined: c.joined, told: c.told, agreed: c.agreed,
      pk: c.pk.concat(c.sh).sort((a, b) => a.date - b.date).map((p) => ({ date: p.date, end: p.end, size: p.size, unit: p.unit, px: p.px, coach: p.coach, coaches: [p.coach, p.coach2, p.coach3].filter(Boolean), type: p.type })),
    })),
    decisions,
  };
  // For the self-test only: which clients landed where. Never written to the page's data.
  if (opts.debug) out.debug = {
    due: upDue.map((c) => c.id).sort(), soon: upSoon.map((c) => c.id).sort(),
    decisions: decisions.map((x) => [x.id, x.outcome]),
  };
  return out;
}
