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
        if (E >= loggedFrom) decisions.push({ id: c.id, coach: owner.coach, E, outcome: next.coach && owner.coach && next.coach !== owner.coach ? "moved" : "kept", at: next.date, to: next.coach, joined: c.joined });
        if (next.end >= E) { E = next.end; owner = next; }
        continue;
      }
      // Cover ran out with no renewal in time (or the client told us they were stopping).
      // Cover still running is no decision yet; cover that ran out inside the grace period is
      // pending - the client is Renewal due.
      if (!toldStop && E >= today) { if (!next) break; E = next.end; owner = next; continue; }
      if (E >= loggedFrom || toldStop) {
        const resolved = toldStop ? c.told : E + grace;
        if (resolved <= today) decisions.push({ id: c.id, coach: owner.coach, E: toldStop ? Math.min(E, c.told) : E, outcome: "lost", at: resolved, joined: c.joined, told: toldStop });
        else decisions.push({ id: c.id, coach: owner.coach, E, outcome: "pending", at: null, joined: c.joined });
      }
      if (!next) break;
      E = next.end; owner = next;
    }
  }

  /* ---- Months: flows and the active count at each month end ---- */
  const firstMonth = Math.max(loggedFrom, toSerial(2025, 1, 1));
  const months = [];
  {
    let d0 = new Date(EPOCH + firstMonth * DAY_MS);
    let y = d0.getUTCFullYear(), m = d0.getUTCMonth() + 1;
    for (;;) {
      const start = toSerial(y, m, 1);
      if (start > today) break;
      const endM = toSerial(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1) - 1;
      const at = Math.min(endM, today);
      let active = 0;
      for (const c of clients.values()) if (statusOn(c, at).status === "Active") active++;
      const inMonth = (p) => p.date >= start && p.date <= endM;
      const distinct = (type) => new Set(packages.filter((p) => inMonth(p) && p.type === type && clients.has(p.id)).map((p) => p.id)).size;
      const lost = decisions.filter((x) => x.outcome === "lost" && x.E >= start && x.E <= endM).length;
      const pending = decisions.filter((x) => x.outcome === "pending" && x.E >= start && x.E <= endM).length;
      const sold = packages.filter(inMonth).length;
      months.push({ p: `${y}-${String(m).padStart(2, "0")}`, active, partial: at < endM, nw: distinct("New"), ret: distinct("Returning"), ren: distinct("Renewal"), lost, pending, sold });
      m++; if (m > 12) { m = 1; y++; }
    }
  }

  /* ---- Coaches ---- */
  const yearStart = toSerial(+iso(today).slice(0, 4), 1, 1);
  const since = today - 365;
  const coachNames = new Set([...Object.keys(coachStatus), ...now.map((c) => c.coach).filter(Boolean)]);
  const coaches = [];
  for (const name of coachNames) {
    const mine = now.filter((c) => c.coach === name);
    const dec = decisions.filter((x) => x.coach === name && (x.outcome === "pending" || x.at >= since));
    const kept = dec.filter((x) => x.outcome === "kept").length, moved = dec.filter((x) => x.outcome === "moved").length;
    const lost = dec.filter((x) => x.outcome === "lost"), pending = dec.filter((x) => x.outcome === "pending").length;
    const pk = packages.filter((p) => p.coach === name && p.date >= yearStart && clients.has(p.id));
    const tenure = lost.filter((x) => x.joined != null).map((x) => x.E - x.joined);
    const row = {
      c: name, status: coachStatus[name] || "Current",
      active: mine.filter((c) => c.now.status === "Active").length,
      due: mine.filter((c) => c.now.status === "Renewal due").length,
      exp: mine.filter((c) => c.now.status === "Active" && c.now.fromPackages && c.now.P - today >= 0 && c.now.P - today <= reminder).length,
      kept, moved, lost: lost.length, pending,
      rate: kept + lost.length ? kept / (kept + lost.length) : null,
      tenure: tenure.length ? Math.round(tenure.reduce((a, b) => a + b, 0) / tenure.length) : null,
      nw: new Set(pk.filter((p) => p.type === "New").map((p) => p.id)).size,
      ret: new Set(pk.filter((p) => p.type === "Returning").map((p) => p.id)).size,
    };
    if (row.status === "Former" && !row.active && !row.due && !dec.length) continue;
    coaches.push(row);
  }
  coaches.sort((a, b) => (a.status === "Former") - (b.status === "Former") || b.active - a.active || a.c.localeCompare(b.c));

  /* ---- Lists to act on ---- */
  const who = (c) => (withNames ? { id: c.id, nm: c.name } : { id: c.id });
  const lastPk = (c) => { const l = c.now.last; return l ? (l.unit === "Months" ? `${l.size} mo` : `${l.size} sess`) : null; };
  const expiring = now.filter((c) => c.now.status === "Active" && c.now.fromPackages && c.now.P - today >= 0 && c.now.P - today <= reminder)
    .sort((a, b) => a.now.P - b.now.P)
    .map((c) => ({ ...who(c), coach: c.coach, svc: c.service, pkg: lastPk(c), ends: iso(c.now.P), left: c.now.P - today }));
  const due = now.filter((c) => c.now.status === "Renewal due")
    .sort((a, b) => a.now.P - b.now.P)
    .map((c) => ({ ...who(c), coach: c.coach, svc: c.service, pkg: lastPk(c), ended: iso(c.now.P), over: today - c.now.P, final: iso(c.now.P + grace) }));
  const recent = now.filter((c) => c.now.status === "Ended" && c.endedOn != null && today - c.endedOn <= 60)
    .sort((a, b) => b.endedOn - a.endedOn)
    .map((c) => ({ ...who(c), coach: c.coach, joined: iso(c.joined), ended: iso(c.endedOn), reason: c.reason || null, told: c.now.stop }));

  /* ---- Who the active clients are, and where this year's new ones came from ---- */
  const active = now.filter((c) => c.now.status === "Active");
  const count = (list, f) => {
    const m = new Map();
    for (const x of list) { const k = f(x) || "(not recorded)"; m.set(k, (m.get(k) || 0) + 1); }
    return [...m.entries()].sort((a, b) => (a[0] === "(not recorded)") - (b[0] === "(not recorded)") || b[1] - a[1]);
  };
  const ageOf = (c) => {
    const a = c.dob != null ? Math.floor((today - c.dob) / 365.25) : c.ageOld;
    if (a == null) return null;
    return a < 18 ? "Under 18" : a < 30 ? "18-29" : a < 40 ? "30-39" : a < 50 ? "40-49" : a < 60 ? "50-59" : "60+";
  };
  // The biggest n, then "Other"; "(not recorded)" always keeps its own row, last.
  const top = (pairs, n) => {
    const nr = pairs.filter((p) => p[0] === "(not recorded)"), rest = pairs.filter((p) => p[0] !== "(not recorded)");
    const head = rest.length <= n ? rest : [...rest.slice(0, n), ["Other", rest.slice(n).reduce((a, b) => a + b[1], 0)]];
    return head.concat(nr);
  };
  const firstBuys = packages.filter((p) => p.date >= yearStart && (p.type === "New" || p.type === "Returning") && clients.has(p.id));
  const AGE_ORDER = ["Under 18", "18-29", "30-39", "40-49", "50-59", "60+", "(not recorded)"];
  const mix = {
    service: count(active, (c) => c.service),
    gender: count(active, (c) => ({ F: "Female", M: "Male" })[c.gender] || c.gender),
    age: count(active, ageOf).sort((a, b) => AGE_ORDER.indexOf(a[0]) - AGE_ORDER.indexOf(b[0])),
    nat: top(count(active, (c) => c.nat), 9),
    lang: top(count(active, (c) => c.lang), 6),
    source: count(firstBuys.filter((p) => p.type === "New"), (p) => p.src),
    by: top(count(firstBuys, (p) => p.by), 10),
    sizes: count(packages.filter((p) => p.date >= yearStart && clients.has(p.id)), (p) => (p.unit === "Months" ? `${p.size} mo` : `${p.size}`))
      .sort((a, b) => (parseFloat(a[0]) || 0) - (parseFloat(b[0]) || 0)),
  };

  /* ---- Headline numbers ---- */
  const monthStart = toSerial(+iso(today).slice(0, 4), +iso(today).slice(5, 7), 1);
  const thisMonth = packages.filter((p) => p.date >= monthStart && p.date <= today && clients.has(p.id));
  const ago30 = [...clients.values()].filter((c) => statusOn(c, today - 30).status === "Active").length;
  const allDec = decisions.filter((x) => x.outcome !== "pending" && x.at >= since);
  const kept = allDec.filter((x) => x.outcome === "kept").length, moved = allDec.filter((x) => x.outcome === "moved").length;
  const lostN = allDec.filter((x) => x.outcome === "lost").length;
  const lostTenure = allDec.filter((x) => x.outcome === "lost" && x.joined != null).map((x) => x.E - x.joined).sort((a, b) => a - b);
  const kpi = {
    active: active.length, active30: ago30,
    due: due.length, expiring: expiring.length,
    ended: now.filter((c) => c.now.status === "Ended").length,
    noPkg: now.filter((c) => c.now.status === "No package").length,
    total: clients.size,
    nw: new Set(thisMonth.filter((p) => p.type === "New").map((p) => p.id)).size,
    ret: new Set(thisMonth.filter((p) => p.type === "Returning").map((p) => p.id)).size,
    ren: new Set(thisMonth.filter((p) => p.type === "Renewal").map((p) => p.id)).size,
    kept, moved, lost: lostN, pending: decisions.filter((x) => x.outcome === "pending").length,
    retention: kept + moved + lostN ? (kept + moved) / (kept + moved + lostN) : null,
    medianTenure: lostTenure.length ? lostTenure[Math.floor(lostTenure.length / 2)] : null,
    carry: now.filter((c) => c.now.status === "Active" && !c.now.fromPackages).length,
  };

  /* ---- Things in the file worth a look ---- */
  const checks = [];
  if (unknownIds) checks.push([`${unknownIds} package row(s) use a Client ID that is not on the Clients tab`, "fix"]);
  const noCoach = active.filter((c) => !c.coach).length;
  if (noCoach) checks.push([`${noCoach} active client(s) have no coach on the Clients tab`, "fix"]);
  if (extNoFee) checks.push([`${extNoFee} extension(s) have no fee typed - 500, or 0 if waived`, "fix"]);
  if (offRule.size) checks.push([`Package size(s) ${[...offRule].sort((a, b) => a - b).join(", ")} have no rule of their own on Settings, so they take the next smaller size's length`, "info"]);
  if (kpi.carry) checks.push([`${kpi.carry} active client(s) are covered only by the move from the old lists, until ${iso(goLive + carryDays)}; their next renewal needs logging`, "info"]);
  if (compared) checks.push([`Statuses agree with the file's own for ${agree} of ${compared} clients` + (agree < compared ? " - the rest differ only because the file was last recalculated on an earlier day" : ""), agree === compared ? "ok" : "info"]);
  notes.forEach((n) => checks.push([n, "fix"]));

  return {
    v: 1, today: iso(today), names: withNames,
    settings: { grace, reminder, goLive: iso(goLive), carryUntil: iso(goLive + carryDays), loggedFrom: iso(loggedFrom) },
    kpi, months, coaches, mix, checks,
    lists: { expiring, due, recent },
  };
}
