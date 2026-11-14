/* Shared by the Overview and Coach review pages: unlocking, loading, formatting, tooltips and
   the small chart pieces. Both pages open with the Performance passphrase (the same one the
   Performance page remembers), and read live.enc.json and overview.enc.json. */
(function (W) {
"use strict";
var GH_OWNER = "hossamyossef701-coder", GH_REPO = "iam-performance", GH_BRANCH = "main";
var PASS_KEY = "iamlive:pass", CFG_KEY = "iamlive:cfg", NOTES_KEY = "iam:notes";
var MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
var C = W.IAM = {};
C.MONTHS = MONTHS;
C.$ = function (id) { return document.getElementById(id); };
C.esc = function (t) { return String(t == null ? "" : t).replace(/[&<>"']/g, function (ch) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]; }); };
C.full = function (v) { return v == null ? "—" : Math.round(v).toLocaleString("en-US"); };
C.aed = function (v) { if (v == null) return "—"; var a = Math.abs(v); return a >= 999500 ? (v / 1e6).toFixed(2) + "M" : a >= 9950 ? Math.round(v / 1000) + "k" : a >= 1000 ? (v / 1000).toFixed(1) + "k" : String(Math.round(v)); };
C.pct = function (a, b) { return b ? Math.round(a / b * 100) + "%" : "—"; };
C.label = function (p, short) { var m = +p.slice(5, 7); return short ? MONTHS[m - 1] : MONTHS[m - 1] + " " + p.slice(0, 4); };
C.day = function (iso) { return iso ? +iso.slice(8, 10) + " " + MONTHS[+iso.slice(5, 7) - 1] : "—"; };
C.date = function (iso) { return iso ? +iso.slice(8, 10) + " " + MONTHS[+iso.slice(5, 7) - 1] + " " + iso.slice(0, 4) : "—"; };
C.median = function (a) { a = a.filter(function (x) { return x != null && isFinite(x); }).sort(function (x, y) { return x - y; }); if (!a.length) return null; var m = Math.floor(a.length / 2); return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2; };
C.daysIn = function (p) { return new Date(Date.UTC(+p.slice(0, 4), +p.slice(5, 7), 0)).getUTCDate(); };
C.prevP = function (p, n) { var y = +p.slice(0, 4), m = +p.slice(5, 7) - n; while (m < 1) { m += 12; y--; } return y + "-" + (m < 10 ? "0" : "") + m; };
C.addDays = function (iso, n) { return new Date(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) + n * 864e5).toISOString().slice(0, 10); };
C.wOf = function (ch) { return ch === "1" || ch === "x" ? 1 : ch === "h" ? 0.5 : (ch >= "2" && ch <= "9") ? +ch : 0; };
C.nameKey = function (n) { return String(n).toLowerCase().replace(/\([^)]*\)/g, " ").replace(/[^a-z0-9؀-ۿ]+/g, " ").trim(); };

/* ---- decryption: identical parameters to scripts/crypto.mjs ---- */
function b2b(s) { var b = atob(s), u = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; }
C.decrypt = function (env, pass) {
  return crypto.subtle.importKey("raw", new TextEncoder().encode(String(pass).trim()), "PBKDF2", false, ["deriveKey"])
    .then(function (base) { return crypto.subtle.deriveKey({ name: "PBKDF2", salt: b2b(env.salt), iterations: env.it || 250000, hash: "SHA-256" }, base, { name: "AES-GCM", length: 256 }, false, ["decrypt"]); })
    .then(function (k) { return crypto.subtle.decrypt({ name: "AES-GCM", iv: b2b(env.iv) }, k, b2b(env.ct)); })
    .then(function (buf) { return JSON.parse(new TextDecoder().decode(buf)); });
};
function getJson(url, headers) { return fetch(url, { cache: "no-store", headers: headers || {} }).then(function (r) { if (!r.ok) throw new Error("HTTP " + r.status); return r.json(); }); }
C.getJson = getJson;
function cfg() { try { return JSON.parse(localStorage.getItem(CFG_KEY) || "{}"); } catch (e) { return {}; } }
// Straight from the repository when possible (the Pages copy can lag a few minutes).
C.loadEnc = function (name) {
  var c = cfg(), h = { Accept: "application/vnd.github.raw" };
  if (c && c.pat) h.Authorization = "Bearer " + c.pat;
  return getJson("https://api.github.com/repos/" + GH_OWNER + "/" + GH_REPO + "/contents/" + name + "?ref=" + GH_BRANCH + "&t=" + Date.now(), h)
    .catch(function () { return getJson("./" + name + "?t=" + Date.now()); });
};
C.load = function (pass) {
  return Promise.all([C.loadEnc("live.enc.json").then(function (e) { return C.decrypt(e, pass); }),
    C.loadEnc("overview.enc.json").then(function (e) { return C.decrypt(e, pass); }, function (e) { if (/HTTP 404/.test(e && e.message || "")) return null; throw e; })])
    .then(function (r) { return { live: r[0], ov: r[1] }; });
};
// The gate: the same passphrase as the Performance page, remembered under the same key.
C.unlock = function (onOpen) {
  var gate = C.$("gate"), form = C.$("gateForm"), input = C.$("gatePass"), rem = C.$("gateRemember"), btn = C.$("gateBtn");
  function status(t, c) { var e = C.$("gateStatus"); e.className = "status " + (c || ""); e.textContent = t; }
  function attempt(pass, keep, fromStore) {
    status("Decrypting…"); btn.disabled = true;
    C.load(pass).then(function (d) {
      if (keep) { try { localStorage.setItem(PASS_KEY, pass); } catch (e) {} }
      gate.style.display = "none"; document.body.classList.remove("locked"); C.$("app").style.display = "";
      onOpen(d, pass);
    }).catch(function (e) {
      btn.disabled = false;
      if (fromStore) { try { localStorage.removeItem(PASS_KEY); } catch (e2) {} }
      var wrong = e && (e.name === "OperationError" || /decrypt/i.test(e.message || ""));
      status(wrong ? "That passphrase doesn’t open the data." : "Could not load the data: " + (e && e.message), "err");
    });
  }
  form.addEventListener("submit", function (ev) { ev.preventDefault(); var v = (input.value || "").trim(); if (!v) { status("Enter the passphrase.", "err"); return; } attempt(v, !!rem.checked, false); });
  var saved = null; try { saved = localStorage.getItem(PASS_KEY); } catch (e) {}
  if (saved) attempt(saved, true, true); else input.focus();
};
// Notes under the panels are hidden unless asked for: one switch for the whole page.
C.notes = function (btn) {
  var on = false; try { on = localStorage.getItem(NOTES_KEY) === "1"; } catch (e) {}
  function set() { document.body.classList.toggle("notes", on); btn.classList.toggle("on", on); }
  btn.addEventListener("click", function () { on = !on; try { localStorage.setItem(NOTES_KEY, on ? "1" : "0"); } catch (e) {} set(); });
  set();
};

/* ---- tooltip and chart pieces ---- */
var tip = null;
C.showTip = function (host, html, x, y) {
  if (!tip) { tip = document.createElement("div"); tip.className = "tip"; document.body.appendChild(tip); }
  tip.innerHTML = html; tip.style.opacity = "1";
  var r = host.getBoundingClientRect(), tw = tip.offsetWidth, th = tip.offsetHeight;
  var left = r.left + window.scrollX + x - tw / 2, top = r.top + window.scrollY + y - th - 12;
  left = Math.max(8, Math.min(left, window.scrollX + document.documentElement.clientWidth - tw - 8));
  if (top < window.scrollY + 8) top = r.top + window.scrollY + y + 16;
  tip.style.left = left + "px"; tip.style.top = top + "px";
};
C.hideTip = function () { if (tip) tip.style.opacity = "0"; };
document.addEventListener("scroll", function () { C.hideTip(); }, { passive: true });
C.trow = function (l, v, col) { return '<div class="row"><span>' + (col ? '<i style="background:' + col + '"></i>' : "") + l + "</span><span>" + v + "</span></div>"; };
C.hits = function (host, areas) {
  var svg = host.querySelector("svg"), NS = "http://www.w3.org/2000/svg";
  areas.forEach(function (a) {
    var r = document.createElementNS(NS, "rect");
    r.setAttribute("x", a.x); r.setAttribute("y", a.y); r.setAttribute("width", a.w); r.setAttribute("height", a.h); r.setAttribute("fill", "transparent");
    function on() { C.showTip(host, a.html, a.ax, a.ay); var hl = svg.querySelector(".hl"); if (hl) { hl.setAttribute("x", a.x); hl.setAttribute("y", a.y); hl.setAttribute("width", a.w); hl.setAttribute("height", a.h); hl.style.opacity = "1"; } }
    r.addEventListener("mouseenter", on); r.addEventListener("mousemove", on); r.addEventListener("touchstart", on, { passive: true });
    r.addEventListener("mouseleave", function () { C.hideTip(); var hl = svg.querySelector(".hl"); if (hl) hl.style.opacity = "0"; });
    svg.appendChild(r);
  });
};
C.niceMax = function (v) { if (v <= 5) return 5; var p = Math.pow(10, Math.floor(Math.log10(v))), n = v / p; return (n <= 1.2 ? 1.2 : n <= 1.5 ? 1.5 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 3 ? 3 : n <= 4 ? 4 : n <= 5 ? 5 : n <= 6 ? 6 : n <= 8 ? 8 : 10) * p; };
C.width = function (host) { return Math.max(320, Math.round(host.clientWidth || 600)); };
C.bar = function (x, y, w, h, fill, down) {
  if (h <= 0.4) return "";
  var r = Math.min(4, h / 2, w / 2);
  if (!down) return '<path d="M' + x + "," + (y + h) + "V" + (y + r) + "Q" + x + "," + y + " " + (x + r) + "," + y + "H" + (x + w - r) + "Q" + (x + w) + "," + y + " " + (x + w) + "," + (y + r) + "V" + (y + h) + 'Z" fill="' + fill + '"/>';
  return '<path d="M' + x + "," + y + "V" + (y + h - r) + "Q" + x + "," + (y + h) + " " + (x + r) + "," + (y + h) + "H" + (x + w - r) + "Q" + (x + w) + "," + (y + h) + " " + (x + w) + "," + (y + h - r) + "V" + y + 'Z" fill="' + fill + '"/>';
};
C.hbar = function (x, y, w, h, fill) {
  if (w <= 0.4) return ""; var r = Math.min(4, w / 2, h / 2);
  return '<path d="M' + x + "," + y + "H" + (x + w - r) + "Q" + (x + w) + "," + y + " " + (x + w) + "," + (y + r) + "V" + (y + h - r) + "Q" + (x + w) + "," + (y + h) + " " + (x + w - r) + "," + (y + h) + "H" + x + 'Z" fill="' + fill + '"/>';
};
C.txt = function (x, y, t, o) { o = o || {}; return '<text x="' + x + '" y="' + y + '" text-anchor="' + (o.a || "middle") + '" style="font-size:' + (o.s || 11) + "px;fill:" + (o.c || "var(--ink-soft)") + '">' + t + "</text>"; };
C.grid = function (W, L, R, y, ticks, fmt) {
  return '<g class="grid">' + ticks.filter(function (t) { return t; }).map(function (t) { return '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(t) + '" y2="' + y(t) + '"/>'; }).join("") + "</g>" +
    '<g class="axis">' + ticks.map(function (t) { return '<text x="' + (L - 8) + '" y="' + (y(t) + 3) + '" text-anchor="end">' + fmt(t) + "</text>"; }).join("") + "</g>";
};

/* ---- the coach sheets, client by client ---- */
// The last day that has happened: the job's day for the month running, else the month's end.
C.lastDay = function (live, p) {
  if (p !== live.cutoffPeriod) return C.daysIn(p);
  return live.cutoffDay || C.daysIn(p);
};
C.clientsIn = function (live, p, c) {
  var v = live.periods[p] && live.periods[p][c], rows = v && v.cl; if (!rows) return null;
  var upto = C.lastDay(live, p);
  return rows.map(function (r) { var s = 0, m = String(r[2] || ""); for (var i = 0; i < Math.min(upto, m.length); i++) s += C.wOf(m[i]); return { name: r[0], key: C.nameKey(r[0]), price: r[1], s: s, m: m }; });
};
// Clients training at half their own usual or less: the 3 weeks to the last day against the 8 before.
C.lessOften = function (live, p, c) {
  var upto = C.lastDay(live, p), anchor = p + "-" + (upto < 10 ? "0" : "") + upto, from = C.addDays(anchor, -76), cut = C.addDays(anchor, -20);
  var cur = C.clientsIn(live, p, c); if (!cur) return null;
  var on = {}; cur.forEach(function (x) { on[x.key] = x.name; });
  var days = {};
  [C.prevP(p, 3), C.prevP(p, 2), C.prevP(p, 1), p].forEach(function (q) {
    var v = live.periods[q] && live.periods[q][c]; if (!v || !v.cl) return;
    var last = q === p ? upto : C.daysIn(q);
    v.cl.forEach(function (r) {
      var k = C.nameKey(r[0]); if (!(k in on)) return;
      var m = String(r[2] || ""), d = days[k] || (days[k] = {});
      for (var i = 0; i < Math.min(last, m.length); i++) { var w = C.wOf(m[i]); if (w) { var iso = q + "-" + (i < 9 ? "0" : "") + (i + 1); if (iso >= from && iso <= anchor) d[iso] = (d[iso] || 0) + w; } }
    });
  });
  var out = [];
  Object.keys(on).forEach(function (k) {
    var d = days[k] || {}, ds = Object.keys(d).sort(); if (!ds.length) return;
    var recent = 0, base = 0; ds.forEach(function (iso) { if (iso >= cut) recent += d[iso]; else base += d[iso]; });
    var usual = base / 8, now = recent / 3;
    if (ds[0] > C.addDays(anchor, -62) || usual < 1.5) return;
    if (now <= usual * 0.5) out.push({ name: on[k], usual: usual, now: now });
  });
  return out;
};
})(window);
