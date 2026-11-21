/* The dashboard switcher: one bar of big tabs on its own row at the top of every page, each
   dashboard with its own icon. On a phone it becomes a 4 x 2 grid, so every tab shows without
   scrolling. IAMNAV.draw(active, allowed): active = this page's file, allowed = the files this
   passphrase may open (empty = all). */
(function () {
  var TABS = [
    ["overview.html", "Overview", '<circle cx="12" cy="12" r="8.5"/><circle cx="12" cy="12" r="4.5"/><circle cx="12" cy="12" r=".8"/>'],
    ["director.html", "Fitness", '<polyline points="2.5,12.5 7,12.5 9.5,6 14,18.5 16.5,12.5 21.5,12.5"/>'],
    ["finance.html", "Finance", '<rect x="3" y="6" width="18" height="12" rx="2"/><circle cx="12" cy="12" r="2.6"/><path d="M6.5 9.5v5M17.5 9.5v5"/>'],
    ["report.html", "Team Performance", '<path d="M4 20h16"/><path d="M6.5 20v-7M12 20V5M17.5 20v-10"/>'],
    ["clients.html", "Clients", '<circle cx="9" cy="8.5" r="3.3"/><path d="M3 19.5c0-3.3 2.7-5.8 6-5.8s6 2.5 6 5.8"/><path d="M15.5 5.2a3.3 3.3 0 0 1 0 6.6M21 19.5c0-2.5-1.5-4.6-3.8-5.4"/>'],
    ["business.html", "Business", '<rect x="3" y="7.5" width="18" height="12.5" rx="2"/><path d="M9 7.5V5.5a1.2 1.2 0 0 1 1.2-1.2h3.6A1.2 1.2 0 0 1 15 5.5v2M3 13h18"/>'],
    ["coach.html", "Coach review", '<rect x="5" y="4.5" width="14" height="16.5" rx="2"/><path d="M9.5 3.5h5v3h-5z"/><path d="M9 13.5l2.2 2.2L15.5 11"/>'],
    ["prize.html", "Prize", '<path d="M8 4.5h8v4.5a4 4 0 0 1-8 0z"/><path d="M8 6H5.2a2.6 2.6 0 0 0 2.9 3.6M16 6h2.8a2.6 2.6 0 0 1-2.9 3.6"/><path d="M12 13v3.5M8.5 20h7M9.5 20l.8-3.5h3.4l.8 3.5"/>'],
    ["roster.html", "Head coach", '<path d="M9.5 6.5h11M9.5 12h11M9.5 17.5h11"/><circle cx="5" cy="6.5" r="1.2"/><circle cx="5" cy="12" r="1.2"/><circle cx="5" cy="17.5" r="1.2"/>']
  ];
  var CSS =
    ".dnav{--n:8;display:grid;grid-template-columns:repeat(var(--n),minmax(0,1fr));gap:6px;padding:6px;margin:0 0 26px;" +
    "background:linear-gradient(180deg,#18181b,#0f0f11);border:1px solid rgba(255,255,255,.09);border-radius:16px;" +
    "box-shadow:0 10px 28px rgba(0,0,0,.4);position:sticky;top:10px;z-index:60;}" +
    ".dnav a{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:7px;min-height:66px;padding:11px 6px 10px;" +
    "border-radius:11px;color:#b9b3ab;text-decoration:none;font-size:14.5px;font-weight:500;line-height:1.15;text-align:center;" +
    "letter-spacing:.01em;transition:background .15s,color .15s,transform .15s;-webkit-tap-highlight-color:transparent;}" +
    ".dnav a svg{width:21px;height:21px;fill:none;stroke:currentColor;stroke-width:1.7;stroke-linecap:round;stroke-linejoin:round;opacity:.8;}" +
    ".dnav a:hover{background:rgba(255,255,255,.06);color:#fff;}" +
    ".dnav a.on{background:linear-gradient(180deg,#f47a33,#e05e14);color:#fff;font-weight:600;box-shadow:0 6px 18px rgba(239,105,29,.38),inset 0 1px 0 rgba(255,255,255,.22);}" +
    ".dnav a.on svg{opacity:1;}" +
    ".dnav a:active{transform:scale(.97);}" +
    // Phones: 4 columns (set from script as --m, which every browser, Safari included, reads).
    "@media (max-width:760px){.dnav{grid-template-columns:repeat(var(--m),minmax(0,1fr));position:static;gap:5px;padding:5px;margin:0 0 18px;border-radius:14px;}" +
    ".dnav a{min-height:58px;font-size:12px;padding:8px 3px 7px;gap:5px;border-radius:10px;}.dnav a svg{width:19px;height:19px;}}" +
    "@media (hover:none) and (pointer:coarse){select,input[type=text],input[type=number],input[type=password],input:not([type]){font-size:16px !important;}}" +
    // Touch screens (phone, iPad): selects and text boxes at 16px, so iOS does not zoom in on a tap.
    // Laptops with a short screen: a slimmer bar.
    "@media (min-width:761px) and (max-height:760px){.dnav a{min-height:54px;padding:8px 6px;gap:5px;font-size:13.5px;}.dnav a svg{width:19px;height:19px;}}" +
    // On a touch screen the hover notes are shown on tap.
    ".tapnote{position:absolute;z-index:200;max-width:min(320px,calc(100vw - 24px));background:#1d1d21;color:#eee;border:1px solid rgba(255,255,255,.14);" +
    "border-radius:10px;padding:9px 11px;font-size:12.5px;line-height:1.45;box-shadow:0 10px 30px rgba(0,0,0,.5);pointer-events:none;}";
  function style() {
    if (document.getElementById("dnavCss")) return;
    var s = document.createElement("style"); s.id = "dnavCss"; s.textContent = CSS; document.head.appendChild(s);
  }
  function draw(active, allowed) {
    style();
    var app = document.getElementById("app") || document.body, el = document.getElementById("dnav");
    if (!el) { el = document.createElement("nav"); el.id = "dnav"; el.className = "dnav"; el.setAttribute("aria-label", "Dashboards"); app.insertBefore(el, app.firstChild); }
    var tabs = TABS.filter(function (t) { return !allowed || !allowed.length || allowed.indexOf(t[0]) >= 0 || t[0] === active; });
    el.style.setProperty("--n", tabs.length); el.style.setProperty("--m", tabs.length > 8 ? 5 : Math.min(tabs.length, 4));
    el.innerHTML = tabs.map(function (t) {
      return '<a href="./' + t[0] + '"' + (t[0] === active ? ' class="on" aria-current="page"' : "") + '><svg viewBox="0 0 24 24" aria-hidden="true">' + t[2] + "</svg><span>" + t[1] + "</span></a>";
    }).join("");
  }
  // Phones and tablets have no hover, so the notes kept in title="" never show. A tap on anything
  // carrying one shows it in a small bubble; a tap elsewhere or a scroll closes it.
  function tapNotes() {
    var touch = false; try { touch = window.matchMedia("(hover: none)").matches; } catch (e) {}
    if (!touch) return;
    var tip = null;
    function close() { if (tip) { tip.remove(); tip = null; } }
    document.addEventListener("click", function (e) {
      close();
      var t = e.target && e.target.closest && e.target.closest("[title],[data-tn]"); if (!t) return;
      if (t.title) { t.setAttribute("data-tn", t.title); t.removeAttribute("title"); }   // no second, native bubble
      var txt = t.getAttribute("data-tn"); if (!txt) return;
      tip = document.createElement("div"); tip.className = "tapnote"; tip.textContent = txt; document.body.appendChild(tip);
      var r = t.getBoundingClientRect(), w = tip.offsetWidth, x = Math.max(12, Math.min(window.innerWidth - w - 12, r.left + r.width / 2 - w / 2));
      var y = r.bottom + 8; if (y + tip.offsetHeight > window.innerHeight - 8) y = r.top - tip.offsetHeight - 8;
      tip.style.left = (x + window.scrollX) + "px"; tip.style.top = (Math.max(8, y) + window.scrollY) + "px";
    }, true);
    window.addEventListener("scroll", close, { passive: true });
  }
  style(); tapNotes();
  window.IAMNAV = { draw: draw, TABS: TABS };
})();
