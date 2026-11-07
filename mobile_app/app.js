/* H2R Nigeria results viewer.
 * Plain JavaScript, no build step. Data comes from data/h2r_data.json, which
 * scripts/build_data.py generates from the R pipeline's Excel outputs. */
(function () {
  "use strict";

  var DATA_URL = "data/h2r_data.json";
  var MAX_OPTIONS = 6;
  var app = document.getElementById("app");
  var D = null;
  var expanded = {};

  // ---------- small helpers ----------
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function pct(v) {
    if (v == null) return "–";
    var p = v * 100;
    return (p > 0 && p < 1 ? "<1" : Math.round(p)) + "%";
  }
  function store(key, val) {
    try {
      if (val === undefined) return localStorage.getItem("h2r." + key);
      localStorage.setItem("h2r." + key, val);
    } catch (e) { return null; }
  }
  function byId(list, id) {
    for (var i = 0; i < list.length; i++) if (list[i].id === id) return list[i];
    return null;
  }
  function sectorLabel(id) {
    var s = byId(D.sectors, id);
    return s ? s.label : (id || "General");
  }
  function normSector(s) {
    s = (s || "").toLowerCase();
    return s === "unmas" ? "unmas" : s;
  }
  function highlight(text, q) {
    var safe = esc(text);
    if (!q || typeof q !== "string") return safe;
    var re = new RegExp("(" + q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")", "ig");
    return safe.replace(re, "<mark>$1</mark>");
  }
  function matches(text, q) {
    return !q || String(text).toLowerCase().indexOf(q.toLowerCase()) !== -1;
  }

  // ---------- routing ----------
  function parseHash() {
    var h = location.hash.replace(/^#\/?/, "");
    var parts = h.split("?");
    var params = {};
    (parts[1] || "").split("&").forEach(function (kv) {
      if (!kv) return;
      var p = kv.split("=");
      params[decodeURIComponent(p[0])] = decodeURIComponent(p[1] || "");
    });
    return { view: parts[0] || "overview", params: params };
  }
  function go(view, params, replace) {
    var qs = Object.keys(params || {}).filter(function (k) { return params[k]; })
      .map(function (k) { return encodeURIComponent(k) + "=" + encodeURIComponent(params[k]); }).join("&");
    var url = "#/" + view + (qs ? "?" + qs : "");
    if (replace) { history.replaceState(null, "", url); render(); }
    else location.hash = url;
  }

  function render() {
    if (!D) return;
    var r = parseHash();
    var views = { overview: viewOverview, explore: viewExplore, compare: viewCompare, findings: viewFindings, about: viewAbout };
    var fn = views[r.view] || viewOverview;
    var tab = views[r.view] ? r.view : "overview";
    document.querySelectorAll(".tabbar a").forEach(function (a) {
      if (a.getAttribute("data-tab") === tab) a.setAttribute("aria-current", "page");
      else a.removeAttribute("aria-current");
    });
    var keepScroll = render.lastView === tab;
    var y = window.scrollY;
    app.innerHTML = fn(r.params);
    bind(tab, r.params);
    if (keepScroll) window.scrollTo(0, y); else window.scrollTo(0, 0);
    render.lastView = tab;
  }

  // ---------- shared fragments ----------
  function chips(name, items, current) {
    return '<div class="chips" role="group" aria-label="' + esc(name) + '">' + items.map(function (it) {
      return '<button type="button" class="chip" data-' + esc(name) + '="' + esc(it.id) + '" aria-pressed="' +
        (it.id === current) + '">' + esc(it.label) + "</button>";
    }).join("") + "</div>";
  }
  function sectorItems(list) {
    return [{ id: "", label: "All sectors" }].concat(list);
  }
  function questionHead(q) {
    return '<div class="q-head"><div>' +
      (q.subtitle ? '<p class="q-sub">' + esc(q.subtitle) + "</p>" : "") +
      '<p class="q-text">' + esc(q.question) + "</p></div>" +
      '<span class="badge' + (q.type === "rank" ? " rank" : "") + '">' + (q.type === "rank" ? "Top 3" : esc(sectorLabel(q.sector))) + "</span></div>";
  }
  function filterQuestions(sector, search) {
    return D.questions.filter(function (q) {
      if (sector && q.sector !== sector) return false;
      if (!search) return true;
      return matches(q.question, search) || q.options.some(function (o) { return matches(o.label, search); });
    });
  }
  function usedSectors(list) {
    var ids = {};
    list.forEach(function (x) { ids[normSector(x.sector)] = true; });
    return D.sectors.filter(function (s) { return ids[s.id]; });
  }

  // ---------- Overview ----------
  function viewOverview() {
    var total = D.lgas.reduce(function (s, l) { return s + (l.settlements || 0); }, 0);
    var max = Math.max.apply(null, D.lgas.map(function (l) { return l.settlements || 0; }));
    var qCount = {};
    D.questions.forEach(function (q) { qCount[q.sector] = (qCount[q.sector] || 0) + 1; });
    var m = D.meta;

    return '<h1>' + esc(m.title) + "</h1>" +
      '<p class="lede">' + esc(m.state) + " State, " + esc(m.country) + " · " + esc(m.round) +
      ". Share of assessed settlements reporting each situation, by Local Government Area.</p>" +
      '<div class="stats">' +
      stat(total.toLocaleString(), "settlements") +
      stat(D.lgas.length, "LGAs") +
      stat(D.questions.length, "indicators") +
      "</div>" +
      "<h2>Settlements assessed per LGA</h2>" +
      '<div class="lga-list">' + D.lgas.map(function (l) {
        var w = max ? Math.round((l.settlements || 0) / max * 100) : 0;
        return '<a class="card lga-row" href="#/explore?lga=' + esc(l.id) + '">' +
          '<span class="lga-name">' + esc(l.label) + "</span>" +
          '<span class="lga-count">' + (l.settlements != null ? l.settlements + " settlements" : "–") + "</span>" +
          '<span class="lga-track"><span class="lga-fill" style="width:' + w + '%"></span></span></a>';
      }).join("") + "</div>" +
      "<h2>Browse by sector</h2>" +
      '<div class="sector-grid">' + D.sectors.filter(function (s) { return qCount[s.id]; }).map(function (s) {
        return '<a class="card sector-tile" href="#/explore?sector=' + esc(s.id) + '"><strong>' + esc(s.label) +
          "</strong><span>" + qCount[s.id] + " indicator" + (qCount[s.id] === 1 ? "" : "s") + "</span></a>";
      }).join("") + "</div>" +
      "<h2>Key findings</h2>" +
      '<div class="stack">' + D.findings.slice(0, 3).map(function (f) { return findingCard(f, ""); }).join("") + "</div>" +
      '<div class="btn-row"><a class="btn secondary" href="#/findings">See all ' + D.findings.length + " findings</a></div>";
  }
  function stat(v, label) {
    return '<div class="card stat"><div class="stat-value">' + esc(v) + '</div><div class="stat-label">' + esc(label) + "</div></div>";
  }

  // ---------- Indicators (one LGA) ----------
  function viewExplore(p) {
    var lga = byId(D.lgas, p.lga) ? p.lga : (byId(D.lgas, store("lga")) ? store("lga") : D.lgas[0].id);
    store("lga", lga);
    var L = byId(D.lgas, lga);
    var list = filterQuestions(p.sector, p.q);
    return '<div class="filters">' +
      chips("lga", D.lgas, lga) +
      chips("sector", sectorItems(usedSectors(D.questions)), p.sector || "") +
      '<input class="search" type="search" id="search" placeholder="Search indicators" value="' + esc(p.q || "") + '" aria-label="Search indicators">' +
      "</div>" +
      '<p class="result-count">' + list.length + " indicator" + (list.length === 1 ? "" : "s") + " · " + esc(L.label) +
      (L.settlements ? " · " + L.settlements + " settlements" : "") + "</p>" +
      (list.length ? '<div class="stack">' + list.map(function (q) { return questionCard(q, L); }).join("") + "</div>"
        : '<p class="empty">No indicators match.</p>');
  }

  function questionCard(q, L) {
    var opts = q.options.map(function (o) { return { label: o.label, v: o.values[L.id] }; });
    var hasData = opts.some(function (o) { return o.v != null; });
    if (q.type === "rank") opts.sort(function (a, b) { return (b.v || 0) - (a.v || 0); });
    var top = {};
    if (q.type === "rank") opts.forEach(function (o, i) { if (i < 3 && o.v) top[i] = true; });
    var open = expanded[q.id];
    var shown = open ? opts : opts.slice(0, MAX_OPTIONS);
    var body = hasData ? '<ul class="bars">' + shown.map(function (o, i) {
      var dim = q.type === "rank" && !top[i];
      return '<li class="bar-row' + (dim ? " dim" : "") + '">' +
        '<span class="bar-label">' + (top[i] ? '<span class="top-mark">' + (i + 1) + "</span>" : "") + esc(o.label) + "</span>" +
        '<span class="bar-value">' + pct(o.v) + "</span>" +
        '<span class="bar-track"><span class="bar-fill" style="width:' + Math.round((o.v || 0) * 100) + '%"></span></span></li>';
    }).join("") + "</ul>" +
      (opts.length > MAX_OPTIONS ? '<button class="more-btn" type="button" data-expand="' + esc(q.id) + '">' +
        (open ? "Show fewer" : "Show all " + opts.length + " options") + "</button>" : "")
      : '<p class="q-meta">No data reported for ' + esc(L.label) + ".</p>";
    return '<article class="card">' + questionHead(q) + body +
      '<p class="q-meta">' + esc(sectorLabel(q.sector)) + " · % of settlements in " + esc(L.label) + "</p></article>";
  }

  // ---------- Compare (all LGAs) ----------
  function viewCompare(p) {
    var list = filterQuestions(p.sector, p.q);
    return '<div class="filters">' +
      chips("sector", sectorItems(usedSectors(D.questions)), p.sector || "") +
      '<input class="search" type="search" id="search" placeholder="Search indicators" value="' + esc(p.q || "") + '" aria-label="Search indicators">' +
      "</div>" +
      '<p class="result-count">' + list.length + " indicator" + (list.length === 1 ? "" : "s") + " across " + D.lgas.length + " LGAs</p>" +
      (list.length ? '<div class="stack">' + list.map(compareCard).join("") + "</div>" : '<p class="empty">No indicators match.</p>') +
      '<div class="legend"><span>0%</span><span class="legend-ramp"></span><span>100%</span></div>';
  }

  function compareCard(q) {
    var best = {};
    if (q.type === "rank") {
      D.lgas.forEach(function (l) {
        var m = -1, idx = -1;
        q.options.forEach(function (o, i) { var v = o.values[l.id]; if (v != null && v > m) { m = v; idx = i; } });
        if (m > 0) best[l.id] = idx;
      });
    }
    var rows = q.options.map(function (o, i) {
      return "<tr><td class=\"opt\">" + esc(o.label) + "</td>" + D.lgas.map(function (l) {
        var v = o.values[l.id];
        if (v == null) return '<td class="cell na">–</td>';
        var a = Math.max(0.06, v);
        var dark = v > 0.55;
        var style = "background:rgba(var(--heat)," + a.toFixed(2) + ");" + (dark ? "color:var(--on-accent);" : "") +
          (best[l.id] === i ? "font-weight:700;box-shadow:inset 0 0 0 2px var(--text);" : "");
        return '<td class="cell" style="' + style + '">' + pct(v) + "</td>";
      }).join("") + "</tr>";
    }).join("");
    return '<article class="card">' + questionHead(q) +
      '<div class="heat-wrap"><table class="heat"><colgroup><col class="opt-col"></colgroup><thead><tr><th class="opt" scope="col">Response</th>' +
      D.lgas.map(function (l) { return '<th scope="col">' + esc(l.label) + "</th>"; }).join("") +
      "</tr></thead><tbody>" + rows + "</tbody></table></div>" +
      (q.type === "rank" ? '<p class="q-meta">Outlined cell = most reported response in that LGA.</p>' : "") +
      "</article>";
  }

  // ---------- Findings ----------
  function findingCard(f, q) {
    return '<article class="card finding"><p>' + highlight(f.text, q) + '</p><span class="tag">' +
      esc(sectorLabel(normSector(f.sector))) + "</span></article>";
  }
  function viewFindings(p) {
    var list = D.findings.filter(function (f) {
      return (!p.sector || normSector(f.sector) === p.sector) && matches(f.text, p.q);
    });
    return '<div class="filters">' +
      chips("sector", sectorItems(usedSectors(D.findings)), p.sector || "") +
      '<input class="search" type="search" id="search" placeholder="Search findings, e.g. Gwoza or malaria" value="' + esc(p.q || "") + '" aria-label="Search findings">' +
      "</div>" +
      '<p class="result-count">' + list.length + " finding" + (list.length === 1 ? "" : "s") + "</p>" +
      (list.length ? '<div class="stack">' + list.map(function (f) { return findingCard(f, p.q); }).join("") + "</div>"
        : '<p class="empty">No findings match.</p>');
  }

  // ---------- About ----------
  function viewAbout() {
    var m = D.meta;
    var standalone = window.matchMedia && window.matchMedia("(display-mode: standalone)").matches;
    return "<h1>About this data</h1>" +
      '<div class="card prose">' + m.methodology.split(/\n+/).map(function (para) {
        return "<p>" + esc(para) + "</p>";
      }).join("") + "</div>" +
      "<h2>Dataset</h2>" +
      '<div class="card"><dl class="kv">' +
      "<dt>Round</dt><dd>" + esc(m.round) + "</dd>" +
      "<dt>Area</dt><dd>" + esc(m.state) + " State, " + esc(m.country) + "</dd>" +
      "<dt>LGAs</dt><dd>" + D.lgas.map(function (l) { return esc(l.label); }).join(", ") + "</dd>" +
      "<dt>Source file</dt><dd>" + esc(m.source) + "</dd>" +
      "<dt>Data built</dt><dd>" + esc(m.generated) + "</dd>" +
      "</dl></div>" +
      "<h2>Reading the charts</h2>" +
      '<div class="card prose"><ul>' +
      "<li>Every value is the <strong>share of assessed settlements</strong> in an LGA that reported the response.</li>" +
      "<li><strong>Top 3</strong> questions are single-choice. Their top three responses are highlighted.</li>" +
      "<li>Small LGAs (for example Ngala, with " + ((byId(D.lgas, "ngala") || {}).settlements || "few") +
      " settlements) can swing widely, so read their percentages with care.</li>" +
      "</ul></div>" +
      "<h2>Use it offline</h2>" +
      '<div class="card prose">' +
      (standalone ? "<p>The app is installed on this device and works without a connection.</p>"
        : "<p>Add the app to your home screen to open it like a native app. After the first visit it works without a connection.</p>" +
          "<ul><li><strong>Android (Chrome):</strong> menu ⋮ → <em>Install app</em> or <em>Add to Home screen</em>.</li>" +
          "<li><strong>iPhone (Safari):</strong> Share → <em>Add to Home Screen</em>.</li></ul>") +
      '<div class="btn-row">' +
      '<button class="btn" type="button" id="install-btn"' + (deferredInstall ? "" : " hidden") + ">Install app</button>" +
      '<button class="btn secondary" type="button" id="refresh-btn">Check for updated data</button>' +
      "</div></div>";
  }

  // ---------- events ----------
  function bind(view, p) {
    app.querySelectorAll("[data-lga]").forEach(function (b) {
      b.addEventListener("click", function () { go(view, { lga: b.getAttribute("data-lga"), sector: p.sector, q: p.q }, true); });
    });
    app.querySelectorAll("[data-sector]").forEach(function (b) {
      b.addEventListener("click", function () { go(view, { lga: p.lga, sector: b.getAttribute("data-sector"), q: p.q }, true); });
    });
    app.querySelectorAll("[data-expand]").forEach(function (b) {
      b.addEventListener("click", function () { var id = b.getAttribute("data-expand"); expanded[id] = !expanded[id]; render(); });
    });
    var s = document.getElementById("search");
    if (s) {
      var t;
      s.addEventListener("input", function () {
        clearTimeout(t);
        t = setTimeout(function () {
          go(view, { lga: p.lga, sector: p.sector, q: s.value.trim() }, true);
          var again = document.getElementById("search");
          if (again) { again.focus(); again.setSelectionRange(again.value.length, again.value.length); }
        }, 200);
      });
    }
    // Scroll the selected chip into view.
    app.querySelectorAll('.chip[aria-pressed="true"]').forEach(function (c) {
      var row = c.parentNode;
      if (c.offsetLeft + c.offsetWidth > row.clientWidth) row.scrollLeft = c.offsetLeft - 16;
    });
    var ib = document.getElementById("install-btn");
    if (ib) ib.addEventListener("click", function () {
      if (!deferredInstall) return;
      deferredInstall.prompt();
      deferredInstall = null;
      ib.hidden = true;
    });
    var rb = document.getElementById("refresh-btn");
    if (rb) rb.addEventListener("click", function () {
      loadData(true).then(function () { toast("Data is up to date"); render(); })
        .catch(function () { toast("You're offline. Showing saved data."); });
    });
  }

  function toast(msg, actionLabel, action) {
    var el = document.getElementById("toast");
    el.innerHTML = esc(msg) + (actionLabel ? ' <button type="button">' + esc(actionLabel) + "</button>" : "");
    if (actionLabel) el.querySelector("button").addEventListener("click", action);
    el.classList.add("show");
    clearTimeout(toast.t);
    toast.t = setTimeout(function () { el.classList.remove("show"); }, actionLabel ? 10000 : 2800);
  }

  // ---------- theme ----------
  function applyTheme(t) {
    if (t === "light" || t === "dark") document.documentElement.setAttribute("data-theme", t);
    else document.documentElement.removeAttribute("data-theme");
  }
  applyTheme(store("theme"));
  document.getElementById("theme-btn").addEventListener("click", function () {
    var cur = document.documentElement.getAttribute("data-theme") ||
      (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
    var next = cur === "dark" ? "light" : "dark";
    applyTheme(next);
    store("theme", next);
  });

  // ---------- install + offline ----------
  var deferredInstall = null;
  window.addEventListener("beforeinstallprompt", function (e) {
    e.preventDefault();
    deferredInstall = e;
    var ib = document.getElementById("install-btn");
    if (ib) ib.hidden = false;
  });

  if ("serviceWorker" in navigator && (location.protocol === "https:" || location.hostname === "localhost" || location.hostname === "127.0.0.1")) {
    window.addEventListener("load", function () {
      navigator.serviceWorker.register("sw.js").then(function (reg) {
        reg.addEventListener("updatefound", function () {
          var nw = reg.installing;
          if (!nw) return;
          nw.addEventListener("statechange", function () {
            if (nw.state === "installed" && navigator.serviceWorker.controller) {
              toast("A new version is available.", "Reload", function () { location.reload(); });
            }
          });
        });
      }).catch(function () { /* offline support unavailable here */ });
    });
  }

  // ---------- boot ----------
  function loadData(fresh) {
    return fetch(DATA_URL, fresh ? { cache: "reload" } : undefined)
      .then(function (r) { if (!r.ok) throw new Error(r.status); return r.json(); })
      .then(function (d) {
        D = d;
        document.getElementById("round-label").textContent = d.meta.state + " · " + d.meta.round;
      });
  }

  window.addEventListener("hashchange", render);
  loadData().then(render).catch(function (err) {
    if (D) { console.error(err); app.innerHTML = '<p class="empty">Something went wrong showing this page.</p>'; return; }
    app.innerHTML = '<p class="empty">Could not load the data. Check your connection and try again.</p>';
  });
})();
