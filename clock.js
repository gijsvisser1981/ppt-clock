(function () {
  "use strict";

  var KEY = "avpclock";
  var DEFAULTS = {
    mode: "clock",      // "clock" | "countdown"
    target: "",         // "HH:MM" of "HH:MM:SS" voor countdown
    overtime: true,
    h24: true,
    seconds: true,
    date: false,
    font: "Segoe UI",
    weight: "600",
    color: "#ffffff",
    bg: "#000000",
    transparent: false,
    shadow: false,
    shadowColor: "#000000",
    shadowSize: "2"     // 1 = subtiel, 2 = normaal, 3 = sterk
  };

  var cfg = Object.assign({}, DEFAULTS);
  var inOffice = false;
  var timer = null;
  var lastShape = "";

  var $ = function (id) { return document.getElementById(id); };
  var clockEl = $("clock"), timeEl = $("time"), dateEl = $("date");
  var form = $("settings");

  // ---------- opslag: in de presentatie zelf, of localStorage buiten Office ----------
  function load() {
    var s = null;
    if (inOffice) {
      s = Office.context.document.settings.get(KEY);
    } else {
      try { s = JSON.parse(localStorage.getItem(KEY)); } catch (e) { s = null; }
    }
    cfg = Object.assign({}, DEFAULTS, s || {});
  }

  function save() {
    if (inOffice) {
      Office.context.document.settings.set(KEY, cfg);
      Office.context.document.settings.saveAsync();
    } else {
      try { localStorage.setItem(KEY, JSON.stringify(cfg)); } catch (e) { /* ignore */ }
    }
  }

  // ---------- weergave ----------
  function pad(n) { return (n < 10 ? "0" : "") + n; }

  function formatClock(d) {
    var h = d.getHours(), suffix = "";
    if (!cfg.h24) {
      suffix = h < 12 ? " AM" : " PM";
      h = h % 12 || 12;
    }
    var t = (cfg.h24 ? pad(h) : h) + ":" + pad(d.getMinutes());
    if (cfg.seconds) t += ":" + pad(d.getSeconds());
    return t + suffix;
  }

  function targetDate(now) {
    var m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(cfg.target || "");
    if (!m) return null;
    var t = new Date(now);
    t.setHours(+m[1], +m[2], +(m[3] || 0), 0);
    return t;
  }

  function formatCountdown(now) {
    var t = targetDate(now);
    if (!t) return { text: cfg.seconds ? "00:00" : "0:00", over: false };
    var diff = Math.round((t - now) / 1000);
    var over = diff < 0;
    if (over && !cfg.overtime) diff = 0;
    var s = Math.abs(diff);
    var h = Math.floor(s / 3600), mi = Math.floor((s % 3600) / 60), se = s % 60;
    var text;
    if (cfg.seconds) {
      text = (h ? h + ":" + pad(mi) : pad(mi)) + ":" + pad(se);
    } else {
      // zonder seconden afronden naar boven, zodat 0:00 pas op het doelmoment verschijnt
      var totalMin = over ? Math.floor(s / 60) : Math.ceil(s / 60);
      text = Math.floor(totalMin / 60) + ":" + pad(totalMin % 60);
    }
    return { text: (over && diff !== 0 ? "+" : "") + text, over: over && diff !== 0 };
  }

  function formatDate(d) {
    try {
      return d.toLocaleDateString("nl-BE", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
    } catch (e) {
      return d.toDateString();
    }
  }

  function render() {
    var now = new Date();
    var over = false, text;
    if (cfg.mode === "countdown") {
      var r = formatCountdown(now);
      text = r.text;
      over = r.over;
      dateEl.textContent = "";
    } else {
      text = formatClock(now);
      dateEl.textContent = cfg.date ? formatDate(now) : "";
    }
    timeEl.textContent = text;
    clockEl.classList.toggle("over", over);

    // alleen opnieuw schalen als de vorm van de tekst verandert (geen gespring per seconde)
    var shape = text.replace(/\d/g, "0") + "|" + dateEl.textContent.length;
    if (shape !== lastShape) { lastShape = shape; fit(); }
  }

  function fit() {
    clockEl.style.fontSize = "100px";
    var r = clockEl.getBoundingClientRect();
    if (!r.width || !r.height) return;
    var scale = Math.min(window.innerWidth * 0.94 / r.width, window.innerHeight * 0.9 / r.height);
    clockEl.style.fontSize = Math.max(6, Math.floor(100 * scale)) + "px";
  }

  function tick() {
    render();
    // uitlijnen op de volgende seconde
    timer = setTimeout(tick, 1000 - (Date.now() % 1000) + 10);
  }

  function applyStyle() {
    var root = document.documentElement.style;
    root.setProperty("--fg", cfg.color);
    root.setProperty("--bg", cfg.transparent ? "transparent" : cfg.bg);
    // schaduw in em, zodat hij meeschaalt met de cijfers
    var k = +cfg.shadowSize || 2;
    root.setProperty("--shadow", cfg.shadow
      ? (0.02 * k) + "em " + (0.02 * k) + "em " + (0.03 * k) + "em " + cfg.shadowColor
      : "none");
    root.setProperty("--font", '"' + cfg.font + '"');
    root.setProperty("--weight", cfg.weight);
    lastShape = "";
    render();
  }

  // ---------- instellingenformulier ----------
  function fillForm() {
    Object.keys(DEFAULTS).forEach(function (k) {
      var el = form.elements[k];
      if (!el) return;
      if (el.type === "checkbox") el.checked = !!cfg[k];
      else el.value = cfg[k];
    });
    formClass();
  }

  function formClass() {
    form.className = "mode-" + cfg.mode + (cfg.shadow ? "" : " no-shadow");
    form.elements.bg.disabled = !!cfg.transparent;
  }

  function readForm() {
    Object.keys(DEFAULTS).forEach(function (k) {
      var el = form.elements[k];
      if (!el) return;
      cfg[k] = el.type === "checkbox" ? el.checked : el.value;
    });
    formClass();
  }

  form.addEventListener("input", function () { readForm(); applyStyle(); save(); });
  form.addEventListener("change", function () { readForm(); applyStyle(); save(); });
  form.addEventListener("submit", function (e) { e.preventDefault(); form.hidden = true; });
  $("reset").addEventListener("click", function () {
    cfg = Object.assign({}, DEFAULTS);
    fillForm(); applyStyle(); save();
  });
  $("gear").addEventListener("click", function () { fillForm(); form.hidden = false; });

  window.addEventListener("resize", function () { lastShape = ""; render(); });

  // ---------- diavoorstelling-detectie: in "read"-weergave geen instellingen ----------
  function setView(view) {
    document.body.classList.toggle("show", view === "read");
    if (view === "read") form.hidden = true;
  }

  function start() {
    load();
    applyStyle();
    if (timer) clearTimeout(timer);
    tick();
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      inOffice = !!(info && info.host);
      if (inOffice) {
        var doc = Office.context.document;
        if (doc.getActiveViewAsync) {
          doc.getActiveViewAsync(function (r) { if (r.status === "succeeded") setView(r.value); });
          doc.addHandlerAsync(Office.EventType.ActiveViewChanged, function (e) { setView(e.activeView); });
        }
      }
      start();
    });
  } else {
    start();
  }

  // offline-cache: na de eerste keer laden werkt de klok ook zonder internet
  if ("serviceWorker" in navigator && location.protocol === "https:") {
    navigator.serviceWorker.register("sw.js").catch(function () { /* niet kritiek */ });
  }
})();
