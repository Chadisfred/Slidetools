/* Slide Tools core: shared state, saved settings, the formula evaluator,
   grid maths, selection tracking and PowerPoint helpers.
   Every tool file talks to the panel through window.ST. */
(function () {
  "use strict";

  var PT = 72;
  var SETTINGS_KEY = "slideToolsSettings";

  var ST = window.ST = {
    PT: PT,
    MAX_COLS: 12,
    MAX_ROWS: 8,
    FRACTIONS: [0.125, 0.25, 0.5, 0.75, 1],
    FRACTION_LABELS: ["1/8", "1/4", "1/2", "3/4", "1"],
    state: {},
    defaults: {},
    grid: {},
    selection: { count: 0, w: 0, h: 0, items: [] },
    inOffice: false,
    tools: []
  };

  ST.$ = function (id) { return document.getElementById(id); };

  // ---------- Events ----------
  var handlers = {};
  ST.on = function (name, fn) { (handlers[name] = handlers[name] || []).push(fn); };
  ST.emit = function (name, data) {
    (handlers[name] || []).forEach(function (fn) {
      try { fn(data); } catch (e) { if (window.console) console.error(e); }
    });
  };

  // ---------- Saved settings ----------
  // Tool values live in one flat object so older saved settings keep working.
  ST.addDefaults = function (values) {
    Object.keys(values).forEach(function (k) { ST.defaults[k] = values[k]; });
  };
  ST.loadState = function () {
    var saved = null;
    try { saved = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "null"); } catch (e) { saved = null; }
    var merged = {};
    Object.keys(ST.defaults).forEach(function (k) { merged[k] = clone(ST.defaults[k]); });
    if (saved) Object.keys(saved).forEach(function (k) { merged[k] = saved[k]; });
    ST.state = merged;
  };
  ST.save = function () {
    try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(ST.state)); } catch (e) { /* storage unavailable */ }
  };
  ST.resetKeys = function (keys) {
    keys.forEach(function (k) { ST.state[k] = clone(ST.defaults[k]); });
    ST.save();
  };
  function clone(v) { return v && typeof v === "object" ? JSON.parse(JSON.stringify(v)) : v; }

  // ---------- Formula evaluator ----------
  // Supports + - * / brackets, numbers, names, and implicit multiplication like 2col.
  ST.evaluate = function (text, vars, ownName) {
    var t = String(text || "").toLowerCase().replace(/\s+/g, "").replace(/"/g, "").replace(/in/g, "");
    if (!t) return null;
    if (ownName && "*/+".indexOf(t.charAt(0)) >= 0) t = ownName + t;
    var pos = 0;
    function peek() { return t.charAt(pos); }
    function fail() { throw new Error("bad"); }
    function sum() {
      var v = product();
      while (peek() === "+" || peek() === "-") {
        var op = t.charAt(pos++);
        v = op === "+" ? v + product() : v - product();
      }
      return v;
    }
    function product() {
      var v = factor();
      for (;;) {
        var c = peek();
        if (c === "*") { pos++; v *= factor(); }
        else if (c === "/") { pos++; var d = factor(); if (d === 0) fail(); v /= d; }
        else if (c === "(" || /[a-z]/.test(c)) { v *= factor(); }
        else break;
      }
      return v;
    }
    function factor() {
      var c = peek();
      if (c === "-") { pos++; return -factor(); }
      if (c === "+") { pos++; return factor(); }
      if (c === "(") { pos++; var v = sum(); if (peek() !== ")") fail(); pos++; return v; }
      var m = /^\d*\.?\d+|^\d+\.?/.exec(t.slice(pos));
      if (m) { pos += m[0].length; return parseFloat(m[0]); }
      m = /^[a-z]+/.exec(t.slice(pos));
      if (m) {
        pos += m[0].length;
        if (!vars || !(m[0] in vars) || vars[m[0]] === null) fail();
        return vars[m[0]];
      }
      fail();
    }
    try {
      var result = sum();
      if (pos !== t.length || !isFinite(result)) return null;
      return result;
    } catch (e) { return null; }
  };
  ST.num = function (text) {
    var v = ST.evaluate(text, null, null);
    return v === null ? 0 : v;
  };
  ST.nice = function (x) { return String(Math.round(x * 1000) / 1000); };
  ST.plural = function (n, word) { return n + " " + word + (n === 1 ? "" : "s"); };

  // ---------- Grid maths ----------
  ST.presetById = function (id) {
    var all = ST.PRESETS.concat(ST.state.customPresets || []);
    return all.filter(function (p) { return p.id === id; })[0];
  };
  ST.calculate = function () {
    var s = ST.state, num = ST.num;
    var spacing = num(s.spacing);
    var maxW = num(s.maxW), maxH = num(s.maxH);
    var cols = Math.min(ST.MAX_COLS, Math.max(1, Math.floor(num(s.cols)) || 1));
    var rows = Math.min(ST.MAX_ROWS, Math.max(1, Math.floor(num(s.rows)) || 1));
    var gapH = spacing * ST.FRACTIONS[s.hIdx];
    var gapV = spacing * ST.FRACTIONS[s.vIdx];
    var available = maxW - (cols - 1) * gapH;
    var cellH = (maxH - (rows - 1) * gapV) / rows;
    var preset = ST.presetById(s.preset);
    var widths = [];
    var weighted = !!(preset && preset.weights && cols === 2);
    if (weighted) {
      var total = preset.weights[0] + preset.weights[1];
      widths = preset.weights.map(function (w) { return available * w / total; });
    } else {
      for (var i = 0; i < cols; i++) widths.push(available / cols);
    }
    var valid = maxW > 0 && maxH > 0 && cellH > 0 && widths.every(function (w) { return w > 0; });
    ST.grid = { spacing: spacing, maxW: maxW, maxH: maxH, cols: cols, rows: rows, gapH: gapH, gapV: gapV,
                widths: widths, cellH: cellH, valid: valid, weighted: weighted,
                left: num(s.left), top: num(s.top) };
    return ST.grid;
  };
  ST.formulaVars = function (shapeW, shapeH) {
    var g = ST.grid;
    return { g: g.spacing, col: g.widths[0], row: g.cellH, gap: g.gapH, vgap: g.gapV,
             w: shapeW === undefined ? null : shapeW, h: shapeH === undefined ? null : shapeH };
  };

  // Call after any value changes: recalculates, saves and lets every tool update.
  ST.changed = function () {
    ST.calculate();
    ST.save();
    ST.emit("change");
  };

  // ---------- Messages ----------
  ST.toast = function (message, isBad) {
    var t = ST.$("toast");
    t.textContent = message;
    t.className = "toast show" + (isBad ? " bad" : "");
    clearTimeout(ST.toast.timer);
    ST.toast.timer = setTimeout(function () { t.className = "toast" + (isBad ? " bad" : ""); }, 3200);
  };

  // ---------- PowerPoint ----------
  ST.supports = function (version) {
    try { return ST.inOffice && Office.context.requirements.isSetSupported("PowerPointApi", version); }
    catch (e) { return false; }
  };
  ST.needOffice = function () {
    if (!ST.inOffice) { ST.toast("Open this panel inside PowerPoint to use this.", true); return false; }
    return true;
  };
  ST.run = function (task) {
    return PowerPoint.run(task).catch(function (e) {
      ST.toast("PowerPoint said: " + (e && e.message ? e.message : e), true);
    });
  };
  // Loads the selected shapes with the given properties, or tells the user to select something.
  ST.withSelection = function (props, task, emptyMessage) {
    if (!ST.needOffice()) return;
    return ST.run(function (context) {
      var shapes = context.presentation.getSelectedShapes();
      shapes.load(props.map(function (p) { return "items/" + p; }).join(","));
      return context.sync().then(function () {
        if (!shapes.items.length) { ST.toast(emptyMessage || "Select the shapes on your slide first.", true); return; }
        return task(context, shapes.items.slice());
      });
    });
  };

  ST.refreshSelection = function () {
    if (!ST.inOffice) return;
    PowerPoint.run(function (context) {
      var shapes = context.presentation.getSelectedShapes();
      shapes.load("items/left,items/top,items/width,items/height");
      return context.sync().then(function () {
        var sel = ST.selection;
        sel.items = shapes.items.map(function (s) {
          return { left: s.left / PT, top: s.top / PT, w: s.width / PT, h: s.height / PT };
        });
        sel.count = sel.items.length;
        if (sel.count) { sel.w = sel.items[0].w; sel.h = sel.items[0].h; }
        ST.emit("selection");
      });
    }).catch(function () {
      ST.selection.count = 0; ST.selection.items = [];
      ST.emit("selection");
    });
  };

  // Converts a corner radius in points into PowerPoint's 0 to 0.5 adjustment value.
  ST.cornerValue = function (widthPt, heightPt, radiusPt) {
    var radius = Math.max(0, radiusPt === undefined ? ST.num(ST.state.radius) : radiusPt);
    return Math.min(0.5, radius / Math.min(widthPt, heightPt));
  };
  ST.cornersSupported = function () { return ST.supports("1.10"); };

  // ---------- Shared field helper ----------
  // Wires a text box to a state key. Number boxes accept maths and tidy up on Enter or blur.
  ST.bindField = function (input, key, opts) {
    opts = opts || {};
    input.value = ST.state[key];
    input.addEventListener("input", function () {
      ST.state[key] = input.value;
      if (opts.number) input.classList.toggle("err", input.value.trim() !== "" && ST.evaluate(input.value, null, null) === null);
      if (opts.onInput) opts.onInput(input.value);
      ST.changed();
    });
    if (opts.number) {
      var commit = function () {
        var v = ST.evaluate(input.value, null, null);
        if (v !== null && input.value.trim() !== ST.nice(v)) {
          input.value = ST.nice(v); ST.state[key] = input.value; ST.changed();
        }
        input.classList.remove("err");
      };
      input.addEventListener("blur", commit);
      input.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); commit(); } });
    }
    ST.on("reset", function () { input.value = ST.state[key]; input.classList.remove("err"); });
  };

  // Small HTML builder so tool files stay readable.
  ST.html = function (markup) {
    var t = document.createElement("template");
    t.innerHTML = markup.trim();
    return t.content;
  };

  ST.registerTool = function (tool) { ST.tools.push(tool); };
})();
