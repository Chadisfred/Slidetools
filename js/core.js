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
    FRACTIONS: [0.125, 0.25, 0.5, 0.75, 1, 0],
    FRACTION_ORDER: [5, 0, 1, 2, 3, 4], // None first; indexes stay stable for saved settings
    FRACTION_LABELS: ["1/8", "1/4", "1/2", "3/4", "1", "None"],
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
    var t = String(text || "").toLowerCase().replace(/\s+/g, "").replace(/"/g, "");
    // Units: points, centimetres and millimetres become inches; "in" is just dropped
    t = t.replace(/(\d*\.?\d+)pt/g, "($1/72)").replace(/(\d*\.?\d+)cm/g, "($1/2.54)").replace(/(\d*\.?\d+)mm/g, "($1/25.4)").replace(/in/g, "");
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
  // Explains why a formula can't be read, in a few words, or returns "" when it reads fine.
  ST.formulaProblem = function (text, vars) {
    var raw = String(text || "").trim();
    if (!raw || ST.evaluate(raw, vars, null) !== null) return "";
    var t = raw.toLowerCase().replace(/\s+/g, "");
    var last = t.charAt(t.length - 1);
    if ("+-*/(".indexOf(last) >= 0) return "ends with " + last;
    var open = (t.match(/\(/g) || []).length, close = (t.match(/\)/g) || []).length;
    if (open !== close) return "has a bracket that isn't closed";
    var names = (t.replace(/(\d*\.?\d+)(pt|cm|mm|in)/g, "$1").match(/[a-z]+/g) || []);
    for (var i = 0; i < names.length; i++) {
      if (!vars || !(names[i] in vars)) return "uses " + names[i] + ", which isn't a name it knows";
      if (vars[names[i]] === null) return "uses " + names[i] + ", which needs a selected shape";
    }
    return "can't be read";
  };
  ST.num = function (text) {
    var v = ST.evaluate(text, null, null);
    return v === null ? 0 : v;
  };
  ST.nice = function (x) { return String(Math.round(x * 1000) / 1000); };
  // Five decimal places, enough that 30 pt (0.41667 in) steps don't drift across a slide.
  ST.exact = function (x) { return String(Math.round(x * 100000) / 100000); };
  ST.plural = function (n, word) { return n + " " + word + (n === 1 ? "" : "s"); };

  // ---------- Grid maths ----------
  ST.presetById = function (id) {
    var all = ST.PRESETS.concat(ST.state.customPresets || []);
    return all.filter(function (p) { return p.id === id; })[0];
  };
  // "2:1:1" (or "2 1 1", "2,1,1") becomes [2, 1, 1]. Empty gives [], anything else null.
  ST.parseRatio = function (text, max) {
    var t = String(text || "").trim();
    if (!t) return [];
    var parts = t.split(/[\s:,;\/]+/).filter(Boolean).map(function (x) { return ST.evaluate(x, null, null); });
    if (parts.length < 1 || parts.length > max || parts.some(function (v) { return !(v > 0); })) return null;
    return parts;
  };
  function share(total, parts) { var sum = parts.reduce(function (a, b) { return a + b; }, 0); return parts.map(function (p) { return total * p / sum; }); }
  ST.calculate = function () {
    var s = ST.state, num = ST.num;
    var spacing = num(s.spacing);
    var maxW = num(s.maxW), maxH = num(s.maxH);
    var colParts = ST.parseRatio(s.colRatio, ST.MAX_COLS), rowParts = ST.parseRatio(s.rowRatio, ST.MAX_ROWS);
    var cols = colParts && colParts.length ? colParts.length : Math.min(ST.MAX_COLS, Math.max(1, Math.floor(num(s.cols)) || 1));
    var rows = rowParts && rowParts.length ? rowParts.length : Math.min(ST.MAX_ROWS, Math.max(1, Math.floor(num(s.rows)) || 1));
    var gapH = spacing * (ST.FRACTIONS[s.hIdx] || 0);
    var gapV = spacing * (ST.FRACTIONS[s.vIdx] || 0);
    var available = maxW - (cols - 1) * gapH, availableH = maxH - (rows - 1) * gapV;
    var preset = ST.presetById(s.preset);
    var widths, heights, i;
    if (colParts && colParts.length) widths = share(available, colParts);
    else if (preset && preset.weights && cols === 2) widths = share(available, preset.weights);
    else { widths = []; for (i = 0; i < cols; i++) widths.push(available / cols); }
    if (rowParts && rowParts.length) heights = share(availableH, rowParts);
    else { heights = []; for (i = 0; i < rows; i++) heights.push(availableH / rows); }
    var uneven = function (list) { return list.some(function (v) { return Math.abs(v - list[0]) > 1e-9; }); };
    var valid = maxW > 0 && maxH > 0 && widths.every(function (w) { return w > 0; }) && heights.every(function (h) { return h > 0; });
    ST.grid = { spacing: spacing, maxW: maxW, maxH: maxH, cols: cols, rows: rows, gapH: gapH, gapV: gapV,
                widths: widths, heights: heights, cellH: heights[0], valid: valid,
                weighted: uneven(widths), tall: uneven(heights), ratioBad: { cols: colParts === null, rows: rowParts === null },
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
  // Plain sentences for the errors PowerPoint reports most often. Anything else shows its own message.
  var ERRORS = {
    ItemNotFound: "A shape or slide changed while this was running. Select it again and retry.",
    InvalidArgument: "PowerPoint didn't accept one of the values. Check the numbers and retry.",
    InvalidSelection: "PowerPoint can't use the current selection. Click a slide or shape and retry.",
    InvalidOperation: "PowerPoint can't do that to this kind of shape.",
    AccessDenied: "The presentation is read only or locked, so it can't be changed.",
    NotImplemented: "This PowerPoint version can't do that yet.",
    ApiNotFound: "This PowerPoint version can't do that yet.",
    GeneralException: "PowerPoint couldn't finish that. Retry, or select fewer shapes."
  };
  // The reason on its own, with no full stop, for use inside a longer message.
  ST.reason = function (e) {
    var code = e && e.code;
    var text = code && ERRORS[code] ? ERRORS[code] : String(e && e.message ? e.message : e);
    return text.replace(/\.\s*$/, "");
  };
  ST.errorText = function (e) {
    var code = e && e.code;
    if (code && ERRORS[code]) return ERRORS[code];
    return "PowerPoint said: " + ST.reason(e) + ".";
  };
  ST.run = function (task) {
    return PowerPoint.run(task).catch(function (e) {
      if (window.console) console.error(e);
      ST.toast(ST.errorText(e), true);
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

  // Reads the selection's sizes, and what kind of thing is selected:
  // picture, table, text (the cursor is in a text box), shape, or none.
  ST.refreshSelection = function () {
    if (!ST.inOffice) return;
    PowerPoint.run(function (context) {
      var shapes = context.presentation.getSelectedShapes();
      shapes.load("items/left,items/top,items/width,items/height,items/type");
      return context.sync().then(function () {
        var sel = ST.selection;
        sel.items = shapes.items.map(function (s) {
          return { left: s.left / PT, top: s.top / PT, w: s.width / PT, h: s.height / PT, line: s.type === "Line" || s.width === 0 || s.height === 0 };
        });
        sel.count = sel.items.length;
        if (sel.count) { sel.w = sel.items[0].w; sel.h = sel.items[0].h; sel.line = sel.items[0].line; }
        var types = shapes.items.map(function (s) { return s.type; });
        return selectionKind(context, shapes.items, types).then(function (kind) { sel.kind = kind; ST.emit("selection"); });
      });
    }).catch(function () {
      ST.selection.count = 0; ST.selection.items = []; ST.selection.kind = "none";
      ST.emit("selection");
    });
  };
  function selectionKind(context, items, types) {
    if (!items.length) return Promise.resolve("none");
    if (types.every(function (t) { return t === "Table"; })) return Promise.resolve("table");
    if (types.every(function (t) { return t === "Image"; })) return Promise.resolve("picture");
    return textCursor(context, items).then(function (inText) {
      if (inText) return "text";
      // A picture placeholder counts as a picture
      if (items.length !== 1 || types[0] !== "Placeholder" || !ST.supports("1.8")) return "shape";
      var ph = items[0].placeholderFormat;
      ph.load("type");
      return context.sync().then(function () { return ph.type === "Picture" ? "picture" : "shape"; }, function () { return "shape"; });
    });
  }
  // True when the cursor or highlighted text is inside a text box, not when whole shapes are selected.
  function textCursor(context, items) {
    if (items.length !== 1 || !ST.supports("1.5")) return Promise.resolve(false);
    try {
      var pres = context.presentation;
      var range = pres.getSelectedTextRangeOrNullObject ? pres.getSelectedTextRangeOrNullObject() : pres.getSelectedTextRange();
      range.load("start,length");
      return context.sync().then(function () { return !range.isNullObject; }, function () { return false; });
    } catch (e) { return Promise.resolve(false); }
  }

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
        if (v !== null && input.value.trim() !== ST.exact(v)) {
          input.value = ST.exact(v); ST.state[key] = input.value; ST.changed();
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

  // ---------- Help ----------
  // A small i icon with a hover tip, for explaining one setting. text may contain simple HTML.
  ST.info = function (text, label) {
    return '<span class="info inline" tabindex="0" role="button" aria-label="' + (label || "More about this") + '">i' +
      '<span class="tip" role="tooltip">' + text + "</span></span>";
  };
  // Opens the user manual at a section, in the system browser where Office allows it.
  ST.SITE_URL = "https://chadisfred.github.io/Slidetools/";
  // Each tool has its own manual page, named after the tool's id.
  ST.openHelp = function (page) {
    var url = ST.SITE_URL + (page ? "manual/" + page + ".html" : "manual.html");
    try {
      if (window.Office && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; }
    } catch (e) { /* fall back to a normal window */ }
    window.open(url, "_blank", "noopener");
  };
})();
