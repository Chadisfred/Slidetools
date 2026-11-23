/* Typography tool: font, size, style and case, exact character spacing, line spacing,
   paragraph spacing, kerning and baseline, plus text box margins, wrap and autofit.
   Font settings, margins, wrap and autofit use the PowerPoint API and apply straight away.
   The API has no spacing, kerning or baseline settings, so those are written to the slide
   XML through the shared slide rebuild (js/rebuild.js), in memory.
   Point boxes take formulas in points: 2g, g/2, 12pt+2, and in Size, +2 or *1.5. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;
  var R = ST.rebuild, NS_P = R.NS_P, NS_A = R.NS_A, kid = R.kid, kids = R.kids, make = R.make;
  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 14L7 4l4 10M4.5 10.5h5"/><path d="M13 7v7M11.5 14h3"/></svg>';
  var TEXT_TYPES = ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"];
  var PPR_ORDER = ["lnSpc", "spcBef", "spcAft"];

  var defaults = { tyFontName: "", tyFontSize: "", tySpc: "", tyLine: "", tyLineUnit: "pt", tyBefore: "", tyAfter: "", tyKern: "", tyBase: "",
                   tyMl: "", tyMr: "", tyMt: "", tyMb: "", tyPreviewOn: true };
  ST.addDefaults(defaults);
  var PT_KEYS = ["tyFontSize", "tySpc", "tyBefore", "tyAfter", "tyKern", "tyMl", "tyMr", "tyMt", "tyMb"];

  // A number in points from a box, "" when empty, null when it can't be read.
  // Points boxes take grid names (g, col, gap) in points. own: the current value, for +2 or *1.5.
  function num(text, own) {
    var t = String(text || "").trim();
    if (!t) return "";
    var v = ST.evaluatePt(t, own);
    return v === null || !isFinite(v) ? null : v;
  }
  // A plain number, for Lines and Baseline %
  function plain(text) {
    var t = String(text || "").trim().replace(/\s*%$/, "");
    if (!t) return "";
    var v = ST.evaluate(t, null, null);
    return v === null || !isFinite(v) ? null : v;
  }
  function isPtBox(key, S) { return PT_KEYS.indexOf(key) >= 0 || (key === "tyLine" && (S || ST.state).tyLineUnit !== "x"); }

  // ---------- Reading the panel (pure) ----------
  // Returns { spec } or { bad }. spec values are undefined when the box is empty (leave as is).
  function readSpec(S) {
    var out = {}, bad = null;
    function take(key, label, test, msg) {
      var pt = isPtBox(key, S), v = pt ? num(S[key]) : plain(S[key]);
      if (v === "") return undefined;
      if (v === null && pt) { var why = ST.ptProblem(S[key]); bad = bad || label + " " + (why || "can\'t be read") + ". Use points or names such as g, g/2 or gap."; return undefined; }
      if (v === null || !test(v)) { bad = bad || label + " " + msg; return undefined; }
      return v;
    }
    out.spc = take("tySpc", "Character spacing", function (v) { return v >= -100 && v <= 400; }, "must be a number of points, from -100 to 400.");
    var line = take("tyLine", "Line spacing", function (v) { return v > 0 && v <= (S.tyLineUnit === "x" ? 9.99 : 1584); }, S.tyLineUnit === "x" ? "must be a multiple from 0.1 to 9.99, such as 0.9." : "must be more than 0 points.");
    if (line !== undefined) out.line = { unit: S.tyLineUnit === "x" ? "x" : "pt", value: line };
    out.before = take("tyBefore", "Space before", function (v) { return v >= 0 && v <= 1584; }, "must be 0 or more points.");
    out.after = take("tyAfter", "Space after", function (v) { return v >= 0 && v <= 1584; }, "must be 0 or more points.");
    out.kern = take("tyKern", "Kerning", function (v) { return v >= 0 && v <= 4000; }, "must be a font size in points, or 0 for off.");
    out.base = take("tyBase", "Baseline", function (v) { return v >= -100 && v <= 100; }, "must be a percentage from -100 to 100.");
    if (bad) return { bad: bad };
    var any = Object.keys(out).some(function (k) { return out[k] !== undefined; });
    return any ? { spec: out } : { bad: "Fill in at least one spacing box first." };
  }

  // ---------- Slide XML ----------
  function ensureFirst(doc, parent, name) {
    var el = kid(parent, NS_A, name);
    if (!el) { el = make(doc, NS_A, name); parent.insertBefore(el, parent.firstChild); }
    return el;
  }
  // Sets lnSpc, spcBef or spcAft in a paragraph's pPr, keeping the schema's order
  function setSpacing(doc, pPr, name, kind, val) {
    kids(pPr, NS_A, name).forEach(function (n) { pPr.removeChild(n); });
    var el = make(doc, NS_A, name), v = make(doc, NS_A, kind);
    v.setAttribute("val", String(Math.round(val))); el.appendChild(v);
    var after = PPR_ORDER.slice(PPR_ORDER.indexOf(name) + 1), ref = null;
    Array.prototype.some.call(pPr.childNodes, function (n) {
      if (n.nodeType === 1 && (after.indexOf(n.localName) >= 0 || PPR_ORDER.indexOf(n.localName) < 0)) { ref = n; return true; }
      return false;
    });
    pPr.insertBefore(el, ref);
  }
  // Applies spec to one text body. Returns true when something changed.
  function applyToBody(doc, body, spec) {
    var paras = kids(body, NS_A, "p");
    if (!paras.length) return false;
    paras.forEach(function (p) {
      if (spec.line || spec.before !== undefined || spec.after !== undefined) {
        var pPr = ensureFirst(doc, p, "pPr");
        if (spec.line) setSpacing(doc, pPr, "lnSpc", spec.line.unit === "x" ? "spcPct" : "spcPts", spec.line.unit === "x" ? spec.line.value * 100000 : spec.line.value * 100);
        if (spec.before !== undefined) setSpacing(doc, pPr, "spcBef", "spcPts", spec.before * 100);
        if (spec.after !== undefined) setSpacing(doc, pPr, "spcAft", "spcPts", spec.after * 100);
      }
      if (spec.spc !== undefined || spec.kern !== undefined || spec.base !== undefined) {
        var props = [];
        kids(p, NS_A).forEach(function (n) {
          if (n.localName === "r" || n.localName === "fld" || n.localName === "br") props.push(ensureFirst(doc, n, "rPr"));
        });
        var end = kid(p, NS_A, "endParaRPr");
        if (!end) { end = make(doc, NS_A, "endParaRPr"); p.appendChild(end); }
        props.push(end);
        props.forEach(function (rPr) {
          if (spec.spc !== undefined) { var s = Math.round(spec.spc * 100); if (s) rPr.setAttribute("spc", s); else rPr.removeAttribute("spc"); }
          if (spec.kern !== undefined) rPr.setAttribute("kern", Math.round(spec.kern * 100));
          if (spec.base !== undefined) { var b = Math.round(spec.base * 1000); if (b) rPr.setAttribute("baseline", b); else rPr.removeAttribute("baseline"); }
        });
      }
    });
    return true;
  }
  // The first explicit values in a text body, for Read from selection
  function readBody(body) {
    var out = {}, p = kid(body, NS_A, "p"), pPr = p && kid(p, NS_A, "pPr");
    function spacing(name) {
      var el = pPr && kid(pPr, NS_A, name), pts = el && kid(el, NS_A, "spcPts"), pct = el && kid(el, NS_A, "spcPct");
      if (pts) return { unit: "pt", value: +pts.getAttribute("val") / 100 };
      if (pct) return { unit: "x", value: +pct.getAttribute("val") / 100000 };
      return null;
    }
    out.line = spacing("lnSpc"); var b = spacing("spcBef"), a = spacing("spcAft");
    out.before = b && b.unit === "pt" ? b.value : null; out.after = a && a.unit === "pt" ? a.value : null;
    var r = p && (kid(p, NS_A, "r") || kid(p, NS_A, "fld")), rPr = r && kid(r, NS_A, "rPr");
    out.spc = rPr && rPr.getAttribute("spc") !== null ? +rPr.getAttribute("spc") / 100 : null;
    out.kern = rPr && rPr.getAttribute("kern") !== null ? +rPr.getAttribute("kern") / 100 : null;
    out.base = rPr && rPr.getAttribute("baseline") !== null ? +rPr.getAttribute("baseline") / 1000 : null;
    return out;
  }

  function boxesOf(items) { return items.map(function (s) { return { left: s.left, top: s.top, width: s.width, height: s.height, name: s.name, id: s.id }; }); }
  function textItems(items) { return items.filter(function (s) { return TEXT_TYPES.indexOf(s.type) >= 0; }); }

  // ---------- Spacing (slide rebuild) ----------
  function applySpacing(saved) {
    var S = saved || {};
    if (!saved) Object.keys(defaults).forEach(function (k) { S[k] = ST.state[k]; });
    var r = readSpec(S);
    if (r.bad) { ST.toast(r.bad, true); return; }
    if (!ST.needOffice()) return;
    if (!R.supported()) { ST.toast("This version of PowerPoint can't change text spacing. It needs PowerPoint API 1.8.", true); return; }
    ST.setRepeat("Apply text spacing", function () { applySpacing(S); });
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      var list = textItems(items);
      if (!list.length) { ST.toast("Select shapes or text boxes with text.", true); return; }
      return ST.confirmSwap("change text spacing").then(function (ok) {
        if (!ok) return null;
        var names = [];
        return R.edit(context, function (doc) {
          var boxes = boxesOf(list), els = R.matchShapes(doc, boxes, ["sp"]), n = 0;
          els.forEach(function (el, i) { var body = el && kid(el, NS_P, "txBody"); if (body && applyToBody(doc, body, r.spec)) { n++; names.push(boxes[i].name); } });
          return { count: n };
        }).then(function (res) {
          if (!res.slideId || !names.length || !ST.supports("1.5")) return res.count;
          var slide = context.presentation.slides.getItem(res.slideId), shapes = slide.shapes;
          shapes.load("items/id,items/name");
          return context.sync().then(function () {
            var ids = shapes.items.filter(function (x) { return names.indexOf(x.name) >= 0; }).map(function (x) { return x.id; });
            if (ids.length) { slide.setSelectedShapes(ids); return context.sync(); }
          }).catch(function () {}).then(function () { return res.count; });
        });
      }).then(function (n) {
        if (n === null) { ST.toast("Nothing was changed."); return; }
        ST.toast(n ? "Set the spacing in " + ST.plural(n, "shape") + "." : "Couldn't find text in the selected shapes.", !n);
        ST.refreshSelection(); queuePreview(true);
      }, function (e) { ST.toast("Couldn't change the slide. " + (ST.errorText ? ST.errorText(e) : ""), true); });
    });
  }

  // ---------- Text box (API) ----------
  function applyMargins() {
    var m = { leftMargin: num(ST.state.tyMl), rightMargin: num(ST.state.tyMr), topMargin: num(ST.state.tyMt), bottomMargin: num(ST.state.tyMb) };
    var keys = Object.keys(m).filter(function (k) { return m[k] !== ""; });
    if (keys.some(function (k) { return m[k] === null || m[k] < 0; })) { ST.toast("Margins must be 0 or more points, or names such as g or g/2.", true); return; }
    if (!keys.length) { ST.toast("Fill in at least one margin first.", true); return; }
    if (!ST.needOffice()) return;
    ST.setRepeat("Apply margins", applyMargins);
    ST.withSelection(["type"], function (context, items) {
      var list = textItems(items);
      if (!list.length) { ST.toast("Select shapes or text boxes.", true); return; }
      list.forEach(function (s) { keys.forEach(function (k) { s.textFrame[k] = Math.round(m[k] * 100) / 100; }); });
      return context.sync().then(function () { ST.toast("Set the margins of " + ST.plural(list.length, "shape") + "."); queuePreview(); });
    });
  }
  function setFrame(prop, value, label) {
    if (!ST.needOffice()) return;
    ST.withSelection(["type"], function (context, items) {
      var list = textItems(items);
      if (!list.length) { ST.toast("Select shapes or text boxes.", true); return; }
      list.forEach(function (s) { s.textFrame[prop] = value; });
      return context.sync().then(function () { ST.toast(label + " for " + ST.plural(list.length, "shape") + "."); queuePreview(); });
    });
  }

  // ---------- Read from selection ----------
  function readSelection() {
    if (!ST.needOffice()) return;
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      var s = textItems(items)[0];
      if (!s) { ST.toast("Select a shape or text box with text.", true); return; }
      var tf = s.textFrame;
      tf.load("leftMargin,rightMargin,topMargin,bottomMargin");
      return context.sync().then(function () {
        var nice = function (v) { return v === null || v === undefined ? "" : ST.nice(v); };
        ST.state.tyMl = nice(tf.leftMargin); ST.state.tyMr = nice(tf.rightMargin); ST.state.tyMt = nice(tf.topMargin); ST.state.tyMb = nice(tf.bottomMargin);
        if (!R.supported()) return null;
        var found = null;
        // Export only: nothing is changed or swapped when the edit returns a count of 0
        return R.edit(context, function (doc) {
          var el = R.matchShapes(doc, boxesOf([s]), ["sp"])[0], body = el && kid(el, NS_P, "txBody");
          if (body) found = readBody(body);
          return { count: 0 };
        }).then(function () { return found; });
      }).then(function (v) {
        if (v) {
          var nice = function (x) { return x === null || x === undefined ? "" : ST.nice(x); };
          ST.state.tySpc = nice(v.spc); ST.state.tyBefore = nice(v.before); ST.state.tyAfter = nice(v.after);
          ST.state.tyKern = nice(v.kern); ST.state.tyBase = nice(v.base);
          ST.state.tyLine = v.line ? ST.nice(v.line.value) : ""; if (v.line) ST.state.tyLineUnit = v.line.unit;
        }
        ST.save(); ST.emit("reset"); ST.changed();
        ST.toast("Read the text settings from " + (s.name || "the selection") + ". Empty boxes follow the template.");
      });
    });
  }

  // ---------- Live preview ----------
  // A scaled copy of the first selected text box, drawn with the values in the boxes. It only
  // reads: the current spacing comes from an export that never swaps the slide.
  var pv = { info: null, cache: {}, timer: null, busy: false, again: false };
  var PX = 4 / 3; // pixels per point
  function esc(t) { return String(t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function previewOn() { return ST.state.tyPreviewOn !== false; }
  function previewVisible() { var b = $("tyPreview"); return !!b && !!b.offsetParent && previewOn(); }
  function queuePreview(fresh) {
    if (fresh) pv.cache = {};
    clearTimeout(pv.timer);
    pv.timer = setTimeout(loadPreview, 200);
  }
  function loadPreview() {
    if (!ST.inOffice || !previewVisible() || !ST.supports("1.4")) { renderPreview(); return; }
    if (pv.busy) { pv.again = true; return; }
    pv.busy = true;
    PowerPoint.run(function (context) {
      var shapes = context.presentation.getSelectedShapes();
      shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
      return context.sync().then(function () {
        var s = textItems(shapes.items)[0];
        if (!s) return null;
        var tf = s.textFrame, tr = tf.textRange, info = { id: s.id, name: s.name, box: { left: s.left, top: s.top, width: s.width, height: s.height } };
        tf.load("leftMargin,rightMargin,topMargin,bottomMargin,wordWrap,verticalAlignment");
        tr.load("text"); tr.font.load("name,size,color,bold,italic,underline"); tr.paragraphFormat.load("horizontalAlignment");
        s.fill.load("type,foregroundColor");
        return context.sync().then(function () {
          info.text = String(tr.text || "").slice(0, 600);
          info.margins = { l: tf.leftMargin, r: tf.rightMargin, t: tf.topMargin, b: tf.bottomMargin };
          info.wrap = tf.wordWrap !== false; info.valign = tf.verticalAlignment || "Top";
          info.align = tr.paragraphFormat.horizontalAlignment || "Left";
          info.fill = s.fill.type === "Solid" ? s.fill.foregroundColor : null;
          info.font = { name: tr.font.name, size: tr.font.size, color: tr.font.color, bold: tr.font.bold, italic: tr.font.italic, underline: tr.font.underline };
          info.mixed = { name: !tr.font.name, size: !tr.font.size, color: !tr.font.color };
          return readCase(context, tr).then(function (c) { info.font.allCaps = c.allCaps; info.font.smallCaps = c.smallCaps; info.font.strikethrough = c.strikethrough; });
        }).then(function () {
          if (info.font.name && info.font.size && info.font.color) return;
          // Mixed formatting: use the first character's font
          var first = tr.getSubstring(0, 1);
          first.font.load("name,size,color,bold,italic");
          return context.sync().then(function () {
            ["name", "size", "color", "bold", "italic"].forEach(function (k) { if (info.font[k] === null || info.font[k] === undefined) info.font[k] = first.font[k]; });
          }).catch(function () {});
        }).then(function () {
          if (pv.cache[s.id] !== undefined || !R.supported()) { info.cur = pv.cache[s.id] || null; return info; }
          return R.edit(context, function (doc) {
            var el = R.matchShapes(doc, boxesOf([s]), ["sp"])[0], body = el && kid(el, NS_P, "txBody");
            pv.cache[s.id] = body ? readBody(body) : null;
            pv.fonts = fontsIn(doc, arguments[1] && arguments[1].files);
            renderFontList();
            return { count: 0 };
          }).then(function () { info.cur = pv.cache[s.id]; return info; }, function () { info.cur = null; return info; });
        });
      });
    }).then(function (info) { pv.info = info; }, function () { pv.info = null; })
      .then(function () {
        pv.busy = false; renderPreview(); renderFontState();
        if (pv.again) { pv.again = false; queuePreview(); }
      });
  }
  // The values the preview shows: a filled box wins, then the shape's own setting, then the default
  function effective(info) {
    var cur = info.cur || {}, S = ST.state, pick = function (key, own, dflt) { var v = key === "tyBase" ? plain(S[key]) : num(S[key]); return typeof v === "number" ? v : (own !== null && own !== undefined ? own : dflt); };
    var lineBox = S.tyLineUnit === "x" ? plain(S.tyLine) : num(S.tyLine), line = typeof lineBox === "number" && lineBox > 0 ? { unit: S.tyLineUnit === "x" ? "x" : "pt", value: lineBox } : (cur.line || { unit: "x", value: 1 });
    var m = info.margins, margin = function (key, own, dflt) { var v = num(S[key]); return typeof v === "number" && v >= 0 ? v : (typeof own === "number" ? own : dflt); };
    return { spc: pick("tySpc", cur.spc, 0), before: pick("tyBefore", cur.before, 0), after: pick("tyAfter", cur.after, 0), base: pick("tyBase", cur.base, 0), line: line,
             ml: margin("tyMl", m.l, 7.2), mr: margin("tyMr", m.r, 7.2), mt: margin("tyMt", m.t, 3.6), mb: margin("tyMb", m.b, 3.6) };
  }
  function renderPreview() {
    var box = $("tyPreview");
    if (!box) return;
    box.hidden = !previewOn();
    if (box.hidden) return;
    var info = pv.info;
    if (!info) { box.innerHTML = '<div class="muted-note">Select a text box to preview its text with these settings.</div>'; return; }
    var e = effective(info), f = info.font || {}, typed = num(S_size(), f.size || 18), size = typeof typed === "number" && typed >= 1 ? typed : (f.size || 18), W = Math.max(info.box.width, 1), H = Math.max(info.box.height, 1);
    var lh = e.line.unit === "pt" ? e.line.value + "pt" : String(e.line.value * 1.2);
    var family = String(ST.state.tyFontName || "").trim() || f.name || "Calibri";
    var deco = [f.underline && f.underline !== "None" ? "underline" : "", f.strikethrough ? "line-through" : ""].filter(Boolean).join(" ") || "none";
    var align = { Center: "center", Centered: "center", Right: "right", Justify: "justify", JustifyLow: "justify", Distributed: "justify" }[info.align] || "left";
    var justify = /Middle|Center/.test(info.valign) ? "center" : /Bottom/.test(info.valign) ? "flex-end" : "flex-start";
    var paras = (info.text || "").split(/\r\n|\r|\n/).map(function (p) {
      var inner = p ? p.split("\v").map(esc).join("<br>") : "&nbsp;";
      return '<p style="margin:' + e.before + 'pt 0 ' + e.after + 'pt;line-height:' + lh + '"><span style="position:relative;top:' + (-e.base / 100) + 'em">' + inner + '</span></p>';
    }).join("");
    var avail = Math.max(120, box.clientWidth - 2), k = Math.min(1, avail / (W * PX));
    box.innerHTML = '<div class="ty-stage"><div class="ty-page" style="width:' + W + 'pt;min-height:' + H + 'pt;padding:' + e.mt + 'pt ' + e.mr + 'pt ' + e.mb + 'pt ' + e.ml + 'pt;' +
      'font-family:&quot;' + esc(family) + '&quot;, sans-serif;font-size:' + size + 'pt;color:' + esc(f.color || "#1d1d1f") + ';font-weight:' + (f.bold ? 700 : 400) + ';font-style:' + (f.italic ? "italic" : "normal") + ';' +
      'text-decoration:' + deco + ';text-transform:' + (f.allCaps ? "uppercase" : "none") + ';font-variant:' + (f.smallCaps ? "small-caps" : "normal") + ';' +
      'letter-spacing:' + e.spc + 'pt;text-align:' + align + ';justify-content:' + justify + ';white-space:' + (info.wrap ? "pre-wrap" : "pre") + ';background:' + esc(info.fill || "#ffffff") + ';transform:scale(' + k + ')">' +
      paras + '<i class="ty-bounds" style="height:' + H + 'pt"></i></div></div><div class="ty-caption">' + esc(info.name || "Selected shape") + ', ' + Math.round(k * 100) + '% size</div>';
    var page = box.querySelector(".ty-page");
    box.querySelector(".ty-stage").style.height = Math.ceil(Math.max(page.offsetHeight, H * PX) * k) + "px";
  }

  function S_size() { return ST.state.tyFontSize; }

  // ---------- Font (API) ----------
  // Case settings need PowerPoint API 1.8. Resolves { allCaps, smallCaps, strikethrough }, null when mixed or unknown.
  function readCase(context, range) {
    var none = { allCaps: null, smallCaps: null, strikethrough: null };
    if (!ST.supports("1.8")) return Promise.resolve(none);
    try { range.font.load("allCaps,smallCaps,strikethrough"); } catch (e) { return Promise.resolve(none); }
    return context.sync().then(function () { return { allCaps: range.font.allCaps, smallCaps: range.font.smallCaps, strikethrough: range.font.strikethrough }; }, function () { return none; });
  }
  // The text to change: the highlighted text when there is some, otherwise all the text in the selected shapes.
  function textTargets(context) {
    var pres = context.presentation, shapes = pres.getSelectedShapes();
    shapes.load("items/id,items/type,items/name");
    return context.sync().then(function () {
      var whole = { list: textItems(shapes.items).map(function (s) { return s.textFrame.textRange; }), part: false };
      if (!ST.supports("1.5")) return whole;
      var r;
      try { r = pres.getSelectedTextRangeOrNullObject ? pres.getSelectedTextRangeOrNullObject() : pres.getSelectedTextRange(); r.load("length"); }
      catch (e) { return whole; }
      return context.sync().then(function () { return !r.isNullObject && r.length > 0 ? { list: [r], part: true } : whole; }, function () { return whole; });
    });
  }
  // Each range's font size, using the first letter when the sizes are mixed
  function sizesOf(context, list) {
    list.forEach(function (r) { r.font.load("size"); });
    return context.sync().then(function () {
      var firsts = list.map(function (r) { if (r.font.size) return null; var f = r.getSubstring(0, 1); f.font.load("size"); return f; });
      return context.sync().then(function () { return list.map(function (r, i) { return r.font.size || (firsts[i] && firsts[i].font.size) || null; }); }, function () { return list.map(function (r) { return r.font.size || null; }); });
    });
  }
  function stepSize(size, dir) { return dir > 0 ? size + (size >= 20 ? 2 : 1) : Math.max(1, size - (size > 20 ? 2 : 1)); }
  function whereText(t) { return t.part ? "the highlighted text" : ST.plural(t.list.length, "shape"); }

  // kind: "name", "size" (text: a formula), "step" (dir), or a style key with value true/false or "Single"/"None"
  function setFont(kind, value) {
    if (!ST.needOffice()) return;
    if (!ST.supports("1.4")) { ST.toast("This version of PowerPoint can\'t change fonts from add-ins.", true); return; }
    var caseKey = ["allCaps", "smallCaps", "strikethrough"].indexOf(kind) >= 0;
    if (caseKey && !ST.supports("1.8")) { ST.toast("Caps and strikethrough need PowerPoint API 1.8. Update PowerPoint.", true); return; }
    if (kind === "name" && !String(value || "").trim()) { ST.toast("Type or pick a font first.", true); return; }
    if (kind === "size") {
      if (!String(value || "").trim()) { ST.toast("Type a size first, such as 18, 2g or +2.", true); return; }
      if (num(value, 12) === null) { ST.toast("Size " + (ST.ptProblem(value, 12) || "can\'t be read") + ". Use points, names such as g, or +2.", true); return; }
    }
    ST.run(function (context) {
      return textTargets(context).then(function (t) {
        if (!t.list.length) { ST.toast("Select text, or shapes with text.", true); return; }
        if (kind === "size" || kind === "step") {
          return sizesOf(context, t.list).then(function (sizes) {
            var bad = 0;
            t.list.forEach(function (r, i) {
              var cur = sizes[i] || 18, v = kind === "step" ? stepSize(cur, value) : num(value, cur);
              if (!(v >= 1 && v <= 4000)) { bad++; return; }
              r.font.size = Math.round(v * 10) / 10;
            });
            return context.sync().then(function () {
              if (bad) ST.toast("Sizes must work out to 1 to 4000 points, so " + ST.plural(bad, "shape") + " kept " + (bad === 1 ? "its" : "their") + " size.", true);
              else ST.toast(kind === "step" ? "Made the text " + (value > 0 ? "larger" : "smaller") + " in " + whereText(t) + "." : "Set the size of " + whereText(t) + ".");
            });
          });
        }
        var first = t.list[0];
        var current = (kind === "name") ? Promise.resolve(null) : caseKey ? readCase(context, first) : (first.font.load(kind), context.sync().then(function () { var o = {}; o[kind] = first.font[kind]; return o; }));
        return current.then(function (cur) {
          var v = value;
          if (v === undefined) v = kind === "underline" ? (cur.underline && cur.underline !== "None" ? "None" : "Single") : !(cur[kind] === true);
          t.list.forEach(function (r) { r.font[kind === "name" ? "name" : kind] = v; });
          ST.setRepeat(repeatLabel(kind, v), function () { setFont(kind, v); });
          return context.sync().then(function () { ST.toast(doneLabel(kind, v) + " for " + whereText(t) + "."); });
        });
      }).then(function () { queuePreview(); });
    });
    if (kind === "size") ST.setRepeat("Font size " + value, function () { setFont("size", value); });
    if (kind === "step") ST.setRepeat(value > 0 ? "Larger text" : "Smaller text", function () { setFont("step", value); });
  }
  var STYLE_NAMES = { bold: "Bold", italic: "Italic", underline: "Underline", allCaps: "All caps", smallCaps: "Small caps", strikethrough: "Strikethrough" };
  function repeatLabel(kind, v) { return kind === "name" ? "Font " + v : STYLE_NAMES[kind] + (v === false || v === "None" ? " off" : " on"); }
  function doneLabel(kind, v) { return kind === "name" ? "Set the font to " + v : (v === false || v === "None" ? "Turned " + STYLE_NAMES[kind].toLowerCase() + " off" : "Turned " + STYLE_NAMES[kind].toLowerCase() + " on"); }

  // Theme fonts from the slide export, then fonts used on the slide
  function fontsIn(doc, files) {
    var out = [], seen = {};
    function add(name, label) { if (!name || name.charAt(0) === "+" || seen[name]) return; seen[name] = true; out.push({ name: name, label: label }); }
    var theme = (files || []).filter(function (f) { return /^ppt\/theme\/theme\d+\.xml$/.test(f.name); })[0];
    if (theme) {
      var t = new DOMParser().parseFromString(theme.text, "application/xml");
      ["majorFont", "minorFont"].forEach(function (k, i) {
        var el = t.getElementsByTagNameNS(NS_A, k)[0], latin = el && kid(el, NS_A, "latin");
        if (latin) add(latin.getAttribute("typeface"), i ? "Body font (theme)" : "Heading font (theme)");
      });
    }
    Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS_A, "latin"), function (el) { add(el.getAttribute("typeface"), "On this slide"); });
    return out;
  }
  function renderFontList() {
    var dl = $("tyFontList");
    if (!dl) return;
    dl.innerHTML = "";
    (pv.fonts || []).forEach(function (f) { var o = document.createElement("option"); o.value = f.name; o.label = f.label; dl.appendChild(o); });
  }
  // Shows the selection's font on the buttons: lit when on, half lit when mixed
  function renderFontState() {
    var f = (pv.info && pv.info.font) || {}, has = !!pv.info;
    Array.prototype.forEach.call(document.querySelectorAll("#tyStyles [data-v], #tyCase [data-v]"), function (b) {
      var k = b.getAttribute("data-v"), v = f[k];
      var on = k === "underline" ? (v && v !== "None") : v === true;
      b.classList.toggle("on", has && !!on);
      b.classList.toggle("mixed", has && (v === null || v === undefined) && k !== "underline");
      b.setAttribute("aria-pressed", String(has && !!on));
    });
    $("tyFontSize").placeholder = has && f.size && !(pv.info.mixed || {}).size ? ST.nice(f.size) + " now" : "e.g. 18, 2g or +2";
    $("tyFontName").placeholder = has && f.name && !(pv.info.mixed || {}).name ? f.name + " now" : "Type or pick a font";
    var sw = $("tyColour").querySelector("span");
    sw.style.background = has && f.color ? f.color : "transparent";
    $("tyColour").classList.toggle("empty", !(has && f.color));
  }
  function openColours() {
    ST.state.colTarget = "text"; ST.save(); ST.emit("reset");
    if (ST.pinTool) ST.pinTool("typography");
    if (ST.openTool) ST.openTool("colours");
  }
  function markField(key) {
    var el = $(key), text = String(ST.state[key] || "").trim();
    if (!el) return;
    var bad = text && (isPtBox(key) ? num(text, key === "tyFontSize" ? 12 : undefined) === null : plain(text) === null);
    el.classList.toggle("err", !!bad);
    el.title = bad ? (isPtBox(key) ? (ST.ptProblem(text, key === "tyFontSize" ? 12 : undefined) || "can't be read") : "can't be read").replace(/^./, function (c) { return c.toUpperCase(); }) + "." : "";
  }

  // ---------- Panel ----------
  var MARKUP =
    '<div class="group-label label-row">Font' + ST.info("Applies straight away to the highlighted text, or to all the text in the selected shapes. Size takes points or formulas: <code>2g</code>, <code>g/2</code>, <code>+2</code> or <code>*1.5</code>. Point boxes below take <code>g</code>, <code>col</code> and <code>gap</code> too.", "About font") + '</div>' +
    '<label class="field">Font<input type="text" id="tyFontName" list="tyFontList" autocomplete="off" placeholder="Type or pick a font"><datalist id="tyFontList"></datalist></label>' +
    '<div class="row">' +
      '<label class="field">Size (pt)<div class="stepper"><input type="text" id="tyFontSize" inputmode="decimal" placeholder="e.g. 18, 2g or +2"><button id="tySizeLess" title="Smaller text">&minus;</button><button id="tySizeMore" title="Larger text">+</button></div></label>' +
      '<div class="field">Style<div class="ty-styles" id="tyStyles"><button data-v="bold" title="Bold" aria-label="Bold"><b>B</b></button><button data-v="italic" title="Italic" aria-label="Italic"><i>I</i></button><button data-v="underline" title="Underline" aria-label="Underline"><u>U</u></button><button id="tyColour" class="ty-swatch" title="Text colour: opens Colours" aria-label="Text colour"><span></span></button></div></div>' +
    '</div>' +
    '<div class="field" data-min="2">Case<div class="btn-row ty-case" id="tyCase"><button class="btn small" data-v="allCaps" title="All capitals">All caps</button><button class="btn small" data-v="smallCaps" title="Small capitals">Small caps</button><button class="btn small" data-v="strikethrough" title="Strikethrough">Strike</button></div></div>' +
    '<div class="group-label label-row">Spacing' + ST.info("Exact spacing for all the text in the selected shapes. Leave a box empty to keep what each shape has. PowerPoint add-ins can\'t set spacing directly, so the slide is swapped for an identical copy with the new spacing. Slide comments may not carry over, and Undo takes more than one step.", "About spacing") + '</div>' +
    '<div class="ty-preview" id="tyPreview" aria-live="polite"></div>' +
    '<label class="field">Character spacing (pt)<input type="text" id="tySpc" inputmode="decimal" placeholder="e.g. -0.5 or 1"></label>' +
    '<div class="row">' +
      '<label class="field">Line spacing<input type="text" id="tyLine" inputmode="decimal" placeholder="e.g. 18"></label>' +
      '<label class="field">Measured in<div class="seg seg2" id="tyLineUnit"><button data-v="pt" title="Line spacing in points">Points</button><button data-v="x" title="Line spacing as a multiple of single spacing">Lines</button></div></label>' +
    '</div>' +
    '<div class="row">' +
      '<label class="field">Space before (pt)<input type="text" id="tyBefore" inputmode="decimal" placeholder="e.g. 0"></label>' +
      '<label class="field">Space after (pt)<input type="text" id="tyAfter" inputmode="decimal" placeholder="e.g. 6"></label>' +
    '</div>' +
    '<div class="row" data-min="2">' +
      '<label class="field"><span class="label-row">Kerning from (pt)' + ST.info("Kerning tightens letter pairs such as AV in text at this size and larger. 0 turns kerning off.", "About kerning") + '</span><input type="text" id="tyKern" inputmode="decimal" placeholder="e.g. 12"></label>' +
      '<label class="field"><span class="label-row">Baseline (%)' + ST.info("Raises text above the baseline (positive) or lowers it (negative), as a percentage of the font size. 0 puts it back.", "About baseline") + '</span><input type="text" id="tyBase" inputmode="decimal" placeholder="e.g. 0"></label>' +
    '</div>' +
    '<button class="btn primary" id="tyApply">Apply spacing</button>' +
    '<div class="group-label">Text box</div>' +
    '<div class="sublabel">Margins (pt)</div>' +
    '<div class="row">' +
      '<label class="field">Left<input type="text" id="tyMl" inputmode="decimal" placeholder="7.2"></label>' +
      '<label class="field">Right<input type="text" id="tyMr" inputmode="decimal" placeholder="7.2"></label>' +
    '</div>' +
    '<div class="row">' +
      '<label class="field">Top<input type="text" id="tyMt" inputmode="decimal" placeholder="3.6"></label>' +
      '<label class="field">Bottom<input type="text" id="tyMb" inputmode="decimal" placeholder="3.6"></label>' +
    '</div>' +
    '<button class="btn" id="tyMargins">Apply margins</button>' +
    '<div class="sublabel">Autofit</div>' +
    '<div class="btn-row" id="tyFit"><button class="btn small" data-v="AutoSizeNone" title="Autofit off: text can overflow the shape">Off</button><button class="btn small" data-v="AutoSizeTextToFitShape" title="Shrink the text to fit the shape">Shrink</button><button class="btn small" data-v="AutoSizeShapeToFitText" title="Resize the shape to fit the text">Resize</button></div>' +
    '<div class="sublabel">Wrap text</div>' +
    '<div class="btn-row" id="tyWrap"><button class="btn small" data-v="1" title="Wrap text inside the shape">Wrap</button><button class="btn small" data-v="0" title="Keep each paragraph on one line">Don\'t wrap</button></div>' +
    '<button class="btn small" id="tyRead" data-min="2" title="Fill the boxes from the selected shape">Read from selection</button>' +
    '<div class="toggle" data-min="2"><span class="label-row">Live preview' + ST.info("Shows a scaled copy of the selected text box with the values in the boxes, before you apply them. Fonts that aren\'t installed on this computer show as a similar font.", "About the preview") + '</span><label class="switch"><input type="checkbox" id="tyPreviewOn"><span></span></label></div>';

  function renderUnit() { Array.prototype.forEach.call($("tyLineUnit").children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === (ST.state.tyLineUnit === "x" ? "x" : "pt")); }); }
  function build(body) {
    body.appendChild(ST.html(MARKUP));
    ["tyFontName", "tyFontSize", "tySpc", "tyLine", "tyBefore", "tyAfter", "tyKern", "tyBase", "tyMl", "tyMr", "tyMt", "tyMb"].forEach(function (k) { ST.bindField($(k), k, { onInput: function () { markField(k); } }); markField(k); });
    $("tyFontName").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); setFont("name", ST.state.tyFontName.trim()); } });
    // Picking from the list applies straight away
    $("tyFontName").addEventListener("input", function (e) { if (!e.inputType || e.inputType === "insertReplacementText") setFont("name", ST.state.tyFontName.trim()); });
    $("tyFontSize").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); setFont("size", ST.state.tyFontSize); } });
    $("tySizeLess").addEventListener("click", function (e) { e.preventDefault(); setFont("step", -1); });
    $("tySizeMore").addEventListener("click", function (e) { e.preventDefault(); setFont("step", 1); });
    Array.prototype.forEach.call(document.querySelectorAll("#tyStyles [data-v], #tyCase [data-v]"), function (b) {
      b.addEventListener("click", function () { setFont(b.getAttribute("data-v")); });
    });
    $("tyColour").addEventListener("click", openColours);
    ["tySpc", "tyLine", "tyBefore", "tyAfter", "tyKern", "tyBase"].forEach(function (k) {
      $(k).addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); applySpacing(); } });
    });
    ["tyMl", "tyMr", "tyMt", "tyMb"].forEach(function (k) {
      $(k).addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); applyMargins(); } });
    });
    Array.prototype.forEach.call($("tyLineUnit").children, function (b) {
      b.addEventListener("click", function () { ST.state.tyLineUnit = b.getAttribute("data-v"); ST.save(); renderUnit(); renderPreview(); });
    });
    var labels = { AutoSizeNone: "Turned autofit off", AutoSizeTextToFitShape: "Text now shrinks to fit", AutoSizeShapeToFitText: "Shapes now resize to fit the text" };
    Array.prototype.forEach.call($("tyFit").children, function (b) {
      b.addEventListener("click", function () { var v = b.getAttribute("data-v"); setFrame("autoSizeSetting", v, labels[v]); });
    });
    Array.prototype.forEach.call($("tyWrap").children, function (b) {
      b.addEventListener("click", function () { var on = b.getAttribute("data-v") === "1"; setFrame("wordWrap", on, on ? "Turned wrapping on" : "Turned wrapping off"); });
    });
    $("tyApply").addEventListener("click", function () { applySpacing(); });
    $("tyMargins").addEventListener("click", applyMargins);
    $("tyRead").addEventListener("click", readSelection);
    var pvOn = $("tyPreviewOn");
    pvOn.checked = previewOn();
    pvOn.addEventListener("change", function () { ST.state.tyPreviewOn = pvOn.checked; ST.save(); renderPreview(); if (pvOn.checked) queuePreview(); });
    renderUnit();
    renderPreview();
    ST.on("reset", function () { renderUnit(); pvOn.checked = previewOn(); renderPreview(); PT_KEYS.concat(["tyLine", "tyBase"]).forEach(markField); });
    ST.on("change", renderPreview);
    renderFontState();
    ST.on("selection", function () { queuePreview(); });
    ST.on("layout", function () { queuePreview(); });
    ST.on("ready", function () { queuePreview(); });
  }

  ST.registerTool({
    id: "typography",
    title: "Typography",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 1,
    defaults: defaults,
    tip: { what: "Sets fonts, sizes and styles, exact text spacing, margins and autofit.", steps: ["Select text, or shapes with text.", "Pick a font, size or style. These apply straight away.", "For spacing, fill in the boxes and click Apply spacing."] },
    build: build
  });

  ST._typography = { renderPreview: renderPreview, loadPreview: loadPreview, readSpec: readSpec, applyToBody: applyToBody, readBody: readBody, applySpacing: applySpacing, readSelection: readSelection,
                     setFont: setFont, fontsIn: fontsIn, num: num, stepSize: stepSize };
})();
