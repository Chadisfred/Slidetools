/* Typography tool: exact character spacing, line spacing, paragraph spacing, kerning and
   baseline, plus text box margins, wrap and autofit. PowerPoint's add-in API has no
   spacing, kerning or baseline settings, so those are written to the slide XML through the
   shared slide rebuild (js/rebuild.js), in memory. Margins, wrap and autofit use the API. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;
  var R = ST.rebuild, NS_P = R.NS_P, NS_A = R.NS_A, kid = R.kid, kids = R.kids, make = R.make;
  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 14L7 4l4 10M4.5 10.5h5"/><path d="M13 7v7M11.5 14h3"/></svg>';
  var TEXT_TYPES = ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"];
  var PPR_ORDER = ["lnSpc", "spcBef", "spcAft"];

  var defaults = { tySpc: "", tyLine: "", tyLineUnit: "pt", tyBefore: "", tyAfter: "", tyKern: "", tyBase: "",
                   tyMl: "", tyMr: "", tyMt: "", tyMb: "", tyPreviewOn: true };
  ST.addDefaults(defaults);

  // A number in points from a box, "" when empty, null when it can't be read
  function num(text) {
    var t = String(text || "").trim().replace(/\s*pt$/i, "").replace(/\s*%$/, "");
    if (!t) return "";
    var v = ST.evaluate(t, null, null);
    return v === null || !isFinite(v) ? null : v;
  }

  // ---------- Reading the panel (pure) ----------
  // Returns { spec } or { bad }. spec values are undefined when the box is empty (leave as is).
  function readSpec(S) {
    var out = {}, bad = null;
    function take(key, label, test, msg) {
      var v = num(S[key]);
      if (v === "") return undefined;
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
    if (keys.some(function (k) { return m[k] === null || m[k] < 0; })) { ST.toast("Margins must be 0 or more points.", true); return; }
    if (!keys.length) { ST.toast("Fill in at least one margin first.", true); return; }
    if (!ST.needOffice()) return;
    ST.setRepeat("Apply margins", applyMargins);
    ST.withSelection(["type"], function (context, items) {
      var list = textItems(items);
      if (!list.length) { ST.toast("Select shapes or text boxes.", true); return; }
      list.forEach(function (s) { keys.forEach(function (k) { s.textFrame[k] = m[k]; }); });
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
        tr.load("text"); tr.font.load("name,size,color,bold,italic"); tr.paragraphFormat.load("horizontalAlignment");
        s.fill.load("type,foregroundColor");
        return context.sync().then(function () {
          info.text = String(tr.text || "").slice(0, 600);
          info.margins = { l: tf.leftMargin, r: tf.rightMargin, t: tf.topMargin, b: tf.bottomMargin };
          info.wrap = tf.wordWrap !== false; info.valign = tf.verticalAlignment || "Top";
          info.align = tr.paragraphFormat.horizontalAlignment || "Left";
          info.fill = s.fill.type === "Solid" ? s.fill.foregroundColor : null;
          info.font = { name: tr.font.name, size: tr.font.size, color: tr.font.color, bold: tr.font.bold, italic: tr.font.italic };
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
            return { count: 0 };
          }).then(function () { info.cur = pv.cache[s.id]; return info; }, function () { info.cur = null; return info; });
        });
      });
    }).then(function (info) { pv.info = info; }, function () { pv.info = null; })
      .then(function () {
        pv.busy = false; renderPreview();
        if (pv.again) { pv.again = false; queuePreview(); }
      });
  }
  // The values the preview shows: a filled box wins, then the shape's own setting, then the default
  function effective(info) {
    var cur = info.cur || {}, S = ST.state, pick = function (key, own, dflt) { var v = num(S[key]); return typeof v === "number" ? v : (own !== null && own !== undefined ? own : dflt); };
    var lineBox = num(S.tyLine), line = typeof lineBox === "number" && lineBox > 0 ? { unit: S.tyLineUnit === "x" ? "x" : "pt", value: lineBox } : (cur.line || { unit: "x", value: 1 });
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
    var e = effective(info), f = info.font || {}, size = f.size || 18, W = Math.max(info.box.width, 1), H = Math.max(info.box.height, 1);
    var lh = e.line.unit === "pt" ? e.line.value + "pt" : String(e.line.value * 1.2);
    var align = { Center: "center", Centered: "center", Right: "right", Justify: "justify", JustifyLow: "justify", Distributed: "justify" }[info.align] || "left";
    var justify = /Middle|Center/.test(info.valign) ? "center" : /Bottom/.test(info.valign) ? "flex-end" : "flex-start";
    var paras = (info.text || "").split(/\r\n|\r|\n/).map(function (p) {
      var inner = p ? p.split("\v").map(esc).join("<br>") : "&nbsp;";
      return '<p style="margin:' + e.before + 'pt 0 ' + e.after + 'pt;line-height:' + lh + '"><span style="position:relative;top:' + (-e.base / 100) + 'em">' + inner + '</span></p>';
    }).join("");
    var avail = Math.max(120, box.clientWidth - 2), k = Math.min(1, avail / (W * PX));
    box.innerHTML = '<div class="ty-stage"><div class="ty-page" style="width:' + W + 'pt;min-height:' + H + 'pt;padding:' + e.mt + 'pt ' + e.mr + 'pt ' + e.mb + 'pt ' + e.ml + 'pt;' +
      'font-family:&quot;' + esc(f.name || "Calibri") + '&quot;, sans-serif;font-size:' + size + 'pt;color:' + esc(f.color || "#1d1d1f") + ';font-weight:' + (f.bold ? 700 : 400) + ';font-style:' + (f.italic ? "italic" : "normal") + ';' +
      'letter-spacing:' + e.spc + 'pt;text-align:' + align + ';justify-content:' + justify + ';white-space:' + (info.wrap ? "pre-wrap" : "pre") + ';background:' + esc(info.fill || "#ffffff") + ';transform:scale(' + k + ')">' +
      paras + '<i class="ty-bounds" style="height:' + H + 'pt"></i></div></div><div class="ty-caption">' + esc(info.name || "Selected shape") + ', ' + Math.round(k * 100) + '% size</div>';
    var page = box.querySelector(".ty-page");
    box.querySelector(".ty-stage").style.height = Math.ceil(Math.max(page.offsetHeight, H * PX) * k) + "px";
  }

  // ---------- Panel ----------
  var MARKUP =
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
    ["tySpc", "tyLine", "tyBefore", "tyAfter", "tyKern", "tyBase", "tyMl", "tyMr", "tyMt", "tyMb"].forEach(function (k) { ST.bindField($(k), k); });
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
    ST.on("reset", function () { renderUnit(); pvOn.checked = previewOn(); renderPreview(); });
    ST.on("change", renderPreview);
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
    tip: { what: "Sets exact text spacing, margins and autofit.", steps: ["Select shapes with text.", "Fill in the spacing you want. Empty boxes are left alone.", "Click Apply spacing."] },
    build: build
  });

  ST._typography = { renderPreview: renderPreview, loadPreview: loadPreview, readSpec: readSpec, applyToBody: applyToBody, readBody: readBody, applySpacing: applySpacing, readSelection: readSelection };
})();
