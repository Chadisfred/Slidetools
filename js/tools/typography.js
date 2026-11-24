/* Typography tool: font, size, style and case, plus text box margins, wrap and autofit.
   Everything here uses the PowerPoint API, so each change shows on the slide straight away.
   Point boxes take formulas in points: 2g, g/2, 12pt+2, and in Size, +2 or *1.5. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;
  var R = ST.rebuild, NS_A = R.NS_A, kid = R.kid;
  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 14L7 4l4 10M4.5 10.5h5"/><path d="M13 7v7M11.5 14h3"/></svg>';
  var TEXT_TYPES = ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"];

  var defaults = { tyFontName: "", tyFontSize: "", tyMl: "", tyMr: "", tyMt: "", tyMb: "" };
  ST.addDefaults(defaults);
  var PT_KEYS = ["tyFontSize", "tyMl", "tyMr", "tyMt", "tyMb"];

  // A number in points from a box, "" when empty, null when it can't be read.
  // Points boxes take grid names (g, col, gap) in points. own: the current value, for +2 or *1.5.
  function num(text, own) {
    var t = String(text || "").trim();
    if (!t) return "";
    var v = ST.evaluatePt(t, own);
    return v === null || !isFinite(v) ? null : v;
  }

  function textItems(items) { return items.filter(function (s) { return TEXT_TYPES.indexOf(s.type) >= 0; }); }

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
      return context.sync().then(function () { ST.toast("Set the margins of " + ST.plural(list.length, "shape") + "."); queueState(); });
    });
  }
  function setFrame(prop, value, label) {
    if (!ST.needOffice()) return;
    ST.withSelection(["type"], function (context, items) {
      var list = textItems(items);
      if (!list.length) { ST.toast("Select shapes or text boxes.", true); return; }
      list.forEach(function (s) { s.textFrame[prop] = value; });
      return context.sync().then(function () { ST.toast(label + " for " + ST.plural(list.length, "shape") + "."); queueState(); });
    });
  }


  // ---------- Read from selection ----------
  function readSelection() {
    if (!ST.needOffice()) return;
    ST.withSelection(["id", "name", "type"], function (context, items) {
      var s = textItems(items)[0];
      if (!s) { ST.toast("Select a shape or text box with text.", true); return; }
      var tf = s.textFrame;
      tf.load("leftMargin,rightMargin,topMargin,bottomMargin");
      return context.sync().then(function () {
        var nice = function (v) { return v === null || v === undefined ? "" : ST.nice(v); };
        ST.state.tyMl = nice(tf.leftMargin); ST.state.tyMr = nice(tf.rightMargin); ST.state.tyMt = nice(tf.topMargin); ST.state.tyMb = nice(tf.bottomMargin);
        ST.save(); ST.emit("reset"); ST.changed();
        ST.toast("Read the margins from " + (s.name || "the selection") + ".");
      });
    });
  }

  // ---------- Current font ----------
  // The selection's font lights the style buttons and fills the size and font hints. The font list
  // adds the theme fonts and the fonts on the slide, read from an export that never changes the slide.
  var st = { info: null, fonts: [], fontSlide: null, timer: null, busy: false, again: false };
  function stateVisible() { return !!$("tyStyles") && (ST.isOpen ? ST.isOpen("typography") : true); }
  function queueState() {
    clearTimeout(st.timer);
    st.timer = setTimeout(loadState, 200);
  }
  function loadState() {
    if (!ST.inOffice || !stateVisible() || !ST.supports("1.4") || document.hidden) { renderFontState(); return; }
    if (st.busy) { st.again = true; return; }
    st.busy = true;
    PowerPoint.run(function (context) {
      var shapes = context.presentation.getSelectedShapes(), slides = context.presentation.getSelectedSlides();
      shapes.load("items/id,items/name,items/type"); slides.load("items/id");
      return context.sync().then(function () {
        var slideId = slides.items[0] && slides.items[0].id, s = textItems(shapes.items)[0], info = null;
        var fonts = slideId && slideId !== st.fontSlide && R.supported() ? R.edit(context, function (doc, extra) {
          st.fonts = fontsIn(doc, extra && extra.files); st.fontSlide = slideId; renderFontList();
          return { count: 0 };
        }).catch(function () {}) : Promise.resolve();
        if (!s) return fonts.then(function () { return null; });
        var tr = s.textFrame.textRange;
        tr.font.load("name,size,color,bold,italic,underline");
        return context.sync().then(function () {
          info = { font: { name: tr.font.name, size: tr.font.size, color: tr.font.color, bold: tr.font.bold, italic: tr.font.italic, underline: tr.font.underline } };
          info.mixed = { name: !tr.font.name, size: !tr.font.size, color: !tr.font.color };
          return readCase(context, tr).then(function (c) { info.font.allCaps = c.allCaps; info.font.smallCaps = c.smallCaps; info.font.strikethrough = c.strikethrough; });
        }).then(function () { return fonts; }).then(function () { return info; });
      });
    }).then(function (info) { st.info = info; }, function () { st.info = null; })
      .then(function () {
        st.busy = false; renderFontState();
        if (st.again) { st.again = false; queueState(); }
      });
  }

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
      }).then(function () { queueState(); });
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
    (st.fonts || []).forEach(function (f) { var o = document.createElement("option"); o.value = f.name; o.label = f.label; dl.appendChild(o); });
  }
  // Shows the selection's font on the buttons: lit when on, half lit when mixed
  function renderFontState() {
    var f = (st.info && st.info.font) || {}, has = !!st.info;
    Array.prototype.forEach.call(document.querySelectorAll("#tyStyles [data-v], #tyCase [data-v]"), function (b) {
      var k = b.getAttribute("data-v"), v = f[k];
      var on = k === "underline" ? (v && v !== "None") : v === true;
      b.classList.toggle("on", has && !!on);
      // null means mixed. Undefined means this PowerPoint can't say, so nothing is lit.
      b.classList.toggle("mixed", has && v === null && k !== "underline");
      b.setAttribute("aria-pressed", String(has && !!on));
    });
    $("tyFontSize").placeholder = has && f.size && !(st.info.mixed || {}).size ? ST.nice(f.size) + " now" : "e.g. 18, 2g or +2";
    $("tyFontName").placeholder = has && f.name && !(st.info.mixed || {}).name ? f.name + " now" : "Type or pick a font";
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
    if (!el || key === "tyFontName") return; // a font name is free text
    var own = key === "tyFontSize" ? 12 : undefined, bad = text && num(text, own) === null;
    el.classList.toggle("err", !!bad);
    el.title = bad ? (ST.ptProblem(text, own) || "can't be read").replace(/^./, function (c) { return c.toUpperCase(); }) + "." : "";
  }


  // ---------- Panel ----------
  var MARKUP =
    '<div class="group-label label-row">Font' + ST.info("Applies straight away to the highlighted text, or to all the text in the selected shapes. Size takes points or formulas: <code>2g</code>, <code>g/2</code>, <code>+2</code> or <code>*1.5</code>. Margin boxes take <code>g</code>, <code>col</code> and <code>gap</code> too.", "About font") + '</div>' +
    '<label class="field">Font<input type="text" id="tyFontName" list="tyFontList" autocomplete="off" placeholder="Type or pick a font"><datalist id="tyFontList"></datalist></label>' +
    '<div class="row">' +
      '<label class="field">Size (pt)<div class="stepper"><input type="text" id="tyFontSize" inputmode="decimal" placeholder="e.g. 18, 2g or +2"><button id="tySizeLess" title="Smaller text">&minus;</button><button id="tySizeMore" title="Larger text">+</button></div></label>' +
      '<div class="field">Style<div class="ty-styles" id="tyStyles"><button data-v="bold" title="Bold" aria-label="Bold"><b>B</b></button><button data-v="italic" title="Italic" aria-label="Italic"><i>I</i></button><button data-v="underline" title="Underline" aria-label="Underline"><u>U</u></button><button id="tyColour" class="ty-swatch" title="Text colour: opens Colours" aria-label="Text colour"><span></span></button></div></div>' +
    '</div>' +
    '<div class="field" data-min="2">Case<div class="btn-row ty-case" id="tyCase"><button class="btn small" data-v="allCaps" title="All capitals">All caps</button><button class="btn small" data-v="smallCaps" title="Small capitals">Small caps</button><button class="btn small" data-v="strikethrough" title="Strikethrough">Strike</button></div></div>' +
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
    '<button class="btn small" id="tyRead" data-min="2" title="Fill the margin boxes from the selected shape">Read from selection</button>' +
    '<div class="help" data-min="2">Line and character spacing aren\'t available to add-ins. Use Paragraph and Font in PowerPoint\'s Home tab.</div>';

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    ["tyFontName", "tyFontSize", "tyMl", "tyMr", "tyMt", "tyMb"].forEach(function (k) { ST.bindField($(k), k, { onInput: function () { markField(k); } }); markField(k); });
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
    ["tyMl", "tyMr", "tyMt", "tyMb"].forEach(function (k) {
      $(k).addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); applyMargins(); } });
    });
    var labels = { AutoSizeNone: "Turned autofit off", AutoSizeTextToFitShape: "Text now shrinks to fit", AutoSizeShapeToFitText: "Shapes now resize to fit the text" };
    Array.prototype.forEach.call($("tyFit").children, function (b) {
      b.addEventListener("click", function () { var v = b.getAttribute("data-v"); setFrame("autoSizeSetting", v, labels[v]); });
    });
    Array.prototype.forEach.call($("tyWrap").children, function (b) {
      b.addEventListener("click", function () { var on = b.getAttribute("data-v") === "1"; setFrame("wordWrap", on, on ? "Turned wrapping on" : "Turned wrapping off"); });
    });
    $("tyMargins").addEventListener("click", applyMargins);
    $("tyRead").addEventListener("click", readSelection);
    ST.on("reset", function () { PT_KEYS.forEach(markField); });
    renderFontState();
    ST.on("selection", queueState);
    ST.on("layout", queueState);
    ST.on("ready", queueState);
  }

  ST.registerTool({
    id: "typography",
    title: "Typography",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 1,
    defaults: defaults,
    tip: { what: "Sets fonts, sizes, styles, case, margins, autofit and wrapping.", steps: ["Select text, or shapes with text.", "Pick a font, size or style. Changes show on the slide straight away.", "Fill in margins and click Apply margins."] },
    build: build
  });

  ST._typography = { readSelection: readSelection, loadState: loadState, setFont: setFont, fontsIn: fontsIn, num: num, stepSize: stepSize };
})();
