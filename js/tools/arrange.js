/* Position tool (id "arrange"): align, distribute with exact gaps, nudge, swap, and copy
   and paste position and size. Snap lives in the Grid tool (js/tools/snap.js). */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;

  var defaults = { arrGap: "g", arrAlignTo: "selection", arrNudge: "g", pasteWhat: "both", pasteAnchor: 0 };
  ST.addDefaults(defaults);

  var ANCHOR_NAMES = ["top left", "top", "top right", "left", "centre", "right", "bottom left", "bottom", "bottom right"];
  var copied = null; // { left, top, width, height } in points, for this session only

  var ICON = '<svg viewBox="0 0 18 18" fill="currentColor"><rect x="2" y="2" width="1.6" height="14" rx="0.8"/><rect x="5" y="4" width="10" height="4" rx="1.2"/><rect x="5" y="10" width="6" height="4" rx="1.2" opacity="0.6"/></svg>';

  function svg(paths) { return '<svg viewBox="0 0 16 16" fill="currentColor">' + paths + "</svg>"; }
  function hint(id) { return ST.shortcuts ? ST.shortcuts.hint(id) : ""; }
  var ALIGN = [
    ["left", "Align left", svg('<rect x="1.5" y="1" width="1.4" height="14" rx="0.7"/><rect x="4" y="3" width="9" height="3.5" rx="1"/><rect x="4" y="9.5" width="5.5" height="3.5" rx="1"/>')],
    ["centre", "Align centres", svg('<rect x="7.3" y="1" width="1.4" height="14" rx="0.7"/><rect x="3" y="3" width="10" height="3.5" rx="1" opacity="0.85"/><rect x="5" y="9.5" width="6" height="3.5" rx="1" opacity="0.85"/>')],
    ["right", "Align right", svg('<rect x="13.1" y="1" width="1.4" height="14" rx="0.7"/><rect x="3" y="3" width="9" height="3.5" rx="1"/><rect x="6.5" y="9.5" width="5.5" height="3.5" rx="1"/>')],
    ["top", "Align top", svg('<rect x="1" y="1.5" width="14" height="1.4" rx="0.7"/><rect x="3" y="4" width="3.5" height="9" rx="1"/><rect x="9.5" y="4" width="3.5" height="5.5" rx="1"/>')],
    ["middle", "Align middles", svg('<rect x="1" y="7.3" width="14" height="1.4" rx="0.7"/><rect x="3" y="3" width="3.5" height="10" rx="1" opacity="0.85"/><rect x="9.5" y="5" width="3.5" height="6" rx="1" opacity="0.85"/>')],
    ["bottom", "Align bottom", svg('<rect x="1" y="13.1" width="14" height="1.4" rx="0.7"/><rect x="3" y="3" width="3.5" height="9" rx="1"/><rect x="9.5" y="6.5" width="3.5" height="5.5" rx="1"/>')]
  ];

  var MARKUP =
    '<div class="group-label">Align</div>' +
    '<div class="icon-row" id="arrAlign"></div>' +
    '<div data-min="2"><div class="sublabel">Align to</div></div>' +
    '<div class="seg seg3" id="arrAlignTo" data-min="2"></div>' +
    '<div class="group-label label-row">Distribute' + ST.info("Across and Down place shapes one after another with exactly this gap. Even gaps keeps the outer shapes still and shares the space.", "About distribute") + '</div>' +
    '<div class="inline"><label class="field">Gap between shapes<input type="text" id="arrGap" placeholder="e.g. g"></label>' +
      '<div class="pair"><button class="btn small" id="arrDistH" title="Place shapes left to right with this gap">Across</button><button class="btn small" id="arrDistV" title="Place shapes top to bottom with this gap">Down</button></div></div>' +
    '<div class="calc" id="arrGapCalc" data-min="2"></div>' +
    '<div class="btn-row" data-min="2"><button class="btn small" id="arrSpreadH" title="Equal gaps, keeping the outer shapes where they are">Even gaps across</button><button class="btn small" id="arrSpreadV" title="Equal gaps, keeping the outer shapes where they are">Even gaps down</button></div>' +
    '<div class="group-label">Move</div>' +
    '<div class="nudge-row">' +
      '<label class="field">Nudge by<input type="text" id="arrNudge" placeholder="e.g. g"></label>' +
      '<div class="nudge-pad" id="arrNudgePad"><button class="btn small" data-nudge="-1,0" title="Nudge left' + hint("NudgeLeft") + '">&larr;</button><button class="btn small" data-nudge="0,-1" title="Nudge up' + hint("NudgeUp") + '">&uarr;</button><button class="btn small" data-nudge="0,1" title="Nudge down' + hint("NudgeDown") + '">&darr;</button><button class="btn small" data-nudge="1,0" title="Nudge right' + hint("NudgeRight") + '">&rarr;</button></div>' +
    '</div>' +
    '<button class="btn small" id="arrSwap" data-min="3" title="Swap the positions of two selected shapes">Swap two shapes</button>' +
    '<div class="group-label">Copy and paste position</div>' +
    '<div class="btn-row"><button class="btn small" id="posCopy" title="Remember the first selected shape\'s position and size">Copy</button><button class="btn small" id="posPaste" title="Apply the copied position and size to the selected shapes" disabled>Paste</button></div>' +
    '<div class="seg seg3" id="pasteWhat"></div>' +
    '<div class="anchor-row" data-min="2"><span class="label-row">Paste from&nbsp;<b id="pasteAnchorName"></b>' + ST.info("The point that lines up with the copied shape when sizes differ.", "About paste from") + '</span><div class="anchor" id="pasteAnchor"></div></div>' +
    '<div class="help" id="posCopied">Nothing copied yet</div>';

  // ---------- Helpers ----------
  function box(items) {
    var l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
    items.forEach(function (s) { l = Math.min(l, s.left); t = Math.min(t, s.top); r = Math.max(r, s.left + s.width); b = Math.max(b, s.top + s.height); });
    return { left: l, top: t, width: r - l, height: b - t };
  }
  function areaBox() {
    var g = ST.calculate();
    return { left: g.left * PT, top: g.top * PT, width: g.maxW * PT, height: g.maxH * PT };
  }
  function amount(key) {
    ST.calculate();
    var v = ST.evaluate(ST.state[key], ST.formulaVars(), null);
    return v === null ? null : v * PT;
  }

  // ---------- Align ----------
  function align(edge) {
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      var mode = ST.state.arrAlignTo, ref, movers = items, note = "";
      if (mode === "area" || (mode === "selection" && items.length === 1)) {
        ref = areaBox();
        if (mode === "selection") note = " One shape, so it was aligned to the grid area.";
      } else if (mode === "key") {
        if (items.length < 2) { ST.toast("Select at least two shapes to align to the largest one.", true); return; }
        var key = items.reduce(function (a, s) { return s.width * s.height > a.width * a.height ? s : a; });
        ref = { left: key.left, top: key.top, width: key.width, height: key.height };
        movers = items.filter(function (s) { return s !== key; });
      } else ref = box(items);
      movers.forEach(function (s) {
        if (edge === "left") s.left = ref.left;
        if (edge === "centre") s.left = ref.left + (ref.width - s.width) / 2;
        if (edge === "right") s.left = ref.left + ref.width - s.width;
        if (edge === "top") s.top = ref.top;
        if (edge === "middle") s.top = ref.top + (ref.height - s.height) / 2;
        if (edge === "bottom") s.top = ref.top + ref.height - s.height;
      });
      return context.sync().then(function () { ST.toast("Aligned " + ST.plural(movers.length, "shape") + "." + note); ST.refreshSelection(); });
    });
  }

  // ---------- Distribute ----------
  function distribute(horizontal) {
    var gap = amount("arrGap");
    if (gap === null) { ST.toast("Check the gap. Use a number or a formula like g or 0.5g.", true); return; }
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      if (items.length < 2) { ST.toast("Select at least two shapes.", true); return; }
      var pos = horizontal ? "left" : "top", size = horizontal ? "width" : "height";
      items.sort(function (a, b) { return a[pos] - b[pos]; });
      var at = items[0][pos] + items[0][size];
      for (var i = 1; i < items.length; i++) { items[i][pos] = at + gap; at = items[i][pos] + items[i][size]; }
      return context.sync().then(function () { ST.toast("Placed " + ST.plural(items.length, "shape") + " " + (horizontal ? "across" : "down") + " with " + ST.nice(gap / PT) + " in gaps."); ST.refreshSelection(); });
    });
  }
  function spread(horizontal) {
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      if (items.length < 3) { ST.toast("Select at least three shapes to spread evenly.", true); return; }
      var pos = horizontal ? "left" : "top", size = horizontal ? "width" : "height";
      items.sort(function (a, b) { return a[pos] - b[pos]; });
      var first = items[0], last = items[items.length - 1];
      var start = first[pos], end = Math.max.apply(null, items.map(function (s) { return s[pos] + s[size]; }));
      var total = items.reduce(function (sum, s) { return sum + s[size]; }, 0);
      var gap = (end - start - total) / (items.length - 1);
      var at = start;
      items.forEach(function (s) { s[pos] = at; at += s[size] + gap; });
      return context.sync().then(function () {
        ST.toast("Spread " + ST.plural(items.length, "shape") + " with " + ST.nice(gap / PT) + " in gaps." + (gap < 0 ? " They overlap, the shapes don't fit." : ""));
        ST.refreshSelection();
      });
    });
  }

  // ---------- Move ----------
  function nudge(dx, dy) {
    var d = amount("arrNudge");
    if (d === null) { ST.toast("Check the nudge amount.", true); return; }
    ST.withSelection(["left", "top"], function (context, items) {
      items.forEach(function (s) { s.left += dx * d; s.top += dy * d; });
      return context.sync().then(function () { ST.refreshSelection(); });
    });
  }
  // Changes width and height by the nudge amount, keeping Size's reference point still
  function resizeBy(dw, dh) {
    var d = amount("arrNudge");
    if (d === null) { ST.toast("Check the nudge amount.", true); return; }
    var a = ST.state.anchor || 0, fx = (a % 3) / 2, fy = Math.floor(a / 3) / 2;
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      var skipped = 0;
      items.forEach(function (s) {
        var w = s.width + dw * d, h = s.height + dh * d;
        // A line keeps its zero side, and nothing shrinks to nothing
        if ((dw && s.width > 0 && w < 1) || (dh && s.height > 0 && h < 1)) { skipped++; return; }
        if (dw && s.width > 0) { s.left -= (w - s.width) * fx; s.width = w; }
        if (dh && s.height > 0) { s.top -= (h - s.height) * fy; s.height = h; }
      });
      return context.sync().then(function () {
        if (skipped) ST.toast(ST.plural(skipped, "shape") + " can't get any smaller.", true);
        ST.refreshSelection();
      });
    });
  }
  function swap() {
    ST.withSelection(["left", "top"], function (context, items) {
      if (items.length !== 2) { ST.toast("Select exactly two shapes to swap.", true); return; }
      var a = items[0], b = items[1], l = a.left, t = a.top;
      a.left = b.left; a.top = b.top; b.left = l; b.top = t;
      return context.sync().then(function () { ST.toast("Swapped the two shapes."); ST.refreshSelection(); });
    });
  }

  // ---------- Copy and paste position and size ----------
  function inches(pt) { return ST.nice(Math.round(pt / PT * 100) / 100); }
  function showCopied() {
    $("posPaste").disabled = !copied;
    $("posCopied").textContent = copied
      ? "Copied: left " + inches(copied.left) + ", top " + inches(copied.top) + " in. " + inches(copied.width) + " x " + inches(copied.height) + " in."
      : "Nothing copied yet";
  }
  function copy() {
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      var s = items[0];
      copied = { left: s.left, top: s.top, width: s.width, height: s.height };
      showCopied();
      ST.toast(items.length > 1 ? "Copied the first selected shape." : "Copied position and size.");
    }, "Select the shape to copy first.");
  }
  // Works out where a shape goes. The chosen anchor point of the shape lands on the same
  // point of the copied box (Position), or stays where it is (Size only).
  function placement(s, box, what, anchor) {
    var fx = (anchor % 3) / 2, fy = Math.floor(anchor / 3) / 2;
    var size = what !== "position", pos = what !== "size";
    var w = size ? box.width : s.width, h = size ? box.height : s.height;
    var ax = pos ? box.left + box.width * fx : s.left + s.width * fx;
    var ay = pos ? box.top + box.height * fy : s.top + s.height * fy;
    return { left: ax - w * fx, top: ay - h * fy, width: w, height: h };
  }
  function paste() {
    if (!copied) { ST.toast("Copy a shape first.", true); return; }
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      var what = ST.state.pasteWhat || "both", anchor = ST.state.pasteAnchor || 0;
      items.forEach(function (s) {
        var p = placement(s, copied, what, anchor);
        if (what !== "position") { s.width = p.width; s.height = p.height; }
        s.left = p.left; s.top = p.top;
      });
      return context.sync().then(function () {
        var word = what === "both" ? "position and size" : what;
        ST.toast("Pasted " + word + " onto " + ST.plural(items.length, "shape") + ".");
        ST.refreshSelection();
      });
    }, "Select the shapes to paste onto.");
  }

  function renderPaste() {
    var seg = $("pasteWhat");
    if (!seg.children.length) {
      [["position", "Position"], ["size", "Size"], ["both", "Both"]].forEach(function (o) {
        var b = document.createElement("button");
        b.textContent = o[1];
        b.setAttribute("data-v", o[0]);
        b.addEventListener("click", function () { ST.state.pasteWhat = o[0]; ST.changed(); });
        seg.appendChild(b);
      });
    }
    Array.prototype.forEach.call(seg.children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === (ST.state.pasteWhat || "both")); });
    var el = $("pasteAnchor");
    if (!el.children.length) {
      ANCHOR_NAMES.forEach(function (name, i) {
        var b = document.createElement("button");
        b.title = "Line up the " + name + " corner";
        if (i === 1 || i === 3 || i === 5 || i === 7) b.title = "Line up the " + name + " edge";
        if (i === 4) b.title = "Line up the centre";
        b.setAttribute("aria-label", b.title);
        b.addEventListener("click", function () { ST.state.pasteAnchor = i; ST.changed(); });
        el.appendChild(b);
      });
    }
    var a = ST.state.pasteAnchor || 0;
    Array.prototype.forEach.call(el.children, function (b, i) { b.classList.toggle("on", i === a); });
    $("pasteAnchorName").textContent = ANCHOR_NAMES[a];
  }
  // ---------- Build ----------
  function seg(id, options, key) {
    var el = $(id);
    options.forEach(function (o) {
      var b = document.createElement("button");
      b.textContent = o[1];
      b.addEventListener("click", function () {
        ST.state[key] = o[0]; ST.save(); renderSegs();
      });
      b.setAttribute("data-v", o[0]);
      el.appendChild(b);
    });
  }
  function renderSegs() {
    Array.prototype.forEach.call($("arrAlignTo").children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === ST.state.arrAlignTo); });
  }
  function updateCalc() {
    var el = $("arrGapCalc");
    if (!el) return;
    var v = amount("arrGap");
    el.textContent = String(ST.state.arrGap || "").trim() ? (v === null ? "Gap: check the formula" : "Gap = " + ST.exact(v / PT) + " in") : "";
    el.className = v === null ? "calc bad" : "calc";
  }

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    var row = $("arrAlign");
    ALIGN.forEach(function (a) {
      var b = document.createElement("button");
      b.className = "tool-btn"; b.title = a[1]; b.setAttribute("aria-label", a[1]); b.innerHTML = a[2];
      b.addEventListener("click", function () { ST.setRepeat(a[1], function () { align(a[0]); }); align(a[0]); });
      row.appendChild(b);
    });
    seg("arrAlignTo", [["selection", "Selection"], ["area", "Grid area"], ["key", "Largest"]], "arrAlignTo");
    ST.bindField($("arrGap"), "arrGap");
    ST.bindField($("arrNudge"), "arrNudge");
    $("arrDistH").addEventListener("click", function () { ST.setRepeat("Distribute across", function () { distribute(true); }); distribute(true); });
    $("arrDistV").addEventListener("click", function () { ST.setRepeat("Distribute down", function () { distribute(false); }); distribute(false); });
    $("arrGap").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); distribute(true); } });
    $("arrSpreadH").addEventListener("click", function () { ST.setRepeat("Even gaps across", function () { spread(true); }); spread(true); });
    $("arrSpreadV").addEventListener("click", function () { ST.setRepeat("Even gaps down", function () { spread(false); }); spread(false); });
    $("arrSwap").addEventListener("click", swap);
    Array.prototype.forEach.call(body.querySelectorAll("[data-nudge]"), function (b) {
      var d = b.getAttribute("data-nudge").split(",").map(Number);
      b.addEventListener("click", function () { nudge(d[0], d[1]); });
    });
    renderSegs();
    $("posCopy").addEventListener("click", copy);
    $("posPaste").addEventListener("click", paste);
    ST.on("change", renderPaste);
    ST.on("reset", renderPaste);
    renderPaste(); showCopied();
    ST.on("change", updateCalc);
    ST.on("reset", renderSegs);
  }

  // Used by keyboard shortcuts (js/shortcuts.js)
  ST._arrange = {
    nudge: nudge,
    resizeBy: resizeBy,
    placement: placement
  };

  ST.registerTool({
    id: "arrange",
    title: "Position",
    icon: ICON,
    levels: [1, 2, 3],
    defaultLevel: 1,
    defaults: defaults,
    tip: { what: "Lines up, spaces, nudges and copies the position of shapes.", steps: ["Select shapes.", "Click an align button, Across or Down.", "Or copy one shape's position and paste it onto others."] },
    build: build
  });
})();
