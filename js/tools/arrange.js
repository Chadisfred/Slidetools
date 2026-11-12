/* Arrange tool: align, distribute with exact gaps, snap to the Layout grid, swap and nudge. */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;

  var defaults = { arrGap: "g", arrAlignTo: "selection", arrSnap: "cells", arrSnapEdges: false, arrSnapResize: false, arrNudge: "g" };
  ST.addDefaults(defaults);

  var ICON = '<svg viewBox="0 0 18 18" fill="currentColor"><rect x="2" y="2" width="1.6" height="14" rx="0.8"/><rect x="5" y="4" width="10" height="4" rx="1.2"/><rect x="5" y="10" width="6" height="4" rx="1.2" opacity="0.6"/></svg>';

  function svg(paths) { return '<svg viewBox="0 0 16 16" fill="currentColor">' + paths + "</svg>"; }
  var ALIGN = [
    ["left", "Align left", svg('<rect x="1.5" y="1" width="1.4" height="14" rx="0.7"/><rect x="4" y="3" width="9" height="3.5" rx="1"/><rect x="4" y="9.5" width="5.5" height="3.5" rx="1"/>')],
    ["centre", "Align centres", svg('<rect x="7.3" y="1" width="1.4" height="14" rx="0.7"/><rect x="3" y="3" width="10" height="3.5" rx="1" opacity="0.85"/><rect x="5" y="9.5" width="6" height="3.5" rx="1" opacity="0.85"/>')],
    ["right", "Align right", svg('<rect x="13.1" y="1" width="1.4" height="14" rx="0.7"/><rect x="3" y="3" width="9" height="3.5" rx="1"/><rect x="6.5" y="9.5" width="5.5" height="3.5" rx="1"/>')],
    ["top", "Align top", svg('<rect x="1" y="1.5" width="14" height="1.4" rx="0.7"/><rect x="3" y="4" width="3.5" height="9" rx="1"/><rect x="9.5" y="4" width="3.5" height="5.5" rx="1"/>')],
    ["middle", "Align middles", svg('<rect x="1" y="7.3" width="14" height="1.4" rx="0.7"/><rect x="3" y="3" width="3.5" height="10" rx="1" opacity="0.85"/><rect x="9.5" y="5" width="3.5" height="6" rx="1" opacity="0.85"/>')],
    ["bottom", "Align bottom", svg('<rect x="1" y="13.1" width="14" height="1.4" rx="0.7"/><rect x="3" y="3" width="3.5" height="9" rx="1"/><rect x="9.5" y="6.5" width="3.5" height="5.5" rx="1"/>')]
  ];

  var MARKUP =
    '<span class="info res-info" tabindex="0" role="button" aria-label="How Arrange works">i' +
      '<span class="tip" role="tooltip">Align lines shapes up with the selection, the Layout grid area, or the largest selected shape.<br><br>' +
      'Distribute places shapes one after another with an exact gap. The gap box takes numbers or formulas such as <code>g</code>, <code>0.5g</code> or <code>gap</code>.<br><br>' +
      'Snap to grid moves shapes onto the Layout grid. Grid cells snaps to the nearest cell corner and can resize shapes to the cells they cover. Fine grid snaps to steps of <code>g</code> measured from the grid area\'s top left.</span>' +
    '</span>' +
    '<div class="group-label">Align</div>' +
    '<div class="icon-row" id="arrAlign"></div>' +
    '<div data-min="2"><div class="sublabel">Align to</div></div>' +
    '<div class="seg seg3" id="arrAlignTo" data-min="2"></div>' +
    '<div class="group-label">Distribute</div>' +
    '<div class="inline"><label class="field">Gap between shapes<input type="text" id="arrGap" placeholder="e.g. g"></label>' +
      '<div class="pair"><button class="btn small" id="arrDistH" title="Place shapes left to right with this gap">Across</button><button class="btn small" id="arrDistV" title="Place shapes top to bottom with this gap">Down</button></div></div>' +
    '<div class="calc" id="arrGapCalc" data-min="2"></div>' +
    '<div class="btn-row" data-min="2"><button class="btn small" id="arrSpreadH" title="Equal gaps, keeping the outer shapes where they are">Even gaps across</button><button class="btn small" id="arrSpreadV" title="Equal gaps, keeping the outer shapes where they are">Even gaps down</button></div>' +
    '<div class="group-label">Grid</div>' +
    '<div class="seg seg2" id="arrSnapMode" data-min="2"></div>' +
    '<div class="toggle arr-cells" data-min="2">Resize to the cells it covers<label class="switch"><input type="checkbox" id="arrSnapResize"><span></span></label></div>' +
    '<div class="toggle arr-fine" data-min="2">Snap all four edges<label class="switch"><input type="checkbox" id="arrSnapEdges"><span></span></label></div>' +
    '<button class="btn primary" id="arrSnap">Snap to grid</button>' +
    '<div class="group-label" data-min="3">Move</div>' +
    '<div class="nudge-row" data-min="3">' +
      '<label class="field">Nudge by<input type="text" id="arrNudge" placeholder="e.g. g"></label>' +
      '<div class="nudge-pad"><button class="btn small" data-nudge="-1,0" title="Left">&larr;</button><button class="btn small" data-nudge="0,-1" title="Up">&uarr;</button><button class="btn small" data-nudge="0,1" title="Down">&darr;</button><button class="btn small" data-nudge="1,0" title="Right">&rarr;</button></div>' +
    '</div>' +
    '<button class="btn small" id="arrSwap" data-min="3" title="Swap the positions of two selected shapes">Swap two shapes</button>';

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

  // ---------- Snap ----------
  function nearest(list, v) {
    var best = 0;
    for (var i = 1; i < list.length; i++) if (Math.abs(list[i] - v) < Math.abs(list[best] - v)) best = i;
    return best;
  }
  function gridLines() {
    var g = ST.calculate(), xs = [], xe = [], ys = [], ye = [], x = g.left, y = g.top;
    for (var c = 0; c < g.cols; c++) { xs.push(x * PT); xe.push((x + g.widths[c]) * PT); x += g.widths[c] + g.gapH; }
    for (var r = 0; r < g.rows; r++) { ys.push(y * PT); ye.push((y + g.cellH) * PT); y += g.cellH + g.gapV; }
    return { xs: xs, xe: xe, ys: ys, ye: ye, grid: g };
  }
  function snap() {
    var g = ST.calculate();
    if (!g.valid) { ST.toast("The Layout grid isn't valid. Check its spacing and area.", true); return; }
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      if (ST.state.arrSnap === "fine") {
        var step = g.spacing * PT, ox = g.left * PT, oy = g.top * PT;
        if (!(step > 0)) { ST.toast("Grid spacing needs to be more than 0.", true); return; }
        var round = function (v, o) { return o + Math.round((v - o) / step) * step; };
        items.forEach(function (s) {
          var r = s.left + s.width, b = s.top + s.height;
          s.left = round(s.left, ox); s.top = round(s.top, oy);
          if (ST.state.arrSnapEdges) {
            s.width = Math.max(step, round(r, ox) - s.left);
            s.height = Math.max(step, round(b, oy) - s.top);
          }
        });
        return context.sync().then(function () { ST.toast("Snapped " + ST.plural(items.length, "shape") + " to steps of " + ST.nice(g.spacing) + " in."); ST.refreshSelection(); });
      }
      var L = gridLines();
      items.forEach(function (s) {
        var c = nearest(L.xs, s.left), r = nearest(L.ys, s.top);
        var right = s.left + s.width, bottom = s.top + s.height;
        s.left = L.xs[c]; s.top = L.ys[r];
        if (ST.state.arrSnapResize) {
          var ce = c + nearest(L.xe.slice(c), right), re = r + nearest(L.ye.slice(r), bottom);
          s.width = L.xe[ce] - L.xs[c]; s.height = L.ye[re] - L.ys[r];
        }
      });
      return context.sync().then(function () { ST.toast("Snapped " + ST.plural(items.length, "shape") + " to the grid cells."); ST.refreshSelection(); });
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
  function swap() {
    ST.withSelection(["left", "top"], function (context, items) {
      if (items.length !== 2) { ST.toast("Select exactly two shapes to swap.", true); return; }
      var a = items[0], b = items[1], l = a.left, t = a.top;
      a.left = b.left; a.top = b.top; b.left = l; b.top = t;
      return context.sync().then(function () { ST.toast("Swapped the two shapes."); ST.refreshSelection(); });
    });
  }

  // ---------- Build ----------
  function seg(id, options, key) {
    var el = $(id);
    options.forEach(function (o) {
      var b = document.createElement("button");
      b.textContent = o[1];
      b.addEventListener("click", function () { ST.state[key] = o[0]; ST.save(); renderSegs(); });
      b.setAttribute("data-v", o[0]);
      el.appendChild(b);
    });
  }
  function renderSegs() {
    [["arrAlignTo", "arrAlignTo"], ["arrSnapMode", "arrSnap"]].forEach(function (p) {
      Array.prototype.forEach.call($(p[0]).children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === ST.state[p[1]]); });
    });
    var fine = ST.state.arrSnap === "fine";
    Array.prototype.forEach.call(document.querySelectorAll(".arr-cells"), function (el) { el.classList.toggle("off", fine); });
    Array.prototype.forEach.call(document.querySelectorAll(".arr-fine"), function (el) { el.classList.toggle("off", !fine); });
  }
  function updateCalc() {
    var el = $("arrGapCalc");
    if (!el) return;
    var v = amount("arrGap");
    el.textContent = String(ST.state.arrGap || "").trim() ? (v === null ? "Gap: check the formula" : "Gap = " + ST.nice(v / PT) + " in") : "";
    el.className = v === null ? "calc bad" : "calc";
  }

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    var row = $("arrAlign");
    ALIGN.forEach(function (a) {
      var b = document.createElement("button");
      b.className = "tool-btn"; b.title = a[1]; b.setAttribute("aria-label", a[1]); b.innerHTML = a[2];
      b.addEventListener("click", function () { align(a[0]); });
      row.appendChild(b);
    });
    seg("arrAlignTo", [["selection", "Selection"], ["area", "Grid area"], ["key", "Largest"]], "arrAlignTo");
    seg("arrSnapMode", [["cells", "Grid cells"], ["fine", "Fine grid (g)"]], "arrSnap");
    ST.bindField($("arrGap"), "arrGap");
    ST.bindField($("arrNudge"), "arrNudge");
    [["arrSnapResize", "arrSnapResize"], ["arrSnapEdges", "arrSnapEdges"]].forEach(function (p) {
      var box = $(p[0]);
      box.checked = !!ST.state[p[1]];
      box.addEventListener("change", function () { ST.state[p[1]] = box.checked; ST.save(); });
      ST.on("reset", function () { box.checked = !!ST.state[p[1]]; });
    });
    $("arrDistH").addEventListener("click", function () { distribute(true); });
    $("arrDistV").addEventListener("click", function () { distribute(false); });
    $("arrGap").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); distribute(true); } });
    $("arrSpreadH").addEventListener("click", function () { spread(true); });
    $("arrSpreadV").addEventListener("click", function () { spread(false); });
    $("arrSnap").addEventListener("click", snap);
    $("arrSwap").addEventListener("click", swap);
    Array.prototype.forEach.call(body.querySelectorAll("[data-nudge]"), function (b) {
      var d = b.getAttribute("data-nudge").split(",").map(Number);
      b.addEventListener("click", function () { nudge(d[0], d[1]); });
    });
    renderSegs();
    ST.on("change", updateCalc);
    ST.on("reset", renderSegs);
  }

  ST.registerTool({
    id: "arrange",
    title: "Arrange",
    icon: ICON,
    levels: [1, 2, 3],
    defaultLevel: 1,
    defaults: defaults,
    build: build
  });
})();
