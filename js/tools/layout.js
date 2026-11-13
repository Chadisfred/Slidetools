/* Layout tool: grid size, presets, preview, position, Draw and Fit. */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;

  ST.PRESETS = [
    { id: "full", name: "Full", cols: 1 },
    { id: "split", name: "Split", cols: 2 },
    { id: "thirds", name: "Thirds", cols: 3 },
    { id: "quarters", name: "Quarters", cols: 4 },
    { id: "golden", name: "Golden", cols: 2, weights: [7.73, 4.77] },
    { id: "sidebar", name: "Sidebar", cols: 2, weights: [3, 9.5] },
    { id: "custom", name: "Custom", cols: 0 },
    { id: "960", name: "960", cols: 12 }
  ];
  var GRID_KEYS = ["spacing", "maxW", "maxH", "cols", "rows", "hIdx", "vIdx"];

  var defaults = {
    spacing: "0.42", maxW: "12.5", maxH: "5", cols: "5", rows: "1",
    left: "0.42", top: "1.25", hIdx: 2, vIdx: 2, preset: "custom",
    drawRounded: false, groupDrawn: false
  };
  ST.addDefaults(defaults);
  // Saved presets are kept out of the tool defaults so Reset doesn't delete them.
  ST.addDefaults({ customPresets: [] });

  var ICON = '<svg viewBox="0 0 18 18" fill="currentColor"><rect x="1.5" y="2.5" width="6.5" height="5.5" rx="1.3"/><rect x="10" y="2.5" width="6.5" height="5.5" rx="1.3"/><rect x="1.5" y="10" width="15" height="5.5" rx="1.3"/></svg>';

  var MARKUP =
    '<div class="chips" id="chips"></div>' +
    '<div class="stage" id="stage" data-min="2"></div>' +
    '<div class="result" id="result" data-min="2"></div>' +
    '<div class="tiles" data-min="3">' +
      '<div class="tile"><span>H GAP</span><b id="tGapH"></b></div>' +
      '<div class="tile"><span>V GAP</span><b id="tGapV"></b></div>' +
      '<div class="tile"><span>GRID</span><b id="tGrid"></b></div>' +
      '<div class="tile"><span>TOTAL</span><b id="tTotal"></b></div>' +
    '</div>' +
    '<div class="row">' +
      '<label class="field">Columns<div class="stepper"><input type="text" id="cols" inputmode="numeric"><button data-step="cols:-1" title="Fewer columns">&minus;</button><button data-step="cols:1" title="More columns">+</button></div></label>' +
      '<label class="field">Rows<div class="stepper"><input type="text" id="rows" inputmode="numeric"><button data-step="rows:-1" title="Fewer rows">&minus;</button><button data-step="rows:1" title="More rows">+</button></div></label>' +
    '</div>' +
    '<label class="field" data-min="2">Grid spacing (g)<input type="text" id="spacing" inputmode="decimal"></label>' +
    '<div class="row" data-min="2">' +
      '<label class="field">Area width<input type="text" id="maxW" inputmode="decimal"></label>' +
      '<label class="field">Area height<input type="text" id="maxH" inputmode="decimal"></label>' +
    '</div>' +
    '<div class="row" data-min="3">' +
      '<label class="field">Left<input type="text" id="left" inputmode="decimal"></label>' +
      '<label class="field">Top<input type="text" id="top" inputmode="decimal"></label>' +
    '</div>' +
    '<button class="btn small" id="useArea" data-min="2" title="Take the position and size of the selected shapes as the grid area">Use selection as area</button>' +
    '<div class="sublabel" data-min="2">Horizontal gap</div>' +
    '<div class="seg" id="segH" data-min="2"></div>' +
    '<div class="sublabel" data-min="2">Vertical gap</div>' +
    '<div class="seg" id="segV" data-min="2"></div>' +
    '<div class="toggle" data-min="3">Draw with rounded corners<label class="switch"><input type="checkbox" id="drawRounded"><span></span></label></div>' +
    '<div class="toggle" data-min="3">Group shapes after drawing<label class="switch"><input type="checkbox" id="groupDrawn"><span></span></label></div>' +
    '<div class="inline" data-min="3">' +
      '<label class="field">Save this grid as a preset<input type="text" id="presetName" placeholder="Preset name" maxlength="24"></label>' +
      '<button class="btn small" id="savePreset">Save</button>' +
    '</div>' +
    '<div class="btn-row">' +
      '<button class="btn primary" id="draw">Draw on slide</button>' +
      '<button class="btn" id="fit">Fit selection</button>' +
    '</div>';

  // ---------- Rendering ----------
  function render() {
    var grid = ST.grid, res = $("result");
    if (!grid.valid) {
      res.textContent = "Spacing is too large for this area";
      res.className = "result bad";
    } else if (grid.weighted) {
      res.textContent = grid.widths[0].toFixed(2) + '" + ' + grid.widths[1].toFixed(2) + '" x ' + grid.cellH.toFixed(2) + '"';
      res.className = "result";
    } else {
      res.textContent = grid.widths[0].toFixed(3) + " x " + grid.cellH.toFixed(3) + " in";
      res.className = "result";
    }
    $("tGapH").textContent = grid.gapH.toFixed(2) + '"';
    $("tGapV").textContent = grid.gapV.toFixed(2) + '"';
    $("tGrid").textContent = grid.cols + " x " + grid.rows;
    $("tTotal").textContent = grid.cols * grid.rows;
    $("draw").disabled = !grid.valid;
    $("fit").disabled = !grid.valid;
    renderSeg("segH", ST.state.hIdx);
    renderSeg("segV", ST.state.vIdx);
    renderChips();
    drawStage();
  }

  function drawStage() {
    var stage = $("stage"), grid = ST.grid;
    if (!stage.clientWidth) return; // hidden at this detail level
    var W = stage.clientWidth, H = stage.clientHeight, pad = 10;
    var cells = stage.querySelectorAll(".cell"), needed = grid.cols * grid.rows;
    for (var n = cells.length; n < needed; n++) {
      var d = document.createElement("div");
      d.className = "cell";
      stage.insertBefore(d, stage.querySelector(".sel-outline"));
    }
    for (n = cells.length - 1; n >= needed; n--) stage.removeChild(cells[n]);
    cells = stage.querySelectorAll(".cell");
    if (!grid.valid) {
      Array.prototype.forEach.call(cells, function (c) { c.style.opacity = 0; });
      drawOutlines(stage, null);
      return;
    }
    var scale = Math.min((W - 2 * pad) / grid.maxW, (H - 2 * pad) / grid.maxH);
    var ox = (W - grid.maxW * scale) / 2, oy = (H - grid.maxH * scale) / 2;
    var i = 0, y = oy;
    for (var r = 0; r < grid.rows; r++) {
      var x = ox;
      for (var c = 0; c < grid.cols; c++) {
        var el = cells[i++];
        el.style.opacity = 1;
        el.style.left = x + "px";
        el.style.top = y + "px";
        el.style.width = Math.max(1, grid.widths[c] * scale) + "px";
        el.style.height = Math.max(1, grid.cellH * scale) + "px";
        x += (grid.widths[c] + grid.gapH) * scale;
      }
      y += (grid.cellH + grid.gapV) * scale;
    }
    drawOutlines(stage, { scale: scale, ox: ox, oy: oy });
  }

  // Dashed outlines show where the selected shapes sit compared with the grid area.
  function drawOutlines(stage, frame) {
    var items = frame && ST.selection.count <= 40 ? ST.selection.items : [];
    var outlines = stage.querySelectorAll(".sel-outline");
    for (var n = outlines.length; n < items.length; n++) {
      stage.appendChild(Object.assign(document.createElement("div"), { className: "sel-outline" }));
    }
    for (n = outlines.length - 1; n >= items.length; n--) stage.removeChild(outlines[n]);
    outlines = stage.querySelectorAll(".sel-outline");
    items.forEach(function (s, i) {
      var o = outlines[i];
      o.style.left = (frame.ox + (s.left - ST.grid.left) * frame.scale) + "px";
      o.style.top = (frame.oy + (s.top - ST.grid.top) * frame.scale) + "px";
      o.style.width = Math.max(2, s.w * frame.scale) + "px";
      o.style.height = Math.max(2, s.h * frame.scale) + "px";
    });
  }

  function renderSeg(id, selected) {
    var el = $(id);
    if (!el.children.length) {
      ST.FRACTION_LABELS.forEach(function (label, i) {
        var b = document.createElement("button");
        b.textContent = label;
        b.addEventListener("click", function () {
          if (id === "segH") ST.state.hIdx = i; else ST.state.vIdx = i;
          markCustom("hIdx");
          ST.changed();
        });
        el.appendChild(b);
      });
    }
    Array.prototype.forEach.call(el.children, function (b, i) { b.classList.toggle("on", i === selected); });
  }

  function chipIcon(p) {
    var parts = p.weights ? p.weights : null, gap = 1.5, svg = "", x = 1, widths = [];
    var n = p.cols || 3;
    if (parts) { var t = parts[0] + parts[1]; widths = parts.map(function (w) { return (26 - gap) * w / t; }); }
    else if (p.id === "custom") { widths = [5, 9, 5, 4]; }
    else { n = Math.min(n, 12); for (var i = 0; i < n; i++) widths.push((26 - gap * (n - 1)) / n); }
    var rows = p.user ? Math.min(3, Math.max(1, Math.floor(ST.num(p.rows)) || 1)) : 1;
    var rh = (12 - (rows - 1) * 1.5) / rows;
    for (var r = 0; r < rows; r++) {
      x = 1;
      widths.forEach(function (w) {
        svg += '<rect x="' + x.toFixed(2) + '" y="' + (1 + r * (rh + 1.5)).toFixed(2) + '" width="' + w.toFixed(2) + '" height="' + rh.toFixed(2) + '" rx="1.2" fill="currentColor" opacity="0.85"/>';
        x += w + gap;
      });
    }
    return '<svg viewBox="0 0 28 14">' + svg + "</svg>";
  }

  var chipSignature = "";
  function renderChips() {
    var el = $("chips");
    var all = ST.PRESETS.concat(ST.state.customPresets || []);
    var sig = all.map(function (p) { return p.id; }).join("|");
    if (sig !== chipSignature) {
      chipSignature = sig;
      el.innerHTML = "";
      all.forEach(function (p) {
        var b = document.createElement("button");
        b.className = "chip";
        b.setAttribute("data-id", p.id);
        b.title = p.user ? p.name + " (your preset)" : p.name;
        b.innerHTML = chipIcon(p) + "<span></span>";
        b.querySelector("span").textContent = p.name;
        b.addEventListener("click", function () { applyPreset(p); });
        if (p.user) {
          var del = document.createElement("span");
          del.className = "del";
          del.title = "Delete this preset";
          del.textContent = "×";
          del.addEventListener("click", function (e) { e.stopPropagation(); deletePreset(p.id); });
          b.appendChild(del);
        }
        el.appendChild(b);
      });
    }
    Array.prototype.forEach.call(el.children, function (b) { b.classList.toggle("on", b.getAttribute("data-id") === ST.state.preset); });
  }

  // ---------- Presets ----------
  function applyPreset(p) {
    var s = ST.state;
    s.preset = p.id;
    if (p.user) {
      GRID_KEYS.forEach(function (k) { if (p[k] !== undefined) s[k] = p[k]; });
      ST.emit("reset"); // refresh the boxes
    } else if (p.cols) {
      s.cols = String(p.cols); $("cols").value = s.cols;
    }
    ST.changed();
  }
  function markCustom(key) {
    var s = ST.state, p = ST.presetById(s.preset);
    if (!p || p.id === "custom") return;
    if (p.user) { s.preset = "custom"; return; }
    if (key === "cols" && String(p.cols) === String(s.cols)) return;
    if (["spacing", "maxW", "maxH", "cols"].indexOf(key) >= 0) s.preset = "custom";
  }
  function savePreset() {
    var name = $("presetName").value.trim();
    if (!name) { ST.toast("Type a name for the preset first.", true); $("presetName").focus(); return; }
    var p = { id: "u" + Date.now(), name: name, user: true };
    GRID_KEYS.forEach(function (k) { p[k] = ST.state[k]; });
    ST.state.customPresets = (ST.state.customPresets || []).concat([p]);
    ST.state.preset = p.id;
    $("presetName").value = "";
    ST.changed();
    ST.toast("Saved preset “" + name + "”.");
  }
  function deletePreset(id) {
    var p = ST.presetById(id);
    ST.state.customPresets = (ST.state.customPresets || []).filter(function (x) { return x.id !== id; });
    if (ST.state.preset === id) ST.state.preset = "custom";
    ST.changed();
    ST.toast("Deleted preset “" + (p ? p.name : "") + "”.");
  }

  // ---------- Slide actions ----------
  function useArea() {
    if (!ST.needOffice()) return;
    var items = ST.selection.items;
    if (!items.length) { ST.toast("Select a shape or placeholder to use as the grid area.", true); return; }
    var l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
    items.forEach(function (s) {
      l = Math.min(l, s.left); t = Math.min(t, s.top);
      r = Math.max(r, s.left + s.w); b = Math.max(b, s.top + s.h);
    });
    var s = ST.state;
    s.left = ST.exact(l); s.top = ST.exact(t); s.maxW = ST.exact(r - l); s.maxH = ST.exact(b - t);
    markCustom("maxW");
    ST.emit("reset");
    ST.changed();
    ST.toast("Grid area set to " + s.maxW + " x " + s.maxH + " in.");
  }

  function drawGrid() {
    if (!ST.needOffice()) return;
    var grid = ST.calculate();
    if (!grid.valid) { ST.toast("The grid doesn't fit its area. Check the spacing, size and number of columns and rows.", true); return; }
    var rounded = ST.state.drawRounded, group = ST.state.groupDrawn;
    ST.run(function (context) {
      var slides = context.presentation.getSelectedSlides();
      slides.load("items/id");
      return context.sync().then(function () {
        if (!slides.items.length) { ST.toast("Select a slide first.", true); return; }
        var slide = slides.items[0], made = [];
        var y = grid.top;
        for (var r = 0; r < grid.rows; r++) {
          var x = grid.left;
          for (var c = 0; c < grid.cols; c++) {
            var shape = slide.shapes.addGeometricShape(
              rounded ? PowerPoint.GeometricShapeType.roundRectangle : PowerPoint.GeometricShapeType.rectangle,
              { left: x * PT, top: y * PT, width: grid.widths[c] * PT, height: grid.cellH * PT });
            shape.name = "Grid R" + (r + 1) + " C" + (c + 1);
            shape.load("id");
            made.push({ shape: shape, w: grid.widths[c] * PT, h: grid.cellH * PT });
            x += grid.widths[c] + grid.gapH;
          }
          y += grid.cellH + grid.gapV;
        }
        return context.sync().then(function () {
          if (rounded && ST.cornersSupported()) {
            made.forEach(function (m) { m.shape.adjustments.set(0, ST.cornerValue(m.w, m.h)); });
            return context.sync();
          }
        }).then(function () {
          return finishDraw(context, slide, made.map(function (m) { return m.shape.id; }), group && made.length > 1);
        }).then(function (note) {
          ST.toast("Drew " + ST.plural(made.length, "shape") + "." + (note || ""));
        });
      });
    });
  }

  // Groups and selects the new shapes where this PowerPoint version allows it.
  function finishDraw(context, slide, ids, group) {
    var note = "";
    var step = Promise.resolve(ids);
    if (group) {
      if (ST.supports("1.8")) {
        step = step.then(function () {
          var g = slide.shapes.addGroup(ids);
          g.load("id");
          g.name = "Grid";
          return context.sync().then(function () { return [g.id]; });
        }).catch(function () { note = " Grouping isn't available here."; return ids; });
      } else {
        note = " This PowerPoint version can't group from the panel.";
      }
    }
    return step.then(function (selectIds) {
      if (!ST.supports("1.5")) return note;
      slide.setSelectedShapes(selectIds);
      return context.sync().then(function () { return note; }, function () { return note; });
    });
  }

  function fitSelection() {
    var grid = ST.calculate();
    if (!grid.valid) { ST.toast("The grid doesn't fit its area. Check the spacing, size and number of columns and rows.", true); return; }
    ST.withSelection(["left", "top"], function (context, items) {
      var cells = grid.cols * grid.rows;
      if (items.length > cells) {
        ST.toast(items.length + " shapes are selected, but the grid has " + cells + " cells.", true);
        return;
      }
      // Read order: top to bottom, then left to right within each row
      items.sort(function (a, b) { return a.top - b.top; });
      var ordered = [];
      for (var i = 0; i < items.length; i += grid.cols) {
        ordered = ordered.concat(items.slice(i, i + grid.cols).sort(function (a, b) { return a.left - b.left; }));
      }
      ordered.forEach(function (s, i) {
        var r = Math.floor(i / grid.cols), c = i % grid.cols;
        var x = grid.left;
        for (var k = 0; k < c; k++) x += grid.widths[k] + grid.gapH;
        s.left = x * PT;
        s.top = (grid.top + r * (grid.cellH + grid.gapV)) * PT;
        s.width = grid.widths[c] * PT;
        s.height = grid.cellH * PT;
      });
      return context.sync().then(function () {
        ST.toast("Fitted " + ST.plural(ordered.length, "shape") + ".");
        ST.refreshSelection();
      });
    });
  }

  // ---------- Build ----------
  function build(body) {
    body.appendChild(ST.html(MARKUP));
    ["spacing", "maxW", "maxH", "cols", "rows", "left", "top"].forEach(function (id) {
      ST.bindField($(id), id, { number: true, onInput: function () { markCustom(id); } });
    });
    Array.prototype.forEach.call(body.querySelectorAll("[data-step]"), function (b) {
      b.addEventListener("click", function (e) {
        e.preventDefault();
        var parts = b.getAttribute("data-step").split(":"), id = parts[0], max = id === "cols" ? ST.MAX_COLS : ST.MAX_ROWS;
        var n = Math.min(max, Math.max(1, (Math.floor(ST.num(ST.state[id])) || 1) + parseInt(parts[1], 10)));
        ST.state[id] = String(n);
        $(id).value = ST.state[id];
        markCustom(id);
        if (id === "cols" && ST.presetById(ST.state.preset) && !ST.presetById(ST.state.preset).user) ST.state.preset = "custom";
        ST.changed();
      });
    });
    ["drawRounded", "groupDrawn"].forEach(function (id) {
      var box = $(id);
      box.checked = !!ST.state[id];
      box.addEventListener("change", function () { ST.state[id] = box.checked; ST.save(); });
      ST.on("reset", function () { box.checked = !!ST.state[id]; });
    });
    $("presetName").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); savePreset(); } });
    $("savePreset").addEventListener("click", savePreset);
    $("useArea").addEventListener("click", useArea);
    $("draw").addEventListener("click", drawGrid);
    $("fit").addEventListener("click", fitSelection);

    ST.on("change", render);
    ST.on("selection", drawStage);
    ST.on("layout", drawStage);
    window.addEventListener("resize", drawStage);
  }

  ST.registerTool({
    id: "layout",
    title: "Layout",
    icon: ICON,
    levels: [1, 2, 3],
    defaultLevel: 2,
    defaults: defaults,
    build: build
  });
})();
