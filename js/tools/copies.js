/* Copies tool (id "copies"): makes moved, scaled, rotated and reflected copies of the
   selected shapes in one step, like Illustrator's Transform effect. The reference point
   is shared with Size. */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;

  var defaults = { cpCount: "3", cpRows: "1", cpMoveH: "w+gap", cpMoveV: "", cpMode: "distance", cpScaleH: "100", cpScaleV: "100", cpUniform: true, cpRotate: "0", cpReflectX: false, cpReflectY: false };
  ST.addDefaults(defaults);

  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round"><rect x="1.5" y="1.5" width="8" height="8" rx="1.5"/><rect x="5" y="5" width="8" height="8" rx="1.5" opacity="0.7"/><rect x="8.5" y="8.5" width="8" height="8" rx="1.5" opacity="0.45"/></svg>';

  var MARKUP =
    '<div class="row">' +
      '<label class="field"><span class="label-row">Copies' + ST.info("Makes copies of each selected shape in one step, like Illustrator's Transform effect. Each copy is moved, scaled, rotated and reflected from the one before. Boxes take formulas, where <code>w</code> and <code>h</code> are the shape's size. Leave a Move box empty for no movement that way. PowerPoint add-ins can't copy shapes directly, so the slide is swapped for an identical copy with the new shapes. Slide comments may not carry over, and Undo takes more than one step.", "About copies") + '</span><div class="stepper"><input type="text" id="cpCount" inputmode="numeric"><button id="cpLess" title="Fewer copies">&minus;</button><button id="cpMore" title="More copies">+</button></div></label>' +
      '<label class="field" data-min="2"><span class="label-row">Rows' + ST.info("More than 1 makes a grid: each row repeats the original and its copies, one Move down lower. With Move down empty, rows are the shape\'s height plus the row gap apart.", "About rows") + '</span><div class="stepper"><input type="text" id="cpRows" inputmode="numeric"><button id="cpRowsLess" title="Fewer rows">&minus;</button><button id="cpRowsMore" title="More rows">+</button></div></label>' +
    '</div>' +
    '<div class="field" data-min="2">Space by<div class="seg seg2" id="cpMode"><button data-v="distance" title="Move each copy by this distance, like Illustrator">Distance</button><button data-v="gap" title="Leave this gap between the edges of each copy">Gap</button></div></div>' +
    '<div class="row">' +
      '<label class="field">Move across<input type="text" id="cpMoveH" placeholder="e.g. w+gap"></label>' +
      '<label class="field">Move down<input type="text" id="cpMoveV" placeholder="e.g. h+vgap"></label>' +
    '</div>' +
    '<div class="row" data-min="2">' +
      '<label class="field">Scale across %<input type="text" id="cpScaleH" inputmode="decimal"></label>' +
      '<label class="field">Scale down %<input type="text" id="cpScaleV" inputmode="decimal"></label>' +
    '</div>' +
    '<div class="toggle" data-min="2">Scale evenly<label class="switch"><input type="checkbox" id="cpUniform"><span></span></label></div>' +
    '<label class="field" data-min="2"><span class="label-row">Rotate degrees' + ST.info("Each copy turns this much more than the one before, around the reference point. Positive turns counterclockwise, like Illustrator.", "About rotating copies") + '</span><input type="text" id="cpRotate" inputmode="decimal"></label>' +
    '<div class="anchor-row" data-min="2"><span class="anchor-text"><span class="label-row">Reference point' + ST.info("The point copies rotate, scale and reflect around. Size uses the same point.", "About the reference point") + '</span><b id="cpAnchorName"></b></span><div class="anchor" id="cpAnchor"></div></div>' +
    '<div class="toggle" data-min="3">Reflect X (top to bottom)<label class="switch"><input type="checkbox" id="cpReflectX"><span></span></label></div>' +
    '<div class="toggle" data-min="3">Reflect Y (left to right)<label class="switch"><input type="checkbox" id="cpReflectY"><span></span></label></div>' +
    '<div class="cp-preview" id="cpPreview" data-min="2"></div>' +
    '<div class="calc" id="cpCalc" data-min="2"></div>' +
    '<button class="btn primary" id="cpGo">Make copies</button>';

  // ---------- Copies ----------
  var MAX_COPIES = 50;
  function copiesText(n) { return n === 1 ? "1 copy" : n + " copies"; }
  function copyCount(S) { var n = Math.floor(ST.evaluate((S || ST.state).cpCount, null, null)); return n >= 1 && n <= MAX_COPIES ? n : null; }
  var MAX_ROWS = 20;
  function rowCount(S) { var v = String((S || ST.state).cpRows || "").trim(); if (!v) return 1; var n = Math.floor(ST.evaluate(v, null, null)); return n >= 1 && n <= MAX_ROWS ? n : null; }
  // Shapes added per selected shape
  function totalCopies(n, rows) { return (n + 1) * rows - 1; }
  // Works out the copy settings for a shape of this size (inches). Returns { t } or { bad }.
  var COPY_KEYS = ["cpCount", "cpRows", "cpMoveH", "cpMoveV", "cpMode", "cpScaleH", "cpScaleV", "cpUniform", "cpRotate", "cpReflectX", "cpReflectY", "anchor"];
  // S: the settings to use, the panel's by default (Repeat passes a saved copy).
  function copySettings(wIn, hIn, S) {
    S = S || ST.state;
    var vars = ST.formulaVars(wIn, hIn), a0 = S.anchor || 0, f = { x: (a0 % 3) / 2, y: Math.floor(a0 / 3) / 2 };
    var num = function (text, fallback) { var x = String(text || "").trim(); if (!x) return fallback; return ST.evaluate(x, vars, null); };
    var mh = num(S.cpMoveH, NaN), mv = num(S.cpMoveV, NaN), sx = num(S.cpScaleH, 100), sy = num(S.cpUniform ? S.cpScaleH : S.cpScaleV, 100), a = num(S.cpRotate, 0);
    var names = "w, h, g, gap, vgap, col and row";
    var moveBad = function (label, text, example, which) {
      var why = ST.formulaProblem(text, vars), hint = /isn't a name/.test(why) ? " The names you can use are " + names + "." : "";
      return { bad: label + " " + (why || "can't be read") + "." + hint + " Change it to something like " + example + ", or clear the box for no movement " + which + ".", field: label === "Move across" ? "cpMoveH" : "cpMoveV" };
    };
    if (mh === null) return moveBad("Move across", S.cpMoveH, "w+gap", "across");
    if (mv === null) return moveBad("Move down", S.cpMoveV, "h+vgap", "down");
    if (!(sx > 0)) return { bad: "Check Scale across. It must be more than 0." };
    if (!(sy > 0)) return { bad: "Check Scale down. It must be more than 0." };
    if (a === null) return { bad: "Check Rotate." };
    return { t: { vgap: (ST.grid ? ST.grid.gapV : 0) * PT, sx: sx / 100, sy: sy / 100, angle: a, reflectX: !!S.cpReflectX, reflectY: !!S.cpReflectY, fx: f.x, fy: f.y,
                  gap: S.cpMode === "gap", mh: isNaN(mh) ? null : mh * PT, mv: isNaN(mv) ? null : mv * PT } };
  }
  function updateCopies() {
    if (!ST.grid || !ST.grid.widths) ST.calculate();
    var sel = ST.selection, n = copyCount(), calc = $("cpCalc"), prev = $("cpPreview");
    var w = sel.count ? Math.max(sel.w, 0.01) : 1.5, h = sel.count ? Math.max(sel.h, 0.01) : 1;
    var rows = rowCount();
    var c = copySettings(w, h), bad = n === null ? "Copies must be 1 to " + MAX_COPIES + "." : rows === null ? "Rows must be 1 to " + MAX_ROWS + "." : c.bad;
    ["cpMoveH", "cpMoveV"].forEach(function (k) { $(k).classList.toggle("err", !!c.field && c.field === k); });
    // A problem shows at every level, so the reason is visible before clicking Make copies
    if (bad) calc.removeAttribute("data-min"); else calc.setAttribute("data-min", "2");
    calc.className = bad ? "calc bad" : "calc";
    $("cpGo").disabled = !!bad;
    if (bad) { calc.textContent = bad; prev.innerHTML = ""; return; }
    var t = c.t, F = ST.clone.frameOf({ left: 0, top: 0, width: w * PT, height: h * PT }), list = [F].concat(ST.clone.gridFrames(F, t, n, rows));
    var parts = [rows > 1 ? (n + 1) + " across x " + rows + " rows, " + copiesText(totalCopies(n, rows)) : copiesText(n)];
    if (t.mh !== null || t.mv !== null) parts.push((t.gap ? "gaps " : "moves ") + [t.mh !== null ? ST.nice(t.mh / PT) + " in across" : "", t.mv !== null ? ST.nice(t.mv / PT) + " in down" : ""].filter(Boolean).join(", "));
    if (t.sx !== 1 || t.sy !== 1) parts.push("scaled " + ST.nice(t.sx * 100) + (t.sx !== t.sy ? " x " + ST.nice(t.sy * 100) : "") + "%");
    if (t.angle) parts.push("turned " + ST.nice(t.angle) + " degrees");
    calc.textContent = parts.join(", ") + (sel.count ? "" : ". Preview uses a sample shape.");
    // Preview: outlines of the original and its copies
    var pts = [];
    list.forEach(function (G) { pts = pts.concat(ST.clone.corners(G)); });
    var xs = pts.map(function (p) { return p[0]; }), ys = pts.map(function (p) { return p[1]; });
    var minX = Math.min.apply(null, xs), minY = Math.min.apply(null, ys), spanX = Math.max.apply(null, xs) - minX || 1, spanY = Math.max.apply(null, ys) - minY || 1;
    var W = 260, H = 110, pad = 8, k = Math.min((W - 2 * pad) / spanX, (H - 2 * pad) / spanY);
    var ox = (W - spanX * k) / 2, oy = (H - spanY * k) / 2, P = function (p) { return ((p[0] - minX) * k + ox).toFixed(1) + "," + ((p[1] - minY) * k + oy).toFixed(1); };
    var svg = list.map(function (G, i) { return '<polygon points="' + ST.clone.corners(G).map(P).join(" ") + '" class="' + (i ? "cp-copy" : "cp-orig") + '"/>'; }).join("");
    var r = P(ST.clone.refPoint(F, t.fx, t.fy)).split(",");
    svg += '<circle cx="' + r[0] + '" cy="' + r[1] + '" r="3" class="cp-ref"/>';
    prev.innerHTML = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Preview of the copies">' + svg + '</svg>';
  }
  function renderMode() { Array.prototype.forEach.call($("cpMode").children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === ST.state.cpMode); }); }
  function syncUniform() { var u = !!ST.state.cpUniform; $("cpScaleV").disabled = u; if (u) { ST.state.cpScaleV = ST.state.cpScaleH; $("cpScaleV").value = ST.state.cpScaleV; } }
  function makeCopies(saved) {
    if (!ST.needOffice()) return;
    if (!ST.clone.supported()) { ST.toast("This version of PowerPoint can't make copies. It needs PowerPoint API 1.8.", true); return; }
    ST.calculate();
    var S = saved || {};
    if (!saved) COPY_KEYS.forEach(function (k) { S[k] = ST.state[k]; });
    var n = copyCount(S), rows = rowCount(S);
    if (n === null) { ST.toast("Copies must be 1 to " + MAX_COPIES + ".", true); return; }
    if (rows === null) { ST.toast("Rows must be 1 to " + MAX_ROWS + ".", true); return; }
    ST.setRepeat("Make " + copiesText(totalCopies(n, rows)), function () { makeCopies(S); });
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      var jobs = [];
      for (var i = 0; i < items.length; i++) {
        var s = items[i], c = copySettings(s.width / PT, s.height / PT, S);
        if (c.bad) { ST.toast(c.bad, true); return; }
        jobs.push({ shape: s, t: c.t, copies: n, rows: rows });
      }
      return ST.confirmSwap("copy shapes").then(function (ok) {
        if (!ok) return { cancelled: true };
        return ST.clone.apply(context, jobs);
      }).then(function (r) {
        if (r.cancelled) { ST.toast("No copies were made."); return; }
        var msg = r.made ? "Made " + copiesText(totalCopies(n, rows)) + " of " + ST.plural(r.shapes, "shape") + "." : "No copies were made.";
        if (r.missed.length) msg += " Couldn't find " + r.missed.join(", ") + " in the slide.";
        if (r.fixed) msg += " Tables, charts and other frames can't turn or flip, and tables keep their size, so those settings were skipped for them.";
        ST.toast(msg, !r.made);
        ST.refreshSelection();
      }, function (e) { ST.toast("Couldn't make the copies. " + (ST.errorText ? ST.errorText(e) : ""), true); });
    });
  }

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    ["cpCount", "cpRows", "cpMoveH", "cpMoveV", "cpScaleH", "cpScaleV", "cpRotate"].forEach(function (k) { ST.bindField($(k), k, k === "cpScaleH" ? { onInput: function () { if (ST.state.cpUniform) { ST.state.cpScaleV = ST.state.cpScaleH; $("cpScaleV").value = ST.state.cpScaleV; } } } : null); });
    var stepCount = function (d) { var n = (copyCount() || 1) + d; n = Math.max(1, Math.min(MAX_COPIES, n)); ST.state.cpCount = String(n); $("cpCount").value = ST.state.cpCount; ST.changed(); };
    $("cpLess").addEventListener("click", function (e) { e.preventDefault(); stepCount(-1); });
    $("cpMore").addEventListener("click", function (e) { e.preventDefault(); stepCount(1); });
    var stepRows = function (d) { var n = (rowCount() || 1) + d; n = Math.max(1, Math.min(MAX_ROWS, n)); ST.state.cpRows = String(n); $("cpRows").value = ST.state.cpRows; ST.changed(); };
    $("cpRowsLess").addEventListener("click", function (e) { e.preventDefault(); stepRows(-1); });
    $("cpRowsMore").addEventListener("click", function (e) { e.preventDefault(); stepRows(1); });
    Array.prototype.forEach.call($("cpMode").children, function (b) { b.addEventListener("click", function () { ST.state.cpMode = b.getAttribute("data-v"); renderMode(); ST.changed(); }); });
    var switches = ["cpUniform", "cpReflectX", "cpReflectY"];
    switches.forEach(function (k) { var el = $(k); el.checked = !!ST.state[k]; el.addEventListener("change", function () { ST.state[k] = el.checked; if (k === "cpUniform") syncUniform(); ST.changed(); }); });
    $("cpGo").addEventListener("click", function () { makeCopies(); });
    if (ST.anchorGrid) ST.anchorGrid("cpAnchor", "cpAnchorName");
    renderMode(); syncUniform();
    ST.on("change", updateCopies);
    ST.on("selection", updateCopies);
    updateCopies();
    ST.on("reset", function () {
      switches.forEach(function (k) { $(k).checked = !!ST.state[k]; });
      ["cpCount", "cpRows", "cpMoveH", "cpMoveV", "cpScaleH", "cpScaleV", "cpRotate"].forEach(function (k) { $(k).value = ST.state[k]; });
      renderMode(); syncUniform(); updateCopies();
    });
  }

  ST.registerTool({
    id: "copies",
    title: "Copies",
    icon: ICON,
    levels: [1, 2, 3],
    defaultLevel: 2,
    defaults: defaults,
    tip: { what: "Makes spaced copies of shapes in one step, in a row or a grid.", steps: ["Select shapes.", "Set Copies and Move across, such as w+gap.", "Click Make copies."] },
    build: build
  });
})();
