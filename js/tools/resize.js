/* Resize tool: set width and height with formulas, lock aspect, anchor, match sizes. */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;
  var MAX_RECENT = 6;

  var defaults = { resW: "", resH: "", lockAspect: false, anchor: 0 };
  ST.addDefaults(defaults);
  ST.addDefaults({ recentFormulas: [] });

  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="6" width="10" height="10" rx="1.5"/><path d="M9 2h7v7"/><path d="M16 2l-6 6"/></svg>';
  var ANCHOR_NAMES = ["top left", "top", "top right", "left", "centre", "right", "bottom left", "bottom", "bottom right"];

  var MARKUP =
    '<span class="info res-info" tabindex="0" role="button" aria-label="How formulas work">i' +
      '<span class="tip" role="tooltip">Press Enter in a box to apply just that size. With the lock on, the other side follows, and Resize uses Width.<br><br>' +
      'Names you can use: <code>g</code> grid spacing, <code>col</code> column width, <code>row</code> row height, ' +
      '<code>gap</code> and <code>vgap</code> the gaps, <code>w</code> and <code>h</code> the shape\'s current size.<br><br>' +
      'Examples: <code>1.5g</code> <code>10g</code> <code>2col+gap</code> <code>*2</code> <code>1.4*2</code>. Every number box in this panel accepts maths too.</span>' +
    '</span>' +
    '<div class="row">' +
      '<label class="field">Width<input type="text" id="resW" placeholder="e.g. 2col+gap"></label>' +
      '<label class="field">Height<input type="text" id="resH" placeholder="e.g. 10g"></label>' +
    '</div>' +
    '<div class="calc" id="calc" data-min="2"></div>' +
    '<div class="toggle" data-min="2">Lock aspect ratio<label class="switch"><input type="checkbox" id="lockAspect"><span></span></label></div>' +
    '<div class="anchor-row" data-min="2"><span>Resize from <b id="anchorName"></b></span><div class="anchor" id="anchor"></div></div>' +
    '<div class="btn-row res-actions">' +
      '<button class="btn" id="useSel" data-min="2" title="Copy the selected shape\'s size into the boxes">Use selection</button>' +
      '<button class="btn primary" id="resize">Resize</button>' +
    '</div>' +
    '<div class="group-label" data-min="2">Match selected shapes to the</div>' +
    '<div class="btn-row" data-min="2">' +
      '<button class="btn small" id="matchLarge">Largest</button>' +
      '<button class="btn small" id="matchSmall">Smallest</button>' +
    '</div>' +
    '<div class="group-label" data-min="3">Recent formulas</div>' +
    '<div class="mini-chips" id="recent" data-min="3"></div>';

  function anchorFactors() {
    var a = ST.state.anchor || 0;
    return { x: (a % 3) / 2, y: Math.floor(a / 3) / 2 };
  }

  // Sets a shape's size, keeping the chosen anchor point still.
  function applySize(s, nw, nh) {
    var f = anchorFactors();
    if (nw !== null) { s.left = s.left + (s.width - nw) * f.x; s.width = nw; }
    if (nh !== null) { s.top = s.top + (s.height - nh) * f.y; s.height = nh; }
  }

  function updateCalc() {
    var parts = [], bad = false, sel = ST.selection;
    [["resW", "Width", "w"], ["resH", "Height", "h"]].forEach(function (f) {
      var text = ST.state[f[0]];
      if (!text || !String(text).trim()) return;
      var v = ST.evaluate(text, ST.formulaVars(sel.count ? sel.w : undefined, sel.count ? sel.h : undefined), f[2]);
      if (v === null || v <= 0) { parts.push(f[1] + ": check the formula"); bad = true; }
      else parts.push(f[1] + " = " + ST.nice(v) + " in");
    });
    var calc = $("calc");
    calc.textContent = parts.join("     ");
    calc.className = bad ? "calc bad" : "calc";
  }

  function renderAnchor() {
    var el = $("anchor");
    if (!el.children.length) {
      ANCHOR_NAMES.forEach(function (name, i) {
        var b = document.createElement("button");
        b.title = "Resize from the " + name;
        b.setAttribute("aria-label", b.title);
        b.addEventListener("click", function () { ST.state.anchor = i; ST.save(); renderAnchor(); });
        el.appendChild(b);
      });
    }
    Array.prototype.forEach.call(el.children, function (b, i) { b.classList.toggle("on", i === (ST.state.anchor || 0)); });
    $("anchorName").textContent = ANCHOR_NAMES[ST.state.anchor || 0];
  }

  function renderRecent() {
    var el = $("recent"), list = ST.state.recentFormulas || [];
    el.innerHTML = "";
    if (!list.length) {
      el.appendChild(Object.assign(document.createElement("span"), { className: "muted-note", textContent: "Formulas you resize with will appear here." }));
      return;
    }
    list.forEach(function (r) {
      var b = document.createElement("button");
      b.className = "mini-chip";
      b.textContent = (r.w ? "W " + r.w : "") + (r.w && r.h ? "   " : "") + (r.h ? "H " + r.h : "");
      b.title = "Put this back in the boxes";
      b.addEventListener("click", function () {
        ST.state.resW = r.w || ""; ST.state.resH = r.h || "";
        $("resW").value = ST.state.resW; $("resH").value = ST.state.resH;
        ST.changed();
      });
      el.appendChild(b);
    });
  }

  function remember(w, h) {
    if (!w && !h) return;
    var list = (ST.state.recentFormulas || []).filter(function (r) { return !(r.w === w && r.h === h); });
    list.unshift({ w: w, h: h });
    ST.state.recentFormulas = list.slice(0, MAX_RECENT);
    renderRecent();
  }

  function useSelection() {
    if (!ST.needOffice()) return;
    if (!ST.selection.count) { ST.toast("Select a shape first.", true); return; }
    ST.state.resW = ST.nice(ST.selection.w); ST.state.resH = ST.nice(ST.selection.h);
    $("resW").value = ST.state.resW; $("resH").value = ST.state.resH;
    ST.changed();
  }

  function resizeSelection(doW, doH) {
    if (!ST.needOffice()) return;
    ST.calculate();
    var lock = !!ST.state.lockAspect;
    var wText = String(ST.state.resW || "").trim(), hText = String(ST.state.resH || "").trim();
    doW = doW && !!wText;
    doH = doH && !!hText;
    if (lock && doW && doH) doH = false; // width leads when the lock is on
    if (!doW && !doH) { ST.toast("Type a width or height first.", true); return; }
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      var updates = [];
      for (var i = 0; i < items.length; i++) {
        var s = items[i], vars = ST.formulaVars(s.width / PT, s.height / PT);
        var nw = doW ? ST.evaluate(wText, vars, "w") : null;
        var nh = doH ? ST.evaluate(hText, vars, "h") : null;
        if ((doW && !(nw > 0)) || (doH && !(nh > 0))) { ST.toast("That formula couldn't be worked out.", true); return; }
        if (nw !== null) nw *= PT;
        if (nh !== null) nh *= PT;
        if (lock && nw !== null) nh = s.height * nw / s.width;
        if (lock && nh !== null && nw === null) nw = s.width * nh / s.height;
        updates.push([s, nw, nh]);
      }
      updates.forEach(function (u) { applySize(u[0], u[1], u[2]); });
      return context.sync().then(function () {
        remember(doW ? wText : "", doH ? hText : "");
        var first = updates[0];
        if (doW) { ST.state.resW = ST.nice(first[1] / PT); $("resW").value = ST.state.resW; }
        if (doH) { ST.state.resH = ST.nice(first[2] / PT); $("resH").value = ST.state.resH; }
        ST.changed();
        ST.toast("Resized " + ST.plural(updates.length, "shape") + ".");
        ST.refreshSelection();
      });
    });
  }

  function matchSize(largest) {
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      if (items.length < 2) { ST.toast("Select at least two shapes to match.", true); return; }
      var target = items.reduce(function (best, s) {
        var a = s.width * s.height, b = best.width * best.height;
        return (largest ? a > b : a < b) ? s : best;
      });
      var tw = target.width, th = target.height, changed = 0;
      items.forEach(function (s) {
        if (s === target) return;
        applySize(s, tw, th);
        changed++;
      });
      return context.sync().then(function () {
        ST.toast("Matched " + ST.plural(changed, "shape") + " to " + ST.nice(tw / PT) + " x " + ST.nice(th / PT) + " in.");
        ST.refreshSelection();
      });
    });
  }

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    ST.bindField($("resW"), "resW");
    ST.bindField($("resH"), "resH");
    $("resW").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); resizeSelection(true, false); } });
    $("resH").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); resizeSelection(false, true); } });
    var lock = $("lockAspect");
    lock.checked = !!ST.state.lockAspect;
    lock.addEventListener("change", function () { ST.state.lockAspect = lock.checked; ST.save(); });
    $("useSel").addEventListener("click", useSelection);
    $("resize").addEventListener("click", function () { resizeSelection(true, true); });
    $("matchLarge").addEventListener("click", function () { matchSize(true); });
    $("matchSmall").addEventListener("click", function () { matchSize(false); });
    renderAnchor();
    renderRecent();
    ST.on("change", updateCalc);
    ST.on("selection", updateCalc);
    ST.on("reset", function () { lock.checked = !!ST.state.lockAspect; renderAnchor(); });
  }

  ST.registerTool({
    id: "resize",
    title: "Resize",
    icon: ICON,
    levels: [1, 2, 3],
    defaultLevel: 2,
    defaults: defaults,
    build: build
  });
})();
