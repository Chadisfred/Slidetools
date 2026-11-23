/* Size tool (id "resize"): set sizes with formulas, lock aspect, reference point,
   match sizes and crop pictures. The reference point is shared with Copies. */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;
  var MAX_RECENT = 6;

  var defaults = { resW: "", resH: "", lockAspect: false, anchor: 0, cropPictures: true };
  ST.addDefaults(defaults);
  ST.addDefaults({ recentFormulas: [] });

  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="6" width="10" height="10" rx="1.5"/><path d="M9 2h7v7"/><path d="M16 2l-6 6"/></svg>';
  var ANCHOR_NAMES = ["top left", "top", "top right", "left", "centre", "right", "bottom left", "bottom", "bottom right"];

  var MARKUP =
    '<div class="row">' +
      '<label class="field"><span class="label-row">Width' + ST.info("Numbers or formulas. <code>g</code> grid spacing, <code>col</code> column width, <code>row</code> row height, <code>gap</code>, <code>w</code> and <code>h</code> the current size. Examples: <code>2col+gap</code>, <code>10g</code>, <code>*2</code>.", "About formulas") + '</span><input type="text" id="resW" placeholder="e.g. 2col+gap"></label>' +
      '<label class="field">Height<input type="text" id="resH" placeholder="e.g. 10g"></label>' +
    '</div>' +
    '<div class="calc" id="calc" data-min="2"></div>' +
    '<div class="toggle" data-min="2">Lock aspect ratio<label class="switch"><input type="checkbox" id="lockAspect"><span></span></label></div>' +
    '<div class="anchor-row" data-min="2"><span class="anchor-text"><span class="label-row">Reference point' + ST.info("The point that stays still when shapes are resized, matched or scaled, and the point copies rotate and reflect around. Size and Copies share it.", "About the reference point") + '</span><b id="anchorName"></b></span><div class="anchor" id="anchor"></div></div>' +
    '<div class="btn-row res-actions">' +
      '<button class="btn" id="useSel" data-min="2" title="Copy the selected shape\'s size into the boxes">Use selection</button>' +
      '<button class="btn primary" id="resize">Resize</button>' +
    '</div>' +
    '<div class="group-label" data-min="2">Match size to</div>' +
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

  // A line has no area: its type says so, or one side is 0.
  function isLine(s) { return s.type === "Line" || s.width === 0 || s.height === 0; }
  function lengthOf(s) { return Math.sqrt(s.width * s.width + s.height * s.height); }
  // New width and height that give a line this length, keeping its direction.
  function lineSize(s, len) { var now = lengthOf(s); return [s.width * len / now, s.height * len / now]; }
  // The frame a resize would give, without touching the shape.
  function frameFor(s, nw, nh) { var b = { left: s.left, top: s.top, width: s.width, height: s.height }; applySize(b, nw, nh); return b; }
  function sameShape(w1, h1, w2, h2) { return Math.abs(w1 * h2 - w2 * h1) <= 0.002 * Math.max(w1 * h2, w2 * h1); }

  // Runs after plain resizes have synced: crops pictures whose frame changes shape,
  // or keeps them in proportion when cropping is off or unavailable.
  // plans: [{ shape, nw, nh }] in points. Resolves { cropped, kept, missed, note }.
  function finishPictures(context, plans) {
    var out = { resized: 0, cropped: 0, kept: 0, missed: [], note: "" };
    if (!plans.length) return Promise.resolve(out);
    var jobs = [];
    var crop = ST.state.cropPictures !== false && ST.crop.canCrop();
    if (ST.state.cropPictures !== false && !ST.crop.canCrop()) out.note = " This version of PowerPoint can't crop pictures, so they kept their proportions.";
    plans.forEach(function (p) {
      var s = p.shape, nw = p.nw === null ? s.width : p.nw, nh = p.nh === null ? s.height : p.nh;
      if (sameShape(s.width, s.height, nw, nh)) { applySize(s, nw, nh); out.resized++; return; }
      if (crop) { jobs.push({ shape: s, frame: frameFor(s, nw, nh), anchor: anchorFactors(), nw: nw, nh: nh, typedW: p.nw !== null }); return; }
      // Keep proportions, led by the box that was typed
      if (p.nw !== null) nh = s.height * nw / s.width; else nw = s.width * nh / s.height;
      applySize(s, nw, nh); out.kept++;
    });
    return context.sync().then(function () {
      if (!jobs.length) return out;
      return ST.confirmSwap("crop pictures").then(function (ok) {
        if (ok) return ST.crop.apply(context, jobs).then(function (r) { out.cropped = r.cropped; out.missed = r.missed; return out; });
        // Not swapping the slide: keep the pictures in proportion instead
        jobs.forEach(function (j) { var sh = j.shape, w = j.nw, h = j.nh; if (j.typedW) h = sh.height * w / sh.width; else w = sh.width * h / sh.height; applySize(sh, w, h); out.kept++; });
        return context.sync().then(function () { return out; });
      });
    });
  }
  function pictureMessage(r) {
    return (r.cropped ? " Cropped " + ST.plural(r.cropped, "picture") + " to fit." : "") +
      (r.kept ? " " + ST.plural(r.kept, "picture") + " kept " + (r.kept === 1 ? "its" : "their") + " proportions." : "") +
      (r.missed.length ? " Couldn't find " + r.missed.join(", ") + " in the slide, so it wasn't cropped." : "") + r.note;
  }
  function failText(e) {
    return "Couldn't crop the pictures. " + (ST.errorText ? ST.errorText(e) : "");
  }

  function updateCalc() {
    var parts = [], bad = false, sel = ST.selection;
    [["resW", "Width", "w"], ["resH", "Height", "h"]].forEach(function (f) {
      var text = ST.state[f[0]];
      if (!text || !String(text).trim()) return;
      var line = sel.count && sel.line ? Math.sqrt(sel.w * sel.w + sel.h * sel.h) : 0;
      var v = ST.evaluate(text, line ? ST.formulaVars(line, line) : ST.formulaVars(sel.count ? sel.w : undefined, sel.count ? sel.h : undefined), f[2]);
      if (v === null || v <= 0) { parts.push(f[1] + ": check the formula"); bad = true; }
      else parts.push((line ? "Length" : f[1]) + " = " + ST.nice(v) + " in");
    });
    var calc = $("calc");
    calc.textContent = parts.join("     ");
    // A problem shows at every level, so the reason is visible before clicking Resize
    if (bad) calc.removeAttribute("data-min"); else calc.setAttribute("data-min", "2");
    calc.className = bad ? "calc bad" : "calc";
  }

  // Reference point grids. Size and Copies each show one, and both set the same value.
  var anchorGrids = [];
  ST.anchorGrid = function (gridId, nameId) {
    var el = $(gridId);
    ANCHOR_NAMES.forEach(function (name, i) {
      var b = document.createElement("button");
      b.title = "Use the " + name + " as the reference point";
      b.setAttribute("aria-label", b.title);
      b.addEventListener("click", function () { ST.state.anchor = i; ST.changed(); renderAnchor(); });
      el.appendChild(b);
    });
    anchorGrids.push([gridId, nameId]);
    renderAnchor();
  };
  function renderAnchor() {
    anchorGrids.forEach(function (g) {
      Array.prototype.forEach.call($(g[0]).children, function (b, i) { b.classList.toggle("on", i === (ST.state.anchor || 0)); });
      $(g[1]).textContent = ANCHOR_NAMES[ST.state.anchor || 0];
    });
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
    var sw = ST.selection.w, sh = ST.selection.h;
    if (ST.selection.line) { ST.state.resW = ST.nice(Math.sqrt(sw * sw + sh * sh)); ST.state.resH = ""; } // a line: its length
    else { ST.state.resW = ST.nice(sw); ST.state.resH = ST.nice(sh); }
    $("resW").value = ST.state.resW; $("resH").value = ST.state.resH;
    ST.changed();
  }

  // over: { w, h } formula texts to use instead of the boxes (Repeat passes the original ones).
  function resizeSelection(doW, doH, over) {
    if (!ST.needOffice()) return;
    ST.calculate();
    var lock = !!ST.state.lockAspect;
    var wText = String(over ? over.w : ST.state.resW || "").trim(), hText = String(over ? over.h : ST.state.resH || "").trim();
    doW = doW && !!wText;
    doH = doH && !!hText;
    if (lock && doW && doH) doH = false; // width leads when the lock is on
    if (!doW && !doH) { ST.toast("Type a width or height first.", true); return; }
    var lineText = doW ? wText : hText; // a line's length comes from whichever box is used
    var rw = doW ? wText : "", rh = doH ? hText : "";
    ST.setRepeat("Resize" + (rw ? " W " + rw : "") + (rh ? " H " + rh : ""), function () { resizeSelection(!!rw, !!rh, { w: rw, h: rh }); });
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      return ST.crop.pictureFlags(context, items).then(function (isPic) {
        var updates = [], pics = [], lines = 0, skipped = 0, firstW = null, firstH = null;
        for (var i = 0; i < items.length; i++) {
          var s = items[i];
          if (isLine(s) && !isPic[i]) {
            var len = lengthOf(s);
            if (!(len > 0)) { skipped++; continue; }
            var L = ST.evaluate(lineText, ST.formulaVars(len / PT, len / PT), "w");
            if (!(L > 0)) { ST.toast("That formula couldn't be worked out.", true); return; }
            var ls = lineSize(s, L * PT);
            updates.push([s, ls[0], ls[1]]); lines++;
            if (firstW === null) firstW = L;
            continue;
          }
          var vars = ST.formulaVars(s.width / PT, s.height / PT);
          var nw = doW ? ST.evaluate(wText, vars, "w") : null;
          var nh = doH ? ST.evaluate(hText, vars, "h") : null;
          if ((doW && !(nw > 0)) || (doH && !(nh > 0))) { ST.toast("That formula couldn't be worked out.", true); return; }
          if (nw !== null) nw *= PT;
          if (nh !== null) nh *= PT;
          if (lock && nw !== null) nh = s.height * nw / s.width;
          if (lock && nh !== null && nw === null) nw = s.width * nh / s.height;
          if (firstW === null && firstH === null) { firstW = doW ? nw / PT : null; firstH = doH ? nh / PT : null; }
          if (isPic[i]) pics.push({ shape: s, nw: nw, nh: nh });
          else updates.push([s, nw, nh]);
        }
        if (!updates.length && !pics.length) { ST.toast(skipped ? "Those lines have no length to scale." : "Select the shapes on your slide first.", true); return; }
        updates.forEach(function (u) { applySize(u[0], u[1], u[2]); });
        return context.sync().then(function () {
          return finishPictures(context, pics).catch(function (e) { return { resized: 0, cropped: 0, kept: 0, missed: [], note: " " + failText(e), failed: true }; });
        }).then(function (r) {
          remember(doW ? wText : "", doH ? hText : "");
          if (doW && firstW !== null) { ST.state.resW = ST.nice(firstW); $("resW").value = ST.state.resW; }
          if (doH && firstH !== null) { ST.state.resH = ST.nice(firstH); $("resH").value = ST.state.resH; }
          ST.changed();
          var plain = updates.length + r.resized;
          var msg = (plain ? "Resized " + ST.plural(plain, "shape") + (lines ? " (" + ST.plural(lines, "line") + " by length)" : "") + "." : "") + pictureMessage(r);
          if (skipped) msg += " Skipped " + ST.plural(skipped, "line") + " with no length.";
          ST.toast(msg.trim(), !!r.failed);
          ST.refreshSelection();
        });
      });
    });
  }

  function matchSize(largest) {
    ST.setRepeat("Match size to " + (largest ? "largest" : "smallest"), function () { matchSize(largest); });
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      if (items.length < 2) { ST.toast("Select at least two shapes to match.", true); return; }
      return ST.crop.pictureFlags(context, items).then(function (isPic) {
        var lineItems = [], others = [];
        items.forEach(function (s, i) { if (isLine(s) && !isPic[i]) { if (lengthOf(s) > 0) lineItems.push(s); } else others.push({ s: s, pic: isPic[i] }); });
        if (lineItems.length < 2 && others.length < 2) { ST.toast("Select at least two shapes, or two lines, to match.", true); return; }
        var parts = [], pics = [], changed = 0;
        if (lineItems.length >= 2) {
          var lt = lineItems.reduce(function (best, s) { return (largest ? lengthOf(s) > lengthOf(best) : lengthOf(s) < lengthOf(best)) ? s : best; });
          var L = lengthOf(lt), n = 0;
          lineItems.forEach(function (s) { if (s === lt) return; var ls = lineSize(s, L); applySize(s, ls[0], ls[1]); n++; });
          parts.push("Matched " + ST.plural(n, "line") + " to " + ST.nice(L / PT) + " in long.");
        }
        var tw, th;
        if (others.length >= 2) {
          var target = others.reduce(function (best, o) {
            var a = o.s.width * o.s.height, b = best.s.width * best.s.height;
            return (largest ? a > b : a < b) ? o : best;
          }).s;
          tw = target.width; th = target.height;
          others.forEach(function (o) {
            if (o.s === target) return;
            if (o.pic) pics.push({ shape: o.s, nw: tw, nh: th }); else { applySize(o.s, tw, th); changed++; }
          });
        }
        return context.sync().then(function () {
          return finishPictures(context, pics).catch(function (e) { return { resized: 0, cropped: 0, kept: 0, missed: [], note: " " + failText(e), failed: true }; });
        }).then(function (r) {
          changed += r.resized + r.kept;
          if (others.length >= 2 && (changed || !pics.length)) parts.unshift("Matched " + ST.plural(changed, "shape") + " to " + ST.nice(tw / PT) + " x " + ST.nice(th / PT) + " in.");
          ST.toast((parts.join(" ") + pictureMessage(r)).trim(), !!r.failed);
          ST.refreshSelection();
        });
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
    ST.anchorGrid("anchor", "anchorName");
    renderRecent();
    ST.on("change", updateCalc);
    ST.on("selection", updateCalc);
    ST.on("reset", function () {
      lock.checked = !!ST.state.lockAspect;
      renderAnchor();
    });
  }

  ST.registerTool({
    id: "resize",
    title: "Size",
    icon: ICON,
    levels: [1, 2, 3],
    defaultLevel: 2,
    defaults: defaults,
    tip: { what: "Sets exact sizes with numbers or formulas, and matches sizes.", steps: ["Select shapes.", "Type a width or height, such as 2col+gap, and press Enter.", "Or click Largest or Smallest to match sizes."] },
    build: build
  });
})();
