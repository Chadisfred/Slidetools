/* Animation tool (id "animate"): custom entrances, exits and emphasis with GSAP-style eases,
   combined effects, exact numbers, a stagger, starter and saved presets, and Capture.
   PowerPoint add-ins have no animation API, so effects are written to the slide's timing XML
   through the shared slide rebuild (js/rebuild.js), in memory. The XML and the eases live in
   js/animcore.js. Presets are saved in the presentation as tags, like shape styles. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$, R = ST.rebuild, A = ST.anim;
  var STORE_KEY = "SLIDETOOLS_ANIMPRESETS", PART_SIZE = 2000;
  var SHAPE_KINDS = ["sp", "pic", "grpSp", "cxnSp", "graphicFrame"];
  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="5" width="7" height="8" rx="1.5"/><path d="M2 6h4M3 9h4M2 12h4"/></svg>';

  var defaults = {
    anKind: "entr", anFade: true, anMove: true, anDir: "up", anDist: "20", anScale: false, anScaleVal: "90", anRotate: false, anRotVal: "-12",
    anPulse: true, anPulseVal: "106", anSpin: false, anSpinVal: "360", anDim: false, anDimVal: "40",
    anEase: "power2.out", anDur: "0.5", anDelay: "0", anStart: "click", anStagger: "0.08", anOrder: "selection", anReplace: true, anPresetName: ""
  };
  ST.addDefaults(defaults);
  var KEYS = Object.keys(defaults).filter(function (k) { return k !== "anPresetName" && k !== "anReplace" && k !== "anOrder"; });

  // Built in starters, read only. Each one fills the controls.
  function S(o) { return Object.assign({ anKind: "entr", anFade: false, anMove: false, anScale: false, anRotate: false, anPulse: false, anSpin: false, anDim: false, anDelay: "0" }, o); }
  var STARTERS = [
    { name: "Fade up", s: S({ anFade: true, anMove: true, anDir: "up", anDist: "20", anEase: "power2.out", anDur: "0.5" }) },
    { name: "Pop", s: S({ anFade: true, anScale: true, anScaleVal: "80", anEase: "back.out(1.7)", anDur: "0.45" }) },
    { name: "Slide in", s: S({ anFade: true, anMove: true, anDir: "right", anDist: "2g", anEase: "expo.out", anDur: "0.6" }) },
    { name: "Drop in", s: S({ anMove: true, anDir: "down", anDist: "30", anEase: "bounce.out", anDur: "0.8" }) },
    { name: "Zoom in", s: S({ anFade: true, anScale: true, anScaleVal: "50", anEase: "power3.out", anDur: "0.5" }) },
    { name: "Swing in", s: S({ anFade: true, anRotate: true, anRotVal: "-12", anEase: "back.out(1.7)", anDur: "0.6" }) },
    { name: "Elastic grow", s: S({ anScale: true, anScaleVal: "60", anEase: "elastic.out(1, 0.4)", anDur: "1" }) },
    { name: "Rise out", s: S({ anKind: "exit", anFade: true, anMove: true, anDir: "up", anDist: "20", anEase: "power2.in", anDur: "0.4" }) },
    { name: "Shrink out", s: S({ anKind: "exit", anFade: true, anScale: true, anScaleVal: "80", anEase: "power2.in", anDur: "0.4" }) },
    { name: "Pulse", s: S({ anKind: "emph", anPulse: true, anPulseVal: "106", anEase: "sine.inOut", anDur: "0.6" }) },
    { name: "Gentle spin", s: S({ anKind: "emph", anSpin: true, anSpinVal: "360", anEase: "power2.inOut", anDur: "1" }) }
  ];
  var FAMILIES = [["none", "None"], ["power1", "Power 1"], ["power2", "Power 2"], ["power3", "Power 3"], ["power4", "Power 4"], ["sine", "Sine"], ["expo", "Expo"], ["circ", "Circ"], ["back", "Back"], ["elastic", "Elastic"], ["bounce", "Bounce"]];

  var presets = [], presetsLoaded = false, captured = null; // captured: effects of the last applied captured preset

  // ---------- Settings to an effect ----------
  function secs(text) { var v = ST.evaluate(String(text || "").trim(), null, null); return v === null || !isFinite(v) ? null : v; }
  function plainNum(text) { var t = String(text || "").trim().replace(/\s*(%|°|deg)$/i, ""); if (!t) return null; var v = ST.evaluate(t, null, null); return v === null || !isFinite(v) ? null : v; }
  var TRAVEL = { up: [0, -1], down: [0, 1], left: [-1, 0], right: [1, 0] };
  // Returns { spec } or { bad }. size: the slide size in points.
  function specFrom(St, size) {
    size = size || { w: 960, h: 540 };
    var kind = St.anKind === "exit" ? "exit" : St.anKind === "emph" ? "emph" : "entr";
    var dur = secs(St.anDur), delay = secs(St.anDelay);
    if (!(dur > 0 && dur <= 60)) return { bad: "Duration must be more than 0 and up to 60 seconds." };
    if (delay === null || delay < 0) return { bad: "Delay must be 0 or more seconds." };
    if (!A.parseEase(St.anEase)) return { bad: "Ease " + St.anEase + " isn't one it knows. Try power2.out, back.out(1.7) or bounce.out." };
    var spec = { kind: kind, ease: A.easeText(A.readEase(St.anEase)), dur: Math.round(dur * 1000), delay: Math.round(delay * 1000) };
    if (kind === "emph") {
      if (St.anPulse) { var p = plainNum(St.anPulseVal); if (!(p > 0 && p <= 1000)) return { bad: "Pulse must be a percentage, such as 106." }; spec.pulse = p / 100; }
      if (St.anSpin) { var sp = plainNum(St.anSpinVal); if (sp === null) return { bad: "Spin must be a number of degrees, such as 360." }; spec.spin = sp; }
      if (St.anDim) { var d = plainNum(St.anDimVal); if (!(d >= 0 && d <= 100)) return { bad: "Dim must be a percentage from 0 to 100." }; spec.dim = d / 100; }
      if (!spec.pulse && !spec.spin && spec.dim === undefined) return { bad: "Turn on Pulse, Spin or Dim first." };
      return { spec: spec };
    }
    spec.fade = !!St.anFade;
    if (St.anMove) {
      var dist = ST.evaluatePt(St.anDist);
      if (dist === null || !isFinite(dist)) return { bad: "Distance " + (ST.ptProblem(St.anDist) || "can't be read") + ". Use points or names such as g." };
      var v = TRAVEL[St.anDir] || TRAVEL.up, sign = kind === "exit" ? 1 : -1; // an entrance starts back along its travel
      spec.movePt = { dx: v[0] * dist * sign, dy: v[1] * dist * sign };
      spec.move = { dx: spec.movePt.dx / size.w, dy: spec.movePt.dy / size.h };
    }
    if (St.anScale) { var s = plainNum(St.anScaleVal); if (!(s > 0 && s <= 1000)) return { bad: "Scale must be a percentage, such as 90." }; spec.scale = s / 100; }
    if (St.anRotate) { var r = plainNum(St.anRotVal); if (r === null) return { bad: "Rotate must be a number of degrees, such as -12." }; spec.rotate = r; }
    if (!spec.fade && !spec.move && !spec.scale && !spec.rotate) return { bad: "Turn on Fade, Move, Scale or Rotate first." };
    return { spec: spec };
  }
  function snapshot() { var o = {}; KEYS.forEach(function (k) { o[k] = ST.state[k]; }); return o; }
  function describe(St) {
    var kind = { entr: "Entrance", exit: "Exit", emph: "Emphasis" }[St.anKind] || "Entrance", parts = [];
    if (St.anKind === "emph") { if (St.anPulse) parts.push("pulse"); if (St.anSpin) parts.push("spin"); if (St.anDim) parts.push("dim"); }
    else { if (St.anFade) parts.push("fade"); if (St.anMove) parts.push("move " + St.anDir); if (St.anScale) parts.push("scale"); if (St.anRotate) parts.push("rotate"); }
    return kind + ": " + parts.join(", ");
  }

  // ---------- Applying ----------
  function orderItems(items, order) {
    var list = items.slice();
    if (order === "lr") list.sort(function (a, b) { return a.left - b.left || a.top - b.top; });
    else if (order === "tb") list.sort(function (a, b) { return a.top - b.top || a.left - b.left; });
    return list;
  }
  function boxesOf(items) { return items.map(function (s) { return { left: s.left, top: s.top, width: s.width, height: s.height, name: s.name, id: s.id }; }); }
  function jobsFor(doc, items, St, firstDelay) {
    var els = R.matchShapes(doc, boxesOf(items), SHAPE_KINDS), stagger = Math.max(0, secs(St.anStagger === undefined ? ST.state.anStagger : St.anStagger) || 0) * 1000, jobs = [], k = 0;
    els.forEach(function (el) {
      if (!el) return;
      var c = R.cNvPr(el), box = R.readBox(el);
      if (!c) return;
      jobs.push({ spid: c.getAttribute("id"), rot0: box ? box.rot : 0, hasText: el.localName === "sp" && !!R.kid(el, A.NS_P, "txBody"),
                  start: k ? "with" : (St.anStart || "click"), delay: Math.round(firstDelay + k * stagger) });
      k++;
    });
    return jobs;
  }
  // preset: a saved or starter preset, or null for the panel settings
  function apply(preset) {
    var St = preset && preset.s ? Object.assign(snapshot(), preset.s) : snapshot();
    var cap = preset && preset.captured ? preset.captured : null;
    var order = ST.state.anOrder, replace = ST.state.anReplace !== false;
    St.anStagger = ST.state.anStagger; St.anStart = cap ? ST.state.anStart : St.anStart;
    var check = cap ? { spec: null } : specFrom(St);
    if (check.bad) { ST.toast(check.bad, true); return; }
    if (!ST.needOffice()) return;
    if (!R.supported()) { ST.toast("This version of PowerPoint can't change animations. It needs PowerPoint API 1.8.", true); return; }
    ST.setRepeat(cap ? "Animate: " + preset.name : describe(St), function () { apply(preset || { s: St }); });
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      items = orderItems(items, order);
      return ST.confirmSwap("add animations").then(function (ok) {
        if (!ok) return null;
        return R.edit(context, function (doc, extra) {
          var size = A.slideSize(extra && extra.files), r = cap ? { spec: null } : specFrom(St, size);
          var firstDelay = cap ? (secs(ST.state.anDelay) || 0) * 1000 : r.spec.delay;
          var jobs = jobsFor(doc, items, St, firstDelay);
          if (!jobs.length) return { count: 0 };
          if (replace) A.removeFor(doc, jobs.map(function (j) { return j.spid; }));
          var n = cap ? A.addCaptured(doc, cap, jobs) : A.addEffects(doc, r.spec, jobs);
          return { count: n ? jobs.length : 0 };
        });
      }).then(function (res) {
        if (res === null) { ST.toast("Nothing was changed."); return; }
        ST.toast(res.count ? "Animated " + ST.plural(res.count, "shape") + (res.count > 1 && secs(ST.state.anStagger) > 0 ? ", " + ST.nice(secs(ST.state.anStagger)) + " s apart" : "") + "." : "Couldn't find the selected shapes in the slide.", !res.count);
        ST.refreshSelection();
      }, function (e) { ST.toast("Couldn't change the slide. " + (ST.errorText ? ST.errorText(e) : ""), true); });
    });
  }
  function removeFromSelection() {
    if (!ST.needOffice()) return;
    if (!R.supported()) { ST.toast("This version of PowerPoint can't change animations. It needs PowerPoint API 1.8.", true); return; }
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      return ST.confirmSwap("remove animations").then(function (ok) {
        if (!ok) return null;
        var gone = 0;
        return R.edit(context, function (doc) {
          var els = R.matchShapes(doc, boxesOf(items), SHAPE_KINDS), ids = els.filter(Boolean).map(function (el) { var c = R.cNvPr(el); return c && c.getAttribute("id"); }).filter(Boolean);
          gone = ids.length ? A.removeFor(doc, ids) : 0;
          return { count: gone };
        }).then(function () { return gone; });
      }).then(function (n) {
        if (n === null) { ST.toast("Nothing was changed."); return; }
        ST.toast(n ? "Removed " + ST.plural(n, "animation") + "." : "The selected shapes have no animations of their own.", !n);
      }, function (e) { ST.toast("Couldn't change the slide. " + (ST.errorText ? ST.errorText(e) : ""), true); });
    });
  }

  // ---------- Capture ----------
  function captureSelection() {
    if (!ST.needOffice()) return;
    if (!R.supported()) { ST.toast("Capture needs PowerPoint API 1.8.", true); return; }
    var name = String(ST.state.anPresetName || "").trim();
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      var found = null;
      return R.edit(context, function (doc) {
        var el = R.matchShapes(doc, boxesOf([items[0]]), SHAPE_KINDS)[0], c = el && R.cNvPr(el);
        if (c) found = A.capture(doc, c.getAttribute("id"));
        return { count: 0 }; // read only: nothing is swapped
      }).then(function () {
        if (!found || !found.length) { ST.toast("The selected shape has no animations of its own to capture.", true); return; }
        var p = { id: "u" + Date.now(), name: name || (items[0].name || "Captured") , captured: found };
        presets.push(p); savePresets();
        ST.state.anPresetName = ""; $("anPresetName").value = ""; ST.save();
        ST.toast("Captured " + ST.plural(found.length, "animation") + " as " + p.name + ".");
      });
    }, "Select a shape with animations first.");
  }

  // ---------- Presets in the presentation ----------
  function loadPresets() {
    if (!ST.inOffice || !ST.supports("1.3")) { presetsLoaded = true; renderPresets(); return; }
    PowerPoint.run(function (context) {
      var tags = context.presentation.tags, count = tags.getItemOrNullObject(STORE_KEY + "_N");
      count.load("value");
      return context.sync().then(function () {
        var n = count.isNullObject ? 0 : parseInt(count.value, 10) || 0, parts = [];
        for (var i = 0; i < n; i++) { var t = tags.getItemOrNullObject(STORE_KEY + "_" + i); t.load("value"); parts.push(t); }
        return context.sync().then(function () {
          var text = parts.map(function (t) { return t.isNullObject ? "" : t.value; }).join("");
          presets = text ? JSON.parse(text) : [];
        });
      });
    }).catch(function () { presets = []; }).then(function () { presetsLoaded = true; renderPresets(); });
  }
  function savePresets() {
    renderPresets();
    if (!ST.inOffice || !ST.supports("1.3")) return Promise.resolve();
    var text = JSON.stringify(presets), parts = [];
    for (var i = 0; i < text.length; i += PART_SIZE) parts.push(text.slice(i, i + PART_SIZE));
    return PowerPoint.run(function (context) {
      var tags = context.presentation.tags, old = tags.getItemOrNullObject(STORE_KEY + "_N");
      old.load("value");
      return context.sync().then(function () {
        var oldN = old.isNullObject ? 0 : parseInt(old.value, 10) || 0;
        parts.forEach(function (p, i) { tags.add(STORE_KEY + "_" + i, p); });
        for (var j = parts.length; j < oldN; j++) tags.delete(STORE_KEY + "_" + j);
        tags.add(STORE_KEY + "_N", String(parts.length));
        return context.sync();
      });
    }).catch(function (e) { ST.toast("Couldn't save animation presets: " + ST.reason(e), true); });
  }
  function savePreset() {
    var name = String(ST.state.anPresetName || "").trim();
    if (!name) { ST.toast("Type a name for the preset first.", true); return; }
    var check = specFrom(snapshot());
    if (check.bad) { ST.toast(check.bad, true); return; }
    presets = presets.filter(function (p) { return p.name.toLowerCase() !== name.toLowerCase(); });
    presets.push({ id: "u" + Date.now(), name: name, s: snapshot() });
    ST.state.anPresetName = ""; $("anPresetName").value = ""; ST.save();
    savePresets(); ST.toast("Saved " + name + ".");
  }
  // Clicking a preset fills the controls (a captured one is applied as it is)
  function usePreset(p) {
    if (p.captured) { apply(p); return; }
    Object.keys(p.s).forEach(function (k) { if (k in defaults) ST.state[k] = p.s[k]; });
    ST.save(); ST.emit("reset"); ST.changed(); play();
  }
  function renderPresets() {
    var el = $("anPresets");
    if (!el) return;
    el.innerHTML = "";
    STARTERS.concat(presets).forEach(function (p) {
      var b = document.createElement("button");
      b.className = "mini-chip an-chip" + (p.id ? " mine" : "");
      b.textContent = p.name;
      b.title = p.captured ? "Apply this captured animation to the selected shapes" : p.id ? "Your preset: fill the controls" : "Fill the controls with " + p.name;
      b.addEventListener("click", function () { usePreset(p); });
      if (p.id) {
        var del = document.createElement("span");
        del.className = "del"; del.textContent = "×"; del.title = "Delete this preset";
        del.addEventListener("click", function (e) { e.stopPropagation(); presets = presets.filter(function (x) { return x.id !== p.id; }); savePresets(); });
        b.appendChild(del);
      }
      el.appendChild(b);
    });
  }

  // ---------- Preview ----------
  var PX_PER_PT = 0.6, running = null;
  function play() {
    var box = $("anSample"), note = $("anPreviewNote");
    if (!box || !box.animate) return;
    var r = specFrom(snapshot());
    if (running) { try { running.cancel(); } catch (e) { /* already done */ } running = null; }
    if (r.bad) { note.textContent = r.bad; note.classList.add("bad"); return; }
    note.textContent = describe(snapshot()) + ", " + ST.nice(r.spec.dur / 1000) + " s, " + r.spec.ease; note.classList.remove("bad");
    var fs = A.frames(r.spec);
    var kf = fs.map(function (f) {
      return { offset: f.t, opacity: f.opacity, transform: "translate(" + (f.dxPt * PX_PER_PT).toFixed(2) + "px," + (f.dyPt * PX_PER_PT).toFixed(2) + "px) scale(" + f.scale.toFixed(4) + ") rotate(" + f.rot.toFixed(2) + "deg)" };
    });
    running = box.animate(kf, { duration: r.spec.dur, delay: 150, fill: "both", easing: "linear" });
    // An exit leaves the sample hidden, so bring it back after a moment
    if (r.spec.kind === "exit") running.onfinish = function () { setTimeout(function () { if (running) { running.cancel(); running = null; } }, 600); };
    drawCurve(r.spec.ease);
  }
  function drawCurve(text) {
    var svg = $("anCurve"), f = A.parseEase(text);
    if (!svg || !f) return;
    var W = 120, H = 64, top = 14, bottom = 10, pts = [];
    for (var i = 0; i <= 60; i++) { var t = i / 60, y = f(t); pts.push((t * W).toFixed(1) + "," + (H - bottom - y * (H - top - bottom)).toFixed(1)); }
    svg.innerHTML = '<line x1="0" y1="' + (H - bottom) + '" x2="' + W + '" y2="' + (H - bottom) + '" class="an-axis"/><line x1="0" y1="' + top + '" x2="' + W + '" y2="' + top + '" class="an-axis"/>' +
      '<polyline points="' + pts.join(" ") + '" class="an-line"/>';
  }

  // ---------- Panel ----------
  function seg(id, opts) { return '<div class="seg seg' + opts.length + '" id="' + id + '">' + opts.map(function (o) { return '<button data-v="' + o[0] + '"' + (o[2] ? ' title="' + o[2] + '"' : '') + '>' + o[1] + '</button>'; }).join("") + '</div>'; }
  function toggle(id, label, field) { return '<div class="an-effect"><div class="toggle">' + label + '<label class="switch"><input type="checkbox" id="' + id + '"><span></span></label></div>' + (field || "") + '</div>'; }
  var MARKUP =
    '<div class="sublabel">Presets</div>' +
    '<div class="mini-chips" id="anPresets"></div>' +
    seg("anKind", [["entr", "Entrance"], ["exit", "Exit"], ["emph", "Emphasis"]]) +
    '<div class="an-preview"><div class="an-stage"><div class="an-sample" id="anSample"></div></div><button class="btn small" id="anPlay">Play</button></div>' +
    '<div class="help an-note" id="anPreviewNote"></div>' +
    '<div class="an-move-group">' +
      toggle("anFade", "Fade") +
      toggle("anMove", "Move", '<div class="row an-sub"><label class="field">Travel' + seg("anDir", [["up", "&uarr;", "Up"], ["down", "&darr;", "Down"], ["left", "&larr;", "Left"], ["right", "&rarr;", "Right"]]) + '</label><label class="field"><span class="label-row">Distance (pt)' + ST.info("Points, or grid names in points such as <code>g</code> or <code>g/2</code>.", "About distance") + '</span><input type="text" id="anDist" inputmode="decimal"></label></div>') +
      toggle("anScale", "Scale", '<label class="field an-sub">' + '<span id="anScaleLabel">From %</span><input type="text" id="anScaleVal" inputmode="decimal"></label>') +
      toggle("anRotate", "Rotate", '<label class="field an-sub"><span id="anRotLabel">From degrees</span><input type="text" id="anRotVal" inputmode="decimal"></label>') +
    '</div>' +
    '<div class="an-emph-group">' +
      toggle("anPulse", "Pulse", '<label class="field an-sub">Grow to %<input type="text" id="anPulseVal" inputmode="decimal"></label>') +
      toggle("anSpin", "Spin", '<label class="field an-sub">Degrees<input type="text" id="anSpinVal" inputmode="decimal"></label>') +
      toggle("anDim", "Dim", '<label class="field an-sub">Fade to %<input type="text" id="anDimVal" inputmode="decimal"></label>') +
    '</div>' +
    '<div class="group-label label-row">Ease' + ST.info("Named the way GSAP names them. Power 1 to 4 get stronger, Back overshoots, Elastic springs and Bounce bounces. In starts slowly, Out ends slowly. You can also type a name such as <code>back.out(1.7)</code>.", "About eases") + '</div>' +
    '<div class="an-ease"><div class="an-ease-controls"><select id="anFamily" aria-label="Ease">' + FAMILIES.map(function (f) { return '<option value="' + f[0] + '">' + f[1] + '</option>'; }).join("") + '</select>' +
      seg("anEaseDir", [["in", "In"], ["out", "Out"], ["inOut", "In out"]]) +
      '<input type="text" id="anEase" spellcheck="false" aria-label="Ease as text"></div><svg id="anCurve" viewBox="0 0 120 64" aria-hidden="true"></svg></div>' +
    '<div class="row"><label class="field">Duration (s)<input type="text" id="anDur" inputmode="decimal"></label><label class="field">Delay (s)<input type="text" id="anDelay" inputmode="decimal"></label></div>' +
    '<div class="field">Start' + seg("anStart", [["click", "On click"], ["with", "With previous"], ["after", "After previous"]]) + '</div>' +
    '<div class="row" data-min="2"><label class="field"><span class="label-row">Stagger (s)' + ST.info("With several shapes selected, each one starts this much after the one before.", "About stagger") + '</span><input type="text" id="anStagger" inputmode="decimal"></label>' +
      '<label class="field">Order' + seg("anOrder", [["selection", "Picked", "In the order PowerPoint gives the selection"], ["lr", "&rarr;", "Left to right"], ["tb", "&darr;", "Top to bottom"]]) + '</label></div>' +
    '<div class="toggle" data-min="2"><span class="label-row">Replace their animations' + ST.info("Removes the shapes\' earlier animations first, so applying again doesn\'t add another.", "About replacing") + '</span><label class="switch"><input type="checkbox" id="anReplace"><span></span></label></div>' +
    '<button class="btn primary" id="anApply">Apply to selection</button>' +
    '<div class="inline" data-min="2"><label class="field">Preset name<input type="text" id="anPresetName" maxlength="32" placeholder="e.g. Card in"></label>' +
      '<div class="pair"><button class="btn small" id="anSave" title="Save these settings as a preset in this presentation">Save</button><button class="btn small" id="anCapture" title="Save the selected shape\'s own animations as a preset">Capture</button></div></div>' +
    '<button class="btn small" id="anRemove" data-min="2" title="Remove the selected shapes\' animations">Remove animations from selection</button>';

  function renderSegs() {
    [["anKind", "anKind"], ["anDir", "anDir"], ["anStart", "anStart"], ["anOrder", "anOrder"]].forEach(function (p) {
      Array.prototype.forEach.call($(p[0]).children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === ST.state[p[1]]); });
    });
    var e = A.readEase(ST.state.anEase) || { family: "power2", dir: "out" };
    $("anFamily").value = e.family === "power0" ? "none" : e.family;
    Array.prototype.forEach.call($("anEaseDir").children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === e.dir); b.disabled = e.family === "none"; });
    var emph = ST.state.anKind === "emph", exit = ST.state.anKind === "exit";
    document.querySelector(".an-move-group").hidden = emph;
    document.querySelector(".an-emph-group").hidden = !emph;
    $("anScaleLabel").textContent = exit ? "To %" : "From %";
    $("anRotLabel").textContent = exit ? "To degrees" : "From degrees";
    ["anFade", "anMove", "anScale", "anRotate", "anPulse", "anSpin", "anDim", "anReplace"].forEach(function (k) {
      var box = $(k); box.checked = k === "anReplace" ? ST.state[k] !== false : !!ST.state[k];
      var sub = box.closest(".an-effect") && box.closest(".an-effect").querySelector(".an-sub");
      if (sub) sub.hidden = !box.checked;
    });
    $("anEase").classList.toggle("err", !A.parseEase(ST.state.anEase));
  }
  function setEaseFromControls(dir) {
    var fam = $("anFamily").value, cur = A.readEase(ST.state.anEase) || { dir: "out", params: [] };
    var e = { family: fam, dir: dir || cur.dir || "out", params: fam === cur.family ? cur.params : (fam === "back" ? [1.7] : fam === "elastic" ? [1, 0.3] : []) };
    ST.state.anEase = A.easeText(e); $("anEase").value = ST.state.anEase;
    ST.changed(); renderSegs(); play();
  }
  function build(body) {
    body.appendChild(ST.html(MARKUP));
    ["anDist", "anScaleVal", "anRotVal", "anPulseVal", "anSpinVal", "anDimVal", "anDur", "anDelay", "anStagger", "anPresetName"].forEach(function (k) {
      ST.bindField($(k), k, { onInput: function () { if (k !== "anPresetName") queuePlay(); } });
    });
    ST.bindField($("anEase"), "anEase", { onInput: function () { renderSegs(); queuePlay(); } });
    [["anKind", "anKind"], ["anDir", "anDir"], ["anStart", "anStart"], ["anOrder", "anOrder"]].forEach(function (p) {
      Array.prototype.forEach.call($(p[0]).children, function (b) {
        b.addEventListener("click", function (e) { e.preventDefault(); ST.state[p[1]] = b.getAttribute("data-v"); ST.save(); renderSegs(); ST.changed(); if (p[0] !== "anOrder" && p[0] !== "anStart") play(); });
      });
    });
    ["anFade", "anMove", "anScale", "anRotate", "anPulse", "anSpin", "anDim", "anReplace"].forEach(function (k) {
      $(k).addEventListener("change", function () { ST.state[k] = $(k).checked; ST.save(); renderSegs(); ST.changed(); if (k !== "anReplace") play(); });
    });
    $("anFamily").addEventListener("change", function () { setEaseFromControls(); });
    Array.prototype.forEach.call($("anEaseDir").children, function (b) { b.addEventListener("click", function (e) { e.preventDefault(); setEaseFromControls(b.getAttribute("data-v")); }); });
    $("anPlay").addEventListener("click", play);
    $("anApply").addEventListener("click", function () { apply(null); });
    $("anSave").addEventListener("click", savePreset);
    $("anCapture").addEventListener("click", captureSelection);
    $("anRemove").addEventListener("click", removeFromSelection);
    $("anPresetName").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); savePreset(); } });
    renderSegs(); renderPresets(); drawCurve(ST.state.anEase);
    var r = specFrom(snapshot()); $("anPreviewNote").textContent = r.bad || describe(snapshot()) + ", " + ST.nice(r.spec.dur / 1000) + " s, " + r.spec.ease;
    ST.on("reset", function () { renderSegs(); drawCurve(ST.state.anEase); });
    ST.on("ready", loadPresets);
    ST.on("layout", function () { if (ST.isOpen && ST.isOpen("animate") && !document.hidden) queuePlay(); });
  }
  var playTimer = null;
  function queuePlay() { clearTimeout(playTimer); playTimer = setTimeout(play, 350); }

  ST.registerTool({
    id: "animate",
    title: "Animation",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 1,
    defaults: defaults,
    tip: { what: "Builds smooth custom animations with GSAP-style eases.", steps: ["Pick a preset, or turn on Fade, Move, Scale or Rotate.", "Choose an ease and a duration, and watch the preview.", "Select shapes and click Apply to selection."] },
    build: build
  });

  ST._animate = { specFrom: specFrom, STARTERS: STARTERS, apply: apply, presets: function () { return presets; }, setPresets: function (p) { presets = p; renderPresets(); }, usePreset: usePreset, play: play };
})();
