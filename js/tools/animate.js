/* Animation tool (id "animate"): custom entrances, exits and emphasis with GSAP-style eases,
   combined effects, exact numbers, a stagger, starter and saved presets, and Capture. Keyframed
   starters (bounces, attention seekers and more) sit in folding groups.
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
    anEase: "power2.out", anDur: "0.5", anDelay: "0", anStart: "click", anStagger: "0.08", anOrder: "selection", anReplace: true, anPresetName: "", anKeys: "", anGroup: "basics", anRepeat: "1", anRepeatMode: "count", anOpen: null
  };
  ST.addDefaults(defaults);
  var KEYS = Object.keys(defaults).filter(function (k) { return k !== "anPresetName" && k !== "anReplace" && k !== "anOrder" && k !== "anGroup" && k !== "anOpen"; });

  // Built in starters, read only. Each one fills the controls.
  function S(o) { return Object.assign({ anKind: "entr", anKeys: "", anFade: false, anMove: false, anScale: false, anRotate: false, anPulse: false, anSpin: false, anDim: false, anDelay: "0" }, o); }
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
  var TRAVEL = { up: [0, -1], down: [0, 1], left: [-1, 0], right: [1, 0] };
  // Keyframed starters: our own keyframes, several inspired by the effects in Animate.css (MIT).
  // Tracks are described in js/animcore.js: o, x and y (shape sizes), X and Y (slide), s, sx, sy, r.
  // TRAVEL gives each direction's unit step: an entrance starts back along its travel.
  function along(dir, d, entering) { var v = TRAVEL[dir], k = entering ? -d : d; return { h: !!v[0], v: v[0] ? v[0] * k : v[1] * k }; }
  function moveTrack(dir, pairs, slide, entering) {
    var a = along(dir, 1, entering), out = {}, name = a.h ? (slide ? "X" : "x") : (slide ? "Y" : "y");
    out[name] = pairs.map(function (p) { return [p[0], p[1] * a.v].concat(p[2] ? [p[2]] : []); });
    return out;
  }
  function K(id, name, kind, dur, ease, keys) { return { id: id, name: name, kind: kind, dur: dur, e: ease, keys: keys }; }
  var DIRS = [["up", "up"], ["down", "down"], ["left", "left"], ["right", "right"]];
  var KEYED = [];
  function add(group, k) { k.group = group; KEYED.push(k); }
  DIRS.forEach(function (d) {
    // One long GSAP curve per motion, or overlapping phases, so the speed never stalls at a keyframe
    add("fade", K("bigfade-" + d[0], "Big fade " + d[1], "entr", 0.9, "power3.out", Object.assign({ o: [[0, 0], [0.6, 1, "power1.out"]] }, moveTrack(d[0], [[0, 0.5], [1, 0, "power3.out"]], true, true))));
    add("fade", K("slide-" + d[0], "Slide " + d[1], "entr", 0.7, "expo.out", moveTrack(d[0], [[0, 1], [1, 0, "expo.out"]], false, true)));
    add("back", K("backin-" + d[0], "Back in " + d[1], "entr", 1, "power3.out", Object.assign({ o: [[0, 0.7], [0.75, 1, "power1.out"]], s: [[0, 0.7], [0.55, 0.7], [1, 1, "back.out(1.7)"]] }, moveTrack(d[0], [[0, 0.8], [0.75, 0, "power3.out"]], true, true))));
    add("bounce", K("bouncein-" + d[0], "Bounce in " + d[1], "entr", 1, "elastic.out(1, 0.5)", Object.assign({ o: [[0, 0], [0.35, 1, "power1.out"]] }, moveTrack(d[0], [[0, 2], [1, 0, "elastic.out(1, 0.5)"]], false, true))));
    add("zoom", K("zoomin-" + d[0], "Zoom in " + d[1], "entr", 0.8, "power3.out", Object.assign({ o: [[0, 0], [0.4, 1, "power1.out"]], s: [[0, 0.1], [1, 1, "power3.out"]] }, moveTrack(d[0], [[0, 2], [1, 0, "back.out(1.4)"]], false, true))));
    add("exit", K("slideout-" + d[0], "Slide out " + d[1], "exit", 0.6, "expo.in", moveTrack(d[0], [[0, 0], [1, 1, "expo.in"]], false, false)));
  });
  add("bounce", K("bouncein", "Bounce in", "entr", 0.9, "elastic.out(1, 0.45)", { o: [[0, 0], [0.35, 1, "power1.out"]], s: [[0, 0.3], [1, 1, "elastic.out(1, 0.45)"]] }));
  add("zoom", K("jackbox", "Jack in the box", "entr", 1, "back.out(1.6)", { o: [[0, 0], [0.3, 1, "power1.out"]], s: [[0, 0.1], [1, 1, "back.out(1.6)"]], r: [[0, 30], [1, 0, "elastic.out(1, 0.35)"]] }));
  add("roll", K("rollin", "Roll in", "entr", 0.8, "power1.out", { o: [[0, 0], [1, 1]], x: [[0, -1], [1, 0]], r: [[0, -120], [1, 0]] }));
  add("roll", K("rotatein", "Rotate in", "entr", 0.8, "power2.out", { o: [[0, 0], [1, 1]], r: [[0, -200], [1, 0]] }));
  add("roll", K("spinin", "Spin in", "entr", 1, "power3.out", { o: [[0, 0], [0.3, 1]], s: [[0, 0.05], [1, 1]], r: [[0, -360], [1, 0]] }));
  var UP = "power2.out", DOWN = "power2.in";
  add("attention", K("bounce", "Bounce", "emph", 1, "power1.inOut", { y: [[0, 0], [0.2, 0], [0.4, -0.3, UP], [0.43, -0.3], [0.53, 0, DOWN], [0.7, -0.15, UP], [0.8, 0, DOWN], [0.9, -0.04, UP], [1, 0, DOWN]], sy: [[0, 1], [0.4, 1.1], [0.53, 0.95], [0.7, 1.05], [0.8, 0.97], [1, 1]] }));
  add("attention", K("flash", "Flash", "emph", 1, "none", { o: [[0, 1], [0.25, 0], [0.5, 1], [0.75, 0], [1, 1]] }));
  add("attention", K("heartbeat", "Heartbeat", "emph", 1.3, "power1.inOut", { s: [[0, 1], [0.14, 1.3], [0.28, 1], [0.42, 1.3], [0.7, 1]] }));
  add("attention", K("rubberband", "Rubber band", "emph", 1, "power1.inOut", { sx: [[0, 1], [0.3, 1.25], [0.4, 0.75], [0.5, 1.15], [0.65, 0.95], [0.75, 1.05], [1, 1]], sy: [[0, 1], [0.3, 0.75], [0.4, 1.25], [0.5, 0.85], [0.65, 1.05], [0.75, 0.95], [1, 1]] }));
  function shake(n, a) { var out = [[0, 0]]; for (var i = 1; i < n; i++) out.push([i / n, i % 2 ? -a : a]); out.push([1, 0]); return out; }
  add("attention", K("shakex", "Shake across", "emph", 1, "none", { x: shake(10, 0.05) }));
  add("attention", K("shakey", "Shake down", "emph", 1, "none", { y: shake(10, 0.08) }));
  add("attention", K("headshake", "Head shake", "emph", 1, "power1.inOut", { x: [[0, 0], [0.065, -0.06], [0.185, 0.05], [0.315, -0.03], [0.435, 0.02], [0.5, 0], [1, 0]], r: [[0, 0], [0.065, -4], [0.185, 3], [0.315, -2], [0.435, 1], [0.5, 0], [1, 0]] }));
  add("attention", K("swing", "Swing", "emph", 1, "power1.inOut", { r: [[0, 0], [0.2, 15], [0.4, -10], [0.6, 5], [0.8, -5], [1, 0]] }));
  add("attention", K("tada", "Tada", "emph", 1, "power1.inOut", { s: [[0, 1], [0.1, 0.9], [0.2, 0.9], [0.3, 1.1], [0.9, 1.1], [1, 1]], r: [[0, 0], [0.1, -3], [0.2, -3], [0.3, 3], [0.4, -3], [0.5, 3], [0.6, -3], [0.7, 3], [0.8, -3], [0.9, 3], [1, 0]] }));
  add("attention", K("wobble", "Wobble", "emph", 1, "power1.inOut", { x: [[0, 0], [0.15, -0.25], [0.3, 0.2], [0.45, -0.15], [0.6, 0.1], [0.75, -0.05], [1, 0]], r: [[0, 0], [0.15, -5], [0.3, 3], [0.45, -3], [0.6, 2], [0.75, -1], [1, 0]] }));
  add("exit", K("backout", "Back out", "exit", 1, "power3.in", { o: [[0, 1], [0.45, 0.7, "power1.in"]], s: [[0, 1], [0.45, 0.7, "back.in(1.7)"]], Y: [[0, 0], [0.25, 0], [1, 0.8, "power3.in"]] }));
  add("exit", K("bounceout", "Bounce out", "exit", 0.75, "power1.inOut", { o: [[0, 1], [0.55, 1], [1, 0]], s: [[0, 1], [0.2, 0.9], [0.5, 1.1], [0.55, 1.1], [1, 0.3]] }));
  add("exit", K("zoomout", "Zoom out", "exit", 0.6, "power2.in", { o: [[0, 1], [0.5, 0], [1, 0]], s: [[0, 1], [0.5, 0.3], [1, 0.3]] }));
  add("exit", K("rollout", "Roll out", "exit", 0.8, "power1.in", { o: [[0, 1], [1, 0]], x: [[0, 0], [1, 1]], r: [[0, 0], [1, 120]] }));
  add("exit", K("bigfadeout", "Big fade out", "exit", 0.8, "power3.in", { o: [[0, 1], [1, 0, "power2.in"]], Y: [[0, 0], [1, 0.5, "power3.in"]] }));
  var GROUPS = [["basics", "Basics"], ["fade", "Fade and slide"], ["back", "Back"], ["bounce", "Bounce"], ["zoom", "Zoom"], ["roll", "Roll and spin"], ["attention", "Attention"], ["exit", "Exits"], ["mine", "Your presets"]];
  function keyed(id) { return KEYED.filter(function (k) { return k.id === id; })[0] || null; }

  var FAMILIES = [["none", "None"], ["power1", "Power 1"], ["power2", "Power 2"], ["power3", "Power 3"], ["power4", "Power 4"], ["sine", "Sine"], ["expo", "Expo"], ["circ", "Circ"], ["back", "Back"], ["elastic", "Elastic"], ["bounce", "Bounce"]];

  var presets = [], presetsLoaded = false, captured = null; // captured: effects of the last applied captured preset

  // ---------- Settings to an effect ----------
  function secs(text) { var v = ST.evaluate(String(text || "").trim(), null, null); return v === null || !isFinite(v) ? null : v; }
  function plainNum(text) { var t = String(text || "").trim().replace(/\s*(%|°|deg)$/i, ""); if (!t) return null; var v = ST.evaluate(t, null, null); return v === null || !isFinite(v) ? null : v; }
  // Returns { spec } or { bad }. size: the slide size in points.
  function specFrom(St, size) {
    size = size || { w: 960, h: 540 };
    var kind = St.anKind === "exit" ? "exit" : St.anKind === "emph" ? "emph" : "entr";
    var dur = secs(St.anDur), delay = secs(St.anDelay);
    if (!(dur > 0 && dur <= 60)) return { bad: "Duration must be more than 0 and up to 60 seconds." };
    if (delay === null || delay < 0) return { bad: "Delay must be 0 or more seconds." };
    var repeat = 1;
    if (St.anRepeatMode === "click" || St.anRepeatMode === "slide") repeat = St.anRepeatMode;
    else {
      var rc = plainNum(St.anRepeat === undefined || St.anRepeat === "" ? "1" : St.anRepeat);
      if (!(rc >= 1 && rc <= 99 && Math.round(rc) === rc)) { if (kind === "emph" || (St.anKeys && keyed(St.anKeys) && keyed(St.anKeys).kind === "emph")) return { bad: "Repeat must be a whole number from 1 to 99." }; rc = 1; }
      repeat = rc;
    }
    if (St.anKeys) {
      var def = keyed(St.anKeys);
      if (!def) return { bad: "That keyframed effect isn't available in this version. Pick another." };
      return { spec: { kind: def.kind, keys: def.keys, ease: def.e, dur: Math.round(dur * 1000), delay: Math.round(delay * 1000), keyed: def.id, repeat: def.kind === "emph" ? repeat : 1 } };
    }
    if (!A.parseEase(St.anEase)) return { bad: "Ease " + St.anEase + " isn't one it knows. Try power2.out, back.out(1.7) or bounce.out." };
    var spec = { kind: kind, ease: A.easeText(A.readEase(St.anEase)), dur: Math.round(dur * 1000), delay: Math.round(delay * 1000), repeat: kind === "emph" ? repeat : 1 };
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
  function repeatText(St) {
    if (St.anRepeatMode === "click") return ", until the next click";
    if (St.anRepeatMode === "slide") return ", until the end of the slide";
    var n = plainNum(St.anRepeat); return n > 1 ? ", " + n + " times" : "";
  }
  function describe(St) {
    var em = St.anKind === "emph" || (St.anKeys && keyed(St.anKeys) && keyed(St.anKeys).kind === "emph");
    return describeBase(St) + (em ? repeatText(St) : "");
  }
  function describeBase(St) {
    var kind = { entr: "Entrance", exit: "Exit", emph: "Emphasis" }[St.anKind] || "Entrance", parts = [];
    var def = St.anKeys && keyed(St.anKeys);
    if (def) return ({ entr: "Entrance", exit: "Exit", emph: "Emphasis" })[def.kind] + ": " + def.name;
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
        ST.toast(res.count ? "Animated " + ST.plural(res.count, "shape") + (res.count > 1 && secs(ST.state.anStagger) > 0 ? ", " + ST.nice(secs(ST.state.anStagger)) + " s apart" : "") + "." : R.missText(), !res.count);
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

  // ---------- Copy, paste and repeat sets ----------
  // The copied set lives in memory and in this session's storage, so it can be pasted on another slide.
  var COPY_KEY = "slideToolsAnimSet", copied = null;
  try { copied = JSON.parse(sessionStorage.getItem(COPY_KEY) || "null"); } catch (e) { copied = null; }
  function renderCopied() {
    var el = $("anCopied"), b = $("anPaste");
    if (!el) return;
    b.disabled = !copied;
    el.textContent = copied ? "Copied: " + ST.plural(copied.effects.length, "animation") + ", " + ST.plural(copied.shapes, "shape") + "." : "Nothing copied yet.";
  }
  function spidsOf(doc, items) {
    return R.matchShapes(doc, boxesOf(items), SHAPE_KINDS).map(function (el) {
      if (!el) return null;
      var c = R.cNvPr(el);
      return c ? { spid: c.getAttribute("id"), hasText: el.localName === "sp" && !!R.kid(el, A.NS_P, "txBody") } : null;
    });
  }
  function copySet() {
    if (!ST.needOffice()) return;
    if (!R.supported()) { ST.toast("Copying animations needs PowerPoint API 1.8.", true); return; }
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      items = orderItems(items, ST.state.anOrder);
      var set = null, missing = false;
      return R.edit(context, function (doc) {
        var t = spidsOf(doc, items);
        missing = t.some(function (x) { return !x; });
        // The set's shapes follow the panel's order rather than the order they first animate
        var got = t.filter(Boolean).map(function (x) { return x.spid; });
        set = A.captureSet(doc, got);
        if (set.effects.length) {
          var remap = set.order.map(function (sp) { return got.indexOf(sp); });
          set.effects.forEach(function (fx) { fx.src = remap[fx.src]; });
          set.shapes = Math.max.apply(null, remap) + 1;
        }
        return { count: 0 }; // read only: nothing is swapped
      }).then(function () {
        if (!set || !set.effects.length) { ST.toast(missing ? R.missText() : "The selected shapes have no animations of their own.", true); return; }
        copied = { effects: set.effects, shapes: set.shapes };
        try { sessionStorage.setItem(COPY_KEY, JSON.stringify(copied)); } catch (e) { /* kept in memory for now */ }
        renderCopied();
        ST.toast("Copied " + ST.plural(set.effects.length, "animation") + " from " + ST.plural(set.shapes, "shape") + ".");
      });
    }, "Select the shapes whose animations you want to copy.");
  }
  // set: a set to add, or null to paste the copied one. same: add each shape's own set again.
  function pasteSet(set, same) {
    if (!same && !(set || copied)) { ST.toast("Copy a set of animations first.", true); return; }
    if (!ST.needOffice()) return;
    if (!R.supported()) { ST.toast("This version of PowerPoint can't change animations. It needs PowerPoint API 1.8.", true); return; }
    var use = set || copied;
    ST.setRepeat(same ? "Repeat animation set" : "Paste animations", function () { if (same) repeatSet(); else pasteSet(use); });
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      items = orderItems(items, ST.state.anOrder);
      var added = 0, shapes = 0, missing = false, nothing = false;
      return ST.confirmSwap(same ? "repeat animations" : "paste animations").then(function (ok) {
        if (!ok) return null;
        return R.edit(context, function (doc) {
          var t = spidsOf(doc, items), targets = t.filter(Boolean);
          missing = t.some(function (x) { return !x; });
          if (!targets.length) return { count: 0 };
          var s = use;
          if (same) {
            s = A.captureSet(doc, targets.map(function (x) { return x.spid; }));
            if (!s.effects.length) { nothing = true; return { count: 0 }; }
            var ids = targets.map(function (x) { return x.spid; });
            s.effects.forEach(function (fx) { fx.src = ids.indexOf(s.order[fx.src]); });
          }
          added = A.addSet(doc, s, targets); shapes = targets.length;
          return { count: added ? targets.length : 0 };
        });
      }).then(function (res) {
        if (res === null) { ST.toast("Nothing was changed."); return; }
        if (nothing) { ST.toast("The selected shapes have no animations of their own to repeat.", true); return; }
        if (!added) { ST.toast(missing ? R.missText() : "Select the shapes to paste onto first.", true); return; }
        var note = !same && use.shapes !== shapes ? " (the copied set had " + ST.plural(use.shapes, "shape") + ", so it " + (use.shapes > shapes ? "wraps round" : "repeats across them") + ")" : "";
        ST.toast((same ? "Repeated " : "Pasted ") + ST.plural(added, "animation") + (same ? "" : " onto " + ST.plural(shapes, "shape")) + note + ".");
        ST.refreshSelection();
      }, function (e) { ST.toast("Couldn't change the slide. " + (ST.errorText ? ST.errorText(e) : ""), true); });
    }, same ? "Select the shapes whose animations you want to repeat." : "Select the shapes to paste onto.");
  }
  function repeatSet() { pasteSet(null, true); }

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
    if (p.keys) {
      ST.state.anKeys = p.id; ST.state.anKind = p.kind; ST.state.anDur = String(p.dur);
      ST.save(); ST.emit("reset"); ST.changed(); play(); return;
    }
    Object.keys(p.s).forEach(function (k) { if (k in defaults) ST.state[k] = p.s[k]; });
    ST.save(); ST.emit("reset"); ST.changed(); play();
  }
  function presetChip(p) {
    var b = document.createElement("button"), on = p.keys ? ST.state.anKeys === p.id : false;
    b.className = "mini-chip an-chip" + (p.id && !p.keys ? " mine" : "") + (on ? " on" : "");
    b.textContent = p.name;
    b.title = p.captured ? "Apply this captured animation to the selected shapes" : p.keys ? "Use " + p.name + ", a keyframed " + ({ entr: "entrance", exit: "exit", emph: "emphasis" })[p.kind] : p.id ? "Your preset: fill the controls" : "Fill the controls with " + p.name;
    if (!p.keys && !p.id && pickedName === p.name && !ST.state.anKeys) b.classList.add("on");
    b.addEventListener("click", function () { pickedName = p.keys ? "" : p.name; usePreset(p); });
    if (p.id && !p.keys) {
      var del = document.createElement("span");
      del.className = "del"; del.textContent = "×"; del.title = "Delete this preset";
      del.addEventListener("click", function (e) { e.stopPropagation(); presets = presets.filter(function (x) { return x.id !== p.id; }); savePresets(); });
      b.appendChild(del);
    }
    return b;
  }
  // ---------- Effect gallery ----------
  // Effects for the chosen kind, in categories, with a search across them
  var CATS = {
    entr: [["all", "All"], ["simple", "Simple"], ["fade", "Fade and slide"], ["back", "Back"], ["bounce", "Bounce"], ["zoom", "Zoom"], ["roll", "Roll and spin"], ["mine", "Mine"]],
    emph: [["all", "All"], ["simple", "Simple"], ["attention", "Attention"], ["mine", "Mine"]],
    exit: [["all", "All"], ["simple", "Simple"], ["exit", "Exits"], ["mine", "Mine"]]
  };
  var search = "", pickedName = "";
  function kindOf(p) { if (p.captured) return "any"; if (p.keys) return p.kind; var k = p.s && p.s.anKeys && keyed(p.s.anKeys); return k ? k.kind : (p.s && p.s.anKind) || "entr"; }
  function effectsFor(kind, cat) {
    var fits = function (p) { var k = kindOf(p); return k === "any" || k === kind; };
    var simple = STARTERS.filter(fits), keyedList = KEYED.filter(fits), mine = presets.filter(fits);
    if (cat === "simple") return simple;
    if (cat === "mine") return mine;
    if (cat === "all" || !cat) return simple.concat(keyedList, mine);
    return keyedList.filter(function (k) { return k.group === cat; });
  }
  function renderPresets() {
    var cats = $("anCats"), gal = $("anGallery");
    if (!cats) return;
    var kind = ST.state.anKind === "exit" ? "exit" : ST.state.anKind === "emph" ? "emph" : "entr";
    var list = CATS[kind].filter(function (c) { return c[0] !== "mine" || effectsFor(kind, "mine").length; });
    var cat = ST.state.anGroup;
    if (!list.some(function (c) { return c[0] === cat; })) cat = "all";
    cats.innerHTML = "";
    list.forEach(function (c) {
      var b = document.createElement("button");
      b.className = "mini-chip an-cat" + (c[0] === cat && !search ? " on" : "");
      b.textContent = c[1]; b.setAttribute("data-cat", c[0]);
      b.addEventListener("click", function () { ST.state.anGroup = c[0]; ST.save(); search = ""; $("anSearch").value = ""; renderPresets(); });
      cats.appendChild(b);
    });
    var q = search.trim().toLowerCase();
    var items = q ? effectsFor(kind, "all").filter(function (p) { return p.name.toLowerCase().indexOf(q) >= 0; }) : effectsFor(kind, cat);
    gal.innerHTML = "";
    if (!q && cat === "all") {
      var custom = document.createElement("button");
      custom.className = "mini-chip an-chip an-custom" + (!ST.state.anKeys && !pickedName ? " on" : "");
      custom.textContent = "Custom"; custom.title = "Build your own from Fade, Move, Scale and Rotate (or Pulse, Spin and Dim)";
      custom.addEventListener("click", function () { pickedName = ""; dropKeys(); setOpen("custom", true); ST.save(); renderSegs(); ST.changed(); play(); });
      gal.appendChild(custom);
    }
    items.forEach(function (p) { gal.appendChild(presetChip(p)); });
    if (!items.length) { var none = document.createElement("div"); none.className = "muted-note"; none.textContent = q ? "No effects match " + search.trim() + "." : "Nothing here yet."; gal.appendChild(none); }
  }

  // ---------- Folding sections ----------
  function openState() { var o = ST.state.anOpen; return o && typeof o === "object" ? o : {}; }
  function setOpen(id, on) { var o = Object.assign({}, openState()); o[id] = on; ST.state.anOpen = o; renderSections(); }
  function sectionMarkup(id, title, body) {
    return '<div class="an-sec" data-sec="' + id + '"><button class="an-sec-head" aria-expanded="false"><span class="an-caret" aria-hidden="true"></span><span class="an-sec-title">' + title + '</span><span class="an-sec-sum" id="anSum_' + id + '"></span></button><div class="an-sec-body">' + body + '</div></div>';
  }
  function summaries() {
    var S = ST.state, def = S.anKeys && keyed(S.anKeys), emph = S.anKind === "emph", out = {};
    out.custom = def ? "Keyframed: " + def.name : describeBase(S).replace(/^[^:]+:\s*/, "") || "Nothing on";
    if (emph) out.custom += repeatText(S);
    out.motion = (def ? "keyframed" : S.anEase) + " · " + ST.nice(secs(S.anDur) || 0) + " s";
    var d = secs(S.anDelay) || 0;
    out.timing = ({ click: "On click", "with": "With previous", after: "After previous" })[S.anStart] + " · " + (d ? ST.nice(d) + " s delay" : "no delay") + ((ST.selection && ST.selection.count > 1) ? " · stagger " + ST.nice(secs(S.anStagger) || 0) + " s" : "");
    out.presets = presets.length ? ST.plural(presets.length, "preset") + " saved" : "None yet";
    return out;
  }
  function renderSections() {
    var o = openState(), sum = summaries();
    Array.prototype.forEach.call(document.querySelectorAll("[data-tool=animate] .an-sec"), function (sec) {
      var id = sec.getAttribute("data-sec"), on = !!o[id];
      sec.classList.toggle("open", on);
      sec.querySelector(".an-sec-head").setAttribute("aria-expanded", String(on));
      var el = $("anSum_" + id); if (el) el.textContent = on ? "" : sum[id] || "";
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
    note.textContent = describe(snapshot()) + ", " + ST.nice(r.spec.dur / 1000) + " s" + (r.spec.keys ? ", keyframed" : ", " + r.spec.ease); note.classList.remove("bad");
    var fs = A.frames(r.spec), stage = box.parentNode, bw = box.offsetWidth || 40, bh = box.offsetHeight || 28, sw = (stage && stage.offsetWidth) || 240, sh = (stage && stage.offsetHeight) || 90;
    var kf = fs.map(function (f) {
      var x = r.spec.keys ? f.dxW * bw + f.dx * sw : f.dxPt * PX_PER_PT, y = r.spec.keys ? f.dyH * bh + f.dy * sh : f.dyPt * PX_PER_PT;
      var sx = f.sx || f.scale, sy = f.sy || f.scale;
      return { offset: f.t, opacity: f.opacity, transform: "translate(" + x.toFixed(2) + "px," + y.toFixed(2) + "px) scale(" + sx.toFixed(4) + "," + sy.toFixed(4) + ") rotate(" + f.rot.toFixed(2) + "deg)" };
    });
    var its = typeof r.spec.repeat === "number" ? r.spec.repeat : r.spec.kind === "emph" && r.spec.repeat ? 3 : 1;
    running = box.animate(kf, { duration: r.spec.dur, delay: 150, fill: "both", easing: "linear", iterations: its });
    // An exit leaves the sample hidden, so bring it back after a moment
    if (r.spec.kind === "exit") running.onfinish = function () { setTimeout(function () { if (running) { running.cancel(); running = null; } }, 600); };
    drawCurve(r.spec.keys ? "none" : r.spec.ease);
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
  var ICONS = {
    copy: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" aria-hidden="true"><rect x="5" y="5" width="8.5" height="8.5" rx="1.5"/><path d="M3.5 10.5h-.5A1.5 1.5 0 0 1 1.5 9V3A1.5 1.5 0 0 1 3 1.5h6A1.5 1.5 0 0 1 10.5 3v.5"/></svg>',
    paste: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" aria-hidden="true"><rect x="3" y="3" width="10" height="11.5" rx="1.5"/><path d="M6 3V1.8h4V3M5.5 7.5h5M5.5 10.5h3.5"/></svg>',
    repeat: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2.5 8a5.5 5.5 0 0 1 9.4-3.9L13.5 6M13.5 2.5V6H10M13.5 8a5.5 5.5 0 0 1-9.4 3.9L2.5 10M2.5 13.5V10H6"/></svg>',
    remove: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" aria-hidden="true"><path d="M2.5 4h11M6 4V2.5h4V4M4 4l.7 9.5h6.6L12 4"/></svg>'
  };
  function actionBtn(id, icon, label, title, disabled) { return '<button class="btn small an-act" id="' + id + '" title="' + title + '"' + (disabled ? " disabled" : "") + '>' + ICONS[icon] + '<span>' + label + '</span></button>'; }
  var MARKUP =
    seg("anKind", [["entr", "Entrance"], ["emph", "Emphasis"], ["exit", "Exit"]]) +
    '<div class="an-search"><input type="search" id="anSearch" placeholder="Search effects" aria-label="Search effects" autocomplete="off"></div>' +
    '<div class="an-cats" id="anCats"></div>' +
    '<div class="mini-chips an-gallery" id="anGallery"></div>' +
    '<div class="an-preview"><div class="an-stage"><div class="an-sample" id="anSample"></div></div><button class="btn small" id="anPlay">Play</button></div>' +
    '<div class="help an-note" id="anPreviewNote"></div>' +
    sectionMarkup("custom", "Effect",
      '<div class="an-keyed-note" id="anKeyedNote" hidden><span></span><button class="link-btn" id="anKeyedOff">Use my own settings</button></div>' +
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
      '<div class="an-repeat"><div class="field"><span class="label-row">Repeat' + ST.info("How many times the emphasis plays, or until the next click or the end of the slide, like PowerPoint\'s Timing settings.", "About repeat") + '</span>' + seg("anRepeatMode", [["count", "Times"], ["click", "Until click", "Until the next click"], ["slide", "Until end", "Until the end of the slide"]]) + '</div>' +
        '<label class="field an-repeat-count">Times<div class="stepper"><input type="text" id="anRepeat" inputmode="numeric"><button id="anRepeatLess" title="Fewer times">&minus;</button><button id="anRepeatMore" title="More times">+</button></div></label></div>') +
    sectionMarkup("motion", "Motion",
      '<div class="an-ease"><div class="an-ease-controls"><span class="label-row sublabel">Ease' + ST.info("Named the way GSAP names them. Power 1 to 4 get stronger, Back overshoots, Elastic springs and Bounce bounces. In starts slowly, Out ends slowly. You can also type a name such as <code>back.out(1.7)</code>.", "About eases") + '</span><select id="anFamily" aria-label="Ease">' + FAMILIES.map(function (f) { return '<option value="' + f[0] + '">' + f[1] + '</option>'; }).join("") + '</select>' +
        seg("anEaseDir", [["in", "In"], ["out", "Out"], ["inOut", "In out"]]) +
        '<input type="text" id="anEase" spellcheck="false" aria-label="Ease as text"></div><svg id="anCurve" viewBox="0 0 120 64" aria-hidden="true"></svg></div>' +
      '<label class="field">Duration (s)<input type="text" id="anDur" inputmode="decimal"></label>') +
    sectionMarkup("timing", "Timing",
      '<div class="field">Start' + seg("anStart", [["click", "On click"], ["with", "With previous"], ["after", "After previous"]]) + '</div>' +
      '<label class="field">Delay (s)<input type="text" id="anDelay" inputmode="decimal"></label>' +
      '<div class="row" data-min="2"><label class="field"><span class="label-row">Stagger (s)' + ST.info("With several shapes selected, each one starts this much after the one before.", "About stagger") + '</span><input type="text" id="anStagger" inputmode="decimal"></label>' +
        '<label class="field">Order' + seg("anOrder", [["selection", "Picked", "In the order PowerPoint gives the selection"], ["lr", "&rarr;", "Left to right"], ["tb", "&darr;", "Top to bottom"]]) + '</label></div>' +
      '<div class="toggle" data-min="2"><span class="label-row">Replace their animations' + ST.info("Removes the shapes\' earlier animations first, so applying again doesn\'t add another.", "About replacing") + '</span><label class="switch"><input type="checkbox" id="anReplace"><span></span></label></div>') +
    sectionMarkup("presets", "Presets",
      '<div class="inline"><label class="field">Preset name<input type="text" id="anPresetName" maxlength="32" placeholder="e.g. Card in"></label>' +
        '<div class="pair"><button class="btn small" id="anSave" title="Save these settings as a preset in this presentation">Save</button><button class="btn small" id="anCapture" data-min="2" title="Save the selected shape\'s own animations as a preset">Capture</button></div></div>' +
      '<div class="help">Saved presets appear under Mine in the effects above.</div>') +
    '<div class="an-actions">' +
      actionBtn("anCopy", "copy", "Copy", "Copy the selected shapes\' animations as a set, with their order and timing") +
      actionBtn("anPaste", "paste", "Paste", "Add the copied set to the selected shapes, after their own animations", true) +
      actionBtn("anRepeatSet", "repeat", "Repeat", "Add the selected shapes\' animations again, after themselves") +
      actionBtn("anRemove", "remove", "Remove", "Remove the selected shapes\' animations") +
    '</div>' +
    '<div class="help an-copied" id="anCopied">Nothing copied yet.</div>' +
    '<div class="an-footer"><button class="btn primary" id="anApply">Apply to selection</button></div>';

  function renderSegs() {
    [["anKind", "anKind"], ["anDir", "anDir"], ["anStart", "anStart"], ["anOrder", "anOrder"]].forEach(function (p) {
      Array.prototype.forEach.call($(p[0]).children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === ST.state[p[1]]); });
    });
    var e = A.readEase(ST.state.anEase) || { family: "power2", dir: "out" };
    $("anFamily").value = e.family === "power0" ? "none" : e.family;
    Array.prototype.forEach.call($("anEaseDir").children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === e.dir); b.disabled = e.family === "none"; });
    var emph = ST.state.anKind === "emph", exit = ST.state.anKind === "exit";
    var mode = ST.state.anRepeatMode || "count";
    Array.prototype.forEach.call($("anRepeatMode").children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === mode); });
    document.querySelector(".an-repeat").hidden = !emph;
    document.querySelector(".an-repeat-count").hidden = mode !== "count";
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
    var def = ST.state.anKeys && keyed(ST.state.anKeys), note = $("anKeyedNote");
    note.hidden = !def;
    if (def) note.querySelector("span").textContent = "Keyframed effect: " + def.name + ". Duration, delay, start and stagger apply. Changing an effect or the ease goes back to your own settings.";
    Array.prototype.forEach.call(document.querySelectorAll(".an-move-group, .an-emph-group, .an-ease"), function (n) { n.classList.toggle("an-dim", !!def); });
  }
  // Leaves a keyframed starter: the panel's own effect settings apply again
  function dropKeys() { if (ST.state.anKeys) { ST.state.anKeys = ""; } renderPresets(); }
  function setEaseFromControls(dir) {
    dropKeys();
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
    ST.bindField($("anEase"), "anEase", { onInput: function () { dropKeys(); renderSegs(); queuePlay(); } });
    $("anKeyedOff").addEventListener("click", function () { dropKeys(); ST.save(); renderSegs(); ST.changed(); play(); });
    ["anDist", "anScaleVal", "anRotVal", "anPulseVal", "anSpinVal", "anDimVal"].forEach(function (k) { $(k).addEventListener("input", function () { if (ST.state.anKeys) { dropKeys(); renderSegs(); } }); });
    [["anKind", "anKind"], ["anDir", "anDir"], ["anStart", "anStart"], ["anOrder", "anOrder"]].forEach(function (p) {
      Array.prototype.forEach.call($(p[0]).children, function (b) {
        b.addEventListener("click", function (e) { e.preventDefault(); if (p[0] === "anKind" || p[0] === "anDir") dropKeys(); if (p[0] === "anKind") pickedName = ""; ST.state[p[1]] = b.getAttribute("data-v"); ST.save(); renderSegs(); if (p[0] === "anKind") renderPresets(); ST.changed(); if (p[0] !== "anOrder" && p[0] !== "anStart") play(); });
      });
    });
    ["anFade", "anMove", "anScale", "anRotate", "anPulse", "anSpin", "anDim", "anReplace"].forEach(function (k) {
      $(k).addEventListener("change", function () { if (k !== "anReplace") { pickedName = ""; dropKeys(); } ST.state[k] = $(k).checked; ST.save(); renderSegs(); ST.changed(); if (k !== "anReplace") play(); });
    });
    $("anFamily").addEventListener("change", function () { setEaseFromControls(); });
    Array.prototype.forEach.call($("anEaseDir").children, function (b) { b.addEventListener("click", function (e) { e.preventDefault(); setEaseFromControls(b.getAttribute("data-v")); }); });
    $("anPlay").addEventListener("click", play);
    Array.prototype.forEach.call(body.querySelectorAll(".an-sec-head"), function (h) {
      h.addEventListener("click", function () { var id = h.parentNode.getAttribute("data-sec"); setOpen(id, !openState()[id]); ST.save(); });
    });
    $("anSearch").addEventListener("input", function () { search = $("anSearch").value; renderPresets(); });
    ST.on("change", renderSections);
    ST.on("selection", renderSections);
    ST.bindField($("anRepeat"), "anRepeat", { onInput: function () { queuePlay(); } });
    Array.prototype.forEach.call($("anRepeatMode").children, function (b) {
      b.addEventListener("click", function (e) { e.preventDefault(); ST.state.anRepeatMode = b.getAttribute("data-v"); ST.save(); renderSegs(); ST.changed(); play(); });
    });
    [["anRepeatLess", -1], ["anRepeatMore", 1]].forEach(function (x) {
      $(x[0]).addEventListener("click", function (e) { e.preventDefault(); var n = Math.round(plainNum(ST.state.anRepeat) || 1) + x[1]; ST.state.anRepeat = String(Math.max(1, Math.min(99, n))); $("anRepeat").value = ST.state.anRepeat; ST.save(); ST.changed(); queuePlay(); });
    });
    $("anCopy").addEventListener("click", copySet);
    $("anPaste").addEventListener("click", function () { pasteSet(null); });
    $("anRepeatSet").addEventListener("click", repeatSet);
    renderCopied();
    $("anApply").addEventListener("click", function () { apply(null); });
    $("anSave").addEventListener("click", savePreset);
    $("anCapture").addEventListener("click", captureSelection);
    $("anRemove").addEventListener("click", removeFromSelection);
    $("anPresetName").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); savePreset(); } });
    renderSegs(); renderPresets(); renderSections(); drawCurve(ST.state.anEase);
    var r = specFrom(snapshot()); $("anPreviewNote").textContent = r.bad || describe(snapshot()) + ", " + ST.nice(r.spec.dur / 1000) + " s" + (r.spec.keys ? ", keyframed" : ", " + r.spec.ease);
    ST.on("reset", function () { renderSegs(); renderPresets(); renderSections(); drawCurve(ST.state.anEase); });
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

  ST._animate = { copySet: copySet, pasteSet: pasteSet, repeatSet: repeatSet, copied: function () { return copied; }, specFrom: specFrom, STARTERS: STARTERS, KEYED: KEYED, GROUPS: GROUPS, CATS: CATS, effectsFor: effectsFor, apply: apply, presets: function () { return presets; }, setPresets: function (p) { presets = p; renderPresets(); }, usePreset: usePreset, play: play };
})();
