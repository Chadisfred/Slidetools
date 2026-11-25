/* Animation core: eases and slide timing XML for the Animation tool. No PowerPoint calls here,
   so everything can be tested on its own.

   Eases follow the published Penner equations, named the way GSAP names them:
   none, power1 to power4, sine, expo, circ, back(overshoot), elastic(amplitude, period) and
   bounce, each with .in, .out or .inOut, for example "back.out(1.7)".

   PowerPoint keeps animations in the slide's <p:timing>. Custom effects are written as
   keyframed <p:anim> behaviours, which is how PowerPoint stores its own motion effects, so
   an ease is sampled into keyframes. */
(function () {
  "use strict";

  var ST = window.ST;
  var NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  var NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

  // ---------- Eases ----------
  function outOf(fin) { return function (t) { return 1 - fin(1 - t); }; }
  function inOutOf(fin) { return function (t) { return t < 0.5 ? fin(t * 2) / 2 : 1 - fin((1 - t) * 2) / 2; }; }
  function bounceOut(t) {
    var n = 7.5625, d = 2.75;
    if (t < 1 / d) return n * t * t;
    if (t < 2 / d) { t -= 1.5 / d; return n * t * t + 0.75; }
    if (t < 2.5 / d) { t -= 2.25 / d; return n * t * t + 0.9375; }
    t -= 2.625 / d; return n * t * t + 0.984375;
  }
  // The "in" form of each family, given its parameters
  var FAMILIES = {
    none: function () { return function (t) { return t; }; },
    power0: function () { return function (t) { return t; }; },
    power1: function () { return function (t) { return t * t; }; },
    power2: function () { return function (t) { return t * t * t; }; },
    power3: function () { return function (t) { return t * t * t * t; }; },
    power4: function () { return function (t) { return t * t * t * t * t; }; },
    sine: function () { return function (t) { return 1 - Math.cos(t * Math.PI / 2); }; },
    expo: function () { return function (t) { return t === 0 ? 0 : Math.pow(2, 10 * (t - 1)); }; },
    circ: function () { return function (t) { return 1 - Math.sqrt(Math.max(0, 1 - t * t)); }; },
    back: function (p) { var s = isFinite(p[0]) ? p[0] : 1.70158; return function (t) { return t * t * ((s + 1) * t - s); }; },
    elastic: function (p) {
      var a = Math.max(1, isFinite(p[0]) ? p[0] : 1), per = isFinite(p[1]) && p[1] > 0 ? p[1] : 0.3, s = per / (2 * Math.PI) * Math.asin(1 / a);
      return function (t) { if (t === 0 || t === 1) return t; t -= 1; return -(a * Math.pow(2, 10 * t) * Math.sin((t - s) * (2 * Math.PI) / per)); };
    },
    bounce: function () { return function (t) { return 1 - bounceOut(1 - t); }; }
  };
  // GSAP's own inOut for elastic uses a longer period, and back's inOut a larger overshoot
  var ALIASES = { linear: "none", quad: "power1", cubic: "power2", quart: "power3", quint: "power4", strong: "power4" };
  // "back.out(1.7)" becomes { family: "back", dir: "out", params: [1.7] }, or null
  function readEase(text) {
    var m = /^\s*([a-z]+[0-4]?)(?:\.(in|out|inout))?\s*(?:\(([^)]*)\))?\s*$/i.exec(String(text || ""));
    if (!m) return null;
    var fam = m[1].toLowerCase(); fam = ALIASES[fam] || fam;
    if (!FAMILIES[fam]) return null;
    var dir = (m[2] || (fam === "none" || fam === "power0" ? "" : "out")).toLowerCase();
    if (dir === "inout") dir = "inOut";
    var params = m[3] ? m[3].split(",").map(function (x) { return parseFloat(x); }).filter(function (x) { return isFinite(x); }) : [];
    return { family: fam, dir: dir, params: params };
  }
  function easeText(e) {
    if (!e) return "none";
    if (e.family === "none" || e.family === "power0") return "none";
    return e.family + "." + e.dir + (e.params && e.params.length ? "(" + e.params.join(", ") + ")" : "");
  }
  // A function from time (0 to 1) to progress, or null when the text isn't an ease
  function parseEase(text) {
    var e = readEase(text);
    if (!e) return null;
    var fin = FAMILIES[e.family](e.params || []);
    if (e.family === "none" || e.family === "power0") return fin;
    return e.dir === "in" ? fin : e.dir === "inOut" ? inOutOf(fin) : outOf(fin);
  }
  // Time points to sample: more for springy eases so the wobble survives
  // Points along a curve for a move lasting ms: about one every 20 ms, more for springy curves.
  // PowerPoint draws straight lines between them, so enough points make the curve look smooth.
  function pointsFor(e, ms, lo, springLo, hi) {
    if (!e || e.family === "none" || e.family === "power0") return 1;
    var springy = e.family === "elastic" || e.family === "bounce" || e.family === "back";
    var n = Math.round((ms > 0 ? ms : 500) / 20);
    return Math.max(springy ? springLo : lo, Math.min(hi, n));
  }
  function samples(text, dur) {
    var e = readEase(text), n = pointsFor(e, dur, 12, e && (e.family === "elastic" || e.family === "bounce") ? 24 : 12, e && (e.family === "elastic" || e.family === "bounce") ? 60 : 50);
    var out = []; for (var i = 0; i <= n; i++) out.push(i / n);
    return out;
  }

  // ---------- Settings to motion ----------
  // The motion of one effect as plain numbers per sample, shared by the slide XML and the
  // panel preview. spec: { kind: "entr"|"exit"|"emph", fade, move: { dx, dy } (fractions of the
  // slide), movePt: { dx, dy } (points, for the preview), scale (fraction), rotate (degrees),
  // pulse (fraction), spin (degrees), dim (opacity fraction), ease, dur (ms) }.
  // Each frame: { t, opacity, dx, dy, dxPt, dyPt, scale, rot } relative to the shape at rest.
  function frames(spec) {
    if (spec.keys) return keyFrames(spec);
    var ease = parseEase(spec.ease) || parseEase("power2.out"), ts = samples(spec.ease, spec.dur), kind = spec.kind;
    var mv = spec.move || { dx: 0, dy: 0 }, mp = spec.movePt || { dx: 0, dy: 0 };
    return ts.map(function (t) {
      var f = { t: t, opacity: 1, dx: 0, dy: 0, dxPt: 0, dyPt: 0, scale: 1, rot: 0 };
      if (kind === "emph") {
        // Out and back: up over the first half, down over the second, eased both ways
        var u = t <= 0.5 ? ease(t * 2) : ease((1 - t) * 2);
        if (spec.pulse) f.scale = 1 + (spec.pulse - 1) * u;
        if (spec.spin) f.rot = spec.spin * ease(t);
        if (spec.dim !== undefined && spec.dim !== null) f.opacity = 1 + (spec.dim - 1) * ease(t);
        return f;
      }
      var p = ease(t), k = kind === "exit" ? p : 1 - p; // k: how far from rest
      if (spec.fade) f.opacity = kind === "exit" ? 1 - Math.min(1, Math.max(0, p)) : Math.min(1, Math.max(0, p));
      f.dx = mv.dx * k; f.dy = mv.dy * k; f.dxPt = mp.dx * k; f.dyPt = mp.dy * k;
      if (spec.scale) f.scale = 1 + (spec.scale - 1) * k;
      if (spec.rotate) f.rot = spec.rotate * k;
      return f;
    });
  }

  // ---------- Keyframed effects ----------
  // spec.keys holds tracks of keyframes [[t, value, ease?], ...], t from 0 to 1. The ease shapes
  // the segment that ends at that key (spec.ease when it has none). Tracks:
  //   o opacity, x and y offsets in shape widths and heights, X and Y offsets as fractions of the
  //   slide, s scale (both ways), sx and sy scale one way (they win over s), r degrees.
  // Values are absolute: an entrance ends at rest (o 1, offsets 0, scale 1, r 0), an exit starts there.
  var REST = { o: 1, x: 0, y: 0, X: 0, Y: 0, s: 1, sx: 1, sy: 1, r: 0 };
  function trackAt(track, t, fallback) {
    if (!track || !track.length) return fallback;
    if (t <= track[0][0]) return track[0][1];
    for (var i = 1; i < track.length; i++) {
      var a = track[i - 1], b = track[i];
      if (t <= b[0]) {
        var span = b[0] - a[0], u = span > 0 ? (t - a[0]) / span : 1, e = parseEase(b[2] || trackAt.ease) || function (v) { return v; };
        return a[1] + (b[1] - a[1]) * e(u);
      }
    }
    return track[track.length - 1][1];
  }
  function keyTimes(keys, ease, dur) {
    var seen = {}, out = [];
    function add(t) { var k = Math.round(t * 10000) / 10000; if (!seen[k]) { seen[k] = true; out.push(k); } }
    add(0); add(1);
    Object.keys(keys).forEach(function (name) {
      var tr = keys[name] || [];
      for (var i = 0; i < tr.length; i++) {
        add(tr[i][0]);
        if (!i) continue;
        var e = readEase(tr[i][2] || ease), n = tr[i][1] === tr[i - 1][1] ? 1 : pointsFor(e, (tr[i][0] - tr[i - 1][0]) * (dur || 1000), 8, 24, 60);
        for (var j = 1; j < n; j++) add(tr[i - 1][0] + (tr[i][0] - tr[i - 1][0]) * j / n);
      }
    });
    return out.sort(function (a, b) { return a - b; });
  }
  // Each frame: { t, opacity, dxW, dyH (shape sizes), dx, dy (slide fractions), sx, sy, scale, rot }
  function keyFrames(spec) {
    var k = spec.keys || {};
    trackAt.ease = spec.ease || "power1.inOut";
    return keyTimes(k, trackAt.ease, spec.dur).map(function (t) {
      var s = trackAt(k.s, t, 1), sx = trackAt(k.sx, t, s), sy = trackAt(k.sy, t, s);
      return { t: t, opacity: Math.min(1, Math.max(0, trackAt(k.o, t, 1))), dxW: trackAt(k.x, t, 0), dyH: trackAt(k.y, t, 0), dx: trackAt(k.X, t, 0), dy: trackAt(k.Y, t, 0),
               sx: Math.max(0.01, sx), sy: Math.max(0.01, sy), scale: Math.max(0.01, sx), rot: trackAt(k.r, t, 0), dxPt: 0, dyPt: 0 };
    });
  }
  function used(track) { return !!(track && track.length); }
  function offsetFormula(name, size, inSize, slide) {
    var f = "#" + name;
    if (Math.abs(inSize) > 1e-7) f += (inSize >= 0 ? "+" : "-") + "#" + size + "*" + num(Math.abs(inSize));
    if (Math.abs(slide) > 1e-7) f += (slide >= 0 ? "+" : "-") + num(Math.abs(slide));
    return f;
  }
  function keyBehaviours(doc, spec, opts, dur) {
    var spid = String(opts.spid), k = spec.keys || {}, fs = keyFrames(spec), kids = [];
    if (used(k.o)) kids.push(anim(doc, spid, dur, "style.opacity", fs, fs.map(function (f) { return f.opacity; }), false));
    if (used(k.x) || used(k.X)) kids.push(anim(doc, spid, dur, "ppt_x", fs, fs.map(function (f) { return offsetFormula("ppt_x", "ppt_w", f.dxW, f.dx); }), true));
    if (used(k.y) || used(k.Y)) kids.push(anim(doc, spid, dur, "ppt_y", fs, fs.map(function (f) { return offsetFormula("ppt_y", "ppt_h", f.dyH, f.dy); }), true));
    if (used(k.s) || used(k.sx)) kids.push(anim(doc, spid, dur, "ppt_w", fs, fs.map(function (f) { return timesFormula("ppt_w", f.sx); }), true));
    if (used(k.s) || used(k.sy)) kids.push(anim(doc, spid, dur, "ppt_h", fs, fs.map(function (f) { return timesFormula("ppt_h", f.sy); }), true));
    if (used(k.r)) { var r0 = opts.rot0 || 0; kids.push(anim(doc, spid, dur, "style.rotation", fs, fs.map(function (f) { return r0 + f.rot; }), false)); }
    return kids;
  }

  // ---------- XML helpers ----------
  function el(doc, ns, name, attrs, kidsList) {
    var e = doc.createElementNS(ns, (ns === NS_P ? "p:" : "a:") + name);
    Object.keys(attrs || {}).forEach(function (k) { if (attrs[k] !== undefined && attrs[k] !== null) e.setAttribute(k, String(attrs[k])); });
    (kidsList || []).forEach(function (c) { if (c) e.appendChild(c); });
    return e;
  }
  function P(doc, name, attrs, kidsList) { return el(doc, NS_P, name, attrs, kidsList); }
  function kidsP(parent, name) { return Array.prototype.filter.call(parent.childNodes, function (n) { return n.nodeType === 1 && n.namespaceURI === NS_P && (!name || n.localName === name); }); }
  function kidP(parent, name) { return kidsP(parent, name)[0] || null; }
  function cond(doc, delay) { return P(doc, "stCondLst", null, [P(doc, "cond", { delay: delay })]); }
  function num(v) { return String(Math.round(v * 100000) / 100000); }
  function target(doc, spid) { return P(doc, "tgtEl", null, [P(doc, "spTgt", { spid: spid })]); }
  function behaviour(doc, spid, dur, attrName, delay, extra) {
    var hasDelay = delay !== undefined && delay !== null;
    var ctn = P(doc, "cTn", Object.assign({ id: 0, dur: dur, fill: "hold" }, extra || {}), hasDelay ? [cond(doc, delay)] : []);
    return P(doc, "cBhvr", null, [ctn, target(doc, spid), attrName ? P(doc, "attrNameLst", null, [P(doc, "attrName")]) : null].filter(Boolean));
  }
  function setAttrName(bhvr, name) { var a = bhvr.getElementsByTagNameNS(NS_P, "attrName")[0]; if (a) a.textContent = name; }
  // A keyframed <p:anim> on one attribute. values: one per frame, as strings (formulas) or numbers.
  function anim(doc, spid, dur, attrName, fs, values, isStr) {
    var b = behaviour(doc, spid, dur, true); setAttrName(b, attrName);
    // Inside a run of equal values only the run's ends matter, since PowerPoint draws straight lines
    var keep = values.map(function (v, i) { return i === 0 || i === values.length - 1 || !(values[i - 1] === v && values[i + 1] === v); });
    fs = fs.filter(function (f, i) { return keep[i]; }); values = values.filter(function (v, i) { return keep[i]; });
    var tav = P(doc, "tavLst", null, fs.map(function (f, i) {
      var v = isStr ? P(doc, "strVal", { val: values[i] }) : P(doc, "fltVal", { val: num(values[i]) });
      return P(doc, "tav", { tm: Math.round(f.t * 100000) }, [P(doc, "val", null, [v])]);
    }));
    return P(doc, "anim", { calcmode: "lin", valueType: "num" }, [b, tav]);
  }
  function setVisibility(doc, spid, value, delay) {
    var b = behaviour(doc, spid, 1, true, delay); setAttrName(b, "style.visibility");
    return P(doc, "set", null, [b, P(doc, "to", null, [P(doc, "strVal", { val: value })])]);
  }
  function plusFormula(name, v) { return Math.abs(v) < 1e-7 ? "#" + name : "#" + name + (v >= 0 ? "+" : "-") + num(Math.abs(v)); }
  function timesFormula(name, v) { return Math.abs(v - 1) < 1e-7 ? "#" + name : "#" + name + "*" + num(v); }

  // ---------- One effect ----------
  // Builds the effect <p:par> for one shape. opts: { spid, rot0 (the shape's own rotation),
  // nodeType, delay (ms), grpId }. Returns { node, end } where end is delay + duration in ms.
  function buildEffect(doc, spec, opts) {
    var spid = String(opts.spid), dur = Math.max(1, Math.round(spec.dur)), fs = spec.keys ? [] : frames(spec), kids = [], kind = spec.kind;
    if (kind === "entr") kids.push(setVisibility(doc, spid, "visible", 0));
    if (spec.keys) {
      kids = kids.concat(keyBehaviours(doc, spec, opts, dur));
      return finishEffect(doc, spec, opts, dur, kids);
    }
    var fadeUsed = (kind !== "emph" && spec.fade) || (kind === "emph" && spec.dim !== undefined && spec.dim !== null);
    if (fadeUsed && spec.fadeMode === "effect" && kind !== "emph") {
      // PowerPoint's own fade filter, eased with its accelerate and decelerate settings
      var e = readEase(spec.ease) || { dir: "out" }, ex = {};
      if (e.dir === "in" || e.dir === "inOut") ex.accel = 50000;
      if (e.dir === "out" || e.dir === "inOut") ex.decel = 50000;
      var fb = behaviour(doc, spid, dur, false, null, ex);
      kids.push(P(doc, "animEffect", { transition: kind === "exit" ? "out" : "in", filter: "fade" }, [fb]));
    } else if (fadeUsed) kids.push(anim(doc, spid, dur, "style.opacity", fs, fs.map(function (f) { return f.opacity; }), false));
    var mv = spec.move || { dx: 0, dy: 0 };
    if (kind !== "emph" && mv.dx) kids.push(anim(doc, spid, dur, "ppt_x", fs, fs.map(function (f) { return plusFormula("ppt_x", f.dx); }), true));
    if (kind !== "emph" && mv.dy) kids.push(anim(doc, spid, dur, "ppt_y", fs, fs.map(function (f) { return plusFormula("ppt_y", f.dy); }), true));
    var scaled = (kind !== "emph" && spec.scale && spec.scale !== 1) || (kind === "emph" && spec.pulse && spec.pulse !== 1);
    if (scaled) {
      kids.push(anim(doc, spid, dur, "ppt_w", fs, fs.map(function (f) { return timesFormula("ppt_w", f.scale); }), true));
      kids.push(anim(doc, spid, dur, "ppt_h", fs, fs.map(function (f) { return timesFormula("ppt_h", f.scale); }), true));
    }
    var turned = (kind !== "emph" && spec.rotate) || (kind === "emph" && spec.spin);
    if (turned) { var r0 = opts.rot0 || 0; kids.push(anim(doc, spid, dur, "style.rotation", fs, fs.map(function (f) { return r0 + f.rot; }), false)); }
    return finishEffect(doc, spec, opts, dur, kids);
  }
  function finishEffect(doc, spec, opts, dur, kids) {
    var spid = String(opts.spid), kind = spec.kind;
    if (kind === "exit") kids.push(setVisibility(doc, spid, "hidden", Math.max(0, dur - 1)));
    var ctn = P(doc, "cTn", { id: 0, presetID: 0, presetClass: kind, presetSubtype: 0, fill: "hold", grpId: opts.grpId || 0, nodeType: opts.nodeType || "clickEffect" },
      [cond(doc, Math.max(0, Math.round(opts.delay || 0))), P(doc, "childTnLst", null, kids)]);
    return { node: P(doc, "par", null, [ctn]), end: Math.round(opts.delay || 0) + dur };
  }

  // ---------- Timing tree ----------
  function timingOf(doc) { var sld = doc.documentElement; return kidP(sld, "timing"); }
  // The main sequence's child list, made when missing
  function mainSeqList(doc, create) {
    var sld = doc.documentElement, timing = kidP(sld, "timing");
    if (!timing) {
      if (!create) return null;
      var seq = P(doc, "seq", { concurrent: 1, nextAc: "seek" }, [
        P(doc, "cTn", { id: 0, dur: "indefinite", nodeType: "mainSeq" }, [P(doc, "childTnLst")]),
        P(doc, "prevCondLst", null, [P(doc, "cond", { evt: "onPrev", delay: 0 }, [P(doc, "tgtEl", null, [P(doc, "sldTgt")])])]),
        P(doc, "nextCondLst", null, [P(doc, "cond", { evt: "onNext", delay: 0 }, [P(doc, "tgtEl", null, [P(doc, "sldTgt")])])])]);
      timing = P(doc, "timing", null, [P(doc, "tnLst", null, [P(doc, "par", null, [P(doc, "cTn", { id: 0, dur: "indefinite", restart: "never", nodeType: "tmRoot" }, [P(doc, "childTnLst", null, [seq])])])])]);
      // Schema order: after clrMapOvr and transition, before extLst
      var ext = kidP(sld, "extLst");
      sld.insertBefore(timing, ext);
    }
    var ctns = timing.getElementsByTagNameNS(NS_P, "cTn");
    for (var i = 0; i < ctns.length; i++) if (ctns[i].getAttribute("nodeType") === "mainSeq") {
      var list = kidP(ctns[i], "childTnLst");
      if (!list) { list = P(doc, "childTnLst"); ctns[i].appendChild(list); }
      return list;
    }
    if (!create) return null;
    // A timing tree without a main sequence: add one under the root
    var root = timing.getElementsByTagNameNS(NS_P, "cTn")[0], rootList = kidP(root, "childTnLst") || root.appendChild(P(doc, "childTnLst"));
    var cTn = P(doc, "cTn", { id: 0, dur: "indefinite", nodeType: "mainSeq" }, [P(doc, "childTnLst")]);
    rootList.insertBefore(P(doc, "seq", { concurrent: 1, nextAc: "seek" }, [cTn]), rootList.firstChild);
    return kidP(cTn, "childTnLst");
  }
  function condDelay(ctn) { var s = kidP(ctn, "stCondLst"), c = s && kidP(s, "cond"); var d = c ? c.getAttribute("delay") : "0"; return d === "indefinite" ? 0 : (+d || 0); }
  // When an effect ends, from its own delay and its behaviours' delays and durations
  function effectEnd(effectPar) {
    var ctn = kidP(effectPar, "cTn"), start = condDelay(ctn), end = start;
    Array.prototype.forEach.call(effectPar.getElementsByTagNameNS(NS_P, "cBhvr"), function (b) {
      var c = kidP(b, "cTn"); if (!c) return;
      var d = +c.getAttribute("dur") || 0; if (c.getAttribute("autoRev") === "1") d *= 2;
      end = Math.max(end, start + condDelay(c) + d);
    });
    return end;
  }
  function clickGroups(list) { return kidsP(list, "par"); }
  function stepsOf(group) { var c = kidP(group, "cTn"), l = c && kidP(c, "childTnLst"); return l ? kidsP(l, "par") : []; }
  function stepList(step) { var c = kidP(step, "cTn"); return kidP(c, "childTnLst") || c.appendChild(P(step.ownerDocument, "childTnLst")); }
  function stepEnd(step) { var start = condDelay(kidP(step, "cTn")), end = start; kidsP(stepList(step), "par").forEach(function (e) { end = Math.max(end, start + effectEnd(e)); }); return end; }
  function newStep(doc, delay) { return P(doc, "par", null, [P(doc, "cTn", { id: 0, fill: "hold" }, [cond(doc, delay), P(doc, "childTnLst")])]); }
  function newGroup(doc, auto) {
    var conds = [P(doc, "cond", { delay: "indefinite" })];
    // An effect that starts by itself on the first step runs when the slide starts
    if (auto) conds.push(P(doc, "cond", { evt: "onBegin", delay: 0 }, [P(doc, "tn", { val: 2 })]));
    return P(doc, "par", null, [P(doc, "cTn", { id: 0, fill: "hold" }, [P(doc, "stCondLst", null, conds), P(doc, "childTnLst")])]);
  }
  // Adds one effect node. start: "click", "with" or "after".
  function placeEffect(doc, list, node, start) {
    var groups = clickGroups(list), ctn = kidP(node, "cTn");
    if (start === "click" || !groups.length) {
      var g = newGroup(doc, start !== "click"), s = newStep(doc, 0);
      kidP(kidP(g, "cTn"), "childTnLst").appendChild(s);
      stepList(s).appendChild(node); list.appendChild(g);
      ctn.setAttribute("nodeType", start === "click" ? "clickEffect" : start === "with" ? "withEffect" : "afterEffect");
      return;
    }
    var group = groups[groups.length - 1], steps = stepsOf(group), last = steps[steps.length - 1];
    if (start === "with" && last) { stepList(last).appendChild(node); ctn.setAttribute("nodeType", "withEffect"); return; }
    var at = last ? stepEnd(last) : 0, step = newStep(doc, at);
    kidP(kidP(group, "cTn"), "childTnLst").appendChild(step);
    stepList(step).appendChild(node); ctn.setAttribute("nodeType", "afterEffect");
  }
  // Every p:cTn gets a unique id, in document order (the root is 1 and the main sequence 2)
  function renumber(doc) {
    var t = timingOf(doc); if (!t) return;
    var all = t.getElementsByTagNameNS(NS_P, "cTn");
    for (var i = 0; i < all.length; i++) all[i].setAttribute("id", String(i + 1));
    var mainId = null;
    for (var j = 0; j < all.length; j++) if (all[j].getAttribute("nodeType") === "mainSeq") { mainId = all[j].getAttribute("id"); break; }
    if (mainId) Array.prototype.forEach.call(t.getElementsByTagNameNS(NS_P, "tn"), function (tn) { tn.setAttribute("val", mainId); });
  }
  function effectNodes(doc) {
    var t = timingOf(doc); if (!t) return [];
    return Array.prototype.filter.call(t.getElementsByTagNameNS(NS_P, "cTn"), function (c) { return c.hasAttribute("presetClass"); }).map(function (c) { return c.parentNode; });
  }
  function spidsIn(node) { return Array.prototype.map.call(node.getElementsByTagNameNS(NS_P, "spTgt"), function (s) { return s.getAttribute("spid"); }); }
  // The next free build group id for a shape
  function nextGrpId(doc, spid) {
    var max = -1;
    effectNodes(doc).forEach(function (p) { if (spidsIn(p).indexOf(String(spid)) >= 0) max = Math.max(max, +kidP(p, "cTn").getAttribute("grpId") || 0); });
    var t = timingOf(doc), bld = t && kidP(t, "bldLst");
    if (bld) kidsP(bld).forEach(function (b) { if (b.getAttribute("spid") === String(spid)) max = Math.max(max, +b.getAttribute("grpId") || 0); });
    return max + 1;
  }
  // Shapes with text need a build entry so PowerPoint treats the whole shape as one
  function addBuild(doc, spid, grpId) {
    var t = timingOf(doc), bld = kidP(t, "bldLst");
    if (!bld) { bld = P(doc, "bldLst"); t.appendChild(bld); }
    if (kidsP(bld).some(function (b) { return b.getAttribute("spid") === String(spid) && b.getAttribute("grpId") === String(grpId); })) return;
    bld.appendChild(P(doc, "bldP", { spid: spid, grpId: grpId, animBg: 1 }));
  }

  // ---------- Remove ----------
  // Removes effects that only target these shapes, then empty steps, click groups and the
  // timing itself when nothing else is left. Returns how many effects were removed.
  function removeFor(doc, spids) {
    spids = spids.map(String);
    var gone = 0;
    effectNodes(doc).forEach(function (p) {
      var on = spidsIn(p);
      if (on.length && on.every(function (s) { return spids.indexOf(s) >= 0; })) { p.parentNode.removeChild(p); gone++; }
    });
    var list = mainSeqList(doc, false);
    if (list) clickGroups(list).forEach(function (g) {
      stepsOf(g).forEach(function (s) { if (!kidsP(stepList(s), "par").length) s.parentNode.removeChild(s); });
      if (!stepsOf(g).length) list.removeChild(g);
    });
    var t = timingOf(doc);
    if (t) {
      var bld = kidP(t, "bldLst"), still = {};
      effectNodes(doc).forEach(function (p) { spidsIn(p).forEach(function (s) { still[s] = true; }); });
      if (bld) { kidsP(bld).forEach(function (b) { var s = b.getAttribute("spid"); if (spids.indexOf(s) >= 0 && !still[s]) bld.removeChild(b); }); if (!kidsP(bld).length) t.removeChild(bld); }
      if (!effectNodes(doc).length && !t.getElementsByTagNameNS(NS_P, "seq")[1]) t.parentNode.removeChild(t);
      else renumber(doc);
    }
    return gone;
  }

  // ---------- Add ----------
  // jobs: [{ spid, rot0, hasText, start, delay }] in order, each with the same spec.
  function addEffects(doc, spec, jobs) {
    var list = mainSeqList(doc, true);
    jobs.forEach(function (j) {
      var grp = nextGrpId(doc, j.spid);
      var e = buildEffect(doc, spec, { spid: j.spid, rot0: j.rot0, delay: j.delay, grpId: grp });
      placeEffect(doc, list, e.node, j.start);
      if (j.hasText) addBuild(doc, j.spid, grp);
    });
    renumber(doc);
    return jobs.length;
  }

  // ---------- Capture ----------
  // The effects on one shape as text, with its id swapped for a placeholder, in play order.
  // Each item: { xml, start } where start is how that effect began.
  function capture(doc, spid) {
    var ser = new XMLSerializer(), out = [];
    effectNodes(doc).forEach(function (p) {
      var on = spidsIn(p);
      if (!on.length || !on.every(function (s) { return s === String(spid); })) return;
      var c = p.cloneNode(true), nt = kidP(c, "cTn").getAttribute("nodeType");
      Array.prototype.forEach.call(c.getElementsByTagNameNS(NS_P, "spTgt"), function (s) { s.setAttribute("spid", "__SPID__"); });
      out.push({ xml: ser.serializeToString(c), start: nt === "withEffect" ? "with" : nt === "afterEffect" ? "after" : "click" });
    });
    return out;
  }
  // Puts captured effects on shapes. The first effect of each shape follows its job's start and
  // delay. Later effects keep how they began.
  function addCaptured(doc, effects, jobs) {
    var list = mainSeqList(doc, true), n = 0;
    jobs.forEach(function (j) {
      effects.forEach(function (fx, i) {
        var text = fx.xml.replace(/__SPID__/g, String(j.spid));
        var parsed = new DOMParser().parseFromString(text, "application/xml").documentElement;
        if (!parsed || parsed.localName !== "par") return;
        var node = doc.importNode(parsed, true), ctn = kidP(node, "cTn"), grp = nextGrpId(doc, j.spid);
        ctn.setAttribute("grpId", String(grp));
        if (i === 0) { var sc = kidP(ctn, "stCondLst"), c0 = sc && kidP(sc, "cond"); if (c0) c0.setAttribute("delay", String(Math.round((+c0.getAttribute("delay") || 0) + (j.delay || 0)))); }
        placeEffect(doc, list, node, i === 0 ? j.start : fx.start);
        if (j.hasText) addBuild(doc, j.spid, grp);
        n++;
      });
    });
    renumber(doc);
    return n;
  }

  // The slide size in points, from ppt/presentation.xml in an export, or 13.333 x 7.5 in
  function slideSize(files) {
    var f = (files || []).filter(function (x) { return x.name === "ppt/presentation.xml"; })[0];
    if (f) {
      var d = new DOMParser().parseFromString(f.text, "application/xml"), s = d.getElementsByTagNameNS(NS_P, "sldSz")[0];
      if (s) return { w: +s.getAttribute("cx") / 12700, h: +s.getAttribute("cy") / 12700 };
    }
    return { w: 960, h: 540 };
  }

  ST.anim = { readEase: readEase, easeText: easeText, parseEase: parseEase, samples: samples, frames: frames, keyFrames: keyFrames, buildEffect: buildEffect,
              addEffects: addEffects, removeFor: removeFor, capture: capture, addCaptured: addCaptured, slideSize: slideSize,
              effectNodes: effectNodes, renumber: renumber, NS_P: NS_P, NS_A: NS_A };
})();
