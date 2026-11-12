/* Shape styles: sample a shape's fill, outline, corners and text box settings, save them
   as a named style in the presentation, apply them, and push edits with Update all.
   A shape style can also carry a character style for the text inside.

   Shapes are linked with a shape tag holding the style ID, so recolouring a shape never
   breaks its link. Only normal slides are read or changed, never masters or layouts. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;
  var STORE_KEY = "SLIDETOOLS_SHAPESTYLES";
  var SHAPE_TAG = "SLIDETOOLS_SHAPESTYLE";
  var PART_SIZE = 2000;
  var BATCH = 150;
  var SHAPE_TYPES = ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"];

  var KEYS = ["fillType", "fillColor", "fillTransparency",
              "lineVisible", "lineColor", "lineWeight", "lineDash", "lineStyle", "lineTransparency",
              "radius", "marginLeft", "marginRight", "marginTop", "marginBottom", "vAlign", "autoSize", "wordWrap",
              "textStyleId"];
  var DASHES = [["Solid", "Solid"], ["RoundDot", "Round dot"], ["SquareDot", "Square dot"], ["Dash", "Dash"], ["DashDot", "Dash dot"],
                ["DashDotDot", "Dash dot dot"], ["LongDash", "Long dash"], ["LongDashDot", "Long dash dot"], ["LongDashDotDot", "Long dash dot dot"],
                ["SystemDash", "System dash"], ["SystemDot", "System dot"], ["SystemDashDot", "System dash dot"]];
  var COMPOUNDS = [["Single", "Single"], ["ThinThin", "Double"], ["ThickThin", "Thick and thin"], ["ThinThick", "Thin and thick"], ["ThickBetweenThin", "Triple"]];
  var VALIGNS = [["Top", "Top"], ["Middle", "Middle"], ["Bottom", "Bottom"], ["TopCentered", "Top centred"], ["MiddleCentered", "Middle centred"], ["BottomCentered", "Bottom centred"]];
  var AUTOSIZES = [["AutoSizeNone", "Don't autofit"], ["AutoSizeTextToFitShape", "Shrink text on overflow"], ["AutoSizeShapeToFitText", "Resize shape to fit text"]];

  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="2" y="3" width="10" height="9" rx="2" fill="currentColor" fill-opacity="0.25"/><rect x="6" y="6.5" width="10" height="9" rx="2" fill="currentColor" stroke="none"/></svg>';
  var EDIT_ICON = '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M10.5 2.5l3 3L6 13H3v-3l7.5-7.5z"/></svg>';

  var styles = [], stylesLoaded = false, editing = null, draft = null, review = null, busy = false;

  // ---------- Helpers ----------
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function byId(id) { return styles.filter(function (s) { return s.id === id; })[0]; }
  function nextId() { var id = 1; while (byId(id)) id++; return id; }
  function pending(s) { return JSON.stringify(s.props) !== JSON.stringify(s.applied); }
  function blank() { var p = {}; KEYS.forEach(function (k) { p[k] = null; }); return p; }
  function has(v) { return v !== null && v !== undefined && v !== ""; }
  function same(a, b) {
    if (!has(a) || !has(b)) return false;
    if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) < 0.05;
    return String(a).toUpperCase() === String(b).toUpperCase();
  }
  var canvas = null;
  function norm(c) {
    if (!has(c)) return null;
    c = String(c).trim();
    if (/^#?[0-9a-f]{6}$/i.test(c)) return ("#" + c.replace("#", "")).toUpperCase();
    try {
      canvas = canvas || document.createElement("canvas").getContext("2d");
      canvas.fillStyle = "#000000"; canvas.fillStyle = c;
      return /^#[0-9a-f]{6}$/i.test(canvas.fillStyle) ? canvas.fillStyle.toUpperCase() : null;
    } catch (e) { return null; }
  }
  function esc(t) { return String(has(t) ? t : "").replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function round1(v) { return Math.round(v * 10) / 10; }
  function textStyles() { return ST.textStyles ? ST.textStyles.list() : []; }
  function textStyleName(id) { var t = ST.textStyles && ST.textStyles.byId(id); return t ? t.name : null; }

  // ---------- Saving in the presentation ----------
  function loadStyles() {
    if (!ST.inOffice || !ST.supports("1.3")) { stylesLoaded = true; render(); return; }
    PowerPoint.run(function (context) {
      var tags = context.presentation.tags, count = tags.getItemOrNullObject(STORE_KEY + "_N");
      count.load("value");
      return context.sync().then(function () {
        var n = count.isNullObject ? 0 : parseInt(count.value, 10) || 0, parts = [];
        for (var i = 0; i < n; i++) { var t = tags.getItemOrNullObject(STORE_KEY + "_" + i); t.load("value"); parts.push(t); }
        return context.sync().then(function () {
          var text = parts.map(function (t) { return t.isNullObject ? "" : t.value; }).join("");
          styles = text ? JSON.parse(text) : [];
        });
      });
    }).catch(function () {
      styles = [];
      ST.toast("Couldn't read the shape styles saved in this presentation.", true);
    }).then(function () { stylesLoaded = true; render(); refreshStatus(); });
  }
  function saveStyles() {
    render();
    if (!ST.inOffice || !ST.supports("1.3")) return Promise.resolve();
    var text = JSON.stringify(styles), parts = [];
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
    }).catch(function (e) { ST.toast("Couldn't save shape styles: " + (e && e.message ? e.message : e), true); });
  }

  // ---------- Reading shapes ----------
  // Loads everything a shape style cares about for a list of shapes, in two round trips.
  function readShapes(context, shapes) {
    var useNullFrame = ST.supports("1.10"), canRound = ST.cornersSupported();
    var items = shapes.map(function (sh) {
      var it = { shape: sh };
      sh.load("id,type,width,height");
      sh.fill.load("type,foregroundColor,transparency");
      sh.lineFormat.load("visible,color,weight,dashStyle,style,transparency");
      it.tagObj = sh.tags.getItemOrNullObject(SHAPE_TAG); it.tagObj.load("value");
      it.textTagObj = sh.tags.getItemOrNullObject(ST.textStyles ? ST.textStyles.shapeTag : "SLIDETOOLS_USES"); it.textTagObj.load("value");
      return it;
    });
    return context.sync().then(function () {
      items.forEach(function (it) {
        var sh = it.shape;
        if (useNullFrame) it.tf = sh.getTextFrameOrNullObject();
        else if (SHAPE_TYPES.indexOf(sh.type) >= 0) it.tf = sh.textFrame;
        if (it.tf) it.tf.load("leftMargin,rightMargin,topMargin,bottomMargin,verticalAlignment,autoSizeSetting,wordWrap,hasText");
        if (canRound && sh.type === "GeometricShape") sh.adjustments.load("count");
      });
      return context.sync();
    }).then(function () {
      items.forEach(function (it) {
        var sh = it.shape;
        if (it.tf && it.tf.isNullObject) it.tf = null;
        if (it.tf && it.tf.hasText && ST.textStyles) { it.font = it.tf.textRange.font; it.font.load(ST.textStyles.fontProps()); }
        if (canRound && sh.type === "GeometricShape" && sh.adjustments.count === 1) it.adj = sh.adjustments.get(0);
      });
      return context.sync();
    }).then(function () {
      return items.map(function (it) {
        var sh = it.shape, f = sh.fill, l = sh.lineFormat, v = blank();
        v.fillType = f.type === "Solid" ? "Solid" : f.type === "NoFill" ? "NoFill" : "Other";
        v.fillColor = v.fillType === "Solid" ? norm(f.foregroundColor) : null;
        v.fillTransparency = v.fillType === "Solid" && has(f.transparency) ? round1(f.transparency * 100) / 100 : null;
        v.lineVisible = has(l.visible) ? !!l.visible : null;
        v.lineColor = norm(l.color);
        v.lineWeight = has(l.weight) ? round1(l.weight) : null;
        v.lineDash = l.dashStyle || null;
        v.lineStyle = l.style || null;
        v.lineTransparency = has(l.transparency) ? round1(l.transparency * 100) / 100 : null;
        if (it.adj) v.radius = round1(it.adj.value * Math.min(sh.width, sh.height));
        if (it.tf) {
          v.marginLeft = round1(it.tf.leftMargin); v.marginRight = round1(it.tf.rightMargin);
          v.marginTop = round1(it.tf.topMargin); v.marginBottom = round1(it.tf.bottomMargin);
          v.vAlign = it.tf.verticalAlignment || null;
          v.autoSize = it.tf.autoSizeSetting && it.tf.autoSizeSetting !== "AutoSizeMixed" ? it.tf.autoSizeSetting : null;
          v.wordWrap = has(it.tf.wordWrap) ? !!it.tf.wordWrap : null;
        }
        var textTag = it.textTagObj.isNullObject ? "" : it.textTagObj.value;
        if (it.font && ST.textStyles) v.textStyleId = ST.textStyles.detect(it.font, textTag);
        return {
          shape: sh, id: sh.id, type: sh.type, w: sh.width, h: sh.height, values: v,
          hasFrame: !!it.tf, canRound: !!it.adj, hasText: !!(it.tf && it.tf.hasText),
          tagId: it.tagObj.isNullObject ? 0 : parseInt(it.tagObj.value, 10) || 0, textTag: textTag
        };
      });
    });
  }

  // ---------- Writing shapes ----------
  // opts.keepOverrides: leave settings changed by hand since the last push.
  // opts.keepColors: keep the shape's own fill and outline colours.
  // opts.applyText: also apply the style's text style to the text inside.
  function writeShape(info, style, opts) {
    opts = opts || {};
    var sh = info.shape, P = style.props, A = style.applied, cur = info.values;
    function keep(k) { return opts.keepOverrides && has(A[k]) && !same(cur[k], A[k]) && !same(cur[k], P[k]); }
    function want(k) { return has(P[k]) && !keep(k); }

    // Fill
    if (P.fillType === "NoFill" && !keep("fillType")) sh.fill.clear();
    else if (P.fillType === "Solid" && !keep("fillType")) {
      var fill = (opts.keepColors && cur.fillColor) || (keep("fillColor") && cur.fillColor) || P.fillColor;
      if (fill && (cur.fillType !== "Solid" || !same(cur.fillColor, fill))) sh.fill.setSolidColor(fill);
      if (want("fillTransparency")) sh.fill.transparency = P.fillTransparency;
    }
    // Outline
    if (P.lineVisible === false && !keep("lineVisible")) sh.lineFormat.visible = false;
    else if (P.lineVisible !== false && !keep("lineVisible")) {
      if (P.lineVisible === true) sh.lineFormat.visible = true;
      var line = (opts.keepColors && cur.lineColor && cur.lineVisible) ? cur.lineColor : (keep("lineColor") ? null : P.lineColor);
      if (line) sh.lineFormat.color = line;
      if (want("lineWeight")) sh.lineFormat.weight = P.lineWeight;
      if (want("lineDash")) sh.lineFormat.dashStyle = P.lineDash;
      if (want("lineStyle")) sh.lineFormat.style = P.lineStyle;
      if (want("lineTransparency")) sh.lineFormat.transparency = P.lineTransparency;
    }
    // Corners
    if (want("radius") && info.canRound) sh.adjustments.set(0, ST.cornerValue(info.w, info.h, P.radius));
    // Text box
    if (info.hasFrame) {
      var tf = sh.textFrame;
      if (want("marginLeft")) tf.leftMargin = P.marginLeft;
      if (want("marginRight")) tf.rightMargin = P.marginRight;
      if (want("marginTop")) tf.topMargin = P.marginTop;
      if (want("marginBottom")) tf.bottomMargin = P.marginBottom;
      if (want("vAlign")) tf.verticalAlignment = P.vAlign;
      if (want("autoSize")) tf.autoSizeSetting = P.autoSize;
      if (want("wordWrap")) tf.wordWrap = P.wordWrap;
      if (opts.applyText && has(P.textStyleId) && info.hasText && ST.textStyles) ST.textStyles.writeToShape(sh, P.textStyleId, info.textTag);
    }
    sh.tags.add(SHAPE_TAG, String(style.id));
  }

  // ---------- Selection ----------
  function withSelectedShapes(task) {
    if (!ST.needOffice()) return;
    if (!ST.supports("1.5")) { ST.toast("This PowerPoint version can't work with selected shapes.", true); return; }
    return ST.run(function (context) {
      var shapes = context.presentation.getSelectedShapes();
      shapes.load("items/type");
      return context.sync().then(function () {
        var list = shapes.items.filter(function (s) { return SHAPE_TYPES.indexOf(s.type) >= 0; });
        if (!list.length) { ST.toast(shapes.items.length ? "Shape styles work on shapes, text boxes and placeholders." : "Select a shape first.", true); return; }
        return readShapes(context, list).then(function (infos) { return task(context, infos); });
      });
    });
  }

  function sample() {
    withSelectedShapes(function (context, infos) {
      var v = clone(infos[0].values), notes = [];
      if (v.fillType === "Other") { v.fillType = null; v.fillColor = null; v.fillTransparency = null; notes.push("gradient, picture and pattern fills can't be copied, so fill was left out"); }
      if (v.lineVisible === false) { v.lineColor = v.lineWeight = v.lineDash = v.lineStyle = v.lineTransparency = null; }
      draft = { props: v, linkIds: infos.slice(0, 1) };
      editing = "new";
      render();
      var name = $("ssName");
      if (name) name.focus();
      if (notes.length) ST.toast("Note: " + notes.join(", ") + ".");
    });
  }

  function newStyle() {
    var p = blank();
    p.fillType = "Solid"; p.fillColor = "#0071E3"; p.fillTransparency = 0;
    p.lineVisible = false;
    draft = { props: p, linkIds: null };
    editing = "new";
    render();
    var name = $("ssName");
    if (name) name.focus();
  }

  function apply(style, opts) {
    opts = opts || {};
    withSelectedShapes(function (context, infos) {
      infos.forEach(function (info) { writeShape(info, style, { applyText: true, keepColors: opts.keepColors }); });
      return context.sync().then(function () {
        ST.toast(opts.message || "Applied " + style.name + " to " + ST.plural(infos.length, "shape") + ".");
        refreshStatus();
      });
    });
  }

  function unlink() {
    withSelectedShapes(function (context, infos) {
      infos.forEach(function (info) { if (info.tagId) info.shape.tags.delete(SHAPE_TAG); });
      return context.sync().then(function () { ST.toast("Unlinked. The shape keeps its look but no longer follows a style."); refreshStatus(); });
    });
  }

  // ---------- Classifying ----------
  var MATCH_KEYS = ["fillType", "fillColor", "lineVisible", "lineColor", "lineWeight", "lineDash"];
  function lookalikeMatch(style, v) {
    return [style.applied, style.props].some(function (d) {
      var keys = MATCH_KEYS.filter(function (k) { return has(d[k]); });
      if (!has(d.fillColor) && !has(d.lineColor)) return false;
      if (d.lineVisible === false) keys = keys.filter(function (k) { return k.indexOf("line") !== 0 || k === "lineVisible"; });
      if (d.fillType === "NoFill") keys = keys.filter(function (k) { return k !== "fillColor"; });
      return keys.every(function (k) { return same(v[k], d[k]) || (typeof d[k] === "boolean" && v[k] === d[k]); });
    });
  }
  function classify(info) {
    var s = info.tagId && byId(info.tagId);
    if (s) return { kind: "linked", style: s };
    for (var i = 0; i < styles.length; i++) if (lookalikeMatch(styles[i], info.values)) return { kind: "lookalike", style: styles[i] };
    return { kind: "none" };
  }
  function differs(info, style) {
    var v = info.values, P = style.props;
    return KEYS.some(function (k) {
      if (!has(P[k]) || k === "textStyleId") return false;
      if (k === "radius" && !info.canRound) return false;
      if (["marginLeft", "marginRight", "marginTop", "marginBottom", "vAlign", "autoSize", "wordWrap"].indexOf(k) >= 0 && !info.hasFrame) return false;
      if (k.indexOf("line") === 0 && k !== "lineVisible" && P.lineVisible === false) return false;
      if (k.indexOf("fill") === 0 && k !== "fillType" && P.fillType === "NoFill") return false;
      return typeof P[k] === "boolean" ? v[k] !== P[k] : !same(v[k], P[k]);
    });
  }

  // ---------- Status ----------
  var statusTimer = null;
  function refreshStatus() {
    var el = $("ssStatus");
    if (!el) return;
    if (!ST.inOffice || !ST.supports("1.5") || !ST.isOpen("shapestyles") || ST.levelOf("shapestyles") < 2) { el.hidden = true; return; }
    clearTimeout(statusTimer);
    statusTimer = setTimeout(readStatus, 150);
  }
  function readStatus() {
    var el = $("ssStatus");
    PowerPoint.run(function (context) {
      var shapes = context.presentation.getSelectedShapes();
      shapes.load("items/type");
      return context.sync().then(function () {
        var list = shapes.items.filter(function (s) { return SHAPE_TYPES.indexOf(s.type) >= 0; }).slice(0, 1);
        if (!list.length) return null;
        return readShapes(context, list).then(function (r) { return r[0]; });
      });
    }).then(showStatus, function () { el.hidden = true; });
  }
  function showStatus(info) {
    var el = $("ssStatus");
    el.hidden = false; el.innerHTML = "";
    var text = document.createElement("span");
    el.appendChild(text);
    if (!info) { text.textContent = "Select a shape to see its style."; return; }
    var c = classify(info);
    if (c.kind === "linked") {
      var changed = differs(info, c.style);
      text.innerHTML = "Uses <b></b>" + (changed ? ' <span class="plus" title="Has local changes">+</span>' : "");
      text.querySelector("b").textContent = c.style.name;
      if (changed) addButton(el, "Reset to style", function () { apply(c.style, { message: "Reset to " + c.style.name + "." }); });
      addButton(el, "Unlink", unlink);
    } else if (c.kind === "lookalike") {
      text.innerHTML = "Looks like <b></b>";
      text.querySelector("b").textContent = c.style.name;
      addButton(el, "Link, keep colours", function () { apply(c.style, { keepColors: true, message: "Linked to " + c.style.name + ", keeping its colours." }); });
    } else {
      text.textContent = "No shape style.";
    }
  }
  function addButton(el, label, fn) {
    var b = document.createElement("button");
    b.className = "link-btn"; b.textContent = label;
    b.addEventListener("click", fn);
    el.appendChild(b);
  }

  // ---------- Scanning (normal slides only) ----------
  function scanDeck(context) {
    var slides = context.presentation.slides;
    slides.load("items/id");
    var found = [], groups = [], useGroups = ST.supports("1.8");
    function collect(sh, si, sid, path) {
      if (sh.type === "Group") { if (useGroups) groups.push({ shape: sh, si: si, sid: sid, path: path }); return; }
      if (SHAPE_TYPES.indexOf(sh.type) >= 0) found.push({ shape: sh, si: si, sid: sid, path: path });
    }
    function expand() {
      if (!groups.length) return Promise.resolve();
      var gs = groups; groups = [];
      gs.forEach(function (g) { g.shape.group.shapes.load("items/id,items/type"); });
      return context.sync().then(function () {
        gs.forEach(function (g) { g.shape.group.shapes.items.forEach(function (c) { collect(c, g.si, g.sid, g.path.concat([c.id])); }); });
        return expand();
      });
    }
    return context.sync().then(function () {
      slides.items.forEach(function (sl) { sl.shapes.load("items/id,items/type"); });
      return context.sync();
    }).then(function () {
      slides.items.forEach(function (sl, si) { sl.shapes.items.forEach(function (sh) { collect(sh, si, sl.id, [sh.id]); }); });
      return expand();
    }).then(function () {
      var out = [], i = 0;
      function next() {
        if (i >= found.length) return Promise.resolve();
        var chunk = found.slice(i, i + BATCH);
        i += BATCH;
        return readShapes(context, chunk.map(function (f) { return f.shape; })).then(function (infos) {
          infos.forEach(function (info, k) { info.si = chunk[k].si; info.sid = chunk[k].sid; info.path = chunk[k].path; out.push(info); });
          return next();
        });
      }
      return next().then(function () { return { infos: out, slideCount: slides.items.length }; });
    });
  }
  function getShape(context, sid, path) {
    var shape = context.presentation.slides.getItem(sid).shapes.getItem(path[0]);
    for (var i = 1; i < path.length; i++) shape = shape.group.shapes.getItem(path[i]);
    return shape;
  }

  // ---------- Update all ----------
  function updateAll() {
    if (!ST.needOffice()) return;
    if (!ST.supports("1.5")) { ST.toast("This PowerPoint version can't update shape styles.", true); return; }
    if (!styles.length) { ST.toast("Save a shape style first.", true); return; }
    if (busy) return;
    busy = true; review = { scanning: true }; render();
    PowerPoint.run(function (context) { return scanDeck(context); }).then(function (result) {
      var groups = {};
      styles.forEach(function (s) { groups[s.id] = { style: s, linked: [], look: [], choice: "leave" }; });
      result.infos.forEach(function (info) {
        var c = classify(info);
        if (c.kind === "linked") groups[c.style.id].linked.push(info);
        else if (c.kind === "lookalike") { info.choice = null; groups[c.style.id].look.push(info); }
      });
      var list = styles.map(function (s) { return groups[s.id]; }).filter(function (g) { return (g.linked.length && pending(g.style)) || g.look.length; });
      busy = false;
      if (!list.length) {
        review = null;
        styles.forEach(function (s) { s.applied = clone(s.props); });
        saveStyles();
        ST.toast("Everything is up to date. No lookalikes found on " + ST.plural(result.slideCount, "slide") + ".");
        return;
      }
      review = { groups: list, slideCount: result.slideCount };
      render();
    }).catch(function (e) {
      busy = false; review = null; render();
      ST.toast("Couldn't scan the slides: " + (e && e.message ? e.message : e), true);
    });
  }

  function confirmUpdate() {
    if (!review || !review.groups || busy) return;
    busy = true;
    var jobs = [];
    review.groups.forEach(function (g) {
      if (pending(g.style)) g.linked.forEach(function (info) { jobs.push({ info: info, style: g.style, mode: "linked" }); });
      g.look.forEach(function (info) { var c = info.choice || g.choice; if (c !== "leave") jobs.push({ info: info, style: g.style, mode: c }); });
    });
    PowerPoint.run(function (context) {
      var i = 0;
      function next() {
        if (i >= jobs.length) return Promise.resolve();
        jobs.slice(i, i + BATCH).forEach(function (j) {
          var info = j.info;
          info.shape = getShape(context, info.sid, info.path);
          if (j.mode === "linked") {
            var textChanged = j.style.props.textStyleId !== j.style.applied.textStyleId;
            writeShape(info, j.style, { keepOverrides: true, applyText: textChanged });
          } else {
            writeShape(info, j.style, { keepColors: j.mode === "keep", applyText: true });
          }
        });
        i += BATCH;
        return context.sync().then(next);
      }
      return next();
    }).then(function () {
      styles.forEach(function (s) { s.applied = clone(s.props); });
      busy = false; review = null;
      return saveStyles().then(function () {
        ST.toast(jobs.length ? "Updated " + ST.plural(jobs.length, "shape") + "." : "Styles saved. No shapes needed changing.");
        refreshStatus();
      });
    }).catch(function (e) {
      busy = false;
      ST.toast("Update stopped: " + (e && e.message ? e.message : e) + ". Some shapes may already be updated, use Undo if needed.", true);
      render();
    });
  }

  function showShape(info) {
    ST.run(function (context) {
      context.presentation.setSelectedSlides([info.sid]);
      context.presentation.slides.getItem(info.sid).setSelectedShapes([info.path[0]]);
      return context.sync();
    });
  }

  // ---------- Editor ----------
  function sel(id, options, value, leave) {
    var v = has(value) ? String(value) : "";
    return '<select id="' + id + '">' + (leave === false ? "" : '<option value="">Leave</option>') +
      options.map(function (o) { return '<option value="' + esc(o[0]) + '"' + (String(o[0]) === v ? " selected" : "") + ">" + esc(o[1]) + "</option>"; }).join("") + "</select>";
  }
  function tri(id, value) { return sel(id, [["1", "On"], ["0", "Off"]], has(value) ? (value ? "1" : "0") : ""); }
  function pct(v) { return has(v) ? String(Math.round(v * 100)) : ""; }
  function colorField(idBase, value) {
    var c = norm(value) || "#000000";
    return '<div class="color-field"><input type="color" id="' + idBase + 'Pick" list="paletteList" value="' + c.toLowerCase() + '"><input type="text" id="' + idBase + '" value="' + esc(value) + '" placeholder="Leave"></div>';
  }
  function textStyleOptions(value) {
    var opts = textStyles().map(function (t) { return [String(t.id), t.name]; });
    return '<select id="ssText"><option value="">Leave text alone</option>' +
      opts.map(function (o) { return '<option value="' + o[0] + '"' + (String(value) === o[0] ? " selected" : "") + ">" + esc(o[1]) + "</option>"; }).join("") + "</select>";
  }

  function editorMarkup(name, p, isNew) {
    var fillMode = p.fillType === "Solid" ? "Solid" : p.fillType === "NoFill" ? "NoFill" : "";
    var lineMode = p.lineVisible === true ? "1" : p.lineVisible === false ? "0" : "";
    return '<div class="style-editor">' +
      '<label class="field">Style name<input type="text" id="ssName" maxlength="32" value="' + esc(name) + '" placeholder="e.g. Card"></label>' +

      '<div class="group-label">Fill</div>' +
      '<div class="row"><label class="field">Fill' + sel("ssFillType", [["Solid", "Solid colour"], ["NoFill", "No fill"]], fillMode) + '</label>' +
        '<label class="field ss-fill">Transparency (%)<input type="text" id="ssFillT" inputmode="decimal" value="' + pct(p.fillTransparency) + '" placeholder="Leave"></label></div>' +
      '<label class="field ss-fill">Fill colour' + colorField("ssFillC", p.fillColor) + '</label>' +

      '<div class="group-label">Outline</div>' +
      '<div class="row"><label class="field">Outline' + sel("ssLineOn", [["1", "Line"], ["0", "No line"]], lineMode) + '</label>' +
        '<label class="field ss-line">Weight (pt)<input type="text" id="ssLineW" inputmode="decimal" value="' + esc(p.lineWeight) + '" placeholder="Leave"></label></div>' +
      '<label class="field ss-line">Outline colour' + colorField("ssLineC", p.lineColor) + '</label>' +
      '<div class="row ss-line"><label class="field">Dash' + sel("ssDash", DASHES, p.lineDash) + '</label>' +
        '<label class="field">Line type' + sel("ssCompound", COMPOUNDS, p.lineStyle) + '</label></div>' +
      '<label class="field ss-line">Outline transparency (%)<input type="text" id="ssLineT" inputmode="decimal" value="' + pct(p.lineTransparency) + '" placeholder="Leave"></label>' +

      '<div class="group-label">Corners</div>' +
      '<label class="field">Corner radius (pt), rounded rectangles only<input type="text" id="ssRadius" inputmode="decimal" value="' + esc(p.radius) + '" placeholder="Leave"></label>' +

      '<div class="group-label">Text box</div>' +
      '<div class="row"><label class="field">Left margin (pt)<input type="text" id="ssML" inputmode="decimal" value="' + esc(p.marginLeft) + '" placeholder="Leave"></label>' +
        '<label class="field">Right margin (pt)<input type="text" id="ssMR" inputmode="decimal" value="' + esc(p.marginRight) + '" placeholder="Leave"></label></div>' +
      '<div class="row"><label class="field">Top margin (pt)<input type="text" id="ssMT" inputmode="decimal" value="' + esc(p.marginTop) + '" placeholder="Leave"></label>' +
        '<label class="field">Bottom margin (pt)<input type="text" id="ssMB" inputmode="decimal" value="' + esc(p.marginBottom) + '" placeholder="Leave"></label></div>' +
      '<div class="row"><label class="field">Vertical alignment' + sel("ssVA", VALIGNS, p.vAlign) + '</label>' +
        '<label class="field">Wrap text' + tri("ssWrap", p.wordWrap) + '</label></div>' +
      '<label class="field">Autofit' + sel("ssAuto", AUTOSIZES, p.autoSize) + '</label>' +
      '<label class="field">Text style for the text inside' + textStyleOptions(p.textStyleId) + '</label>' +

      (isNew && draft && draft.linkIds ? '<div class="muted-note">Saving links the sampled shape to this style.</div>' : "") +
      '<div class="btn-row"><button class="btn" id="ssCancel">Cancel</button><button class="btn primary" id="ssSave">' + (isNew ? "Save style" : "Save changes") + '</button></div>' +
      (isNew ? "" : '<button class="btn small danger" id="ssDelete">Delete style</button>') +
    '</div>';
  }

  function readEditor() {
    var p = blank(), err = null;
    function numOf(id, scale, min, max) {
      var t = $(id).value.trim();
      if (!t) return null;
      var v = ST.evaluate(t, null, null);
      if (v === null || v < min || (max !== undefined && v > max)) { err = err || id; return null; }
      return round1(v) / (scale || 1);
    }
    function colorOf(id) {
      var t = $(id).value.trim();
      if (!t) return null;
      var c = norm(t);
      if (!c) err = err || id;
      return c;
    }
    var fm = $("ssFillType").value;
    p.fillType = fm || null;
    if (fm === "Solid") { p.fillColor = colorOf("ssFillC"); p.fillTransparency = numOf("ssFillT", 100, 0, 100); if (!p.fillColor) err = err || "ssFillC"; }
    var lm = $("ssLineOn").value;
    p.lineVisible = lm === "" ? null : lm === "1";
    if (lm !== "0") {
      p.lineColor = colorOf("ssLineC"); p.lineWeight = numOf("ssLineW", 1, 0, 100);
      p.lineDash = $("ssDash").value || null; p.lineStyle = $("ssCompound").value || null;
      p.lineTransparency = numOf("ssLineT", 100, 0, 100);
    }
    p.radius = numOf("ssRadius", 1, 0, 500);
    p.marginLeft = numOf("ssML", 1, 0, 500); p.marginRight = numOf("ssMR", 1, 0, 500);
    p.marginTop = numOf("ssMT", 1, 0, 500); p.marginBottom = numOf("ssMB", 1, 0, 500);
    p.vAlign = $("ssVA").value || null;
    p.autoSize = $("ssAuto").value || null;
    var w = $("ssWrap").value; p.wordWrap = w === "" ? null : w === "1";
    var t = $("ssText").value; p.textStyleId = t ? parseInt(t, 10) : null;
    if (err) return { error: err };
    if (!KEYS.some(function (k) { return has(p[k]); })) return { error: "empty" };
    return { props: p, name: $("ssName").value.trim() };
  }

  function saveEditor() {
    var r = readEditor();
    if (r.error === "empty") { ST.toast("Set at least one thing for the style to change.", true); return; }
    if (r.error) {
      var box = $(r.error);
      ST.toast(/C$/.test(r.error) ? "Colours need a hex code like #0071E3, or a colour name." : "Check that number. Transparency is 0 to 100.", true);
      box.classList.add("err"); box.focus();
      return;
    }
    if (!r.name) { ST.toast("Give the style a name.", true); $("ssName").focus(); return; }
    var clash = styles.filter(function (s) { return s.name.toLowerCase() === r.name.toLowerCase() && s.id !== editing; })[0];
    if (clash) { ST.toast("There's already a shape style called " + clash.name + ".", true); return; }
    if (editing === "new") {
      var s = { id: nextId(), name: r.name, props: r.props, applied: clone(r.props) };
      var link = draft && draft.linkIds;
      styles.push(s);
      editing = null; draft = null;
      saveStyles().then(function () {
        if (link) apply(s, { message: "Saved " + s.name + " and applied it to the sampled shape." });
        else ST.toast("Saved " + s.name + ". Select shapes and click it to apply.");
      });
    } else {
      var st = byId(editing);
      st.name = r.name; st.props = r.props;
      editing = null;
      saveStyles();
      ST.toast(pending(st) ? "Saved. Click Update all to change the shapes that use " + st.name + "." : "Saved " + st.name + ".");
    }
  }

  function wireEditor(id) {
    [["ssFillC", "ssFillCPick"], ["ssLineC", "ssLineCPick"]].forEach(function (pair) {
      var text = $(pair[0]), pick = $(pair[1]);
      pick.addEventListener("input", function () { text.value = pick.value.toUpperCase(); text.classList.remove("err"); });
      text.addEventListener("input", function () { var c = norm(text.value); if (c) pick.value = c.toLowerCase(); text.classList.remove("err"); });
    });
    function sync() {
      var fill = $("ssFillType").value === "Solid", line = $("ssLineOn").value !== "0";
      Array.prototype.forEach.call(document.querySelectorAll(".ss-fill"), function (el) { el.hidden = !fill; });
      Array.prototype.forEach.call(document.querySelectorAll(".ss-line"), function (el) { el.hidden = !line; });
    }
    $("ssFillType").addEventListener("change", function () {
      if ($("ssFillType").value === "Solid" && !$("ssFillC").value) { $("ssFillC").value = "#0071E3"; $("ssFillCPick").value = "#0071e3"; }
      sync();
    });
    $("ssLineOn").addEventListener("change", sync);
    sync();
    $("ssCancel").addEventListener("click", function () { editing = null; draft = null; render(); });
    $("ssSave").addEventListener("click", saveEditor);
    $("ssName").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); saveEditor(); } });
    if (id) $("ssDelete").addEventListener("click", function () {
      var s = byId(id);
      styles = styles.filter(function (x) { return x.id !== id; });
      editing = null;
      saveStyles();
      ST.toast("Deleted " + s.name + ". Shapes keep their look.");
    });
  }

  // ---------- Rendering ----------
  function summary(p) {
    var bits = [];
    if (p.fillType === "NoFill") bits.push("no fill");
    else if (p.fillColor) bits.push(p.fillColor + (p.fillTransparency ? " " + Math.round(p.fillTransparency * 100) + "% clear" : ""));
    if (p.lineVisible === false) bits.push("no outline");
    else if (p.lineColor || p.lineWeight) bits.push((p.lineWeight ? p.lineWeight + " pt " : "") + "outline" + (p.lineColor ? " " + p.lineColor : ""));
    if (has(p.radius)) bits.push(p.radius + " pt corners");
    if (has(p.marginLeft) || has(p.marginTop)) bits.push("margins");
    var t = textStyleName(p.textStyleId);
    if (t) bits.push("text: " + t);
    return bits.join(", ") || "No settings";
  }
  function tileStyle(p) {
    var css = "";
    if (p.fillType === "Solid" && p.fillColor) {
      var c = norm(p.fillColor), a = 1 - (p.fillTransparency || 0);
      var r = parseInt(c.substr(1, 2), 16), g = parseInt(c.substr(3, 2), 16), b = parseInt(c.substr(5, 2), 16);
      css += "background:rgba(" + r + "," + g + "," + b + "," + a + ");";
    } else css += "background:transparent;";
    if (p.lineVisible !== false && (p.lineColor || p.lineWeight)) {
      var dash = p.lineDash && p.lineDash !== "Solid" ? (/Dot$/.test(p.lineDash) ? "dotted" : "dashed") : (p.lineStyle && p.lineStyle !== "Single" ? "double" : "solid");
      css += "border:" + Math.max(1, Math.min(4, p.lineWeight || 1)) + "px " + dash + " " + (p.lineColor || "#888") + ";";
    } else css += "border:1px dashed var(--field-line);";
    if (has(p.radius)) css += "border-radius:" + Math.min(6, p.radius / 4) + "px;";
    return css;
  }

  function render() {
    var list = $("ssList"), form = $("ssForm"), rev = $("ssReview");
    if (!list) return;
    form.innerHTML = ""; rev.innerHTML = ""; rev.hidden = true; list.hidden = false; $("ssActions").hidden = false;
    if (review) { renderReview(rev); list.hidden = true; $("ssActions").hidden = true; return; }
    if (editing === "new") { form.innerHTML = editorMarkup("", draft.props, true); wireEditor(); }
    list.innerHTML = "";
    if (!stylesLoaded && ST.inOffice) { list.innerHTML = '<div class="muted-note">Reading shape styles from this presentation.</div>'; return; }
    if (!styles.length && editing !== "new") list.innerHTML = '<div class="muted-note">No shape styles yet. Select a shape and click Sample to make one from it.</div>';
    styles.forEach(function (s) {
      var row = document.createElement("div");
      row.className = "style-row";
      row.innerHTML =
        '<button class="style-apply" title="Apply to the selected shapes"><span class="shape-tile"><i></i></span>' +
          '<span class="style-text"><span class="style-name"></span><span class="style-sum" data-min="2"></span></span>' +
          (pending(s) ? '<span class="badge" title="Edited since the last Update all">Edited</span>' : "") +
        '</button><button class="icon-btn" title="Edit style" aria-label="Edit style">' + EDIT_ICON + "</button>";
      row.querySelector(".shape-tile i").setAttribute("style", tileStyle(s.props));
      row.querySelector(".style-name").textContent = s.name;
      row.querySelector(".style-sum").textContent = summary(s.props);
      row.querySelector(".style-apply").addEventListener("click", function () { apply(s); });
      row.querySelector(".icon-btn").addEventListener("click", function () { editing = editing === s.id ? null : s.id; render(); });
      list.appendChild(row);
      if (editing === s.id) {
        var holder = document.createElement("div");
        holder.innerHTML = editorMarkup(s.name, s.props, false);
        list.appendChild(holder);
        wireEditor(s.id);
      }
    });
    var n = styles.filter(pending).length;
    $("ssUpdate").textContent = n ? "Update all (" + n + " edited)" : "Update all";
  }

  var CHOICES = [["leave", "Leave them alone"], ["full", "Apply style"], ["keep", "Apply style, keep their colours"]];
  function renderReview(el) {
    el.hidden = false;
    if (review.scanning) { el.innerHTML = '<div class="muted-note">Checking every slide. Masters and layouts are left alone.</div>'; return; }
    el.innerHTML = '<div class="group-label">Update all, ' + ST.plural(review.slideCount, "slide") + " checked</div>";
    review.groups.forEach(function (g) {
      var box = document.createElement("div");
      box.className = "review-group";
      var slides = {};
      g.linked.forEach(function (i) { slides[i.si] = 1; });
      var html = '<div class="review-title"></div>';
      if (g.linked.length) html += '<div class="review-line">' + (pending(g.style)
        ? ST.plural(g.linked.length, "linked shape") + " on " + ST.plural(Object.keys(slides).length, "slide") + " will update. Local changes are kept."
        : ST.plural(g.linked.length, "linked shape") + ", nothing to change.") + "</div>";
      if (g.look.length) html += '<div class="review-line">' + ST.plural(g.look.length, "shape") + (g.look.length === 1 ? " looks like this style but isn't linked." : " look like this style but aren't linked.") + "</div>" +
        '<label class="field">For those<select class="look-choice">' + CHOICES.map(function (c) { return '<option value="' + c[0] + '">' + c[1] + "</option>"; }).join("") + "</select></label>" +
        '<button class="link-btn review-toggle">Review each one</button><div class="review-items" hidden></div>';
      box.innerHTML = html;
      box.querySelector(".review-title").textContent = g.style.name;
      var s = box.querySelector(".look-choice");
      if (s) {
        s.value = g.choice;
        s.addEventListener("change", function () { g.choice = s.value; });
        var items = box.querySelector(".review-items");
        box.querySelector(".review-toggle").addEventListener("click", function () {
          items.hidden = !items.hidden;
          if (!items.hidden && !items.children.length) g.look.forEach(function (info) {
            var row = document.createElement("div");
            row.className = "review-item";
            row.innerHTML = '<button class="link-btn item-show"></button><select class="item-choice"><option value="">Same as above</option>' +
              CHOICES.map(function (c) { return '<option value="' + c[0] + '">' + c[1] + "</option>"; }).join("") + "</select>";
            row.querySelector(".item-show").textContent = "Slide " + (info.si + 1) + ": " + (info.type === "TextBox" ? "text box" : info.type === "Placeholder" ? "placeholder" : "shape") + (info.values.fillColor ? " " + info.values.fillColor : "");
            row.querySelector(".item-show").title = "Select this shape on the slide";
            row.querySelector(".item-show").addEventListener("click", function () { showShape(info); });
            var c = row.querySelector(".item-choice");
            c.value = info.choice || "";
            c.addEventListener("change", function () { info.choice = c.value || null; });
            items.appendChild(row);
          });
        });
      }
      el.appendChild(box);
    });
    var actions = document.createElement("div");
    actions.className = "btn-row";
    actions.innerHTML = '<button class="btn" id="ssReviewCancel">Cancel</button><button class="btn primary" id="ssReviewGo">Update</button>';
    el.appendChild(actions);
    $("ssReviewCancel").addEventListener("click", function () { review = null; render(); });
    $("ssReviewGo").addEventListener("click", confirmUpdate);
  }

  // ---------- Build ----------
  var MARKUP =
    '<span class="info res-info" tabindex="0" role="button" aria-label="How shape styles work">i' +
      '<span class="tip" role="tooltip">Select a shape and click Sample to save its fill, outline, corners and text box settings as a shape style. Click a style to apply it to the selected shapes.<br><br>' +
      'A shape style can also apply one of your text styles to the text inside.<br><br>' +
      'Edit a style, then click Update all to change every shape that uses it. Settings you changed by hand on linked shapes are kept. Shapes that look like a style but aren\'t linked are offered too.<br><br>' +
      'Styles are saved in this presentation. Only normal slides are changed, never masters or layouts. Gradients, shadows and other effects can\'t be set by add-ins.</span>' +
    '</span>' +
    '<div class="style-status" id="ssStatus" data-min="2" hidden></div>' +
    '<div id="ssForm"></div>' +
    '<div class="style-list" id="ssList"></div>' +
    '<div id="ssReview" class="ss-review" hidden></div>' +
    '<div id="ssActions" class="ss-actions">' +
      '<div class="btn-row"><button class="btn" id="ssSample" title="Make a style from the selected shape">Sample</button>' +
      '<button class="btn" id="ssUpdate" title="Push style edits to every slide">Update all</button></div>' +
      '<button class="btn small" id="ssNew" data-min="2">New blank style</button>' +
    '</div>';

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    $("ssSample").addEventListener("click", sample);
    $("ssNew").addEventListener("click", newStyle);
    $("ssUpdate").addEventListener("click", updateAll);
    ST.on("ready", loadStyles);
    ST.on("selection", refreshStatus);
    ST.on("layout", refreshStatus);
    ST.on("textstyles", function () {
      var box = $("ssText");
      if (box) { var v = box.value; box.outerHTML = textStyleOptions(v); }
      if (!editing) render();
    });
    render();
  }

  ST.registerTool({
    id: "shapestyles",
    title: "Shape styles",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 2,
    defaults: {},
    build: build
  });

  // Exposed for testing only
  ST._shapeStyles = { get: function () { return styles; }, set: function (s) { styles = s; stylesLoaded = true; render(); }, classify: classify, norm: norm };
})();
