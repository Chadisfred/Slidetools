/* Consistency check: one scan of normal slides that lists layout and formatting drift,
   with a click to jump to each item and Fix buttons where a fix is safe.
   Masters and layouts are never read or changed. */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;

  var defaults = { chkGrid: true, chkText: true, chkLook: true, chkRadius: true, chkColours: true, chkTol: "0.1" };
  ST.addDefaults(defaults);

  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="8" cy="8" r="5.5"/><path d="M12.2 12.2L16 16"/><path d="M5.6 8.1l1.7 1.7 3.2-3.4"/></svg>';
  var CHECKS = [
    ["chkGrid", "grid", "Nearly on the grid", "Shapes whose left or top edge is just off a Layout grid line or a step of g."],
    ["chkText", "text", "Text with no style", "Text whose look matches none of your text styles."],
    ["chkLook", "look", "Unlinked lookalikes", "Text and shapes that look like a style but aren't linked."],
    ["chkRadius", "radius", "Odd corner radii", "Rounded rectangles whose radius differs from the most common one."],
    ["chkColours", "colours", "Colours not in your palette", "Fills, outlines and text in colours outside your palette, theme and styles."]
  ];
  var EPS = 0.004; // inches: closer than this counts as on the line
  var MAX_ITEMS = 60;

  var results = null, running = false;

  // ---------- Helpers ----------
  function tol() { var v = ST.evaluate(ST.state.chkTol, ST.formulaVars(), null); return v > 0 ? v : 0.1; }
  function gridTargets() {
    var g = ST.calculate(), xs = [], ys = [], x = g.left, y = g.top, c, r;
    if (!g.valid) return null;
    for (c = 0; c < g.cols; c++) { xs.push(x, x + g.widths[c]); x += g.widths[c] + g.gapH; }
    for (r = 0; r < g.rows; r++) { ys.push(y, y + g.cellH); y += g.cellH + g.gapV; }
    return { xs: xs, ys: ys, step: g.spacing, ox: g.left, oy: g.top };
  }
  // Distance (inches, signed) from v to the nearest grid line or step of g.
  function offset(v, lines, step, origin) {
    var best = null;
    lines.forEach(function (l) { var d = l - v; if (best === null || Math.abs(d) < Math.abs(best)) best = d; });
    if (step > 0) { var f = origin + Math.round((v - origin) / step) * step - v; if (best === null || Math.abs(f) < Math.abs(best)) best = f; }
    return best === null ? 0 : best;
  }
  function knownColours() {
    var set = {};
    (ST.palette ? ST.palette.list() : []).forEach(function (c) { set[c.hex] = 1; });
    (ST.textStyles ? ST.textStyles.list() : []).forEach(function (s) { if (s.font.color) set[base(s.font.color)] = 1; });
    (ST.shapeStyles ? ST.shapeStyles.list() : []).forEach(function (s) {
      if (s.props.fillColor) set[base(s.props.fillColor)] = 1;
      if (s.props.lineColor) set[base(s.props.lineColor)] = 1;
    });
    return set;
  }
  function base(c) { return ST.styleCode ? ST.styleCode.base(c) : String(c).toUpperCase(); }
  function known(set, c) {
    if (!c) return true;
    c = String(c).toUpperCase();
    return !!(set[c] || set[base(c)] || Object.keys(set).some(function (k) { return base(k) === base(c); }));
  }
  function snippet(t) { t = String(t || "").replace(/\s+/g, " ").trim(); return t.length > 28 ? t.slice(0, 27) + "…" : t; }
  function shapeWord(info) { return info.type === "TextBox" ? "Text box" : info.type === "Placeholder" ? "Placeholder" : "Shape"; }

  // ---------- Running ----------
  function run() {
    if (!ST.needOffice()) return;
    if (!ST.supports("1.5")) { ST.toast("This PowerPoint version can't run the check.", true); return; }
    if (running) return;
    var on = {};
    CHECKS.forEach(function (c) { on[c[1]] = !!ST.state[c[0]]; });
    if (!Object.keys(on).some(function (k) { return on[k]; })) { ST.toast("Turn on at least one check.", true); return; }
    running = true; results = { scanning: true }; render();
    var needText = on.text || on.look || on.colours, needShapes = on.grid || on.look || on.radius || on.colours;
    var textRes = null, shapeRes = null;
    PowerPoint.run(function (context) {
      var step = Promise.resolve();
      if (needText && ST.textStyles) step = step.then(function () { return ST.textStyles.scan(context).then(function (r) { textRes = r; }); });
      if (needShapes && ST.shapeStyles) step = step.then(function () { return ST.shapeStyles.scan(context).then(function (r) { shapeRes = r; }); });
      return step;
    }).then(function () {
      running = false;
      results = build(on, textRes, shapeRes);
      render();
      var total = results.groups.reduce(function (n, g) { return n + g.items.length; }, 0);
      ST.toast(total ? "Found " + ST.plural(total, "thing") + " to look at on " + ST.plural(results.slideCount, "slide") + "." : "All clear on " + ST.plural(results.slideCount, "slide") + ".");
    }).catch(function (e) {
      running = false; results = null; render();
      ST.toast("Couldn't check the slides: " + (e && e.message ? e.message : e), true);
    });
  }

  function build(on, textRes, shapeRes) {
    var groups = [], slideCount = (shapeRes || textRes || { slideCount: 0 }).slideCount;
    var infos = shapeRes ? shapeRes.infos : [], segs = textRes ? textRes.segs.filter(function (s) { return /\S/.test(s.text); }) : [];
    function group(key) { var c = CHECKS.filter(function (x) { return x[1] === key; })[0]; var g = { key: key, title: c[2], items: [] }; groups.push(g); return g; }

    if (on.grid) {
      var g = group("grid"), T = gridTargets(), t = tol();
      if (!T) g.note = "The Layout grid isn't valid, so this check was skipped.";
      else infos.forEach(function (info) {
        var dx = offset(info.x / PT, T.xs, T.step, T.ox), dy = offset(info.y / PT, T.ys, T.step, T.oy);
        var fx = Math.abs(dx) > EPS && Math.abs(dx) <= t, fy = Math.abs(dy) > EPS && Math.abs(dy) <= t;
        if (!fx && !fy) return;
        var parts = [];
        if (fx) parts.push("left " + ST.nice(Math.abs(dx)) + " in off");
        if (fy) parts.push("top " + ST.nice(Math.abs(dy)) + " in off");
        g.items.push({ si: info.si, label: shapeWord(info) + ", " + parts.join(", "), select: shapeSelect(info),
          fixLabel: "Snap", fix: function (context) {
            var sh = ST.shapeStyles.getShape(context, info.sid, info.path);
            if (fx) sh.left = info.x + dx * PT;
            if (fy) sh.top = info.y + dy * PT;
          } });
      });
    }
    if (on.text && ST.textStyles) {
      var gt = group("text");
      if (!ST.textStyles.list().length) gt.note = "You have no text styles yet, so this check was skipped.";
      else groupText(segs.filter(function (seg) { return ST.textStyles.classifySeg(seg).kind === "none"; }),
        function (f) { return (f.name || "mixed font") + (f.size ? " " + f.size + " pt" : ""); }).forEach(function (it) { gt.items.push(it); });
    }
    if (on.look) {
      var gl = group("look");
      segs.forEach(function (seg) {
        var c = ST.textStyles.classifySeg(seg);
        if (c.kind !== "lookalike") return;
        gl.items.push({ si: seg.e.si, label: "Text “" + snippet(seg.text) + "” looks like " + c.style.name, select: textSelect(seg),
          fixLabel: "Link", fix: function (context) { ST.textStyles.linkSeg(context, seg, c.style.id, true); } });
      });
      infos.forEach(function (info) {
        var c = ST.shapeStyles.classify(info);
        if (c.kind !== "lookalike") return;
        gl.items.push({ si: info.si, label: shapeWord(info) + " looks like " + c.style.name, select: shapeSelect(info),
          fixLabel: "Link", fix: function (context) { ST.shapeStyles.link(context, info, c.style.id); } });
      });
    }
    if (on.radius) {
      var gr = group("radius"), counts = {}, rounded = infos.filter(function (i) { return i.canRound && i.values.radius !== null; });
      rounded.forEach(function (i) { var k = Math.round(i.values.radius * 2) / 2; counts[k] = (counts[k] || 0) + 1; });
      var common = Object.keys(counts).sort(function (a, b) { return counts[b] - counts[a]; })[0];
      if (rounded.length < 2) gr.note = "Fewer than two rounded rectangles, so there's nothing to compare.";
      else {
        common = parseFloat(common);
        gr.note = "Most common radius: " + common + " pt.";
        rounded.forEach(function (info) {
          if (Math.abs(info.values.radius - common) < 0.5) return;
          gr.items.push({ si: info.si, label: shapeWord(info) + ", " + info.values.radius + " pt radius", select: shapeSelect(info),
            fixLabel: "Set " + common + " pt", fix: function (context) {
              var sh = ST.shapeStyles.getShape(context, info.sid, info.path);
              sh.adjustments.set(0, ST.cornerValue(info.w, info.h, common));
            } });
        });
      }
    }
    if (on.colours) {
      var gc = group("colours"), set = knownColours();
      if (!ST.palette || !ST.palette.list().length) gc.note = "You have no palette or theme colours, so this check was skipped.";
      else {
        infos.forEach(function (info) {
          var v = info.values, bad = [];
          if (v.fillType === "Solid" && !known(set, v.fillColor)) bad.push("fill " + v.fillColor);
          if (v.lineVisible && v.lineColor && !known(set, v.lineColor)) bad.push("outline " + v.lineColor);
          if (bad.length) gc.items.push({ si: info.si, label: shapeWord(info) + ", " + bad.join(", "), select: shapeSelect(info) });
        });
        groupText(segs.filter(function (seg) { return seg.font.color && !known(set, seg.font.color); }),
          function (f) { return f.color; }, "Text ").forEach(function (it) { gc.items.push(it); });
      }
    }
    groups.forEach(function (g) { g.items.sort(function (a, b) { return a.si - b.si; }); g.more = Math.max(0, g.items.length - MAX_ITEMS); });
    return { groups: groups, slideCount: slideCount };
  }

  // One item per text box and description, so a paragraph split by a bold word is listed once.
  function groupText(list, describe, prefix) {
    var map = {}, order = [];
    list.forEach(function (seg) {
      var key = seg.e.sid + "/" + seg.e.path.join("/") + "/" + describe(seg.font);
      if (!map[key]) { map[key] = { segs: [], desc: describe(seg.font) }; order.push(key); }
      map[key].segs.push(seg);
    });
    return order.map(function (k) {
      var g = map[k], first = g.segs[0], more = g.segs.length - 1;
      return { si: first.e.si, label: (prefix || "") + "“" + snippet(first.text) + "”" + (more ? " and " + more + " more" : "") + ", " + g.desc,
        select: more ? function (context) {
          context.presentation.setSelectedSlides([first.e.sid]);
          context.presentation.slides.getItem(first.e.sid).setSelectedShapes([first.e.path[0]]);
        } : textSelect(first) };
    });
  }

  function shapeSelect(info) {
    return function (context) {
      context.presentation.setSelectedSlides([info.sid]);
      context.presentation.slides.getItem(info.sid).setSelectedShapes([info.path[0]]);
    };
  }
  function textSelect(seg) {
    return function (context) {
      context.presentation.setSelectedSlides([seg.e.sid]);
      var shape = ST.shapeStyles ? ST.shapeStyles.getShape(context, seg.e.sid, seg.e.path) : context.presentation.slides.getItem(seg.e.sid).shapes.getItem(seg.e.path[0]);
      shape.textFrame.textRange.getSubstring(seg.start, seg.len).setSelected();
    };
  }

  function doFix(items, label) {
    var todo = items.filter(function (i) { return i.fix && !i.fixed; });
    if (!todo.length) return;
    ST.run(function (context) {
      todo.forEach(function (i) { i.fix(context); });
      return context.sync().then(function () {
        todo.forEach(function (i) { i.fixed = true; });
        render();
        ST.toast(label || "Fixed " + ST.plural(todo.length, "item") + ".");
      });
    });
  }
  function select(item) {
    ST.run(function (context) { item.select(context); return context.sync(); });
  }

  // ---------- Rendering ----------
  function render() {
    var out = $("chkResults");
    if (!out) return;
    out.innerHTML = "";
    if (!results) return;
    if (results.scanning) { out.innerHTML = '<div class="muted-note">Checking every slide. Masters and layouts are left alone.</div>'; return; }
    results.groups.forEach(function (g) {
      var box = document.createElement("div");
      box.className = "review-group";
      var open = g.items.filter(function (i) { return !i.fixed; });
      box.innerHTML = '<div class="chk-head"><span class="review-title"></span><span class="chk-count"></span></div>';
      box.querySelector(".review-title").textContent = g.title;
      var count = box.querySelector(".chk-count");
      count.textContent = open.length ? String(open.length) : "✓";
      count.className = "chk-count" + (open.length ? "" : " ok");
      if (g.note) box.appendChild(Object.assign(document.createElement("div"), { className: "review-line", textContent: g.note }));
      var list = document.createElement("div");
      list.className = "review-items";
      g.items.slice(0, MAX_ITEMS).forEach(function (item) {
        var row = document.createElement("div");
        row.className = "review-item chk-item" + (item.fixed ? " fixed" : "");
        var b = document.createElement("button");
        b.className = "link-btn"; b.textContent = "Slide " + (item.si + 1) + ": " + item.label; b.title = "Select it on the slide";
        b.addEventListener("click", function () { select(item); });
        row.appendChild(b);
        if (item.fix) {
          var f = document.createElement("button");
          f.className = "btn small"; f.textContent = item.fixed ? "Fixed" : item.fixLabel; f.disabled = !!item.fixed;
          f.addEventListener("click", function () { doFix([item]); });
          row.appendChild(f);
        }
        list.appendChild(row);
      });
      if (g.more) list.appendChild(Object.assign(document.createElement("div"), { className: "review-line", textContent: "And " + g.more + " more. Fix some and run the check again." }));
      if (g.items.length) box.appendChild(list);
      var fixable = open.filter(function (i) { return i.fix; });
      if (fixable.length > 1) {
        var all = document.createElement("button");
        all.className = "btn small"; all.textContent = "Fix all " + fixable.length;
        all.addEventListener("click", function () { doFix(fixable); });
        box.appendChild(all);
      }
      out.appendChild(box);
    });
  }

  var MARKUP =
    '<span class="info res-info" tabindex="0" role="button" aria-label="How the check works">i' +
      '<span class="tip" role="tooltip">Checks every normal slide for small inconsistencies and lists them by slide. Click an item to select it.<br><br>' +
      'Snap moves an edge onto the grid, Link joins a lookalike to its style while keeping its colour, and Set changes a corner radius to the most common one. Other items are for you to review.<br><br>' +
      'Nothing changes until you click a fix. Masters and layouts are never checked or changed.</span>' +
    '</span>' +
    '<button class="btn primary" id="chkRun">Check slides</button>' +
    '<div class="chk-options" data-min="2"></div>' +
    '<label class="field" data-min="2">Grid tolerance (in), how far off still counts as nearly on<input type="text" id="chkTol" inputmode="decimal"></label>' +
    '<div class="chk-results" id="chkResults"></div>';

  function buildTool(body) {
    body.appendChild(ST.html(MARKUP));
    var opts = body.querySelector(".chk-options");
    CHECKS.forEach(function (c) {
      var row = document.createElement("div");
      row.className = "toggle";
      row.title = c[3];
      row.innerHTML = '<span></span><label class="switch"><input type="checkbox"><span></span></label>';
      row.querySelector("span").textContent = c[2];
      var box = row.querySelector("input");
      box.checked = !!ST.state[c[0]];
      box.addEventListener("change", function () { ST.state[c[0]] = box.checked; ST.save(); });
      ST.on("reset", function () { box.checked = !!ST.state[c[0]]; });
      opts.appendChild(row);
    });
    ST.bindField($("chkTol"), "chkTol", { number: true });
    $("chkRun").addEventListener("click", run);
  }

  ST.registerTool({
    id: "check",
    title: "Consistency check",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 1,
    defaults: defaults,
    build: buildTool
  });

  // Exposed for testing only
  ST._check = { results: function () { return results; }, offset: offset };
})();
