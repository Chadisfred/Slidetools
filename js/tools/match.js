/* Select similar tool (id "match"): select shapes like the selected one, slide by slide.
   Copy and paste position lives in the Position tool. */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;

  var defaults = { matchBy: "fill,type" };
  ST.addDefaults(defaults);

  var CRITERIA = [["fill", "Fill colour"], ["line", "Outline"], ["shapeStyle", "Shape style"], ["textStyle", "Text style"], ["size", "Size"], ["type", "Type"]];
  var SIZE_TOLERANCE = 0.5; // points

  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="2" y="2" width="8" height="8" rx="1.5"/><rect x="8" y="8" width="8" height="8" rx="1.5" fill="currentColor" fill-opacity="0.25"/></svg>';

  var MARKUP =
    '<div class="sublabel label-row">Match by' + ST.info("Shapes must match on every chip that's on. Size allows half a point of difference. Pictures and tables only match by Type and Size.", "About match by") + '</div>' +
    '<div class="mini-chips" id="matchBy"></div>' +
    '<button class="btn primary" id="matchSelect">Select similar on slide</button>' +
    '<div class="btn-row" data-min="2"><button class="btn small" id="matchPrev" title="Go to the previous slide with matches and select them">Previous slide</button><button class="btn small" id="matchNext" title="Go to the next slide with matches and select them">Next slide</button></div>' +
    '<div class="help" id="matchStatus"></div>';

  // ---------- State kept for this session only ----------
  var reference = null; // signature of the example shape

  function by() { return String(ST.state.matchBy || "").split(",").filter(Boolean); }

  // ---------- Reading and comparing ----------
  // Reads what matching needs for a list of shapes. Only shape kinds the Shape styles reader
  // understands get fill, outline and style details. Pictures, tables and groups are compared
  // by type and size only.
  function readInfos(context, shapes, needDetails) {
    shapes.forEach(function (s) { s.load("id,type,left,top,width,height"); });
    return context.sync().then(function () {
      var infos = shapes.map(function (s) { return { id: s.id, type: s.type, w: s.width, h: s.height, values: null, tagId: 0 }; });
      var reader = ST.shapeReader;
      if (!needDetails || !reader) return infos;
      var readable = [];
      shapes.forEach(function (s, i) { if (reader.types.indexOf(s.type) >= 0) readable.push(i); });
      var at = 0;
      function next() {
        if (at >= readable.length) return Promise.resolve(infos);
        var chunk = readable.slice(at, at + reader.batch);
        at += reader.batch;
        return reader.read(context, chunk.map(function (i) { return shapes[i]; })).then(function (out) {
          out.forEach(function (r, k) { infos[chunk[k]].values = r.values; infos[chunk[k]].tagId = r.tagId; });
          return next();
        });
      }
      return next();
    });
  }
  function needsDetails(list) { return list.some(function (c) { return c !== "size" && c !== "type"; }); }

  function signature(info) {
    var v = info.values;
    return {
      type: info.type, w: info.w, h: info.h,
      fill: v ? v.fillType + ":" + (v.fillColor || "") : "",
      line: v ? (v.lineVisible ? "on:" + (v.lineColor || "") + ":" + (v.lineWeight || "") + ":" + (v.lineDash || "") : "off") : "",
      shapeStyle: info.tagId || 0,
      textStyle: v && v.textStyleId ? v.textStyleId : 0
    };
  }
  function matches(ref, sig, list) {
    return list.every(function (c) {
      if (c === "size") return Math.abs(ref.w - sig.w) <= SIZE_TOLERANCE && Math.abs(ref.h - sig.h) <= SIZE_TOLERANCE;
      return ref[c] === sig[c];
    });
  }

  function canSelect() { return ST.supports("1.5"); }
  function setStatus(text) { $("matchStatus").textContent = text; }

  // ---------- Select similar ----------
  function selectOnSlide() {
    if (!ST.needOffice()) return;
    if (!canSelect()) { ST.toast("This PowerPoint version can't select shapes from the panel.", true); return; }
    var list = by();
    if (!list.length) { ST.toast("Tick at least one setting to match by.", true); return; }
    ST.run(function (context) {
      var selected = context.presentation.getSelectedShapes();
      var slides = context.presentation.getSelectedSlides();
      selected.load("items/id");
      slides.load("items/id");
      return context.sync().then(function () {
        if (!selected.items.length) { ST.toast("Select an example shape first.", true); return; }
        if (!slides.items.length) { ST.toast("Go to a slide first.", true); return; }
        var slide = slides.items[0];
        slide.shapes.load("items/id");
        return context.sync().then(function () {
          var shapes = slide.shapes.items.slice();
          var details = needsDetails(list);
          return readInfos(context, [selected.items[0]].concat(shapes), details).then(function (infos) {
            reference = signature(infos[0]);
            var ids = infos.slice(1).filter(function (inf) { return matches(reference, signature(inf), list); }).map(function (inf) { return inf.id; });
            if (!ids.length) { setStatus("No matches on this slide."); ST.toast("No shapes on this slide match.", true); return; }
            slide.setSelectedShapes(ids);
            return context.sync().then(function () {
              setStatus(ST.plural(ids.length, "match") + " on this slide.");
              ST.toast("Selected " + ST.plural(ids.length, "matching shape") + ".");
              ST.refreshSelection();
            });
          });
        });
      });
    });
  }

  function step(dir) {
    if (!ST.needOffice()) return;
    if (!canSelect()) { ST.toast("This PowerPoint version can't select shapes from the panel.", true); return; }
    if (!reference) { ST.toast("Click Select similar on slide first to set the example.", true); return; }
    var list = by();
    if (!list.length) { ST.toast("Tick at least one setting to match by.", true); return; }
    ST.run(function (context) {
      var all = context.presentation.slides, current = context.presentation.getSelectedSlides();
      all.load("items/id");
      current.load("items/id");
      return context.sync().then(function () {
        var n = all.items.length;
        if (!n) return;
        var curId = current.items.length ? current.items[0].id : all.items[0].id;
        var idx = Math.max(0, all.items.map(function (s) { return s.id; }).indexOf(curId));
        all.items.forEach(function (sl) { sl.shapes.load("items/id"); });
        return context.sync().then(function () {
          var flat = [], owner = [];
          all.items.forEach(function (sl, si) { sl.shapes.items.forEach(function (sh) { flat.push(sh); owner.push(si); }); });
          return readInfos(context, flat, needsDetails(list)).then(function (infos) {
            var perSlide = all.items.map(function () { return []; });
            infos.forEach(function (inf, k) { if (matches(reference, signature(inf), list)) perSlide[owner[k]].push(inf.id); });
            var total = 0, withMatches = 0;
            perSlide.forEach(function (ids) { total += ids.length; if (ids.length) withMatches++; });
            if (!total) { setStatus("No matches in the presentation."); ST.toast("No shapes in the presentation match.", true); return; }
            var target = -1;
            for (var i = 1; i <= n; i++) {
              var j = ((idx + dir * i) % n + n) % n;
              if (perSlide[j].length) { target = j; break; }
            }
            var slide = all.items[target];
            context.presentation.setSelectedSlides([slide.id]);
            return context.sync().then(function () {
              slide.setSelectedShapes(perSlide[target]);
              return context.sync();
            }).then(function () {
              var summary = ST.plural(perSlide[target].length, "match") + " on slide " + (target + 1) + ". " + total + " on " + ST.plural(withMatches, "slide") + ".";
              setStatus(summary);
              ST.toast(target === idx ? "Only this slide has matches." : "Slide " + (target + 1) + ": selected " + ST.plural(perSlide[target].length, "shape") + ".");
              ST.refreshSelection();
            });
          });
        });
      });
    });
  }

  // ---------- Rendering ----------
  function renderChips() {
    var el = $("matchBy"), on = by();
    if (!el.children.length) {
      CRITERIA.forEach(function (c) {
        var b = document.createElement("button");
        b.className = "mini-chip";
        b.textContent = c[1];
        b.setAttribute("data-v", c[0]);
        b.setAttribute("aria-pressed", "false");
        b.addEventListener("click", function () {
          var list = by(), i = list.indexOf(c[0]);
          if (i >= 0) list.splice(i, 1); else list.push(c[0]);
          ST.state.matchBy = CRITERIA.map(function (x) { return x[0]; }).filter(function (k) { return list.indexOf(k) >= 0; }).join(",");
          ST.changed();
        });
        el.appendChild(b);
      });
    }
    Array.prototype.forEach.call(el.children, function (b) {
      var isOn = on.indexOf(b.getAttribute("data-v")) >= 0;
      b.classList.toggle("on", isOn);
      b.setAttribute("aria-pressed", String(isOn));
    });
  }
  function renderSupport() {
    var ok = !ST.inOffice || canSelect();
    ["matchSelect", "matchPrev", "matchNext"].forEach(function (id) { $(id).disabled = !ok; });
    if (!ok) setStatus("Selecting shapes from the panel needs a newer version of PowerPoint.");
  }

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    $("matchSelect").addEventListener("click", selectOnSlide);
    $("matchPrev").addEventListener("click", function () { step(-1); });
    $("matchNext").addEventListener("click", function () { step(1); });
    ST.on("change", renderChips);
    ST.on("reset", renderChips);
    ST.on("ready", renderSupport);
    renderChips(); renderSupport();
  }

  ST.registerTool({
    id: "match",
    title: "Select similar",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 1,
    defaults: defaults,
    tip: { what: "Selects shapes like the one you pick, on this slide or slide by slide.", steps: ["Select an example shape.", "Choose what to match.", "Click Select similar."] },
    build: build
  });

  // Exposed for testing only
  ST._match = { signature: signature, matches: matches, placement: function () { return ST._arrange.placement.apply(null, arguments); } };
})();
