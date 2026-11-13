/* Match tool: select shapes like the selected one, slide by slide, and copy one shape's position and size onto others. */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;

  var defaults = { matchBy: "fill,type", pasteWhat: "both", pasteAnchor: 0 };
  ST.addDefaults(defaults);

  var CRITERIA = [["fill", "Fill colour"], ["line", "Outline"], ["shapeStyle", "Shape style"], ["textStyle", "Text style"], ["size", "Size"], ["type", "Type"]];
  var ANCHOR_NAMES = ["top left", "top", "top right", "left", "centre", "right", "bottom left", "bottom", "bottom right"];
  var SIZE_TOLERANCE = 0.5; // points

  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="2" y="2" width="8" height="8" rx="1.5"/><rect x="8" y="8" width="8" height="8" rx="1.5" fill="currentColor" fill-opacity="0.25"/></svg>';

  var MARKUP =
    '<span class="info res-info" tabindex="0" role="button" aria-label="How Match works">i' +
      '<span class="tip" role="tooltip">Select similar uses the first selected shape as the example and selects every shape on the slide that matches it on the ticked settings.<br><br>' +
      'PowerPoint can only select shapes on one slide at a time, so Next slide and Previous slide jump to the next slide with matches and select them there. Masters and layouts are never searched.<br><br>' +
      'Copy remembers the first selected shape\'s position and size until you close the panel. Paste puts every selected shape in the same spot, the same size, or both. Paste from chooses which point lines up when the sizes differ.</span>' +
    '</span>' +
    '<div class="group-label">Select similar</div>' +
    '<div class="sublabel">Match by</div>' +
    '<div class="mini-chips" id="matchBy"></div>' +
    '<button class="btn primary" id="matchSelect">Select similar on slide</button>' +
    '<div class="btn-row" data-min="2"><button class="btn small" id="matchPrev" title="Go to the previous slide with matches and select them">Previous slide</button><button class="btn small" id="matchNext" title="Go to the next slide with matches and select them">Next slide</button></div>' +
    '<div class="help" id="matchStatus"></div>' +
    '<div class="group-label">Position and size</div>' +
    '<div class="btn-row"><button class="btn small" id="posCopy" title="Remember the first selected shape\'s position and size">Copy</button><button class="btn small" id="posPaste" title="Apply the copied position and size to the selected shapes" disabled>Paste</button></div>' +
    '<div class="seg seg3" id="pasteWhat"></div>' +
    '<div class="anchor-row" data-min="2"><span>Paste from <b id="pasteAnchorName"></b></span><div class="anchor" id="pasteAnchor"></div></div>' +
    '<div class="help" id="posCopied">Nothing copied yet.</div>';

  // ---------- State kept for this session only ----------
  var reference = null; // signature of the example shape
  var copied = null;    // { left, top, width, height } in points

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

  // ---------- Copy and paste position and size ----------
  function inches(pt) { return ST.nice(Math.round(pt / PT * 100) / 100); }
  function showCopied() {
    $("posPaste").disabled = !copied;
    $("posCopied").textContent = copied
      ? "Copied: left " + inches(copied.left) + ", top " + inches(copied.top) + " in. " + inches(copied.width) + " x " + inches(copied.height) + " in."
      : "Nothing copied yet.";
  }
  function copy() {
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      var s = items[0];
      copied = { left: s.left, top: s.top, width: s.width, height: s.height };
      showCopied();
      ST.toast(items.length > 1 ? "Copied the first selected shape." : "Copied position and size.");
    }, "Select the shape to copy first.");
  }
  // Works out where a shape goes. The chosen anchor point of the shape lands on the same
  // point of the copied box (Position), or stays where it is (Size only).
  function placement(s, box, what, anchor) {
    var fx = (anchor % 3) / 2, fy = Math.floor(anchor / 3) / 2;
    var size = what !== "position", pos = what !== "size";
    var w = size ? box.width : s.width, h = size ? box.height : s.height;
    var ax = pos ? box.left + box.width * fx : s.left + s.width * fx;
    var ay = pos ? box.top + box.height * fy : s.top + s.height * fy;
    return { left: ax - w * fx, top: ay - h * fy, width: w, height: h };
  }
  function paste() {
    if (!copied) { ST.toast("Copy a shape first.", true); return; }
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      var what = ST.state.pasteWhat || "both", anchor = ST.state.pasteAnchor || 0;
      items.forEach(function (s) {
        var p = placement(s, copied, what, anchor);
        if (what !== "position") { s.width = p.width; s.height = p.height; }
        s.left = p.left; s.top = p.top;
      });
      return context.sync().then(function () {
        var word = what === "both" ? "position and size" : what;
        ST.toast("Pasted " + word + " onto " + ST.plural(items.length, "shape") + ".");
        ST.refreshSelection();
      });
    }, "Select the shapes to paste onto.");
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
  function renderPaste() {
    var seg = $("pasteWhat");
    if (!seg.children.length) {
      [["position", "Position"], ["size", "Size"], ["both", "Both"]].forEach(function (o) {
        var b = document.createElement("button");
        b.textContent = o[1];
        b.setAttribute("data-v", o[0]);
        b.addEventListener("click", function () { ST.state.pasteWhat = o[0]; ST.changed(); });
        seg.appendChild(b);
      });
    }
    Array.prototype.forEach.call(seg.children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === (ST.state.pasteWhat || "both")); });
    var el = $("pasteAnchor");
    if (!el.children.length) {
      ANCHOR_NAMES.forEach(function (name, i) {
        var b = document.createElement("button");
        b.title = "Line up the " + name + " corner";
        if (i === 1 || i === 3 || i === 5 || i === 7) b.title = "Line up the " + name + " edge";
        if (i === 4) b.title = "Line up the centre";
        b.setAttribute("aria-label", b.title);
        b.addEventListener("click", function () { ST.state.pasteAnchor = i; ST.changed(); });
        el.appendChild(b);
      });
    }
    var a = ST.state.pasteAnchor || 0;
    Array.prototype.forEach.call(el.children, function (b, i) { b.classList.toggle("on", i === a); });
    $("pasteAnchorName").textContent = ANCHOR_NAMES[a];
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
    $("posCopy").addEventListener("click", copy);
    $("posPaste").addEventListener("click", paste);
    ST.on("change", function () { renderChips(); renderPaste(); });
    ST.on("reset", function () { renderChips(); renderPaste(); });
    ST.on("ready", renderSupport);
    renderChips(); renderPaste(); showCopied(); renderSupport();
  }

  ST.registerTool({
    id: "match",
    title: "Match",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 1,
    defaults: defaults,
    build: build
  });

  // Exposed for testing only
  ST._match = { signature: signature, matches: matches, placement: placement };
})();
