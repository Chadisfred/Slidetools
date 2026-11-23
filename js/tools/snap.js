/* Snap: moves shapes onto the template guides, the grid's columns and rows, or steps of g.
   Shown inside the Grid tool (js/tools/layout.js), which calls ST._snap.build. */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;

  var defaults = { arrSnap: "cells", arrSnapEdges: false, arrSnapResize: false };
  ST.addDefaults(defaults);

  var MARKUP =
    '<div class="group-label label-row">Snap' + ST.info("Moves each shape's top left corner to the nearest line.<br><b>Guides</b> the guides on your slide.<br><b>Columns</b> the grid's column and row corners.<br><b>Steps</b> every g from the grid area's top left.", "About snap") + '</div>' +
    '<div class="seg seg3" id="arrSnapMode"></div>' +
    '<div class="toggle arr-cells" data-min="2">Resize to the cells it covers<label class="switch"><input type="checkbox" id="arrSnapResize"><span></span></label></div>' +
    '<div class="toggle arr-edges" data-min="2">Snap all four edges<label class="switch"><input type="checkbox" id="arrSnapEdges"><span></span></label></div>' +
    '<div class="arr-guides" data-min="2"><div class="help" id="arrGuideInfo"></div>' +
      '<div class="btn-row"><button class="btn small" id="arrGuideRead" title="Read the guides from this slide\'s layout and master again">Read again</button>' +
      '<button class="btn small" id="arrGuideUse" title="Set the grid spacing and the grid area\'s left and top from the guides">Use for grid</button></div></div>' +
    '<button class="btn primary" id="arrSnap">Snap</button>';

  // ---------- Snap ----------
  function nearest(list, v) {
    var best = 0;
    for (var i = 1; i < list.length; i++) if (Math.abs(list[i] - v) < Math.abs(list[best] - v)) best = i;
    return best;
  }
  function gridLines() {
    var g = ST.calculate(), xs = [], xe = [], ys = [], ye = [], x = g.left, y = g.top;
    for (var c = 0; c < g.cols; c++) { xs.push(x * PT); xe.push((x + g.widths[c]) * PT); x += g.widths[c] + g.gapH; }
    for (var r = 0; r < g.rows; r++) { ys.push(y * PT); ye.push((y + g.heights[r]) * PT); y += g.heights[r] + g.gapV; }
    return { xs: xs, xe: xe, ys: ys, ye: ye, grid: g };
  }
  // ---------- Template guides ----------
  // Guides live in two places, with positions in eighths of a point:
  //   ppt/viewProps.xml  <p:guide>    presentation guides, draggable on any slide
  //   master and layout  <p15:guide>  template guides, locked outside Slide Master view
  // The whole open file is read (read only) so both kinds are found. If that isn't possible,
  // a copy of the slide gives its layout and master, but not the presentation guides.
  var guides = null; // { slideId, v: [pt], h: [pt], sources: [label], whole: bool }
  var fileBytes = null; // cached copy of the whole file, cleared by Read again
  function canReadWholeFile() {
    try { return !!(Office.context.document && Office.context.document.getFileAsync && Office.FileType && Office.FileType.Compressed); }
    catch (e) { return false; }
  }
  function guidesSupported() {
    return ST.inOffice && typeof DecompressionStream === "function" && (canReadWholeFile() || ST.supports("1.8"));
  }
  function guidesIn(text) {
    var v = [], h = [];
    var doc = new DOMParser().parseFromString(text, "application/xml");
    Array.prototype.forEach.call(doc.getElementsByTagName("*"), function (el) {
      if (el.localName !== "guide") return;
      var pos = parseInt(el.getAttribute("pos") || "0", 10) / 8;
      if (!isFinite(pos)) return;
      (el.getAttribute("orient") === "horz" ? h : v).push(pos);
    });
    return { v: v, h: h };
  }
  // parts: [{ label, text }]. Returns sorted, de-duplicated guides and the labels that had any.
  function collectGuides(parts) {
    var v = {}, h = {}, sources = [];
    parts.forEach(function (p) {
      if (!p.text) return;
      var g = guidesIn(p.text);
      g.v.forEach(function (x) { v[x] = true; });
      g.h.forEach(function (y) { h[y] = true; });
      if (g.v.length || g.h.length) sources.push(p.label);
    });
    function list(o) { return Object.keys(o).map(Number).sort(function (a, b) { return a - b; }); }
    return { v: list(v), h: list(h), sources: sources };
  }

  // Reads the open presentation as a zip, slice by slice. Resolves to bytes, or null.
  function readWholeFile(force) {
    if (!force && fileBytes) return Promise.resolve(fileBytes);
    if (!canReadWholeFile()) return Promise.resolve(null);
    return new Promise(function (resolve) {
      try {
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, function (res) {
          if (res.status !== "succeeded" || !res.value) { resolve(null); return; }
          var file = res.value, count = file.sliceCount, chunks = [], total = 0;
          function done(bytes) { try { file.closeAsync(function () {}); } catch (e) { /* already closed */ } resolve(bytes); }
          function next(i) {
            if (i >= count) {
              var out = new Uint8Array(total), at = 0;
              chunks.forEach(function (c) { out.set(c, at); at += c.length; });
              fileBytes = out;
              done(out);
              return;
            }
            file.getSliceAsync(i, function (s) {
              if (s.status !== "succeeded" || !s.value) { done(null); return; }
              var c = s.value.data instanceof Uint8Array ? s.value.data : new Uint8Array(s.value.data);
              chunks.push(c); total += c.length;
              next(i + 1);
            });
          }
          next(0);
        });
      } catch (e) { resolve(null); }
    });
  }

  // Zip paths: a target in a .rels file is relative to the folder of the part it belongs to.
  function resolvePath(fromPart, target) {
    if (target.charAt(0) === "/") return target.slice(1);
    var parts = fromPart.split("/").slice(0, -1);
    target.split("/").forEach(function (seg) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); });
    return parts.join("/");
  }
  function relsPath(part) { var i = part.lastIndexOf("/"); return part.slice(0, i) + "/_rels/" + part.slice(i + 1) + ".rels"; }
  function relations(text) {
    var doc = new DOMParser().parseFromString(text || "<x/>", "application/xml");
    return Array.prototype.filter.call(doc.getElementsByTagName("*"), function (el) { return el.localName === "Relationship"; })
      .map(function (el) { return { id: el.getAttribute("Id"), type: el.getAttribute("Type") || "", target: el.getAttribute("Target") || "" }; });
  }
  function relTarget(files, part, typeEnd) {
    var r = relations(files[relsPath(part)]).filter(function (x) { return x.type.slice(x.type.lastIndexOf("/") + 1) === typeEnd; })[0];
    return r ? resolvePath(part, r.target) : null;
  }

  // Finds the guides for the slide at index from the whole file's bytes.
  function guidesFromFile(bytes, index) {
    var zip = ST.zip, files = {};
    var first = /^(ppt\/viewProps\.xml|ppt\/presentation\.xml|ppt\/_rels\/presentation\.xml\.rels|ppt\/slides\/_rels\/[^\/]+\.rels|ppt\/slideLayouts\/_rels\/[^\/]+\.rels)$/;
    return zip.readZip(bytes, function (n) { return first.test(n); }).then(function (list) {
      list.forEach(function (f) { files[f.name] = f.text; });
      var pres = new DOMParser().parseFromString(files["ppt/presentation.xml"] || "<x/>", "application/xml");
      var ids = Array.prototype.filter.call(pres.getElementsByTagName("*"), function (el) { return el.localName === "sldId"; });
      var sld = ids[index], rid = null;
      if (sld) Array.prototype.forEach.call(sld.attributes, function (a) { if (a.localName === "id" && a.namespaceURI && /relationships$/.test(a.namespaceURI)) rid = a.value; });
      var rel = rid ? relations(files["ppt/_rels/presentation.xml.rels"]).filter(function (r) { return r.id === rid; })[0] : null;
      var slidePart = rel ? resolvePath("ppt/presentation.xml", rel.target) : null;
      var layoutPart = slidePart ? relTarget(files, slidePart, "slideLayout") : null;
      var masterPart = layoutPart ? relTarget(files, layoutPart, "slideMaster") : null;
      var wanted = [layoutPart, masterPart].filter(Boolean);
      return (wanted.length ? zip.readZip(bytes, function (n) { return wanted.indexOf(n) >= 0; }) : Promise.resolve([])).then(function (more) {
        more.forEach(function (f) { files[f.name] = f.text; });
        var g = collectGuides([
          { label: "presentation guides", text: files["ppt/viewProps.xml"] },
          { label: "slide layout", text: layoutPart && files[layoutPart] },
          { label: "slide master", text: masterPart && files[masterPart] }
        ]);
        g.whole = true;
        return g;
      });
    });
  }
  function guidesFromCopy(context, slide) {
    if (!ST.supports("1.8")) return Promise.resolve(null);
    var data = slide.exportAsBase64();
    return context.sync().then(function () {
      return ST.zip.readZip(ST.zip.b64ToBytes(data.value), function (n) { return /^ppt\/(slideMasters\/slideMaster\d+|slideLayouts\/slideLayout\d+)\.xml$/.test(n); });
    }).then(function (files) {
      var g = collectGuides(files.map(function (f) { return { label: /Master/.test(f.name) ? "slide master" : "slide layout", text: f.text }; }));
      g.whole = false;
      return g;
    });
  }
  function loadGuides(context, force) {
    if (!guidesSupported() || !ST.zip) return Promise.resolve(null);
    var sel = context.presentation.getSelectedSlides(), all = context.presentation.slides;
    sel.load("items/id"); all.load("items/id");
    return context.sync().then(function () {
      var slide = sel.items[0] || all.items[0];
      if (!slide) return null;
      if (!force && guides && guides.slideId === slide.id) return guides;
      var index = Math.max(0, all.items.map(function (s) { return s.id; }).indexOf(slide.id));
      return readWholeFile(force).then(function (bytes) {
        return bytes ? guidesFromFile(bytes, index).catch(function () { return null; }) : null;
      }).then(function (g) {
        return g || guidesFromCopy(context, slide);
      }).then(function (g) {
        if (!g) return null;
        g.slideId = slide.id;
        guides = g;
        renderGuideInfo();
        return g;
      });
    });
  }
  // Describes the spacing between guides and points out any that are slightly off the pattern.
  function analyse(list) {
    if (list.length < 2) return { step: null, off: [] };
    var counts = {}, gaps = [];
    for (var i = 1; i < list.length; i++) { var d = Math.round((list[i] - list[i - 1]) * 8) / 8; gaps.push(d); counts[d] = (counts[d] || 0) + 1; }
    var step = Number(Object.keys(counts).sort(function (a, b) { return counts[b] - counts[a]; })[0]);
    // Guides on the pattern share the same remainder after dividing by the step. Ones a little
    // away from it are reported with how far off they are. Far away ones are taken as deliberate.
    var rem = {}, best = null;
    list.forEach(function (p) { var r = Math.round((((p % step) + step) % step) * 8) / 8; rem[r] = (rem[r] || 0) + 1; if (best === null || rem[r] > rem[best]) best = r; });
    var off = [];
    list.forEach(function (p) {
      var d = (((p - best) % step) + step) % step;
      if (d > step / 2) d -= step;
      d = Math.round(d * 8) / 8;
      if (d !== 0 && Math.abs(d) <= Math.max(1, step * 0.1)) off.push({ pos: p, by: d });
    });
    return { step: step, off: off, regular: counts[step] === gaps.length };
  }
  function inch(pt) { return ST.exact(pt / PT); }
  function joinWords(list) { return list.length < 2 ? list.join("") : list.slice(0, -1).join(", ") + " and " + list[list.length - 1]; }
  function describe(list, word) {
    if (!list.length) return "No " + word + " guides.";
    var a = analyse(list), text = ST.plural(list.length, word + " guide");
    if (a.step) text += a.regular ? ", evenly every " + a.step + " pt (" + inch(a.step) + " in)" : ", mostly " + a.step + " pt (" + inch(a.step) + " in) apart";
    text += ". ";
    var after = word === "vertical" ? ["right", "left"] : ["down", "up"];
    text += a.off.length ? ST.plural(a.off.length, "guide") + " slightly off: " + a.off.map(function (o) {
      return inch(o.pos) + " in (" + Math.abs(o.by) + " pt " + (o.by > 0 ? after[0] : after[1]) + ")";
    }).join(", ") + "." : (a.step ? "None off the pattern." : "");
    return text;
  }
  function renderGuideInfo() {
    var el = $("arrGuideInfo");
    if (!el) return;
    if (!ST.inOffice) { el.textContent = "Guides are read from your template once the panel is open in PowerPoint."; return; }
    if (!guidesSupported()) { el.textContent = "This PowerPoint version can't read template guides."; return; }
    if (!guides) { el.textContent = "Guides haven't been read yet. They're read the first time you snap."; return; }
    if (!guides.v.length && !guides.h.length) {
      el.textContent = guides.whole
        ? "No guides found in the presentation guides, this slide's layout or its master."
        : "No guides found in this slide's layout or master. The whole file couldn't be read, so guides drawn on the slides weren't checked.";
      return;
    }
    el.textContent = describe(guides.v, "vertical") + " " + describe(guides.h, "horizontal") + " Found in the " + joinWords(guides.sources) + ".";
  }
  function readGuidesNow(force) {
    if (!ST.needOffice()) return;
    if (!guidesSupported()) { ST.toast("This PowerPoint version can't read template guides.", true); return; }
    ST.run(function (context) {
      return loadGuides(context, force).then(function (g) {
        if (g) ST.toast(g.v.length + g.h.length ? "Read " + ST.plural(g.v.length + g.h.length, "guide") + "." : "No guides found.", !(g.v.length + g.h.length));
      });
    });
  }
  function useInLayout() {
    if (!guides || (!guides.v.length && !guides.h.length)) { ST.toast("Read the guides first.", true); return; }
    var a = analyse(guides.v.length >= guides.h.length ? guides.v : guides.h), s = ST.state, parts = [];
    if (a.step) { s.spacing = ST.exact(a.step / PT); parts.push("spacing " + s.spacing + " in"); }
    if (guides.v.length) { s.left = ST.exact(guides.v[0] / PT); parts.push("left " + s.left + " in"); }
    if (guides.h.length) { s.top = ST.exact(guides.h[0] / PT); parts.push("top " + s.top + " in"); }
    if (s.preset !== undefined) s.preset = "custom";
    ST.emit("reset");
    ST.changed();
    ST.toast("Grid set from guides: " + parts.join(", ") + ".");
  }
  function snapToGuides(edges) {
    ST.run(function (context) {
      var shapes = context.presentation.getSelectedShapes();
      shapes.load("items/left,items/top,items/width,items/height");
      return loadGuides(context, false).then(function (g) {
        return context.sync().then(function () {
          if (!g) { ST.toast("This PowerPoint version can't read template guides.", true); return; }
          if (!g.v.length && !g.h.length) { ST.toast("No guides found. See the guide report at the More level.", true); return; }
          var items = shapes.items;
          if (!items.length) { ST.toast("Select the shapes on your slide first.", true); return; }
          items.forEach(function (s) {
            var r = s.left + s.width, b = s.top + s.height;
            if (g.v.length) {
              s.left = g.v[nearest(g.v, s.left)];
              if (edges) { var rs = g.v.filter(function (x) { return x > s.left; }); if (rs.length) s.width = rs[nearest(rs, r)] - s.left; }
            }
            if (g.h.length) {
              s.top = g.h[nearest(g.h, s.top)];
              if (edges) { var bs = g.h.filter(function (y) { return y > s.top; }); if (bs.length) s.height = bs[nearest(bs, b)] - s.top; }
            }
          });
          return context.sync().then(function () { ST.toast("Snapped " + ST.plural(items.length, "shape") + " to the template guides."); ST.refreshSelection(); });
        });
      });
    });
  }

  function snap() {
    if (ST.state.arrSnap === "guides") {
      if (!ST.needOffice()) return;
      if (!guidesSupported()) { ST.toast("This PowerPoint version can't read template guides.", true); return; }
      snapToGuides(!!ST.state.arrSnapEdges);
      return;
    }
    var g = ST.calculate();
    if (!g.valid) { ST.toast("The grid isn't valid. Check its spacing and area.", true); return; }
    ST.withSelection(["left", "top", "width", "height"], function (context, items) {
      if (ST.state.arrSnap === "fine") {
        var step = g.spacing * PT, ox = g.left * PT, oy = g.top * PT;
        if (!(step > 0)) { ST.toast("Grid spacing needs to be more than 0.", true); return; }
        var round = function (v, o) { return o + Math.round((v - o) / step) * step; };
        items.forEach(function (s) {
          var r = s.left + s.width, b = s.top + s.height;
          s.left = round(s.left, ox); s.top = round(s.top, oy);
          if (ST.state.arrSnapEdges) {
            s.width = Math.max(step, round(r, ox) - s.left);
            s.height = Math.max(step, round(b, oy) - s.top);
          }
        });
        return context.sync().then(function () { ST.toast("Snapped " + ST.plural(items.length, "shape") + " to steps of " + ST.nice(g.spacing) + " in."); ST.refreshSelection(); });
      }
      var L = gridLines();
      items.forEach(function (s) {
        var c = nearest(L.xs, s.left), r = nearest(L.ys, s.top);
        var right = s.left + s.width, bottom = s.top + s.height;
        s.left = L.xs[c]; s.top = L.ys[r];
        if (ST.state.arrSnapResize) {
          var ce = c + nearest(L.xe.slice(c), right), re = r + nearest(L.ye.slice(r), bottom);
          s.width = L.xe[ce] - L.xs[c]; s.height = L.ye[re] - L.ys[r];
        }
      });
      return context.sync().then(function () { ST.toast("Snapped " + ST.plural(items.length, "shape") + " to the grid cells."); ST.refreshSelection(); });
    });
  }

  function renderSnapNote() {
    var el = $("arrSnapNote"), mode = ST.state.arrSnap, g = ST.calculate();
    if (!el) return;
    if (mode === "guides") el.textContent = "Moves each shape's top left corner to the nearest guide in your template.";
    else if (mode === "cells") el.textContent = "Moves each shape's top left corner to the nearest column and row corner of the grid (" + g.cols + " x " + g.rows + ").";
    else el.textContent = "Moves each shape's top left corner to the nearest step of g (" + ST.exact(g.spacing) + " in), counted from the grid area's top left.";
  }

  function renderSegs() {
    Array.prototype.forEach.call($("arrSnapMode").children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === ST.state.arrSnap); });
    var mode = ST.state.arrSnap;
    function show(sel, on) { Array.prototype.forEach.call(document.querySelectorAll(sel), function (el) { el.classList.toggle("off", !on); }); }
    show(".arr-cells", mode === "cells");
    show(".arr-edges", mode !== "cells");
    show(".arr-guides", mode === "guides");
    renderSnapNote();
  }

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    var el = $("arrSnapMode");
    [["guides", "Guides"], ["cells", "Columns"], ["fine", "Steps"]].forEach(function (o) {
      var b = document.createElement("button");
      b.textContent = o[1];
      b.setAttribute("data-v", o[0]);
      b.addEventListener("click", function () {
        ST.state.arrSnap = o[0]; ST.save(); renderSegs();
        if (o[0] === "guides" && !guides && ST.inOffice && guidesSupported()) readGuidesNow(false);
      });
      el.appendChild(b);
    });
    $("arrGuideRead").addEventListener("click", function () { readGuidesNow(true); });
    $("arrGuideUse").addEventListener("click", useInLayout);
    ["arrSnapResize", "arrSnapEdges"].forEach(function (k) {
      var box = $(k);
      box.checked = !!ST.state[k];
      box.addEventListener("change", function () { ST.state[k] = box.checked; ST.save(); });
      ST.on("reset", function () { box.checked = !!ST.state[k]; });
    });
    $("arrSnap").addEventListener("click", function () { ST.setRepeat("Snap", function () { snap(); }); snap(); });
    renderSegs();
    ST.on("change", renderSnapNote);
    ST.on("reset", renderSegs);
    ST.on("ready", renderGuideInfo);
    renderGuideInfo();
  }

  ST._snap = { defaults: defaults, build: build };
})();
