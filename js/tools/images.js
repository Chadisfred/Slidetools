/* Images tool: paste a picture into a shape as a real, croppable picture, move the crop of
   pictures, reset crops, and choose whether resizing crops pictures. Pictures are placed
   through the shared slide rebuild (js/rebuild.js): the image is added to an in-memory
   copy of the slide, nothing is sent anywhere, and only the slide itself is edited. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;
  var R = ST.rebuild, NS_P = R.NS_P, NS_A = R.NS_A, kid = R.kid, kids = R.kids, make = R.make;
  var NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round"><rect x="2.5" y="3.5" width="13" height="11" rx="2"/><circle cx="6.5" cy="7.5" r="1.3"/><path d="M3 13l3.8-3.6 2.7 2.4 2.2-2 3.3 3.2"/></svg>';
  var NAMES = ["top left", "top middle", "top right", "middle left", "centre", "middle right", "bottom left", "bottom middle", "bottom right"];
  var SHAPES = ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform", "Image"];
  var KEEP = { png: 1, jpeg: 1, gif: 1, bmp: 1 };
  var FILLS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];

  var defaults = { imgPos: 4 };
  ST.addDefaults(defaults);

  function factors(i) { return { x: (i % 3) / 2, y: Math.floor(i / 3) / 2 }; }

  // ---------- Maths (pure) ----------
  // Crop (fractions) that fills a w x h frame with an iw x ih picture, placed by f
  function coverCrop(iw, ih, w, h, f) {
    var ia = iw / ih, fa = w / h, c = { l: 0, t: 0, r: 0, b: 0 };
    if (ia > fa) { var x = 1 - fa / ia; c.l = x * f.x; c.r = x - c.l; }
    else { var y = 1 - ia / fa; c.t = y * f.y; c.b = y - c.t; }
    return c;
  }
  // Moves an existing crop to position f, keeping how much is cropped each way
  function movedCrop(c, f) {
    var x = c.l + c.r, y = c.t + c.b;
    return { l: x * f.x, r: x * (1 - f.x), t: y * f.y, b: y * (1 - f.y) };
  }
  // The frame that shows the whole picture where a cropped one now sits
  function uncroppedBox(box, c) {
    var vw = 1 - c.l - c.r, vh = 1 - c.t - c.b;
    if (!(vw > 0) || !(vh > 0)) return box;
    var W = box.width / vw, H = box.height / vh;
    return { left: box.left - c.l * W, top: box.top - c.t * H, width: W, height: H };
  }

  // ---------- Slide XML ----------
  function readCrop(pic) {
    var bf = kid(pic, NS_P, "blipFill"), r = bf && kid(bf, NS_A, "srcRect"), v = function (k) { return r && r.getAttribute(k) ? +r.getAttribute(k) / 100000 : 0; };
    return { l: v("l"), t: v("t"), r: v("r"), b: v("b") };
  }
  function writeCrop(doc, pic, c) {
    var bf = kid(pic, NS_P, "blipFill"), r = kid(bf, NS_A, "srcRect");
    if (!r) { r = make(doc, NS_A, "srcRect"); var blip = kid(bf, NS_A, "blip"); bf.insertBefore(r, blip ? blip.nextSibling : bf.firstChild); }
    ["l", "t", "r", "b"].forEach(function (k) { var n = Math.round((c[k] || 0) * 100000); if (n) r.setAttribute(k, n); else r.removeAttribute(k); });
    // Stretch to the frame, as PowerPoint does for cropped pictures
    kids(bf, NS_A, "tile").forEach(function (t) { bf.removeChild(t); });
    if (!kid(bf, NS_A, "stretch")) { var st = make(doc, NS_A, "stretch"); st.appendChild(make(doc, NS_A, "fillRect")); bf.appendChild(st); }
  }
  function maxId(doc) {
    var m = 0;
    Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS_P, "cNvPr"), function (c) { m = Math.max(m, +c.getAttribute("id") || 0); });
    return m;
  }
  // A picture element that takes the shape's place, frame, outline, corners and effects
  function pictureFrom(doc, sp, rId, crop, apiBox, id) {
    var pic = make(doc, NS_P, "pic");
    var nv = make(doc, NS_P, "nvPicPr"), oldC = R.cNvPr(sp), c = make(doc, NS_P, "cNvPr");
    c.setAttribute("id", id);
    c.setAttribute("name", oldC && oldC.getAttribute("name") || "Picture");
    if (oldC && oldC.getAttribute("descr")) c.setAttribute("descr", oldC.getAttribute("descr"));
    var cp = make(doc, NS_P, "cNvPicPr"), locks = make(doc, NS_A, "picLocks");
    locks.setAttribute("noChangeAspect", "1"); cp.appendChild(locks);
    var nvPr = make(doc, NS_P, "nvPr"), oldNv = kid(kid(sp, NS_P, "nvSpPr") || sp, NS_P, "nvPr"), ph = oldNv && kid(oldNv, NS_P, "ph");
    // A picture placeholder stays a placeholder, now holding the picture
    var keepPh = !!ph && ph.getAttribute("type") === "pic";
    if (keepPh) nvPr.appendChild(ph.cloneNode(true));
    nv.appendChild(c); nv.appendChild(cp); nv.appendChild(nvPr);
    var bf = make(doc, NS_P, "blipFill"), blip = make(doc, NS_A, "blip");
    blip.setAttributeNS(NS_R, "r:embed", rId); bf.appendChild(blip);
    var spPr = (kid(sp, NS_P, "spPr") || make(doc, NS_P, "spPr")).cloneNode(true);
    kids(spPr, NS_A).forEach(function (n) { if (FILLS.indexOf(n.localName) >= 0) spPr.removeChild(n); });
    // A placeholder without its own frame or outline shape keeps inheriting them from the layout
    if (!keepPh && !kid(spPr, NS_A, "xfrm")) { var x = make(doc, NS_A, "xfrm"); spPr.insertBefore(x, spPr.firstChild); }
    if (!keepPh && !kid(spPr, NS_A, "prstGeom") && !kid(spPr, NS_A, "custGeom")) {
      var g = make(doc, NS_A, "prstGeom"); g.setAttribute("prst", "rect"); g.appendChild(make(doc, NS_A, "avLst"));
      var xf = kid(spPr, NS_A, "xfrm"); spPr.insertBefore(g, xf.nextSibling);
    }
    pic.appendChild(nv); pic.appendChild(bf); pic.appendChild(spPr);
    // Keep the theme outline the shape had, but not its theme fill behind the picture
    var style = kid(sp, NS_P, "style");
    if (style) { style = style.cloneNode(true); var fr = kid(style, NS_A, "fillRef"); if (fr) fr.setAttribute("idx", "0"); pic.appendChild(style); }
    if (!keepPh && !R.readBox(pic)) R.writeBox(pic, apiBox);
    writeCrop(doc, pic, crop);
    return pic;
  }

  // ---------- Reading pictures ----------
  function imageOf(file) {
    var type = String(file.type || "").toLowerCase(), ext = type.split("/")[1] === "jpg" ? "jpeg" : type.split("/")[1];
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(file), img = new Image();
      img.onload = function () { resolve(img); };
      img.onerror = function () { URL.revokeObjectURL(url); reject(new Error("unreadable")); };
      img.src = url;
    }).then(function (img) {
      var size = { w: img.naturalWidth, h: img.naturalHeight };
      if (KEEP[ext]) return file.arrayBuffer().then(function (buf) { return { bytes: new Uint8Array(buf), ext: ext, w: size.w, h: size.h }; });
      // PowerPoint may not read other formats, such as WebP, so they become PNG
      var c = document.createElement("canvas"); c.width = size.w; c.height = size.h; c.getContext("2d").drawImage(img, 0, 0);
      return new Promise(function (res) { c.toBlob(res, "image/png"); }).then(function (b) { return b.arrayBuffer(); })
        .then(function (buf) { return { bytes: new Uint8Array(buf), ext: "png", w: size.w, h: size.h }; });
    });
  }

  function ready() {
    if (!ST.needOffice()) return false;
    if (!R.supported()) { ST.toast("This version of PowerPoint can't place or crop pictures. It needs PowerPoint API 1.8.", true); return false; }
    return true;
  }
  function failText(e) { return "Couldn't change the slide. " + (ST.errorText ? ST.errorText(e) : ""); }
  function selectNames(context, slideId, names) {
    if (!slideId || !names.length || !ST.supports("1.5")) return Promise.resolve();
    var slide = context.presentation.slides.getItem(slideId), shapes = slide.shapes;
    shapes.load("items/id,items/name");
    return context.sync().then(function () {
      var ids = shapes.items.filter(function (x) { return names.indexOf(x.name) >= 0; }).map(function (x) { return x.id; });
      if (ids.length) { slide.setSelectedShapes(ids); return context.sync(); }
    }).catch(function () {});
  }

  // ---------- Paste into shape ----------
  function place(file) {
    if (!file || !/^image\//.test(file.type || "")) { ST.toast("That isn't a picture. Paste or drop a PNG, JPEG, GIF or similar.", true); return; }
    if (!ready()) return;
    var f = factors(ST.state.imgPos), image;
    imageOf(file).then(function (im) {
      image = im;
      ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
        var list = items.filter(function (s) { return SHAPES.indexOf(s.type) >= 0 && s.width > 0 && s.height > 0; });
        if (!list.length) { ST.toast("Select a shape or picture to place the picture in. Groups, tables and lines can't hold one.", true); return; }
        return ST.confirmSwap("place the picture").then(function (ok) {
          if (!ok) return { cancelled: true };
          var out = { placed: 0, text: 0, names: [] };
          return R.edit(context, function (doc, extra) {
            var boxes = list.map(function (s) { return { left: s.left, top: s.top, width: s.width, height: s.height, name: s.name, id: s.id }; });
            var els = R.matchShapes(doc, boxes, ["sp", "pic"]), rId = null, next = maxId(doc) + 1;
            els.forEach(function (el, i) {
              if (!el) return;
              rId = rId || extra.addImage(image.bytes, image.ext);
              var box = R.readBox(el) || boxes[i], crop = coverCrop(image.w, image.h, box.width, box.height, f);
              if (el.localName === "pic") {
                var blip = kid(kid(el, NS_P, "blipFill"), NS_A, "blip");
                blip.setAttributeNS(NS_R, "r:embed", rId);
                writeCrop(doc, el, crop);
              } else {
                var tb = kid(el, NS_P, "txBody");
                if (tb && /\S/.test(tb.textContent || "")) out.text++;
                var pic = pictureFrom(doc, el, rId, crop, boxes[i], next++);
                el.parentNode.replaceChild(pic, el);
              }
              out.names.push(boxes[i].name); out.placed++;
            });
            return { count: out.placed };
          }).then(function (r) {
            return selectNames(context, r.slideId, out.names).then(function () { return out; });
          });
        }).then(function (r) {
          if (r.cancelled) { ST.toast("The picture wasn't placed."); return; }
          if (!r.placed) { ST.toast(R.missText(), true); return; }
          ST.toast("Placed the picture in " + ST.plural(r.placed, "shape") + ". Use PowerPoint's Crop button to pan or zoom it." + (r.text ? " Text in the shape was removed, since pictures can't hold text." : ""));
          ST.refreshSelection();
        }, function (e) { ST.toast(failText(e), true); });
      });
    }, function () { ST.toast("Couldn't read that picture. Try saving it as PNG or JPEG.", true); });
  }

  // ---------- Crop position and reset ----------
  // fn(doc, pic, box) changes one picture. Resolves the number changed.
  // quiet: say nothing when no pictures are selected (a crop position click with a shape selected)
  function editPictures(what, fn, done, quiet) {
    if (quiet ? !(ST.inOffice && R.supported()) : !ready()) return;
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      return ST.crop.pictureFlags(context, items).then(function (isPic) {
        var pics = items.filter(function (s, i) { return isPic[i]; });
        if (!pics.length) { if (!quiet) ST.toast("Select one or more pictures first.", true); return; }
        return ST.confirmSwap(what).then(function (ok) {
          if (!ok) return null;
          var names = [];
          return R.edit(context, function (doc) {
            var boxes = pics.map(function (s) { return { left: s.left, top: s.top, width: s.width, height: s.height, name: s.name, id: s.id }; });
            var els = R.matchShapes(doc, boxes, ["pic"]), n = 0;
            els.forEach(function (el, i) { if (el && fn(doc, el, R.readBox(el) || boxes[i])) { n++; names.push(boxes[i].name); } });
            return { count: n };
          }).then(function (r) { return selectNames(context, r.slideId, names).then(function () { return r.count; }); });
        }).then(function (n) {
          if (n === null) { ST.toast("Nothing was changed."); return; }
          ST.toast(n ? done(n) : "Those pictures already look like that.");
          ST.refreshSelection();
        }, function (e) { ST.toast(failText(e), true); });
      });
    });
  }
  function moveCrop(i, quiet) {
    var f = factors(i);
    editPictures("move the crop", function (doc, pic) {
      var c = readCrop(pic), m = movedCrop(c, f);
      if (["l", "t", "r", "b"].every(function (k) { return Math.abs(m[k] - c[k]) < 1e-5; })) return false;
      writeCrop(doc, pic, m); return true;
    }, function (n) {
      ST.setRepeat("Move crop to " + NAMES[i], function () { moveCrop(i); });
      return "Moved the crop of " + ST.plural(n, "picture") + " to the " + NAMES[i] + ".";
    }, quiet);
  }
  function resetCrop() {
    editPictures("reset the crop", function (doc, pic, box) {
      var c = readCrop(pic);
      if (!c.l && !c.t && !c.r && !c.b) return false;
      R.writeBox(pic, uncroppedBox(box, c));
      writeCrop(doc, pic, { l: 0, t: 0, r: 0, b: 0 }); return true;
    }, function (n) { return "Reset the crop of " + ST.plural(n, "picture") + ". The whole picture shows again."; });
  }

  // ---------- Panel ----------
  var MARKUP =
    '<div class="group-label label-row">Paste into shape' + ST.info("Select a shape, such as a rounded card, and paste a picture into it. The shape becomes a real picture with the same position, size, corners and outline, cropped to fill it without stretching. The whole picture is kept, so PowerPoint\'s Crop button can pan or zoom it afterwards. Pasting onto a picture swaps its image. The picture stays on this computer. PowerPoint add-ins can\'t place cropped pictures directly, so the slide is swapped for an identical copy. Slide comments may not carry over, and Undo takes more than one step.", "About paste into shape") + '</div>' +
    '<div class="pic-drop" id="imgDrop" tabindex="0" role="button" aria-label="Paste or drop a picture into the selected shapes"><b>Click here and press Ctrl+V</b><span>or drop a picture, or <a href="#" id="imgChoose">choose a file</a></span></div>' +
    '<input type="file" id="imgFile" accept="image/*" hidden>' +
    '<div class="anchor-row"><span class="anchor-text"><span class="label-row">Crop position' + ST.info("Where new pictures are cropped from. With pictures selected, clicking a point also moves their crop there, keeping how much is cropped.", "About crop position") + '</span><b id="imgPosName"></b></span><div class="anchor" id="imgPos"></div></div>' +
    '<button class="btn" id="imgReset" title="Remove the crop so the whole picture shows">Reset crop</button>' +
    '<div class="toggle" data-min="2"><span class="label-row">Crop pictures to fit' + ST.info("When Size changes a picture\'s frame shape, the picture is cropped to fill it instead of being stretched. The slide is swapped for an identical copy with the new crop. Turn this off to keep pictures in proportion instead.", "About cropping pictures") + '</span><label class="switch"><input type="checkbox" id="cropPictures"><span></span></label></div>';

  function renderPos() {
    var el = $("imgPos"), i0 = ST.state.imgPos === undefined ? 4 : ST.state.imgPos;
    if (!el.children.length) NAMES.forEach(function (name, i) {
      var b = document.createElement("button");
      b.title = "Crop from the " + name; b.setAttribute("aria-label", b.title);
      b.addEventListener("click", function () {
        ST.state.imgPos = i; ST.save(); renderPos();
        if (ST.selection && ST.selection.count && ST.inOffice) moveCrop(i, true);
      });
      el.appendChild(b);
    });
    Array.prototype.forEach.call(el.children, function (b, i) { b.classList.toggle("on", i === i0); });
    $("imgPosName").textContent = NAMES[i0];
  }

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    var drop = $("imgDrop"), input = $("imgFile");
    drop.addEventListener("paste", function (e) {
      var items = (e.clipboardData && e.clipboardData.items) || [], file = null;
      for (var i = 0; i < items.length && !file; i++) if (items[i].kind === "file" && /^image\//.test(items[i].type)) file = items[i].getAsFile();
      e.preventDefault();
      if (file) place(file); else ST.toast("The clipboard has no picture. Copy a picture, then press Ctrl+V here.", true);
    });
    drop.addEventListener("click", function (e) { if (e.target.id !== "imgChoose") drop.focus(); });
    drop.addEventListener("dragover", function (e) { e.preventDefault(); drop.classList.add("over"); });
    drop.addEventListener("dragleave", function () { drop.classList.remove("over"); });
    drop.addEventListener("drop", function (e) { e.preventDefault(); drop.classList.remove("over"); var fl = e.dataTransfer && e.dataTransfer.files; place(fl && fl[0]); });
    $("imgChoose").addEventListener("click", function (e) { e.preventDefault(); input.click(); });
    input.addEventListener("change", function () { if (input.files[0]) place(input.files[0]); input.value = ""; });
    $("imgReset").addEventListener("click", resetCrop);
    var crop = $("cropPictures");
    crop.checked = ST.state.cropPictures !== false;
    crop.addEventListener("change", function () { ST.state.cropPictures = crop.checked; ST.save(); });
    renderPos();
    ST.on("reset", function () { crop.checked = ST.state.cropPictures !== false; renderPos(); });
  }

  ST.registerTool({
    id: "images",
    title: "Images",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 1,
    defaults: defaults,
    tip: { what: "Puts pictures into shapes and adjusts their crop.", steps: ["Select a shape.", "Click the paste box and press Ctrl+V.", "Click a crop position to move what shows."] },
    build: build
  });

  ST._images = { coverCrop: coverCrop, movedCrop: movedCrop, uncroppedBox: uncroppedBox, place: place, moveCrop: moveCrop, resetCrop: resetCrop, _pictureFrom: pictureFrom };
})();
