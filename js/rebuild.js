/* Slide rebuild, shared by Transform's picture cropping and copies. PowerPoint's add-in
   API can't crop or duplicate shapes, so the current slide is exported, its XML is
   edited, and the edited copy replaces the original using the deck's own layout.
   Everything happens in memory inside the panel: nothing is sent anywhere and nothing
   is stored. Only the slide's own XML is edited, never masters or layouts. */
(function () {
  "use strict";

  var ST = window.ST;
  var NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  var NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  var EMU = 12700; // per point

  // ---------- Zip writing (stored, no compression) ----------
  var CRC = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  // files: [{ name, bytes }] in order. Returns a Uint8Array.
  function writeZip(files) {
    var enc = new TextEncoder(), parts = [], central = [], offset = 0;
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = f.bytes, crc = crc32(data);
      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint16(8, 0, true); h.setUint32(14, crc, true); h.setUint32(18, data.length, true);
      h.setUint32(22, data.length, true); h.setUint16(26, name.length, true);
      parts.push(new Uint8Array(h.buffer), name, data);
      var d = new DataView(new ArrayBuffer(46));
      d.setUint32(0, 0x02014b50, true); d.setUint16(4, 20, true); d.setUint16(6, 20, true); d.setUint16(8, 0x0800, true);
      d.setUint32(16, crc, true); d.setUint32(20, data.length, true); d.setUint32(24, data.length, true);
      d.setUint16(28, name.length, true); d.setUint32(42, offset, true);
      central.push(new Uint8Array(d.buffer), name);
      offset += 30 + name.length + data.length;
    });
    var size = central.reduce(function (s, p) { return s + p.length; }, 0);
    var e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, size, true); e.setUint32(16, offset, true);
    var all = parts.concat(central, [new Uint8Array(e.buffer)]);
    var out = new Uint8Array(offset + size + 22), at = 0;
    all.forEach(function (p) { out.set(p, at); at += p.length; });
    return out;
  }
  function bytesToB64(bytes) {
    var s = "";
    for (var i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }

  // ---------- Slide XML helpers ----------
  function kids(el, ns, name) { return Array.prototype.filter.call(el.childNodes, function (n) { return n.nodeType === 1 && n.namespaceURI === ns && (!name || n.localName === name); }); }
  function kid(el, ns, name) { return kids(el, ns, name)[0] || null; }
  function make(doc, ns, name) { return doc.createElementNS(ns, (ns === NS_A ? "a:" : "p:") + name); }

  var KINDS = { sp: ["nvSpPr", "spPr"], pic: ["nvPicPr", "spPr"], cxnSp: ["nvCxnSpPr", "spPr"], grpSp: ["nvGrpSpPr", "grpSpPr"], graphicFrame: ["nvGraphicFramePr", null] };
  function cNvPr(el) { var k = KINDS[el.localName], nv = k && kid(el, NS_P, k[0]); return nv ? kid(nv, NS_P, "cNvPr") : null; }
  // The element holding off and ext: a:xfrm in the shape properties, or p:xfrm for frames.
  function xfrmOf(el, create) {
    var k = KINDS[el.localName];
    if (!k) return null;
    if (!k[1]) return kid(el, NS_P, "xfrm");
    var props = kid(el, NS_P, k[1]);
    if (!props) return null;
    var x = kid(props, NS_A, "xfrm");
    if (!x && create) { x = make(el.ownerDocument, NS_A, "xfrm"); props.insertBefore(x, props.firstChild); }
    return x;
  }
  function readBox(el) {
    var x = xfrmOf(el), off = x && kid(x, NS_A, "off"), ext = x && kid(x, NS_A, "ext");
    if (!off || !ext) return null;
    return { left: +off.getAttribute("x") / EMU, top: +off.getAttribute("y") / EMU, width: +ext.getAttribute("cx") / EMU, height: +ext.getAttribute("cy") / EMU,
             rot: (+x.getAttribute("rot") || 0) / 60000, flipH: x.getAttribute("flipH") === "1", flipV: x.getAttribute("flipV") === "1" };
  }
  // box: { left, top, width, height } in points, plus optional rot (degrees), flipH, flipV.
  function writeBox(el, box, opts) {
    opts = opts || {};
    var doc = el.ownerDocument, x = xfrmOf(el, true);
    var off = kid(x, NS_A, "off") || x.insertBefore(make(doc, NS_A, "off"), x.firstChild);
    var ext = kid(x, NS_A, "ext") || x.insertBefore(make(doc, NS_A, "ext"), off.nextSibling);
    off.setAttribute("x", Math.round(box.left * EMU)); off.setAttribute("y", Math.round(box.top * EMU));
    if (!opts.positionOnly) { ext.setAttribute("cx", Math.round(box.width * EMU)); ext.setAttribute("cy", Math.round(box.height * EMU)); }
    if (opts.turn) {
      var r = Math.round((((box.rot || 0) % 360) + 360) % 360 * 60000);
      if (r && r !== 21600000) x.setAttribute("rot", r); else x.removeAttribute("rot");
      if (box.flipH) x.setAttribute("flipH", "1"); else x.removeAttribute("flipH");
      if (box.flipV) x.setAttribute("flipV", "1"); else x.removeAttribute("flipV");
    }
  }

  // Finds the slide element for each API shape: by frame within 1 pt, then by name, then id.
  // tags: element names to look at (top level of the shape tree only).
  function matchShapes(doc, shapes, tags) {
    var tree = doc.getElementsByTagNameNS(NS_P, "spTree")[0];
    var all = tree ? kids(tree, NS_P).filter(function (el) { return tags.indexOf(el.localName) >= 0; }) : [];
    var info = all.map(function (el) { var c = cNvPr(el); return { el: el, name: c ? c.getAttribute("name") : "", id: c ? c.getAttribute("id") : "", box: readBox(el), used: false }; });
    return shapes.map(function (p) {
      var near = function (i) { return i.box && Math.abs(i.box.left - p.left) < 1 && Math.abs(i.box.top - p.top) < 1 && Math.abs(i.box.width - p.width) < 1 && Math.abs(i.box.height - p.height) < 1; };
      var hit = info.filter(function (i) { return !i.used && near(i); });
      if (hit.length > 1) hit = hit.filter(function (i) { return i.name === p.name; }).concat(hit).slice(0, 1);
      if (!hit.length) hit = info.filter(function (i) { return !i.used && !i.box && p.name && i.name === p.name; });
      if (!hit.length) hit = info.filter(function (i) { return !i.used && p.id && i.id === String(p.id); });
      if (!hit.length) return null;
      hit[0].used = true;
      return hit[0].el;
    });
  }

  function supported() { return ST.inOffice && ST.supports("1.8") && typeof DecompressionStream === "function" && typeof TextEncoder === "function" && !!ST.zip; }

  // Exports the selected slide, lets editFn change its XML document, then swaps the copy in.
  // editFn(doc) returns a result with a count; nothing is swapped when the count is 0.
  // Resolves to that result, plus slideId: the new slide's id when one was inserted.
  function edit(context, editFn) {
    var sel = context.presentation.getSelectedSlides(), deck = context.presentation.slides;
    sel.load("items/id");
    deck.load("items/id");
    var slide, result;
    return context.sync().then(function () {
      if (!sel.items.length) throw new Error("No slide");
      slide = sel.items[0];
      var data = slide.exportAsBase64();
      return context.sync().then(function () { return data.value; });
    }).then(function (value) {
      return ST.zip.readZip(ST.zip.b64ToBytes(value), function () { return true; });
    }).then(function (files) {
      var slideFiles = files.filter(function (f) { return /^ppt\/slides\/slide\d+\.xml$/.test(f.name); });
      if (slideFiles.length !== 1) throw new Error("Unexpected export");
      var part = slideFiles[0], doc = new DOMParser().parseFromString(part.text, "application/xml");
      result = editFn(doc) || { count: 0 };
      if (!result.count) return result;
      var xml = new XMLSerializer().serializeToString(doc);
      if (!/^<\?xml/.test(xml)) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
      var enc = new TextEncoder();
      var b64 = bytesToB64(writeZip(files.map(function (f) { return { name: f.name, bytes: f === part ? enc.encode(xml) : f.bytes }; })));
      var index = deck.items.map(function (s) { return s.id; }).indexOf(slide.id);
      context.presentation.insertSlidesFromBase64(b64, { formatting: "UseDestinationTheme", targetSlideId: slide.id });
      b64 = null;
      return context.sync().then(function () {
        slide.delete();
        return context.sync();
      }).then(function () {
        var after = context.presentation.slides;
        after.load("items/id");
        return context.sync().then(function () {
          var fresh = after.items[index];
          if (fresh) result.slideId = fresh.id;
          if (fresh && ST.supports("1.5")) { context.presentation.setSelectedSlides([fresh.id]); return context.sync(); }
        });
      }).then(function () { return result; });
    });
  }

  ST.rebuild = { NS_P: NS_P, NS_A: NS_A, EMU: EMU, crc32: crc32, writeZip: writeZip, bytesToB64: bytesToB64,
                 kids: kids, kid: kid, make: make, cNvPr: cNvPr, xfrmOf: xfrmOf, readBox: readBox, writeBox: writeBox,
                 matchShapes: matchShapes, supported: supported, edit: edit };
})();
