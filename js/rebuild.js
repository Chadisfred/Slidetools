/* Slide rebuild, shared by picture cropping, Copies, Images and Typography. PowerPoint's add-in
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

  function serialize(doc) {
    var xml = new XMLSerializer().serializeToString(doc);
    return /^<\?xml/.test(xml) ? xml : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
  }
  // Lets an edit add a picture to the exported slide: the image goes into ppt/media, the
  // slide gets a relationship to it, and the content types list its extension.
  var NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
  var NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
  var IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  var TYPES = { png: "image/png", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp" };
  function mediaHelper(files, slideName) {
    var relsName = slideName.replace(/slides\/(slide\d+\.xml)$/, "slides/_rels/$1.rels"), docs = {}, added = [], n = 0;
    function open(name, fallback) {
      if (docs[name]) return docs[name];
      var f = files.filter(function (x) { return x.name === name; })[0];
      docs[name] = new DOMParser().parseFromString(f ? f.text : fallback, "application/xml");
      return docs[name];
    }
    return {
      // bytes: Uint8Array, ext: png, jpeg, gif or bmp. Returns the relationship id.
      addImage: function (bytes, ext) {
        var names = files.map(function (f) { return f.name; }).concat(added.map(function (f) { return f.name; }));
        var name; do { n++; name = "ppt/media/slidetools" + n + "." + ext; } while (names.indexOf(name) >= 0);
        added.push({ name: name, bytes: bytes });
        var rels = open(relsName, '<Relationships xmlns="' + NS_REL + '"/>'), root = rels.documentElement, ids = {};
        Array.prototype.forEach.call(root.getElementsByTagNameNS(NS_REL, "Relationship"), function (r) { ids[r.getAttribute("Id")] = 1; });
        var id, k = 0; do { k++; id = "rIdST" + k; } while (ids[id]);
        var rel = rels.createElementNS(NS_REL, "Relationship");
        rel.setAttribute("Id", id); rel.setAttribute("Type", IMAGE_REL); rel.setAttribute("Target", "../media/" + name.split("/").pop());
        root.appendChild(rel);
        var ct = open("[Content_Types].xml", '<Types xmlns="' + NS_CT + '"/>'), types = ct.documentElement;
        var has = Array.prototype.some.call(types.getElementsByTagNameNS(NS_CT, "Default"), function (d) { return String(d.getAttribute("Extension")).toLowerCase() === ext; });
        if (!has) { var d = ct.createElementNS(NS_CT, "Default"); d.setAttribute("Extension", ext); d.setAttribute("ContentType", TYPES[ext]); types.insertBefore(d, types.firstChild); }
        return id;
      },
      changed: function () { var out = {}, enc = new TextEncoder(); Object.keys(docs).forEach(function (k) { out[k] = enc.encode(serialize(docs[k])); }); return out; },
      added: function () {
        // A rels file that didn't exist in the export is new
        var have = files.map(function (f) { return f.name; }), out = added.slice(), enc = new TextEncoder();
        Object.keys(docs).forEach(function (k) { if (have.indexOf(k) < 0) out.push({ name: k, bytes: enc.encode(serialize(docs[k])) }); });
        return out;
      }
    };
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
      var media = mediaHelper(files, part.name);
      result = editFn(doc, { addImage: media.addImage }) || { count: 0 };
      if (!result.count) return result;
      var enc = new TextEncoder(), changed = media.changed();
      changed[part.name] = enc.encode(serialize(doc));
      var list = files.map(function (f) { return { name: f.name, bytes: changed[f.name] || f.bytes }; }).concat(media.added());
      var b64 = bytesToB64(writeZip(list));
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
                 matchShapes: matchShapes, supported: supported, edit: edit, _mediaHelper: mediaHelper };
})();
