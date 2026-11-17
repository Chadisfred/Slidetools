/* Picture cropping for Resize. PowerPoint's add-in API can't crop a picture, so the
   slide is exported, the picture's crop and frame are set in the slide XML, and the
   copy replaces the original. Everything happens in memory: nothing is sent anywhere
   and nothing is stored. Only the slide's own XML is edited, never masters or layouts. */
(function () {
  "use strict";

  var ST = window.ST;
  var NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  var NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  var EMU = 12700; // per point

  // ---------- Crop maths (pure) ----------
  // pic: { left, top, width, height, crop: {l,t,r,b} as fractions }
  // frame: the new { left, top, width, height }; f: anchor factors { x, y } from 0 to 1.
  // Keeps the picture's scale and place, so the frame shows more or less of it. If the
  // frame would run past the picture's edges, the picture grows just enough to cover it.
  function cropFor(pic, frame, f) {
    var c = pic.crop || { l: 0, t: 0, r: 0, b: 0 };
    // The whole picture as it sits on the slide now, before cropping
    var full = function (pos, size, lo, hi) { var vis = 1 - lo - hi; if (!(vis > 0)) { vis = 1; lo = 0; } var whole = size / vis; return { start: pos - lo * whole, size: whole }; };
    var X = full(pic.left, pic.width, c.l, c.r), Y = full(pic.top, pic.height, c.t, c.b);
    // One scale for both directions, so the picture is never stretched
    var k = Math.max(1, frame.width / X.size, frame.height / Y.size);
    var fit = function (A, npos, nsize, fa) {
      var start = A.start, size = A.size;
      if (k > 1) { var p = npos + fa * nsize; start = p - (p - start) * k; size *= k; }
      start = Math.min(npos, Math.max(npos + nsize - size, start));
      return [Math.max(0, (npos - start) / size), Math.max(0, (start + size - npos - nsize) / size)];
    };
    var x = fit(X, frame.left, frame.width, f.x), y = fit(Y, frame.top, frame.height, f.y);
    return { l: x[0], r: x[1], t: y[0], b: y[1] };
  }

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

  // ---------- Slide XML ----------
  function kids(el, ns, name) { return Array.prototype.filter.call(el.childNodes, function (n) { return n.nodeType === 1 && n.namespaceURI === ns && n.localName === name; }); }
  function kid(el, ns, name) { return kids(el, ns, name)[0] || null; }
  function make(doc, ns, name) { return doc.createElementNS(ns, (ns === NS_A ? "a:" : "p:") + name); }

  // Finds the <p:pic> for each picture, first by its frame, then by name.
  function matchPictures(doc, pics) {
    var all = Array.prototype.slice.call(doc.getElementsByTagNameNS(NS_P, "pic"));
    var info = all.map(function (el) {
      var nv = kid(el, NS_P, "nvPicPr"), c = nv && kid(nv, NS_P, "cNvPr");
      var sp = kid(el, NS_P, "spPr"), x = sp && kid(sp, NS_A, "xfrm");
      var off = x && kid(x, NS_A, "off"), ext = x && kid(x, NS_A, "ext");
      return { el: el, name: c ? c.getAttribute("name") : "", id: c ? c.getAttribute("id") : "",
        box: off && ext ? [+off.getAttribute("x") / EMU, +off.getAttribute("y") / EMU, +ext.getAttribute("cx") / EMU, +ext.getAttribute("cy") / EMU] : null, used: false };
    });
    return pics.map(function (p) {
      var near = function (i) { return i.box && Math.abs(i.box[0] - p.left) < 1 && Math.abs(i.box[1] - p.top) < 1 && Math.abs(i.box[2] - p.width) < 1 && Math.abs(i.box[3] - p.height) < 1; };
      var hit = info.filter(function (i) { return !i.used && near(i); });
      if (hit.length > 1) hit = hit.filter(function (i) { return i.name === p.name; }).concat(hit).slice(0, 1);
      if (!hit.length) hit = info.filter(function (i) { return !i.used && !i.box && p.name && i.name === p.name; });
      if (!hit.length) hit = info.filter(function (i) { return !i.used && p.id && i.id === String(p.id); });
      if (!hit.length) return null;
      hit[0].used = true;
      return hit[0].el;
    });
  }
  function readCrop(pic) {
    var bf = kid(pic, NS_P, "blipFill"), r = bf && kid(bf, NS_A, "srcRect"), v = function (k) { return r && r.getAttribute(k) ? +r.getAttribute(k) / 100000 : 0; };
    return { l: v("l"), t: v("t"), r: v("r"), b: v("b") };
  }
  function writePicture(doc, pic, frame, crop) {
    var bf = kid(pic, NS_P, "blipFill");
    var r = kid(bf, NS_A, "srcRect");
    if (!r) { r = make(doc, NS_A, "srcRect"); var blip = kid(bf, NS_A, "blip"); bf.insertBefore(r, blip ? blip.nextSibling : bf.firstChild); }
    ["l", "t", "r", "b"].forEach(function (k) { var n = Math.round(crop[k] * 100000); if (n) r.setAttribute(k, n); else r.removeAttribute(k); });
    var sp = kid(pic, NS_P, "spPr");
    var x = kid(sp, NS_A, "xfrm");
    if (!x) { x = make(doc, NS_A, "xfrm"); sp.insertBefore(x, sp.firstChild); }
    var off = kid(x, NS_A, "off") || x.insertBefore(make(doc, NS_A, "off"), x.firstChild);
    var ext = kid(x, NS_A, "ext") || x.insertBefore(make(doc, NS_A, "ext"), off.nextSibling);
    off.setAttribute("x", Math.round(frame.left * EMU)); off.setAttribute("y", Math.round(frame.top * EMU));
    ext.setAttribute("cx", Math.round(frame.width * EMU)); ext.setAttribute("cy", Math.round(frame.height * EMU));
  }

  // ---------- Detection ----------
  function canCrop() { return ST.inOffice && ST.supports("1.8") && typeof DecompressionStream === "function" && typeof TextEncoder === "function" && !!ST.zip; }

  // Resolves to a list of booleans: true where the shape is a picture.
  function pictureFlags(context, items) {
    var ph = items.filter(function (s) { return s.type === "Placeholder"; });
    if (!ph.length || !ST.supports("1.8")) return Promise.resolve(items.map(function (s) { return s.type === "Image"; }));
    ph.forEach(function (s) { s.placeholderFormat.load("containedType"); });
    return context.sync().then(function () {
      return items.map(function (s) { return s.type === "Image" || (s.type === "Placeholder" && s.placeholderFormat.containedType === "Image"); });
    }).catch(function () { return items.map(function (s) { return s.type === "Image"; }); });
  }

  // jobs: [{ shape, frame }] where frame is the new size in points. Resolves { cropped, missed }.
  function apply(context, jobs) {
    var sel = context.presentation.getSelectedSlides(), deck = context.presentation.slides;
    sel.load("items/id");
    deck.load("items/id");
    var slide, b64, missed = [];
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
      // Shape values are already in points
      var pics = jobs.map(function (j) { var s = j.shape; return { left: s.left, top: s.top, width: s.width, height: s.height, name: s.name, id: s.id }; });
      var els = matchPictures(doc, pics), done = 0;
      els.forEach(function (el, i) {
        if (!el) { missed.push(jobs[i].shape.name || "a picture"); return; }
        var p = pics[i]; p.crop = readCrop(el);
        writePicture(doc, el, jobs[i].frame, cropFor(p, jobs[i].frame, jobs[i].anchor));
        done++;
      });
      if (!done) return 0;
      var xml = new XMLSerializer().serializeToString(doc);
      if (!/^<\?xml/.test(xml)) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
      var enc = new TextEncoder();
      b64 = bytesToB64(writeZip(files.map(function (f) { return { name: f.name, bytes: f === part ? enc.encode(xml) : f.bytes }; })));
      var index = deck.items.map(function (s) { return s.id; }).indexOf(slide.id);
      context.presentation.insertSlidesFromBase64(b64, { formatting: "UseDestinationTheme", targetSlideId: slide.id });
      return context.sync().then(function () {
        slide.delete();
        return context.sync();
      }).then(function () {
        b64 = null;
        var after = context.presentation.slides;
        after.load("items/id");
        return context.sync().then(function () {
          var fresh = after.items[index];
          if (fresh && ST.supports("1.5")) { context.presentation.setSelectedSlides([fresh.id]); return context.sync(); }
        });
      }).then(function () { return done; });
    }).then(function (done) { return { cropped: done, missed: missed }; });
  }

  ST.crop = { cropFor: cropFor, writeZip: writeZip, crc32: crc32, bytesToB64: bytesToB64, canCrop: canCrop, pictureFlags: pictureFlags, apply: apply,
              _matchPictures: matchPictures, _readCrop: readCrop, _writePicture: writePicture };
})();
