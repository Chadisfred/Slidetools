/* Picture cropping for the Size tool. Pictures are cropped to fill a new frame instead of
   being stretched. The crop is written to the slide XML through the shared slide rebuild
   (js/rebuild.js), which works in memory and edits only the slide. */
(function () {
  "use strict";

  var ST = window.ST, R = ST.rebuild, NS_P = R.NS_P, NS_A = R.NS_A, kid = R.kid, make = R.make;

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

  // ---------- Slide XML ----------
  function readCrop(pic) {
    var bf = kid(pic, NS_P, "blipFill"), r = bf && kid(bf, NS_A, "srcRect"), v = function (k) { return r && r.getAttribute(k) ? +r.getAttribute(k) / 100000 : 0; };
    return { l: v("l"), t: v("t"), r: v("r"), b: v("b") };
  }
  function writePicture(doc, pic, frame, crop) {
    var bf = kid(pic, NS_P, "blipFill");
    var r = kid(bf, NS_A, "srcRect");
    if (!r) { r = make(doc, NS_A, "srcRect"); var blip = kid(bf, NS_A, "blip"); bf.insertBefore(r, blip ? blip.nextSibling : bf.firstChild); }
    ["l", "t", "r", "b"].forEach(function (k) { var n = Math.round(crop[k] * 100000); if (n) r.setAttribute(k, n); else r.removeAttribute(k); });
    R.writeBox(pic, frame);
  }

  // Resolves to a list of booleans: true where the shape is a picture.
  function pictureFlags(context, items) {
    var ph = items.filter(function (s) { return s.type === "Placeholder"; });
    if (!ph.length || !ST.supports("1.8")) return Promise.resolve(items.map(function (s) { return s.type === "Image"; }));
    ph.forEach(function (s) { s.placeholderFormat.load("containedType"); });
    return context.sync().then(function () {
      return items.map(function (s) { return s.type === "Image" || (s.type === "Placeholder" && s.placeholderFormat.containedType === "Image"); });
    }).catch(function () { return items.map(function (s) { return s.type === "Image"; }); });
  }

  // jobs: [{ shape, frame, anchor }] where frame is the new size in points. Resolves { cropped, missed }.
  function apply(context, jobs) {
    var missed = [];
    return R.edit(context, function (doc) {
      // Shape values are already in points
      var pics = jobs.map(function (j) { var s = j.shape; return { left: s.left, top: s.top, width: s.width, height: s.height, name: s.name, id: s.id }; });
      var els = R.matchShapes(doc, pics, ["pic"]), done = 0;
      els.forEach(function (el, i) {
        if (!el) { missed.push(jobs[i].shape.name || "a picture"); return; }
        var p = pics[i]; p.crop = readCrop(el);
        writePicture(doc, el, jobs[i].frame, cropFor(p, jobs[i].frame, jobs[i].anchor));
        done++;
      });
      return { count: done };
    }).then(function (r) { return { cropped: r.count, missed: missed }; });
  }

  ST.crop = { cropFor: cropFor, writeZip: R.writeZip, crc32: R.crc32, bytesToB64: R.bytesToB64, canCrop: R.supported, pictureFlags: pictureFlags, apply: apply,
              _readCrop: readCrop, _writePicture: writePicture };
})();
