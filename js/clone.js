/* Copies for Transform, like Illustrator's Transform effect: one set of Scale, Reflect,
   Rotate and Move values around a reference point, applied again for every copy, so
   each copy builds on the one before. Copies are added to the slide XML through the
   shared slide rebuild (js/rebuild.js), which works in memory and edits only the slide. */
(function () {
  "use strict";

  var ST = window.ST, R = ST.rebuild, NS_P = R.NS_P, NS_A = R.NS_A;
  var TAGS = ["sp", "pic", "cxnSp", "grpSp", "graphicFrame"];
  var RAD = Math.PI / 180;

  // ---------- Maths (pure) ----------
  // A frame is { cx, cy, w, h, rot, flipH, flipV }: centre and unrotated size in points,
  // rotation in degrees clockwise (as PowerPoint stores it) and the flips.
  function turn(x, y, deg) { var c = Math.cos(deg * RAD), s = Math.sin(deg * RAD); return [x * c - y * s, x * s + y * c]; }
  function refPoint(F, fx, fy) { var v = turn((fx - 0.5) * F.w, (fy - 0.5) * F.h, F.rot); return [F.cx + v[0], F.cy + v[1]]; }
  function bounds(F) {
    var c = Math.abs(Math.cos(F.rot * RAD)), s = Math.abs(Math.sin(F.rot * RAD));
    var hw = (F.w * c + F.h * s) / 2, hh = (F.w * s + F.h * c) / 2;
    return { left: F.cx - hw, right: F.cx + hw, top: F.cy - hh, bottom: F.cy + hh };
  }
  // t: { sx, sy (factors), angle (degrees, counterclockwise like Illustrator), reflectX, reflectY,
  //      fx, fy (reference point, 0 to 1), gap (true: moves are gaps between edges), mh, mv (points or null) }
  function step(F, t) {
    var N = { cx: F.cx, cy: F.cy, w: F.w, h: F.h, rot: F.rot, flipH: F.flipH, flipV: F.flipV };
    var P = refPoint(F, t.fx, t.fy);
    // Scale along the shape's own sides, keeping the reference point still
    N.w = F.w * t.sx; N.h = F.h * t.sy;
    var v = turn((t.fx - 0.5) * N.w, (t.fy - 0.5) * N.h, N.rot);
    N.cx = P[0] - v[0]; N.cy = P[1] - v[1];
    // Reflect X mirrors top to bottom, Reflect Y mirrors left to right, through the reference point
    if (t.reflectX) { N.cy = 2 * P[1] - N.cy; N.rot = -N.rot; N.flipV = !N.flipV; }
    if (t.reflectY) { N.cx = 2 * P[0] - N.cx; N.rot = -N.rot; N.flipH = !N.flipH; }
    // Rotate around the reference point. Positive is counterclockwise, like Illustrator.
    if (t.angle) {
      var a = -t.angle, d = turn(N.cx - P[0], N.cy - P[1], a);
      N.cx = P[0] + d[0]; N.cy = P[1] + d[1]; N.rot += a;
    }
    // Move
    if (t.gap) {
      var was = bounds(F), now = bounds(N);
      if (t.mh !== null) N.cx += t.mh >= 0 ? was.right + t.mh - now.left : was.left + t.mh - now.right;
      if (t.mv !== null) N.cy += t.mv >= 0 ? was.bottom + t.mv - now.top : was.top + t.mv - now.bottom;
    } else {
      N.cx += t.mh || 0; N.cy += t.mv || 0;
    }
    N.rot = ((N.rot % 360) + 360) % 360;
    if (Math.abs(N.rot - 360) < 1e-9) N.rot = 0;
    return N;
  }
  function frames(F, t, n) { var out = []; for (var i = 0; i < n; i++) { F = step(F, t); out.push(F); } return out; }
  // Copies in rows: each row is the original moved down, followed by n copies across.
  // Move down becomes the row step; when it's empty a row is the shape's height plus t.vgap.
  function gridFrames(F, t, n, rows) {
    if (!(rows > 1)) return frames(F, t, n);
    var across = {}, out = [];
    Object.keys(t).forEach(function (k) { across[k] = t[k]; });
    across.mv = null;
    var h = bounds(F).bottom - bounds(F).top, vgap = t.vgap || 0;
    var stepDown = t.gap ? h + (t.mv !== null ? t.mv : vgap) : (t.mv !== null ? t.mv : h + vgap);
    for (var r = 0; r < rows; r++) {
      var base = { cx: F.cx, cy: F.cy + r * stepDown, w: F.w, h: F.h, rot: F.rot, flipH: F.flipH, flipV: F.flipV };
      if (r) out.push(base);
      out = out.concat(frames(base, across, n));
    }
    return out;
  }
  function frameOf(box) { return { cx: box.left + box.width / 2, cy: box.top + box.height / 2, w: box.width, h: box.height, rot: box.rot || 0, flipH: !!box.flipH, flipV: !!box.flipV }; }
  function boxOf(F) { return { left: F.cx - F.w / 2, top: F.cy - F.h / 2, width: F.w, height: F.h, rot: F.rot, flipH: F.flipH, flipV: F.flipV }; }
  function corners(F) { return [[-.5, -.5], [.5, -.5], [.5, .5], [-.5, .5]].map(function (k) { var v = turn(k[0] * F.w, k[1] * F.h, F.rot); return [F.cx + v[0], F.cy + v[1]]; }); }

  // ---------- Slide XML ----------
  function isTable(el) {
    if (el.localName !== "graphicFrame") return false;
    var g = el.getElementsByTagNameNS(NS_A, "graphicData")[0];
    return !!g && /table$/.test(g.getAttribute("uri") || "");
  }
  function maxId(doc) {
    var m = 0;
    Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS_P, "cNvPr"), function (c) { m = Math.max(m, +c.getAttribute("id") || 0); });
    return m;
  }
  function removeAll(root, ns, name) { Array.prototype.slice.call(root.getElementsByTagNameNS(ns, name)).forEach(function (n) { n.parentNode.removeChild(n); }); }
  // A copy needs its own ids and name, and must not claim the original's placeholder slot or connections.
  function tidyCopy(el, ids, name) {
    Array.prototype.forEach.call(el.getElementsByTagNameNS(NS_P, "cNvPr"), function (c) {
      c.setAttribute("id", ids.next++);
      R.kids(c, NS_A, "extLst").forEach(function (x) { c.removeChild(x); });
    });
    R.cNvPr(el).setAttribute("name", name);
    removeAll(el, NS_P, "ph");
    removeAll(el, NS_A, "stCxn");
    removeAll(el, NS_A, "endCxn");
  }

  // jobs: [{ shape, t, copies }]. Resolves { made, shapes, missed, fixed, names, slideId }.
  function apply(context, jobs) {
    var out = { made: 0, shapes: 0, missed: [], fixed: 0, names: [] };
    return R.edit(context, function (doc) {
      var els = R.matchShapes(doc, jobs.map(function (j) { var s = j.shape; return { left: s.left, top: s.top, width: s.width, height: s.height, name: s.name, id: s.id }; }), TAGS);
      var ids = { next: maxId(doc) + 1 };
      els.forEach(function (el, i) {
        var j = jobs[i], s = j.shape;
        if (!el) { out.missed.push(s.name || "a shape"); return; }
        var xml = R.readBox(el) || {}, frame = el.localName === "graphicFrame", table = isTable(el), t = j.t;
        if (frame && (t.angle || t.reflectX || t.reflectY || (table && (t.sx !== 1 || t.sy !== 1)))) out.fixed++;
        // Frames can't rotate or flip, and a table's size comes from its rows and columns
        if (frame) t = Object.assign({}, t, { angle: 0, reflectX: false, reflectY: false, sx: table ? 1 : t.sx, sy: table ? 1 : t.sy });
        var F = frameOf({ left: s.left, top: s.top, width: s.width, height: s.height, rot: frame ? 0 : xml.rot, flipH: xml.flipH, flipV: xml.flipV });
        var base = (R.cNvPr(el).getAttribute("name") || "Shape").replace(/ copy \d+$/, ""), after = el;
        gridFrames(F, t, j.copies, j.rows || 1).forEach(function (G, k) {
          var c = el.cloneNode(true), name = base + " copy " + (k + 1);
          tidyCopy(c, ids, name);
          R.writeBox(c, boxOf(G), { turn: !frame, positionOnly: table });
          after.parentNode.insertBefore(c, after.nextSibling);
          after = c;
          out.names.push(name); out.made++;
        });
        out.shapes++;
      });
      return { count: out.made };
    }).then(function (r) {
      out.slideId = r.slideId;
      if (!r.slideId || !out.names.length || !ST.supports("1.5")) return out;
      // Select the new copies
      var slide = context.presentation.slides.getItem(r.slideId), shapes = slide.shapes;
      shapes.load("items/id,items/name");
      return context.sync().then(function () {
        var ids = shapes.items.filter(function (x) { return out.names.indexOf(x.name) >= 0; }).map(function (x) { return x.id; });
        if (ids.length) { slide.setSelectedShapes(ids); return context.sync(); }
      }).catch(function () {}).then(function () { return out; });
    });
  }

  ST.clone = { step: step, frames: frames, gridFrames: gridFrames, frameOf: frameOf, boxOf: boxOf, bounds: bounds, corners: corners, refPoint: refPoint, apply: apply, supported: R.supported, _tidyCopy: tidyCopy };
})();
