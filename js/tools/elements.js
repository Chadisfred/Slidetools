/* Elements tool (id "elements"): prebuilt design elements, dropped onto the slide as one group of
   shapes with the text inside each shape. Colours are theme colours, so an element follows the
   template it's placed in. The element replaces the selected shape at its size and position, or
   takes a typed width (g, col and the other grid names) at the grid area's top left.
   Add-ins can't set bullets or theme tints directly, so the group is written into the slide XML
   through the shared slide rebuild (js/rebuild.js), in memory. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$, R = ST.rebuild, PT = ST.PT;
  var NS_P = R.NS_P, NS_A = R.NS_A;
  var SHAPE_KINDS = ["sp", "pic", "grpSp", "cxnSp", "graphicFrame"];
  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"><rect x="3" y="2.5" width="12" height="4" rx="1" fill="currentColor" fill-opacity="0.25"/><rect x="3" y="7.5" width="12" height="5" rx="1"/><rect x="3" y="13.5" width="12" height="2.5" rx="1"/></svg>';

  var defaults = {
    elKind: "stat", elColor: "accent1", elPlace: "selection", elW: "3col", elH: "",
    elStTitle: "Region name", elStValue: "+5.0%", elStArrow: true, elStSub: "Short summary of the change",
    elStBullets: "First point about the result\nSecond point", elStFootB: "120 units", elStFootN: "(15%)",
    elPtTitle: "Key insight", elPtBullets: "First point\nSecond point\nThird point", elPtLabel: "Action needed:", elPtFoot: "Next step",
    elBnNum: "75%", elBnLabel: "What the number measures", elBnNote: "A short note about where it comes from"
  };
  ST.addDefaults(defaults);

  var KINDS = [
    { id: "stat", name: "Stat card", tip: "A bordered card: a coloured header with a title and a value with an up or down arrow, a summary band, bullets and a footer figure." },
    { id: "pointer", name: "Pointer card", tip: "A dark header, a band of bullets and a light footer, joined by small pointers." },
    { id: "bignum", name: "Big number", tip: "A large figure in the theme colour with a label and a note, beside a colour bar." }
  ];
  var COLORS = [["accent1", "Accent 1"], ["accent2", "Accent 2"], ["accent3", "Accent 3"], ["accent4", "Accent 4"], ["accent5", "Accent 5"], ["accent6", "Accent 6"], ["tx2", "Dark 2"]];
  var themeHex = {}; // slot: "#RRGGBB", read from the presentation's theme for the chips

  // ---------- XML pieces (pure) ----------
  function e(pt) { return Math.round(pt * 12700); }
  function x(t) { return String(t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  // A theme colour, optionally darker (lumMod) or lighter (lumMod plus lumOff)
  function clr(slot, mod, off, alpha) {
    return '<a:schemeClr val="' + slot + '">' + (mod ? '<a:lumMod val="' + mod + '"/>' : "") + (off ? '<a:lumOff val="' + off + '"/>' : "") + (alpha ? '<a:alpha val="' + alpha + '"/>' : "") + "</a:schemeClr>";
  }
  function palette(slot) {
    return {
      dark: clr(slot, 75000), base: clr(slot), mid: clr(slot, 60000, 40000), light: clr(slot, 40000, 60000), pale: clr(slot, 20000, 80000),
      grey: clr("bg1", 95000), rule: clr("bg1", 75000), white: clr("bg1"), text: clr("tx1"), soft: clr("tx1", 65000, 35000),
      // See-through bands, so a picture or texture behind the card shows a little
      tint: clr(slot, null, null, 30000), mist: clr("bg1", 95000, null, 85000)
    };
  }
  // runs: [{ t, b, i, sz, c (colour XML), mj (heading font) }]
  function para(runs, o) {
    o = o || {};
    var ppr = "";
    var attrs = (o.align ? ' algn="' + o.align + '"' : "") + (o.bullet ? ' marL="' + e(o.marL) + '" indent="' + e(-o.marL) + '"' : ' marL="0" indent="0"');
    var inner = (o.spcBef ? '<a:spcBef><a:spcPts val="' + Math.round(o.spcBef * 100) + '"/></a:spcBef>' : "") +
      (o.bullet ? '<a:buClr>' + o.bullet.c + '</a:buClr><a:buSzPct val="100000"/><a:buFont typeface="Arial"/><a:buChar char="' + o.bullet.ch + '"/>' : "<a:buNone/>");
    ppr = "<a:pPr" + attrs + ">" + inner + "</a:pPr>";
    var body = runs.map(function (r) {
      var rpr = '<a:rPr lang="en-US" sz="' + Math.round(r.sz * 100) + '"' + (r.b ? ' b="1"' : ' b="0"') + (r.i ? ' i="1"' : "") + ' dirty="0"><a:solidFill>' + r.c + '</a:solidFill><a:latin typeface="' + (r.mj ? "+mj-lt" : "+mn-lt") + '"/></a:rPr>';
      return "<a:r>" + rpr + "<a:t>" + x(r.t) + "</a:t></a:r>";
    }).join("");
    var end = runs[0] ? '<a:endParaRPr lang="en-US" sz="' + Math.round(runs[0].sz * 100) + '" dirty="0"/>' : "";
    return "<a:p>" + ppr + body + end + "</a:p>";
  }
  function txBody(paras, inset, anchor) {
    return '<p:txBody><a:bodyPr wrap="square" lIns="' + e(inset.l) + '" tIns="' + e(inset.t) + '" rIns="' + e(inset.r) + '" bIns="' + e(inset.b) + '" anchor="' + (anchor || "t") + '" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>' + paras.join("") + "</p:txBody>";
  }
  function sp(id, name, b, o) {
    o = o || {};
    var xf = '<a:xfrm' + (o.rot ? ' rot="' + o.rot + '"' : "") + '><a:off x="' + e(b.x) + '" y="' + e(b.y) + '"/><a:ext cx="' + e(b.w) + '" cy="' + e(b.h) + '"/></a:xfrm>';
    var fill = o.fill ? "<a:solidFill>" + o.fill + "</a:solidFill>" : "<a:noFill/>";
    var tx = o.text || txBody([para([{ t: "", sz: 12, c: clr("tx1") }])], { l: 0, t: 0, r: 0, b: 0 });
    return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + x(name) + '"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr>' + xf + '<a:prstGeom prst="' + (o.geom || "rect") + '"><a:avLst/></a:prstGeom>' + fill + (o.line ? '<a:ln w="' + e(o.line.w) + '"><a:solidFill>' + o.line.c + "</a:solidFill></a:ln>" : "<a:ln><a:noFill/></a:ln>") + "</p:spPr>" + tx + "</p:sp>";
  }
  function group(id, name, b, kids) {
    return '<p:grpSp xmlns:p="' + NS_P + '" xmlns:a="' + NS_A + '"><p:nvGrpSpPr><p:cNvPr id="' + id + '" name="' + x(name) + '"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="' + e(b.x) + '" y="' + e(b.y) + '"/><a:ext cx="' + e(b.w) + '" cy="' + e(b.h) + '"/><a:chOff x="' + e(b.x) + '" y="' + e(b.y) + '"/><a:chExt cx="' + e(b.w) + '" cy="' + e(b.h) + '"/></a:xfrm></p:grpSpPr>' + kids.join("") + "</p:grpSp>";
  }

  // ---------- Text size estimates (pure) ----------
  // How tall wrapped text is. PowerPoint does the real wrapping; this keeps bands close to it.
  function linesFor(text, sz, width) {
    var perLine = Math.max(4, Math.floor(width / (sz * 0.5)));
    return String(text || "").split("\n").reduce(function (n, p) { return n + Math.max(1, Math.ceil(p.length / perLine)); }, 0);
  }
  function textH(text, sz, width) { return linesFor(text, sz, width) * sz * 1.2; }
  function bullets(text) { return String(text || "").split(/\r?\n/).map(function (s) { return s.replace(/^\s*[-•▪*]\s*/, "").trim(); }).filter(Boolean); }
  // The arrow for a value such as +7.0% or -2: up, down, flat for zero, none without a number
  function arrowFor(value) {
    var t = String(value || "").trim(), m = /^([+\-−])?\s*[^\d]*?(\d+(?:[.,]\d+)?)/.exec(t);
    if (!m) return "";
    if (m[1] === "+") return "▲";
    if (m[1] === "-" || m[1] === "−") return "▼";
    return parseFloat(m[2].replace(",", ".")) === 0 ? "▶" : "";
  }

  // ---------- Elements (pure) ----------
  // Each returns { name, kids: [xml], h } for a box { x, y, w, h (optional) }. ids: { next }.
  function scaleFor(w) { return Math.max(0.8, Math.min(1.6, w / 216)); }
  // Spreads extra height into the body band so the element fills a given height
  function fit(heights, body, total) {
    var sum = heights.reduce(function (a, b) { return a + b; }, 0);
    if (total && total > sum) heights[body] += total - sum;
    return heights;
  }
  function statCard(S, b, slot, ids) {
    var k = scaleFor(b.w), c = palette(slot), pad = Math.round(12 * k), W = b.w, inW = W - 2 * pad;
    var tSz = 22 * k, sSz = 14 * k, bSz = 14 * k, fSz = 14 * k, valW = Math.round(W * 0.38), gap = 8 * k;
    var arrow = S.elStArrow !== false ? arrowFor(S.elStValue) : "";
    var list = bullets(S.elStBullets), marL = 16 * k;
    var hHead = Math.max(textH(S.elStTitle, tSz, W - valW - 2 * pad), (arrow ? 2 : 1) * tSz * 1.2) + 2 * pad;
    var hSub = textH(S.elStSub, sSz, inW) + 2 * pad;
    var hList = list.reduce(function (n, t) { return n + textH(t, bSz, inW - marL) + gap; }, 0) - gap + 2.4 * pad;
    var hFoot = fSz * 1.2 + 2.4 * pad;
    var hs = fit([hHead, hSub, hList, hFoot], 2, b.h), y = b.y, kids = [];
    var inset = { l: pad, t: pad, r: pad, b: pad };
    kids.push(sp(ids.next++, "Header", { x: b.x, y: y, w: W, h: hs[0] }, { fill: c.dark, text: txBody([para([{ t: S.elStTitle, b: true, sz: tSz, c: c.white, mj: true }])], { l: pad, t: pad, r: valW, b: pad }) }));
    var valParas = [para([{ t: S.elStValue, b: true, sz: tSz, c: c.white, mj: true }], { align: "r" })];
    if (arrow) valParas.push(para([{ t: arrow, sz: tSz * 0.8, c: c.white }], { align: "r" }));
    kids.push(sp(ids.next++, "Value", { x: b.x + W - valW, y: y, w: valW, h: hs[0] }, { text: txBody(valParas, inset) }));
    y += hs[0];
    kids.push(sp(ids.next++, "Summary", { x: b.x, y: y, w: W, h: hs[1] }, { fill: c.tint, text: txBody([para([{ t: S.elStSub, i: true, sz: sSz, c: c.text }])], inset, "ctr") }));
    y += hs[1];
    kids.push(sp(ids.next++, "Points", { x: b.x, y: y, w: W, h: hs[2] }, { fill: c.mist, text: txBody(list.map(function (t, i) { return para([{ t: t, sz: bSz, c: c.text }], { bullet: { ch: "▪", c: c.dark }, marL: marL, spcBef: i ? gap : 0 }); }), { l: pad, t: pad * 1.2, r: pad, b: pad * 1.2 }) }));
    y += hs[2];
    // Rules between the bands: the header colour, as thick as the outline
    var rule = Math.max(1, Math.round(k * 4) / 4 + 0.25);
    kids.push(sp(ids.next++, "Rule", { x: b.x, y: b.y + hs[0] + hs[1] - rule / 2, w: W, h: rule }, { fill: c.dark }));
    kids.push(sp(ids.next++, "Rule", { x: b.x, y: y - rule / 2, w: W, h: rule }, { fill: c.dark }));
    kids.push(sp(ids.next++, "Footer", { x: b.x, y: y, w: W, h: hs[3] }, { fill: c.mist, text: txBody([para([{ t: S.elStFootB, b: true, sz: fSz, c: c.text, mj: true }, { t: S.elStFootN ? " " + S.elStFootN : "", sz: fSz, c: c.text }].filter(function (r) { return r.t; }))], inset, "ctr") }));
    y += hs[3];
    // An outline round the whole card, on top
    kids.push(sp(ids.next++, "Border", { x: b.x, y: b.y, w: W, h: y - b.y }, { line: { c: c.dark, w: rule } }));
    return { name: "Stat card", kids: kids, h: y - b.y };
  }
  function pointerCard(S, b, slot, ids) {
    var k = scaleFor(b.w), c = palette(slot), pad = Math.round(14 * k), W = b.w, inW = W - 2 * pad;
    var tSz = 22 * k, bSz = 16 * k, lSz = 11 * k, fSz = 16 * k, list = bullets(S.elPtBullets), marL = 16 * k;
    var hHead = textH(S.elPtTitle, tSz, inW) + 2 * pad;
    var hList = list.reduce(function (n, t) { return n + textH(t, bSz, inW - marL) + 10 * k; }, 0) + 2 * pad;
    var hFoot = (S.elPtLabel ? lSz * 1.2 : 0) + textH(S.elPtFoot, fSz, inW) + 2 * pad;
    var hs = fit([hHead, hList, hFoot], 1, b.h), y = b.y, kids = [], inset = { l: pad, t: pad, r: pad, b: pad };
    kids.push(sp(ids.next++, "Header", { x: b.x, y: y, w: W, h: hs[0] }, { fill: c.dark, text: txBody([para([{ t: S.elPtTitle, b: true, sz: tSz, c: c.white, mj: true }])], inset, "ctr") }));
    y += hs[0];
    var top2 = y;
    kids.push(sp(ids.next++, "Points", { x: b.x, y: y, w: W, h: hs[1] }, { fill: c.base, text: txBody(list.map(function (t, i) { return para([{ t: t, sz: bSz, c: c.white }], { bullet: { ch: "▪", c: c.white }, marL: marL, spcBef: i ? 10 * k : 0 }); }), { l: pad, t: pad * 1.2, r: pad, b: pad }) }));
    y += hs[1];
    var top3 = y, fp = [];
    if (S.elPtLabel) fp.push(para([{ t: S.elPtLabel, sz: lSz, c: c.text }]));
    fp.push(para([{ t: S.elPtFoot, b: true, sz: fSz, c: c.text, mj: true }]));
    kids.push(sp(ids.next++, "Footer", { x: b.x, y: y, w: W, h: hs[2] }, { fill: c.light, text: txBody(fp, { l: pad, t: pad * 1.2, r: pad, b: pad }) }));
    y += hs[2];
    // Small pointers: each band points down into the next one
    var pw = Math.round(16 * k), ph = Math.round(8 * k), px = b.x + pad + 4 * k;
    kids.push(sp(ids.next++, "Pointer", { x: px, y: top2, w: pw, h: ph }, { geom: "triangle", rot: 10800000, fill: c.dark }));
    kids.push(sp(ids.next++, "Pointer", { x: px, y: top3, w: pw, h: ph }, { geom: "triangle", rot: 10800000, fill: c.base }));
    return { name: "Pointer card", kids: kids, h: y - b.y };
  }
  function bigNumber(S, b, slot, ids) {
    var k = scaleFor(b.w), c = palette(slot), bar = Math.max(3, Math.round(4 * k)), gap = Math.round(12 * k), W = b.w, inW = W - bar - gap;
    // The number stays on one line: long figures get a smaller size
    var chars = Math.max(2, String(S.elBnNum || "").length), nSz = Math.min(54 * k, inW / (chars * 0.6)), lSz = 14 * k, tSz = 11 * k;
    var h = nSz * 1.1 + textH(S.elBnLabel, lSz, inW) + (S.elBnNote ? 4 * k + textH(S.elBnNote, tSz, inW) : 0);
    h = Math.max(h, b.h || 0);
    var ps = [para([{ t: S.elBnNum, b: true, sz: nSz, c: c.base, mj: true }]), para([{ t: S.elBnLabel, b: true, sz: lSz, c: c.text, mj: true }])];
    if (S.elBnNote) ps.push(para([{ t: S.elBnNote, sz: tSz, c: c.soft }], { spcBef: 4 * k }));
    return { name: "Big number", h: h, kids: [
      sp(ids.next++, "Bar", { x: b.x, y: b.y, w: bar, h: h }, { fill: c.base }),
      sp(ids.next++, "Number", { x: b.x + bar + gap, y: b.y, w: inW, h: h }, { text: txBody(ps, { l: 0, t: 0, r: 0, b: 0 }) })
    ] };
  }
  var BUILDERS = { stat: statCard, pointer: pointerCard, bignum: bigNumber };
  // The whole group's XML for a box. Returns { xml, name, h }.
  function elementXml(kind, S, box, slot, firstId) {
    var ids = { next: firstId + 1 }, built = BUILDERS[kind](S, box, slot, ids);
    var name = "Element " + built.name + " " + firstId;
    return { xml: group(firstId, name, { x: box.x, y: box.y, w: box.w, h: built.h }, built.kids), name: name, h: built.h, used: ids.next - firstId };
  }
  function maxId(doc) {
    var m = 1;
    Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS_P, "cNvPr"), function (c) { m = Math.max(m, +c.getAttribute("id") || 0); });
    return m;
  }
  // The theme's colours, for the chips
  function readTheme(files) {
    var t = (files || []).filter(function (f) { return /^ppt\/theme\/theme\d+\.xml$/.test(f.name); })[0], out = {};
    if (!t) return out;
    var d = new DOMParser().parseFromString(t.text, "application/xml"), scheme = d.getElementsByTagNameNS(NS_A, "clrScheme")[0];
    if (!scheme) return out;
    var map = { accent1: "accent1", accent2: "accent2", accent3: "accent3", accent4: "accent4", accent5: "accent5", accent6: "accent6", dk2: "tx2" };
    Array.prototype.forEach.call(scheme.childNodes, function (n) {
      if (n.nodeType !== 1 || !map[n.localName]) return;
      var v = n.getElementsByTagNameNS(NS_A, "srgbClr")[0], s = n.getElementsByTagNameNS(NS_A, "sysClr")[0];
      var hex = v ? v.getAttribute("val") : s ? s.getAttribute("lastClr") : null;
      if (hex) out[map[n.localName]] = "#" + hex.toUpperCase();
    });
    return out;
  }

  // ---------- Placing ----------
  function targetBox(S) {
    if (S.elPlace === "size") {
      var w = ST.evaluatePt(S.elW), h = String(S.elH || "").trim() ? ST.evaluatePt(S.elH) : null;
      if (w === null || !(w >= 36 && w <= 2000)) return { bad: "Width " + (ST.ptProblem(S.elW) || "must work out to 36 to 2000 points") + ". Try 3col, 8g or 240." };
      if (h !== null && !(h >= 18 && h <= 2000)) return { bad: "Height must work out to 18 to 2000 points, or stay empty to fit the text." };
      var g = ST.calculate();
      return { box: { x: g.left * PT, y: g.top * PT, w: w, h: h || 0 } };
    }
    return { fromSelection: true };
  }
  function place(saved) {
    var S = saved || {};
    if (!saved) Object.keys(defaults).forEach(function (key) { S[key] = ST.state[key]; });
    var t = targetBox(S);
    if (t.bad) { ST.toast(t.bad, true); return; }
    if (!ST.needOffice()) return;
    if (!R.supported()) { ST.toast("Placing elements needs PowerPoint API 1.8.", true); return; }
    var kindName = (KINDS.filter(function (k) { return k.id === S.elKind; })[0] || KINDS[0]).name;
    ST.setRepeat("Place " + kindName, function () { place(S); });
    var placedName = null, replaced = false, missing = false;
    ST.withSelection(["id", "name", "type", "left", "top", "width", "height"], function (context, items) {
      if (t.fromSelection && !items.length) { ST.toast("Select a shape to replace, or choose Size and type a width.", true); return; }
      return ST.confirmSwap("place an element").then(function (ok) {
        if (!ok) return null;
        return R.edit(context, function (doc, extra) {
          var tree = doc.getElementsByTagNameNS(NS_P, "spTree")[0], box = t.box, old = null;
          if (t.fromSelection) {
            var s = items[0];
            old = R.matchShapes(doc, [{ left: s.left, top: s.top, width: s.width, height: s.height, name: s.name, id: s.id }], SHAPE_KINDS)[0];
            if (!old) { missing = true; return { count: 0 }; }
            box = { x: s.left, y: s.top, w: s.width, h: s.height };
          }
          var built = elementXml(S.elKind, S, box, S.elColor || "accent1", maxId(doc) + 1);
          var node = doc.importNode(new DOMParser().parseFromString(built.xml, "application/xml").documentElement, true);
          if (old) { old.parentNode.insertBefore(node, old); old.parentNode.removeChild(old); replaced = true; }
          else tree.appendChild(node);
          placedName = built.name;
          return { count: 1 };
        }).then(function (res) {
          if (!res || !res.count || !res.slideId || !ST.supports("1.5")) return res;
          var slide = context.presentation.slides.getItem(res.slideId), shapes = slide.shapes;
          shapes.load("items/id,items/name");
          return context.sync().then(function () {
            var hit = shapes.items.filter(function (x) { return x.name === placedName; })[0];
            if (hit) { slide.setSelectedShapes([hit.id]); return context.sync(); }
          }).catch(function () {}).then(function () { return res; });
        });
      }).then(function (res) {
        if (res === null) { ST.toast("Nothing was changed."); return; }
        if (missing) { ST.toast(R.missText(), true); return; }
        if (!res || !res.count) { ST.toast("Couldn't place the element.", true); return; }
        ST.toast((replaced ? "Replaced the shape with a " : "Placed a ") + kindName.toLowerCase() + ". Edit its text on the slide.");
        ST.refreshSelection();
      }, function (err) { ST.toast("Couldn't change the slide. " + (ST.errorText ? ST.errorText(err) : ""), true); });
    }, t.fromSelection ? "Select a shape to replace, or choose Size and type a width." : null);
  }
  // Reads the theme colours once, from a copy of the slide that is never swapped
  var themeRead = false;
  function loadTheme() {
    if (themeRead || !ST.inOffice || !R.supported() || !(ST.isOpen && ST.isOpen("elements"))) return;
    themeRead = true;
    PowerPoint.run(function (context) {
      return R.edit(context, function (doc, extra) { themeHex = readTheme(extra && extra.files); return { count: 0 }; });
    }).catch(function () { themeRead = false; }).then(renderColors);
  }

  // ---------- Panel ----------
  function seg(id, opts) { return '<div class="seg seg' + opts.length + '" id="' + id + '">' + opts.map(function (o) { return '<button data-v="' + o[0] + '"' + (o[2] ? ' title="' + o[2] + '"' : "") + ">" + o[1] + "</button>"; }).join("") + "</div>"; }
  function field(id, label, ph, area) {
    return '<label class="field">' + label + (area ? '<textarea id="' + id + '" rows="3" placeholder="' + ph + '"></textarea>' : '<input type="text" id="' + id + '" placeholder="' + ph + '">') + "</label>";
  }
  var FIELDS = {
    stat: '<div class="el-fields" data-kind="stat">' + '<div class="row">' + field("elStTitle", "Title", "e.g. Region name") + field("elStValue", "Value", "e.g. +5.0%") + "</div>" +
      '<div class="toggle"><span class="label-row">Arrow from the sign' + ST.info("A value starting with + gets an up arrow, one starting with - gets a down arrow, and 0 gets a flat arrow.", "About the arrow") + '</span><label class="switch"><input type="checkbox" id="elStArrow"><span></span></label></div>' +
      field("elStSub", "Summary", "One line about the change") + field("elStBullets", "Bullets, one per line", "First point\nSecond point", true) +
      '<div class="row">' + field("elStFootB", "Footer figure", "e.g. 120 units") + field("elStFootN", "Footer note", "e.g. (15%)") + "</div></div>",
    pointer: '<div class="el-fields" data-kind="pointer">' + field("elPtTitle", "Title", "e.g. Key insight") + field("elPtBullets", "Bullets, one per line", "First point\nSecond point", true) +
      '<div class="row">' + field("elPtLabel", "Footer label", "e.g. Action needed:") + field("elPtFoot", "Footer text", "e.g. Next step") + "</div></div>",
    bignum: '<div class="el-fields" data-kind="bignum">' + field("elBnNum", "Number", "e.g. 75%") + field("elBnLabel", "Label", "What it measures") + field("elBnNote", "Note", "Optional") + "</div>"
  };
  var MARKUP =
    '<div class="mini-chips el-kinds" id="elKinds"></div>' +
    '<div class="help" id="elKindTip"></div>' +
    '<div class="field"><span class="label-row">Colour' + ST.info("Theme colours, so the element follows your template. Lighter bands are tints of the colour you pick.", "About colour") + '</span><div class="el-colors" id="elColors"></div></div>' +
    FIELDS.stat + FIELDS.pointer + FIELDS.bignum +
    '<div class="group-label">Size and place</div>' +
    seg("elPlace", [["selection", "Replace shape", "Draw a shape, select it, and the element takes its size and place"], ["size", "Type a size", "A width in g, col or points, at the grid area's top left"]]) +
    '<div class="row el-size"><label class="field"><span class="label-row">Width' + ST.info("Points or formulas with the grid names: <code>3col</code>, <code>2col+gap</code>, <code>8g</code> or <code>240</code>.", "About width") + '</span><input type="text" id="elW" placeholder="e.g. 3col"></label>' +
      '<label class="field">Height<input type="text" id="elH" placeholder="Fits the text"></label></div>' +
    '<div class="help el-place-help" id="elPlaceHelp"></div>' +
    '<button class="btn primary" id="elPlaceBtn">Place element</button>';

  function renderKinds() {
    var box = $("elKinds");
    box.innerHTML = "";
    KINDS.forEach(function (k) {
      var b = document.createElement("button");
      b.className = "mini-chip" + (ST.state.elKind === k.id ? " on" : "");
      b.textContent = k.name; b.title = k.tip; b.setAttribute("data-kind", k.id);
      b.addEventListener("click", function () { ST.state.elKind = k.id; ST.save(); render(); });
      box.appendChild(b);
    });
    var cur = KINDS.filter(function (k) { return k.id === ST.state.elKind; })[0] || KINDS[0];
    $("elKindTip").textContent = cur.tip;
    Array.prototype.forEach.call(document.querySelectorAll(".el-fields"), function (f) { f.hidden = f.getAttribute("data-kind") !== cur.id; });
  }
  function renderColors() {
    var box = $("elColors");
    if (!box) return;
    box.innerHTML = "";
    COLORS.forEach(function (c) {
      var b = document.createElement("button");
      b.className = "el-color" + (ST.state.elColor === c[0] ? " on" : "");
      b.title = c[1]; b.setAttribute("aria-label", c[1]); b.setAttribute("data-slot", c[0]);
      b.innerHTML = '<span style="background:' + (themeHex[c[0]] || "var(--field-line)") + '"></span>';
      b.addEventListener("click", function () { ST.state.elColor = c[0]; ST.save(); renderColors(); });
      box.appendChild(b);
    });
  }
  function renderPlace() {
    var size = ST.state.elPlace === "size";
    Array.prototype.forEach.call($("elPlace").children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === ST.state.elPlace); });
    document.querySelector(".el-size").hidden = !size;
    var help = $("elPlaceHelp");
    if (!size) { help.textContent = "Draw a shape where the element should go, select it, then click Place element."; help.classList.remove("bad"); return; }
    var r = targetBox(ST.state);
    help.textContent = r.bad ? r.bad : "Width " + ST.nice(r.box.w / PT) + " in (" + ST.gText(r.box.w / PT) + ")" + (r.box.h ? ", height " + ST.nice(r.box.h / PT) + " in" : ", height fits the text") + ", at the grid area's top left.";
    help.classList.toggle("bad", !!r.bad);
  }
  function render() { renderKinds(); renderColors(); renderPlace(); $("elStArrow").checked = ST.state.elStArrow !== false; }
  function build(body) {
    body.appendChild(ST.html(MARKUP));
    ["elStTitle", "elStValue", "elStSub", "elStBullets", "elStFootB", "elStFootN", "elPtTitle", "elPtBullets", "elPtLabel", "elPtFoot", "elBnNum", "elBnLabel", "elBnNote", "elW", "elH"].forEach(function (k) {
      ST.bindField($(k), k, k === "elW" || k === "elH" ? { onInput: renderPlace } : null);
    });
    $("elStArrow").addEventListener("change", function () { ST.state.elStArrow = $("elStArrow").checked; ST.save(); });
    Array.prototype.forEach.call($("elPlace").children, function (b) {
      b.addEventListener("click", function (ev) { ev.preventDefault(); ST.state.elPlace = b.getAttribute("data-v"); ST.save(); renderPlace(); });
    });
    $("elPlaceBtn").addEventListener("click", function () { place(null); });
    render();
    ST.on("reset", render);
    ST.on("change", renderPlace);
    ST.on("layout", loadTheme);
    ST.on("ready", loadTheme);
  }

  ST.registerTool({
    id: "elements",
    title: "Elements",
    icon: ICON,
    levels: [1],
    defaultLevel: 1,
    defaults: defaults,
    tip: { what: "Drops in prebuilt design elements in your theme colours: stat cards, pointer cards and big numbers.", steps: ["Pick an element and fill in its text.", "Draw a shape where it should go and select it, or type a width.", "Click Place element, then edit the text on the slide."] },
    build: build
  });

  ST._elements = { elementXml: elementXml, arrowFor: arrowFor, bullets: bullets, readTheme: readTheme, targetBox: targetBox, place: place, KINDS: KINDS };
})();
