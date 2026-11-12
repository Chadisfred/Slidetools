/* Character styles: sample text, save named styles in the presentation, apply them,
   and push edits to every use with Update all.

   How text is linked to a style
   PowerPoint add-ins can't attach data to individual characters, so each style's ID
   (1 to 63) is hidden in the lowest two bits of each colour channel. That shifts the
   colour by at most 3 out of 255 per channel, which can't be seen. Text counts as
   linked when its colour carries a style's ID and its shape is tagged with that ID.
   Text that has lost its code (usually recoloured from the ribbon) but still matches
   a style's font settings is offered as a lookalike during Update all.
   Only normal slides are read or changed. Masters and layouts are never touched. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;
  var STORE_KEY = "SLIDETOOLS_STYLES";   // presentation tags, split into parts
  var SHAPE_TAG = "SLIDETOOLS_USES";     // shape tag: comma separated style IDs
  var PART_SIZE = 2000;
  var MAX_ID = 63;
  var BATCH = 300;
  var TEXT_TYPES = ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"];

  var ICON = '<svg viewBox="0 0 18 18" fill="currentColor"><path d="M2.2 15L6.3 3.5h1.6L12 15h-1.9l-1-3H5l-1 3H2.2zm3.4-4.6h3L7.1 6 5.6 10.4z"/><rect x="12.6" y="9" width="3.6" height="6" rx="0.8" opacity="0.55"/></svg>';
  var EDIT_ICON = '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M10.5 2.5l3 3L6 13H3v-3l7.5-7.5z"/></svg>';

  var styles = [];        // [{ id, name, font, applied }]
  var stylesLoaded = false;
  var editing = null;     // id being edited, or "new"
  var draft = null;       // sampled or new style waiting for a name
  var review = null;      // Update all results waiting for confirmation
  var busy = false;

  // ---------- Font properties ----------
  var BASIC_PROPS = ["name", "size", "bold", "italic", "underline"];
  var EXTRA_PROPS = ["allCaps", "smallCaps", "strikethrough", "superscript", "subscript"];
  function props() { return ST.supports("1.8") || !ST.inOffice ? BASIC_PROPS.concat(EXTRA_PROPS) : BASIC_PROPS.slice(); }
  function loadList() { return props().concat(["color"]).join(","); }

  function readFont(font) {
    var f = { color: font.color ? String(font.color).toUpperCase() : null };
    props().forEach(function (p) { f[p] = font[p] === undefined ? null : font[p]; });
    return f;
  }
  function same(a, b) {
    if (a === null || b === null || a === undefined || b === undefined) return false;
    if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) < 0.05;
    return a === b;
  }
  function uniform(f) {
    if (!f.color) return false;
    return props().every(function (p) { return f[p] !== null; });
  }

  // ---------- Colour code ----------
  function rgb(hex) { hex = String(hex || "#000000").replace("#", ""); return [0, 2, 4].map(function (i) { return parseInt(hex.substr(i, 2), 16) || 0; }); }
  function hex(c) { return "#" + c.map(function (v) { return ("0" + v.toString(16)).slice(-2); }).join("").toUpperCase(); }
  function encode(color, id) {
    var c = rgb(color);
    return hex([(c[0] & ~3) | ((id >> 4) & 3), (c[1] & ~3) | ((id >> 2) & 3), (c[2] & ~3) | (id & 3)]);
  }
  function decodeId(color) { var c = rgb(color); return ((c[0] & 3) << 4) | ((c[1] & 3) << 2) | (c[2] & 3); }
  function base(color) { return hex(rgb(color).map(function (v) { return v & ~3; })); }
  ST.styleCode = { encode: encode, decodeId: decodeId, base: base }; // exposed for testing

  // ---------- Font weights ----------
  // Fonts whose name already carries a weight (Segoe UI Semibold, Arial Black, Roboto Light)
  // shouldn't also get PowerPoint's bold, which thickens them artificially.
  var WEIGHT_RE = /\b(thin|hairline|(extra|ultra|semi)[ -]?light|light|book|medium|(demi|semi|extra|ultra)[ -]?bold|bold|black|heavy)\b/i;
  function isWeighted(name) { return !!name && WEIGHT_RE.test(String(name)); }

  // ---------- Font list ----------
  // Fonts used in this deck (collected while scanning), fonts in saved styles, and
  // installed fonts when the browser inside PowerPoint allows listing them.
  var FONT_CHECK_KEY = "slideToolsFontCheck";
  var FONT_CHECK_HOURS = 12;
  var deckFonts = [], installedFonts = [], deckFontsState = "idle", installedState = null;
  function rememberedFontCheck() {
    if (installedState) return installedState;
    try {
      var s = sessionStorage.getItem(FONT_CHECK_KEY);
      if (s) return s;
      var saved = JSON.parse(localStorage.getItem(FONT_CHECK_KEY) || "null");
      if (saved && Date.now() - saved.at < FONT_CHECK_HOURS * 3600 * 1000) return saved.result;
    } catch (e) { /* storage unavailable */ }
    return null;
  }
  function rememberFontCheck(result) {
    installedState = result;
    try {
      sessionStorage.setItem(FONT_CHECK_KEY, result);
      localStorage.setItem(FONT_CHECK_KEY, JSON.stringify({ result: result, at: Date.now() }));
    } catch (e) { /* storage unavailable */ }
  }
  function addDeckFonts(names) {
    names.forEach(function (n) { if (n && deckFonts.indexOf(n) < 0) deckFonts.push(n); });
    deckFonts.sort(function (a, b) { return a.localeCompare(b); });
    updateFontList();
  }
  // PowerPoint names non standard weights by their full name (Segoe UI Semibold),
  // and Regular, Bold and Italic by the family name.
  function pptName(font) {
    var style = String(font.style || "").toLowerCase().replace(/\s+/g, " ").trim();
    if (["regular", "normal", "bold", "italic", "oblique", "bold italic", "bold oblique"].indexOf(style) >= 0) return font.family;
    return String(font.fullName || font.family).replace(/ (Italic|Oblique)$/i, "");
  }
  function tryInstalledFonts() {
    if (installedFonts.length) return;
    var state = rememberedFontCheck();
    if (state === "denied" || state === "unsupported" || state === "failed") return;
    if (typeof window.queryLocalFonts !== "function") { rememberFontCheck("unsupported"); return; }
    // The browser only allows the request straight after a click or key press.
    if (navigator.userActivation && !navigator.userActivation.isActive) return;
    var go = function () {
      window.queryLocalFonts().then(function (fonts) {
        var names = [];
        fonts.forEach(function (f) { var n = pptName(f); if (n && names.indexOf(n) < 0) names.push(n); });
        installedFonts = names.sort(function (a, b) { return a.localeCompare(b); });
        rememberFontCheck(installedFonts.length ? "ok" : "failed");
        updateFontList();
      }).catch(function (e) {
        rememberFontCheck(e && e.name === "NotAllowedError" ? "denied" : "failed");
      });
    };
    if (navigator.permissions && navigator.permissions.query) {
      navigator.permissions.query({ name: "local-fonts" }).then(function (p) {
        if (p.state === "denied") rememberFontCheck("denied"); else go();
      }, go);
    } else go();
  }
  function collectDeckFonts() {
    if (!ST.inOffice || !ST.supports("1.5") || deckFontsState !== "idle") return;
    deckFontsState = "busy";
    PowerPoint.run(function (context) { return scanDeck(context, ["name"]); }).then(function (r) {
      deckFontsState = "done";
      addDeckFonts(r.segs.map(function (x) { return x.font.name; }));
    }, function () { deckFontsState = "idle"; });
  }
  function updateFontList() {
    var dl = $("fontList");
    if (!dl) return;
    var seen = {}, html = "";
    var add = function (n, label) {
      if (!n || seen[n.toLowerCase()]) return;
      seen[n.toLowerCase()] = 1;
      html += '<option value="' + esc(n) + '"' + (label ? ' label="' + label + '"' : "") + "></option>";
    };
    deckFonts.forEach(function (n) { add(n, "In this presentation"); });
    styles.forEach(function (st) { add(st.font.name, "In a style"); });
    installedFonts.forEach(function (n) { add(n, ""); });
    dl.innerHTML = html;
  }

  function byId(id) { return styles.filter(function (s) { return s.id === id; })[0]; }
  function nextId() { for (var i = 1; i <= MAX_ID; i++) if (!byId(i)) return i; return 0; }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function pending(s) { return JSON.stringify(s.font) !== JSON.stringify(s.applied); }

  // ---------- Saving in the presentation ----------
  function loadStyles() {
    if (!ST.inOffice || !ST.supports("1.3")) { stylesLoaded = true; render(); return; }
    PowerPoint.run(function (context) {
      var tags = context.presentation.tags;
      var count = tags.getItemOrNullObject(STORE_KEY + "_N");
      count.load("value");
      return context.sync().then(function () {
        var n = count.isNullObject ? 0 : parseInt(count.value, 10) || 0;
        var parts = [];
        for (var i = 0; i < n; i++) { var t = tags.getItemOrNullObject(STORE_KEY + "_" + i); t.load("value"); parts.push(t); }
        return context.sync().then(function () {
          var text = parts.map(function (t) { return t.isNullObject ? "" : t.value; }).join("");
          styles = text ? JSON.parse(text) : [];
        });
      });
    }).catch(function () {
      styles = [];
      ST.toast("Couldn't read the styles saved in this presentation.", true);
    }).then(function () { stylesLoaded = true; render(); refreshStatus(); ST.emit("textstyles"); });
  }

  function saveStyles() {
    render();
    ST.emit("textstyles");
    if (!ST.inOffice || !ST.supports("1.3")) return Promise.resolve();
    var text = JSON.stringify(styles), parts = [];
    for (var i = 0; i < text.length; i += PART_SIZE) parts.push(text.slice(i, i + PART_SIZE));
    return PowerPoint.run(function (context) {
      var tags = context.presentation.tags;
      var old = tags.getItemOrNullObject(STORE_KEY + "_N");
      old.load("value");
      return context.sync().then(function () {
        var oldN = old.isNullObject ? 0 : parseInt(old.value, 10) || 0;
        parts.forEach(function (p, i) { tags.add(STORE_KEY + "_" + i, p); });
        for (var j = parts.length; j < oldN; j++) tags.delete(STORE_KEY + "_" + j);
        tags.add(STORE_KEY + "_N", String(parts.length));
        return context.sync();
      });
    }).catch(function (e) { ST.toast("Couldn't save styles: " + (e && e.message ? e.message : e), true); });
  }

  // ---------- Writing to text ----------
  // Sets the style's settings on a text range. keepColor keeps the text's own colour
  // (re-coded with the style ID). current is the range's font as read, used to keep overrides.
  function writeFont(range, style, opts) {
    opts = opts || {};
    var f = style.font, old = style.applied, cur = opts.current;
    props().forEach(function (p) {
      if (f[p] === null || f[p] === undefined) return;
      if (opts.keepOverrides && cur && old[p] !== null && old[p] !== undefined && !same(cur[p], old[p]) && !same(cur[p], f[p])) return;
      range.font[p] = f[p];
    });
    var color;
    if (opts.keepColor && cur && cur.color) color = encode(cur.color, style.id);
    else if (opts.keepOverrides && cur && cur.color && base(cur.color) !== base(encode(old.color, style.id)) && base(cur.color) !== base(encode(f.color, style.id))) color = encode(cur.color, style.id);
    else color = encode(f.color, style.id);
    range.font.color = color;
    // A weighted font with Bold on Leave: switch bold off so leftover bold doesn't stack.
    if (isWeighted(f.name) && (f.bold === null || f.bold === undefined) && (!opts.keepOverrides || !cur || !same(cur.name, f.name))) {
      range.font.bold = false;
    }
  }
  // Shape tag format: "3,7:00A651" means styles 3 and 7 are used here, and style 7
  // also appears with a kept colour whose coded base is #00A651.
  function parseTag(value) {
    var map = {};
    String(value || "").split(",").filter(Boolean).forEach(function (part) {
      var bits = part.split(":"), id = parseInt(bits[0], 10);
      if (!(id > 0)) return;
      map[id] = map[id] || [];
      bits.slice(1).forEach(function (b) { if (/^[0-9A-F]{6}$/i.test(b)) map[id].push("#" + b.toUpperCase()); });
    });
    return map;
  }
  function tagValue(map) {
    return Object.keys(map).map(Number).sort(function (a, b) { return a - b; }).map(function (id) {
      var uniq = map[id].filter(function (b, i, arr) { return arr.indexOf(b) === i; });
      return [id].concat(uniq.map(function (b) { return b.replace("#", ""); })).join(":");
    }).join(",");
  }
  function tagWith(value, id, keptColor) {
    var map = parseTag(value);
    map[id] = map[id] || [];
    if (keptColor) map[id].push(base(keptColor));
    return tagValue(map);
  }
  function addTag(shape, existing, id, keptColor) {
    var value = tagWith(existing, id, keptColor);
    shape.tags.add(SHAPE_TAG, value);
    return value;
  }

  // Gets the selected text, or the full text of each selected shape when no text is selected.
  function withSelectedText(task) {
    if (!ST.needOffice()) return;
    if (!ST.supports("1.5")) { ST.toast("This PowerPoint version can't work with selected text.", true); return; }
    return ST.run(function (context) {
      var range = context.presentation.getSelectedTextRangeOrNullObject();
      range.load("length,text");
      return context.sync().then(function () {
        if (!range.isNullObject && range.length > 0) {
          var shape = range.getParentTextFrame().getParentShape();
          var tag = shape.tags.getItemOrNullObject(SHAPE_TAG);
          tag.load("value");
          return context.sync().then(function () {
            return task(context, [{ range: range, shape: shape, tag: tag.isNullObject ? "" : tag.value }]);
          });
        }
        var shapes = context.presentation.getSelectedShapes();
        shapes.load("items/type,items/id");
        return context.sync().then(function () {
          var items = shapes.items.filter(function (s) { return TEXT_TYPES.indexOf(s.type) >= 0; });
          if (!items.length) { ST.toast("Select some text, or a text box.", true); return; }
          var list = items.map(function (s) {
            var r = s.textFrame.textRange, tag = s.tags.getItemOrNullObject(SHAPE_TAG);
            r.load("length,text"); tag.load("value");
            return { range: r, shape: s, tagObj: tag };
          });
          return context.sync().then(function () {
            list = list.filter(function (x) { return x.range.length > 0; });
            list.forEach(function (x) { x.tag = x.tagObj.isNullObject ? "" : x.tagObj.value; });
            if (!list.length) { ST.toast("The selected shapes have no text.", true); return; }
            return task(context, list);
          });
        });
      });
    });
  }

  // ---------- Sample, apply, relink, unlink ----------
  function sample() {
    withSelectedText(function (context, list) {
      var first = list[0].range;
      first.font.load(loadList());
      var firstChar = first.getSubstring(0, 1);
      firstChar.font.load("color");
      return context.sync().then(function () {
        var f = readFont(first.font);
        if (!f.color) f.color = firstChar.font.color ? String(firstChar.font.color).toUpperCase() : "#000000";
        var linked = decodeId(f.color);
        f.color = byId(linked) ? base(f.color) : f.color;
        var left = props().filter(function (p) { return f[p] === null; });
        var boldNote = false;
        if (isWeighted(f.name) && f.bold !== false) { f.bold = false; boldNote = true; }
        draft = { font: f, link: true, boldNote: boldNote };
        editing = "new";
        render();
        var name = $("styleName");
        if (name) { name.value = ""; name.focus(); }
        if (left.length) ST.toast("Mixed settings left out: " + left.join(", ") + ".");
      });
    });
  }

  function newStyle() {
    draft = { font: { name: "Segoe UI", size: 14, color: "#1D1D1F", bold: false, italic: false, underline: "None",
                      allCaps: null, smallCaps: null, strikethrough: null, superscript: null, subscript: null }, link: false };
    editing = "new";
    render();
    var name = $("styleName");
    if (name) name.focus();
  }

  function apply(style, opts) {
    withSelectedText(function (context, list) {
      list.forEach(function (x) {
        if (opts && opts.keepColor) x.range.font.load(loadList());
      });
      return context.sync().then(function () {
        list.forEach(function (x) {
          var cur = opts && opts.keepColor ? readFont(x.range.font) : null;
          writeFont(x.range, style, { keepColor: opts && opts.keepColor, current: cur });
          addTag(x.shape, x.tag, style.id, cur && cur.color && base(cur.color) !== base(style.font.color) ? cur.color : null);
        });
        return context.sync();
      }).then(function () {
        ST.toast((opts && opts.message) || "Applied " + style.name + ".");
        refreshStatus();
      });
    });
  }

  function unlink() {
    withSelectedText(function (context, list) {
      list.forEach(function (x) { x.range.font.load("color"); });
      var chars = list.map(function (x) { var c = x.range.getSubstring(0, 1); c.font.load("color"); return c; });
      return context.sync().then(function () {
        list.forEach(function (x, i) {
          var c = x.range.font.color || chars[i].font.color;
          if (c) x.range.font.color = encode(c, 0);
        });
        return context.sync();
      }).then(function () { ST.toast("Unlinked. The text keeps its look but no longer follows a style."); refreshStatus(); });
    });
  }

  // ---------- Selection status ----------
  var statusTimer = null;
  function refreshStatus() {
    var el = $("styleStatus");
    if (!el) return;
    if (!ST.inOffice || !ST.supports("1.5") || !ST.isOpen("styles") || ST.levelOf("styles") < 2) { el.hidden = true; return; }
    clearTimeout(statusTimer);
    statusTimer = setTimeout(readStatus, 150);
  }
  function readStatus() {
    var el = $("styleStatus");
    PowerPoint.run(function (context) {
      var range = context.presentation.getSelectedTextRangeOrNullObject();
      range.load("length");
      return context.sync().then(function () {
        var target = range, shape;
        if (range.isNullObject || !range.length) {
          var shapes = context.presentation.getSelectedShapes();
          shapes.load("items/type");
          return context.sync().then(function () {
            var s = shapes.items.filter(function (x) { return TEXT_TYPES.indexOf(x.type) >= 0; })[0];
            if (!s) return null;
            return finish(s.textFrame.textRange, s);
          });
        }
        shape = range.getParentTextFrame().getParentShape();
        return finish(target, shape);
      });
      function finish(r, s) {
        r.font.load(loadList());
        var tag = s.tags.getItemOrNullObject(SHAPE_TAG);
        tag.load("value");
        return context.sync().then(function () { return { font: readFont(r.font), tags: parseTag(tag.isNullObject ? "" : tag.value) }; });
      }
    }).then(function (info) { showStatus(info); }, function () { el.hidden = true; });
  }
  function showStatus(info) {
    var el = $("styleStatus");
    el.hidden = false;
    el.innerHTML = "";
    var text = document.createElement("span");
    el.appendChild(text);
    if (!info) { text.textContent = "Select text to see its style."; return; }
    var c = classify(info.font, info.tags);
    if (c.kind === "linked") {
      var changed = props().some(function (p) { return c.style.font[p] !== null && c.style.font[p] !== undefined && !same(info.font[p], c.style.font[p]); }) ||
        base(info.font.color) !== base(encode(c.style.font.color, c.style.id));
      text.innerHTML = "Uses <b></b>" + (changed ? " <span class=\"plus\" title=\"Has local changes\">+</span>" : "");
      text.querySelector("b").textContent = c.style.name;
      if (changed) addStatusButton(el, "Reset to style", function () { apply(c.style, { message: "Reset to " + c.style.name + "." }); });
      addStatusButton(el, "Unlink", unlink);
    } else if (c.kind === "lookalike") {
      text.innerHTML = "Looks like <b></b> (" + (c.coded ? "copied" : "colour changed") + ")";
      text.querySelector("b").textContent = c.style.name;
      addStatusButton(el, "Relink, keep colour", function () { apply(c.style, { keepColor: true, message: "Relinked to " + c.style.name + ", keeping its colour." }); });
    } else if (!info.font.color || props().some(function (p) { return info.font[p] === null; })) {
      text.textContent = "Mixed formatting selected.";
    } else {
      text.textContent = "No style.";
    }
  }
  function addStatusButton(el, label, fn) {
    var b = document.createElement("button");
    b.className = "link-btn";
    b.textContent = label;
    b.addEventListener("click", fn);
    el.appendChild(b);
  }

  // ---------- Classifying text ----------
  function lookalikeMatch(style, f) {
    var defs = [style.applied, style.font];
    return defs.some(function (d) {
      var keys = props().filter(function (p) { return d[p] !== null && d[p] !== undefined; });
      if (keys.indexOf("name") < 0 && keys.indexOf("size") < 0) return false;
      return keys.every(function (p) { return same(f[p], d[p]); });
    });
  }
  // tags is the parsed shape tag. Linked text must carry the style's code, sit in a
  // shape tagged with that style, and have a colour the panel set: the style's own
  // colour (current or last pushed) or a kept colour recorded in the tag.
  function classify(f, tags) {
    if (!f.color) return { kind: "none" };
    var id = decodeId(f.color), s = byId(id), b = base(f.color);
    if (s && tags[id] && (b === base(s.font.color) || b === base(s.applied.color) || tags[id].indexOf(b) >= 0)) return { kind: "linked", style: s };
    if (s && lookalikeMatch(s, f)) return { kind: "lookalike", style: s, coded: true };
    for (var i = 0; i < styles.length; i++) {
      if (lookalikeMatch(styles[i], f)) return { kind: "lookalike", style: styles[i], coded: false };
    }
    return { kind: "none" };
  }

  // ---------- Scanning the deck (normal slides only) ----------
  // only: optional list of font properties to read, e.g. ["name"] for a quick font list.
  function scanDeck(context, only) {
    var list = only ? only.join(",") : loadList();
    var isUniform = only ? function (f) { return only.every(function (p) { return f[p] !== null; }); } : uniform;
    var slides = context.presentation.slides;
    slides.load("items/id");
    var entries = [], groups = [], segs = [];
    var useNullFrame = ST.supports("1.10"), useGroups = ST.supports("1.8");

    function collect(sh, si, sid, path) {
      if (sh.type === "Group") { if (useGroups) groups.push({ shape: sh, si: si, sid: sid, path: path }); return; }
      if (useNullFrame ? sh.type !== "Table" : TEXT_TYPES.indexOf(sh.type) >= 0) entries.push({ shape: sh, si: si, sid: sid, path: path });
    }
    function expandGroups() {
      if (!groups.length) return Promise.resolve();
      var gs = groups; groups = [];
      gs.forEach(function (g) { g.shape.group.shapes.load("items/id,items/type"); });
      return context.sync().then(function () {
        gs.forEach(function (g) { g.shape.group.shapes.items.forEach(function (c) { collect(c, g.si, g.sid, g.path.concat([c.id])); }); });
        return expandGroups();
      });
    }

    return context.sync().then(function () {
      slides.items.forEach(function (sl) { sl.shapes.load("items/id,items/type"); });
      return context.sync();
    }).then(function () {
      slides.items.forEach(function (sl, si) { sl.shapes.items.forEach(function (sh) { collect(sh, si, sl.id, [sh.id]); }); });
      return expandGroups();
    }).then(function () {
      entries.forEach(function (e) {
        e.tf = useNullFrame ? e.shape.getTextFrameOrNullObject() : e.shape.textFrame;
        e.tf.load("hasText");
        e.tagObj = e.shape.tags.getItemOrNullObject(SHAPE_TAG);
        e.tagObj.load("value");
      });
      return context.sync();
    }).then(function () {
      entries = entries.filter(function (e) { return !e.tf.isNullObject && e.tf.hasText; });
      entries.forEach(function (e) { e.tags = parseTag(e.tagObj.isNullObject ? "" : e.tagObj.value); e.tagValue = e.tagObj.isNullObject ? "" : e.tagObj.value; e.tf.textRange.load("text"); });
      return context.sync();
    }).then(function () {
      // Split each text into paragraphs, then halve any piece with mixed formatting until each piece is uniform.
      var queue = [];
      entries.forEach(function (e) {
        var text = e.tf.textRange.text || "", re = /[^\r\n\v\u000b]+/g, m;
        e.text = text;
        while ((m = re.exec(text))) queue.push({ e: e, start: m.index, len: m[0].length });
      });
      function round() {
        if (!queue.length) return Promise.resolve();
        var batch = queue.splice(0, BATCH);
        batch.forEach(function (q) { q.rng = q.e.tf.textRange.getSubstring(q.start, q.len); q.rng.font.load(list); });
        return context.sync().then(function () {
          batch.forEach(function (q) {
            var f = only ? readOnly(q.rng.font, only) : readFont(q.rng.font);
            if (isUniform(f) || q.len === 1) segs.push({ e: q.e, start: q.start, len: q.len, font: f });
            else {
              var h = Math.floor(q.len / 2);
              queue.push({ e: q.e, start: q.start, len: h }, { e: q.e, start: q.start + h, len: q.len - h });
            }
          });
          return round();
        });
      }
      return round();
    }).then(function () {
      // Join neighbouring pieces that ended up with identical formatting.
      segs.sort(function (a, b) { return a.e === b.e ? a.start - b.start : entries.indexOf(a.e) - entries.indexOf(b.e); });
      var merged = [];
      segs.forEach(function (s) {
        var last = merged[merged.length - 1];
        if (last && last.e === s.e && last.start + last.len === s.start && JSON.stringify(last.font) === JSON.stringify(s.font)) last.len += s.len;
        else merged.push(s);
      });
      merged.forEach(function (s) { s.text = s.e.text.substr(s.start, s.len); });
      return { segs: merged, slideCount: slides.items.length };
    });
  }

  function readOnly(font, only) {
    var f = {};
    only.forEach(function (p) { f[p] = font[p] === undefined ? null : font[p]; });
    return f;
  }

  function getShape(context, sid, path) {
    var shape = context.presentation.slides.getItem(sid).shapes.getItem(path[0]);
    for (var i = 1; i < path.length; i++) shape = shape.group.shapes.getItem(path[i]);
    return shape;
  }

  // ---------- Update all ----------
  function updateAll() {
    if (!ST.needOffice()) return;
    if (!ST.supports("1.5")) { ST.toast("This PowerPoint version can't update styles.", true); return; }
    if (!styles.length) { ST.toast("Save a style first.", true); return; }
    if (busy) return;
    busy = true;
    review = { scanning: true };
    render();
    PowerPoint.run(function (context) {
      return scanDeck(context);
    }).then(function (result) {
      var groupsById = {};
      styles.forEach(function (s) { groupsById[s.id] = { style: s, linked: [], look: [], choice: "leave" }; });
      addDeckFonts(result.segs.map(function (x) { return x.font.name; }));
      if (deckFontsState === "idle") deckFontsState = "done";
      result.segs.forEach(function (seg) {
        if (!/\S/.test(seg.text)) return;
        var c = classify(seg.font, seg.e.tags);
        if (c.kind === "linked") groupsById[c.style.id].linked.push(seg);
        else if (c.kind === "lookalike") { seg.choice = null; groupsById[c.style.id].look.push(seg); }
      });
      var list = styles.map(function (s) { return groupsById[s.id]; }).filter(function (g) {
        return (g.linked.length && pending(g.style)) || g.look.length;
      });
      busy = false;
      if (!list.length) {
        review = null;
        styles.forEach(function (s) { s.applied = clone(s.font); });
        saveStyles();
        ST.toast("Everything is up to date. No lookalikes found on " + ST.plural(result.slideCount, "slide") + ".");
        return;
      }
      review = { groups: list, slideCount: result.slideCount };
      render();
    }).catch(function (e) {
      busy = false; review = null; render();
      ST.toast("Couldn't scan the slides: " + (e && e.message ? e.message : e), true);
    });
  }

  function confirmUpdate() {
    if (!review || !review.groups || busy) return;
    busy = true;
    var jobs = [];
    review.groups.forEach(function (g) {
      if (pending(g.style)) g.linked.forEach(function (seg) { jobs.push({ seg: seg, style: g.style, mode: "linked" }); });
      g.look.forEach(function (seg) {
        var choice = seg.choice || g.choice;
        if (choice !== "leave") jobs.push({ seg: seg, style: g.style, mode: choice });
      });
    });
    var changed = jobs.length;
    PowerPoint.run(function (context) {
      var shapes = {}, tagsToAdd = {};
      var i = 0;
      function next() {
        if (i >= jobs.length) {
          Object.keys(tagsToAdd).forEach(function (k) {
            var t = tagsToAdd[k], value = t.e.tagValue;
            t.adds.forEach(function (a) { value = tagWith(value, a.id, a.kept); });
            t.shape.tags.add(SHAPE_TAG, value);
          });
          return context.sync();
        }
        jobs.slice(i, i + BATCH).forEach(function (j) {
          var e = j.seg.e, key = e.sid + "/" + e.path.join("/");
          var shape = shapes[key] || (shapes[key] = getShape(context, e.sid, e.path));
          var range = shape.textFrame.textRange.getSubstring(j.seg.start, j.seg.len);
          if (j.mode === "linked") writeFont(range, j.style, { keepOverrides: true, current: j.seg.font });
          else {
            writeFont(range, j.style, { keepColor: j.mode === "keep", current: j.seg.font });
            var t = tagsToAdd[key] || (tagsToAdd[key] = { shape: shape, e: e, adds: [] });
            var kept = j.mode === "keep" && base(j.seg.font.color) !== base(j.style.font.color) ? j.seg.font.color : null;
            t.adds.push({ id: j.style.id, kept: kept });
          }
        });
        i += BATCH;
        return context.sync().then(next);
      }
      return next();
    }).then(function () {
      styles.forEach(function (s) { s.applied = clone(s.font); });
      busy = false; review = null;
      return saveStyles().then(function () {
        ST.toast(changed ? "Updated " + ST.plural(changed, "piece") + " of text." : "Styles saved. Nothing on the slides needed changing.");
        refreshStatus();
      });
    }).catch(function (e) {
      busy = false;
      ST.toast("Update stopped: " + (e && e.message ? e.message : e) + ". Some text may already be updated, use Undo if needed.", true);
      render();
    });
  }

  function showSeg(seg) {
    ST.run(function (context) {
      context.presentation.setSelectedSlides([seg.e.sid]);
      var shape = getShape(context, seg.e.sid, seg.e.path);
      shape.textFrame.textRange.getSubstring(seg.start, seg.len).setSelected();
      return context.sync();
    });
  }

  // ---------- Editing ----------
  var TRI = [["", "Leave"], ["1", "On"], ["0", "Off"]];
  function triSelect(id, value) {
    var v = value === null || value === undefined ? "" : (value ? "1" : "0");
    return '<select id="' + id + '">' + TRI.map(function (o) { return '<option value="' + o[0] + '"' + (o[0] === v ? " selected" : "") + ">" + o[1] + "</option>"; }).join("") + "</select>";
  }
  function optSelect(id, options, value) {
    return '<select id="' + id + '">' + options.map(function (o) { return '<option value="' + o[0] + '"' + (o[0] === value ? " selected" : "") + ">" + o[1] + "</option>"; }).join("") + "</select>";
  }
  function esc(t) { return String(t === null || t === undefined ? "" : t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

  function editorMarkup(name, f, isNew) {
    var caps = f.allCaps ? "all" : f.smallCaps ? "small" : (f.allCaps === false && f.smallCaps === false ? "none" : "");
    var pos = f.superscript ? "super" : f.subscript ? "sub" : (f.superscript === false && f.subscript === false ? "normal" : "");
    var underline = f.underline === null || f.underline === undefined ? "" : (f.underline && f.underline !== "None" ? "1" : "0");
    var extra = props().indexOf("allCaps") >= 0;
    return '<div class="style-editor">' +
      '<label class="field">Style name<input type="text" id="styleName" maxlength="32" value="' + esc(name) + '" placeholder="e.g. Header"></label>' +
      '<div class="row">' +
        '<label class="field">Font<input type="text" id="sFont" list="fontList" autocomplete="off" value="' + esc(f.name) + '" placeholder="Leave"></label>' +
        '<label class="field">Size (pt)<input type="text" id="sSize" inputmode="decimal" value="' + esc(f.size) + '" placeholder="Leave"></label>' +
      '</div>' +
      '<label class="field">Colour<div class="color-field"><input type="color" id="sColorPick" value="' + esc(base(f.color).toLowerCase()) + '"><input type="text" id="sColor" value="' + esc(f.color) + '"></div></label>' +
      '<div class="prop-grid">' +
        '<label>Bold' + triSelect("sBold", f.bold) + '</label>' +
        '<label>Italic' + triSelect("sItalic", f.italic) + '</label>' +
        '<label>Underline' + optSelect("sUnderline", TRI, underline) + '</label>' +
        (extra ? '<label>Strikethrough' + triSelect("sStrike", f.strikethrough) + '</label>' +
          '<label>Caps' + optSelect("sCaps", [["", "Leave"], ["none", "Normal"], ["all", "All caps"], ["small", "Small caps"]], caps) + '</label>' +
          '<label>Position' + optSelect("sPos", [["", "Leave"], ["normal", "Normal"], ["super", "Superscript"], ["sub", "Subscript"]], pos) + '</label>' : "") +
      '</div>' +
      '<datalist id="fontList"></datalist>' +
      '<div class="weight-note" id="sWarn" hidden></div>' +
      (isNew && draft && draft.link ? '<div class="muted-note">Saving links the sampled text to this style.</div>' : "") +
      '<div class="btn-row">' +
        '<button class="btn" id="sCancel">Cancel</button>' +
        '<button class="btn primary" id="sSave">' + (isNew ? "Save style" : "Save changes") + '</button>' +
      '</div>' +
      (isNew ? "" : '<button class="btn small danger" id="sDelete">Delete style</button>') +
    '</div>';
  }

  function readEditor() {
    var f = {};
    var name = $("sFont").value.trim(); f.name = name || null;
    var size = ST.evaluate($("sSize").value, null, null); f.size = size > 0 ? Math.round(size * 10) / 10 : null;
    var color = $("sColor").value.trim();
    if (!/^#?[0-9a-f]{6}$/i.test(color)) return { error: "Colour needs to be a 6 digit hex code like #0071E3." };
    f.color = ("#" + color.replace("#", "")).toUpperCase();
    var tri = function (id) { var v = $(id).value; return v === "" ? null : v === "1"; };
    f.bold = tri("sBold"); f.italic = tri("sItalic");
    var u = $("sUnderline").value; f.underline = u === "" ? null : (u === "1" ? "Single" : "None");
    if ($("sCaps")) {
      f.strikethrough = tri("sStrike");
      var caps = $("sCaps").value;
      f.allCaps = caps === "" ? null : caps === "all"; f.smallCaps = caps === "" ? null : caps === "small";
      var pos = $("sPos").value;
      f.superscript = pos === "" ? null : pos === "super"; f.subscript = pos === "" ? null : pos === "sub";
    } else {
      f.strikethrough = f.allCaps = f.smallCaps = f.superscript = f.subscript = null;
    }
    if (!f.name && !f.size) return { error: "A style needs at least a font or a size." };
    return { font: f, name: $("styleName").value.trim() };
  }

  function saveEditor() {
    var r = readEditor();
    if (r.error) { ST.toast(r.error, true); return; }
    if (!r.name) { ST.toast("Give the style a name.", true); $("styleName").focus(); return; }
    var clash = styles.filter(function (s) { return s.name.toLowerCase() === r.name.toLowerCase() && s.id !== editing; })[0];
    if (clash) { ST.toast("There's already a style called " + clash.name + ".", true); return; }
    if (editing === "new") {
      var id = nextId();
      if (!id) { ST.toast("This presentation already has the maximum of " + MAX_ID + " styles.", true); return; }
      var s = { id: id, name: r.name, font: r.font, applied: clone(r.font) };
      var link = draft && draft.link;
      styles.push(s);
      editing = null; draft = null;
      saveStyles().then(function () {
        if (link) apply(s, { message: "Saved " + s.name + " and linked the sampled text." });
        else ST.toast("Saved " + s.name + ". Select text and click it to apply.");
      });
    } else {
      var st = byId(editing);
      st.name = r.name; st.font = r.font;
      editing = null;
      saveStyles();
      ST.toast(pending(st) ? "Saved. Click Update all to change the text that uses " + st.name + "." : "Saved " + st.name + ".");
    }
  }

  function deleteStyle(id) {
    var s = byId(id);
    styles = styles.filter(function (x) { return x.id !== id; });
    editing = null;
    saveStyles();
    ST.toast("Deleted " + s.name + ". Text keeps its look.");
  }

  // ---------- Rendering ----------
  function summary(f) {
    var bits = [];
    if (f.name) bits.push(f.name);
    if (f.size) bits.push(f.size + " pt");
    if (f.bold) bits.push("bold");
    if (f.italic) bits.push("italic");
    if (f.underline && f.underline !== "None") bits.push("underline");
    if (f.allCaps) bits.push("all caps");
    if (f.smallCaps) bits.push("small caps");
    if (f.strikethrough) bits.push("strikethrough");
    if (f.superscript) bits.push("superscript");
    if (f.subscript) bits.push("subscript");
    bits.push(f.color);
    return bits.join(", ");
  }
  function sampleStyle(f) {
    var css = "color:" + f.color + ";";
    if (f.name) css += "font-family:'" + String(f.name).replace(/'/g, "") + "',sans-serif;";
    if (f.bold) css += "font-weight:700;";
    if (f.italic) css += "font-style:italic;";
    if (f.underline && f.underline !== "None") css += "text-decoration:underline;";
    if (f.strikethrough) css += "text-decoration:line-through;";
    return css;
  }

  function render() {
    var list = $("styleList"), form = $("styleForm"), rev = $("styleReview");
    if (!list) return;
    form.innerHTML = ""; rev.innerHTML = "";
    list.hidden = false; rev.hidden = true;
    $("styleActions").hidden = false;

    if (review) { renderReview(rev); list.hidden = true; $("styleActions").hidden = true; return; }

    if (editing === "new") {
      form.innerHTML = editorMarkup("", draft.font, true);
      wireEditor();
    }

    list.innerHTML = "";
    if (!stylesLoaded && ST.inOffice) { list.innerHTML = '<div class="muted-note">Reading styles from this presentation.</div>'; return; }
    if (!styles.length && editing !== "new") {
      list.innerHTML = '<div class="muted-note">No styles yet. Select some text and click Sample to make one from it.</div>';
    }
    styles.forEach(function (s) {
      var row = document.createElement("div");
      row.className = "style-row";
      row.innerHTML =
        '<button class="style-apply" title="Apply to the selected text">' +
          '<span class="aa"></span>' +
          '<span class="style-text"><span class="style-name"></span><span class="style-sum" data-min="2"></span></span>' +
          (pending(s) ? '<span class="badge" title="Edited since the last Update all">Edited</span>' : "") +
        '</button>' +
        '<button class="icon-btn" title="Edit style" aria-label="Edit style">' + EDIT_ICON + '</button>';
      var aa = row.querySelector(".aa");
      aa.textContent = s.font.allCaps ? "AA" : "Aa";
      var c = rgb(s.font.color), light = 0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2] > 200;
      aa.setAttribute("style", sampleStyle(s.font) + (light ? "background:#3a3a3c;" : ""));
      row.querySelector(".style-name").textContent = s.name;
      row.querySelector(".style-sum").textContent = summary(s.font);
      row.querySelector(".style-apply").addEventListener("click", function () { apply(s); });
      row.querySelector(".icon-btn").addEventListener("click", function () { editing = editing === s.id ? null : s.id; render(); });
      list.appendChild(row);
      if (editing === s.id) {
        var holder = document.createElement("div");
        holder.innerHTML = editorMarkup(s.name, s.font, false);
        list.appendChild(holder);
        wireEditor(s.id);
      }
    });
    var n = styles.filter(pending).length;
    $("updateAll").textContent = n ? "Update all (" + n + " edited)" : "Update all";
  }

  function wireEditor(id) {
    var pick = $("sColorPick"), text = $("sColor");
    var fontBox = $("sFont"), boldBox = $("sBold"), warn = $("sWarn");
    var lastAutoFont = null;
    function checkWeight(fromFont) {
      var name = fontBox.value.trim(), bold = boldBox.value;
      var msg = "", warnLevel = false;
      if (fromFont && isWeighted(name) && bold !== "0" && name !== lastAutoFont) {
        boldBox.value = "0"; bold = "0"; lastAutoFont = name;
        msg = "Bold set to Off, because " + name + " is already a weighted font.";
      } else if (isWeighted(name) && bold === "1") {
        msg = "Double bold: " + name + " is already a weighted font, so Bold On makes PowerPoint thicken it artificially. It can look smeared and print badly.";
        warnLevel = true;
      } else if (!name && bold === "1") {
        msg = "Bold On without a font may give a fake bold on light or regular fonts. For a true weight, pick the font's own bold version, such as Segoe UI Semibold or Segoe UI Bold.";
        warnLevel = true;
      } else if (isWeighted(name) && bold === "0" && warn.getAttribute("data-auto") === "1") {
        return; // keep the earlier note visible
      }
      warn.hidden = !msg;
      warn.textContent = msg;
      warn.className = "weight-note" + (warnLevel ? " warn" : "");
      warn.setAttribute("data-auto", msg && !warnLevel ? "1" : "0");
    }
    fontBox.addEventListener("change", function () { checkWeight(true); });
    fontBox.addEventListener("blur", function () { checkWeight(true); });
    fontBox.addEventListener("input", function () { if (!fontBox.value.trim()) checkWeight(false); });
    fontBox.addEventListener("focus", tryInstalledFonts);
    fontBox.addEventListener("pointerdown", tryInstalledFonts);
    boldBox.addEventListener("change", function () { warn.setAttribute("data-auto", "0"); checkWeight(false); });
    updateFontList();
    collectDeckFonts();
    tryInstalledFonts();
    if (draft && draft.boldNote && editing === "new") {
      lastAutoFont = fontBox.value.trim();
      warn.hidden = false;
      warn.textContent = "Bold set to Off, because " + lastAutoFont + " is already a weighted font.";
      warn.setAttribute("data-auto", "1");
    } else {
      checkWeight(false);
    }
    pick.addEventListener("input", function () { text.value = pick.value.toUpperCase(); });
    text.addEventListener("input", function () { if (/^#?[0-9a-f]{6}$/i.test(text.value.trim())) pick.value = ("#" + text.value.trim().replace("#", "")).toLowerCase(); });
    $("sCancel").addEventListener("click", function () { editing = null; draft = null; render(); });
    $("sSave").addEventListener("click", saveEditor);
    $("styleName").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); saveEditor(); } });
    if (id) $("sDelete").addEventListener("click", function () { deleteStyle(id); });
  }

  var CHOICES = [["leave", "Leave them alone"], ["full", "Apply style and colour"], ["keep", "Apply style, keep their colour"]];
  function renderReview(el) {
    el.hidden = false;
    if (review.scanning) { el.innerHTML = '<div class="muted-note">Checking every slide. Masters and layouts are left alone.</div>'; return; }
    var html = '<div class="group-label">Update all, ' + ST.plural(review.slideCount, "slide") + ' checked</div>';
    el.innerHTML = html;
    review.groups.forEach(function (g, gi) {
      var box = document.createElement("div");
      box.className = "review-group";
      var slidesLinked = {};
      g.linked.forEach(function (s) { slidesLinked[s.e.si] = 1; });
      var lines = '<div class="review-title"></div>';
      if (g.linked.length) {
        lines += '<div class="review-line">' + (pending(g.style)
          ? ST.plural(g.linked.length, "linked piece") + " on " + ST.plural(Object.keys(slidesLinked).length, "slide") + " will update. Local changes are kept."
          : ST.plural(g.linked.length, "linked piece") + ", nothing to change.") + "</div>";
      }
      if (g.look.length) {
        lines += '<div class="review-line">' + ST.plural(g.look.length, "piece") + (g.look.length === 1 ? " looks like this style but isn't linked.</div>" : " look like this style but aren't linked.</div>") +
          '<label class="field">For those<select class="look-choice">' + CHOICES.map(function (c) { return '<option value="' + c[0] + '">' + c[1] + "</option>"; }).join("") + "</select></label>" +
          '<button class="link-btn review-toggle">Review each one</button><div class="review-items" hidden></div>';
      }
      box.innerHTML = lines;
      box.querySelector(".review-title").textContent = g.style.name;
      var sel = box.querySelector(".look-choice");
      if (sel) {
        sel.value = g.choice;
        sel.addEventListener("change", function () { g.choice = sel.value; });
        var items = box.querySelector(".review-items");
        box.querySelector(".review-toggle").addEventListener("click", function () {
          items.hidden = !items.hidden;
          if (!items.hidden && !items.children.length) renderItems(items, g);
        });
      }
      el.appendChild(box);
    });
    var actions = document.createElement("div");
    actions.className = "btn-row";
    actions.innerHTML = '<button class="btn" id="reviewCancel">Cancel</button><button class="btn primary" id="reviewGo">Update</button>';
    el.appendChild(actions);
    $("reviewCancel").addEventListener("click", function () { review = null; render(); });
    $("reviewGo").addEventListener("click", confirmUpdate);
  }
  function renderItems(el, g) {
    g.look.forEach(function (seg) {
      var row = document.createElement("div");
      row.className = "review-item";
      row.innerHTML = '<button class="link-btn item-show"></button><select class="item-choice"><option value="">Same as above</option>' +
        CHOICES.map(function (c) { return '<option value="' + c[0] + '">' + c[1] + "</option>"; }).join("") + "</select>";
      var snippet = seg.text.length > 34 ? seg.text.slice(0, 33) + "…" : seg.text;
      row.querySelector(".item-show").textContent = "Slide " + (seg.e.si + 1) + ": " + snippet;
      row.querySelector(".item-show").title = "Select this text on the slide";
      row.querySelector(".item-show").addEventListener("click", function () { showSeg(seg); });
      var sel = row.querySelector(".item-choice");
      sel.value = seg.choice || "";
      sel.addEventListener("change", function () { seg.choice = sel.value || null; });
      el.appendChild(row);
    });
  }

  // ---------- Build ----------
  var MARKUP =
    '<span class="info res-info" tabindex="0" role="button" aria-label="How styles work">i' +
      '<span class="tip" role="tooltip">Select text and click Sample to save its look as a character style. Click a style to apply it to selected text.<br><br>' +
      'Edit a style, then click Update all to change every piece of text that uses it. Settings you changed by hand on linked text are kept.<br><br>' +
      'The link is hidden in a tiny, invisible shift of the text colour. If you recolour styled text from the ribbon, Update all offers it as a lookalike, and you can choose to keep its colour.<br><br>' +
      'Styles are saved in this presentation. Only normal slides are changed, never masters or layouts.</span>' +
    '</span>' +
    '<div class="style-status" id="styleStatus" data-min="2" hidden></div>' +
    '<div id="styleForm"></div>' +
    '<div class="style-list" id="styleList"></div>' +
    '<div id="styleReview" hidden></div>' +
    '<div id="styleActions">' +
      '<div class="btn-row">' +
        '<button class="btn" id="sampleStyle" title="Make a style from the selected text">Sample</button>' +
        '<button class="btn" id="updateAll" title="Push style edits to every slide">Update all</button>' +
      '</div>' +
      '<button class="btn small new-style" id="newStyle" data-min="2">New blank style</button>' +
    '</div>';

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    $("sampleStyle").addEventListener("click", sample);
    $("newStyle").addEventListener("click", newStyle);
    $("updateAll").addEventListener("click", updateAll);
    ST.on("ready", loadStyles);
    ST.on("selection", refreshStatus);
    ST.on("layout", refreshStatus);
    render();
  }

  ST.registerTool({
    id: "styles",
    title: "Styles",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 2,
    defaults: {},
    build: build
  });

  // Used by the Shape styles tool to apply a character style to a shape's text.
  ST.textStyles = {
    list: function () { return styles; },
    byId: byId,
    shapeTag: SHAPE_TAG,
    fontProps: loadList,
    // Applies text style id to all text in shape. existingTag is the shape's current SLIDETOOLS_USES value.
    writeToShape: function (shape, id, existingTag) {
      var st = byId(id);
      if (!st) return false;
      writeFont(shape.textFrame.textRange, st);
      addTag(shape, existingTag, id);
      return true;
    },
    // Returns the text style a whole text uses, or null.
    detect: function (font, tagValue) {
      var c = classify(readFont(font), parseTag(tagValue));
      return c.kind === "linked" ? c.style.id : null;
    }
  };

  // Exposed for testing only
  ST._styles = { get: function () { return styles; }, set: function (s) { styles = s; stylesLoaded = true; render(); }, classify: classify, scanDeck: scanDeck, parseTag: parseTag, tagWith: tagWith,
    isWeighted: isWeighted, pptName: pptName, fonts: function () { return { deck: deckFonts, installed: installedFonts, state: rememberedFontCheck() }; } };
})();
