/* Colours tool: a palette saved in the presentation, plus the theme colours, applied to
   fill, outline or text in one click. Text recolouring keeps Styles links intact.
   Theme colours are only read, never changed, because they live on the slide master. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;
  var STORE_KEY = "SLIDETOOLS_COLOURS";
  var PART_SIZE = 2000;
  var SHAPE_TYPES = ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"];
  // PowerPoint's picker order and names.
  var THEME_SLOTS = ["Light1", "Dark1", "Light2", "Dark2", "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6"];
  var THEME_NAMES = { Light1: "Background 1", Dark1: "Text 1", Light2: "Background 2", Dark2: "Text 2", Accent1: "Accent 1", Accent2: "Accent 2", Accent3: "Accent 3", Accent4: "Accent 4", Accent5: "Accent 5", Accent6: "Accent 6" };
  var XML_SLOTS = { Light1: "lt1", Dark1: "dk1", Light2: "lt2", Dark2: "dk2", Accent1: "accent1", Accent2: "accent2", Accent3: "accent3", Accent4: "accent4", Accent5: "accent5", Accent6: "accent6" };
  var DML = "http://schemas.openxmlformats.org/drawingml/2006/main";

  var defaults = { colTarget: "fill", colTab: "theme" };
  ST.addDefaults(defaults);

  var ICON = '<svg viewBox="0 0 18 18"><circle cx="6" cy="6.5" r="4" fill="currentColor"/><circle cx="12" cy="6.5" r="4" fill="currentColor" opacity="0.55"/><circle cx="9" cy="11.8" r="4" fill="currentColor" opacity="0.3"/></svg>';

  var palette = [], theme = [], template = [], loaded = false;

  function norm(c) {
    c = String(c || "").trim();
    if (/^#?[0-9a-f]{6}$/i.test(c)) return ("#" + c.replace("#", "")).toUpperCase();
    try {
      var ctx = document.createElement("canvas").getContext("2d");
      ctx.fillStyle = "#000000"; ctx.fillStyle = c;
      return /^#[0-9a-f]{6}$/i.test(ctx.fillStyle) ? ctx.fillStyle.toUpperCase() : null;
    } catch (e) { return null; }
  }
  function textOn(hex) {
    var r = parseInt(hex.substr(1, 2), 16), g = parseInt(hex.substr(3, 2), 16), b = parseInt(hex.substr(5, 2), 16);
    return 0.299 * r + 0.587 * g + 0.114 * b > 150 ? "#1d1d1f" : "#ffffff";
  }

  // ---------- Tints and shades, the way Office works them out ----------
  function rgbToHsl(hex) {
    var r = parseInt(hex.substr(1, 2), 16) / 255, g = parseInt(hex.substr(3, 2), 16) / 255, b = parseInt(hex.substr(5, 2), 16) / 255;
    var mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2, h = 0, s = 0, d = mx - mn;
    if (d) {
      s = d / (1 - Math.abs(2 * l - 1));
      if (mx === r) h = ((g - b) / d) % 6; else if (mx === g) h = (b - r) / d + 2; else h = (r - g) / d + 4;
      h *= 60; if (h < 0) h += 360;
    }
    return [h, s, l];
  }
  function hslToHex(h, s, l) {
    var c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs((h / 60) % 2 - 1)), m = l - c / 2, r = 0, g = 0, b = 0;
    if (h < 60) { r = c; g = x; } else if (h < 120) { r = x; g = c; } else if (h < 180) { g = c; b = x; }
    else if (h < 240) { g = x; b = c; } else if (h < 300) { r = x; b = c; } else { r = c; b = x; }
    // Office rounds halves down, so 50% grey is #7F7F7F.
    return "#" + [r, g, b].map(function (v) { return ("0" + Math.round((v + m) * 255 - 1e-6).toString(16)).slice(-2); }).join("").toUpperCase();
  }
  // Five variants per theme colour, matching the rows under PowerPoint's theme colours.
  function variants(c) {
    var hsl = rgbToHsl(c.hex), l = hsl[2], steps;
    if (l >= 0.999) steps = [-5, -15, -25, -35, -50];
    else if (l <= 0.001) steps = [50, 35, 25, 15, 5];
    else if (l < 0.2) steps = [90, 75, 50, 25, 10];
    else if (l > 0.8) steps = [-10, -25, -50, -75, -90];
    else steps = [80, 60, 40, -25, -50];
    return steps.map(function (st) {
      var p = st / 100, nl = p > 0 ? l * (1 - p) + p : l * (1 + p);
      return { hex: hslToHex(hsl[0], hsl[1], Math.min(1, Math.max(0, nl))), name: c.name + ", " + (p > 0 ? "Lighter " : "Darker ") + Math.abs(st) + "%", theme: true };
    });
  }
  ST.colourVariants = variants; // exposed for testing

  // ---------- Template colours from the theme file ----------
  // Exports a copy of the first slide (never changes the deck), unzips it in the browser
  // and reads the theme's colour scheme and Custom Colors list.
  function b64ToBytes(b64) {
    var bin = atob(b64), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function inflate(bytes) {
    var ds = new DecompressionStream("deflate-raw");
    var stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Response(stream).arrayBuffer().then(function (buf) { return new Uint8Array(buf); });
  }
  // Minimal ZIP reader: returns the text of entries whose names pass the filter.
  function readZip(bytes, want) {
    var view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength), eocd = -1;
    for (var i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) { if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
    if (eocd < 0) return Promise.reject(new Error("not a zip"));
    var count = view.getUint16(eocd + 10, true), p = view.getUint32(eocd + 16, true), dec = new TextDecoder(), jobs = [];
    for (var n = 0; n < count; n++) {
      if (view.getUint32(p, true) !== 0x02014b50) break;
      var method = view.getUint16(p + 10, true), csize = view.getUint32(p + 20, true);
      var nameLen = view.getUint16(p + 28, true), extraLen = view.getUint16(p + 30, true), commentLen = view.getUint16(p + 32, true);
      var local = view.getUint32(p + 42, true), name = dec.decode(bytes.subarray(p + 46, p + 46 + nameLen));
      p += 46 + nameLen + extraLen + commentLen;
      if (!want(name)) continue;
      var start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
      var data = bytes.subarray(start, start + csize);
      jobs.push((method === 8 ? inflate(data) : Promise.resolve(data)).then((function (nm) {
        return function (raw) { return { name: nm, text: dec.decode(raw) }; };
      })(name)));
    }
    return Promise.all(jobs);
  }
  function colourOf(el) {
    if (!el) return null;
    var s = el.getElementsByTagNameNS(DML, "srgbClr")[0];
    if (s) return norm(s.getAttribute("val"));
    var y = el.getElementsByTagNameNS(DML, "sysClr")[0];
    return y ? norm(y.getAttribute("lastClr")) : null;
  }
  function parseTheme(xml) {
    var doc = new DOMParser().parseFromString(xml, "application/xml"), out = { scheme: [], custom: [] };
    var scheme = doc.getElementsByTagNameNS(DML, "clrScheme")[0];
    if (scheme) THEME_SLOTS.forEach(function (slot) {
      var hex = colourOf(scheme.getElementsByTagNameNS(DML, XML_SLOTS[slot])[0]);
      if (hex) out.scheme.push({ hex: hex, name: THEME_NAMES[slot], theme: true });
    });
    Array.prototype.forEach.call(doc.getElementsByTagNameNS(DML, "custClr"), function (el, i) {
      var hex = colourOf(el);
      if (hex) out.custom.push({ hex: hex, name: el.getAttribute("name") || "Custom colour " + (i + 1), template: true });
    });
    return out;
  }
  function readTemplate(context) {
    if (!ST.supports("1.8") || typeof DecompressionStream !== "function") return Promise.resolve(null);
    var slides = context.presentation.slides;
    slides.load("items/id");
    return context.sync().then(function () {
      if (!slides.items.length) return null;
      var data = slides.items[0].exportAsBase64();
      return context.sync().then(function () {
        return readZip(b64ToBytes(data.value), function (n) { return /^ppt\/theme\/theme\d+\.xml$/.test(n); });
      }).then(function (files) {
        var themes = files.map(function (f) { return parseTheme(f.text); });
        return themes.filter(function (t) { return t.custom.length; })[0] || themes.filter(function (t) { return t.scheme.length; })[0] || null;
      });
    }).catch(function () { return null; });
  }

  // Shared with the style editors: their colour pickers suggest these colours.
  function allColours() {
    var tints = [];
    theme.forEach(function (c) { tints = tints.concat(variants(c)); });
    return theme.concat(tints, template, palette);
  }
  ST.palette = { list: allColours };
  function updateDatalist() {
    var dl = document.getElementById("paletteList");
    if (!dl) { dl = document.createElement("datalist"); dl.id = "paletteList"; document.body.appendChild(dl); }
    var seen = {};
    dl.innerHTML = ST.palette.list().filter(function (c) { if (seen[c.hex]) return false; seen[c.hex] = 1; return true; })
      .map(function (c) { return '<option value="' + c.hex.toLowerCase() + '"></option>'; }).join("");
  }

  // ---------- Saving and loading ----------
  function load() {
    if (!ST.inOffice || !ST.supports("1.3")) { loaded = true; render(); return; }
    PowerPoint.run(function (context) {
      var themeResults = [];
      var step = Promise.resolve();
      if (ST.supports("1.10")) {
        step = step.then(function () {
          var slides = context.presentation.slides;
          slides.load("items/id");
          return context.sync().then(function () {
            if (!slides.items.length) return;
            var scheme = slides.items[0].themeColorScheme;
            THEME_SLOTS.forEach(function (slot) { themeResults.push({ slot: slot, r: scheme.getThemeColor(slot) }); });
            return context.sync();
          }).then(function () {
            theme = themeResults.map(function (t) { var hex = norm(t.r.value); return hex ? { hex: hex, name: THEME_NAMES[t.slot], theme: true } : null; }).filter(Boolean);
          }).catch(function () { theme = []; });
        });
      }
      step = step.then(function () {
        return readTemplate(context).then(function (t) {
          template = t ? t.custom : [];
          if (t && t.scheme.length && theme.length < THEME_SLOTS.length) theme = t.scheme;
        });
      });
      return step.then(function () {
        var tags = context.presentation.tags, count = tags.getItemOrNullObject(STORE_KEY + "_N");
        count.load("value");
        return context.sync().then(function () {
          var n = count.isNullObject ? 0 : parseInt(count.value, 10) || 0, parts = [];
          for (var i = 0; i < n; i++) { var t = tags.getItemOrNullObject(STORE_KEY + "_" + i); t.load("value"); parts.push(t); }
          return context.sync().then(function () {
            var text = parts.map(function (t) { return t.isNullObject ? "" : t.value; }).join("");
            palette = text ? JSON.parse(text) : [];
          });
        });
      });
    }).catch(function () {
      ST.toast("Couldn't read the colours saved in this presentation.", true);
    }).then(function () { loaded = true; render(); });
  }
  function save() {
    render();
    if (!ST.inOffice || !ST.supports("1.3")) return Promise.resolve();
    var text = JSON.stringify(palette), parts = [];
    for (var i = 0; i < text.length; i += PART_SIZE) parts.push(text.slice(i, i + PART_SIZE));
    return PowerPoint.run(function (context) {
      var tags = context.presentation.tags, old = tags.getItemOrNullObject(STORE_KEY + "_N");
      old.load("value");
      return context.sync().then(function () {
        var oldN = old.isNullObject ? 0 : parseInt(old.value, 10) || 0;
        parts.forEach(function (p, i) { tags.add(STORE_KEY + "_" + i, p); });
        for (var j = parts.length; j < oldN; j++) tags.delete(STORE_KEY + "_" + j);
        tags.add(STORE_KEY + "_N", String(parts.length));
        return context.sync();
      });
    }).catch(function (e) { ST.toast("Couldn't save colours: " + ST.reason(e), true); });
  }

  // ---------- Applying ----------
  function applyColour(c) {
    if (!ST.needOffice()) return;
    if (!ST.supports("1.5")) { ST.toast("This PowerPoint version can't work with the selection.", true); return; }
    var target = ST.state.colTarget;
    ST.run(function (context) {
      if (target === "text") return applyText(context, c);
      var shapes = context.presentation.getSelectedShapes();
      shapes.load("items/type");
      return context.sync().then(function () {
        var list = shapes.items.filter(function (s) { return target === "line" ? s.type !== "Image" && s.type !== "Group" : SHAPE_TYPES.indexOf(s.type) >= 0; });
        if (!list.length) { ST.toast("Select a shape first.", true); return; }
        list.forEach(function (s) {
          if (target === "fill") s.fill.setSolidColor(c.hex);
          else { s.lineFormat.visible = true; s.lineFormat.color = c.hex; }
        });
        return context.sync().then(function () {
          ST.toast((target === "fill" ? "Filled " : "Outlined ") + ST.plural(list.length, "shape") + " with " + (c.name || c.hex) + ".");
        });
      });
    });
  }

  function applyText(context, c) {
    var tagKey = ST.textStyles ? ST.textStyles.shapeTag : "SLIDETOOLS_USES";
    var range = context.presentation.getSelectedTextRangeOrNullObject();
    range.load("length");
    return context.sync().then(function () {
      var jobs = [];
      if (!range.isNullObject && range.length > 0) {
        jobs.push({ range: range, shape: range.getParentTextFrame().getParentShape() });
        return run(jobs);
      }
      var shapes = context.presentation.getSelectedShapes();
      shapes.load("items/type");
      return context.sync().then(function () {
        shapes.items.filter(function (s) { return SHAPE_TYPES.indexOf(s.type) >= 0; }).forEach(function (s) {
          jobs.push({ range: s.textFrame.textRange, shape: s });
        });
        if (!jobs.length) { ST.toast("Select some text, or a text box.", true); return; }
        return run(jobs);
      });
    });
    function run(jobs) {
      jobs.forEach(function (j) { j.tag = j.shape.tags.getItemOrNullObject(tagKey); j.tag.load("value"); });
      return context.sync().then(function () {
        var linked = 0, i = 0;
        function next() {
          if (i >= jobs.length) return Promise.resolve();
          var j = jobs[i++], tagValue = j.tag.isNullObject ? "" : j.tag.value;
          var step = ST.textStyles
            ? ST.textStyles.recolor(context, j.range, j.shape, tagValue, c.hex)
            : (j.range.font.color = c.hex, context.sync().then(function () { return 0; }));
          return step.then(function (n) { linked += n || 0; return next(); });
        }
        return next().then(function () {
          ST.toast("Coloured the text " + (c.name || c.hex) + "." + (linked ? " Styled text stays linked to its style." : ""));
        });
      });
    }
  }

  // ---------- Adding ----------
  function addColour(hex, name) {
    hex = norm(hex);
    if (!hex) { ST.toast("Use a hex code like #0071E3, or a colour name.", true); return; }
    if (palette.some(function (c) { return c.hex === hex; })) { ST.toast(hex + " is already in your colours.", true); return; }
    palette.push({ hex: hex, name: name || "" });
    save();
    ST.toast("Added " + hex + ".");
  }
  function addFromSelection() {
    if (!ST.needOffice()) return;
    ST.run(function (context) {
      var range = ST.supports("1.5") ? context.presentation.getSelectedTextRangeOrNullObject() : null;
      if (range) range.load("length");
      return context.sync().then(function () {
        if (range && !range.isNullObject && range.length > 0) {
          range.font.load("color");
          return context.sync().then(function () {
            var c = range.font.color;
            if (!c) { ST.toast("The selected text has more than one colour.", true); return; }
            var code = ST.styleCode;
            // Styled text carries a tiny hidden shift. Add the style's own colour instead where we can.
            var id = code ? code.decodeId(c) : 0, st = id && ST.textStyles && ST.textStyles.byId(id);
            addColour(st && code.base(c) === code.base(st.font.color) ? st.font.color : c);
          });
        }
        var shapes = context.presentation.getSelectedShapes();
        shapes.load("items/type");
        return context.sync().then(function () {
          var s = shapes.items.filter(function (x) { return SHAPE_TYPES.indexOf(x.type) >= 0; })[0];
          if (!s) { ST.toast("Select a shape or some text first.", true); return; }
          s.fill.load("type,foregroundColor");
          return context.sync().then(function () {
            if (s.fill.type !== "Solid") { ST.toast("That shape doesn't have a solid fill.", true); return; }
            addColour(s.fill.foregroundColor);
          });
        });
      });
    });
  }

  // PowerPoint lays Custom Colors out as rows of 10, padding empty slots with white.
  // A column whose slots are all filled is treated as a family (a colour and its tints);
  // everything else is a single colour. Empty white slots are dropped.
  // True when every colour in the list is a tint or shade of the same hue.
  function sameHue(col) {
    var hs = col.map(function (c) { return rgbToHsl(c.hex); });
    var grey = function (h) { return h[1] < 0.12; };
    if (hs.every(grey)) return true;
    if (hs.some(grey)) return false;
    return hs.every(function (h) { var d = Math.abs(h[0] - hs[0][0]); return Math.min(d, 360 - d) <= 18; });
  }
  function splitTemplate(list) {
    var COLS = 10, out = { singles: [], families: [], rows: 0 };
    var empty = function (c) { return !c || c.hex === "#FFFFFF"; };
    if (list.length < COLS * 2 || list.length % COLS) {
      out.singles = list.filter(function (c) { return list.length < COLS * 2 ? true : !empty(c); });
      return out;
    }
    var rows = list.length / COLS, famCols = [];
    for (var c = 0; c < COLS; c++) {
      var col = [];
      for (var r = 0; r < rows; r++) col.push(list[r * COLS + c]);
      if (col.every(function (x) { return !empty(x); }) && sameHue(col)) famCols.push(c);
    }
    for (var r2 = 0; r2 < rows; r2++) for (var c2 = 0; c2 < COLS; c2++) {
      var x = list[r2 * COLS + c2];
      if (famCols.indexOf(c2) < 0 && !empty(x)) out.singles.push(x);
    }
    out.families = famCols.map(function (fc) { var f = []; for (var r3 = 0; r3 < rows; r3++) f.push(list[r3 * COLS + fc]); return f; });
    out.rows = rows;
    return out;
  }
  ST._splitTemplate = splitTemplate; // exposed for testing

  // ---------- Rendering ----------
  function swatch(c, i) {
    var b = document.createElement("button");
    b.className = "swatch";
    b.style.background = c.hex;
    b.style.color = textOn(c.hex);
    b.title = (c.name ? c.name + "  " : "") + c.hex + ". Click to apply.";
    b.setAttribute("aria-label", b.title);
    b.addEventListener("click", function () { applyColour(c); });
    return b;
  }
  function render() {
    var themeRow = $("colTheme"), mine = $("colMine"), edit = $("colEdit");
    if (!themeRow) return;
    updateDatalist();
    Array.prototype.forEach.call($("colTarget").children, function (b) { b.classList.toggle("on", b.getAttribute("data-v") === ST.state.colTarget); });
    themeRow.innerHTML = ""; mine.innerHTML = ""; edit.innerHTML = "";
    // Base theme colours, then five rows of tints and shades under each, like PowerPoint.
    theme.forEach(function (c, i) { themeRow.appendChild(swatch(c, i)); });
    var tints = theme.map(variants), grid = $("colTints");
    grid.innerHTML = "";
    grid.style.gridTemplateColumns = "repeat(" + Math.max(1, theme.length) + ", 1fr)";
    themeRow.style.gridTemplateColumns = grid.style.gridTemplateColumns;
    for (var r = 0; r < 5; r++) tints.forEach(function (col) { if (col[r]) grid.appendChild(swatch(col[r])); });
    // Template colours: single colours in one row, then each family as a column with its tints.
    var split = splitTemplate(template), singles = $("colSingles"), fams = $("colFamilies");
    singles.innerHTML = ""; fams.innerHTML = "";
    split.singles.forEach(function (c) { singles.appendChild(swatch(c)); });
    fams.style.gridTemplateColumns = "repeat(" + Math.max(1, split.families.length) + ", 1fr)";
    for (var fr = 0; fr < split.rows; fr++) split.families.forEach(function (f) { if (f[fr]) fams.appendChild(swatch(f[fr])); });
    $("colSinglesLabel").hidden = !split.singles.length || !split.families.length;
    $("colFamiliesLabel").hidden = !split.families.length;
    // Tabs: hide Theme or Template when the deck doesn't have them.
    var tabs = { theme: theme.length > 0, template: template.length > 0, mine: true };
    var tab = tabs[ST.state.colTab] ? ST.state.colTab : (tabs.theme ? "theme" : tabs.template ? "template" : "mine");
    Array.prototype.forEach.call($("colTabs").children, function (b) {
      var v = b.getAttribute("data-v");
      b.hidden = !tabs[v]; b.classList.toggle("on", v === tab);
    });
    $("colTabs").style.gridTemplateColumns = "repeat(" + Object.keys(tabs).filter(function (k) { return tabs[k]; }).length + ", 1fr)";
    $("colThemeWrap").hidden = tab !== "theme";
    $("colTemplateWrap").hidden = tab !== "template";
    $("colMineWrap").hidden = tab !== "mine";
    if (!palette.length) mine.innerHTML = '<div class="muted-note">' + (loaded || !ST.inOffice ? "No colours yet. Add one below, or take one from the selection." : "Reading colours from this presentation.") + "</div>";
    palette.forEach(function (c, i) { mine.appendChild(swatch(c, i)); });
    palette.forEach(function (c, i) {
      var row = document.createElement("div");
      row.className = "col-row";
      row.innerHTML = '<span class="col-chip"></span><input type="text" maxlength="24" placeholder="Name"><span class="col-hex"></span>' +
        '<button class="icon-btn" title="Move up">&uarr;</button><button class="icon-btn" title="Move down">&darr;</button><button class="icon-btn" title="Remove">&times;</button>';
      row.querySelector(".col-chip").style.background = c.hex;
      row.querySelector(".col-hex").textContent = c.hex;
      var name = row.querySelector("input");
      name.value = c.name || "";
      name.addEventListener("change", function () { c.name = name.value.trim(); save(); });
      var btns = row.querySelectorAll(".icon-btn");
      btns[0].disabled = i === 0; btns[1].disabled = i === palette.length - 1;
      btns[0].addEventListener("click", function () { palette.splice(i - 1, 0, palette.splice(i, 1)[0]); save(); });
      btns[1].addEventListener("click", function () { palette.splice(i + 1, 0, palette.splice(i, 1)[0]); save(); });
      btns[2].addEventListener("click", function () { palette.splice(i, 1); save(); ST.toast("Removed " + c.hex + "."); });
      edit.appendChild(row);
    });
  }

  var MARKUP =
    '<span class="info res-info" tabindex="0" role="button" aria-label="How Colours works">i' +
      '<span class="tip" role="tooltip">Pick what to colour, then click a swatch to apply it to the selected shapes or text.<br><br>' +
      'Theme colours, with the same tints and shades as PowerPoint\'s colour menu, and the template\'s custom colours come from this presentation\'s theme. They can\'t be changed here. Your colours are saved in the presentation.<br><br>' +
      'Colouring styled text here keeps it linked to its text style, unlike the ribbon\'s font colour button. Shapes keep their shape style link either way.<br><br>' +
      'The colour pickers in the style editors suggest these colours too.</span>' +
    '</span>' +
    '<div class="seg seg3" id="colTarget"></div>' +
    '<div class="seg col-tabs" id="colTabs"><button data-v="theme">Theme</button><button data-v="template">Template</button><button data-v="mine">Yours</button></div>' +
    '<div id="colThemeWrap" class="col-pane"><div class="theme-grid" id="colTheme"></div><div class="theme-grid tints" id="colTints"></div></div>' +
    '<div id="colTemplateWrap" class="col-pane" hidden>' +
      '<div class="sublabel" id="colSinglesLabel">Colours</div><div class="fixed-grid" id="colSingles"></div>' +
      '<div class="sublabel" id="colFamiliesLabel">Families</div><div class="theme-grid families" id="colFamilies"></div>' +
    '</div>' +
    '<div id="colMineWrap" class="col-pane" hidden>' +
      '<div class="fixed-grid" id="colMine"></div>' +
      '<div class="inline"><div class="color-field"><input type="color" id="colPick" value="#0071e3"><input type="text" id="colHex" placeholder="#0071E3"></div>' +
        '<button class="btn small" id="colAdd">Add</button></div>' +
      '<button class="btn small" id="colFromSel">Add from selection</button>' +
      '<div class="group-label" data-min="2">Edit your colours</div>' +
      '<div class="col-edit" id="colEdit" data-min="2"></div>' +
    '</div>';

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    var seg = $("colTarget");
    [["fill", "Fill"], ["line", "Outline"], ["text", "Text"]].forEach(function (o) {
      var b = document.createElement("button");
      b.textContent = o[1]; b.setAttribute("data-v", o[0]);
      b.addEventListener("click", function () { ST.state.colTarget = o[0]; ST.save(); render(); });
      seg.appendChild(b);
    });
    Array.prototype.forEach.call($("colTabs").children, function (b) {
      b.addEventListener("click", function () { ST.state.colTab = b.getAttribute("data-v"); ST.save(); render(); });
    });
    var pick = $("colPick"), hex = $("colHex");
    pick.addEventListener("input", function () { hex.value = pick.value.toUpperCase(); });
    hex.addEventListener("input", function () { var c = norm(hex.value); if (c) pick.value = c.toLowerCase(); });
    $("colAdd").addEventListener("click", function () { addColour(hex.value || pick.value); hex.value = ""; });
    hex.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); addColour(hex.value); hex.value = ""; } });
    $("colFromSel").addEventListener("click", addFromSelection);
    ST.on("ready", load);
    ST.on("reset", render);
    render();
  }

  ST.registerTool({
    id: "colours",
    title: "Colours",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 1,
    defaults: defaults,
    build: build
  });

  // Exposed for testing only
  ST._colours = { get: function () { return palette; }, theme: function () { return theme; } };
})();
