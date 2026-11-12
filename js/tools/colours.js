/* Colours tool: a palette saved in the presentation, plus the theme colours, applied to
   fill, outline or text in one click. Text recolouring keeps Styles links intact.
   Theme colours are only read, never changed, because they live on the slide master. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;
  var STORE_KEY = "SLIDETOOLS_COLOURS";
  var PART_SIZE = 2000;
  var SHAPE_TYPES = ["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform"];
  var THEME_SLOTS = ["Dark1", "Light1", "Dark2", "Light2", "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6"];
  var THEME_NAMES = { Dark1: "Dark 1", Light1: "Light 1", Dark2: "Dark 2", Light2: "Light 2", Accent1: "Accent 1", Accent2: "Accent 2", Accent3: "Accent 3", Accent4: "Accent 4", Accent5: "Accent 5", Accent6: "Accent 6" };

  var defaults = { colTarget: "fill" };
  ST.addDefaults(defaults);

  var ICON = '<svg viewBox="0 0 18 18"><circle cx="6" cy="6.5" r="4" fill="currentColor"/><circle cx="12" cy="6.5" r="4" fill="currentColor" opacity="0.55"/><circle cx="9" cy="11.8" r="4" fill="currentColor" opacity="0.3"/></svg>';

  var palette = [], theme = [], loaded = false;

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

  // Shared with the style editors: their colour pickers suggest palette and theme colours.
  ST.palette = { list: function () { return theme.concat(palette); } };
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
      var tags = context.presentation.tags, count = tags.getItemOrNullObject(STORE_KEY + "_N");
      count.load("value");
      var themeResults = [];
      if (ST.supports("1.10")) {
        var slides = context.presentation.slides;
        slides.load("items");
        return context.sync().then(function () {
          if (slides.items.length) {
            var scheme = slides.items[0].themeColorScheme;
            THEME_SLOTS.forEach(function (slot) { themeResults.push({ slot: slot, r: scheme.getThemeColor(slot) }); });
          }
          return readParts();
        });
      }
      return readParts();
      function readParts() {
        return context.sync().then(function () {
          var n = count.isNullObject ? 0 : parseInt(count.value, 10) || 0, parts = [];
          for (var i = 0; i < n; i++) { var t = tags.getItemOrNullObject(STORE_KEY + "_" + i); t.load("value"); parts.push(t); }
          return context.sync().then(function () {
            var text = parts.map(function (t) { return t.isNullObject ? "" : t.value; }).join("");
            palette = text ? JSON.parse(text) : [];
            theme = themeResults.map(function (t) { var hex = norm(t.r.value); return hex ? { hex: hex, name: THEME_NAMES[t.slot], theme: true } : null; }).filter(Boolean);
          });
        });
      }
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
    }).catch(function (e) { ST.toast("Couldn't save colours: " + (e && e.message ? e.message : e), true); });
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
    $("colThemeWrap").hidden = !theme.length;
    theme.forEach(function (c, i) { themeRow.appendChild(swatch(c, i)); });
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
      'Theme colours come from this presentation\'s theme and can\'t be changed here. Your colours are saved in the presentation.<br><br>' +
      'Colouring styled text here keeps it linked to its text style, unlike the ribbon\'s font colour button. Shapes keep their shape style link either way.<br><br>' +
      'The colour pickers in the style editors suggest these colours too.</span>' +
    '</span>' +
    '<div class="seg seg3" id="colTarget"></div>' +
    '<div id="colThemeWrap"><div class="sublabel">Theme</div><div class="swatches" id="colTheme"></div></div>' +
    '<div class="sublabel">Your colours</div>' +
    '<div class="swatches" id="colMine"></div>' +
    '<div class="inline"><div class="color-field"><input type="color" id="colPick" value="#0071e3"><input type="text" id="colHex" placeholder="#0071E3"></div>' +
      '<button class="btn small" id="colAdd">Add</button></div>' +
    '<button class="btn small" id="colFromSel">Add from selection</button>' +
    '<div class="group-label" data-min="2">Edit your colours</div>' +
    '<div class="col-edit" id="colEdit" data-min="2"></div>';

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    var seg = $("colTarget");
    [["fill", "Fill"], ["line", "Outline"], ["text", "Text"]].forEach(function (o) {
      var b = document.createElement("button");
      b.textContent = o[1]; b.setAttribute("data-v", o[0]);
      b.addEventListener("click", function () { ST.state.colTarget = o[0]; ST.save(); render(); });
      seg.appendChild(b);
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
