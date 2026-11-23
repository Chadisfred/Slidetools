/* Tables tool: named table styles saved in the presentation. Insert a table that is already
   formatted, or format the selected table. Each table's columns have a role:
     Category  stacked labels (shown once per group), left aligned
     Label     row text that doesn't repeat, left aligned
     Data      numbers, right aligned, with a thin line on their left
   The rules (header and bottom rules, group rules, total rows, column bands, padding and
   fonts) come from the style. Only normal slides are read or changed. */
(function () {
  "use strict";

  var ST = window.ST, PT = ST.PT, $ = ST.$;
  var STORE_KEY = "SLIDETOOLS_TABLESTYLES";
  var TABLE_TAG = "SLIDETOOLS_TABLE";
  var PART_SIZE = 2000;
  var ROLE_NAMES = { C: "Category", L: "Label", D: "Data" };
  var PARTS = [["header", "Header row"], ["first", "First column"], ["body", "Body"], ["total", "Total rows"]];
  var LINES = [["header", "Rule under header"], ["bottom", "Rule under table"], ["group", "Rule above a new group"],
               ["total", "Rule above total rows"], ["row", "Lines between rows"], ["vertical", "Lines left of data columns"]];

  var defaults = { tblRows: "6", tblCols: "4" };
  ST.addDefaults(defaults);

  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.4"><path d="M2 4.5h14" stroke-width="2"/><path d="M2 8.5h14M2 12h14" stroke-width="0.9"/><path d="M2 15.5h14" stroke-width="2"/><rect x="7" y="5.5" width="4" height="9.5" fill="currentColor" fill-opacity="0.22" stroke="none"/></svg>';
  var EDIT_ICON = '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M10.5 2.5l3 3L6 13H3v-3l7.5-7.5z"/></svg>';

  var styles = [], stylesLoaded = false, editing = null, busy = false;
  var current = null; // the selected table: { shapeId, slideId, cols, values, roles, styleId }

  // ---------- Style model ----------
  function font(name, bold) { return { ts: null, name: name, size: 12, color: "#000000", bold: !!bold, italic: false }; }
  function line(w) { return { on: true, w: w, c: "#000000" }; }
  function defaultProps() {
    return {
      fonts: { header: font("Avenir Next LT Pro", true), first: font("Avenir Next LT Pro Demi"), body: font("Avenir Next LT Pro"), total: font("Avenir Next LT Pro Demi") },
      lines: { header: line(1), bottom: line(1), group: line(1), total: line(1), row: line(0.25), vertical: line(0.25) },
      band: { on: true, color: "#F2F2F2" },
      pad: { all: 3.6, firstLeft: 0, lastRight: 0 },
      totalPrefix: "Total",
      thinSingleGroups: true,
      trimSpaces: true,
      alignDecimals: true,
      size: { on: true, rowMin: 1, rowStep: 0.25, fitColumns: true }
    };
  }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function byId(id) { return styles.filter(function (s) { return s.id === id; })[0]; }
  function nextId() { var id = 1; while (byId(id)) id++; return id; }
  function pending(s) { return JSON.stringify(s.props) !== JSON.stringify(s.applied); }
  function esc(t) { return String(t === null || t === undefined ? "" : t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function seed() { if (!styles.length) styles.push({ id: 1, name: "Default table", props: defaultProps(), applied: null }); }

  // ---------- The formatting rules (pure, shared by Insert and Format) ----------
  function blank(v) { return !String(v === null || v === undefined ? "" : v).trim(); }
  function normRoles(roles, cols) {
    var out = [];
    for (var c = 0; c < cols; c++) out.push(roles && /^[CLD]$/.test(roles[c]) ? roles[c] : (c === 0 ? "C" : "D"));
    return out;
  }
  // For each body row that is a total row, the column where its "Total" text sits.
  function totalColumns(values, P) {
    var prefix = String(P.totalPrefix || "").trim().toLowerCase(), out = {};
    if (!prefix) return out;
    for (var r = 1; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) {
        if (blank(values[r][c])) continue;
        if (String(values[r][c]).trim().toLowerCase().indexOf(prefix) === 0) out[r] = c;
        break;
      }
    }
    return out;
  }
  function fontFor(part, P) {
    var f = P.fonts[part], out = { name: f.name, size: f.size, color: f.color, bold: f.bold, italic: f.italic };
    var t = f.ts && ST.textStyles ? ST.textStyles.byId(f.ts) : null;
    if (t && t.font) ["name", "size", "color", "bold", "italic"].forEach(function (k) {
      var v = t.font[k];
      if (v !== null && v !== undefined && v !== "" && v !== "leave") out[k] = k === "color" && ST.styleCode ? ST.styleCode.base(v) : v;
    });
    Object.keys(out).forEach(function (k) { if (out[k] === null || out[k] === undefined || out[k] === "") delete out[k]; });
    return out;
  }
  var NONE = { color: "#000000", weight: 0.25, transparency: 1, dashStyle: "Solid" };
  function edge(spec) { return spec && spec.on ? { color: spec.c, weight: Number(spec.w), transparency: 0, dashStyle: "Solid" } : NONE; }

  // values: 2D text (row 0 is the header). roles: array of C, L, D. merged: { "r,c": true }.
  function cellProps(values, roles, P, merged) {
    var R = values.length, C = R ? values[0].length : 0, L = P.lines, totals = totalColumns(values, P);
    roles = normRoles(roles, C);
    merged = merged || {};
    // With a single category column, group lines use the thin row line (header, total and
    // bottom rules stay heavy). Older styles without the setting get it on.
    var singleCategory = P.thinSingleGroups !== false && roles.filter(function (x) { return x === "C"; }).length === 1;
    // The line above row r (r = 1..R, where R is the bottom edge of the table), per column.
    function hline(r) {
      var out = [], c;
      if (r === 1 && R > 1) { for (c = 0; c < C; c++) out.push(L.header); return out; }
      if (r === R) { for (c = 0; c < C; c++) out.push(L.bottom); return out; }
      var start = C, kind = null;
      if (totals[r] !== undefined) { start = totals[r]; kind = L.total; }
      else for (c = 0; c < C; c++) if (roles[c] === "C" && !blank(values[r][c])) { start = c; kind = singleCategory ? L.row : L.group; break; }
      for (c = 0; c < C; c++) {
        if (c >= start) out.push(kind);
        else out.push(roles[c] === "C" ? null : L.row);
      }
      return out;
    }
    function vline(c) { return c > 0 && c < C && roles[c] === "D" ? L.vertical : null; }
    var lines = [];
    for (var r = 1; r <= R; r++) lines[r] = hline(r);
    var out = [];
    for (var i = 0; i < R; i++) {
      var row = [], total = totals[i] !== undefined;
      for (var c = 0; c < C; c++) {
        var part = i === 0 ? "header" : total ? "total" : c === 0 ? "first" : "body";
        var cell = {
          font: fontFor(part, P),
          fill: i > 0 && P.band.on && (c + 1) % 2 === 0 ? { color: P.band.color, transparency: 0 } : { transparency: 1 },
          margins: {
            top: Number(P.pad.all), bottom: Number(P.pad.all),
            left: c === 0 ? Number(P.pad.firstLeft) : Number(P.pad.all),
            right: c === C - 1 ? Number(P.pad.lastRight) : Number(P.pad.all)
          },
          horizontalAlignment: roles[c] === "D" ? "Right" : "Left",
          verticalAlignment: i === 0 ? "Bottom" : merged[i + "," + c] ? "Top" : "Middle",
          borders: {
            top: i === 0 ? NONE : edge(lines[i][c]),
            bottom: edge(lines[i + 1][c]),
            left: edge(vline(c)),
            right: edge(vline(c + 1))
          }
        };
        row.push(cell);
      }
      out.push(row);
    }
    return out;
  }

  // The same settings for addTable, leaving out lines and fills that would only hide something.
  function insertProps(values, roles, P) {
    return cellProps(values, roles, P, {}).map(function (row) {
      return row.map(function (p) {
        var out = { font: p.font, margins: p.margins, horizontalAlignment: p.horizontalAlignment, verticalAlignment: p.verticalAlignment, borders: {} };
        if (p.fill.transparency !== 1) out.fill = { color: p.fill.color };
        Object.keys(p.borders).forEach(function (k) { var e = p.borders[k]; if (e.transparency !== 1) out.borders[k] = { color: e.color, weight: e.weight }; });
        return out;
      });
    });
  }

  // Blanks a category label when it repeats the row above within the same parent group.
  // Removes spaces, tabs and line breaks at the start and end of each cell (text pasted from
  // Excel often has them), leaving everything inside the text alone.
  // Pads numbers in each data column to the column's largest number of decimal places, so the
  // points line up in right aligned cells: 1.5 next to 2.25 becomes 1.50. Never removes digits.
  var NUMBER = /^(\s*[(\-\u2212+]?\s*[$\u20ac\u00a3]?\s*)(\d{1,3}(?:,\d{3})+|\d+)(?:\.(\d+))?(\s*(?:%|[kKmMbB]n?)?\)?\s*)$/;
  function decimalValues(values, roles) {
    var out = values.map(function (row) { return row.slice(); }), changed = [];
    for (var c = 0; c < roles.length; c++) {
      if (roles[c] !== "D") continue;
      var most = 0, hits = [];
      for (var r = 1; r < values.length; r++) {
        var m = NUMBER.exec(String(values[r][c] === null || values[r][c] === undefined ? "" : values[r][c]));
        if (m) { hits.push([r, m]); most = Math.max(most, (m[3] || "").length); }
      }
      if (!most) continue;
      hits.forEach(function (h) {
        var m = h[1], dec = m[3] || "";
        if (dec.length >= most) return;
        out[h[0]][c] = m[1] + m[2] + "." + dec + new Array(most - dec.length + 1).join("0") + m[4];
        changed.push([h[0], c]);
      });
    }
    return { values: out, changed: changed };
  }
  // Rows and columns from text copied out of Excel: tabs between cells, line breaks between
  // rows, and quotes around cells that hold a tab, a line break or a quote.
  function parseClipboard(text) {
    var t = String(text || "").replace(/\r\n?/g, "\n"), rows = [], row = [], cell = "", i = 0, quoted = false;
    if (!t.trim()) return [];
    while (i < t.length) {
      var ch = t.charAt(i);
      if (quoted) {
        if (ch === '"' && t.charAt(i + 1) === '"') { cell += '"'; i += 2; continue; }
        if (ch === '"') { quoted = false; i++; continue; }
        cell += ch; i++; continue;
      }
      if (ch === '"' && cell === "") { quoted = true; i++; continue; }
      if (ch === "\t") { row.push(cell); cell = ""; i++; continue; }
      if (ch === "\n") { row.push(cell); rows.push(row); row = []; cell = ""; i++; continue; }
      cell += ch; i++;
    }
    if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
    while (rows.length && rows[rows.length - 1].every(function (v) { return !String(v).trim(); })) rows.pop();
    var width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    return rows.map(function (r) { while (r.length < width) r.push(""); return r; });
  }

  function trimmedValues(values) {
    var out = values.map(function (row) { return row.slice(); }), changed = [];
    values.forEach(function (row, r) {
      row.forEach(function (v, c) {
        var text = v === null || v === undefined ? "" : String(v), clean = text.trim();
        if (clean !== text) { out[r][c] = clean; changed.push([r, c]); }
      });
    });
    return { values: out, changed: changed };
  }

  function stackedValues(values, roles, P) {
    var R = values.length, C = R ? values[0].length : 0, totals = totalColumns(values, P), out = clone(values), filled = [];
    roles = normRoles(roles, C);
    for (var r = 0; r < R; r++) {
      filled[r] = [];
      for (var c = 0; c < C; c++) filled[r][c] = !blank(values[r][c]) ? String(values[r][c]).trim() : (r > 1 && totals[r - 1] === undefined ? filled[r - 1][c] : "");
    }
    var changed = 0;
    for (r = 2; r < R; r++) {
      if (totals[r] !== undefined || totals[r - 1] !== undefined) continue;
      for (c = 0; c < C; c++) {
        if (roles[c] !== "C" || blank(values[r][c])) continue;
        if (filled[r].slice(0, c + 1).join("\u0001") === filled[r - 1].slice(0, c + 1).join("\u0001")) { out[r][c] = ""; changed++; }
      }
    }
    return { values: out, changed: changed };
  }

  // A first guess at each column's role from its contents.
  function guessRoles(values) {
    var R = values.length, C = R ? values[0].length : 0, roles = [], totals = totalColumns(values, defaultProps()), num = /^[\s(\-−+]*[$€£]?\s*[\d.,]+\s*(%|[kKmMbB]n?)?\)?\s*$/;
    for (var c = 0; c < C; c++) {
      var filled = 0, numbers = 0, blanks = 0, seen = {}, repeats = 0;
      for (var r = 1; r < R; r++) {
        if (totals[r] !== undefined) continue; // total rows often leave label cells empty
        var v = values[r][c];
        if (blank(v)) { blanks++; continue; }
        filled++;
        if (num.test(String(v))) numbers++;
        var k = String(v).trim().toLowerCase();
        if (seen[k]) repeats++; seen[k] = 1;
      }
      var numeric = filled && numbers / filled >= 0.6;
      var leading = roles.every(function (x) { return x === "C"; });
      // Category columns sit at the left and repeat or leave blanks (years count too)
      if (c === 0 || (leading && (blanks || repeats))) roles.push("C");
      else if (numeric) roles.push("D");
      else roles.push("L");
    }
    return roles;
  }
  function insertRoles(cols) {
    if (cols <= 1) return ["L"];
    if (cols === 2) return ["C", "D"];
    if (cols === 3) return ["C", "L", "D"];
    var r = ["C", "C", "L"];
    while (r.length < cols) r.push("D");
    return r;
  }

  // ---------- Saving styles in the presentation ----------
  function loadStyles() {
    if (!ST.inOffice || !ST.supports("1.3")) { seed(); stylesLoaded = true; render(); return; }
    PowerPoint.run(function (context) {
      var tags = context.presentation.tags, count = tags.getItemOrNullObject(STORE_KEY + "_N");
      count.load("value");
      return context.sync().then(function () {
        var n = count.isNullObject ? 0 : parseInt(count.value, 10) || 0, parts = [];
        for (var i = 0; i < n; i++) { var t = tags.getItemOrNullObject(STORE_KEY + "_" + i); t.load("value"); parts.push(t); }
        return context.sync().then(function () {
          var text = parts.map(function (t) { return t.isNullObject ? "" : t.value; }).join("");
          styles = text ? JSON.parse(text) : [];
          styles.forEach(function (st) {
            if (!st.props.size) st.props.size = defaultProps().size;
            if (st.applied && !st.applied.size) st.applied.size = clone(st.props.size);
          });
        });
      });
    }).catch(function () {
      styles = [];
      ST.toast("Couldn't read the table styles saved in this presentation.", true);
    }).then(function () { seed(); stylesLoaded = true; render(); refreshTable(); });
  }
  function saveStyles() {
    render();
    if (!ST.inOffice || !ST.supports("1.3")) return Promise.resolve();
    var text = JSON.stringify(styles), parts = [];
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
    }).catch(function (e) { ST.toast("Couldn't save table styles: " + ST.reason(e), true); });
  }

  // ---------- Reading and writing tables ----------
  function canInsert() { return ST.supports("1.8"); }
  function canFormat() { return ST.supports("1.9"); }
  function parseTag(v) { try { var o = JSON.parse(v || "null"); return o && typeof o === "object" ? o : {}; } catch (e) { return {}; } }

  // Loads a table shape's text, size, merged areas and Slide Tools tag.
  function readTable(context, shape) {
    var table = shape.getTable(), tag = shape.tags.getItemOrNullObject(TABLE_TAG);
    table.load("rowCount,columnCount,values");
    tag.load("value");
    var areas = table.getMergedAreas();
    areas.load("items/rowIndex,items/columnIndex,items/rowCount,items/columnCount");
    return context.sync().then(function () {
      var merged = {}, anchors = {};
      areas.items.forEach(function (a) {
        anchors[a.rowIndex + "," + a.columnIndex] = true;
        for (var r = a.rowIndex; r < a.rowIndex + a.rowCount; r++)
          for (var c = a.columnIndex; c < a.columnIndex + a.columnCount; c++) merged[r + "," + c] = true;
      });
      var t = tag.isNullObject ? {} : parseTag(tag.value);
      return { table: table, shape: shape, rows: table.rowCount, cols: table.columnCount, values: table.values, merged: merged, anchors: anchors, tag: t };
    });
  }
  // ---------- Line check ----------
  // Some PowerPoint versions apply table line transparency the opposite way round, so a
  // line meant to show is hidden and the other way round. A one-time check finds out.
  var LINE_KEY = "slideToolsLineMode", lineModeMem = "", lineCheck = null;
  function lineMode() { try { return localStorage.getItem(LINE_KEY) || lineModeMem; } catch (e) { return lineModeMem; } }
  function setLineMode(m) { lineModeMem = m; try { localStorage.setItem(LINE_KEY, m); } catch (e) { /* remembered for this session only */ } }
  function lineAlpha(t) { return lineMode() === "inverted" ? 1 - t : t; }
  function writeBorder(b, e) {
    b.color = e.color; b.weight = e.weight; b.transparency = lineAlpha(e.transparency); b.dashStyle = e.dashStyle;
  }
  // Runs task straight away when the line mode is known, otherwise after the line check.
  function withLines(task) {
    if (!canFormat() || lineMode()) { task(); return; }
    startLineCheck(task);
  }
  function startLineCheck(then) {
    if (!ST.needOffice() || busy) return;
    if (!canFormat()) { ST.toast("This PowerPoint version doesn't need the line check.", false); return; }
    busy = true;
    ST.run(function (context) {
      var slides = context.presentation.getSelectedSlides();
      slides.load("items/id");
      return context.sync().then(function () {
        if (!slides.items.length) { ST.toast("Go to a slide first.", true); return; }
        var slide = slides.items[0];
        var shape = slide.shapes.addTable(1, 2, { values: [["A", "B"]], left: 36, top: 36, width: 216, style: "NoStyleNoGrid" });
        shape.load("id");
        return context.sync().then(function () {
          var t = shape.getTable();
          [[0, 0], [0, 1]].forEach(function (rc) {
            var cell = t.getCellOrNullObject(rc[0], rc[1]), b = cell.borders.bottom;
            b.color = "#000000"; b.weight = 4; b.dashStyle = "Solid"; b.transparency = rc[1] === 0 ? 0 : 1;
            cell.font.size = 24; cell.font.color = "#000000";
            cell.horizontalAlignment = "Center";
          });
          return context.sync();
        }).then(function () {
          lineCheck = { slideId: slide.id, shapeId: shape.id, then: then, reselect: target };
          renderLineCheck();
        });
      });
    }).then(function () { busy = false; });
  }
  function answerLineCheck(answer) {
    var check = lineCheck;
    if (!check) return;
    lineCheck = null;
    setLineMode(answer === "B" ? "inverted" : "normal");
    renderLineCheck();
    ST.run(function (context) {
      context.presentation.slides.getItem(check.slideId).shapes.getItem(check.shapeId).delete();
      var back = check.reselect;
      if (back && ST.supports("1.5")) context.presentation.slides.getItem(back.slideId).setSelectedShapes([back.shapeId]);
      return context.sync();
    }).then(function () {
      if (answer === "none") ST.toast("Lines will be set the standard way. If they look wrong, run Check lines again at the More level.", true);
      else ST.toast(answer === "B" ? "Got it. Lines are set the other way round for this PowerPoint." : "Got it. Lines are set the standard way.");
      if (check.then) check.then();
    });
  }
  function renderLineCheck() {
    var box = $("tblCheck");
    if (!box) return;
    box.hidden = !lineCheck;
    if (lineCheck) box.querySelector(".check-q").textContent = "A small test table with cells A and B is on the slide. Which cell has a thick line under it?";
  }
  // Each kind of setting is sent to PowerPoint on its own, so if one is refused the others
  // still apply, and the panel can say which one it was.
  var GROUPS = [
    ["lines", function (cell, p, fresh) {
      ["top", "bottom", "left", "right"].forEach(function (side) {
        var e = p.borders[side];
        if (fresh && e.transparency === 1) return; // a new table has no lines to hide
        writeBorder(cell.borders[side], e);
      });
    }],
    ["fill", function (cell, p, fresh) {
      if (p.fill.transparency === 1) { if (!fresh) cell.fill.clear(); }
      else cell.fill.setSolidColor(p.fill.color);
    }],
    ["fonts", function (cell, p) { Object.keys(p.font).forEach(function (k) { cell.font[k] = p.font[k]; }); }],
    ["padding", function (cell, p) {
      cell.margins.left = p.margins.left; cell.margins.right = p.margins.right;
      cell.margins.top = p.margins.top; cell.margins.bottom = p.margins.bottom;
    }],
    ["alignment", function (cell, p) { cell.horizontalAlignment = p.horizontalAlignment; cell.verticalAlignment = p.verticalAlignment; }]
  ];
  // ---------- Grid-based sizes ----------
  var measureCtx = null;
  function textWidth(text, f) {
    measureCtx = measureCtx || document.createElement("canvas").getContext("2d");
    measureCtx.font = (f.italic ? "italic " : "") + (f.bold ? "bold " : "") + (f.size || 12) + 'pt "' + String(f.name || "Segoe UI").replace(/"/g, "") + '", "Segoe UI", sans-serif';
    return String(text || "").split(/[\r\n\v\u000b]+/).reduce(function (m, line) { return Math.max(m, measureCtx.measureText(line).width * 0.75); }, 0);
  }
  // Column widths in whole steps of g that fill the grid area width. Returns null when the
  // grid has no usable spacing.
  function columnWidths(values, roles, P, merged) {
    var grid = ST.calculate(), g = grid.spacing * PT;
    if (!(g > 0) || !(grid.maxW > 0)) return null;
    var R = values.length, C = R ? values[0].length : 0, props = cellProps(values, roles, P, merged);
    roles = normRoles(roles, C);
    var W = Math.floor(grid.maxW * PT / g + 1e-6) * g, widths = [], wider = false, c, r;
    if (!P.size.fitColumns) {
      var each = Math.max(g, Math.floor(W / C / g + 1e-6) * g);
      for (c = 0; c < C; c++) widths.push(each);
      widths[0] += Math.max(0, W - each * C);
      return { widths: widths, W: W, g: g, wider: each * C > W + 0.01 };
    }
    var need = [];
    for (c = 0; c < C; c++) {
      need[c] = 0;
      for (r = 0; r < R; r++) {
        if (merged && merged[r + "," + c]) continue;
        var p = props[r][c];
        need[c] = Math.max(need[c], textWidth(values[r][c], p.font) + p.margins.left + p.margins.right + 2);
      }
      widths[c] = Math.max(g, Math.ceil(need[c] / g - 1e-6) * g);
    }
    var labels = [], datas = [], fixed = 0;
    for (c = 0; c < C; c++) { if (roles[c] === "L") labels.push(c); else { fixed += widths[c]; if (roles[c] === "D") datas.push(c); } }
    if (labels.length) {
      // Label columns share whatever is left, in whole g
      var unit = Math.floor((W - fixed) / labels.length / g + 1e-6) * g;
      if (unit < g) { wider = true; labels.forEach(function (i) { widths[i] = g; }); }
      else {
        labels.forEach(function (i) { widths[i] = unit; });
        widths[labels[0]] += Math.round((W - fixed - unit * labels.length) / g) * g;
      }
    } else {
      // No label column: spread the spare room over the data columns (or the last column)
      var share = datas.length ? datas : [C - 1], extra = W - fixed;
      if (extra < 0) wider = true;
      else {
        var add = Math.floor(extra / share.length / g + 1e-6) * g;
        share.forEach(function (i) { widths[i] += add; });
        widths[share[0]] += Math.round((extra - add * share.length) / g) * g;
      }
    }
    return { widths: widths, W: W, g: g, wider: wider };
  }

  // Rewrites only the cells whose text has spaces at the start or end. Resolves to the count.
  function trimCells(context, info) {
    var res = trimmedValues(info.values), count = 0;
    res.changed.forEach(function (rc) {
      if (info.merged[rc[0] + "," + rc[1]] && !info.anchors[rc[0] + "," + rc[1]]) return;
      info.table.getCellOrNullObject(rc[0], rc[1]).text = res.values[rc[0]][rc[1]];
      count++;
    });
    if (!count) return Promise.resolve(0);
    return context.sync().then(function () { info.values = res.values; return count; });
  }

  // Writes the style's rules into a table, cell by cell. Resolves to the list of setting
  // kinds PowerPoint refused (empty when everything applied). fresh: a table just inserted.
  // Unless the style says otherwise, spaces at the start and end of cells are trimmed first,
  // so column widths and wrapping are worked out on the clean text.
  // Rewrites only the number cells that need padding to line up their decimals.
  function decimalCells(context, info, roles) {
    var res = decimalValues(info.values, roles), count = 0;
    res.changed.forEach(function (rc) {
      if (info.merged[rc[0] + "," + rc[1]] && !info.anchors[rc[0] + "," + rc[1]]) return;
      info.table.getCellOrNullObject(rc[0], rc[1]).text = res.values[rc[0]][rc[1]];
      count++;
    });
    if (!count) return Promise.resolve(0);
    return context.sync().then(function () { info.values = res.values; return count; });
  }
  function formatTable(context, info, style, roles, fresh) {
    var refused = [];
    var first = style.props.trimSpaces !== false && !fresh
      ? trimCells(context, info).catch(function (e) { if (window.console) console.error("Table spaces:", e); refused.push("trim spaces"); })
      : Promise.resolve();
    if (style.props.alignDecimals !== false && !fresh) {
      first = first.then(function () { return decimalCells(context, info, roles); })
        .catch(function (e) { if (window.console) console.error("Table decimals:", e); refused.push("line up decimals"); });
    }
    return first.then(function () { return formatRules(context, info, style, roles, fresh); })
      .then(function (failed) { refused.forEach(function (f) { failed.unshift(f); }); return failed; });
  }
  function formatRules(context, info, style, roles, fresh) {
    var props = cellProps(info.values, roles, style.props, info.merged), failed = [];
    function settings() {
      try {
        var s = info.table.styleSettings;
        s.style = "NoStyleNoGrid";
        s.areRowsBanded = false; s.areColumnsBanded = false;
        s.isFirstRowHighlighted = false; s.isFirstColumnHighlighted = false;
        s.isLastRowHighlighted = false; s.isLastColumnHighlighted = false;
      } catch (e) { return Promise.resolve(); }
      return context.sync().catch(function () { /* older hosts: the cell settings still apply */ });
    }
    function group(g) {
      var count = 0;
      function from(r) {
        for (; r < info.rows; r++) {
          for (var c = 0; c < info.cols; c++) {
            // Inside a merged area only the top left cell can be changed
            if (info.merged[r + "," + c] && !info.anchors[r + "," + c]) continue;
            g[1](info.table.getCellOrNullObject(r, c), props[r][c], fresh);
            count++;
          }
          // Send big tables in pieces so each request stays small
          if (count >= 60 && r < info.rows - 1) { count = 0; var next = r + 1; return context.sync().then(function () { return from(next); }); }
        }
        return context.sync();
      }
      return from(0).catch(function (e) {
        if (window.console) console.error("Table " + g[0] + ":", e);
        failed.push(g[0]);
      });
    }
    var size = style.props.size && style.props.size.on ? columnWidths(info.values, roles, style.props, info.merged) : null;
    function widthsStep() {
      if (!size) return Promise.resolve();
      size.widths.forEach(function (w, c) { info.table.columns.getItemAt(c).width = w; });
      return context.sync().catch(function (e) { if (window.console) console.error("Table column widths:", e); failed.push("column width"); });
    }
    function heightsStep() {
      if (!size) return Promise.resolve();
      var min = size.g * (Number(style.props.size.rowMin) || 1), step = size.g * (Number(style.props.size.rowStep) || 0.25), rows = [];
      for (var r = 0; r < info.rows; r++) { var row = info.table.rows.getItemAt(r); row.height = min; rows.push(row); }
      return context.sync().then(function () {
        rows.forEach(function (row) { row.load("currentHeight"); });
        return context.sync();
      }).then(function () {
        var changed = false;
        rows.forEach(function (row) {
          if (row.currentHeight > min + 0.5) { row.height = min + Math.ceil((row.currentHeight - min) / step - 1e-6) * step; changed = true; }
        });
        return changed ? context.sync() : null;
      }).catch(function (e) { if (window.console) console.error("Table row heights:", e); failed.push("row height"); });
    }
    var chain = settings().then(widthsStep);
    GROUPS.forEach(function (g) { chain = chain.then(function () { return group(g); }); });
    chain = chain.then(heightsStep);
    return chain.then(function () {
      info.shape.tags.add(TABLE_TAG, JSON.stringify({ s: style.id, r: normRoles(roles, info.cols).join("") }));
      return context.sync();
    }).then(function () { failed.wider = !!(size && size.wider); return failed; });
  }
  function joinWords(list) { return list.length < 2 ? list.join("") : list.slice(0, -1).join(", ") + " and " + list[list.length - 1]; }
  function refusedNote(failed) {
    var note = failed && failed.length ? " PowerPoint didn't accept the " + joinWords(failed) + " settings, so those were left as they were." : "";
    if (failed && failed.wider) note += " The columns need more room than the grid area, so the table is wider.";
    return note;
  }
  // Finds the table to work on. PowerPoint doesn't always report the table as the selected
  // shape (for example while the cursor is in a cell, or after the line check), so when no
  // shape is selected this falls back to the table the panel last saw on this slide, or the
  // only table on the slide.
  var lastTable = null, target = null;
  var NO_TABLE = "Click the table's border to select it, or put the cursor in one of its cells, then try again.";
  // A table can be its own shape, or sit inside a content placeholder (the usual result of
  // the table icon in a template's layout). Loads what's needed and returns the tables.
  function tableShapes(context, list) {
    var holders = list.filter(function (x) { return x.type === "Placeholder"; });
    holders.forEach(function (x) { try { x.placeholderFormat.load("containedType"); } catch (e) { /* older host */ } });
    var done = holders.length ? context.sync().catch(function () { /* placeholder details unavailable */ }) : Promise.resolve();
    return done.then(function () { return list.filter(isTableShape); });
  }
  function isTableShape(x) {
    if (x.type === "Table") return true;
    if (x.type !== "Placeholder") return false;
    try { return x.placeholderFormat.containedType === "Table"; } catch (e) { return false; }
  }
  function kindOf(x) {
    if (x.type !== "Placeholder") return x.type;
    var held = null;
    try { held = x.placeholderFormat.containedType; } catch (e) { held = null; }
    return held ? "Placeholder holding " + held : "Placeholder";
  }
  var lastFind = { report: "", other: null };
  function findTable(context) {
    var shapes = context.presentation.getSelectedShapes(), slides = context.presentation.getSelectedSlides(), all = context.presentation.slides;
    shapes.load("items/id,items/type");
    slides.load("items/id");
    all.load("items/id");
    var selTables = [], slideTables = [], slide = null;
    return context.sync().then(function () {
      slide = slides.items[0] || null;
      return tableShapes(context, shapes.items);
    }).then(function (t) {
      selTables = t;
      if (!slide) return [];
      slide.shapes.load("items/id,items/type");
      return context.sync().then(function () { return tableShapes(context, slide.shapes.items); });
    }).then(function (t) {
      slideTables = t;
      var n = shapes.items.length, index = slide ? all.items.map(function (x) { return x.id; }).indexOf(slide.id) : -1;
      lastFind.report = "PowerPoint reports: " +
        (n ? ST.plural(n, "selected shape") + " (" + shapes.items.map(kindOf).join(", ") + ")" : "no selected shape") +
        (slide ? ", slide " + (index + 1) + ", " + ST.plural(slideTables.length, "table") + " on the slide." : ", no slide selected.");
      lastFind.other = n && !selTables.length ? kindOf(shapes.items[0]) : null;
      if (selTables.length) { if (slide) lastTable = { slideId: slide.id, shapeId: selTables[0].id }; return selTables[0]; }
      if (n || !slide) return null; // something that isn't a table is selected
      var pin = target || lastTable;
      var pinned = pin && pin.slideId === slide.id ? slideTables.filter(function (x) { return x.id === pin.shapeId; })[0] : null;
      if (pinned) { lastTable = pin; return pinned; }
      if (slideTables.length === 1) { lastTable = { slideId: slide.id, shapeId: slideTables[0].id }; return slideTables[0]; }
      return null;
    });
  }
  function noTableMessage() {
    return (lastFind.other ? "The selected shape is a " + lastFind.other + ", not a table. " : "") + NO_TABLE;
  }
  // Resolves the table first, then runs the line check, so the check can't lose the table.
  function checkThen(task) {
    ST.run(function (context) {
      return findTable(context).then(function (shape) {
        if (!shape) { ST.toast(noTableMessage(), true); return; }
        target = lastTable;
      });
    }).then(function () { if (target) startLineCheck(task); });
  }

  // ---------- Actions ----------
  function apply(style) {
    if (!ST.needOffice() || busy) return;
    if (!canFormat()) { ST.toast("This PowerPoint version can't format existing tables. Use Insert table instead.", true); return; }
    if (!lineMode()) { checkThen(function () { apply(style); }); return; }
    busy = true;
    ST.run(function (context) {
      return findTable(context).then(function (shape) {
        if (!shape) { ST.toast(noTableMessage() + " Or use Insert table.", true); return; }
        return readTable(context, shape).then(function (info) {
          var roles = info.tag.r ? info.tag.r.split("") : (current && current.shapeId === shape.id ? current.roles : guessRoles(info.values));
          return formatTable(context, info, style, roles).then(function (failed) {
            if (!style.applied) { style.applied = clone(style.props); saveStyles(); }
            ST.toast("Formatted the table with " + style.name + "." + refusedNote(failed), failed.length > 0);
            refreshTable();
          });
        });
      });
    }).then(function () { busy = false; target = null; });
  }

  function insert() {
    if (!ST.needOffice() || busy) return;
    if (!canInsert()) { ST.toast("This PowerPoint version can't add tables from the panel.", true); return; }
    if (canFormat() && !lineMode()) { withLines(insert); return; }
    var style = styles.filter(function (s) { return s.id === Number($("tblInsertStyle").value); })[0] || styles[0];
    if (!style) return;
    var pasted = parseClipboard($("tblPaste").value), rows, cols, roles, values;
    if (pasted.length) {
      if (pasted.length > 40 || pasted[0].length > 20) { ST.toast("Pasted data can be up to 40 rows and 20 columns. This has " + pasted.length + " x " + pasted[0].length + ".", true); return; }
      values = trimmedValues(pasted).values;
      rows = values.length; cols = values[0].length;
      roles = guessRoles(values);
      if (style.props.alignDecimals !== false) values = decimalValues(values, roles).values;
    } else {
      rows = Math.max(2, Math.min(40, Math.round(ST.num(ST.state.tblRows)) || 6));
      cols = Math.max(1, Math.min(20, Math.round(ST.num(ST.state.tblCols)) || 4));
      roles = insertRoles(cols);
      var cats = 0, datas = 0;
      var header = roles.map(function (r) {
        if (r === "C") { cats++; return cats === 1 ? "Category" : "Sub-category"; }
        if (r === "L") return "Item";
        datas++; return "Value " + datas;
      });
      values = [header];
      for (var i = 1; i < rows; i++) { var row = []; for (var c = 0; c < cols; c++) row.push(""); values.push(row); }
    }
    var g = ST.calculate();
    busy = true;
    ST.run(function (context) {
      var slides = context.presentation.getSelectedSlides();
      slides.load("items/id");
      return context.sync().then(function () {
        if (!slides.items.length) { ST.toast("Go to a slide first.", true); return; }
        var slide = slides.items[0];
        var where = g.valid ? { left: g.left * PT, top: g.top * PT, width: g.maxW * PT } : {};
        // Adds the table, trying a simpler request if PowerPoint refuses the first one.
        function add(attempts) {
          var opts = attempts[0], shape = slide.shapes.addTable(rows, cols, opts);
          shape.load("id");
          return context.sync().then(function () { return { shape: shape, opts: opts }; }, function (e) {
            if (window.console) console.error("addTable:", e);
            if (attempts.length > 1) return add(attempts.slice(1));
            throw e;
          });
        }
        var attempts = [];
        // Hosts that can't format cells afterwards get the formatting with the table itself
        if (!canFormat()) {
          var oneShot = Object.assign({ values: values, specificCellProperties: insertProps(values, roles, style.props), style: "NoStyleNoGrid" }, where);
          var sz = style.props.size && style.props.size.on ? columnWidths(values, roles, style.props, {}) : null;
          if (sz) {
            oneShot.columns = sz.widths.map(function (w) { return { columnWidth: w }; });
            oneShot.rows = values.map(function () { return { rowHeight: sz.g * (Number(style.props.size.rowMin) || 1) }; });
            delete oneShot.width;
          }
          attempts.push(oneShot);
        }
        attempts.push(Object.assign({ values: values }, where), { values: values });
        return add(attempts).then(function (res) {
          var shape = res.shape, formatted = !!res.opts.specificCellProperties;
          var step = Promise.resolve([]);
          if (canFormat()) step = readTable(context, shape).then(function (info) { formatted = true; return formatTable(context, info, style, roles, true); });
          return step.then(function (failed) {
            shape.tags.add(TABLE_TAG, JSON.stringify({ s: style.id, r: roles.join("") }));
            if (ST.supports("1.5")) slide.setSelectedShapes([shape.id]);
            return context.sync().then(function () {
              if (!style.applied) { style.applied = clone(style.props); saveStyles(); }
              var placed = res.opts.left !== undefined ? " at the Layout grid area." : ".";
              if (!formatted) ST.toast("Added a " + rows + " x " + cols + " table" + placed + " This PowerPoint version couldn't format it.", true);
              else ST.toast("Added a " + rows + " x " + cols + " table" + placed + refusedNote(failed), failed.length > 0);
              refreshTable();
            });
          });
        });
      });
    }).then(function () { busy = false; });
  }

  function stackRepeats() {
    if (!ST.needOffice() || busy) return;
    if (canFormat() && !lineMode()) { checkThen(stackRepeats); return; }
    if (!canInsert()) { ST.toast("This PowerPoint version can't change table text from the panel.", true); return; }
    busy = true;
    ST.run(function (context) {
      return findTable(context).then(function (shape) {
        if (!shape) { ST.toast(noTableMessage(), true); return; }
        return readTable(context, shape).then(function (info) {
          var roles = info.tag.r ? info.tag.r.split("") : guessRoles(info.values);
          var style = byId(info.tag.s) || styles[0] || { props: defaultProps() };
          var res = stackedValues(info.values, roles, style.props);
          if (!res.changed) { ST.toast("No repeated category labels to stack."); return; }
          for (var r = 0; r < info.rows; r++) for (var c = 0; c < info.cols; c++) {
            if (res.values[r][c] !== info.values[r][c]) info.table.getCellOrNullObject(r, c).text = res.values[r][c];
          }
          return context.sync().then(function () {
            info.values = res.values;
            // A linked table is reformatted so group rules follow the new layout
            if (info.tag.s && byId(info.tag.s) && canFormat()) return formatTable(context, info, byId(info.tag.s), roles);
            return [];
          }).then(function (failed) {
            ST.toast("Stacked " + ST.plural(res.changed, "repeated label") + "." + refusedNote(failed), failed.length > 0);
            refreshTable();
          });
        });
      });
    }).then(function () { busy = false; target = null; });
  }

  function trimNow() {
    if (!ST.needOffice() || busy) return;
    if (canFormat() && !lineMode()) { checkThen(trimNow); return; }
    if (!canInsert()) { ST.toast("This PowerPoint version can't change table text from the panel.", true); return; }
    busy = true;
    ST.run(function (context) {
      return findTable(context).then(function (shape) {
        if (!shape) { ST.toast(noTableMessage(), true); return; }
        return readTable(context, shape).then(function (info) {
          return trimCells(context, info).then(function (count) {
            if (!count) { ST.toast("No extra spaces at the start or end of cells."); return; }
            var style = byId(info.tag.s);
            // A linked table is reformatted so sizes follow the cleaner text
            var step = style && canFormat() ? formatTable(context, info, style, info.tag.r ? info.tag.r.split("") : guessRoles(info.values)) : Promise.resolve([]);
            return step.then(function (failed) {
              ST.toast("Trimmed spaces in " + ST.plural(count, "cell") + "." + refusedNote(failed), failed.length > 0);
              refreshTable();
            });
          });
        });
      });
    }).then(function () { busy = false; target = null; });
  }

  function setRole(c) {
    if (!current || busy) return;
    var order = ["C", "L", "D"];
    current.roles[c] = order[(order.indexOf(current.roles[c]) + 1) % 3];
    renderRoles();
    var id = current.shapeId, roles = current.roles.slice(), styleId = current.styleId;
    ST.run(function (context) {
      return findTable(context).then(function (shape) {
        if (!shape || shape.id !== id) return;
        return readTable(context, shape).then(function (info) {
          var style = byId(styleId);
          if (style && canFormat()) return formatTable(context, info, style, roles).then(function (failed) { if (failed.length) ST.toast(refusedNote(failed).trim(), true); });
          shape.tags.add(TABLE_TAG, JSON.stringify({ s: styleId || 0, r: roles.join("") }));
          return context.sync();
        });
      });
    });
  }
  function unlink() {
    if (!current) return;
    var id = current.shapeId;
    ST.run(function (context) {
      return findTable(context).then(function (shape) {
        if (!shape || shape.id !== id) return;
        shape.tags.add(TABLE_TAG, JSON.stringify({ s: 0, r: current.roles.join("") }));
        return context.sync().then(function () { ST.toast("The table is no longer linked to a table style."); refreshTable(); });
      });
    });
  }

  function updateAll() {
    if (!ST.needOffice() || busy) return;
    if (!canFormat()) { ST.toast("This PowerPoint version can't format existing tables.", true); return; }
    if (!lineMode()) { withLines(updateAll); return; }
    busy = true;
    var done = 0, refused = [];
    ST.run(function (context) {
      var slides = context.presentation.slides;
      slides.load("items/id");
      return context.sync().then(function () {
        slides.items.forEach(function (sl) { sl.shapes.load("items/id,items/type"); });
        return context.sync();
      }).then(function () {
        var everything = [];
        slides.items.forEach(function (sl) { everything = everything.concat(sl.shapes.items); });
        return tableShapes(context, everything);
      }).then(function (tables) {
        var i = 0;
        function next() {
          if (i >= tables.length) return Promise.resolve();
          var shape = tables[i++];
          return readTable(context, shape).then(function (info) {
            var style = byId(info.tag.s);
            if (!style) return next();
            done++;
            return formatTable(context, info, style, info.tag.r ? info.tag.r.split("") : guessRoles(info.values)).then(function (failed) {
              failed.forEach(function (f) { if (refused.indexOf(f) < 0) refused.push(f); });
              return next();
            });
          });
        }
        return next();
      }).then(function () {
        styles.forEach(function (s) { s.applied = clone(s.props); });
        return saveStyles();
      }).then(function () {
        ST.toast(done ? "Updated " + ST.plural(done, "linked table") + "." + refusedNote(refused) : "No tables are linked to a table style yet.", refused.length > 0);
        refreshTable();
      });
    }).then(function () { busy = false; });
  }

  // Reads the selected table so the panel can show its column roles and style.
  function refreshTable() {
    if (!ST.inOffice || !canInsert()) { current = null; renderRoles(); return; }
    PowerPoint.run(function (context) {
      return findTable(context).then(function (shape) {
        if (!shape) { current = null; return; }
        return readTable(context, shape).then(function (info) {
          var roles = info.tag.r ? normRoles(info.tag.r.split(""), info.cols) : guessRoles(info.values || []);
          if (roles.length !== info.cols) roles = normRoles(roles, info.cols);
          current = { shapeId: shape.id, cols: info.cols, values: info.values, roles: roles, styleId: info.tag.s || 0, saved: !!info.tag.r };
        });
      });
    }).catch(function (e) {
      current = null;
      lastFind.report += " Couldn't read the table: " + ST.reason(e) + ".";
    }).then(renderRoles);
  }

  // ---------- Rendering ----------
  function colName(c) { var s = ""; c++; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
  function renderRoles() {
    var box = $("tblRoles"), status = $("tblStatus"), diag = $("tblDiag");
    if (!box) return;
    if (diag) diag.textContent = ST.inOffice ? lastFind.report : "";
    if (!current) {
      box.innerHTML = '<div class="muted-note">Select a table to see its columns</div>';
      status.hidden = true;
      return;
    }
    box.innerHTML = "";
    current.roles.forEach(function (r, c) {
      var b = document.createElement("button");
      b.className = "mini-chip role-chip role-" + r;
      var head = current.values && current.values[0] ? String(current.values[0][c] || "").trim() : "";
      b.innerHTML = "<b>" + colName(c) + "</b> " + ROLE_NAMES[r];
      b.title = (head ? head + ": " : "") + ROLE_NAMES[r] + ". Click to change.";
      b.setAttribute("data-col", c);
      b.addEventListener("click", function () { setRole(c); });
      box.appendChild(b);
    });
    var st = byId(current.styleId);
    status.hidden = false;
    status.innerHTML = st ? "Table uses <b></b>" : "Table isn't linked to a table style.";
    if (st) {
      status.querySelector("b").textContent = st.name;
      var u = document.createElement("button");
      u.className = "link-btn"; u.textContent = "Unlink";
      u.addEventListener("click", unlink);
      status.appendChild(u);
    }
  }

  function tile(P) {
    var L = P.lines, w = 30, y = [4, 10, 15, 20, 25];
    function ln(spec, y1, x1, x2) { return spec.on ? '<line x1="' + x1 + '" y1="' + y1 + '" x2="' + x2 + '" y2="' + y1 + '" stroke="' + spec.c + '" stroke-width="' + Math.max(0.5, Math.min(2, spec.w * 1.4)) + '"/>' : ""; }
    var band = P.band.on ? '<rect x="8" y="10" width="7" height="15" fill="' + P.band.color + '"/><rect x="22" y="10" width="7" height="15" fill="' + P.band.color + '"/>' : "";
    return '<svg viewBox="0 0 ' + w + ' 28" width="34" height="30">' + band +
      ln(L.header, y[1], 1, 29) + ln(L.row, y[2], 8, 29) + ln(L.group, y[3], 1, 29) + ln(L.bottom, y[4], 1, 29) +
      (L.vertical.on ? '<line x1="22" y1="2" x2="22" y2="25" stroke="' + L.vertical.c + '" stroke-width="0.5"/>' : "") + "</svg>";
  }
  function summary(P) {
    return [P.fonts.body.name || "Body font", P.band.on ? "bands " + P.band.color : "no bands", "padding " + P.pad.all + " pt"].join(", ");
  }

  function render() {
    var list = $("tblList");
    if (!list) return;
    list.innerHTML = "";
    if (!stylesLoaded && ST.inOffice) { list.innerHTML = '<div class="muted-note">Reading table styles from this presentation.</div>'; return; }
    styles.forEach(function (s) {
      var row = document.createElement("div");
      row.className = "style-row";
      row.innerHTML = '<button class="style-apply" title="Format the selected table with this style"><span class="shape-tile tbl-tile"></span>' +
        '<span class="style-text"><span class="style-name"></span><span class="style-sum" data-min="2"></span></span>' +
        (s.applied && pending(s) ? '<span class="badge" title="Edited since the last Update all">Edited</span>' : "") +
        '</button><button class="icon-btn" data-min="2" title="Edit style" aria-label="Edit style">' + EDIT_ICON + "</button>";
      row.querySelector(".tbl-tile").innerHTML = tile(s.props);
      row.querySelector(".style-name").textContent = s.name;
      row.querySelector(".style-sum").textContent = summary(s.props);
      row.querySelector(".style-apply").addEventListener("click", function () {
        var sid = s.id;
        ST.setRepeat("Apply table style " + s.name, function () { var cur = styles.filter(function (x) { return x.id === sid; })[0]; if (cur) apply(cur); });
        apply(s);
      });
      row.querySelector(".icon-btn").addEventListener("click", function () { editing = editing === s.id ? null : s.id; render(); });
      list.appendChild(row);
      if (editing === s.id) {
        var holder = document.createElement("div");
        holder.innerHTML = editorMarkup(s);
        list.appendChild(holder);
        wireEditor(s);
      }
    });
    var sel = $("tblInsertStyle"), keep = sel.value;
    sel.innerHTML = styles.map(function (s) { return '<option value="' + s.id + '">' + esc(s.name) + "</option>"; }).join("");
    if (keep && byId(Number(keep))) sel.value = keep;
    var n = styles.filter(function (s) { return s.applied && pending(s); }).length;
    $("tblUpdate").textContent = n ? "Update all (" + n + " edited)" : "Update all";
    renderRoles();
  }

  // ---------- Editor ----------
  function textStyleSelect(id, value) {
    var list = ST.textStyles ? ST.textStyles.list() : [];
    return '<select id="' + id + '"><option value="">Set below</option>' + list.map(function (t) {
      return '<option value="' + t.id + '"' + (String(value) === String(t.id) ? " selected" : "") + ">" + esc(t.name) + "</option>";
    }).join("") + "</select>";
  }
  function onOff(id, v) { return '<select id="' + id + '"><option value="1"' + (v ? " selected" : "") + '>On</option><option value="0"' + (v ? "" : " selected") + ">Off</option></select>"; }
  function colorField(id, value) {
    return '<div class="color-field"><input type="color" id="' + id + 'Pick" list="paletteList" value="' + esc(String(value || "#000000").toLowerCase()) + '"><input type="text" id="' + id + '" value="' + esc(value) + '"></div>';
  }
  function editorMarkup(s) {
    var P = s.props, html = '<div class="style-editor tbl-editor">' +
      '<label class="field">Style name<input type="text" id="tbName" maxlength="32" value="' + esc(s.name) + '"></label>';
    html += '<div class="group-label label-row">Fonts' + ST.info("A chosen text style overrides the font settings under it.", "About table fonts") + '</div>';
    PARTS.forEach(function (p) {
      var f = P.fonts[p[0]], k = "tbF_" + p[0];
      html += '<div class="sublabel">' + p[1] + '</div><div class="prop-grid">' +
        '<label>Text style' + textStyleSelect(k + "_ts", f.ts) + '</label>' +
        '<label>Font<input type="text" id="' + k + '_name" value="' + esc(f.name) + '" list="fontList"></label>' +
        '<label>Size (pt)<input type="text" id="' + k + '_size" inputmode="decimal" value="' + esc(f.size) + '"></label>' +
        '<label>Colour' + colorField(k + "_color", f.color) + '</label>' +
        '<label>Bold' + onOff(k + "_bold", f.bold) + '</label>' +
        '<label>Italic' + onOff(k + "_italic", f.italic) + '</label></div>';
    });
    html += '<div class="group-label">Lines</div>';
    LINES.forEach(function (l) {
      var v = P.lines[l[0]], k = "tbL_" + l[0];
      html += '<div class="sublabel">' + l[1] + '</div><div class="prop-grid tbl-line">' +
        '<label>Show' + onOff(k + "_on", v.on) + '</label>' +
        '<label>Weight (pt)<input type="text" id="' + k + '_w" inputmode="decimal" value="' + esc(v.w) + '"></label>' +
        '<label class="span2">Colour' + colorField(k + "_c", v.c) + '</label></div>';
    });
    html += '<div class="group-label label-row">Column fill' + ST.info("Body rows only, counted from the table's left edge.", "About column fill") + '</div><div class="prop-grid">' +
      '<label>Fill 2nd, 4th, 6th... columns' + onOff("tbBandOn", P.band.on) + '</label>' +
      '<label>Colour' + colorField("tbBandC", P.band.color) + '</label></div>';
    html += '<div class="group-label">Padding inside cells (pt)</div><div class="prop-grid">' +
      '<label>All sides<input type="text" id="tbPadAll" inputmode="decimal" value="' + esc(P.pad.all) + '"></label>' +
      '<label>First column, left<input type="text" id="tbPadFirst" inputmode="decimal" value="' + esc(P.pad.firstLeft) + '"></label>' +
      '<label>Last column, right<input type="text" id="tbPadLast" inputmode="decimal" value="' + esc(P.pad.lastRight) + '"></label></div>';
    var Z = P.size || defaultProps().size;
    html += '<div class="group-label label-row">Sizes' + ST.info("g is the Layout spacing. Columns are whole steps of g and the table fills the grid area. Label columns get the room left over.", "About table sizes") + '</div><div class="prop-grid">' +
      '<label>Grid-based sizes' + onOff("tbSizeOn", Z.on) + '</label>' +
      '<label>Fit column widths' + onOff("tbSizeFit", Z.fitColumns) + '</label>' +
      '<label>Minimum row height (g)<input type="text" id="tbRowMin" inputmode="decimal" value="' + esc(Z.rowMin) + '"></label>' +
      '<label>Row height step (g)<input type="text" id="tbRowStep" inputmode="decimal" value="' + esc(Z.rowStep) + '"></label></div>';
    html += '<label class="field">With one category column, thin lines between groups' + onOff("tbThinSingle", P.thinSingleGroups !== false) + '</label>';
    html += '<label class="field">Trim spaces at the start and end of cells' + onOff("tbTrim", P.trimSpaces !== false) + '</label>';
    html += '<label class="field">Line up decimals in data columns' + onOff("tbDecimals", P.alignDecimals !== false) + '</label>';
    html += '<label class="field">Total rows start with<input type="text" id="tbTotal" value="' + esc(P.totalPrefix) + '"></label>';
    html += '<div class="btn-row"><button class="btn" id="tbCancel">Cancel</button><button class="btn primary" id="tbSave">Save</button></div>' +
      '<div class="btn-row"><button class="btn small" id="tbDup">Duplicate</button><button class="btn small danger" id="tbDel">Delete</button></div></div>';
    return html;
  }
  function wireEditor(s) {
    Array.prototype.forEach.call(document.querySelectorAll(".tbl-editor input[type=color]"), function (pick) {
      var text = $(pick.id.replace(/Pick$/, ""));
      pick.addEventListener("input", function () { text.value = pick.value.toUpperCase(); });
      text.addEventListener("change", function () { if (/^#?[0-9a-f]{6}$/i.test(text.value.trim())) pick.value = ("#" + text.value.trim().replace("#", "")).toLowerCase(); });
    });
    $("tbCancel").addEventListener("click", function () { editing = null; render(); });
    $("tbSave").addEventListener("click", function () { saveEditor(s); });
    $("tbDup").addEventListener("click", function () {
      var copy = { id: nextId(), name: (s.name + " copy").slice(0, 32), props: clone(s.props), applied: null };
      styles.push(copy); editing = copy.id; saveStyles();
    });
    $("tbDel").addEventListener("click", function () {
      if (styles.length === 1) { ST.toast("Keep at least one table style.", true); return; }
      styles = styles.filter(function (x) { return x !== s; }); editing = null; saveStyles();
      ST.toast("Deleted " + s.name + ". Tables that used it keep their formatting.");
    });
  }
  function num(id, fallback) { var v = ST.evaluate($(id).value, null, null); return v === null || v < 0 ? fallback : Math.round(v * 100) / 100; }
  function hex(id, fallback) { var v = $(id).value.trim(); return /^#?[0-9a-f]{6}$/i.test(v) ? ("#" + v.replace("#", "")).toUpperCase() : fallback; }
  function saveEditor(s) {
    var P = clone(s.props);
    var name = $("tbName").value.trim();
    if (!name) { ST.toast("Give the style a name.", true); return; }
    PARTS.forEach(function (p) {
      var k = "tbF_" + p[0], f = P.fonts[p[0]];
      f.ts = Number($(k + "_ts").value) || null;
      f.name = $(k + "_name").value.trim();
      f.size = num(k + "_size", f.size);
      f.color = hex(k + "_color", f.color);
      f.bold = $(k + "_bold").value === "1";
      f.italic = $(k + "_italic").value === "1";
    });
    LINES.forEach(function (l) {
      var k = "tbL_" + l[0], v = P.lines[l[0]];
      v.on = $(k + "_on").value === "1";
      v.w = num(k + "_w", v.w) || 0.25;
      v.c = hex(k + "_c", v.c);
    });
    P.band.on = $("tbBandOn").value === "1";
    P.band.color = hex("tbBandC", P.band.color);
    P.pad.all = num("tbPadAll", P.pad.all);
    P.pad.firstLeft = num("tbPadFirst", P.pad.firstLeft);
    P.pad.lastRight = num("tbPadLast", P.pad.lastRight);
    P.totalPrefix = $("tbTotal").value.trim();
    P.thinSingleGroups = $("tbThinSingle").value === "1";
    P.trimSpaces = $("tbTrim").value === "1";
    P.alignDecimals = $("tbDecimals").value === "1";
    P.size = P.size || defaultProps().size;
    P.size.on = $("tbSizeOn").value === "1";
    P.size.fitColumns = $("tbSizeFit").value === "1";
    P.size.rowMin = num("tbRowMin", P.size.rowMin) || 1;
    P.size.rowStep = num("tbRowStep", P.size.rowStep) || 0.25;
    s.name = name; s.props = P; editing = null;
    saveStyles();
    ST.toast(s.applied && pending(s) ? "Saved. Click Update all to change the tables that use it." : "Saved " + s.name + ".");
  }
  function newStyle() {
    var s = { id: nextId(), name: "Table style " + nextId(), props: defaultProps(), applied: null };
    styles.push(s); editing = s.id; render();
  }

  // ---------- Build ----------
  var MARKUP =
    '<div class="line-check" id="tblCheck" hidden><div class="check-q"></div>' +
      '<div class="check-btns"><button class="btn" data-ans="A">A</button><button class="btn" data-ans="B">B</button><button class="btn small" data-ans="none" title="Neither cell, or both cells, have a thick line">Not sure</button></div></div>' +
    '<div class="style-status" id="tblStatus" hidden></div>' +
    '<div class="sublabel label-row">Columns' + ST.info("Click a chip to switch it. <b>Category</b> stacked labels such as a year. <b>Label</b> row names. <b>Data</b> numbers, right aligned.", "About column roles") + '</div>' +
    '<div class="mini-chips" id="tblRoles"></div>' +
    '<div class="help tbl-diag" id="tblDiag" data-min="2"></div>' +
    '<div class="btn-row"><button class="btn small" id="tblStack" title="Blank category labels that repeat the row above">Stack repeats</button>' +
      '<button class="btn small" id="tblTrim" title="Remove spaces at the start and end of every cell">Trim spaces</button></div>' +
    '<div class="group-label">Table styles</div>' +
    '<div class="style-list" id="tblList"></div>' +
    '<div class="btn-row" data-min="2"><button class="btn small" id="tblNew">New style</button><button class="btn small" id="tblUpdate" title="Reformat every table linked to a style">Update all</button></div>' +
    '<button class="btn small" id="tblCheckBtn" data-min="2" title="Check how this PowerPoint shows table lines">Check lines</button>' +
    '<div class="group-label">Insert table</div>' +
    '<label class="field"><span class="label-row">Paste data' + ST.info("Copy cells in Excel and paste them here. Insert table then builds the table from them: the first row becomes the header, and column roles are worked out from the values. Leave it empty for a blank table of the size below.", "About pasting data") + '</span><textarea id="tblPaste" class="paste-box" rows="3" placeholder="Paste cells from Excel here"></textarea></label>' +
    '<div class="paste-info" id="tblPasteInfo" hidden><span id="tblPasteSize"></span><button class="link-btn" id="tblPasteClear">Clear</button></div>' +
    '<div class="row"><label class="field">Rows<input type="text" id="tblRows" inputmode="numeric"></label>' +
      '<label class="field">Columns<input type="text" id="tblCols" inputmode="numeric"></label></div>' +
    '<label class="field" data-min="2">Style<select id="tblInsertStyle"></select></label>' +
    '<button class="btn primary" id="tblInsert">Insert table</button>';

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    ST.bindField($("tblRows"), "tblRows", { number: true });
    ST.bindField($("tblCols"), "tblCols", { number: true });
    $("tblInsert").addEventListener("click", insert);
    var showPaste = function () {
      var data = parseClipboard($("tblPaste").value), has = data.length > 0;
      $("tblPasteInfo").hidden = !has;
      $("tblRows").disabled = has; $("tblCols").disabled = has;
      if (has) {
        $("tblRows").value = data.length; $("tblCols").value = data[0].length;
        $("tblPasteSize").textContent = data.length + " rows x " + data[0].length + " columns, first row as the header";
      } else { $("tblRows").value = ST.state.tblRows; $("tblCols").value = ST.state.tblCols; }
    };
    $("tblPaste").addEventListener("input", showPaste);
    $("tblPasteClear").addEventListener("click", function () { $("tblPaste").value = ""; showPaste(); });
    $("tblCheckBtn").addEventListener("click", function () { startLineCheck(null); });
    Array.prototype.forEach.call(body.querySelectorAll("#tblCheck [data-ans]"), function (b) {
      b.addEventListener("click", function () { answerLineCheck(b.getAttribute("data-ans")); });
    });
    $("tblStack").addEventListener("click", stackRepeats);
    $("tblTrim").addEventListener("click", trimNow);
    $("tblNew").addEventListener("click", newStyle);
    $("tblUpdate").addEventListener("click", updateAll);
    ST.on("ready", loadStyles);
    ST.on("selection", refreshTable);
    ST.on("textstyles", function () { if (!editing) render(); });
    render();
  }

  ST.registerTool({
    id: "tables",
    title: "Tables",
    icon: ICON,
    levels: [1, 2],
    defaultLevel: 1,
    defaults: defaults,
    tip: { what: "Formats a table in one click.", steps: ["Select a table.", "Check the column chips.", "Click a style, or use Insert table."] },
    build: build
  });

  // Exposed for testing only
  ST._tables = { decimalValues: decimalValues, parseClipboard: parseClipboard, trimmedValues: trimmedValues, columnWidths: columnWidths, lineMode: lineMode, cellProps: cellProps, stackedValues: stackedValues, guessRoles: guessRoles, defaultProps: defaultProps,
    get: function () { return styles; }, current: function () { return current; }, put: function (s) { styles = s; stylesLoaded = true; return saveStyles(); } };
})();
