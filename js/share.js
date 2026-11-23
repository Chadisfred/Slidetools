/* Shared dialogs: share settings as a code, confirm before a slide swap, and the
   top bar Repeat button. Everything stays in the panel: the settings code is plain
   text the person copies and pastes themselves. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;
  var SWAP_KEY = "slideToolsSwapOk", PREFIX = "SLIDETOOLS1.";

  // ---------- Sheet: a small dialog over the panel ----------
  var current = null;
  function closeSheet() { if (current) { current.remove(); current = null; document.removeEventListener("keydown", onKey); } }
  function onKey(e) { if (e.key === "Escape") closeSheet(); }
  ST.sheet = function (title, body) {
    closeSheet();
    var wrap = document.createElement("div");
    wrap.className = "sheet-wrap";
    wrap.innerHTML = '<div class="sheet" role="dialog" aria-modal="true"><div class="sheet-head"><h2></h2><button class="icon-btn sheet-close" title="Close" aria-label="Close">' +
      '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"><path d="M4 4l8 8M12 4l-8 8"/></svg></button></div><div class="sheet-body"></div></div>';
    wrap.querySelector("h2").textContent = title;
    wrap.querySelector(".sheet-body").appendChild(body);
    wrap.querySelector(".sheet-close").addEventListener("click", closeSheet);
    wrap.addEventListener("click", function (e) { if (e.target === wrap) closeSheet(); });
    document.body.appendChild(wrap);
    document.addEventListener("keydown", onKey);
    current = wrap;
    var first = wrap.querySelector("textarea, input, button.btn");
    if (first) first.focus();
    return { close: closeSheet, el: wrap };
  };

  // ---------- Confirm before replacing a slide ----------
  function swapOk() { try { return localStorage.getItem(SWAP_KEY) === "1"; } catch (e) { return false; } }
  // Resolves true to go ahead. what: "crop the pictures" or "add the copies".
  ST.confirmSwap = function (what) {
    if (swapOk()) return Promise.resolve(true);
    return new Promise(function (resolve) {
      var body = ST.html(
        '<p class="sheet-text">PowerPoint add-ins can\'t ' + what + ' directly, so Slide Tools replaces this slide with an edited copy. The copy is made in this panel and nothing leaves your computer.</p>' +
        '<ul class="sheet-list"><li>Comments on this slide may not carry over</li><li>Undo takes more than one step</li></ul>' +
        '<label class="sheet-check"><input type="checkbox" id="swapAgain"> Don\'t ask again</label>' +
        '<div class="btn-row"><button class="btn" id="swapCancel">Cancel</button><button class="btn primary" id="swapGo">Continue</button></div>');
      var done = false, s = ST.sheet("Replace this slide?", body);
      var finish = function (ok) {
        if (done) return; done = true;
        if (ok && $("swapAgain").checked) { try { localStorage.setItem(SWAP_KEY, "1"); } catch (e) { /* storage unavailable */ } }
        s.close(); resolve(ok);
      };
      $("swapCancel").addEventListener("click", function () { finish(false); });
      $("swapGo").addEventListener("click", function () { finish(true); });
      // Closing the sheet another way counts as Cancel
      new MutationObserver(function (m, obs) { if (!document.body.contains(s.el)) { obs.disconnect(); finish(false); } }).observe(document.body, { childList: true });
    });
  };
  ST.resetSwapConfirm = function () { try { localStorage.removeItem(SWAP_KEY); } catch (e) { /* storage unavailable */ } };

  // ---------- Repeat last action ----------
  var repeat = null;
  ST.setRepeat = function (label, run) {
    repeat = { label: label, run: run };
    var b = $("repeatBtn");
    if (b) { b.disabled = false; b.title = "Repeat: " + label; b.setAttribute("aria-label", "Repeat " + label); }
  };
  ST.runRepeat = function () { if (repeat) repeat.run(); };

  // ---------- Settings code ----------
  function clone(x) { return JSON.parse(JSON.stringify(x)); }
  function key(name) { return String(name || "").trim().toLowerCase(); }
  function nextId(list) { return list.reduce(function (m, s) { return Math.max(m, +s.id || 0); }, 0) + 1; }

  // Each part: what it's called, how to read it and how to store it.
  function parts() {
    var out = [];
    if (ST._styles) out.push({ id: "textStyles", label: "Text styles", get: function () { return ST._styles.get().map(function (s) { return { id: s.id, name: s.name, font: s.font }; }); } });
    if (ST._shapeStyles) out.push({ id: "shapeStyles", label: "Shape styles", get: function () { return ST._shapeStyles.get().map(function (s) { return { id: s.id, name: s.name, props: s.props }; }); } });
    if (ST._tables) out.push({ id: "tableStyles", label: "Table styles", get: function () { return ST._tables.get().map(function (s) { return { id: s.id, name: s.name, props: s.props }; }); } });
    if (ST._colours) out.push({ id: "colours", label: "Colours", get: function () { return ST._colours.get().map(function (c) { return { hex: c.hex, name: c.name || "" }; }); } });
    out.push({ id: "presets", label: "Layout presets", get: function () { return (ST.state.customPresets || []).map(function (p) { return clone(p); }); } });
    return out;
  }
  function encode(data) {
    var bytes = new TextEncoder().encode(JSON.stringify(data)), s = "";
    for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return PREFIX + btoa(s);
  }
  function decode(text) {
    var t = String(text || "").replace(/\s+/g, "");
    var at = t.indexOf(PREFIX);
    if (at < 0) return null;
    try {
      var bin = atob(t.slice(at + PREFIX.length)), bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      var data = JSON.parse(new TextDecoder().decode(bytes));
      return data && data.v === 1 ? data : null;
    } catch (e) { return null; }
  }
  ST.settingsCode = function (which) {
    var data = { v: 1 };
    parts().forEach(function (p) { if (!which || which.indexOf(p.id) >= 0) data[p.id] = p.get(); });
    return encode(data);
  };

  // Merges or replaces each chosen part. Resolves a summary per part.
  function importData(data, which, replace) {
    var summary = [], idMap = {}, jobs = [];
    var mergeNamed = function (have, incoming, make, update) {
      var list = replace ? [] : have.map(clone), added = 0, updated = 0;
      incoming.forEach(function (item) {
        var hit = list.filter(function (x) { return key(x.name) === key(item.name); })[0];
        if (hit) { update(hit, item); updated++; } else { list.push(make(item, nextId(list))); added++; }
      });
      return { list: list, added: added, updated: updated };
    };
    if (which.indexOf("textStyles") >= 0 && data.textStyles && ST._styles) {
      var r1 = mergeNamed(ST._styles.get(), data.textStyles,
        function (s, id) { idMap[s.id] = id; return { id: id, name: s.name, font: clone(s.font), applied: clone(s.font) }; },
        function (h, s) { idMap[s.id] = h.id; h.font = clone(s.font); });
      jobs.push(ST._styles.put(r1.list)); summary.push(["Text styles", r1]);
    }
    if (which.indexOf("shapeStyles") >= 0 && data.shapeStyles && ST._shapeStyles) {
      var remap = function (props) {
        var p = clone(props);
        if (p.textStyleId !== undefined && p.textStyleId !== null) p.textStyleId = idMap[p.textStyleId] !== undefined ? idMap[p.textStyleId] : null;
        return p;
      };
      var r2 = mergeNamed(ST._shapeStyles.get(), data.shapeStyles,
        function (s, id) { var p = remap(s.props); return { id: id, name: s.name, props: p, applied: clone(p) }; },
        function (h, s) { h.props = remap(s.props); });
      jobs.push(ST._shapeStyles.put(r2.list)); summary.push(["Shape styles", r2]);
    }
    if (which.indexOf("tableStyles") >= 0 && data.tableStyles && ST._tables) {
      var r3 = mergeNamed(ST._tables.get(), data.tableStyles,
        function (s, id) { return { id: id, name: s.name, props: clone(s.props), applied: null }; },
        function (h, s) { h.props = clone(s.props); });
      jobs.push(ST._tables.put(r3.list)); summary.push(["Table styles", r3]);
    }
    if (which.indexOf("colours") >= 0 && data.colours && ST._colours) {
      var have = replace ? [] : ST._colours.get().map(clone), added = 0;
      data.colours.forEach(function (c) {
        if (!have.some(function (x) { return String(x.hex).toUpperCase() === String(c.hex).toUpperCase(); })) { have.push({ hex: c.hex, name: c.name || "" }); added++; }
      });
      jobs.push(ST._colours.put(have)); summary.push(["Colours", { added: added, updated: 0 }]);
    }
    if (which.indexOf("presets") >= 0 && data.presets) {
      var stamp = Date.now();
      var r5 = mergeNamed(ST.state.customPresets || [], data.presets,
        function (p, id) { var n = clone(p); n.id = "u" + stamp + "_" + id; n.user = true; return n; },
        function (h, p) { Object.keys(p).forEach(function (k) { if (k !== "id") h[k] = clone(p[k]); }); });
      ST.state.customPresets = r5.list; ST.changed(); ST.emit("reset");
      summary.push(["Layout presets", r5]);
    }
    return Promise.all(jobs).then(function () { return summary; });
  }

  function partRows(prefix, counts) {
    return '<div class="sheet-checks">' + counts.map(function (c) {
      return '<label class="sheet-check"><input type="checkbox" id="' + prefix + c.id + '"' + (c.n ? " checked" : " disabled") + '> ' + c.label + ' <span class="muted">' + c.n + '</span></label>';
    }).join("") + '</div>';
  }
  function checked(prefix, counts) { return counts.filter(function (c) { var b = $(prefix + c.id); return b && b.checked; }).map(function (c) { return c.id; }); }

  ST.openShare = function () {
    var counts = parts().map(function (p) { return { id: p.id, label: p.label, n: p.get().length }; });
    var body = ST.html(
      '<div class="seg seg2" id="shareTabs"><button data-v="export" class="on">Export</button><button data-v="import">Import</button></div>' +
      '<div id="shareExport" class="sheet-pane">' +
        '<p class="sheet-text">Copy a settings code and send it to someone. They choose Import and paste it. Styles and colours come from this presentation, and Layout presets from this computer.</p>' +
        partRows("shx_", counts) +
        '<textarea id="shareCode" class="sheet-code" readonly rows="4" aria-label="Settings code"></textarea>' +
        '<button class="btn primary" id="shareCopy">Copy code</button>' +
      '</div>' +
      '<div id="shareImport" class="sheet-pane" hidden>' +
        '<p class="sheet-text">Paste a settings code. Styles and colours are added to this presentation, and Layout presets to this computer.</p>' +
        '<textarea id="shareIn" class="sheet-code" rows="4" placeholder="SLIDETOOLS1." aria-label="Paste a settings code"></textarea>' +
        '<div id="shareFound"></div>' +
      '</div>');
    var sheet = ST.sheet("Share settings", body);
    var refresh = function () { $("shareCode").value = ST.settingsCode(checked("shx_", counts)); };
    counts.forEach(function (c) { var b = $("shx_" + c.id); if (b) b.addEventListener("change", refresh); });
    refresh();
    Array.prototype.forEach.call($("shareTabs").children, function (b) {
      b.addEventListener("click", function () {
        Array.prototype.forEach.call($("shareTabs").children, function (x) { x.classList.toggle("on", x === b); });
        $("shareExport").hidden = b.getAttribute("data-v") !== "export";
        $("shareImport").hidden = b.getAttribute("data-v") !== "import";
      });
    });
    $("shareCopy").addEventListener("click", function () {
      var text = $("shareCode").value, done = function () { ST.toast("Settings code copied. Paste it into an email or message."); };
      var fallback = function () { var t = $("shareCode"); t.focus(); t.select(); try { document.execCommand("copy"); done(); } catch (e) { ST.toast("Select the code and press Ctrl+C to copy it.", true); } };
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(done, fallback); else fallback();
    });
    $("shareIn").addEventListener("input", function () {
      var data = decode($("shareIn").value), box = $("shareFound");
      box.innerHTML = "";
      if (!$("shareIn").value.trim()) return;
      if (!data) { box.appendChild(ST.html('<p class="calc bad">That isn\'t a Slide Tools settings code. Check it was copied in full.</p>')); return; }
      var found = parts().map(function (p) { return { id: p.id, label: p.label, n: (data[p.id] || []).length }; });
      box.appendChild(ST.html(
        '<div class="sheet-checks"><div class="group-label">In this code</div></div>' + partRows("shi_", found) +
        '<div class="seg seg2" id="shareMode"><button data-v="merge" class="on" title="Add new items and update ones with the same name">Merge</button><button data-v="replace" title="Replace your lists with the ones in the code">Replace</button></div>' +
        '<p class="sheet-text" id="shareModeNote">Adds new items and updates ones with the same name. Nothing else is removed.</p>' +
        '<button class="btn primary" id="shareGo">Import</button>'));
      var mode = "merge";
      Array.prototype.forEach.call($("shareMode").children, function (b) {
        b.addEventListener("click", function () {
          mode = b.getAttribute("data-v");
          Array.prototype.forEach.call($("shareMode").children, function (x) { x.classList.toggle("on", x === b); });
          $("shareModeNote").textContent = mode === "merge" ? "Adds new items and updates ones with the same name. Nothing else is removed." : "Replaces each chosen list with the one in the code. Items not in the code are removed.";
        });
      });
      $("shareGo").addEventListener("click", function () {
        var which = checked("shi_", found);
        if (!which.length) { ST.toast("Choose at least one part to import.", true); return; }
        $("shareGo").disabled = true;
        importData(data, which, mode === "replace").then(function (summary) {
          sheet.close();
          ST.toast("Imported " + summary.map(function (s) { var r = s[1]; return s[0].toLowerCase() + " (" + r.added + " new" + (r.updated ? ", " + r.updated + " updated" : "") + ")"; }).join(", ") + ".");
        }, function (e) { $("shareGo").disabled = false; ST.toast("Couldn't import: " + ST.reason(e), true); });
      });
    });
  };

  ST._share = { encode: encode, decode: decode, importData: importData, parts: parts };
})();
