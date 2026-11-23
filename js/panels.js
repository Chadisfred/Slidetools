/* Panel framework: builds a panel for each registered tool, with detail levels, folding,
   drag to reorder, hiding and a saved arrangement.

   Two views:
   Launcher (default)  a grid of tool buttons at the top. Clicking one opens that tool and
                       closes the others, except pinned tools, which stay open.
   Stacked             every tool in one scrolling list, each folded or unfolded. */
(function () {
  "use strict";

  var ST = window.ST;
  var LAYOUT_KEY = "slideToolsPanels";
  var LEVEL_NAMES = { 1: "Basic", 2: "More", 3: "Full" };

  var ICONS = {
    grip: '<svg viewBox="0 0 10 16" fill="currentColor"><circle cx="3" cy="3" r="1.3"/><circle cx="7" cy="3" r="1.3"/><circle cx="3" cy="8" r="1.3"/><circle cx="7" cy="8" r="1.3"/><circle cx="3" cy="13" r="1.3"/><circle cx="7" cy="13" r="1.3"/></svg>',
    chev: '<svg class="chev" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6l4 4 4-4"/></svg>',
    more: '<svg viewBox="0 0 16 16" fill="currentColor"><circle cx="3.5" cy="8" r="1.4"/><circle cx="8" cy="8" r="1.4"/><circle cx="12.5" cy="8" r="1.4"/></svg>',
    check: '<svg viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 8.5l3 3 7-7"/></svg>',
    blank: '<svg viewBox="0 0 16 16" width="14" height="14"></svg>',
    close: '<svg class="close-x" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"><path d="M4.5 4.5l7 7M11.5 4.5l-7 7"/></svg>',
    pin: '<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M6 2.5h4M7 2.5v4L4.5 9h7L9 6.5v-4M8 9v4.5"/></svg>'
  };
  // The first order, grouped by job: placing shapes, text and colour, then content.
  var DEFAULT_ORDER = ["layout", "arrange", "resize", "copies", "typography", "colours", "styles", "shapestyles", "images", "tables", "corners", "match"];
  var LAYOUT_VERSION = 2;
  // What a tool's launcher button is called, when that differs from its title.
  var SHORT = { styles: "Text styles", match: "Similar" };
  ST.ICONS = ICONS;

  var layout = { order: [], levels: {}, folded: {}, hidden: {}, pinned: {}, open: "layout", stacked: false, noAutoOpen: false, v: LAYOUT_VERSION };
  var lastPick = 0; // when a launcher button was last clicked, so the selection doesn't override it
  var cards = {};
  var openMenu = null;

  function loadLayout() {
    var saved = null;
    try { saved = JSON.parse(localStorage.getItem(LAYOUT_KEY) || "null"); } catch (e) { saved = null; }
    if (saved) Object.keys(layout).forEach(function (k) { if (saved[k] !== undefined) layout[k] = saved[k]; });
    var ids = ST.tools.map(function (t) { return t.id; });
    // Arrangements saved before the launcher get the new grouped order once.
    if (!saved || !(saved.v >= LAYOUT_VERSION)) { layout.order = defaultOrder(); layout.v = LAYOUT_VERSION; }
    if (layout.open && ids.indexOf(layout.open) < 0) layout.open = null;
    // Keep saved order, then add any tools that are new since last time.
    layout.order = layout.order.filter(function (id) { return ids.indexOf(id) >= 0; });
    defaultOrder().forEach(function (id) { if (layout.order.indexOf(id) < 0) layout.order.push(id); });
  }
  function saveLayout() {
    try { localStorage.setItem(LAYOUT_KEY, JSON.stringify(layout)); } catch (e) { /* storage unavailable */ }
  }
  function defaultOrder() {
    var ids = ST.tools.map(function (t) { return t.id; });
    return DEFAULT_ORDER.filter(function (id) { return ids.indexOf(id) >= 0; }).concat(ids.filter(function (id) { return DEFAULT_ORDER.indexOf(id) < 0; }));
  }
  function tool(id) { return ST.tools.filter(function (t) { return t.id === id; })[0]; }
  function levelsOf(t) { return t.levels || [1]; }
  function levelOf(t) {
    var lv = layout.levels[t.id], allowed = levelsOf(t);
    return allowed.indexOf(lv) >= 0 ? lv : (t.defaultLevel || allowed[0]);
  }

  // ---------- Building ----------
  function buildCard(t) {
    var card = document.createElement("section");
    card.className = "tool";
    card.setAttribute("data-tool", t.id);
    var single = levelsOf(t).length < 2;
    card.innerHTML =
      '<div class="tool-head">' +
        '<button class="grip" title="Drag to reorder" aria-label="Drag to reorder">' + ICONS.grip + '</button>' +
        '<span class="tool-icon">' + t.icon + '</span>' +
        '<button class="tool-title" title="Fold or unfold">' + t.title + '</button>' +
        (single ? "" : '<button class="level-btn" title="Show more or less detail"><span class="level-dots"></span><span class="level-name"></span></button>') +
        '<button class="icon-btn pin-btn" title="Pin: keep this tool open when you open another" aria-label="Pin this tool" aria-pressed="false">' + ICONS.pin + '</button>' +
        '<button class="icon-btn menu-btn" title="Panel options" aria-label="Panel options">' + ICONS.more + '</button>' +
        '<button class="icon-btn fold-btn" title="Fold or unfold" aria-label="Fold or unfold">' + ICONS.chev + ICONS.close + '</button>' +
      '</div>' +
      '<div class="tool-body"></div>';
    var body = card.querySelector(".tool-body");
    ST.$("stack").appendChild(card); // tools look up their boxes by id, so attach first
    t.build(body, card);
    // A tool's i tip lives in its header, so it never covers the tool's first row. Tools
    // describe it as { what: "one line", steps: ["...", "..."] }.
    if (t.tip && !body.querySelector(".res-info")) {
      body.insertAdjacentHTML("afterbegin", '<span class="info res-info" tabindex="0" role="button" aria-label="What ' + t.title + ' does">i' +
        '<span class="tip" role="tooltip"><b class="tip-what"></b><ol class="tip-steps"></ol></span></span>');
      var tipBox = body.querySelector(".res-info .tip");
      tipBox.querySelector(".tip-what").textContent = t.tip.what;
      (t.tip.steps || []).forEach(function (step) { var li = document.createElement("li"); li.textContent = step; tipBox.querySelector(".tip-steps").appendChild(li); });
    }
    var info = body.querySelector(".res-info");
    if (info) {
      var head = card.querySelector(".tool-head");
      head.insertBefore(info, card.querySelector(".level-btn") || card.querySelector(".menu-btn"));
      info.addEventListener("click", function (e) { e.stopPropagation(); });
    }

    card.querySelector(".tool-title").addEventListener("click", function () { toggleFold(t.id); });
    card.querySelector(".fold-btn").addEventListener("click", function () { toggleFold(t.id); });
    if (!single) card.querySelector(".level-btn").addEventListener("click", function () { cycleLevel(t.id); });
    card.querySelector(".menu-btn").addEventListener("click", function (e) { e.stopPropagation(); showToolMenu(t.id, e.currentTarget); });
    card.querySelector(".pin-btn").addEventListener("click", function () { togglePin(t.id); });
    wireDrag(card, card.querySelector(".grip"));
    cards[t.id] = card;
    return card;
  }

  // In the launcher view a tool shows when it's the open one or pinned.
  function shown(id) {
    if (layout.hidden[id]) return false;
    return layout.stacked ? true : (layout.open === id || !!layout.pinned[id]);
  }
  function applyCard(t) {
    var card = cards[t.id], lv = levelOf(t);
    card.setAttribute("data-level", lv);
    card.classList.toggle("folded", layout.stacked && !!layout.folded[t.id]);
    card.hidden = !shown(t.id);
    var pinned = !layout.stacked && !!layout.pinned[t.id];
    card.classList.toggle("pinned", pinned);
    var pin = card.querySelector(".pin-btn");
    pin.setAttribute("aria-pressed", String(pinned));
    var fold = card.querySelector(".fold-btn");
    fold.title = layout.stacked ? "Fold or unfold" : "Close " + t.title;
    fold.setAttribute("aria-label", fold.title);
    card.querySelector(".tool-title").title = layout.stacked ? "Fold or unfold" : "Close " + t.title;
    pin.title = pinned ? "Unpin: close this tool when you open another" : "Pin: keep this tool open when you open another";
    var btn = card.querySelector(".level-btn");
    if (btn) {
      var allowed = levelsOf(t);
      btn.querySelector(".level-name").textContent = LEVEL_NAMES[lv];
      btn.querySelector(".level-dots").innerHTML = allowed.map(function (l) {
        return '<i class="' + (allowed.indexOf(l) <= allowed.indexOf(lv) ? "on" : "") + '"></i>';
      }).join("");
      var next = allowed[(allowed.indexOf(lv) + 1) % allowed.length];
      btn.title = "Showing " + LEVEL_NAMES[lv] + ". Click for " + LEVEL_NAMES[next] + ".";
    }
  }

  // Display order: everything in saved order when stacked. In the launcher view, pinned
  // tools first, then the open one.
  function displayOrder() {
    if (layout.stacked) return layout.order.slice();
    var pinned = layout.order.filter(function (id) { return layout.pinned[id]; });
    var rest = layout.order.filter(function (id) { return !layout.pinned[id]; });
    if (layout.open && !layout.pinned[layout.open]) { rest.splice(rest.indexOf(layout.open), 1); pinned.push(layout.open); }
    return pinned.concat(rest);
  }
  function applyAll() {
    var stack = ST.$("stack");
    document.body.classList.toggle("stacked", !!layout.stacked);
    displayOrder().forEach(function (id) { stack.appendChild(cards[id]); applyCard(tool(id)); });
    var anyTool = layout.order.some(function (id) { return !layout.hidden[id]; });
    var anyShown = layout.order.some(shown);
    var note = ST.$("emptyNote");
    note.hidden = anyShown;
    note.textContent = !anyTool ? "All tools are hidden. Use the button at the top right to show them." : "Pick a tool above to open it.";
    renderLauncher();
    ST.emit("layout");
  }

  // ---------- Launcher ----------
  function renderLauncher() {
    var el = ST.$("launcher");
    if (!el) return;
    el.hidden = !!layout.stacked;
    if (layout.stacked) return;
    el.innerHTML = "";
    layout.order.forEach(function (id) {
      if (layout.hidden[id]) return;
      var t = tool(id), b = document.createElement("button");
      b.className = "launch-btn";
      b.setAttribute("data-tool", id);
      b.title = t.title + (t.tip ? ". " + t.tip.what : "");
      b.innerHTML = '<span class="launch-icon">' + t.icon + '</span><span class="launch-name"></span>';
      b.querySelector(".launch-name").textContent = SHORT[id] || t.title;
      b.classList.toggle("on", layout.open === id);
      b.classList.toggle("pinned", !!layout.pinned[id]);
      b.classList.toggle("suggest", suggested === id);
      b.setAttribute("aria-pressed", String(shown(id)));
      b.addEventListener("click", function () { lastPick = Date.now(); pick(id); });
      el.appendChild(b);
    });
  }
  function scrollToCard(id) {
    var card = cards[id];
    if (!card || card.hidden) return;
    var top = card.getBoundingClientRect().top + window.scrollY - ST.$("launcher").getBoundingClientRect().height - 8;
    if (card.getBoundingClientRect().top < 60 || card.getBoundingClientRect().top > window.innerHeight - 80) window.scrollTo(0, Math.max(0, top));
  }
  // A launcher click: opens the tool, or closes it when it's already the open one.
  function pick(id) {
    if (layout.open === id) layout.open = null;
    else layout.open = id;
    saveLayout(); applyAll();
    if (layout.open) scrollToCard(id);
  }
  function togglePin(id) {
    if (layout.pinned[id]) {
      delete layout.pinned[id];
    } else {
      layout.pinned[id] = true;
      if (layout.open === id) layout.open = null; // it stays shown as a pinned tool
    }
    saveLayout(); applyAll();
  }
  // Opens a tool from code (the selection, or another tool's button). Returns true when shown.
  function openTool(id, auto) {
    var t = tool(id);
    if (!t || layout.hidden[id]) return false;
    if (layout.stacked) {
      if (auto) return false; // the stacked view never moves on its own
      layout.folded[id] = false;
    } else if (!layout.pinned[id]) {
      layout.open = id;
    }
    saveLayout(); applyAll();
    scrollToCard(id);
    return true;
  }
  function closeTool(id) {
    if (layout.stacked) { layout.folded[id] = true; }
    else { delete layout.pinned[id]; if (layout.open === id) layout.open = null; }
    saveLayout(); applyAll();
  }

  // ---------- Following the selection ----------
  // A picture opens Images, a table opens Tables, and a text cursor opens Typography.
  var FOLLOW = { picture: "images", table: "tables", text: "typography" };
  var lastKind = null, suggested = null;
  function followSelection() {
    var kind = ST.selection.kind || "none";
    var id = FOLLOW[kind] || null;
    if (id && layout.hidden[id]) id = null;
    suggested = id;
    var changed = kind !== lastKind;
    lastKind = kind;
    if (changed && id && !layout.stacked && !layout.noAutoOpen && Date.now() - lastPick > 4000 && !shown(id)) { openTool(id, true); return; }
    renderLauncher();
  }

  // ---------- Actions ----------
  function toggleFold(id) {
    if (!layout.stacked) { closeTool(id); return; }
    layout.folded[id] = !layout.folded[id];
    saveLayout(); applyCard(tool(id)); ST.emit("layout");
  }
  function cycleLevel(id) {
    var t = tool(id), allowed = levelsOf(t);
    if (layout.stacked && layout.folded[id]) layout.folded[id] = false;
    else layout.levels[id] = allowed[(allowed.indexOf(levelOf(t)) + 1) % allowed.length];
    saveLayout(); applyCard(t); ST.emit("layout");
  }
  function move(id, delta) {
    var visible = layout.order.filter(function (x) { return !layout.hidden[x]; });
    var i = visible.indexOf(id), j = i + delta;
    if (j < 0 || j >= visible.length) return;
    var a = layout.order.indexOf(id), b = layout.order.indexOf(visible[j]);
    layout.order[a] = visible[j]; layout.order[b] = id;
    saveLayout(); applyAll();
  }
  function setHidden(id, hidden) {
    layout.hidden[id] = hidden;
    if (hidden) { delete layout.pinned[id]; if (layout.open === id) layout.open = null; }
    saveLayout(); applyAll();
  }
  function resetTool(id) {
    var t = tool(id);
    ST.resetKeys(Object.keys(t.defaults || {}));
    ST.emit("reset");
    ST.changed();
    ST.toast(t.title + " is back to its defaults.");
  }

  // ---------- Menus ----------
  function closeMenu() {
    if (openMenu) { openMenu.remove(); openMenu = null; }
  }
  function showMenu(anchor, items) {
    closeMenu();
    var menu = document.createElement("div");
    menu.className = "menu";
    items.forEach(function (it) {
      if (it === "sep") { menu.appendChild(Object.assign(document.createElement("div"), { className: "sep" })); return; }
      if (it.title) { menu.appendChild(Object.assign(document.createElement("div"), { className: "menu-title", textContent: it.title })); return; }
      var b = document.createElement("button");
      b.innerHTML = (it.checked === undefined ? "" : (it.checked ? ICONS.check : ICONS.blank)) + "<span></span>";
      b.querySelector("span").textContent = it.label;
      b.disabled = !!it.disabled;
      b.addEventListener("click", function (e) { e.stopPropagation(); closeMenu(); it.action(); });
      menu.appendChild(b);
    });
    document.body.appendChild(menu);
    var r = anchor.getBoundingClientRect();
    var left = Math.min(r.right - menu.offsetWidth, window.innerWidth - menu.offsetWidth - 8);
    var top = r.bottom + 4;
    if (top + menu.offsetHeight > window.innerHeight - 8) top = Math.max(8, r.top - menu.offsetHeight - 4);
    menu.style.left = Math.max(8, left) + "px";
    menu.style.top = top + "px";
    openMenu = menu;
  }
  function showToolMenu(id, anchor) {
    var visible = layout.order.filter(function (x) { return !layout.hidden[x]; });
    var i = visible.indexOf(id), t = tool(id), word = layout.stacked ? ["Move up", "Move down"] : ["Move earlier", "Move later"];
    showMenu(anchor, [
      { label: "How to use " + t.title, action: function () { ST.openHelp(t.help || t.id); } },
      "sep",
      { label: word[0], disabled: i <= 0, action: function () { move(id, -1); } },
      { label: word[1], disabled: i >= visible.length - 1, action: function () { move(id, 1); } },
      "sep",
      { label: "Reset " + t.title + " to defaults", disabled: !Object.keys(t.defaults || {}).length, action: function () { resetTool(id); } },
      { label: "Hide " + t.title, action: function () { setHidden(id, true); ST.toast(t.title + " hidden. Use the gear button to show it again."); } }
    ]);
  }
  function showSettingsMenu(anchor) {
    var items = ST.update ? ST.update.menuItems().concat(["sep"]) : [];
    items.push({ title: "Show tools" });
    layout.order.forEach(function (id) {
      var t = tool(id);
      items.push({ label: t.title, checked: !layout.hidden[id], action: function () { setHidden(id, !layout.hidden[id]); } });
    });
    items.push("sep");
    items.push({ label: "Show all tools in one list", checked: !!layout.stacked, action: function () { layout.stacked = !layout.stacked; saveLayout(); applyAll(); } });
    items.push({ label: "Open tools with the selection", checked: !layout.noAutoOpen, action: function () { layout.noAutoOpen = !layout.noAutoOpen; saveLayout(); } });
    if (layout.stacked) {
      items.push({ label: "Unfold all", action: function () { layout.folded = {}; saveLayout(); applyAll(); } });
      items.push({ label: "Fold all", action: function () { layout.order.forEach(function (id) { layout.folded[id] = true; }); saveLayout(); applyAll(); } });
    } else {
      items.push({ label: "Unpin all", disabled: !Object.keys(layout.pinned).length, action: function () { layout.pinned = {}; saveLayout(); applyAll(); } });
    }
    items.push("sep");
    items.push({ label: "Share settings", action: function () { ST.openShare(); } });
    items.push({ label: "Keyboard shortcuts", action: function () { ST.shortcuts.show(); } });
    items.push({ label: "Ask before replacing a slide again", action: function () { ST.resetSwapConfirm(); ST.toast("Slide Tools will ask before cropping pictures or making copies."); } });
    items.push({ label: "Reset panel arrangement", action: function () {
      layout = { order: defaultOrder(), levels: {}, folded: {}, hidden: {}, pinned: {}, open: "layout", stacked: false, noAutoOpen: false, v: LAYOUT_VERSION };
      saveLayout(); applyAll(); ST.toast("Panel view, order, detail and pins are back to the defaults.");
    } });
    showMenu(anchor, items);
  }

  // ---------- Drag to reorder ----------
  function wireDrag(card, grip) {
    var scrollTimer = null, lastY = 0;
    // Listeners go on the window: moving the panel in the page would drop pointer capture.
    function onMove(e) {
      lastY = e.clientY;
      reposition(lastY);
      clearInterval(scrollTimer);
      if (lastY < 50 || lastY > window.innerHeight - 30) {
        scrollTimer = setInterval(function () { window.scrollBy(0, lastY < 50 ? -8 : 8); reposition(lastY); }, 30);
      }
    }
    function reposition(y) {
      var stack = ST.$("stack");
      var others = Array.prototype.filter.call(stack.children, function (c) { return c !== card && !c.hidden; });
      var before = null;
      for (var i = 0; i < others.length; i++) {
        var r = others[i].getBoundingClientRect();
        if (y < r.top + r.height / 2) { before = others[i]; break; }
      }
      if (before ? card.nextElementSibling !== before : stack.lastElementChild !== card) {
        stack.insertBefore(card, before);
      }
    }
    function onEnd() {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onEnd);
      window.removeEventListener("pointercancel", onEnd);
      clearInterval(scrollTimer);
      card.classList.remove("dragging");
      var dom = Array.prototype.map.call(ST.$("stack").children, function (c) { return c.getAttribute("data-tool"); });
      if (layout.stacked) layout.order = dom;
      else {
        // Only the open and pinned tools were moved: put them back in their own slots, in the new order
        var moved = dom.filter(shown), slots = moved.map(function (id) { return layout.order.indexOf(id); }).sort(function (a, b) { return a - b; });
        slots.forEach(function (slot, k) { layout.order[slot] = moved[k]; });
      }
      saveLayout(); applyAll();
    }
    grip.addEventListener("pointerdown", function (e) {
      if (e.button !== 0) return;
      e.preventDefault();
      closeMenu();
      card.classList.add("dragging");
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onEnd);
      window.addEventListener("pointercancel", onEnd);
    });
  }

  // ---------- Info tips ----------
  // Tips are positioned against the whole panel, so a folded or short tool can't clip them.
  // They open below the i icon, or above it when there isn't room.
  function placeTip(info) {
    var tip = info.querySelector(".tip");
    if (!tip) return;
    tip.style.display = "block";
    var r = info.getBoundingClientRect(), w = tip.offsetWidth, h = tip.offsetHeight, gap = 2, margin = 8;
    var left = Math.min(Math.max(margin, r.right - w), window.innerWidth - w - margin);
    var below = window.innerHeight - r.bottom - gap - margin, above = r.top - gap - margin;
    var top = h <= below || below >= above ? r.bottom + gap : Math.max(margin, r.top - gap - h);
    tip.style.left = left + "px";
    tip.style.top = top + "px";
    tip.style.maxHeight = Math.max(below, above, 120) + "px";
    tip.style.display = "";
  }
  function wireTips() {
    var place = function (e) {
      var info = e.target.closest && e.target.closest(".info");
      if (info) placeTip(info);
    };
    document.addEventListener("mouseover", place);
    document.addEventListener("focusin", place);
    window.addEventListener("scroll", function () {
      var open = document.querySelector(".info:hover, .info:focus");
      if (open) placeTip(open);
    }, true);
  }

  // ---------- Start ----------
  ST.buildPanels = function () {
    loadLayout();
    layout.order.forEach(function (id) { buildCard(tool(id)); });
    applyAll();
    ST.$("settingsBtn").addEventListener("click", function (e) { e.stopPropagation(); showSettingsMenu(e.currentTarget); });
    ST.$("repeatBtn").addEventListener("click", function () { ST.runRepeat(); });
    document.addEventListener("click", function (e) { if (openMenu && !openMenu.contains(e.target)) closeMenu(); });
    document.addEventListener("keydown", function (e) { if (e.key === "Escape") closeMenu(); });
    window.addEventListener("resize", closeMenu);
    ST.on("selection", followSelection);
    wireTips();
  };
  ST.levelOf = function (id) { return levelOf(tool(id)); };
  ST.isOpen = function (id) { return shown(id) && !(layout.stacked && layout.folded[id]); };
  ST.openTool = function (id) { return openTool(id, false); };
  ST.pinTool = function (id) {
    if (layout.stacked || layout.pinned[id] || layout.hidden[id]) return;
    layout.pinned[id] = true;
    if (layout.open === id) layout.open = null;
    saveLayout(); applyAll();
  };
})();
