/* Panel framework: builds a dockable panel for each registered tool,
   with detail levels, folding, drag to reorder, hiding and a saved arrangement. */
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
    blank: '<svg viewBox="0 0 16 16" width="14" height="14"></svg>'
  };
  ST.ICONS = ICONS;

  var layout = { order: [], levels: {}, folded: {}, hidden: {} };
  var cards = {};
  var openMenu = null;

  function loadLayout() {
    var saved = null;
    try { saved = JSON.parse(localStorage.getItem(LAYOUT_KEY) || "null"); } catch (e) { saved = null; }
    if (saved) Object.keys(layout).forEach(function (k) { if (saved[k]) layout[k] = saved[k]; });
    // Keep saved order, then add any tools that are new since last time.
    var ids = ST.tools.map(function (t) { return t.id; });
    layout.order = layout.order.filter(function (id) { return ids.indexOf(id) >= 0; });
    ids.forEach(function (id) { if (layout.order.indexOf(id) < 0) layout.order.push(id); });
  }
  function saveLayout() {
    try { localStorage.setItem(LAYOUT_KEY, JSON.stringify(layout)); } catch (e) { /* storage unavailable */ }
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
        '<button class="icon-btn menu-btn" title="Panel options" aria-label="Panel options">' + ICONS.more + '</button>' +
        '<button class="icon-btn fold-btn" title="Fold or unfold" aria-label="Fold or unfold">' + ICONS.chev + '</button>' +
      '</div>' +
      '<div class="tool-body"></div>';
    var body = card.querySelector(".tool-body");
    ST.$("stack").appendChild(card); // tools look up their boxes by id, so attach first
    t.build(body, card);

    card.querySelector(".tool-title").addEventListener("click", function () { toggleFold(t.id); });
    card.querySelector(".fold-btn").addEventListener("click", function () { toggleFold(t.id); });
    if (!single) card.querySelector(".level-btn").addEventListener("click", function () { cycleLevel(t.id); });
    card.querySelector(".menu-btn").addEventListener("click", function (e) { e.stopPropagation(); showToolMenu(t.id, e.currentTarget); });
    wireDrag(card, card.querySelector(".grip"));
    cards[t.id] = card;
    return card;
  }

  function applyCard(t) {
    var card = cards[t.id], lv = levelOf(t);
    card.setAttribute("data-level", lv);
    card.classList.toggle("folded", !!layout.folded[t.id]);
    card.hidden = !!layout.hidden[t.id];
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

  function applyAll() {
    var stack = ST.$("stack");
    layout.order.forEach(function (id) { stack.appendChild(cards[id]); applyCard(tool(id)); });
    var anyVisible = layout.order.some(function (id) { return !layout.hidden[id]; });
    ST.$("emptyNote").hidden = anyVisible;
    ST.emit("layout");
  }

  // ---------- Actions ----------
  function toggleFold(id) {
    layout.folded[id] = !layout.folded[id];
    saveLayout(); applyCard(tool(id)); ST.emit("layout");
  }
  function cycleLevel(id) {
    var t = tool(id), allowed = levelsOf(t);
    if (layout.folded[id]) layout.folded[id] = false;
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
    var i = visible.indexOf(id), t = tool(id);
    showMenu(anchor, [
      { label: "Move up", disabled: i <= 0, action: function () { move(id, -1); } },
      { label: "Move down", disabled: i >= visible.length - 1, action: function () { move(id, 1); } },
      "sep",
      { label: "Reset " + t.title + " to defaults", disabled: !Object.keys(t.defaults || {}).length, action: function () { resetTool(id); } },
      { label: "Hide " + t.title, action: function () { setHidden(id, true); ST.toast(t.title + " hidden. Use the gear button to show it again."); } }
    ]);
  }
  function showSettingsMenu(anchor) {
    var items = [{ title: "Show tools" }];
    layout.order.forEach(function (id) {
      var t = tool(id);
      items.push({ label: t.title, checked: !layout.hidden[id], action: function () { setHidden(id, !layout.hidden[id]); } });
    });
    items.push("sep");
    items.push({ label: "Unfold all", action: function () { layout.folded = {}; saveLayout(); applyAll(); } });
    items.push({ label: "Fold all", action: function () { layout.order.forEach(function (id) { layout.folded[id] = true; }); saveLayout(); applyAll(); } });
    items.push({ label: "Reset panel arrangement", action: function () {
      layout = { order: ST.tools.map(function (t) { return t.id; }), levels: {}, folded: {}, hidden: {} };
      saveLayout(); applyAll(); ST.toast("Panel order, detail and folding are back to the defaults.");
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
      layout.order = Array.prototype.map.call(ST.$("stack").children, function (c) { return c.getAttribute("data-tool"); });
      saveLayout();
      ST.emit("layout");
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

  // ---------- Start ----------
  ST.buildPanels = function () {
    loadLayout();
    layout.order.forEach(function (id) { buildCard(tool(id)); });
    applyAll();
    ST.$("settingsBtn").addEventListener("click", function (e) { e.stopPropagation(); showSettingsMenu(e.currentTarget); });
    document.addEventListener("click", function (e) { if (openMenu && !openMenu.contains(e.target)) closeMenu(); });
    document.addEventListener("keydown", function (e) { if (e.key === "Escape") closeMenu(); });
    window.addEventListener("resize", closeMenu);
  };
  ST.levelOf = function (id) { return levelOf(tool(id)); };
  ST.isOpen = function (id) { return !layout.folded[id] && !layout.hidden[id]; };
})();
