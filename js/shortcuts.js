/* Keyboard shortcuts. The keys are declared in shortcuts.json, linked from the manifest,
   and work while the slide has focus (PowerPoint for Windows 2601 or later, PowerPoint
   for Mac 16.105.2 or later, and PowerPoint on the web). This file names each action and
   connects it to the tool code once Office is ready. */
(function () {
  "use strict";

  var ST = window.ST;
  // id, label, key, and what it runs. Keep in step with shortcuts.json.
  var LIST = [
    ["NudgeLeft", "Nudge left", "Left", function () { arr().nudge(-1, 0); }],
    ["NudgeUp", "Nudge up", "Up", function () { arr().nudge(0, -1); }],
    ["NudgeDown", "Nudge down", "Down", function () { arr().nudge(0, 1); }],
    ["NudgeRight", "Nudge right", "Right", function () { arr().nudge(1, 0); }],
    ["GrowWidth", "Wider", "Right", function () { arr().resizeBy(1, 0); }, "Alt+Shift+"],
    ["ShrinkWidth", "Narrower", "Left", function () { arr().resizeBy(-1, 0); }, "Alt+Shift+"],
    ["GrowHeight", "Taller", "Down", function () { arr().resizeBy(0, 1); }, "Alt+Shift+"],
    ["ShrinkHeight", "Shorter", "Up", function () { arr().resizeBy(0, -1); }, "Alt+Shift+"],
    ["RepeatLast", "Repeat last action", "D", function () { ST.runRepeat(); }],
    ["AlignLeft", "Align left", "L", function () { arr().alignKey("left"); }],
    ["AlignCentre", "Align centres", "C", function () { arr().alignKey("centre"); }],
    ["AlignRight", "Align right", "R", function () { arr().alignKey("right"); }],
    ["AlignTop", "Align top", "T", function () { arr().alignKey("top"); }],
    ["AlignMiddle", "Align middles", "M", function () { arr().alignKey("middle"); }],
    ["AlignBottom", "Align bottom", "B", function () { arr().alignKey("bottom"); }],
    ["DistributeAcross", "Distribute across", "H", function () { arr().distributeKey(true); }],
    ["DistributeDown", "Distribute down", "V", function () { arr().distributeKey(false); }]
  ];
  function arr() { return ST._arrange; }

  function isMac() {
    try { if (window.Office && Office.context && Office.context.diagnostics) return Office.context.diagnostics.platform === "Mac"; } catch (e) {}
    return /Mac/i.test(navigator.platform || "");
  }
  // "Ctrl+Alt+Shift+L", or "Cmd+Option+Shift+L" on a Mac. Resizing uses Alt+Shift (Option+Shift).
  function keyText(key, mods) {
    if (mods) return (isMac() ? mods.replace("Alt", "Option") : mods) + key;
    return (isMac() ? "Cmd+Option+Shift+" : "Ctrl+Alt+Shift+") + key;
  }
  function byId(id) { return LIST.filter(function (s) { return s[0] === id; })[0]; }

  ST.shortcuts = {
    list: function () { return LIST.map(function (s) { return { id: s[0], label: s[1], keys: keyText(s[2], s[4]) }; }); },
    // Text for a tooltip, such as " (Ctrl+Alt+Shift+L)"
    hint: function (id) { var s = byId(id); return s ? " (" + keyText(s[2], s[4]) + ")" : ""; },
    supported: function () { return !!(window.Office && Office.actions && Office.actions.associate); },
    show: function () {
      var rows = LIST.map(function (s) { return '<tr><td>' + s[1] + '</td><td class="keys">' + keyText(s[2], s[4]) + '</td></tr>'; }).join("");
      ST.sheet("Keyboard shortcuts", ST.html('<p class="sheet-text">These work while the slide has focus, so click the slide first. They need PowerPoint for Windows 2601 or later, PowerPoint for Mac 16.105.2 or later, or PowerPoint on the web. Nudge and the size keys step by the Nudge by amount in Arrange. Sizes change around Transform\'s reference point.</p>' +
        '<table class="shortcut-table">' + rows + '</table>'));
    }
  };

  function register() {
    if (!ST.shortcuts.supported()) return;
    LIST.forEach(function (s) {
      try {
        Office.actions.associate(s[0], function () {
          try { s[3](); } catch (e) { if (window.console) console.error(e); }
          return Promise.resolve();
        });
      } catch (e) { if (window.console) console.error(e); }
    });
  }
  // Register straight away, as Microsoft's sample does, so a shortcut that starts the
  // add-in finds its action. Office.onReady repeats it in case Office.actions arrives late.
  var done = false;
  function registerOnce() { if (!done && ST.shortcuts.supported()) { done = true; register(); } }
  registerOnce();
  if (window.Office && Office.onReady) Office.onReady(registerOnce);
})();
