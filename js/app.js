/* Starts the panel once every tool file has registered itself. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;

  function setPill() {
    var pill = $("selPill"), sel = ST.selection;
    if (!ST.inOffice) { pill.textContent = "Preview mode"; pill.className = "pill"; return; }
    if (!sel.count) { pill.textContent = "Nothing selected"; pill.className = "pill"; return; }
    pill.textContent = sel.count === 1
      ? "1 shape  " + ST.nice(sel.w) + " x " + ST.nice(sel.h) + " in"
      : sel.count + " shapes selected";
    pill.className = "pill on";
  }

  // Follows PowerPoint's dark themes when Office tells us the colours.
  function applyTheme() {
    var dark = false;
    try {
      var bg = Office.context.officeTheme && Office.context.officeTheme.bodyBackgroundColor;
      if (bg) {
        var hex = bg.replace("#", "");
        var r = parseInt(hex.substr(0, 2), 16), g = parseInt(hex.substr(2, 2), 16), b = parseInt(hex.substr(4, 2), 16);
        dark = (0.299 * r + 0.587 * g + 0.114 * b) < 128;
      }
    } catch (e) { dark = false; }
    if (/[?&]theme=dark/.test(location.search)) dark = true; // for previewing in a browser
    document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
  }

  function start(info) {
    ST.inOffice = !!(info && info.host === Office.HostType.PowerPoint);
    var banner = $("banner");
    if (!ST.inOffice) {
      banner.hidden = false;
      banner.textContent = "Preview mode. The calculator works here, and the slide buttons work once this panel is opened from PowerPoint.";
    } else if (!ST.supports("1.5")) {
      banner.hidden = false;
      banner.textContent = "This version of PowerPoint is too old for the slide buttons. Update Office to use Draw, Fit and Resize.";
    } else {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, ST.refreshSelection);
      ST.refreshSelection();
    }
    applyTheme();
    setPill();
    ST.changed();
    ST.emit("ready");
  }

  ST.loadState();
  ST.buildPanels();
  ST.on("selection", setPill);
  ST.changed();
  applyTheme();

  var started = false;
  function startOnce(info) { if (!started) { started = true; start(info); } }
  if (window.Office && Office.onReady) Office.onReady(startOnce); else startOnce(null);

  // Exposed for testing only
  window.__slideTools = { evaluate: ST.evaluate, grid: function () { return ST.grid; }, ST: ST };
})();
