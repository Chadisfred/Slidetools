/* Clears the "start with this presentation" choice that 3.11.0 could save in a file. Loading with
   a presentation made PowerPoint report the add-in as unavailable for installs made with the
   installer, so the choice was removed. When the panel opens in a file that still asks to start
   Slide Tools, it turns that off, and the next save keeps the file clean. */
(function () {
  "use strict";

  var ST = window.ST;
  var SHOW_KEY = "slideTools.showOnStart";

  function supported() {
    try { return !!(ST.inOffice && Office.addin && Office.addin.setStartupBehavior && Office.context.requirements.isSetSupported("SharedRuntime", "1.1")); }
    catch (e) { return false; }
  }
  function clearShow() {
    try {
      var s = Office.context.document.settings;
      if (s && s.get(SHOW_KEY)) { s.remove(SHOW_KEY); if (s.saveAsync) s.saveAsync(function () {}); return true; }
    } catch (e) { /* nothing saved */ }
    return false;
  }
  function cleanUp() {
    if (!supported()) return;
    var shown = clearShow();
    Promise.resolve().then(function () { return Office.addin.getStartupBehavior(); }).then(function (b) {
      if (b !== Office.StartupBehavior.load) return;
      return Promise.resolve(Office.addin.setStartupBehavior(Office.StartupBehavior.none)).then(function () {
        ST.toast("This presentation no longer starts Slide Tools by itself. Save it to keep the change.");
      });
    }).catch(function () { if (shown) ST.toast("Save this presentation to keep it from opening Slide Tools by itself."); });
  }
  function showPanel() {
    try { if (window.Office && Office.addin && Office.addin.showAsTaskpane) return Office.addin.showAsTaskpane(); } catch (e) { /* not available here */ }
    return Promise.resolve();
  }

  ST.startup = { showPanel: showPanel };
  ST.on("ready", cleanUp);
})();
