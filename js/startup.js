/* Starting with a presentation. PowerPoint doesn't let an add-in open itself whenever PowerPoint
   starts, but an add-in with a shared runtime can ask to load quietly whenever a presentation
   opens (SharedRuntime 1.1). PowerPoint keeps that choice in the file, so turning it on in a
   template carries it to every deck made from it. A second choice, kept in the document's own
   settings, shows the panel once it has loaded. */
(function () {
  "use strict";

  var ST = window.ST;
  var SHOW_KEY = "slideTools.showOnStart";
  var state = { ok: false, load: false, show: false };

  function supported() {
    try { return !!(ST.inOffice && Office.addin && Office.addin.setStartupBehavior && Office.context.requirements.isSetSupported("SharedRuntime", "1.1")); }
    catch (e) { return false; }
  }
  function settings() { try { return Office.context.document.settings; } catch (e) { return null; } }
  function readShow() { var s = settings(); try { return !!(s && s.get(SHOW_KEY)); } catch (e) { return false; } }
  function saveSettings() {
    var s = settings();
    return new Promise(function (done) { if (!s || !s.saveAsync) { done(); return; } s.saveAsync(function () { done(); }); });
  }
  // Reads both choices for the settings menu
  function refresh() {
    state.ok = supported();
    state.show = state.ok && readShow();
    if (!state.ok) { state.load = false; return Promise.resolve(state); }
    return Promise.resolve().then(function () { return Office.addin.getStartupBehavior(); })
      .then(function (b) { state.load = b === Office.StartupBehavior.load; }, function () { state.load = false; })
      .then(function () { return state; });
  }
  function setLoad(on) {
    if (!supported()) { ST.toast("Starting with a presentation needs a newer PowerPoint.", true); return Promise.resolve(); }
    return Promise.resolve().then(function () { return Office.addin.setStartupBehavior(on ? Office.StartupBehavior.load : Office.StartupBehavior.none); })
      .then(function () {
        state.load = on;
        if (!on && state.show) return setShow(false, true);
      }).then(function () {
        ST.toast(on ? "Slide Tools will start when this presentation opens, once it's saved. Turn this on in your company template and new decks start with it too." : "Slide Tools won't start by itself with this presentation.");
      }, function (e) { ST.toast("Couldn't change how Slide Tools starts: " + ST.reason(e), true); });
  }
  function setShow(on, quiet) {
    var s = settings();
    if (!s) return Promise.resolve();
    try { if (on) s.set(SHOW_KEY, true); else s.remove(SHOW_KEY); } catch (e) { return Promise.resolve(); }
    state.show = on;
    return saveSettings().then(function () {
      if (!quiet) ST.toast(on ? "The panel will open when this presentation opens." : "The panel will stay hidden when it starts. Use the shortcut or the Home tab button to show it.");
    });
  }
  // On a start with the presentation, show the panel when the file asks for it
  function showIfAsked() {
    if (!supported() || !readShow() || !Office.addin.showAsTaskpane) return;
    try { Office.addin.showAsTaskpane(); } catch (e) { /* the panel can still be opened from the Home tab */ }
  }
  function showPanel() {
    try { if (window.Office && Office.addin && Office.addin.showAsTaskpane) return Office.addin.showAsTaskpane(); } catch (e) { /* not available here */ }
    return Promise.resolve();
  }

  ST.startup = { refresh: refresh, setLoad: setLoad, setShow: setShow, state: function () { return state; }, showPanel: showPanel, supported: supported };
  ST.on("ready", function () { refresh().then(showIfAsked); });
})();
