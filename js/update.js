/* Versions and updates. The panel's version comes from the ?v= on this script, and the
   install version from ?install= on the panel address in the install file. version.json on
   the site says what's current. Panel updates apply by reloading the panel. A new install
   file needs the installer, which the panel can download but can't run. */
(function () {
  "use strict";

  var ST = window.ST;
  var INSTALLER = "https://chadisfred.github.io/Slidetools/install/Install-SlideTools.cmd";
  var latest = null, checking = null;

  // This script's own ?v=, which every release raises along with the others
  function scriptVersion() {
    var s = document.currentScript || document.querySelector('script[src*="js/update.js"]'), m = s && /[?&]v=([\d.]+)/.exec(s.getAttribute("src"));
    return m ? m[1] : "";
  }
  function installVersion() { var m = /[?&]install=(\d+)/.exec(location.search); return m ? +m[1] : 0; }
  // 1 if a is newer than b, -1 if older, 0 if the same
  function compare(a, b) {
    var x = String(a).split("."), y = String(b).split(".");
    for (var i = 0; i < Math.max(x.length, y.length); i++) { var d = (+x[i] || 0) - (+y[i] || 0); if (d) return d > 0 ? 1 : -1; }
    return 0;
  }
  ST.version = scriptVersion();

  // The changes listed in CHANGELOG.md for versions newer than the panel's, up to five versions
  function parseChanges(md, since) {
    var out = [], cur = null;
    String(md).split(/\r?\n/).forEach(function (line) {
      var h = /^## \[([\d.]+)\]/.exec(line);
      if (h) { cur = compare(h[1], since) > 0 && out.length < 5 ? { v: h[1], items: [] } : null; if (cur) out.push(cur); return; }
      var it = cur && /^\d+\.\s+(.*)$/.exec(line);
      if (it) cur.items.push(it[1].replace(/`/g, ""));
    });
    return out;
  }
  function fresh(url) { return fetch(url + (url.indexOf("?") < 0 ? "?" : "&") + "t=" + Date.now(), { cache: "no-store" }); }

  function state() {
    if (!latest) return { known: false };
    var panel = ST.version && compare(latest.panel, ST.version) > 0, install = installVersion() < (+latest.install || 0);
    return { known: true, panel: !!panel, install: install, any: !!(panel || install) };
  }
  function showDot() { var b = ST.$("settingsBtn"); if (b) b.classList.toggle("has-update", state().any); }

  function check(quiet) {
    if (checking) return checking;
    checking = fresh("version.json").then(function (r) { if (!r.ok) throw new Error(r.status); return r.json(); }).then(function (v) {
      latest = v; showDot();
      if (!quiet) { if (state().any) open(); else ST.toast("Slide Tools is up to date (panel " + ST.version + ", install " + installVersion() + ")."); }
    }).catch(function () {
      if (!quiet) ST.toast("Couldn't check for updates. Check your connection and try again.", true);
    }).then(function () { checking = null; });
    return checking;
  }

  function reloadPanel() {
    var q = location.search.replace(/([?&])r=\d+&?/, "$1").replace(/[?&]$/, "");
    location.replace(location.pathname + (q ? q + "&" : "?") + "r=" + Date.now());
  }
  function downloadInstaller() {
    try {
      if (window.Office && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(INSTALLER); return; }
    } catch (e) {}
    window.open(INSTALLER, "_blank");
  }

  function esc(t) { return String(t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function open() {
    var st = state(), inst = installVersion();
    var body = '<div class="sheet-text">This panel is version <b>' + esc(ST.version || "unknown") + '</b>, install ' + (inst ? '<b>' + inst + '</b>' : 'from before version tracking') + '.' +
      (st.known ? ' The latest is panel <b>' + esc(latest.panel) + '</b>, install <b>' + esc(latest.install) + '</b>.' : '') + '</div>';
    if (st.known && !st.any) body += '<div class="sheet-text">Everything is up to date.</div>';
    body += '<div class="update-changes" id="updChanges"></div>';
    if (st.panel) body += '<button class="btn primary" id="updPanel">Update panel to ' + esc(latest.panel) + '</button><div class="sheet-text">Reloads the panel. Your settings, styles and colours are kept.</div>';
    if (st.install) body += '<button class="btn' + (st.panel ? '' : ' primary') + '" id="updInstall">Download installer</button>' +
      '<div class="sheet-text">A new install file is ready' + (latest.installNote ? ': ' + esc(latest.installNote) : '') + '. Download the installer, close PowerPoint, then double click the file. Open PowerPoint again when it says it\'s done.</div>';
    if (!st.known) body += '<button class="btn" id="updCheck">Check for updates</button>';
    ST.sheet(st.any ? "Update available" : "About Slide Tools", ST.html(body));
    var b;
    if ((b = ST.$("updPanel"))) b.addEventListener("click", reloadPanel);
    if ((b = ST.$("updInstall"))) b.addEventListener("click", function () { downloadInstaller(); ST.toast("The installer is downloading in your browser. Close PowerPoint, then double click it."); });
    if ((b = ST.$("updCheck"))) b.addEventListener("click", function () { check(false); });
    if (st.panel) fresh("CHANGELOG.md").then(function (r) { return r.ok ? r.text() : ""; }).then(function (md) {
      var list = parseChanges(md, ST.version), box = ST.$("updChanges");
      if (!box || !list.length) return;
      box.innerHTML = '<div class="menu-title">What\'s new</div>' + list.map(function (c) {
        return '<div class="update-version">' + esc(c.v) + '</div><ul class="sheet-list">' + c.items.map(function (i) { return '<li>' + esc(i) + '</li>'; }).join("") + '</ul>';
      }).join("");
    }).catch(function () {});
  }

  ST.update = {
    // Items for the top of the settings menu
    menuItems: function () {
      var st = state(), items = [{ title: "Slide Tools " + (ST.version || "") + " · install " + (installVersion() || "old") }];
      if (st.panel) items.push({ label: "Update panel to " + latest.panel, action: open });
      else if (st.install) items.push({ label: "New install file available", action: open });
      items.push({ label: st.any ? "What's new" : "About and updates", action: function () { if (st.known) open(); else check(false); } });
      return items;
    },
    check: check, open: open, state: state,
    _: { compare: compare, parseChanges: parseChanges, installVersion: installVersion }
  };

  // Check quietly once the panel has loaded
  setTimeout(function () { check(true); }, 1500);
})();
