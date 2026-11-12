/* Corners tool: round selected rounded rectangles to an exact radius in points. */
(function () {
  "use strict";

  var ST = window.ST, $ = ST.$;
  var PRESETS = [2, 4, 8, 12, 16];

  var defaults = { radius: "8" };
  ST.addDefaults(defaults);

  var ICON = '<svg viewBox="0 0 18 18" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M3 16V9a6 6 0 0 1 6-6h7"/></svg>';

  var MARKUP =
    '<div class="inline">' +
      '<label class="field">Radius (pt)<input type="text" id="radius" inputmode="decimal"></label>' +
      '<button class="btn" id="round">Round selection</button>' +
    '</div>' +
    '<div class="mini-chips" id="radiusChips" data-min="3"></div>' +
    '<button class="btn small" id="readRadius" data-min="3" title="Set the radius box from the selected rounded rectangle">Read radius from selection</button>' +
    '<div class="help" data-min="3">Works on rounded rectangles. Square rectangles can\'t be converted from the panel, so draw them rounded or use Round in your VBA Slide Tools.</div>';

  function renderChips() {
    var el = $("radiusChips");
    if (!el.children.length) {
      PRESETS.forEach(function (v) {
        var b = document.createElement("button");
        b.className = "mini-chip";
        b.textContent = v + " pt";
        b.addEventListener("click", function () {
          ST.state.radius = String(v); $("radius").value = ST.state.radius; ST.changed();
        });
        el.appendChild(b);
      });
    }
    var current = ST.num(ST.state.radius);
    Array.prototype.forEach.call(el.children, function (b, i) { b.classList.toggle("on", PRESETS[i] === current); });
  }

  function roundSelection() {
    if (!ST.needOffice()) return;
    if (!ST.cornersSupported()) {
      ST.toast("This PowerPoint version can't change corners yet. Use Round in your VBA Slide Tools.", true);
      return;
    }
    ST.withSelection(["type", "width", "height", "adjustments/count"], function (context, items) {
      var changed = 0;
      items.forEach(function (s) {
        // Rounded rectangles have exactly one adjustment handle
        if (s.type === "GeometricShape" && s.adjustments.count === 1) {
          s.adjustments.set(0, ST.cornerValue(s.width, s.height));
          changed++;
        }
      });
      return context.sync().then(function () {
        if (changed) ST.toast("Rounded " + ST.plural(changed, "shape") + ".");
        else ST.toast("No rounded rectangles selected. Square rectangles need Round in your VBA Slide Tools.", true);
      });
    });
  }

  function readRadius() {
    if (!ST.needOffice()) return;
    if (!ST.cornersSupported()) { ST.toast("This PowerPoint version can't read corners yet.", true); return; }
    ST.withSelection(["type", "width", "height", "adjustments/count"], function (context, items) {
      var s = items.filter(function (x) { return x.type === "GeometricShape" && x.adjustments.count === 1; })[0];
      if (!s) { ST.toast("Select a rounded rectangle to read its radius.", true); return; }
      var value = s.adjustments.get(0);
      return context.sync().then(function () {
        var pt = value.value * Math.min(s.width, s.height);
        ST.state.radius = String(Math.round(pt * 10) / 10);
        $("radius").value = ST.state.radius;
        ST.changed();
        ST.toast("Radius is " + ST.state.radius + " pt.");
      });
    });
  }

  function build(body) {
    body.appendChild(ST.html(MARKUP));
    ST.bindField($("radius"), "radius", { number: true });
    $("round").addEventListener("click", roundSelection);
    $("readRadius").addEventListener("click", readRadius);
    $("radius").addEventListener("keydown", function (e) { if (e.key === "Enter") roundSelection(); });
    ST.on("change", renderChips);
  }

  ST.registerTool({
    id: "corners",
    title: "Corners",
    icon: ICON,
    levels: [1, 3],
    defaultLevel: 1,
    defaults: defaults,
    build: build
  });
})();
