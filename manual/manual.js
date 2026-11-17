/* Slide Tools manual: theme switch, "On this page" list, and the phone menu. */
(function () {
  "use strict";
  var root = document.documentElement;

  // Theme: follows the system unless the reader picks one, which is remembered.
  try { var saved = localStorage.getItem("manualTheme"); if (saved) root.setAttribute("data-theme", saved); } catch (e) { /* storage off */ }
  function current() {
    return root.getAttribute("data-theme") || (window.matchMedia && matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
  }
  function label(btn) { btn.textContent = current() === "dark" ? "Light mode" : "Dark mode"; }

  document.addEventListener("DOMContentLoaded", function () {
    var btn = document.getElementById("themeBtn");
    if (btn) {
      label(btn);
      btn.addEventListener("click", function () {
        var next = current() === "dark" ? "light" : "dark";
        root.setAttribute("data-theme", next);
        try { localStorage.setItem("manualTheme", next); } catch (e) { /* storage off */ }
        label(btn);
      });
    }

    // Phone menu
    var toggle = document.getElementById("menuToggle"), side = document.querySelector(".side");
    if (toggle && side) toggle.addEventListener("click", function () { side.classList.toggle("open"); });

    // On this page: built from the h2 headings that have an id
    var list = document.getElementById("onpageList");
    var heads = Array.prototype.slice.call(document.querySelectorAll("main h2[id]"));
    if (list) {
      heads.forEach(function (h) {
        var li = document.createElement("li"), a = document.createElement("a");
        a.href = "#" + h.id; a.textContent = h.textContent;
        li.appendChild(a); list.appendChild(li);
      });
      var links = Array.prototype.slice.call(list.querySelectorAll("a"));
      var mark = function () {
        var y = window.scrollY + 100, on = 0;
        heads.forEach(function (h, i) { if (h.offsetTop <= y) on = i; });
        links.forEach(function (a, i) { a.classList.toggle("on", i === on); });
      };
      window.addEventListener("scroll", mark, { passive: true });
      mark();
    }
  });
})();
