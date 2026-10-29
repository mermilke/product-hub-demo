/* Demo sign-in: a pretend PM session. The demo has no Microsoft sign-in at all — this replaces signin.js
   and hands app.js the same globals the real sign-in would (PF_user, PF_getToken, PF_getApiToken). */
(function () {
  "use strict";
  if (window.self !== window.top) return; // same guard as signin.js: never start inside a frame
  window.PF_user = "wren@grumbleton.example";
  window.PF_getToken = window.PF_getApiToken = function () { return Promise.resolve("demo"); };
  var bar = document.createElement("div");
  bar.id = "demoBanner"; bar.className = "demo-banner"; bar.setAttribute("role", "note");
  bar.textContent = "Demo — Grumbleton & Sons is a fictional company. Nothing you change is saved.";
  document.body.insertBefore(bar, document.body.firstChild);
  document.body.classList.remove("pre-auth");
  if (window.__startPF) window.__startPF();
})();
