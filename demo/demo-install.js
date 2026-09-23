/* Demo: switch on the in-browser backend before app.js makes its first request. */
(function () {
  "use strict";
  var backend = window.PF_DEMO.createBackend(window.PF_DEMO_SEED);
  window.PF_DEMO.install(backend, function (name) { return "/demo/files/" + encodeURIComponent(name); });
})();
