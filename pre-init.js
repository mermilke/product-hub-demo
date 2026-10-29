
  // Apply the user's saved light/dark choice before first paint, so there's no theme flash.
  // If they never chose, no attribute is set and the OS preference (prefers-color-scheme) wins.
  try { var _t = localStorage.getItem("pf-theme"); if (_t === "dark" || _t === "light") document.documentElement.setAttribute("data-theme", _t); } catch (e) {}

  // Register the service worker (installable app on Android + a basic offline fallback). Deferred to load so
  // it never competes with first paint; non-fatal — the app works identically without it. The SW is
  // network-first and ignores all auth/Graph/API traffic (see sw.js): it can't break sign-in, and it answers
  // from its cache only for a file whose fetch failed.
  if ("serviceWorker" in navigator) { window.addEventListener("load", function () { navigator.serviceWorker.register("/sw.js").catch(function () {}); }); }

  // "Signing you in…" instead of a live button when the session is about to resume on its own: MSAL has a
  // cached account (its account-keys list in localStorage) or this is the return trip from Microsoft (an
  // auth response in the hash). Set on <html> here, before the stylesheet applies, so the button never
  // flashes; signin.js clears it the moment a click is actually needed. (F022/F082)
  try {
    var _keys = localStorage.getItem("msal.account.keys");
    var _returning = /[#&?](code|error|state)=/.test(location.hash);
    if ((_keys && _keys !== "[]") || _returning) document.documentElement.classList.add("resuming");
  } catch (e) {}

// Decision 8 (F029): catch errors from the very first script. app.js only adds its own listeners once it has booted, so
// a throw while it boots (the blank-page case) went unseen. Errors are queued here (the Hub's own scripts only, de-duplicated,
// at most 10); app.js takes the queue over when it runs, and signin.js posts it itself if app.js never started.
(function () {
  try {
    var q = window.__pfErrQ = [];
    var own = function (s) { return String(s || "").indexOf(location.origin) !== -1; };
    var add = function (m) { m = String(m || "").replace(/\s+/g, " ").slice(0, 200); if (m && q.length < 10 && q.indexOf(m) === -1) q.push(m); };
    var onErr = function (e) { if (e && e.message && e.filename && e.filename.indexOf(location.origin) === 0) add(e.message + " @ " + e.filename.split("/").pop().split("?")[0] + ":" + (e.lineno || 0)); };
    var onRej = function (e) { var r = e && e.reason; if (r && r.stack && own(r.stack)) add("unhandled: " + (r.message || String(r))); };
    window.addEventListener("error", onErr);
    window.addEventListener("unhandledrejection", onRej);
    window.__pfErrOff = function () { window.removeEventListener("error", onErr); window.removeEventListener("unhandledrejection", onRej); };
  } catch (e) {}
})();
