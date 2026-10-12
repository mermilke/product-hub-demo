
  // Apply the user's saved light/dark choice before first paint, so there's no theme flash.
  // If they never chose, no attribute is set and the OS preference (prefers-color-scheme) wins.
  try { var _t = localStorage.getItem("pf-theme"); if (_t === "dark" || _t === "light") document.documentElement.setAttribute("data-theme", _t); } catch (e) {}

  // Register the service worker (installable app on Android + a basic offline fallback). Deferred to load so
  // it never competes with first paint; non-fatal — the app works identically without it. The SW is
  // network-first and ignores all auth/Graph/API traffic (see sw.js), so it can't serve stale code or break
  // sign-in.
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
