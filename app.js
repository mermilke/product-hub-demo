
(function () {
  "use strict";

  /* ============================================================
     DATA — the catalog state. Everything product-related arrives
     from the auth-gated /api/catalog after sign-in (it is read from
     the "Product Family Tree" Smartsheet server-side); nothing is
     baked into this file, so the public source carries no data.
     ============================================================ */
  var CATALOG = []; // filled by loadCatalog()

  var RESOURCE_TYPES = [
    { name: "Flyer",                 desc: "Promotional sell sheet",                     icon: "flyer" },
    { name: "Product Overview",      desc: "At-a-glance product summary",                icon: "overview" },
    { name: "Spec Sheet",            desc: "Dimensions, materials & technical data",     icon: "spec" },
    { name: "Cut Sheet",             desc: "Reference sheet for specific configurations", icon: "cut" },
    { name: "Renders",               desc: "Product images & 3D renders",                icon: "render" },
    { name: "In-Store Images",       desc: "Photos of the product in real stores",       icon: "camera" },
    { name: "Assembly Instructions", desc: "Build & installation guides",                icon: "wrench" },
    { name: "Compliance & Testing",  desc: "Certifications, regulatory & test results",  icon: "shield" },
    { name: "Warranty",              desc: "Warranty terms & guarantees",                icon: "badge" },
    { name: "Training",              desc: "Guides, videos & sales enablement",          icon: "train" }
  ];

  // Live SharePoint links. Key = "MODEL::Resource". A value is a URL string, or an
  // array of { label, url } when one resource has multiple documents.
  var LINKS = {}; // filled by loadCatalog(); SharePoint links are internal, so they only ever arrive after sign-in
  var RUNTIME_ALIASES = []; // PM-added search synonyms ([{from,to}]) from /api/catalog; applied in expandSearchQuery
  var RUNTIME_HANDLED = []; // search misses a PM marked handled, from /api/catalog (shared across PMs/devices — T7-21)

  // When each link was added (YYYY-MM-DD, from the link row's created date). Drives the
  // "Recently added" strip, which shows an item only for RECENT_DAYS after this date.
  var RECENT_DAYS = 7;
  var RECENT_MAX = 6; // cap the What's New strip to the N most-recent so a bulk-wiring day can't flood it
  var ADDED = {}; // "MODEL::Resource" -> date added; filled by loadCatalog()
  var ORPHANS = []; // files in a shared auto-linked folder that match no model (from the daily cron; flagged on Progress)
  var SHEET_HYGIENE = hygieneFrom(null); // what in the Product Family Tree the app could not read as intended (from /api/catalog `diagnostics`); shown to PMs on Progress
  var PF_ADMIN_ALERTS = []; // PM-only admin heads-up alerts from /api/me (e.g. Graph secret nearing expiry)
  var PF_serverFavs = null; // this user's cross-device favorites from /api/me (null until loaded; [] = none)
  var NA = {}; // "model::resource" -> 1 for resource types marked "not applicable" (excluded from coverage math)

  var ICONS = {
    flyer: '<path d="m3 11 18-5v12L3 14v-3z"/><path d="M11.6 16.8a3 3 0 1 1-5.8-1.6"/>',
    // Product Overview = a magnifying glass with two summary lines in the lens ("look at the at-a-glance
    // summary"). Deliberately distinct from the search-box magnifying glass (which has no lines) so the
    // same glyph doesn't read as "search".
    overview: '<circle cx="10.5" cy="10.5" r="6.5"/><path d="m20.5 20.5-4.4-4.4"/><path d="M7.6 9.3h5.8M7.6 12h5.8"/>',
    spec: '<path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M17 21H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h7l5 5v11a2 2 0 0 1-2 2z"/><path d="M9 9h1M9 13h6M9 17h6"/>',
    render: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="m21 16-5-5L5 20"/>',
    // In-Store Images = a camera (photos taken in real stores) — deliberately distinct from the render
    // icon above (a framed picture), so "renders" vs "real store photos" read differently at a glance.
    camera: '<path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3z"/><circle cx="12" cy="13" r="3"/>',
    // Cut Sheet = a pair of scissors (a "cut" sheet) — distinct from the "spec" document.
    cut: '<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M20 4 8.12 15.88"/><path d="M14.47 14.48 20 20"/><path d="M8.12 8.12 12 12"/>',
    shield: '<path d="M12 3l7 3v5c0 4.4-3 8.4-7 9.5C8 19.4 5 15.4 5 11V6z"/><path d="m9 12 2 2 4-4"/>',
    train: '<path d="M22 10 12 5 2 10l10 5 10-5z"/><path d="M6 12v5c0 1 2.7 2.5 6 2.5s6-1.5 6-2.5v-5"/>',
    ext: '<path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/>',
    alert: '<path d="M10.3 3.3 1.8 18a2 2 0 0 0 1.7 3h16.9a2 2 0 0 0 1.7-3L13.7 3.3a2 2 0 0 0-3.4 0z"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
    edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z"/>',
    sync: '<path d="M21 2v6h-6"/><path d="M3 12a9 9 0 0 1 15-6.7L21 8"/><path d="M3 22v-6h6"/><path d="M21 12a9 9 0 0 1-15 6.7L3 16"/>',
    folder: '<path d="M2 6a2 2 0 0 1 2-2h5l2 2h9a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2z"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7.5V12l3 1.8"/>',
    book: '<path d="M12 7v14"/><path d="M3 18a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h5a4 4 0 0 1 4 4 4 4 0 0 1 4-4h5a1 1 0 0 1 1 1v13a1 1 0 0 1-1 1h-6a3 3 0 0 0-3 3 3 3 0 0 0-3-3z"/>',
    wrench: '<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>',
    // Warranty = an award medal: the scalloped seal (shrunk + lifted via the inner group; its stroke-width
    // is pre-scaled to 2.449 so the 0.735 scale renders it back to the icon's 1.8) over a banner ribbon
    // with a V-notch. Reads as "guarantee/seal of approval".
    badge: '<g transform="translate(12,8) scale(0.735) translate(-12,-12)" stroke-width="2.449"><path d="M3.85 8.62a4 4 0 0 1 4.78-4.77 4 4 0 0 1 6.74 0 4 4 0 0 1 4.78 4.78 4 4 0 0 1 0 6.74 4 4 0 0 1-4.77 4.78 4 4 0 0 1-6.75 0 4 4 0 0 1-4.78-4.77 4 4 0 0 1 0-6.76Z"/></g><path d="M8.6 13.4 8.6 21.6 12 19.6 15.4 21.6 15.4 13.4"/>',
    // File-type glyphs so a PDF vs a slide deck vs a doc is legible at a glance (a document outline plus a
    // distinguishing mark: PDF = a solid band, PPT = little bars).
    filepdf: '<path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M17 21H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h7l5 5v11a2 2 0 0 1-2 2z"/><rect x="7.6" y="13.8" width="8.8" height="3.6" rx="1"/>',
    fileppt: '<path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M17 21H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h7l5 5v11a2 2 0 0 1-2 2z"/><path d="M8.6 17.2v-2.7M11.6 17.2v-4M14.6 17.2v-1.6"/>'
  };
  // Pick a file icon from the name's extension: PDF and PowerPoint get their own glyph; everything else
  // (doc/docx/xls/…) uses the generic document icon.
  function fileTypeIcon(name) {
    var m = String(name || "").toLowerCase().match(/\.([a-z0-9]+)$/), ext = m ? m[1] : "";
    if (ext === "pdf") return ICONS.filepdf;
    if (ext === "ppt" || ext === "pptx" || ext === "pps" || ext === "ppsx") return ICONS.fileppt;
    return ICONS.spec;
  }

  var content = document.getElementById("content");
  var searchInput = document.getElementById("searchInput");
  var toastWrap = document.getElementById("toastWrap");
  var toastStatus = document.getElementById("toastStatus");
  var toastAlerts = document.getElementById("toastAlerts");
  var toastTopWrap = document.getElementById("toastTopWrap");
  var srAnnounce = document.getElementById("srAnnounce");
  // ---- CSP-safe dynamic styling -----------------------------------------------------------------
  // A strict style-src (no 'unsafe-inline') forbids inline style attributes, but the CSSOM
  // (el.style.x = …) is NOT gated by CSP. So bars/segments carry their numeric value in a data-*
  // attribute and get width/height/--i applied here via the CSSOM. A MutationObserver covers every
  // render path (full render, in-place optimistic updates, within-view nav) so none can be missed.
  function applyDynEl(el) {
    if (!el || el.nodeType !== 1 || !el.hasAttribute) return;
    if (el.hasAttribute("data-barw")) el.style.width = el.getAttribute("data-barw") + "%";
    if (el.hasAttribute("data-barh")) el.style.height = "calc(" + el.getAttribute("data-barh") + "px * var(--demo-scale, 1))";
    if (el.hasAttribute("data-cseg")) el.style.setProperty("--i", el.getAttribute("data-cseg"));
  }
  function applyDynStyles(root) {
    if (!root || root.nodeType !== 1) return;
    applyDynEl(root);
    if (root.querySelectorAll) [].forEach.call(root.querySelectorAll("[data-barw],[data-barh],[data-cseg]"), applyDynEl);
  }
  if (content && window.MutationObserver) {
    new MutationObserver(function (muts) {
      for (var mi = 0; mi < muts.length; mi++) { var an = muts[mi].addedNodes; for (var ni = 0; ni < an.length; ni++) applyDynStyles(an[ni]); }
    }).observe(content, { childList: true, subtree: true });
  }
  var RM = window.matchMedia("(prefers-reduced-motion: reduce)");
  // Land a fresh top-level view at the top of the page (tab clicks and hash-driven view switches).
  function resetScroll() { window.scrollTo({ top: 0, behavior: RM.matches ? "auto" : "smooth" }); }

  // selection path (keys), single path — column view is inherently single-select
  var sel = { prod: null, model: null };
  var openFams = {}; // which families are expanded in the left panel — multiple allowed
  var flashKey = null;
  var pendingFocusKey = null; // set ONLY by direct user nav clicks → restore keyboard focus after the full re-render (not on cold load / deep link)
  var view = "catalog";           // "catalog" | "progress" | "metrics" | "admin" | "styleguide" (#styleguide, PM-only, no tab)
  var _preSearchView = null;      // the PM tab the viewer was on when they started typing a search; restored on clear
  var catalogError = false;      // true if the auth-gated catalog fetch failed
  var catalogDenied = false;     // …and it was a 403: this ACCOUNT isn't authorized (refreshing can't help)
  var covIncompleteOnly = false;  // progress filter toggle
  // Cache of the last Link-health run (per browser) so the catalog can flag resources that
  // resolve to nothing ("Open that shows nothing") or are broken, without re-hitting Graph.
  var HEALTH_KEY = "pf_linkhealth_v1";
  var healthMap = {}, healthTs = 0; // last Link-health run (persisted); healthStatusFor reads healthMap
  (function () { try { var j = JSON.parse(localStorage.getItem(HEALTH_KEY) || "{}"); if (j && j.map) { healthMap = j.map; healthTs = j.ts || 0; } } catch (e) {} })();
  // Server-side link health from the daily cron, served by /api/catalog to EVERY user — so a
  // salesperson who never runs "Check all links" still sees broken/empty flags. `serverHealthTs` is
  // when the cron last ran (ms); a more recent local manual run (healthTs) wins over it. `healthCleared`
  // suppresses a flag for a link the PM just re-wired this session, so the post-save catalog reload
  // (which re-seeds serverHealth before the cron re-checks) can't resurrect the stale "Link issue".
  var serverHealth = {}, serverHealthTs = 0, serverHealthDateOnly = false, healthCleared = {};

  /* ============================================================
     RENDER DISPATCHER — view state, render(), document title, tab sync
     ============================================================ */
  // Whether the signed-in user is on the product-management team. The allowlist lives ONLY on the
  // server (lib/msAuth.js); the client learns its own status from /api/me at startup (window.PF_isPM),
  // so no employee list is shipped in the page. Tab visibility is cosmetic — every PM API is gated
  // server-side regardless.
  function isPM() { return window.PF_isPM === true; }

  // top-level render dispatch by current view
  var lastRenderedView = null;
  var lastColDepth = 1; // drill depth last render (1 tree / 2 +models / 3 +resources)
  var mobileRevealPending = false; // set by a product/model TAP so mobile scrolls to the opened block even on a same-depth switch
  function render() {
    // Arriving at a view fresh should animate it in; only clicks WITHIN the catalog (which call
    // renderColumns directly, not render) keep the "animate just the new column/rows" scoping.
    if (view !== lastRenderedView) { lastAnimKeys = null; lastRenderedView = view; }
    if (view === "progress") renderCoverage();
    else if (view === "admin") renderAdmin();
    else if (view === "metrics") renderMetrics();
    else if (view === "styleguide") renderStyleguide();
    else renderColumns();
    syncViewTabs();
    setDocTitle();
  }
  // The document title never changed (WCAG 2.4.2): every view, deep link and search shared "Grumbleton
  // Product Hub", so browser history entries, a row of open tabs, and a screen reader's page announcement
  // were all indistinguishable. Name what's on screen, most specific part first.
  function setDocTitle(searchQuery) {
    var base = "Grumbleton & Sons Hub", lead = "";
    if (searchQuery) lead = "Search: " + searchQuery;
    else if (view === "progress") lead = "Progress";
    else if (view === "metrics") lead = "Metrics";
    else if (view === "admin") lead = "Admin";
    else if (sel.model) lead = String(sel.model).split("§").pop();
    else if (sel.prod) lead = String(sel.prod).split("§").pop();
    try { document.title = lead ? (lead + " · " + base) : base; } catch (e) {}
  }
  function syncViewTabs() {
    var tabs = document.querySelectorAll(".vtab");
    for (var i = 0; i < tabs.length; i++) {
      var on = tabs[i].getAttribute("data-view") === view;
      tabs[i].classList.toggle("on", on);
      tabs[i].setAttribute("aria-selected", on ? "true" : "false");
      tabs[i].tabIndex = on ? 0 : -1; // roving tabindex: only the selected tab is in the tab order
      if (on) { var c = document.getElementById("content"); if (c) c.setAttribute("aria-labelledby", tabs[i].id); }
    }
  }

  function svg(inner) { return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + inner + '</svg>'; }
  var CHEV = svg('<path d="m9 6 6 6-6 6"/>');
  // "Global Project Intake Form" opens the company Smartsheet intake form (button in the Find header).
  // Same form previously labeled "Request a resource" — confirmed unchanged URL, just renamed (2026-08-07).
  var REQ_FORM_URL = "/demo/elsewhere.html";
  var REQICON = svg('<path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M17 21H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h7l5 5v11a2 2 0 0 1-2 2z"/><path d="M12 11v6M9 14h6"/>');
  function openReqForm() { window.open(REQ_FORM_URL, "_blank", "noopener"); }
  // A search that found nothing → open the intake form with the searched text prefilled into
  // the form's "Additional Details" field, so the request carries exactly what
  // staff were looking for. Request type is left blank — a miss could be a product OR a collateral
  // request, so the requester picks. Requester resolves to the signed-in user. Unknown fields are
  // ignored by the form, so this can only ever help.
  function searchRequestFormUrl(query) {
    var q = String(query || "").trim();
    var params = [];
    if (window.PF_user) params.push(["Requester", window.PF_user]);
    params.push(["Additional Details", (q ? 'Searched Product Hub for "' + q + '" — no match found. ' : "") + "Requested via Product Hub."]);
    return REQ_FORM_URL + "?" + params.map(function (p) { return encodeURIComponent(p[0]) + "=" + encodeURIComponent(p[1]); }).join("&");
  }
  function openSearchRequest(query) { window.open(searchRequestFormUrl(query), "_blank", "noopener"); }
  // A click on a "Work in progress" resource → the intake form prefilled with the EXACT model + resource the
  // user needs (they navigated to the precise slot, so this is the strongest possible demand signal). Same
  // form + safe unknown-field handling as searchRequestFormUrl.
  function resourceRequestFormUrl(model, resource) {
    var params = [];
    if (window.PF_user) params.push(["Requester", window.PF_user]);
    params.push(["Additional Details", "Requesting " + resource + " for " + model + " — not yet in Product Hub. Requested via Product Hub."]);
    return REQ_FORM_URL + "?" + params.map(function (p) { return encodeURIComponent(p[0]) + "=" + encodeURIComponent(p[1]); }).join("&");
  }

  /* ============================================================
     CATALOG ACCESSORS — link lookup, live/usable counts, selection, tree columns
     ============================================================ */
  // ---- shared async helpers ----------------------------------------------------------------------------
  // One door for every same-origin /api call: the token, the headers, a JSON body, a timeout and ONE error
  // idiom (22 copies of the token dance and three different error extractions used to live inline).
  // Resolves with the parsed JSON body ({} when there is none). Rejects with an Error whose message is the
  // server's `error` when it sent one and "api <status>" otherwise — a Vercel HTML 502 page is not JSON,
  // so the status is the fallback — and whose .status carries the HTTP code for callers that route a 403
  // differently from an outage. opts: { method, body (object → JSON), fresh (cache: no-store), keepalive,
  // timeout (ms; default 35s, just past the longest server budget so a slow Smartsheet write is never
  // abandoned client-side while it commits) }.
  function apiFetch(path, opts) {
    opts = opts || {};
    if (!window.PF_getApiToken) return Promise.reject(new Error("Please sign in again."));
    return window.PF_getApiToken().then(function (token) {
      var init = { method: opts.method || "GET", headers: { Authorization: "Bearer " + token } };
      if (opts.body !== undefined) { init.headers["Content-Type"] = "application/json"; init.body = JSON.stringify(opts.body); }
      if (opts.fresh) init.cache = "no-store";
      if (opts.keepalive) init.keepalive = true;
      return fetchTimeout(path, init, opts.timeout || 35000);
    }).then(function (r) {
      if (r.ok) return r.json().catch(function () { return {}; });
      return r.json().then(function (j) { return j || {}; }, function () { return {}; }).then(function (j) {
        // `code` tells a compare-and-swap refusal ("conflict") from a shrink-guard one ("would_shrink").
        var err = new Error(j.error || ("api " + r.status)); err.status = r.status; if (j.code) err.code = j.code; throw err;
      });
    });
  }
  // Disable a button for the life of an async action and put it back EXACTLY as it was. Returns the restore
  // function; call it on both the success and the failure path (a button left dead after a failed save was
  // a real bug — T1-14 — because each site hand-rolled its own pair and one path forgot).
  function busyButton(btn, label) {
    if (!btn) return function () {};
    var html = btn.innerHTML, was = btn.disabled;
    btn.disabled = true; if (label != null) btn.textContent = label;
    return function () { btn.disabled = was; btn.innerHTML = html; };
  }
  // Generation guard for a panel that can be reloaded while an earlier load is still in flight: only the
  // NEWEST load may paint, or a slower earlier response clobbers the list with stale rows (T1-08).
  //   var live = loadGen(box); …fetch…; if (!live()) return; paint();
  function loadGen(box) { var gen = (box._loadGen = (box._loadGen || 0) + 1); return function () { return box._loadGen === gen; }; }

  function keyOf(arr) { return arr.join("§"); }
  function iconFor(rname) { for (var i = 0; i < RESOURCE_TYPES.length; i++) if (RESOURCE_TYPES[i].name === rname) return RESOURCE_TYPES[i].icon; return "spec"; }

  function normLinks(v) {
    if (!v) return [];
    var arr = typeof v === "string" ? [{ label: null, url: v }] : v;
    // Sanitize every catalog URL at this single choke-point: names come from a staff-edited
    // Smartsheet, so a "SP Folder URL" value could be javascript:/data:. safeUrl() strips those,
    // and this is the only place catalog links are read for href/window.open/Graph resolution.
    // DROP any value safeUrl() blanked (non-http(s), e.g. a scheme-less/typo'd link) so it can't
    // masquerade as a real link — a phantom "Open" that opens nothing + inflated coverage. Mirrors
    // the server's normLinks (lib/linkcheck.js), which also filters empties.
    return arr.map(function (x) { return { label: x.label, url: safeUrl(x.url) }; }).filter(function (x) { return x.url; });
  }
  function linksFor(model, rname) { return normLinks(LINKS[model + "::" + rname]); }
  // The ONE writer for an optimistic LINKS update after a successful write. LINKS carries the catalog's own
  // shape — a bare string for a single link, [{label,url}] for several — and normLinks above understands
  // nothing else. Four call sites used to inline this rule and one of them stored bare strings, which
  // normLinks silently dropped to []: the resource read as EMPTY in that tab (coverage fell, the row showed
  // "Work in progress") and the next write planned from nothing and was refused by the server's shrink
  // guard until a reload. One helper makes that whole class of drift impossible.
  // Each entry may be a bare url string OR a { label, url }. Keeping the label matters because a
  // multi-file resource renders one chip PER FILE using it, so dropping it made three renders read
  // "Renders, Renders, Renders" until the next catalog load.
  function setLinks(key, urls) {
    var arr = (urls || []).map(function (u) {
      if (!u) return null;
      return (typeof u === "string") ? { label: null, url: u } : { label: u.label || null, url: u.url };
    }).filter(function (x) { return x && x.url; });
    if (!arr.length) { delete LINKS[key]; return; }
    // The catalog's own shape: one link is a bare string, several are [{label,url}] (see normLinks).
    LINKS[key] = arr.length > 1 ? arr : arr[0].url;
  }
  // "Usable" = has a link that actually resolves to files. A wired link the last Link-health run
  // found EMPTY (or BROKEN) is treated as not-there in the Find view, so a resource with no real
  // files behaves consistently with one that was never linked ("Work in progress").
  // N/A wins over a stray link: a resource marked not-applicable never counts as usable-linked, even if a
  // link was left on it. Without this, the Find count (usableCount) included an N/A cell that applicableCount
  // (the denominator) excluded, so a model could read "3 of 2 linked" and disagree with the Progress matrix
  // (covLinked already skips N/A). Guarding here aligns every counter + the Live/segment status in one place.
  function usableLinked(model, rname) { if (naFor(model, rname)) return false; if (!linksFor(model, rname).length) return false; var s = healthStatusFor(model, rname); return s !== "empty" && s !== "broken"; }
  function usableCount(model) { return RESOURCE_TYPES.reduce(function (a, r) { return a + (usableLinked(model, r.name) ? 1 : 0); }, 0); }
  // A resource type a model marked "not applicable" (via the sheet) is excluded from the coverage
  // denominator — so "% complete" / the finish gate reflect only the resources the model actually needs.
  function naFor(model, rname) { return !!NA[model + "::" + rname]; }
  function applicableCount(model) { return RESOURCE_TYPES.reduce(function (a, r) { return a + (naFor(model, r.name) ? 0 : 1); }, 0); }
  // "Live" = has at least one USABLE resource (a link the last health run found to resolve to real files),
  // so the family/product dots agree with the coverage matrix, the % and the model-column dot (which all
  // use usable-only). A model whose only link is broken/empty reads "In progress", not a green "Live" dot.
  function modelLive(m) { return usableCount(m) > 0; }
  function productContent(p) { return p.models.length > 0; }
  function productLive(p) { return p.models.some(modelLive); }
  function familyContent(f) { return f.products.some(productContent); }
  function familyLive(f) { return f.products.some(productLive); }
  function dot(live) { return '<span class="' + (live ? "okdot" : "mdot") + '"><span class="sr-only">' + (live ? "Live" : "Work in progress") + '</span></span>'; }

  function famByKey(k) { if (!k) return null; var p = k.split("§"); for (var i = 0; i < CATALOG.length; i++) if (CATALOG[i].name === p[0]) { var c = CATALOG[i]; for (var j = 0; j < c.families.length; j++) if (c.families[j].name === p[1]) return { c: c, f: c.families[j] }; } return null; }
  function prodByKey(k) { var fb = famByKey(k); if (!fb) return null; var p = k.split("§"); for (var j = 0; j < fb.f.products.length; j++) if (fb.f.products[j].name === p[2]) return { c: fb.c, f: fb.f, p: fb.f.products[j] }; return null; }

  function head(eyebrow, title, lead) {
    return '<div class="level-head"><div class="eyebrow">' + eyebrow + '</div><h1>' + title + '</h1>' + (lead ? '<p class="lead">' + lead + '</p>' : '') + '</div>';
  }

  // ---- selection ----
  // openFams = which families are expanded (multiple allowed); sel.prod/sel.model open columns to the right
  function famOfKey(k) { var p = String(k).split("§"); return keyOf([p[0], p[1]]); }
  function toggleFamily(k) {
    if (openFams[k]) {
      delete openFams[k];
      if (sel.prod && famOfKey(sel.prod) === k) { sel.prod = null; sel.model = null; } // close columns of a collapsed family
      flashKey = null;
    } else { openFams[k] = true; flashKey = k; }
    pendingFocusKey = k; renderColumns(); syncHash();
  }
  function selectProduct(k) { sel.prod = k; sel.model = null; openFams[famOfKey(k)] = true; flashKey = k; pendingFocusKey = k; mobileRevealPending = true; renderColumns(); syncHash(); }
  function selectModel(k) { sel.model = k; flashKey = k; pendingFocusKey = k; mobileRevealPending = true; renderColumns(); syncHash(); }
  // Return to the top-level Find home (nothing selected) — where "Your Favorites" lives.
  function goHome() {
    sel = { prod: null, model: null }; openFams = {}; flashKey = null; view = "catalog"; _preSearchView = null;
    if (searchInput) searchInput.value = "";
    render(); syncHash(); resetScroll();
  }

  // ---- column item ----
  function colItem(o) {
    var el = document.createElement(o.onClick ? "button" : "div");
    if (o.onClick) el.type = "button";
    el.className = "col-item" + (o.selected ? " sel" : "") + (o.soon ? " soon" : "") + (o.onClick ? "" : " static");
    if (o.key) el.setAttribute("data-key", o.key);
    if (o.title) el.setAttribute("title", o.title);
    if (o.onClick) el.setAttribute("aria-current", o.selected ? "true" : "false");
    el.innerHTML =
      '<span class="ci-name' + (o.mono ? " mono" : "") + '">' + escapeHtml(o.label) + (o.suffix || "") + '</span>' +
      (o.meta ? '<span class="ci-meta">' + o.meta + '</span>' : '') +
      (o.chev ? '<span class="ci-chev">' + CHEV + '</span>' : '');
    if (o.onClick) el.addEventListener("click", o.onClick);
    return el;
  }
  function column(cls, headHtml, headClass) {
    var col = document.createElement("div"); col.className = "col" + (cls ? " " + cls : "");
    col.innerHTML = '<div class="col-head' + (headClass ? " " + headClass : "") + '"' + (headHtml ? ' role="heading" aria-level="2"' : '') + '>' + headHtml + '</div>';
    var body = document.createElement("div"); body.className = "col-body";
    col.appendChild(body); col._body = body; return col;
  }

  // "Other Collateral" isn't a product category — it's channel-based collateral (Convenience/Grocery, not
  // products), shown in its own full-width collapsible panel above the breadcrumb, so it's excluded here.
  var OTHER_COLLATERAL = "Other Collateral";

  // Left panel: families as a list; the open family expands its products underneath.
  function familyTreePanel() {
    var stack = document.createElement("div"); stack.className = "col-tree"; var nestSeq = 0;
    CATALOG.forEach(function (c) {
      if (c.name === OTHER_COLLATERAL) return; // rendered separately as a channel panel (see otherCollateralPanel)
      var box = document.createElement("div"); box.className = "cat-box"; box.setAttribute("data-anim", "cat:" + c.name);
      var head = document.createElement("div"); head.className = "cat-box-head"; head.textContent = c.name;
      head.setAttribute("role", "heading"); head.setAttribute("aria-level", "2");
      box.appendChild(head);
      var body = document.createElement("div"); body.className = "col-body"; box.appendChild(body);
      c.families.forEach(function (f) {
        var fk = keyOf([c.name, f.name]), fOpen = !!openFams[fk], content = familyContent(f);
        var fam = document.createElement("button"); fam.type = "button";
        fam.className = "col-item fam-row" + (fOpen ? " open" : "") + (content ? "" : " soon");
        fam.setAttribute("data-key", fk); fam.setAttribute("aria-expanded", fOpen ? "true" : "false");
        fam.innerHTML =
          '<span class="ci-name">' + escapeHtml(f.name) + '</span>' +
          '<span class="ci-meta">' + (content ? dot(familyLive(f)) : '<span class="soon-tag">Work in progress</span>') + '</span>' +
          '<span class="ci-chev">' + CHEV + '</span>';
        fam.addEventListener("click", function () { toggleFamily(fk); });
        body.appendChild(fam);
        if (fOpen) {
          var nestId = "fnest-" + (nestSeq++);
          fam.setAttribute("aria-controls", nestId);
          var nest = document.createElement("div"); nest.className = "tree-nest"; nest.id = nestId;
          if (f.products.length === 0) {
            nest.innerHTML = '<div class="col-empty" data-anim="prowempty:' + escapeAttr(fk) + '">No products in this family yet — work in progress.</div>';
          } else {
            f.products.forEach(function (p, i) {
              var pk = keyOf([c.name, f.name, p.name]), pc = productContent(p);
              var prod = document.createElement("button"); prod.type = "button";
              prod.className = "col-item prod-row" + (sel.prod === pk ? " sel" : "") + (pc ? "" : " soon");
              // Expose the selection: the highlight alone measures 1.18:1 in the light theme, so a
              // low-vision user and a screen-reader user both had no reliable cue about which product is
              // open. (The Models column already sets this.)
              if (sel.prod === pk) prod.setAttribute("aria-current", "true");
              prod.setAttribute("data-key", pk); prod.setAttribute("data-anim", "prow:" + pk); prod.style.setProperty("--i", i);
              prod.innerHTML =
                '<span class="ci-name">' + escapeHtml(p.name) + '</span>' +
                '<span class="ci-meta">' + (pc ? dot(productLive(p)) + p.models.length : '<span class="soon-tag">Work in progress</span>') + '</span>' +
                '<span class="ci-chev">' + CHEV + '</span>';
              prod.addEventListener("click", function () { selectProduct(pk); });
              nest.appendChild(prod);
            });
          }
          body.appendChild(nest);
        }
      });
      stack.appendChild(box);
    });
    return stack;
  }

  // Full-width, collapsed-by-default "Other Collateral" panel shown ABOVE the breadcrumb. Unlike the
  // product tree, its contents are CHANNELS (Convenience, Grocery — sheet "families" under this category),
  // not products/models. Expanding the panel lists the channels vertically; expanding a channel reveals its
  // collateral inline (or a "no collateral yet" note). Native <details> gives free collapse + keyboard a11y.
  function otherCollateralPanel() {
    var C = null;
    for (var i = 0; i < CATALOG.length; i++) if (CATALOG[i].name === OTHER_COLLATERAL) { C = CATALOG[i]; break; }
    if (!C || !C.families.length) return null;
    var det = document.createElement("details"); det.className = "oc-panel"; det.setAttribute("data-anim", "oc");
    var sum = document.createElement("summary"); sum.className = "oc-head";
    sum.innerHTML = '<span class="oc-title">' + escapeHtml(C.name) + '</span><span class="oc-chev">' + CHEV + '</span>';
    det.appendChild(sum);
    var body = document.createElement("div"); body.className = "oc-body";
    C.families.forEach(function (ch) {
      var has = familyContent(ch); // any linked collateral under this channel yet?
      var cd = document.createElement("details"); cd.className = "oc-channel";
      var cs = document.createElement("summary"); cs.className = "oc-channel-head";
      cs.innerHTML =
        '<span class="oc-ch-name">' + escapeHtml(ch.name) + '</span>' +
        '<span class="oc-ch-meta">' + (has ? dot(true) : '<span class="soon-tag">Work in progress</span>') + '</span>' +
        '<span class="oc-chev">' + CHEV + '</span>';
      cd.appendChild(cs);
      var cb = document.createElement("div"); cb.className = "oc-channel-body";
      if (!has) {
        cb.innerHTML = '<div class="oc-empty">No collateral in this channel yet.</div>';
      } else {
        // Flat list of every linked document under the channel (channels hold collateral directly, not a
        // product/model hierarchy). Each opens in the same viewer as any other resource.
        collectChannelCollateral(ch).forEach(function (it) {
          var b = document.createElement("button"); b.type = "button"; b.className = "oc-doc";
          b.innerHTML = '<span class="oc-doc-ico">' + svg(ICONS[iconFor(it.resource)]) + '</span><span class="oc-doc-name">' + escapeHtml(it.resource) + (it.links.length > 1 ? ' (' + it.links.length + ')' : '') + '</span><span class="oc-doc-open">' + svg(ICONS.ext) + '</span>';
          b.addEventListener("click", function () { openResource(it.links, it.resource, it.model); });
          cb.appendChild(b);
        });
      }
      cd.appendChild(cb);
      body.appendChild(cd);
    });
    det.appendChild(body);
    return det;
  }
  // Every linked resource under a channel, flattened (whatever product/model rows the sheet uses to carry it).
  function collectChannelCollateral(ch) {
    var out = [];
    (ch.products || []).forEach(function (p) {
      (p.models || []).forEach(function (m) {
        RESOURCE_TYPES.forEach(function (r) {
          var links = linksFor(m, r.name);
          if (links && links.length) out.push({ model: m, resource: r.name, links: links });
        });
      });
    });
    return out;
  }

  /* ============================================================
     MODEL NAMES & FIND — suffix decode, filter/sort toolbar, model + resource columns
     ============================================================ */
  // Model-name suffix decode — a company-wide naming convention. Only genuine coded models
  // (prefix HMX/GTX/SNX + number + code) are decoded, so descriptive names like "…for Sir Reginald"
  // are never mis-matched. Codes combine; only "PO" is two letters, plus a leading "Mini".
  var MODEL_CODES = { X: "Extra Loot", F: "Fireproof", D: "Double Hoard", G: "Gem-Safe" };
  var MODEL_CODE_ORDER = ["I", "D", "R", "PO", "W", "C", "L", "T", "MINI"];
  function modelCodeLabel(code) { return code === "MINI" ? "Mini" : code; }
  function decodeModelName(name) {
    var m = /^(?:Hoardmaster|Gemtumbler)\s+\d+\s+((?:Mini\s+)?[A-Za-z]+)$/i.exec(String(name || "").trim());
    if (!m) return null;
    var rest = m[1].trim(), tokens = [];
    var mini = /^Mini\s+/i.exec(rest);
    if (mini) { tokens.push("MINI"); rest = rest.slice(mini[0].length); }
    rest = rest.toUpperCase();
    for (var i = 0; i < rest.length;) {
      if (rest.slice(i, i + 2) === "PO") { tokens.push("PO"); i += 2; }
      else { tokens.push(rest.charAt(i)); i += 1; }
    }
    if (!tokens.length) return null;
    var parts = tokens.map(function (t) { return { code: t, label: MODEL_CODES[t] }; });
    if (parts.some(function (p) { return !p.label; })) return null; // unknown token → don't guess
    return parts;
  }
  // Inline decode shown right after a coded model name (same wording/order as the hover tooltip),
  // e.g. " <span…>(Integrated Dividers · Center · Left Pusher)</span>". Returns "" for names that don't decode.
  function decodeSuffixHtml(name) {
    var dec = decodeModelName(name);
    if (!dec) return "";
    var labels = dec.map(function (p) { return p.label; });
    // "Pusher" should read once, on the LAST label that carries it — not repeated per code
    // (e.g. Center · Left · T Pusher, not "Center Pusher · Left Pusher · T Pusher").
    var last = -1;
    labels.forEach(function (l, i) { if (/ Pusher$/.test(l)) last = i; });
    if (last !== -1) labels = labels.map(function (l, i) { return (i !== last) ? l.replace(/ Pusher$/, "") : l; });
    return ' <span class="ci-decode">(' + escapeHtml(labels.join(" · ")) + ')</span>';
  }

  // ---- Find filter/sort toolbar (model column + search results) ----
  // The pure filter/sort logic lives in PF.filterSortModels (pf-shared.js); the app supplies accessors
  // built from its own closure helpers, so nothing about the sheet/catalog leaks into the public shared file.
  var TOOLBAR_MIN = 4; // a list this short doesn't warrant a filter bar (tunable)
  var searchFilter = { status: "any", attrs: [], recent: false, sort: "" };
  // Short, chip-sized labels for the attribute facets (full wording is the title/tooltip via MODEL_CODES).
  var MODEL_CODE_CHIP = { X: "Extra Loot", F: "Fireproof", D: "Double", G: "Gem-Safe" };
  function newFilterState() { return { status: "any", attrs: [], recent: false, sort: "" }; }
  function filterActive(st) { return st.status !== "any" || st.attrs.length > 0 || st.recent; }
  // Newest "added" timestamp (ms) across a model's resources, or null if none is dated (drives Newest sort
  // + the Recently-added facet, reusing the same date data as the What's New strip).
  function modelAddedAt(m) {
    var best = null;
    RESOURCE_TYPES.forEach(function (r) {
      var s = ADDED[m + "::" + r.name]; if (!s) return;
      var t = dayStart(s); if (isNaN(t)) return;
      if (best === null || t > best) best = t;
    });
    return best;
  }
  function modelAttrs(m) { var d = decodeModelName(m); return d ? d.map(function (p) { return p.code; }) : []; }
  // Accessors over a MODEL-NAME string (model column).
  function modelAccessors() {
    return {
      hasFiles: function (m) { return usableCount(m) > 0; },
      attrs: modelAttrs,
      addedAt: modelAddedAt,
      completeness: usableCount,
      name: function (m) { return m; },
      recentCutoff: Date.now() - RECENT_DAYS * 864e5
    };
  }
  // Accessors over a search-hit row (its model name is n.model || n.name).
  function hitAccessors() {
    var a = modelAccessors();
    return {
      hasFiles: function (n) { return a.hasFiles(n.model || n.name); },
      attrs: function (n) { return a.attrs(n.model || n.name); },
      addedAt: function (n) { return a.addedAt(n.model || n.name); },
      completeness: function (n) { return a.completeness(n.model || n.name); },
      name: function (n) { return n.name; },
      recentCutoff: a.recentCutoff
    };
  }
  // Build the filter/sort bar. `availCodes` = attribute codes present in the pool; `count` = {shown,total,noun};
  // `defaultLabel` = label for the "" (unsorted) option; `rerender` repaints the surface in place.
  function findToolbar(st, availCodes, count, defaultLabel, rerender) {
    var bar = document.createElement("div"); bar.className = "ftbar";
    // A radiogroup, not a row of toggles: these three are mutually exclusive, but aria-pressed announced
    // each as an independent on/off switch, so a screen-reader user couldn't tell that choosing one
    // releases the others.
    var seg = document.createElement("div"); seg.className = "ft-seg"; seg.setAttribute("role", "radiogroup"); seg.setAttribute("aria-label", "Filter by status");
    [["any", "All"], ["has", "Has files"], ["wip", "Not in the Hub yet"]].forEach(function (o) {
      var on = st.status === o[0];
      var b = document.createElement("button"); b.type = "button"; b.className = "ft-seg-btn" + (on ? " on" : "");
      b.textContent = o[1];
      b.setAttribute("role", "radio"); b.setAttribute("aria-checked", on ? "true" : "false");
      b.tabIndex = on ? 0 : -1; // roving tabindex: the whole group is one Tab stop; arrows move within it
      b.addEventListener("click", function () { st.status = o[0]; rerender(); });
      seg.appendChild(b);
    });
    seg.addEventListener("keydown", function (e) {
      var step = { ArrowRight: 1, ArrowDown: 1, ArrowLeft: -1, ArrowUp: -1 }[e.key];
      if (!step) return;
      var btns = [].slice.call(seg.querySelectorAll(".ft-seg-btn"));
      var i = btns.indexOf(document.activeElement);
      if (i === -1) return;
      e.preventDefault();
      var next = btns[(i + step + btns.length) % btns.length];
      if (next) next.click(); // selection follows focus — the standard radiogroup behaviour
    });
    bar.appendChild(seg);
    if (availCodes.length) {
      var chips = document.createElement("div"); chips.className = "ft-chips"; chips.setAttribute("role", "group"); chips.setAttribute("aria-label", "Filter by attribute");
      availCodes.forEach(function (c) {
        var on = st.attrs.indexOf(c) !== -1;
        var b = document.createElement("button"); b.type = "button"; b.className = "ft-chip" + (on ? " on" : "");
        b.textContent = MODEL_CODE_CHIP[c] || c; b.title = MODEL_CODES[c] || c;
        b.setAttribute("aria-pressed", on ? "true" : "false");
        b.addEventListener("click", function () { var i = st.attrs.indexOf(c); if (i === -1) st.attrs.push(c); else st.attrs.splice(i, 1); rerender(); });
        chips.appendChild(b);
      });
      bar.appendChild(chips);
    }
    var rec = document.createElement("button"); rec.type = "button"; rec.className = "ft-chip ft-recent" + (st.recent ? " on" : "");
    rec.textContent = "Recently added"; rec.setAttribute("aria-pressed", st.recent ? "true" : "false");
    rec.addEventListener("click", function () { st.recent = !st.recent; rerender(); });
    bar.appendChild(rec);
    var sortWrap = document.createElement("label"); sortWrap.className = "ft-sort";
    sortWrap.appendChild(Object.assign(document.createElement("span"), { className: "ft-sort-lbl", textContent: "Sort" }));
    var ssel = document.createElement("select"); ssel.className = "ft-sort-sel"; ssel.setAttribute("aria-label", "Sort");
    [["", defaultLabel], ["name", "Name (A–Z)"], ["complete-desc", "Most complete"], ["complete-asc", "Least complete"], ["newest", "Newest"]].forEach(function (o) {
      var op = document.createElement("option"); op.value = o[0]; op.textContent = o[1]; if (st.sort === o[0]) op.selected = true; ssel.appendChild(op);
    });
    ssel.addEventListener("change", function () { st.sort = ssel.value; rerender(); });
    sortWrap.appendChild(ssel); bar.appendChild(sortWrap);
    var right = document.createElement("div"); right.className = "ft-right";
    // Only when a filter has narrowed the set: unfiltered, this repeated the lead sentence word
    // for word two lines above it ("11 results" under "11 results for “…”").
    if (count.shown !== count.total) {
      var cnt = document.createElement("span"); cnt.className = "ft-count";
      cnt.textContent = count.shown + " of " + count.total + " " + count.noun;
      right.appendChild(cnt);
    }
    if (filterActive(st) || st.sort) {
      var clr = document.createElement("button"); clr.type = "button"; clr.className = "ft-clear"; clr.textContent = "Clear";
      clr.addEventListener("click", function () { st.status = "any"; st.attrs = []; st.recent = false; st.sort = ""; rerender(); });
      right.appendChild(clr);
    }
    bar.appendChild(right);
    return bar;
  }

  function modelColumn() {
    var pb = prodByKey(sel.prod); if (!pb) return null;
    var col = column("col-models", escapeHtml(pb.p.name), "col-lead"); col.setAttribute("data-anim", "mcol:" + sel.prod);
    // No product-wide "Copy all N links" here (removed 2026-09-14, owner decision): a rep sends files for ONE
    // model, and the model panel now lets them pick which — a whole-product dump was "too much".
    if (pb.p.models.length === 0) {
      col._body.innerHTML = '<div class="col-empty">No models yet — work in progress.</div>';
      return col;
    }
    var usedCodes = {};
    pb.p.models.forEach(function (m) {
      var k = keyOf([pb.c.name, pb.f.name, pb.p.name, m]), lc = usableCount(m);
      var dec = decodeModelName(m);
      if (dec) { dec.forEach(function (p) { usedCodes[p.code] = true; }); }
      col._body.appendChild(colItem({
        key: k, label: m, mono: true, selected: sel.model === k, chev: true, suffix: decodeSuffixHtml(m),
        meta: dot(lc > 0) + lc + "/" + applicableCount(m),
        onClick: function () { selectModel(k); }
      }));
    });
    // Collapsible key showing only the codes this product's models actually use.
    var codes = MODEL_CODE_ORDER.filter(function (c) { return usedCodes[c]; });
    if (codes.length) {
      var det = document.createElement("details"); det.className = "model-key"; det.setAttribute("data-anim", "modelkey:" + sel.prod);
      det.innerHTML = '<summary>Model name key</summary><dl>' + codes.map(function (c) {
        return '<div><dt>' + escapeHtml(modelCodeLabel(c)) + '</dt><dd>' + escapeHtml(MODEL_CODES[c]) + '</dd></div>';
      }).join("") + '</dl>';
      col._body.appendChild(det); // key sits UNDER the models, not above them
    }
    return col;
  }

  function resourceColumn() {
    var p = sel.model.split("§"), m = p[p.length - 1];
    var col = column("col-res", '<span class="lead-name">' + escapeHtml(m) + decodeSuffixHtml(m) + '</span>', "col-lead mono"); col.setAttribute("data-anim", "rcol:" + sel.model);
    col._body.appendChild(buildResources(m, col.querySelector(".col-head")));
    return col;
  }

  /* ============================================================
     DATES, FAVORITES & RECENTLY VIEWED
     ============================================================ */
  // Resource "currency": turn a resource's added/updated date (ADDED, from the newest file in its linked
  // folder or the sheet row's last edit) into a human "Updated …" label, so a user can trust they're
  // opening something current (recent = relative, older = month + year).
  // A date-only "YYYY-MM-DD" from the sheet names a CALENDAR DAY, not an instant. Parsing it as UTC
  // midnight (what appending "Z" does) puts it in the past for anyone west of Greenwich — so at 8pm
  // Central a file added TODAY read "Updated yesterday", while What's New parsed local midnight and
  // disagreed on the same file. One parser, local midnight, shared by both.
  function dayStart(s) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s || ""));
    if (!m) { var t = Date.parse(String(s || "")); return isNaN(t) ? NaN : t; }
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getTime();
  }
  // Whole days between two instants, counted on the local CALENDAR — so "yesterday" means the previous
  // date, not 24 hours ago.
  function daysBetween(fromMs, toMs) {
    var a = new Date(fromMs), b = new Date(toMs);
    var d1 = new Date(a.getFullYear(), a.getMonth(), a.getDate()).getTime();
    var d2 = new Date(b.getFullYear(), b.getMonth(), b.getDate()).getTime();
    return Math.round((d2 - d1) / 86400000);
  }
  function fmtUpdated(dateStr) {
    var t = dayStart(dateStr); if (isNaN(t)) return "";
    var days = daysBetween(t, Date.now());
    if (days <= 0) return "Updated today";
    if (days === 1) return "Updated yesterday";
    if (days < 30) return "Updated " + days + " days ago";
    var d = new Date(t), MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return "Updated " + MON[d.getMonth()] + " " + d.getFullYear();
  }
  function resUpdatedLabel(m, r) { var s = ADDED[m + "::" + r]; return s ? fmtUpdated(s) : ""; }
  // A linked resource whose newest update is older than the staleness window (12 months) is flagged for
  // "review" — likely superseded. Only meaningful for resources that actually have a date.
  var STALE_DAYS = 365;
  function resIsStale(m, r) {
    var s = ADDED[m + "::" + r]; if (!s) return false;
    var t = dayStart(s); if (isNaN(t)) return false;   // same local-midnight parse as fmtUpdated
    return daysBetween(t, Date.now()) > STALE_DAYS;
  }
  // Favorites: pinned models, synced across a user's devices via /api/me (keyed to their sign-in),
  // with a localStorage mirror for instant UI + an offline fallback. localStorage is written first (the
  // pin shows immediately), then the change writes through to the server; a failed write leaves the local
  // pin in place and next load's merge re-uploads it, so a Smartsheet hiccup can't lose a favorite.
  function favGet() { try { var a = JSON.parse(localStorage.getItem("pf-favs") || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
  function favHas(m) { return favGet().indexOf(m) !== -1; }
  // Tombstones: models the user un-pinned. Because the cross-device merge is a union, a removal whose
  // delete-write failed/lagged would be resurrected on the next load; recording it here lets the merge and
  // sync subtract it until the server confirms the delete. Re-pinning clears the tombstone.
  function favRemoved() { try { var a = JSON.parse(localStorage.getItem("pf-favs-removed") || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
  function favSetRemoved(a) { try { localStorage.setItem("pf-favs-removed", JSON.stringify((a || []).slice(0, 50))); } catch (e) {} }
  // What the SERVER last reported for this user. Lets favSyncFromServer tell a genuine local-only pin (never
  // seen on the server → push it up) from a stale copy of a pin that was UN-pinned on another device (was
  // seen, now gone → drop it, don't resurrect it). Without this, the "push local-only up" step re-added an
  // un-pinned favorite to the server every time a second device loaded.
  function favServerSeen() { try { var a = JSON.parse(localStorage.getItem("pf-favs-server-seen") || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
  function favSetServerSeen(a) { try { localStorage.setItem("pf-favs-server-seen", JSON.stringify((a || []).slice(0, 60))); } catch (e) {} }
  function favMarkRemoved(m, removed) {
    var a = favRemoved(), i = a.indexOf(m);
    if (removed && i === -1) a.push(m);            // record the removal
    else if (!removed && i !== -1) a.splice(i, 1); // re-pinned → clear the tombstone
    favSetRemoved(a);
  }
  var FAV_MAX = 24; // the strip's capacity; the server keeps its own, larger, per-user cap
  function favToggle(m) {
    var a = favGet(), i = a.indexOf(m), adding = i === -1;
    // Refuse the 25th instead of silently dropping the oldest. The slice below used to evict it LOCALLY
    // with no matching server remove, so the evicted pin returned on the next load and a different one
    // fell off — favorites appeared to shuffle by themselves.
    if (adding && a.length >= FAV_MAX) {
      showError('<span>You’ve pinned the maximum of ' + FAV_MAX + ' favorites — remove one first.</span>', 6000);
      return false;
    }
    if (adding) a.unshift(m); else a.splice(i, 1);
    try { localStorage.setItem("pf-favs", JSON.stringify(a.slice(0, FAV_MAX))); } catch (e) {}
    favMarkRemoved(m, !adding); // tombstone a removal so a failed delete-write isn't resurrected on reload; re-pinning clears it
    favSync(adding ? "add" : "remove", m); // write through to the cross-device store (best-effort)
    return adding;
  }
  // Fire-and-forget write-through to the per-user store (POST /api/me {action, model}). A failure is
  // non-fatal — the local pin still works on this device, and favSyncFromServer re-uploads any local-only
  // pins on the next load. (Favorites ride on /api/me, not a separate function, to stay within the Hobby
  // plan's 12-function cap.)
  function favSync(action, model) {
    if (!window.PF_getApiToken) return;
    return apiFetch("/api/me", { method: "POST", body: { action: action, model: model } })
      // A 4xx (a rejected model name, a server-side cap) used to look like success and was re-sent on every
      // load. Don't interrupt the user — the local pin still works and the next load's merge retries — but
      // say so in the console instead of letting it vanish.
      .catch(function (e) { if (e && e.status) console.warn("[pf] favorite " + action + " rejected (" + e.status + ")"); });
  }
  // On load: reconcile this browser's favorites with the server-synced set (delivered by /api/me as
  // PF_serverFavs). Union them, upload any local-only pins (one-time migration), and mirror the merged set
  // to localStorage. If /me didn't return favorites (offline / not signed in) -> keep local, non-fatal.
  function favSyncFromServer() {
    if (!Array.isArray(PF_serverFavs)) return; // /me gave us no favorites this load — keep the local mirror
    var server = PF_serverFavs, local = favGet(), removed = favRemoved(), seen = favServerSeen();
    var serverSet = {}; server.forEach(function (m) { serverSet[m] = 1; });
    var seenSet = {}; seen.forEach(function (m) { seenSet[m] = 1; });
    // A local pin the server USED to have but no longer lists was un-pinned on another device — drop it here
    // and don't re-push it (that resurrection was the bug). A local pin the server NEVER had is a genuine
    // new/offline pin → keep it and push it up below.
    local = local.filter(function (m) { return !(seenSet[m] && !serverSet[m]); });
    var merged = window.PF.mergeFavorites(local, server, removed).slice(0, 24); // tombstoned removals stay dropped
    local.forEach(function (m) { if (!serverSet[m] && !seenSet[m] && removed.indexOf(m) === -1) favSync("add", m); }); // push genuine local-only pins up
    // Re-send the delete for any tombstoned model the server STILL lists (the original write didn't land);
    // once the server no longer reports it the removal has propagated — drop the tombstone so it can't
    // suppress a genuine re-pin from another device forever.
    var stillRemoved = removed.filter(function (m) { if (serverSet[m]) { favSync("remove", m); return true; } return false; });
    favSetRemoved(stillRemoved);
    favSetServerSeen(server); // remember what the server had THIS load, to diff against next load
    try { localStorage.setItem("pf-favs", JSON.stringify(merged)); } catch (e) {}
    if (view === "catalog") refreshFavoritesStrip(); // reflect newly-synced pins if we're on Find
  }
  // "Your Favorites" strip on the Find home — pinned models as one-click chips (skips any no longer in the catalog).
  function favoritesStrip() {
    var favs = favGet().filter(function (m) { return !!findModelPath(m); });
    if (!favs.length) return null;
    var box = document.createElement("div"); box.className = "favs view"; box.setAttribute("data-anim", "favs");
    box.innerHTML = '<div class="favs-head"><span class="fav-star" aria-hidden="true">★</span>Your Favorites</div><div class="favs-row">' +
      favs.map(function (m) { return '<button class="fav-chip" type="button" data-m="' + escapeAttr(m) + '">' + escapeHtml(m) + '</button>'; }).join("") + '</div>';
    [].forEach.call(box.querySelectorAll(".fav-chip"), function (b) {
      b.addEventListener("click", function () { var p = findModelPath(b.getAttribute("data-m")); if (p) openTo(p, true); }); // true = restore keyboard focus after re-render
    });
    return box;
  }
  // Update the "Your Favorites" strip IN PLACE (no full re-render) after a favorite is added/removed from
  // a resource panel, so a newly-favorited model appears at the top immediately. Replaces the existing strip,
  // inserts a fresh one (just above Recently viewed) if none was showing, or removes it once the last fav is gone.
  function refreshFavoritesStrip() {
    var old = content.querySelector(".favs.view:not(.recent)");
    var fresh = favoritesStrip();
    if (old) { if (fresh) old.parentNode.replaceChild(fresh, old); else old.remove(); return; }
    if (!fresh) return;
    var head = content.querySelector(".level-head-row");
    if (head) head.insertAdjacentElement("afterend", fresh);
    else content.insertBefore(fresh, content.firstChild);
  }
  // Recently viewed: automatic per-browser history of opened models (localStorage, like favorites — no
  // backend, no PII). Most-recent-first, deduped, capped. A "view" is recorded whenever a model's resource
  // panel is shown (a click, a pasted deep link, back/forward — all funnel through buildResources).
  function recentGet() { try { var a = JSON.parse(localStorage.getItem("pf-recent") || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
  function recentPush(m) { if (!m) return; var a = recentGet().filter(function (x) { return x !== m; }); a.unshift(m); try { localStorage.setItem("pf-recent", JSON.stringify(a.slice(0, 12))); } catch (e) {} }
  function recentRemove(m) { var a = recentGet().filter(function (x) { return x !== m; }); try { localStorage.setItem("pf-recent", JSON.stringify(a)); } catch (e) {} }
  // "Recently viewed" strip — automatic history chips, shown on every Find view. Excludes favorited models
  // (they show under Your Favorites) and any no longer in the catalog. Each chip has a hover × to remove it.
  function recentlyViewedStrip() {
    var pinned = {}; favGet().forEach(function (m) { pinned[m] = 1; });
    var rec = recentGet().filter(function (m) { return !pinned[m] && !!findModelPath(m); }).slice(0, 8);
    if (!rec.length) return null;
    var box = document.createElement("div"); box.className = "favs recent view"; box.setAttribute("data-anim", "recent");
    box.innerHTML = '<div class="favs-head">' + svg(ICONS.clock) + 'Recently viewed</div><div class="favs-row">' +
      rec.map(function (m) {
        var a = escapeAttr(m), h = escapeHtml(m);
        return '<span class="rec-wrap">' +
          '<button class="fav-chip" type="button" data-m="' + a + '">' + h + '</button>' +
          '<button class="rec-x" type="button" data-m="' + a + '" title="Remove from recently viewed" aria-label="Remove ' + a + ' from recently viewed">×</button>' +
        '</span>';
      }).join("") + '</div>';
    [].forEach.call(box.querySelectorAll(".fav-chip"), function (b) {
      b.addEventListener("click", function () { var p = findModelPath(b.getAttribute("data-m")); if (p) openTo(p, true); }); // true = restore keyboard focus after re-render
    });
    [].forEach.call(box.querySelectorAll(".rec-x"), function (x) {
      x.addEventListener("click", function (e) {
        e.stopPropagation(); recentRemove(x.getAttribute("data-m"));
        var wrap = x.closest(".rec-wrap"); if (wrap) wrap.remove();
        if (!box.querySelector(".rec-wrap")) box.remove(); // strip empties out → drop it
      });
    });
    return box;
  }
  // "New since your last visit" baseline: the timestamp of the viewer's PREVIOUS visit, used to badge
  // What's New cards added since then. Computed ONCE per page load and cached, so the badge is stable across
  // in-session navigation (every renderColumns calls availableStrip). A BRAND-NEW visitor has no baseline →
  // returns now() so NOTHING is badged (they can never get a wall of "new" items); we record now() so the
  // NEXT visit has a baseline to compare against.
  var _visitBaseline = null;
  function visitBaseline() {
    if (_visitBaseline !== null) return _visitBaseline;
    var raw = null; try { raw = localStorage.getItem("pf-lastvisit"); } catch (e) {}
    var now = Date.now(), prev = raw ? parseInt(raw, 10) : NaN;
    _visitBaseline = isNaN(prev) ? now : prev;         // first visit → now (badge nothing); else the prior visit
    try { localStorage.setItem("pf-lastvisit", String(now)); } catch (e) {} // advance once per load for next time
    return _visitBaseline;
  }
  function STARICON(filled) {
    return '<svg viewBox="0 0 24 24" fill="' + (filled ? "currentColor" : "none") + '" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
      + '<path d="m12 3.6 2.6 5.3 5.8.85-4.2 4.1 1 5.75L12 16.9l-5.2 2.7 1-5.75-4.2-4.1 5.8-.85z"/></svg>';
  }

  /* ============================================================
     RESOURCE PANEL, HOME STRIPS & MAIN RENDER
     ============================================================ */
  // Model-panel select mode: pick resource ROWS, then Copy links / Share to customer take only those. A row
  // tick means the whole resource (a Renders folder = all its images); to pick individual files, open the
  // resource and use the gallery's own select mode. Keys = resource names. Reset on every panel build.
  var modelSel = window.PF.makeSelection();
  function buildResources(m, headEl) {
    modelSel.stop();
    recentPush(m); // record this model as recently viewed (click, deep link, or back/forward all land here)
    var box = document.createElement("div"); box.className = "res-panel view"; box.setAttribute("data-anim", "respanel:" + sel.model);
    var pin = document.createElement("button"); pin.type = "button";
    // The star is decoration next to the label — left in the accessible name it reads as "black star
    // Favorited". aria-hidden keeps the glyph visual-only; aria-pressed already carries the on/off state.
    function paintPin() {
      var on = favHas(m);
      pin.className = "res-pin" + (on ? " on" : "");
      pin.innerHTML = STARICON(on);
      pin.setAttribute("aria-pressed", on ? "true" : "false");
      pin.setAttribute("aria-label", on ? "Favorited — remove " + m + " from favorites" : "Add " + m + " to favorites");
      pin.title = on ? "Remove from favorites" : "Add to favorites";
    }
    paintPin();
    pin.addEventListener("click", function () {
      favToggle(m); paintPin();
      refreshFavoritesStrip(); // reflect the add/remove in the top "Your Favorites" strip right away (no page refresh)
      // Confirm the action. "Your Favorites" now shows at the top of every Find view, so no need to direct home.
      if (favHas(m)) showToast('<span><b>★ Added to favorites.</b></span>');
      else showToast('<span>Removed from favorites.</span>');
    });
    // Pin sits inline in the column header (same row as the model name); fall back to an in-panel row if no header given.
    // A "Copy links" share-pack button sits just left of the pin (skipped when the model has no openable resources).
    var mpack = modelSharePack(m);
    // Read at click time: in select mode the pack is only the picked resources ("Copy selected links (N)").
    function packNow() {
      var p = modelSharePack(m);
      if (modelSel.isOn()) { p.resources = p.resources.filter(function (r) { return modelSel.has(r.name); }); p.count = p.resources.length; p.empty = !p.count; }
      return p;
    }
    var shareModelBtn = mpack.empty ? null : sharePackBtn("share-model", "Copy links", "Copy every in-app link for " + m, packNow);
    var selBtn = null;
    if (!mpack.empty) {
      selBtn = document.createElement("button"); selBtn.type = "button"; selBtn.className = "share-pack res-select";
      selBtn.setAttribute("aria-pressed", "false"); selBtn.title = "Pick which files Share and Copy links take";
      selBtn.innerHTML = TICKICON + '<span>Select</span>';
      selBtn.addEventListener("click", function () { if (modelSel.isOn()) modelSel.stop(); else modelSel.start(); paintModelSel(); });
    }
    // "Share to customer" for the WHOLE model — sends the actual files (docs first, capped) to the phone's
    // share sheet. Only where the device can share files (mobile/tablet); desktop has no file-share API, so
    // (like the per-file Share button) it isn't shown there — reps download+email on desktop instead.
    var shareCustBtn = (!mpack.empty && navigator.canShare && navigator.share) ? modelShareCustomerBtn(m) : null;
    if (headEl) {
      headEl.classList.add("col-head-res");
      var acts = document.createElement("span"); acts.className = "col-head-actions";
      if (selBtn) acts.appendChild(selBtn);
      if (shareCustBtn) acts.appendChild(shareCustBtn);
      if (shareModelBtn) acts.appendChild(shareModelBtn);
      acts.appendChild(pin);
      headEl.appendChild(acts);
      // Move the heading semantics off the CONTAINER and onto the model name. The buttons live inside
      // this row for layout reasons, and with role="heading" on the row the whole thing announced as
      // "Hoardmaster Share to customer Copy links ☆ Add to favorites, heading level 2" — the heading text
      // stopped being the model name. The name span carries the role instead; the row is just a row.
      var leadName = headEl.querySelector(".lead-name");
      if (leadName) {
        headEl.removeAttribute("role"); headEl.removeAttribute("aria-level");
        leadName.setAttribute("role", "heading"); leadName.setAttribute("aria-level", "2");
      }
    } else {
      var pinRow = document.createElement("div"); pinRow.className = "res-pinrow";
      if (selBtn) pinRow.appendChild(selBtn);
      if (shareCustBtn) pinRow.appendChild(shareCustBtn);
      if (shareModelBtn) pinRow.appendChild(shareModelBtn);
      pinRow.appendChild(pin); box.appendChild(pinRow);
    }
    var linked = usableCount(m), total = applicableCount(m), complete = total > 0 && linked === total;
    var cov = document.createElement("div"); cov.className = "coverage";
    var segs = RESOURCE_TYPES.filter(function (r) { return !naFor(m, r.name); }).map(function (r, i) {
      return '<span class="cseg' + (usableLinked(m, r.name) ? ' on' : '') + '" data-cseg="' + i + '" data-anim="cseg:' + escapeAttr(sel.model) + ':' + i + '"></span>';
    }).join("");
    var label = complete ? '<b>Complete</b> · all ' + total + ' linked'
      : linked === 0 ? 'Work in progress · <b>0</b> of ' + total + ' linked'
      : '<b>' + linked + '</b> of ' + total + ' linked';
    cov.innerHTML = '<div class="cov-segs" aria-hidden="true">' + segs + '</div><div class="coverage-num' + (complete ? ' complete' : '') + '">' + label + '</div>';
    box.appendChild(cov);
    var selbar = document.createElement("div"); selbar.className = "res-selbar"; selbar.hidden = true;
    selbar.innerHTML = '<button type="button" class="wire-btn res-sel-all">Select all</button><button type="button" class="wire-btn res-sel-clear" hidden>Clear</button><span class="res-sel-count" aria-live="polite"></span>';
    box.appendChild(selbar);
    var selKeys = []; // resource names a tick can pick, in row order (healthy, linked rows only)
    function paintModelSel() {
      var on = modelSel.isOn(), n = modelSel.count();
      box.classList.toggle("selecting", on);
      if (selBtn) { selBtn.innerHTML = TICKICON + '<span>' + (on ? "Done" : "Select") + '</span>'; selBtn.setAttribute("aria-pressed", on ? "true" : "false"); }
      selbar.hidden = !on;
      selbar.querySelector(".res-sel-clear").hidden = !n;
      selbar.querySelector(".res-sel-count").textContent = on ? (n + " of " + selKeys.length + " selected") : "";
      [].forEach.call(list.querySelectorAll(".res[data-res]"), function (row) {
        var k = row.getAttribute("data-res"), p = on && modelSel.has(k);
        row.classList.toggle("picked", p);
        // In select mode a row IS a checkbox (a[href] may carry that role); off, it is the link it was.
        if (on) { row.setAttribute("role", "checkbox"); row.setAttribute("aria-checked", p ? "true" : "false"); }
        else { row.removeAttribute("role"); row.removeAttribute("aria-checked"); }
      });
      if (shareModelBtn) { shareModelBtn.innerHTML = PACKICON + '<span>' + escapeHtml(window.PF.selectionLabel("Copy links", "Copy selected links", modelSel)) + '</span>'; shareModelBtn.disabled = on && !n; }
      if (shareCustBtn) setModelShareBusy(shareCustBtn, false);
    }
    selbar.querySelector(".res-sel-all").addEventListener("click", function () { modelSel.all(selKeys); paintModelSel(); });
    selbar.querySelector(".res-sel-clear").addEventListener("click", function () { modelSel.clear(); paintModelSel(); });
    var list = document.createElement("div"); list.className = "res-list";
    RESOURCE_TYPES.forEach(function (r) {
      if (naFor(m, r.name)) return; // resource type not applicable to this model — don't show it
      var arr = linksFor(m, r.name);
      var hs = arr.length ? healthStatusFor(m, r.name) : null; // 'broken' | 'empty' | null (last Link-health run)
      var icon = '<span class="res-icon">' + svg(ICONS[r.icon]) + '</span>';
      var body = '<span class="res-body"><span class="res-name">' + r.name + '</span><span class="res-desc">' + r.desc + '</span>'; // "Updated …" now shows in the preview header
      var wipBody = '<span class="res-body"><span class="res-name">' + r.name + '</span><span class="res-desc res-wipnote"><span class="pdot"></span>Not in the Hub yet</span>';
      if (arr.length === 0 || hs === "empty") {
        // No usable files yet — either never linked, OR a link that resolves to nothing for this model.
        // Instead of a dead-end toast, offer an explicit "Request this" that opens the intake form prefilled
        // with THIS model + resource. The row is a container (not a button); the focusable Request control is
        // the action, so a keyboard user reaches it directly. Clicking it is DEMAND for the exact slot the
        // user navigated to (stronger than a search miss — no guessing at intent), aggregated on Metrics.
        var el = document.createElement("div"); el.className = "res pending";
        // "Work in progress" reads as "someone is actively on it" — a promise the Hub can't keep. The
        // honest statement is that the file isn't here yet. And a request that opens a tab while nothing
        // changes in-app looks ignored, so remember the click and acknowledge it.
        var reqKey = "pf-req:" + m + "::" + r.name, alreadyReq = false;
        try { alreadyReq = !!localStorage.getItem(reqKey); } catch (e) {}
        el.title = r.desc + " — not linked yet. Request this opens a pre-filled request form.";
        el.innerHTML = icon + wipBody + '</span><span class="res-wip">' +
          '<button type="button" class="res-req" aria-label="Request ' + escapeAttr(r.name + " for " + m) + '">' + REQICON + '<span class="res-req-label">' + (alreadyReq ? 'Requested ✓' : 'Request this') + '</span></button></span>';
        el.querySelector(".res-req").addEventListener("click", function () {
          if (window.PF_logEvent) window.PF_logEvent("wip_click", m + " · " + r.name);
          window.open(resourceRequestFormUrl(m, r.name), "_blank", "noopener");
          try { localStorage.setItem(reqKey, String(Date.now())); } catch (e) {}
          var lbl = this.querySelector(".res-req-label"); if (lbl) lbl.textContent = "Requested ✓";
          showToast(svg(ICONS.ext) + '<span>Request form opened in a new tab, with the model and resource filled in. The product team replies by email.</span>', 7000);
        });
        list.appendChild(el);
      } else if (hs === "broken") {
        var el = document.createElement("a"); el.className = "res res-broken"; el.href = arr[0].url; el.target = "_blank"; el.rel = "noopener";
        // Audience-correct tooltip: PMs get the actual fix location (the link check lives in the To Do box on
        // Progress); everyone else (who can't see the PM tabs) is told it's already flagged + how to report it.
        var brokenTip = window.PF_isPM
          ? "This didn’t open when links were last checked" + healthWhen() + " — the file may be moved, deleted, or restricted. Re-run the link check from the To Do box on the Progress tab."
          : "This didn’t open when links were last checked" + healthWhen() + " — the file may be moved, deleted, or you may not have access. The product team has been notified; you can also report it with the “Report a bug” button (bottom-right)."; // plain text: a title attribute renders markup literally
        el.innerHTML = icon + body + '</span><span class="pill bad" title="' + escapeAttr(brokenTip) + '">' + svg(ICONS.alert) + ' Link issue</span>';
        var title = m + " · " + r.name;
        el.addEventListener("click", function (e) { resourceClick(e, arr, title, m); });
        list.appendChild(el);
      } else {
        var el = document.createElement("a"); el.className = "res"; el.href = arr[0].url; el.target = "_blank"; el.rel = "noopener";
        el.setAttribute("data-res", r.name); selKeys.push(r.name);
        // "Open ↗" promises a trip to SharePoint; this opens the in-app viewer. Reserve the ↗ glyph for
        // links that really do leave the Hub (the lightbox's own "Open in SharePoint").
        el.innerHTML = '<span class="res-tick" aria-hidden="true">' + TICKICON + '</span>' + icon + body + '</span><span class="pill ok">Preview</span>';
        var title = m + " · " + r.name;
        el.addEventListener("click", function (e) {
          if (modelSel.isOn()) { e.preventDefault(); modelSel.toggle(r.name); paintModelSel(); return; } // a pick, not a preview
          resourceClick(e, arr, title, m);
        });
        el.addEventListener("keydown", function (e) { if (modelSel.isOn() && e.key === " ") { e.preventDefault(); el.click(); } }); // checkbox semantics: Space toggles
        list.appendChild(el);
      }
    });
    box.appendChild(list);
    return box;
  }

  // ---- Recently added strip (items appear here for RECENT_DAYS after their added date) ----
  function availableStrip() {
    var now = Date.now(), windowMs = RECENT_DAYS * 864e5, items = [];
    Object.keys(LINKS).forEach(function (k) {
      var addedStr = ADDED[k]; if (!addedStr) return;
      var t = dayStart(addedStr);
      if (isNaN(t) || (now - t) > windowMs) return;
      var arr = normLinks(LINKS[k]); if (!arr.length) return;
      var parts = k.split("::");
      var s = healthStatusFor(parts[0], parts[1]); if (s === "empty" || s === "broken") return; // don't advertise resources with no usable files
      items.push({ k: k, m: parts[0], rname: parts[1], arr: arr, t: t });
    });
    if (!items.length) return null;
    items.sort(function (a, b) { return b.t - a.t; });
    var totalRecent = items.length;
    if (items.length > RECENT_MAX) items = items.slice(0, RECENT_MAX); // show only the most-recent few
    var lastVisit = visitBaseline(); // badge cards added since the viewer's previous visit (none on a first visit)
    var cards = items.map(function (it) {
      var sub = it.rname + (it.arr.length > 1 ? " · " + it.arr.length + " files" : "");
      // Added dates are DATE-only, so it.t is that day's LOCAL midnight. Compare the END of the added day
      // to lastVisit, so a resource added TODAY is still badged "New" for a viewer who also visited earlier
      // today (a midnight it.t would sort before that visit and the badge would be missed).
      var isNew = (it.t + 86399999) > lastVisit ? '<span class="now-new">New</span>' : '';
      // Model codes alone don't say WHERE a thing lives ("HMX-400" of which product?); prefix the product.
      var mp = findModelPath(it.m), trail = (mp && mp[2] && mp[2] !== it.m) ? '<span class="now-trail">' + escapeHtml(mp[2]) + ' \u00b7 </span>' : '';
      return '<a class="now-card" href="' + escapeAttr(it.arr[0].url) + '" target="_blank" rel="noopener" data-key="' + escapeAttr(it.k) + '" data-m="' + escapeAttr(it.m) + '" data-r="' + escapeAttr(it.rname) + '">' +
        '<span class="now-ic">' + svg(ICONS[iconFor(it.rname)]) + '</span>' +
        '<span class="now-body"><span class="now-model">' + escapeHtml(it.m) + isNew + '</span><span class="now-res">' + trail + escapeHtml(sub) + '</span></span>' +
        '<span class="now-open">' + svg(ICONS.ext) + '</span></a>';
    });
    var strip = document.createElement("div"); strip.className = "now view"; strip.setAttribute("data-anim", "now");
    // Report the TRUE recent count, not the capped display count — "6 of 20" when truncated, so a busy
    // wiring week doesn't look like only 6 things changed.
    var countLabel = totalRecent > items.length ? (items.length + ' of ' + totalRecent + ' resources')
      : (totalRecent + ' resource' + (totalRecent === 1 ? '' : 's'));
    strip.innerHTML = '<div class="now-head">What\'s New · ' + countLabel + '</div><div class="now-grid">' + cards.join("") + '</div>';
    strip.querySelectorAll(".now-card").forEach(function (a) {
      a.addEventListener("click", function (e) {
        // Let the browser handle every modified click. shift-click (open in a new window) was being
        // hijacked into the in-app viewer; e.button is always 0 in a "click" handler, so the middle-click
        // clause it used to carry never did anything — a middle click fires "auxclick", not "click".
        if (e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) return;
        e.preventDefault();
        var arr = normLinks(LINKS[a.getAttribute("data-key")]);
        var title = a.getAttribute("data-m") + " · " + a.getAttribute("data-r");
        openResource(arr, title, a.getAttribute("data-m"));
      });
    });
    return strip;
  }

  // Shared "catalog isn't available" panel \u2014 used by every view so an outage never reads as "empty".
  function catalogUnavailable(eyebrow, title) {
    // A 403 is not an outage: the person is signed in to Microsoft but this account isn't allowed here (a
    // B2B guest, or an address outside the allowed domains). Telling them to "refresh" produced a loop
    // with no way out, so say what is actually true and offer the one action that can help.
    if (catalogDenied) {
      return head(eyebrow, title, "") +
        '<div class="view view-note note-strong" role="alert">This account isn\u2019t authorized for Product Hub \u2014 it\u2019s for Grumbleton staff sign-ins. ' +
        'If you have a separate work account, sign out and use that one; otherwise ask the product team for access.</div>';
    }
    return head(eyebrow, title, "") +
      '<div class="view view-note note-strong" role="alert">The product catalog couldn\u2019t be loaded just now. ' +
      '<button class="wire-btn" id="catRetry" type="button">Try again</button> If it keeps happening, sign out and sign back in.</div>';
  }
  // A loading state so the first paint (before the Smartsheet fetch resolves) isn't a blank void.
  function renderLoading() {
    // Match the loading header to the deep-link target so a #progress/#admin cold-load doesn't flash
    // "Products & resources" before switching.
    var h = location.hash || "", eyebrow = "Fine Goods for Dragons", title = "Products & resources";
    if (h === "#progress") { eyebrow = "Product team"; title = "Progress"; }
    else if (h === "#metrics") { eyebrow = "Product team"; title = "Metrics"; }
    else if (h === "#styleguide") { eyebrow = "Product team"; title = "Style guide"; }
    else if (h === "#admin") { eyebrow = "Product team"; title = "Admin"; }
    // Skeleton the Find layout (favorites chips \u2192 What's-New cards \u2192 the family tree) so the first paint
    // reads as "content arriving", not a blank wait. Non-Find deep-loads keep a simple status line.
    function skRep(html, n) { var s = ""; for (var i = 0; i < n; i++) s += html; return s; }
    var body;
    if (title === "Products & resources") {
      body = '<div class="view" role="status" aria-live="polite" aria-label="Loading the catalog">' +
        '<span class="sr-only">Loading the catalog\u2026</span>' +
        '<div class="sk-row" aria-hidden="true">' + skRep('<span class="sk sk-chip"></span>', 3) + '</div>' +
        '<div class="sk-row" aria-hidden="true">' + skRep('<span class="sk sk-card"></span>', 4) + '</div>' +
        '<div class="sk-row sk-row-col" aria-hidden="true">' + skRep('<span class="sk sk-line"></span>', 5) + '</div>' +
        '</div>';
    } else {
      body = '<div class="view view-note" role="status" aria-live="polite">Loading\u2026</div>';
    }
    content.innerHTML = head(eyebrow, title, "") + body;
  }

  // ---- main render (columns) ----
  // Remembers which animatable elements ([data-anim]) were on screen last render, so on the next
  // render we can suppress the entrance animation for ones that persist — only new things animate.
  var lastAnimKeys = null;
  function scopeEntranceAnim(root) {
    var cur = {};
    [].forEach.call(root.querySelectorAll("[data-anim]"), function (el) {
      var idn = el.getAttribute("data-anim"); cur[idn] = 1;
      if (lastAnimKeys && lastAnimKeys[idn]) el.classList.add("no-anim");
    });
    lastAnimKeys = cur;
  }
  function renderColumns() {
    if (catalogError) { content.innerHTML = catalogUnavailable("Fine Goods for Dragons", "Products & resources"); lastAnimKeys = null; syncViewTabs(); return; }
    if (!CATALOG.length) { // loaded fine, but the sheet has no products yet — show an empty state, not a bare tree
      content.innerHTML = head("Fine Goods for Dragons", "Products & resources", "") +
        '<div class="view view-note">No products are in the catalog yet. Once they’re added to the Product Family Tree, they’ll appear here.</div>';
      lastAnimKeys = null; syncViewTabs(); return;
    }
    content.innerHTML =
      '<div class="level-head level-head-row">' +
        '<div><div class="eyebrow">Fine Goods for Dragons</div><h1>Products & resources</h1></div>' +
        '<div class="head-actions"><button class="head-btn" id="reqBtn" type="button">' + REQICON + '<span>' + (isPM() ? 'Global Project Intake Form' : 'Request a product or file') + '</span></button></div>' +
      '</div>';
    var reqBtn = document.getElementById("reqBtn");
    if (reqBtn) reqBtn.addEventListener("click", openReqForm);
    // "Your Favorites" shows on EVERY Find view (home or drilled into a product/model), so a pinned model
    // is always one click away — previously it only rendered on the cleared home view, so pins looked lost.
    var favs = favoritesStrip(); if (favs) content.appendChild(favs);
    var recent = recentlyViewedStrip(); if (recent) content.appendChild(recent); // shown on every Find view
    // "What's New" is a home-view welcome strip; once you've drilled into a product, hide it so you're not
    // scrolling past it every time to reach the tree. (Other Collateral is a real category — it stays.)
    if (!sel.prod) { var strip = availableStrip(); if (strip) content.appendChild(strip); }

    // "Other Collateral" (channel-based) sits full-width above the breadcrumb, collapsed until clicked.
    var oc = otherCollateralPanel();
    if (oc) content.appendChild(oc);

    // Once you've drilled in, a breadcrumb replaces the pick-a-family hint — orientation when deep
    // (especially on mobile, where the columns stack and the path scrolls off the top).
    var crumb = breadcrumbBar();
    if (crumb) content.appendChild(crumb);
    else {
      var hint = document.createElement("p"); hint.className = "cols-hint";
      hint.textContent = "Pick a product line, then a product, then a model to see its files.";
      content.appendChild(hint);
    }

    // With nothing picked yet, .cols holds ONLY the category tree, which is ~378px of a ~1500px row —
    // the rest of the width sat empty for the whole scroll. The class lets the categories flow into
    // columns until a product is chosen, at which point the three-column drill-down takes over.
    var cols = document.createElement("div"); cols.className = "cols" + (sel.prod ? "" : " cols-root");
    var treeEl = familyTreePanel();
    cols.appendChild(treeEl);
    var mc = sel.prod ? modelColumn() : null; if (mc) cols.appendChild(mc);
    var rc = sel.model ? resourceColumn() : null; if (rc) cols.appendChild(rc);
    content.appendChild(cols);

    // MOBILE: the columns can't sit side-by-side. Place the opened Models/Resources as full-width blocks
    // UNDERNEATH the WHOLE category the product belongs to (right after its .cat-box — so above the next
    // category, separate from the last family in this one), NOT inside the family's product accordion. The
    // family→product accordion inside the category is left untouched. Desktop keeps the side-by-side columns.
    var stackedMobile = !!(window.matchMedia && window.matchMedia("(max-width: 860px)").matches);
    var anchorCol = null;
    if (stackedMobile) {
      var catBox = null;
      if (sel.prod) { var prow = treeEl.querySelector('[data-key="' + cssEsc(sel.prod) + '"]'); catBox = prow ? prow.closest(".cat-box") : null; }
      var after = catBox; // insertion point: just after the category box (falls back to leaving them in .cols)
      if (mc && after) { after.insertAdjacentElement("afterend", mc); after = mc; }
      if (rc && after) { after.insertAdjacentElement("afterend", rc); after = rc; }
      anchorCol = rc || mc; // scroll target = the opened block
    }
    scopeEntranceAnim(content); // suppress re-animation of elements that were already on screen

    if (flashKey) {
      var fe = cols.querySelector('[data-key="' + cssEsc(flashKey) + '"]');
      if (fe) { fe.classList.add("flash"); setTimeout(function () { fe.classList.remove("flash"); }, 1000); }
    }
    // Restore keyboard focus to the just-activated node (the innerHTML rebuild dropped it to <body>).
    if (pendingFocusKey) {
      var pf = cols.querySelector('[data-key="' + cssEsc(pendingFocusKey) + '"]');
      if (pf) { try { pf.focus({ preventScroll: true }); } catch (e) {} }
    }
    pendingFocusKey = null;
    // Reveal the just-opened content, but ONLY when the user drilled DEEPER (not on same-depth re-renders,
    // which used to jump the scroll on every click). Depth: 1 = tree, 2 = +models, 3 = +resources.
    var newDepth = sel.model ? 3 : (sel.prod ? 2 : 1);
    var deeper = newDepth > lastColDepth;
    if (stackedMobile && anchorCol && (deeper || mobileRevealPending)) {
      // Mobile: land on the opened block on EVERY product/model TAP (not only when going deeper), so
      // switching products/models also brings the new block into view. Clear the FULL sticky topbar —
      // it wraps to several rows on a phone, so a fixed guess let the block's header slide up under the bar
      // and get cut off; measure the bar's real height so the header lands just below it.
      var tbEl = document.querySelector(".topbar");
      var tbH = tbEl ? Math.ceil(tbEl.getBoundingClientRect().height) : 0;
      anchorCol.style.scrollMarginTop = (tbH + 12) + "px";
      anchorCol.scrollIntoView({ block: "start", behavior: RM.matches ? "auto" : "smooth" });
    } else if (!stackedMobile && deeper && cols.lastChild && cols.lastChild.scrollIntoView) {
      cols.lastChild.scrollIntoView({ inline: "end", block: "nearest", behavior: RM.matches ? "auto" : "smooth" });
    }
    mobileRevealPending = false;
    lastColDepth = newDepth;
    flashKey = null;
    // Name the selection in the tab title HERE, not only in render(). toggleFamily/selectProduct/
    // selectModel all change sel and then call renderColumns() directly — and syncHash() sets
    // _internalHash so the hashchange it causes never reaches applyHash()/render() either. The title
    // therefore never followed a user who browsed by CLICKING; it only updated on a cold load, a deep
    // link, back/forward, a tab switch or a search (WCAG 2.4.2). Putting it at the point the catalog is
    // actually painted covers every caller, including any added later. render() still calls it for the
    // Progress/Metrics/Admin views, which don't come through here — for the catalog it just runs twice,
    // and it's an idempotent string assignment.
    setDocTitle();
  }

  /* ============================================================
     COVERAGE STATS
     ============================================================ */
  // ---- coverage tracker (product-team view) ----
  function coverageStats() {
    var totalModels = 0, linkedSlots = 0, totalSlots = 0, groups = [];
    var perType = {}, perTypeApplicable = {}; RESOURCE_TYPES.forEach(function (r) { perType[r.name] = 0; perTypeApplicable[r.name] = 0; });
    CATALOG.forEach(function (c) {
      c.families.forEach(function (f) {
        f.products.forEach(function (p) {
          if (!p.models.length) return;
          var g = { cat: c.name, fam: f.name, prod: p.name, models: [] };
          p.models.forEach(function (m) {
            totalModels++;
            var linkedTypes = {}, naTypes = {}, cnt = 0, applicable = 0;
            RESOURCE_TYPES.forEach(function (r) {
              if (naFor(m, r.name)) { naTypes[r.name] = true; return; } // not applicable — excluded from the denominator
              applicable++; perTypeApplicable[r.name]++; // per-type denominator honors N/A too (D4)
              var on = usableLinked(m, r.name); // usable-only: a link the last health run found empty/broken doesn't count (matches the Find view)
              linkedTypes[r.name] = on;
              if (on) { cnt++; perType[r.name]++; }
            });
            linkedSlots += cnt; totalSlots += applicable;
            g.models.push({ name: m, path: [c.name, f.name, p.name, m], linkedTypes: linkedTypes, naTypes: naTypes, count: cnt, applicable: applicable });
          });
          groups.push(g);
        });
      });
    });
    return { totalModels: totalModels, linkedSlots: linkedSlots, totalSlots: totalSlots, perType: perType, perTypeApplicable: perTypeApplicable, groups: groups };
  }
  // One per-type coverage row (the summary paints it on Progress and refreshCovStats repaints it in place).
  function covTypeHtml(r, s) {
    var n = s.perType[r.name], den = s.perTypeApplicable[r.name], tp = den ? Math.floor(n / den * 100) : 0; // floor + N/A-aware denominator (D4)
    return '<div class="cov-type-name">' + escapeHtml(r.name) + ' <b>' + n + '/' + den + '</b></div>' + covBar(tp);
  }
  function covBar(pct) { return '<div class="cov-bar"><div class="cov-bar-fill" data-barw="' + pct + '"></div></div>'; }
  // Short column/chip labels. The matrix HEADER now uses the resource icons instead (T3-14), but the
  // coverage chips still need a text form that fits.
  var COV_SHORT = { "Flyer": "Flyer", "Product Overview": "Ovrvw", "Spec Sheet": "Spec", "Cut Sheet": "Cut", "Renders": "Rndr", "In-Store Images": "In-Store", "Assembly Instructions": "Asmbly", "Compliance & Testing": "C&T", "Warranty": "Warr", "Training": "Trng" };
  // Repaint ONLY the Progress summary %/counts + per-type bars in place after an inline link edit —
  // the top summary would otherwise stay stale until a tab switch. Deliberately NOT a full
  // renderCoverage() so any in-progress wire-tool drag work survives (same reasoning as fillMatrix).
  function refreshCovStats() {
    if (view !== "progress") return;
    var s = coverageStats();
    var pct = s.totalSlots ? Math.floor(s.linkedSlots / s.totalSlots * 100) : 0; // floor so 199/200 reads 99%, not a misleading 100%
    var big = document.querySelector(".cov-summary .cov-big");
    if (big) big.innerHTML = pct + '%<small>' + s.linkedSlots + ' of ' + s.totalSlots + ' linked</small>';
    var sub = document.querySelector(".cov-summary .cov-sub");
    if (sub) sub.textContent = s.totalModels + ' models · ' + s.totalSlots + ' applicable resources';
    var barwrap = document.querySelector(".cov-summary .cov-barwrap");
    if (barwrap) barwrap.innerHTML = covBar(pct);
    var typeEls = document.querySelectorAll(".cov-types .cov-type");
    RESOURCE_TYPES.forEach(function (r, i) {
      var el = typeEls[i]; if (!el) return;
      el.innerHTML = covTypeHtml(r, s);
    });
  }

  /* ============================================================
     PM · WIRE TOOL (Add files) & MANAGE-DOCUMENT PANEL
     ============================================================ */
  // ---- Progress tool: wire a SharePoint folder to a product's resource, with a preview ----
  function wirePanel(bare) {
    var box = document.createElement("div"); box.className = "wire";
    box.innerHTML =
      (bare ? '' : '<div class="wire-head" role="heading" aria-level="2">' + svg(ICONS.sync) + '<span>Link a folder</span></div>') +
      '<p class="wire-lead">Link a SharePoint <b>folder</b> to a product’s resource, and new files added to that folder show up on this site automatically.</p>' +
      '<p class="wire-steps">Choose the product&nbsp;→&nbsp;choose the resource type&nbsp;→&nbsp;paste the folder link&nbsp;→&nbsp;Preview, then Approve.</p>' +
      '<div class="wire-form">' +
        '<select class="wire-sel" id="wireProduct" aria-label="Product"></select>' +
        '<select class="wire-sel" id="wireResource" aria-label="Resource type"></select>' +
        '<input class="wire-input" id="wireUrl" type="url" aria-label="SharePoint folder link" placeholder="Paste the SharePoint folder link">' +
        '<button class="wire-btn" id="wirePreview" type="button">Preview</button>' +
      '</div>' +
      '<details class="wire-help">' +
        '<summary>How do I get the folder link?</summary>' +
        '<div class="wire-help-body">' +
          '<ol class="wire-help-steps">' +
            '<li>In SharePoint, open the <b>folder</b> that holds the files — not a single file.</li>' +
            '<li>Click <b>Copy link</b> at the top of the folder <em>(not <b>Share</b>)</em>.</li>' +
            '<li><b>Check who can open it.</b> If the link says <em>“Specific people”</em>, click that and choose <b>“People in Grumbleton USA”</b> — otherwise the site (and your colleagues) can’t open the files and they’ll show up as a <b>Link issue</b>.</li>' +
            '<li>Paste it into the box above and hit <b>Preview</b>.</li>' +
          '</ol>' +
          '<img class="wire-help-img" src="img/copy-link-help.png" loading="lazy" decoding="async" alt="In SharePoint, open the folder and click Copy link at the top">' +
        '</div>' +
      '</details>' +
      '<details class="wire-allfam">' +
        '<summary>' + svg(ICONS.sync) + ' One folder for every family (e.g. spec sheets / trainings)</summary>' +
        '<div class="wire-allfam-body">' +
          '<p class="wire-help-note">Have one folder with a single file per family (named to start with the family name)? Match them all at once — each file goes to <b>every model</b> in its family, across products. You’ll confirm the matches before anything saves.</p>' +
          '<div class="wire-form">' +
            '<select class="wire-sel" id="wireFamResource" aria-label="Resource type"></select>' +
            '<input class="wire-input" id="wireFamUrl" type="url" aria-label="Folder link" placeholder="Paste the folder link (one file per family)">' +
            '<button class="wire-btn" id="wireFamMatch" type="button">Match to families</button>' +
          '</div>' +
          '<div id="wireFamResult"></div>' +
        '</div>' +
      '</details>' +
      '<div class="wire-result" id="wireResult"></div>';
    var helpImg = box.querySelector(".wire-help-img");
    if (helpImg) helpImg.addEventListener("error", function () { this.style.display = "none"; });
    var prodSel = box.querySelector("#wireProduct");
    prodSel.innerHTML = '<option value="">Choose a product…</option>';
    CATALOG.forEach(function (c) {
      c.families.forEach(function (f) {
        f.products.forEach(function (p) {
          var o = document.createElement("option");
          o.value = c.name + "§" + f.name + "§" + p.name;
          o.textContent = f.name + " › " + p.name;
          prodSel.appendChild(o);
        });
      });
    });
    var resSel = box.querySelector("#wireResource");
    // Forced choice: without a placeholder the select silently defaults to the first type ("Spec Sheet"),
    // so a PM wiring renders who doesn't touch this dropdown would save them into the wrong resource slot.
    resSel.innerHTML = '<option value="">Choose a resource type…</option>';
    RESOURCE_TYPES.forEach(function (r) { var o = document.createElement("option"); o.value = r.name; o.textContent = r.name; resSel.appendChild(o); });
    box.querySelector("#wirePreview").addEventListener("click", function () { wirePreview(box); });
    // Enter submits, like #docPlaceUrl already does. Pasting a link and pressing Enter did nothing here.
    box.querySelector("#wireUrl").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); wirePreview(box); } });
    var famResSel = box.querySelector("#wireFamResource");
    famResSel.innerHTML = '<option value="">Choose a resource type…</option>';
    RESOURCE_TYPES.forEach(function (r) { var o = document.createElement("option"); o.value = r.name; o.textContent = r.name; famResSel.appendChild(o); });
    box.querySelector("#wireFamMatch").addEventListener("click", function () { wireFamMatch(box); });
    box.querySelector("#wireFamUrl").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); wireFamMatch(box); } });
    return box;
  }
  // ---- Progress tool: place ONE file across many products/models (opens the Manage Document panel) ----
  function docPlacePanel(bare) {
    var box = document.createElement("div"); box.className = "wire";
    box.innerHTML =
      (bare ? '' : '<div class="wire-head" role="heading" aria-level="2">' + svg('<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/>') + '<span>Place a file across products</span></div>') +
      '<p class="wire-lead">Have one file that belongs to several products? Paste it here to see everywhere it’s used and add it to more models at once — or move, replace, or remove it everywhere.</p>' +
      '<p class="wire-steps">Choose the resource type&nbsp;→&nbsp;paste the SharePoint <b>file</b> link&nbsp;→&nbsp;Manage.</p>' +
      '<div class="wire-form">' +
        '<select class="wire-sel" id="docPlaceResource" aria-label="Resource type"></select>' +
        '<input class="wire-input" id="docPlaceUrl" type="url" aria-label="SharePoint file link" placeholder="Paste the SharePoint file link">' +
        '<button class="wire-btn primary" id="docPlaceGo" type="button">Manage</button>' +
      '</div>';
    var resSel = box.querySelector("#docPlaceResource");
    resSel.innerHTML = '<option value="">Choose a resource type…</option>';
    RESOURCE_TYPES.forEach(function (r) { var o = document.createElement("option"); o.value = r.name; o.textContent = r.name; resSel.appendChild(o); });
    function go() {
      var resource = resSel.value, url = box.querySelector("#docPlaceUrl").value.trim();
      if (!resource) { showToast('<span>Choose a resource type first.</span>'); return; }
      if (!/^https?:\/\//i.test(url)) { showToast('<span>Paste a full https:// SharePoint file link.</span>'); return; }
      if (looksLikeFolder(url)) { showToast('<span>That looks like a folder link — use <b>Link a folder</b> above. Paste a single <b>file</b> link here.</span>', 6000); return; }
      openDocPanel(url, resource, { mode: "new" });
    }
    box.querySelector("#docPlaceGo").addEventListener("click", go);
    box.querySelector("#docPlaceUrl").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); go(); } });
    return box;
  }
  function wirePreview(box) {
    var pvBtn = box.querySelector("#wirePreview");
    if (pvBtn.disabled) return; // a run is in flight — Enter must not start a second one (F036)
    var prodVal = box.querySelector("#wireProduct").value;
    var resource = box.querySelector("#wireResource").value;
    var url = box.querySelector("#wireUrl").value.trim();
    var result = box.querySelector("#wireResult");
    if (!prodVal || !resource || !url) { result.innerHTML = '<div class="wire-msg">Pick a product, choose a resource type, and paste a folder link first.</div>'; return; }
    var parts = prodVal.split("§"), fam = parts[1], prod = parts[2], models = [];
    CATALOG.forEach(function (c) { if (c.name !== parts[0]) return; c.families.forEach(function (f) { if (f.name !== fam) return; f.products.forEach(function (p) { if (p.name === prod) models = p.models; }); }); });
    if (!models.length) { result.innerHTML = '<div class="wire-msg">That product has no models in the sheet yet.</div>'; return; }
    result.innerHTML = '<div class="wire-msg">Opening the folder &amp; checking what’s already linked…</div>';
    if (!window.PF_getToken) { result.innerHTML = '<div class="wire-msg err">Please sign in again.</div>'; return; }
    var pvDone = busyButton(pvBtn, "Opening…"); // disable during the fetch — no double-submit
    window.PF_getToken().then(function (token) {
      var bearer = "Bearer " + token;
      return fetchFolderChildren(url, bearer).then(function (files) {
        // Also resolve each model's already-wired files (to their identities) so we can MERGE + DE-DUPE.
        return Promise.all(models.map(function (m) {
          return wireLoadExisting(m, resource, bearer).then(function (ex) { return { model: m, ex: ex }; })
            // error:true marks "we couldn't read this model's current links" — distinct from "genuinely empty".
            // On Approve we must NEVER clear a model whose existing state we failed to load (was silently erasing links).
            .catch(function () { return { model: m, ex: { files: [], hasFolder: false, folderUrl: "", error: true } }; });
        })).then(function (exList) {
          var existing = {}; exList.forEach(function (e) { existing[e.model] = e.ex; });
          renderWirePreview(box, { fam: fam, prod: prod, resource: resource, url: url, models: models, files: files, existing: existing });
        });
      });
    }).catch(function (err) { result.innerHTML = '<div class="wire-msg err">Couldn’t open that as a folder (' + escapeHtml(String((err && err.message) || err)) + '). Make sure it’s a <b>folder</b> link (not a single file) — use <b>Copy link</b> on the folder in SharePoint. See “How do I get the folder link?” above.</div>'; })
      .then(function () { pvDone(); });
  }
  // ---- "One folder for every family" — spec-sheet / training bulk wire ----
  // Every family (name -> its products+models across the catalog), for matching + fan-out.
  function wireAllFamilies() {
    var out = [], byName = {};
    CATALOG.forEach(function (c) { c.families.forEach(function (f) {
      if (!byName[f.name]) { byName[f.name] = { name: f.name, prods: [] }; out.push(byName[f.name]); }
      f.products.forEach(function (p) { if (p.models && p.models.length) byName[f.name].prods.push({ prod: p.name, models: p.models }); });
    }); });
    return out.filter(function (fam) { return fam.prods.length; });
  }
  // Match a filename to a family: normalized name STARTS WITH the family name (longest wins); else CONTAINS.
  function wireMatchFamily(filename, fams) {
    var nf = normModel(filename), best = null, bestLen = 0;
    fams.forEach(function (fam) { var nfam = normModel(fam.name); if (nfam && nf.indexOf(nfam) === 0 && nfam.length > bestLen) { best = fam.name; bestLen = nfam.length; } });
    if (best) return best;
    // The starts-with pass above is safe at any length. This CONTAINS fallback is not: a 2-3 character
    // family name appears inside all sorts of filenames by coincidence, which would file a render under
    // the wrong family. Require a substantial name before trusting a mid-string hit.
    fams.forEach(function (fam) { var nfam = normModel(fam.name); if (nfam && nfam.length >= 4 && nf.indexOf(nfam) !== -1 && nfam.length > bestLen) { best = fam.name; bestLen = nfam.length; } });
    return best;
  }
  // Open the folder, auto-match each file to a family, and render a confirm-the-matches list.
  function wireFamMatch(box) {
    var mbBtn = box.querySelector("#wireFamMatch");
    if (mbBtn.disabled) return; // a run is in flight (F036)
    var resource = box.querySelector("#wireFamResource").value;
    var url = box.querySelector("#wireFamUrl").value.trim();
    var res = box.querySelector("#wireFamResult");
    if (!resource || !url) { res.innerHTML = '<div class="wire-msg">Choose a resource type and paste a folder link first.</div>'; return; }
    if (!window.PF_getToken) { res.innerHTML = '<div class="wire-msg err">Please sign in again.</div>'; return; }
    res.innerHTML = '<div class="wire-msg">Opening the folder…</div>';
    var mbDone = busyButton(mbBtn, "Opening…"); // disable during the fetch — no double-submit
    window.PF_getToken().then(function (token) { return fetchFolderChildren(url, "Bearer " + token); }).then(function (files) {
      if (!files.length) { res.innerHTML = '<div class="wire-msg">That folder has no files.</div>'; return; }
      var fams = wireAllFamilies();
      var m0 = files.map(function (f) { return wireMatchFamily(f.name, fams); });
      var opts = '<option value="">(skip)</option>' + fams.map(function (f) { return '<option value="' + escapeAttr(f.name) + '">' + escapeHtml(f.name) + '</option>'; }).join("");
      var rows = files.map(function (f, i) {
        return '<div class="wire-fam-row"><span class="wire-fam-file' + (m0[i] ? '' : ' unmatched') + '" title="' + escapeAttr(f.name) + '">' + escapeHtml(f.name) + '</span>' +
          '<select class="wire-sel wire-fam-sel" data-i="' + i + '" aria-label="Family for ' + escapeAttr(f.name) + '">' + opts + '</select></div>';
      }).join("");
      var matched = m0.filter(Boolean).length;
      res.innerHTML = '<div class="wire-summary">Auto-matched <b>' + matched + ' of ' + files.length + '</b> files to a family by name — <b>check and fix any</b>, then save. Each file goes to <b>every model</b> in the chosen family; “(skip)” files are ignored.</div>' +
        '<div class="wire-fam-list">' + rows + '</div>' +
        '<div class="wire-approve"><button class="wire-btn primary" id="wireFamSave" type="button">Link to all families &amp; save</button></div>';
      [].forEach.call(res.querySelectorAll(".wire-fam-sel"), function (s) { var i = +s.getAttribute("data-i"); if (m0[i]) s.value = m0[i]; });
      res.querySelector("#wireFamSave").addEventListener("click", function () {
        var matches = [];
        [].forEach.call(res.querySelectorAll(".wire-fam-sel"), function (s) { var i = +s.getAttribute("data-i"); if (s.value) matches.push({ file: files[i], family: s.value }); });
        if (!matches.length) { showToast('<span>Pick a family for at least one file.</span>'); return; }
        wireFamSave(box, matches, resource, url, this);
      });
    }).catch(function (err) { res.innerHTML = '<div class="wire-msg err">Couldn’t open that as a folder (' + escapeHtml(String((err && err.message) || err)) + '). Make sure it’s a <b>folder</b> link.</div>'; })
      .then(function () { mbDone(); });
  }
  // Fan each matched file out to EVERY model in its family and save — reusing the data-loss-safe
  // wireApprove path (reads each model's current links, MERGES, skips any it couldn't read).
  function wireFamSave(box, matches, resource, url, btn) {
    var fams = wireAllFamilies(), famMap = {}; fams.forEach(function (f) { famMap[f.name] = f; });
    var raw = []; // {fam, prod, model, file}
    matches.forEach(function (m) { var fam = famMap[m.family]; if (!fam) return; fam.prods.forEach(function (pp) { pp.models.forEach(function (model) { raw.push({ fam: m.family, prod: pp.prod, model: model, file: m.file }); }); }); });
    if (!raw.length) { showToast('<span>Nothing matched.</span>'); return; }
    // Confirm the blast radius before the widest write in the app (one file → every model in each family,
    // across products). It's all additions (nothing removed), but the PM should see the scope first.
    var uModels = {}, uProds = {}; raw.forEach(function (t) { uModels[t.fam + "§" + t.prod + "§" + t.model] = 1; uProds[t.fam + "§" + t.prod] = 1; });
    var nM = Object.keys(uModels).length, nP = Object.keys(uProds).length, nF = matches.length;
    pfConfirm({
      title: "Link to every model in these families?",
      body: ["Add " + nF + " file" + (nF === 1 ? "" : "s") + " as the " + resource + " for " + nM + " model" + (nM === 1 ? "" : "s") + " across " + nP + " product group" + (nP === 1 ? "" : "s") + ".",
             "Nothing is removed — this only adds."],
      confirmLabel: "Save"
    }).then(function (ok) {
    if (!ok) return;
    if (!window.PF_getToken) { showError('<span>Please sign in again.</span>'); return; }
    if (btn) { btn.disabled = true; btn.textContent = "Saving…"; }
    // Plan from the sheet as it is NOW, not this tab's copy (which can be hours old): this is the widest write
    // in the app and, being multi-product, sends no compare-and-swap base — so a stale copy silently dropped a
    // link another PM had added since (F186).
    return loadCatalog({ fresh: true }).then(function () { return window.PF_getToken(); }).then(function (token) {
      var bearer = "Bearer " + token;
      var uniq = {}; raw.forEach(function (t) { if (!uniq[t.model]) uniq[t.model] = t; });
      var items = Object.keys(uniq).map(function (k) { return uniq[k]; });
      // Resolve each model's current links (from the catalog just re-read) to file identities, so the merge
      // de-dupes and can't drop anything.
      return lhRunPool(items, function (it) {
        return wireLoadExisting(it.model, resource, bearer).then(function (ex) { return { model: it.model, ex: ex }; },
          function () { return { model: it.model, ex: { files: [], hasFolder: false, folderUrl: "", error: true } }; });
      }, 5).then(function (loaded) {
        var existing = {}; loaded.forEach(function (r) { existing[r.model] = r.ex; });
        var targets = [], assign = { __none: [] }, seen = {};
        raw.forEach(function (t) {
          var key = wireKey(t.fam, t.prod, t.model);
          if (!seen[key]) {
            seen[key] = 1; targets.push({ key: key, fam: t.fam, prod: t.prod, model: t.model, own: false }); assign[key] = [];
            var ex = existing[t.model] || { files: [], hasFolder: false, folderUrl: "" };
            wireFolderChips(ex).forEach(function (c) { assign[key].push(c); });
            (ex.files || []).forEach(function (f) { assign[key].push(wireInst(f, true)); });
          }
          var ident = wireIdent(t.file);
          if (!assign[key].some(function (i) { return wireIdent(i) === ident; })) assign[key].push(wireInst(t.file, false));
        });
        wireApprove({ resource: resource, url: url, targets: targets, assign: assign, existing: existing }, btn || box.querySelector("#wireFamMatch"));
      });
    }).catch(function (err) { if (btn) { btn.disabled = false; btn.textContent = "Link to all families & save"; } showError('<span>Couldn’t prepare that (' + escapeHtml(String((err && err.message) || err)) + ').</span>', 7000); });
    });
  }
  var wireDrag = null, wireIid = 0; // id of the instance being dragged; instance-id counter
  var wireMenu = null; // the one open custom Move dropdown, if any
  var DUPICON = svg('<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/>');
  var TICKICON = svg('<path d="M20 6 9 17l-5-5"/>'); // select-mode ticks (gallery tiles, model rows) and the Select toggles
  var DELICON = svg('<path d="M18 6 6 18M6 6l12 12"/>');
  var MOVEICON = svg('<path d="M5 9 2 12l3 3"/><path d="M9 5l3-3 3 3"/><path d="M15 19l-3 3-3-3"/><path d="M19 9l3 3-3 3"/><path d="M2 12h20M12 2v20"/>'); // 4-way move arrows
  // --- File identity (so the SAME file wired via two different SharePoint links de-dupes) ---
  function wireNormUrl(u) { return window.PF.canonFileUrl(u); } // the shared identity — keeps an Office link's ?sourcedoc, drops "Copy link" junk
  function wireIdent(f) { return (f && f.fileId) ? ("id:" + f.fileId) : ("u:" + wireNormUrl(f && f.url)); }
  // Session memo for resolved driveItems, keyed by canonical URL. Opening "Manage" on a Spec Sheet resolves
  // EVERY distinct Spec Sheet link in the catalog through Graph to find where the file is used — hundreds
  // of round-trips across ~240 models — and it threw the results away on close, so re-opening the same
  // panel paid the whole cost again. A driveItem's id/name/thumbnail is stable for a session; invalidated
  // on a fresh catalog load (loadCatalog({fresh:true})), which is when links can have changed.
  var _wireResolveMemo = Object.create(null);
  function wireResolveMemoClear() { _wireResolveMemo = Object.create(null); }
  // Resolve a wired URL to a driveItem so we learn its stable id (+ thumbnail/name/type).
  function wireResolveItem(url, bearer) {
    var key = wireNormUrl(url);
    if (key && _wireResolveMemo[key]) return _wireResolveMemo[key];
    // fetchTimeout (not bare fetch): a hung Graph resolve must not occupy a "Check all links" pool slot
    // forever — 15s → "timed out", which lhTransient maps to "unknown" (state preserved, not false-broken).
    var p = graphJson("https://graph.microsoft.com/v1.0/shares/" + shareId(url) + "/driveItem?$expand=thumbnails($select=small,medium)", bearer)
      .then(function (item) {
        var ts = (item.thumbnails && item.thumbnails[0]) || {};
        // 48×36 wire tiles / 40px doc chips: `small` (96px) covers them at 2×. See PF.pickThumb.
        return { name: item.name, url: url, thumb: window.PF.pickThumb(ts, 96), isImage: isImageFile(item.name, (item.file && item.file.mimeType) || ""), fileId: item.id || null };
      });
    // Memo the PROMISE, so concurrent callers in the same pool share one request. A rejection is dropped
    // from the memo so a transient Graph failure isn't cached as a permanent one.
    if (key) { _wireResolveMemo[key] = p; p.catch(function () { delete _wireResolveMemo[key]; }); }
    return p;
  }
  // A model's already-wired FILES for a resource (resolved to identities), plus whether it's folder-wired.
  function wireLoadExisting(model, resource, bearer) {
    var arr = linksFor(model, resource) || []; // [{label,url}]
    // EVERY folder link, not just the first: a resource can hold two auto-updating folders, and keeping one
    // silently dropped the other (and every render it surfaced) on the next save (F041/F297).
    var files = [], hasFolder = false, folderUrl = "", folderUrls = [];
    var jobs = arr.map(function (lnk) {
      if (looksLikeFolder(lnk.url)) { hasFolder = true; folderUrls.push(lnk.url); folderUrl = folderUrl || lnk.url; return Promise.resolve(); }
      return wireResolveItem(lnk.url, bearer).then(function (f) { files.push(f); })
        .catch(function () { files.push({ name: (lnk.label || "linked file"), url: lnk.url, thumb: null, isImage: false, fileId: null }); });
    });
    return Promise.all(jobs).then(function () { return { files: files, hasFolder: hasFolder, folderUrl: folderUrl, folderUrls: folderUrls }; });
  }
  // One locked "auto-updating folder" chip per folder link the model holds.
  function wireFolderChips(ex) {
    var list = (ex && ex.folderUrls && ex.folderUrls.length) ? ex.folderUrls : ((ex && ex.hasFolder && ex.folderUrl) ? [ex.folderUrl] : []);
    return list.map(function (u) { return wireInst({ name: "Current folder — auto-updating", url: u, folder: true }, true); });
  }
  // A model's current links as the sheet holds them — every folder, then every file. The compare-and-swap
  // base and the removal count must both cover all of them.
  function wireExistingUrls(ex) {
    var folders = (ex && ex.folderUrls && ex.folderUrls.length) ? ex.folderUrls : ((ex && ex.hasFolder && ex.folderUrl) ? [ex.folderUrl] : []);
    return folders.concat(((ex && ex.files) || []).map(function (f) { return f.url; }));
  }
  // Resolve a pasted SharePoint URL to its canonical driveItem webUrl + stable id, so what we STORE is
  // identity-stable no matter which "Copy link" form was pasted. NEVER blocks a save: any failure (offline
  // file, permissions, timeout) falls back to the raw URL. Folder links are left untouched — canonicalizing
  // a folder to a driveItem webUrl would break its auto-updating semantics.
  function normalizeLink(rawUrl) {
    if (!rawUrl || looksLikeFolder(rawUrl) || !window.PF_getToken) return Promise.resolve({ url: rawUrl, id: null });
    return window.PF_getToken().then(function (token) {
      return graphJson("https://graph.microsoft.com/v1.0/shares/" + shareId(rawUrl) + "/driveItem?$select=id,webUrl", "Bearer " + token)
        .then(function (it) { return { url: it.webUrl || rawUrl, id: it.id || null }; });
    }).catch(function () { return { url: rawUrl, id: null }; });
  }
  // Everywhere a file is used, robust to "Copy link" URL variance. Resolves the start link + every distinct
  // same-type link to a Graph id (bounded to ONE resource type — tens of links — through the ~5-wide pool),
  // then matches by PF.sameFile (id when known, else canonical URL). Unresolvable links fall back to
  // URL-only identity. Folder-wired links are excluded (a document is a file, not a folder). Returns
  // { doc:{url,id,name,thumb,type}, placements:[{cat,fam,prod,model,url}] }.
  function docPlacements(startUrl, resource) {
    if (!window.PF_getToken) return Promise.reject(new Error("Please sign in again."));
    return window.PF_getToken().then(function (token) {
      var bearer = "Bearer " + token;
      return wireResolveItem(startUrl, bearer)
        .catch(function () { return { name: null, url: startUrl, thumb: null, isImage: false, fileId: null }; })
        .then(function (self) {
          var doc = { url: self.url, id: self.fileId, name: self.name, thumb: self.thumb, type: resource };
          var byUrl = {}; // distinct file url -> [models that link it under this resource]
          Object.keys(LINKS).forEach(function (k) {
            var parts = k.split("::");
            if (parts[1] !== resource) return;
            var model = parts[0];
            linksFor(model, resource).forEach(function (l) {
              if (looksLikeFolder(l.url)) return; // a document is a file, not a folder
              (byUrl[l.url] = byUrl[l.url] || []).push(model);
            });
          });
          var urls = Object.keys(byUrl);
          return window.PF.lhRunPool(urls, function (u) {
            return wireResolveItem(u, bearer).then(function (f) { return { url: u, id: f.fileId || null }; })
              .catch(function () { return { url: u, id: null }; });
          }, 5).then(function (resolved) {
            var placements = [], seenModel = {};
            resolved.forEach(function (r) {
              if (!r || !r.url || !window.PF.sameFile({ url: doc.url, id: doc.id }, { url: r.url, id: r.id })) return;
              (byUrl[r.url] || []).forEach(function (model) {
                if (seenModel[model]) return; // one placement per model even if linked twice
                seenModel[model] = 1;
                var p = findModelPath(model);
                if (p) placements.push({ cat: p[0], fam: p[1], prod: p[2], model: p[3], url: r.url });
              });
            });
            return { doc: doc, placements: placements };
          });
        });
    });
  }
  // ---- Manage Document panel: see/where-used + add-to-many + move + remove-everywhere + replace ----
  // A document-centric view over the model::resource catalog. Every mutation is STAGED (nothing writes
  // until Approve), compiled by the pure PF.planDocumentOps, and executed through the same shrink-guarded
  // postWire → /api/wire-links path the section editor uses. No new endpoint.
  var docPanel = null, docState = null;
  function docKey(fam, prod, model) { return fam + "§" + prod + "§" + model; }
  function ensureDocPanel() {
    if (docPanel) return docPanel;
    docPanel = document.createElement("div");
    docPanel.className = "docp"; docPanel.id = "docPanel"; docPanel.hidden = true;
    docPanel.innerHTML =
      '<div class="docp-backdrop"></div>' +
      '<div class="docp-panel" role="dialog" aria-modal="true" aria-labelledby="docpTitle">' +
        '<div class="docp-head">' +
          '<img class="docp-thumb" alt="" hidden />' +
          '<div class="docp-hmeta"><span class="docp-title" id="docpTitle"></span><span class="docp-badge"></span></div>' +
          '<button class="docp-close lb-close" type="button" aria-label="Close">✕</button>' +
        '</div>' +
        '<div class="docp-body"></div>' +
        '<div class="docp-foot"><span class="docp-summary" aria-live="polite"></span>' +
          '<button class="wire-btn docp-cancel" type="button">Cancel</button>' +
          '<button class="wire-btn primary docp-approve" type="button">Approve &amp; save</button></div>' +
      '</div>';
    document.body.appendChild(docPanel);
    // Every user-initiated close asks first when changes are staged: a stray backdrop click or Escape threw
    // away twenty ticked models and a pasted replacement with no prompt (F146).
    docPanel.querySelector(".docp-backdrop").addEventListener("click", docRequestClose);
    docPanel.querySelector(".docp-close").addEventListener("click", docRequestClose);
    docPanel.querySelector(".docp-cancel").addEventListener("click", docRequestClose);
    docPanel.querySelector(".docp-approve").addEventListener("click", docApprove);
    // Escape closes the panel too (parity with the preview + bug modals). Added once — ensureDocPanel is
    // idempotent. The preview's own Escape handler no-ops here because the preview is closed while this is open.
    document.addEventListener("keydown", function (e) { if (e.key === "Escape" && docPanel && !docPanel.hidden) docRequestClose(); });
    return docPanel;
  }
  // Close the panel; if it was opened from a file preview, run onClose to restore that preview (so Cancel
  // or Save returns the PM to the file they were looking at, instead of dumping them back to the catalog).
  function docHasPending() {
    var st = docState;
    return !!st && (Object.keys(st.pending.add).length > 0 || Object.keys(st.pending.remove).length > 0 || !!st.replaceWith);
  }
  function docRequestClose() {
    if (!docHasPending()) { closeDocPanel(); return; }
    var st = docState;
    pfConfirm({ title: "Discard these changes?", body: ["You’ve staged changes to where this file is used. Closing now throws them away — nothing has been saved yet."], confirmLabel: "Discard", danger: true })
      .then(function (ok) { if (ok && docState === st) closeDocPanel(); });
  }
  function closeDocPanel() {
    if (!docPanel) return;
    // A confirm still up over the panel (Escape while it was opening, a route change) is this panel's
    // question — answer it No rather than leave it floating over whatever comes next (F007).
    if (pfConfirmEl && !pfConfirmEl.hidden) pfConfirmEl._done(false);
    var onClose = docState && docState.onClose;
    docPanel.hidden = true; docState = null; a11yModal.close(docPanel.querySelector(".docp-panel"));
    if (onClose) { try { onClose(); } catch (e) {} }
  }
  // opts.mode: "manage" (default) | "new" (place a new file — starts with the add-picker in view).
  // opts.onClose: optional callback fired after the panel closes (used to reopen the originating preview).
  function openDocPanel(startUrl, resource, opts) {
    if (!isPM()) return;
    opts = opts || {};
    ensureDocPanel();
    docState = { doc: null, placements: [], resource: resource, mode: opts.mode || "manage", meta: {}, pending: { add: {}, remove: {} }, replaceWith: null, onClose: opts.onClose || null };
    var body = docPanel.querySelector(".docp-body");
    docPanel.querySelector(".docp-title").textContent = "Loading…";
    docPanel.querySelector(".docp-badge").textContent = resource || "";
    // The panel element is cached across opens, so a successful save left its button reading "Saving…"
    // (disabled) for the NEXT open — only the error path reset it. Reset the control state on every open.
    var approveBtn = docPanel.querySelector(".docp-approve");
    if (approveBtn) { approveBtn.disabled = false; approveBtn.textContent = "Approve & save"; }
    var thumb = docPanel.querySelector(".docp-thumb"); thumb.hidden = true;
    body.innerHTML = '<div class="wire-msg">Finding everywhere this file is used…</div>';
    docPanel.querySelector(".docp-summary").textContent = "";
    docPanel.hidden = false;
    a11yModal.open(docPanel.querySelector(".docp-panel"), docPanel.querySelector(".docp-close"));
    // Identity check, not just a null check: open Manage for file A, close it, open it for file B, and A's
    // slow response would arrive to find docState non-null — and paint A's placements and A's URL into B's
    // panel. Approving in that window linked the WRONG file. Compare against the state this open created.
    var st = docState;
    docPlacements(startUrl, resource).then(function (res) {
      if (docState !== st) return; // superseded (or closed) while loading
      docState.doc = res.doc; docState.placements = res.placements;
      var t = docPanel.querySelector(".docp-title");
      t.textContent = res.doc.name || "This file";
      if (res.doc.thumb && safeUrl(res.doc.thumb)) { thumb.src = safeUrl(res.doc.thumb); thumb.hidden = false; }
      paintDoc();
    }).catch(function (err) {
      if (docState !== st) return;
      body.innerHTML = '<div class="wire-msg err">Couldn’t load this file’s placements (' + escapeHtml(String((err && err.message) || err)) + ').</div>';
    });
  }
  function docPaintSummary() {
    if (!docState || !docPanel) return;
    var add = 0, rem = 0, rep = 0, k;
    for (k in docState.pending.add) if (docState.pending.add.hasOwnProperty(k)) add++;
    for (k in docState.pending.remove) if (docState.pending.remove.hasOwnProperty(k)) rem++;
    if (docState.replaceWith) docState.placements.forEach(function (p) { if (!docState.pending.remove[docKey(p.fam, p.prod, p.model)]) rep++; });
    var parts = [];
    if (add) parts.push("+" + add + " added");
    if (rem) parts.push("−" + rem + " removed");
    if (rep) parts.push("~" + rep + " replaced");
    docPanel.querySelector(".docp-summary").textContent = parts.length ? parts.join(" · ") : "No changes staged";
    docPanel.querySelector(".docp-approve").disabled = !(add || rem || rep);
  }
  function paintDoc() {
    if (!docState || !docPanel) return;
    var body = docPanel.querySelector(".docp-body"), st = docState, meta = st.meta = {};
    var placedByModel = {}; st.placements.forEach(function (p) { placedByModel[p.model] = p; });
    // --- Used-in list, grouped by family › product ---
    var groups = {}, order = [];
    st.placements.forEach(function (p) {
      var g = p.fam + " › " + p.prod;
      if (!groups[g]) { groups[g] = []; order.push(g); }
      groups[g].push(p);
      // Carry the placement's ACTUAL stored url — where-used matching is id-based, so a placement can be
      // stored under a different url spelling than doc.url (Office vs Copy-link vs /sites/ webUrl). Remove
      // must target that stored url or planDocumentOps (canonFileUrl-based) won't strip it. See docApprove.
      meta[docKey(p.fam, p.prod, p.model)] = { fam: p.fam, prod: p.prod, model: p.model, url: p.url };
    });
    var used = order.map(function (g) {
      return '<div class="docp-grp"><div class="docp-grp-h">' + escapeHtml(g) + '</div>' +
        groups[g].map(function (p) {
          var k = docKey(p.fam, p.prod, p.model), staged = !!st.pending.remove[k];
          return '<div class="docp-row' + (staged ? ' removing' : '') + '"><span class="docp-m">' + escapeHtml(p.model) + '</span>' +
            // aria-label, not title: the button's CONTENT ("×") wins over title for the accessible name,
            // so every one of these announced as "multiplication sign button" with no idea which model.
            '<button class="docp-x" type="button" data-k="' + escapeAttr(k) + '" aria-label="' + escapeAttr((staged ? 'Keep this file on ' : 'Remove this file from ') + p.model) + '" title="' + (staged ? 'Keep this placement' : 'Remove from this model') + '"><span aria-hidden="true">' + (staged ? '↩ undo' : '×') + '</span></button></div>';
        }).join("") + '</div>';
    }).join("");
    var usedHtml = st.placements.length
      ? '<div class="docp-h">Used in ' + st.placements.length + ' place' + (st.placements.length === 1 ? '' : 's') + '</div>' + used
      : '<div class="docp-h">Not used anywhere yet</div><div class="wire-msg">This file isn’t linked to any model. Add it below.</div>';
    // --- Add to more: catalog checkbox tree ---
    var tree = CATALOG.map(function (c) {
      var fams = c.families.map(function (f) {
        var prods = f.products.map(function (p) {
          var models = (p.models || []).map(function (m) {
            var k = docKey(f.name, p.name, m), isPlaced = !!placedByModel[m], isAdd = !!st.pending.add[k];
            if (!isPlaced) meta[k] = { fam: f.name, prod: p.name, model: m };
            return '<label class="docp-mck' + (isPlaced ? ' placed' : '') + '">' +
              '<input type="checkbox" data-k="' + escapeAttr(k) + '"' + (isPlaced ? ' checked disabled' : (isAdd ? ' checked' : '')) + ' />' +
              '<span>' + escapeHtml(m) + (isPlaced ? ' <em>· here</em>' : '') + '</span></label>';
          }).join("");
          return '<div class="docp-prod"><div class="docp-prod-h">' + escapeHtml(p.name) + '</div><div class="docp-models">' + models + '</div></div>';
        }).join("");
        return '<details class="docp-fam"><summary>' + escapeHtml(f.name) + '</summary>' + prods + '</details>';
      }).join("");
      return '<div class="docp-cat"><div class="docp-cat-h">' + escapeHtml(c.name) + '</div>' + fams + '</div>';
    }).join("");
    var actions = '<div class="docp-actions">' +
      '<button class="wire-btn docp-removeall" type="button"' + (st.placements.length ? '' : ' disabled') + '>Remove everywhere</button>' +
      (st.replaceWith
        ? '<span class="docp-replacing">Replacing with <b>' + escapeHtml(st.replaceWith.name || "new file") + '</b> <button class="wire-btn docp-replace-undo" type="button">undo</button></span>'
        : '<button class="wire-btn docp-replace" type="button"' + (st.placements.length ? '' : ' disabled') + '>Replace with a newer file…</button>') +
      '</div>';
    // This function rebuilds the WHOLE body on every click (each × toggle, each undo), which collapsed
    // every family the PM had opened under "Add to more…" and reset the scroll position — so ticking three
    // models meant re-opening the tree and re-finding their place three times. Remember both, restore both.
    var openFams = {}, prevScroll = body.scrollTop;
    [].forEach.call(body.querySelectorAll("details.docp-fam"), function (d) {
      if (d.open) openFams[((d.querySelector("summary") || {}).textContent) || ""] = 1;
    });

    body.innerHTML = '<div class="docp-sec docp-usedsec">' + usedHtml + '</div>' +
      actions +
      '<div class="docp-sec"><div class="docp-h">Add to more…</div><div class="docp-tree">' + tree + '</div></div>';

    [].forEach.call(body.querySelectorAll("details.docp-fam"), function (d) {
      if (openFams[((d.querySelector("summary") || {}).textContent) || ""]) d.open = true;
    });
    if (prevScroll) body.scrollTop = prevScroll;
    // wire the used-list × toggles
    [].forEach.call(body.querySelectorAll(".docp-x"), function (b) {
      b.addEventListener("click", function () {
        var k = b.getAttribute("data-k");
        if (st.pending.remove[k]) delete st.pending.remove[k]; else st.pending.remove[k] = 1;
        paintDoc();
      });
    });
    // wire the add-tree checkboxes
    [].forEach.call(body.querySelectorAll(".docp-mck input:not([disabled])"), function (cb) {
      cb.addEventListener("change", function () {
        var k = cb.getAttribute("data-k");
        if (cb.checked) st.pending.add[k] = 1; else delete st.pending.add[k];
        docPaintSummary();
      });
    });
    var ra = body.querySelector(".docp-removeall");
    if (ra) ra.addEventListener("click", function () {
      st.placements.forEach(function (p) { st.pending.remove[docKey(p.fam, p.prod, p.model)] = 1; });
      paintDoc();
    });
    var rp = body.querySelector(".docp-replace");
    if (rp) rp.addEventListener("click", docStartReplace);
    var ru = body.querySelector(".docp-replace-undo");
    if (ru) ru.addEventListener("click", function () { st.replaceWith = null; paintDoc(); });
    docPaintSummary();
  }
  // Replace-everywhere: paste a newer SharePoint link → normalize → swap it into every current placement.
  function docStartReplace() {
    if (!docState) return;
    var body = docPanel.querySelector(".docp-body"), holder = body.querySelector(".docp-actions");
    if (!holder) return;
    holder.innerHTML = '<div class="docp-replace-form"><input class="docp-replace-url" type="url" placeholder="Paste the newer SharePoint link" aria-label="Newer SharePoint link" />' +
      '<button class="wire-btn docp-replace-ok" type="button">Use this file</button><button class="wire-btn docp-replace-cancel" type="button">Cancel</button></div>';
    var input = holder.querySelector(".docp-replace-url"); try { input.focus(); } catch (e) {}
    holder.querySelector(".docp-replace-cancel").addEventListener("click", paintDoc);
    holder.querySelector(".docp-replace-ok").addEventListener("click", function () {
      var url = input.value.trim();
      if (!/^https?:\/\//i.test(url)) { input.focus(); return; }
      if (looksLikeFolder(url)) { showError('<span>That’s a folder link — paste the link to the newer file itself.</span>'); input.focus(); return; }
      var okDone = busyButton(holder.querySelector(".docp-replace-ok"), "Checking…");
      var st = docState; // the panel this form belongs to — a bare null check isn't enough (see openDocPanel)
      // Really open the link: this file is about to be written into every placement of the document. The
      // old check (normalizeLink) never rejects, so a mistyped link, one the PM can't open, or a folder all
      // passed as "Replacing with Doc.aspx" (F040).
      (window.PF_getToken ? window.PF_getToken() : Promise.reject(new Error("Please sign in again"))).then(function (token) {
        return graphJson("https://graph.microsoft.com/v1.0/shares/" + shareId(url) + "/driveItem?$select=id,name,webUrl,folder", "Bearer " + token);
      }).then(function (item) {
        if (docState !== st) return; // closed, or reopened for a different file, while resolving
        if (item.folder) { okDone(); showError('<span>That’s a folder link — paste the link to the newer file itself.</span>'); return; }
        var n = { url: item.webUrl || url, id: item.id || null };
        // Refuse the file as its own replacement: it looks like a harmless no-op but rewrites every
        // placement of the document.
        if (window.PF.sameFile({ url: n.url, id: n.id }, { url: st.doc && st.doc.url, id: st.doc && st.doc.id })) {
          okDone();
          showError('<span>That’s the same file this panel is already managing.</span>');
          return;
        }
        docState.replaceWith = { url: n.url, id: n.id, name: item.name || "new file" };
        paintDoc();
      }, function () {
        if (docState !== st) return;
        okDone();
        showError('<span>Couldn’t open that link — check it’s a file you can access, then paste it again.</span>', 7000);
      });
    });
    input.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); holder.querySelector(".docp-replace-ok").click(); } });
  }
  function docApprove() {
    if (!docState || !docState.doc) return;
    // Plan against the sheet as it is NOW: this save sends no compare-and-swap base (see below), so planning
    // from a stale tab copy could drop a link another PM added since (F186). One read, PM-only.
    var st0 = docState, checkDone = busyButton(docPanel.querySelector(".docp-approve"), "Checking…");
    loadCatalog({ fresh: true }).then(function () { checkDone(); if (docState === st0) docApprovePlanned(); });
  }
  function docApprovePlanned() {
    if (!docState || !docState.doc) return;
    var st = docState, doc = st.doc, resource = st.resource, meta = st.meta, intents = [], k;
    for (k in st.pending.add) if (st.pending.add.hasOwnProperty(k) && meta[k]) intents.push({ fam: meta[k].fam, prod: meta[k].prod, model: meta[k].model, addUrl: doc.url });
    for (k in st.pending.remove) if (st.pending.remove.hasOwnProperty(k) && meta[k]) intents.push({ fam: meta[k].fam, prod: meta[k].prod, model: meta[k].model, removeUrl: meta[k].url || doc.url });
    if (st.replaceWith) st.placements.forEach(function (p) {
      var pk = docKey(p.fam, p.prod, p.model);
      if (st.pending.remove[pk]) return; // being removed — don't re-add the new file
      intents.push({ fam: p.fam, prod: p.prod, model: p.model, removeUrl: p.url || doc.url, addUrl: st.replaceWith.url });
    });
    if (!intents.length) return;
    var groups = window.PF.planDocumentOps(intents, resource, function (m, r) { return linksFor(m, r).map(function (l) { return l.url; }); });
    if (!groups.length) { showToast('<span>Nothing to change — every model already has this the way you set it.</span>'); return; }
    // Confirm with real numbers and names, and ALWAYS when anything is removed or replaced — keyed off the
    // intent, since a replace now declares 0 net removals (F038/F146). "Some models will lose a link" told
    // a PM nothing about a remove-everywhere across 40 models.
    var hasRemoval = intents.some(function (it) { return !!it.removeUrl; });
    var rm = [], rp = [], nAdd = 0;
    intents.forEach(function (it) { if (it.removeUrl && it.addUrl) rp.push(it.model); else if (it.removeUrl) rm.push(it.model); else if (it.addUrl) nAdd++; });
    var names = function (list) { return list.slice(0, 8).join(", ") + (list.length > 8 ? " and " + (list.length - 8) + " more" : ""); };
    var plural = function (n) { return n + " model" + (n === 1 ? "" : "s"); };
    var body = [];
    if (rm.length) body.push("Remove this file from " + plural(rm.length) + ": " + names(rm) + ".");
    if (rp.length) body.push("Replace it with “" + ((st.replaceWith && st.replaceWith.name) || "the new file") + "” on " + plural(rp.length) + ": " + names(rp) + ".");
    if (nAdd) body.push("Add it to " + plural(nAdd) + ".");
    body.push("The files stay in SharePoint — only the links change.");
    pfConfirmIf(hasRemoval, { title: "Save these changes?", body: body, confirmLabel: "Save", danger: rm.length > 0 }).then(function (ok) {
    if (!ok) return;
    var approveDone = busyButton(docPanel.querySelector(".docp-approve"), "Saving…");
    // One product at a time, each with its compare-and-swap base — the pre-image planDocumentOps built from
    // the fresh read docApprove just did. Firing them all at once collided on the one sheet and blew through
    // the per-user rate limit on a company-wide file (F039/F092). A model already in its wanted state is not
    // re-sent, so after a partial failure "Approve again" — re-planned from a fresh read — sends only what's
    // left, and its base is current.
    var done = 0;
    groups.reduce(function (chain, g) {
      return chain.then(function () {
        return postWire(g.fam, g.prod, g.resource, g.assignments, { expectRemoved: g.expectRemoved, expectRemovedByModel: g.expectRemovedByModel, base: g.base }).then(function () {
          Object.keys(g.finalByKey).forEach(function (key) { setLinks(key, g.finalByKey[key]); clearHealthFor(key); });
          done++;
        });
      });
    }, Promise.resolve()).then(function () {
      toastSaved('Document updated.');
      refreshCovStats();
      repaintProgressSafely(); // not a full re-render: the section editor this panel was opened from may hold staged edits (F027)
      // Only close the panel this save belongs to: Cancel/Escape stay live while it saves, and the PM may
      // have closed it or opened Manage on another file since (F185).
      if (docState === st) closeDocPanel();
    }).catch(function (err) {
      if (docState === st) approveDone();
      refreshCovStats();
      var lead = done ? 'Saved ' + done + ' of ' + groups.length + ' products, then c' : 'C';
      if (isConflict(err)) showError('<span>' + lead + 'ouldn’t save the rest: another PM changed some of these links meanwhile. The panel now has the current links — click “Approve &amp; save” again.</span>', 12000);
      else showError('<span>' + lead + 'ouldn’t finish saving (' + escapeHtml(String((err && err.message) || err)) + '). Click “Approve &amp; save” again to finish the rest.</span>', 9000);
      loadCatalog({ fresh: true }).then(function () { refreshCovStats(); repaintProgressSafely(); });
    });
    });
  }
  // A self-contained bucket instance (carries its own url/thumb/id so existing files render without the folder listing).
  function wireInst(f, existing) {
    return { id: ++wireIid, name: f.name, url: f.url, thumb: f.thumb || null, isImage: !!f.isImage, fileId: (f.fileId != null ? f.fileId : null), existing: !!existing, folder: !!f.folder };
  }
  function wireFindInst(ctx, id) {
    var found = null;
    Object.keys(ctx.assign).forEach(function (mk) { ctx.assign[mk].forEach(function (x) { if (String(x.id) === String(id)) found = x; }); });
    return found;
  }
  function wireBucketLabel(ctx, key) {
    if (key === "__none") return "Not on any model";
    var t = null; (ctx.targets || []).forEach(function (x) { if (x.key === key) t = x; });
    if (!t) return key;
    return t.own ? t.model : (t.fam + " › " + t.prod + " · " + t.model);
  }
  function wireMoveBtn(inst, currentKey) {
    // A move ICON that opens a custom themed dropdown (see wireOpenMoveMenu) listing the
    // other model buckets. data-cur is the file's current bucket, excluded from the menu.
    return '<button class="wire-move" type="button" data-id="' + inst.id + '" data-cur="' + escapeAttr(currentKey) + '"' +
      ' aria-haspopup="menu" aria-expanded="false" title="Move to another model"' +
      ' aria-label="Move ' + escapeAttr(inst.name) + ' to another model">' + MOVEICON + '</button>';
  }
  function wireFileChip(inst, ctx, currentKey) {
    var isFolder = inst.folder;
    var stateCls = isFolder ? ' folder' : (inst.existing ? ' existing' : ' isnew');
    var tag = isFolder ? '<span class="wire-tag">folder</span>'
      : (inst.existing ? '<span class="wire-tag">existing</span>' : '<span class="wire-tag new">new</span>');
    var body = (!isFolder && inst.isImage && inst.thumb ? '<img src="' + escapeAttr(inst.thumb) + '" alt="" loading="lazy" decoding="async" draggable="false">' : '<span class="wire-ficon">' + svg(isFolder ? ICONS.render : ICONS.spec) + '</span>') +
      '<span class="wire-fname">' + escapeHtml(inst.name) + '</span>';
    // Opener is a SPAN (not a <button>): a <button> as the chip's main grab area swallows the parent's
    // HTML5 drag (browsers won't start a drag from a form control), which made files impossible to drag
    // into another model/product. As a span, the whole chip drags from its body; click/Enter still previews.
    var opener = isFolder
      ? '<span class="wire-open wire-open-static">' + body + '</span>'
      : '<span class="wire-open" role="button" tabindex="0" data-id="' + inst.id + '" title="' + escapeAttr(inst.name) + '" aria-label="Preview ' + escapeAttr(inst.name) + '">' + body + '</span>';
    // Only the folder chip carries an outer title (it explains the dashed state — genuinely additive).
    // For a file the outer title just repeated the inner opener's, giving AT the filename twice; the
    // opener keeps it, which is also where the pointer lands to read an ellipsised .wire-fname.
    return '<span class="wire-file' + stateCls + '"' +
      ' draggable="' + (isFolder ? 'false' : 'true') + '" data-id="' + inst.id + '"' +
      (isFolder ? ' title="This model is linked to a whole auto-updating folder"' : '') + '>' +
      opener + tag +
      (isFolder ? '' : wireMoveBtn(inst, currentKey)) +
      (isFolder ? '' : '<button class="wire-dup" type="button" data-id="' + inst.id + '" title="Duplicate — then move the copy to another model" aria-label="Duplicate ' + escapeAttr(inst.name) + '">' + DUPICON + '</button>') +
      (isFolder ? '' : '<button class="wire-del" type="button" data-id="' + inst.id + '" title="Remove this file" aria-label="Remove ' + escapeAttr(inst.name) + '">' + DELICON + '</button>') +
      '</span>';
  }
  function renderWirePreview(box, ctx) {
    var result = box.querySelector("#wireResult");
    var single = ctx.models.length === 1; // one model → everything goes to it
    ctx.targets = ctx.models.map(function (m) { return { key: wireKey(ctx.fam, ctx.prod, m), fam: ctx.fam, prod: ctx.prod, model: m, own: true }; });
    var assign = { __none: [] }; ctx.targets.forEach(function (t) { assign[t.key] = []; });
    var seen = { __none: {} }; ctx.targets.forEach(function (t) { seen[t.key] = {}; }); // per-bucket identity set
    // 1) Show each model's ALREADY-WIRED content first (preserved on save); folder → a locked chip.
    ctx.targets.forEach(function (t) {
      var ex = (ctx.existing && ctx.existing[t.model]) || { files: [], hasFolder: false, folderUrl: "" };
      wireFolderChips(ex).forEach(function (c) { assign[t.key].push(c); });
      ex.files.forEach(function (f) { var k = wireIdent(f); if (!seen[t.key][k]) { seen[t.key][k] = 1; assign[t.key].push(wireInst(f, true)); } });
    });
    // 2) New folder files, auto-sorted — skip any already wired to that model (de-dupe by file identity).
    ctx.files.forEach(function (f) {
      var target = "__none";
      if (single) target = ctx.targets[0].key;
      else { for (var i = 0; i < ctx.models.length; i++) { if (matchesModel(f.name, ctx.models[i])) { target = ctx.targets[i].key; break; } } }
      var k = wireIdent(f);
      if (seen[target] && seen[target][k]) return; // already present in this bucket → don't duplicate
      if (seen[target]) seen[target][k] = 1;
      assign[target].push(wireInst(f, false));
    });
    ctx.assign = assign;
    result.innerHTML =
      '<div class="wire-scope">Linking the <b>' + escapeHtml(ctx.resource) + '</b> resource for <b>' + escapeHtml(ctx.fam) + ' › ' + escapeHtml(ctx.prod) + '</b></div>' +
      '<div class="wire-summary"><b>Nothing is saved yet</b> — review below, then <b>Approve &amp; save</b> at the bottom. <b>Drag</b> a file onto the right model, <b>click</b> it to preview, or use the <b>copy</b> icon to put one file on several models. Files tagged <b>“existing”</b> are already saved and kept (nothing is duplicated); <b>“Sort into another product”</b> sends a stray file elsewhere; “Not on any model” files won’t show on the site.</div>' +
      '<div id="wireBuckets"></div>' +
      '<div class="wire-elsewhere">' +
        '<span class="wire-elsewhere-lbl">Move a file to another product?</span>' +
        '<select class="wire-sel" id="wireOtherProd" aria-label="Another product"></select>' +
        '<select class="wire-sel" id="wireOtherModel" aria-label="Model in the other product"></select>' +
        '<button class="wire-btn" id="wireAddTarget" type="button">Add this model</button>' +
      '</div>' +
      '<div class="wire-elsewhere-hint">Pick the product &amp; model, then <b>Add this model</b> — a new bucket appears below (highlighted); <b>drag the file into it</b> (or use each file’s <b>Move</b> button). There’s no drop target until you add the model.</div>' +
      '<div class="wire-approve"><button class="wire-btn primary" id="wireApprove" type="button">Approve &amp; save</button></div>';
    wireInitElsewhere(box, ctx);
    renderWireBuckets(box, ctx);
    result.querySelector("#wireApprove").addEventListener("click", function () { wireApprove(ctx, this); });
  }
  function renderWireBuckets(box, ctx) {
    var host = box.querySelector("#wireBuckets"); if (!host) return;
    // Record whose context this panel is currently showing, so a slow async load started for an EARLIER
    // context can tell it has been superseded and stop repainting (see wireAddTarget).
    box._wireCtx = ctx;
    function bucketByKey(key, label, insts, isNone) {
      var hasFolder = insts.some(function (i) { return i.folder; });
      var hasFile = insts.some(function (i) { return !i.folder; });
      var note = (hasFolder && hasFile) ? '<div class="wire-note wire-warn">' + svg(ICONS.alert) + '<span><b>Heads up:</b> this model is linked to an <b>auto-updating folder</b>. Approving switches it to the <b>specific files</b> below — new files added to the folder later will <b>no longer appear automatically</b>. Remove the added files to keep the folder link.</span></div>' : '';
      var actions = '';
      if (isNone && insts.length) {
        note = '<div class="wire-note">These aren’t on any model. On <b>Approve</b> they’ll be parked in the <b>Unsorted</b> pile to sort later — or handle them now:</div>';
        actions = '<div class="wire-none-actions">' +
          '<button class="wire-btn" id="wireKeepAll" type="button">Keep all → Unsorted</button>' +
          '<button class="wire-btn danger" id="wireDeleteAll" type="button">Discard all</button></div>';
      }
      return '<div class="wire-bucket' + (isNone ? ' wire-unmatched' : '') + '" data-key="' + escapeAttr(key) + '">' +
        '<div class="wire-model-name">' + label + ' <b>' + insts.length + '</b></div>' + note + actions +
        '<div class="wire-files">' + (insts.length ? insts.map(function (i) { return wireFileChip(i, ctx, key); }).join("") : '<span class="wire-none">drop files here</span>') + '</div></div>';
    }
    var html = ctx.targets.map(function (t) {
      var label = t.own ? escapeHtml(t.model)
        : '<span class="wire-otherpath">' + escapeHtml(t.fam) + ' › ' + escapeHtml(t.prod) + ' · </span>' + escapeHtml(t.model);
      return bucketByKey(t.key, label, ctx.assign[t.key] || [], false);
    }).join("");
    html += bucketByKey("__none", "Not on any model", ctx.assign.__none, true);
    host.innerHTML = html;
    [].forEach.call(host.querySelectorAll(".wire-bucket"), function (bk) {
      bk.addEventListener("dragover", function (e) { e.preventDefault(); bk.classList.add("drop"); });
      // dragleave also fires when the pointer moves onto a CHILD of the bucket, which made the drop
      // highlight flicker across a bucket that already had chips. Only clear when the pointer has really
      // left the bucket's subtree (relatedTarget is the node it moved TO).
      bk.addEventListener("dragleave", function (e) {
        var to = e && e.relatedTarget;
        if (to && bk.contains(to)) return;
        bk.classList.remove("drop");
      });
      bk.addEventListener("drop", function (e) { e.preventDefault(); bk.classList.remove("drop"); wireDrop(box, ctx, bk.getAttribute("data-key")); });
    });
    [].forEach.call(host.querySelectorAll(".wire-file"), function (chip) {
      var id = chip.getAttribute("data-id"), inst = wireFindInst(ctx, id);
      if (!inst || inst.folder) return; // folder chip is a locked, non-draggable marker
      chip.addEventListener("dragstart", function (e) { wireDrag = id; try { e.dataTransfer.setData("text/plain", id); } catch (x) {} chip.classList.add("dragging"); });
      // Clear wireDrag on dragend: a drag released outside a bucket left the id set, so the NEXT drop of
      // anything at all on a bucket (text, an image, a file from the desktop) silently moved that chip.
      chip.addEventListener("dragend", function () { chip.classList.remove("dragging"); wireDrag = null; });
    });
    [].forEach.call(host.querySelectorAll(".wire-open"), function (op) {
      var inst = wireFindInst(ctx, op.getAttribute("data-id"));
      if (!inst) return; // folder opener has no data-id → skipped
      op.addEventListener("click", function () { if (inst.url) openPreview(inst.url, inst.name); });
      op.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); if (inst.url) openPreview(inst.url, inst.name); } });
    });
    [].forEach.call(host.querySelectorAll(".wire-move"), function (btn) {
      btn.addEventListener("click", function (e) { e.stopPropagation(); wireOpenMoveMenu(btn, box, ctx); });
    });
    [].forEach.call(host.querySelectorAll(".wire-dup"), function (btn) {
      btn.addEventListener("click", function (e) { e.stopPropagation(); wireDuplicate(box, ctx, btn.getAttribute("data-id")); });
    });
    [].forEach.call(host.querySelectorAll(".wire-del"), function (btn) {
      btn.addEventListener("click", function (e) { e.stopPropagation(); wireRemove(box, ctx, btn.getAttribute("data-id")); });
    });
    var keepBtn = host.querySelector("#wireKeepAll");
    if (keepBtn) keepBtn.addEventListener("click", function () { wireKeepAllUnsorted(box, ctx, this); });
    var delAllBtn = host.querySelector("#wireDeleteAll");
    if (delAllBtn) delAllBtn.addEventListener("click", function () {
      var n = (ctx.assign.__none || []).filter(function (i) { return !i.folder; }).length;
      pfConfirmIf(n > 0, {
        title: "Discard unassigned files?",
        body: ["Discard " + n + " file" + (n === 1 ? "" : "s") + " that aren’t on any model?",
               "They won’t be saved, and they won’t be parked in the pile either."],
        confirmLabel: "Discard", danger: true
      }).then(function (ok) {
        if (!ok) return;
        ctx.assign.__none = []; renderWireBuckets(box, ctx);
      });
    });
  }
  function wireRemove(box, ctx, id) {
    Object.keys(ctx.assign).forEach(function (mk) { ctx.assign[mk] = ctx.assign[mk].filter(function (x) { return String(x.id) !== String(id); }); });
    renderWireBuckets(box, ctx);
  }
  function wireDrop(box, ctx, targetKey) {
    var id = wireDrag; wireDrag = null; if (!id) return;
    var inst = null;
    Object.keys(ctx.assign).forEach(function (mk) {
      ctx.assign[mk] = ctx.assign[mk].filter(function (x) { if (String(x.id) === String(id)) { inst = x; return false; } return true; });
    });
    if (!inst) return;
    (ctx.assign[targetKey] || ctx.assign.__none).push(inst);
    renderWireBuckets(box, ctx);
  }
  // Keyboard/screen-reader move (via the per-chip "Move" select): same effect as a drag+drop.
  function wireMoveTo(box, ctx, id, targetKey) {
    var inst = null;
    Object.keys(ctx.assign).forEach(function (mk) {
      ctx.assign[mk] = ctx.assign[mk].filter(function (x) { if (String(x.id) === String(id)) { inst = x; return false; } return true; });
    });
    if (!inst) return;
    (ctx.assign[targetKey] || ctx.assign.__none).push(inst);
    renderWireBuckets(box, ctx);
    if (srAnnounce) srAnnounce.textContent = "Moved " + inst.name + " to " + wireBucketLabel(ctx, targetKey);
    // Re-render replaced the DOM — return focus to the moved chip so keyboard flow continues.
    var host = box.querySelector("#wireBuckets");
    var moved = host && host.querySelector('.wire-file[data-id="' + inst.id + '"] .wire-open');
    if (moved && moved.focus) moved.focus();
  }
  function wireCloseMoveMenu(focusTrigger) {
    if (!wireMenu) return;
    var m = wireMenu, trigger = m._trigger; wireMenu = null;
    document.removeEventListener("keydown", m._onKey, true);
    document.removeEventListener("mousedown", m._onDoc, true);
    window.removeEventListener("scroll", m._onScroll, true);
    window.removeEventListener("resize", m._onScroll, true);
    if (m.parentNode) m.parentNode.removeChild(m);
    if (trigger) { trigger.setAttribute("aria-expanded", "false"); trigger.removeAttribute("aria-controls"); if (focusTrigger && trigger.focus) trigger.focus(); }
  }
  // Custom themed Move dropdown: full control over the look (unlike a native <select> popup),
  // positioned fixed near the trigger so it's never clipped, keyboard + screen-reader operable.
  var _wireMenuSeq = 0;
  function wireOpenMoveMenu(btn, box, ctx) {
    if (wireMenu) { var same = wireMenu._trigger === btn; wireCloseMoveMenu(false); if (same) return; } // toggle
    var id = btn.getAttribute("data-id"), curKey = btn.getAttribute("data-cur");
    var keys = (ctx.targets || []).map(function (t) { return t.key; }).concat(["__none"]);
    var menu = document.createElement("div");
    menu.id = "wireMoveMenu" + (++_wireMenuSeq); // so the trigger can point aria-controls at it
    menu.className = "wire-menu"; menu.setAttribute("role", "menu"); menu.setAttribute("aria-orientation", "vertical"); menu.setAttribute("aria-label", "Move to a model");
    var html = '<div class="wire-menu-head">Move to…</div>';
    // "Copy to every model" — fan this file out to all of the selected product's models at once.
    var ownCount = (ctx.targets || []).filter(function (t) { return t.own; }).length;
    if (ownCount >= 2) html += '<button class="wire-menu-item wire-menu-all" type="button" role="menuitem" data-key="__all">Copy to every model</button>';
    keys.forEach(function (k) {
      if (k === curKey) return; // can't move to the bucket it's already in
      var none = (k === "__none");
      html += '<button class="wire-menu-item' + (none ? ' wire-menu-none' : '') + '" type="button" role="menuitem" data-key="' + escapeAttr(k) + '">' + escapeHtml(wireBucketLabel(ctx, k)) + '</button>';
    });
    menu.innerHTML = html;
    document.body.appendChild(menu);
    // Reposition on scroll/resize instead of closing (the menu is position:fixed); close only if the trigger
    // scrolls out of view. Internal menu scrolling is excluded in _onScroll below.
    function place() {
      var r = btn.getBoundingClientRect();
      var vw = document.documentElement.clientWidth, vh = document.documentElement.clientHeight;
      if (r.bottom < 0 || r.top > vh) { wireCloseMoveMenu(false); return; }
      var left = Math.min(r.left, vw - menu.offsetWidth - 8); if (left < 8) left = 8;
      var top = r.bottom + 4;
      if (top + menu.offsetHeight > vh - 8) { var up = r.top - menu.offsetHeight - 4; top = up >= 8 ? up : Math.max(8, vh - menu.offsetHeight - 8); }
      menu.style.left = left + "px"; menu.style.top = top + "px";
    }
    place();
    var items = [].slice.call(menu.querySelectorAll(".wire-menu-item"));
    function focusItem(i) { if (items.length) items[(i + items.length) % items.length].focus(); }
    items.forEach(function (it) {
      it.addEventListener("click", function () { var key = it.getAttribute("data-key"); wireCloseMoveMenu(false); if (key === "__all") wireToAllModels(box, ctx, id); else wireMoveTo(box, ctx, id, key); });
    });
    menu._trigger = btn;
    menu._onKey = function (e) {
      if (e.key === "Escape") { e.preventDefault(); wireCloseMoveMenu(true); return; }
      var idx = items.indexOf(document.activeElement);
      if (e.key === "ArrowDown") { e.preventDefault(); focusItem(idx < 0 ? 0 : idx + 1); }
      else if (e.key === "ArrowUp") { e.preventDefault(); focusItem(idx < 0 ? items.length - 1 : idx - 1); }
      else if (e.key === "Home") { e.preventDefault(); focusItem(0); }
      else if (e.key === "End") { e.preventDefault(); focusItem(items.length - 1); }
      else if (e.key === "Tab") { wireCloseMoveMenu(true); } // restore focus to the trigger, not <body>
    };
    menu._onDoc = function (e) { if (!menu.contains(e.target) && !btn.contains(e.target)) wireCloseMoveMenu(false); };
    // Close when the PAGE scrolls (the menu is position:fixed), but NOT when scrolling INSIDE the menu
    // itself (it can be taller than 60vh and scroll) — else lower items become unreachable. Capture-phase
    // scroll fires for descendants too, so exclude scrolls originating in the menu by target.
    menu._onScroll = function (e) { if (e && e.type !== "resize" && menu.contains(e.target)) return; place(); }; // reposition, don't close
    document.addEventListener("keydown", menu._onKey, true);
    document.addEventListener("mousedown", menu._onDoc, true);
    window.addEventListener("scroll", menu._onScroll, true);
    window.addEventListener("resize", menu._onScroll, true);
    wireMenu = menu;
    btn.setAttribute("aria-expanded", "true"); btn.setAttribute("aria-controls", menu.id);
    focusItem(0);
  }
  // Composite bucket key so models from different products can coexist as drop targets.
  function wireKey(fam, prod, model) { return fam + "§" + prod + "§" + model; }
  function wireModelsOf(fam, prod) {
    var out = [];
    CATALOG.forEach(function (c) { c.families.forEach(function (f) { if (f.name !== fam) return; f.products.forEach(function (p) { if (p.name === prod) out = p.models.slice(); }); }); });
    return out;
  }
  // "Sort into another product" picker: populate product + dependent model selects, wire Add.
  function wireInitElsewhere(box, ctx) {
    var ps = box.querySelector("#wireOtherProd"), ms = box.querySelector("#wireOtherModel");
    if (!ps || !ms) return;
    ps.innerHTML = '<option value="">Choose a product…</option>';
    CATALOG.forEach(function (c) { c.families.forEach(function (f) { f.products.forEach(function (p) {
      var o = document.createElement("option"); o.value = f.name + "§" + p.name; o.textContent = f.name + " › " + p.name; ps.appendChild(o);
    }); }); });
    function fillModels() {
      var v = ps.value;
      if (!v) { ms.innerHTML = '<option value="">Model…</option>'; return; }
      var parts = v.split("§"), models = wireModelsOf(parts[0], parts[1]);
      ms.innerHTML = models.length
        ? '<option value="">Model…</option>' + models.map(function (m) { return '<option value="' + escapeAttr(m) + '">' + escapeHtml(m) + '</option>'; }).join("")
        : '<option value="">(no models in sheet)</option>';
    }
    ps.addEventListener("change", fillModels); fillModels();
    box.querySelector("#wireAddTarget").addEventListener("click", function () { wireAddTarget(box, ctx); });
  }
  function wireScrollToBucket(box, key) {
    var buckets = box.querySelectorAll(".wire-bucket");
    for (var i = 0; i < buckets.length; i++) {
      if (buckets[i].getAttribute("data-key") === key) {
        var bk = buckets[i];
        if (bk.scrollIntoView) bk.scrollIntoView({ block: "nearest" });
        bk.classList.remove("justadded"); void bk.offsetWidth; bk.classList.add("justadded"); // (re)start the flash
        break;
      }
    }
  }
  function wireAddTarget(box, ctx) {
    var ps = box.querySelector("#wireOtherProd"), ms = box.querySelector("#wireOtherModel");
    var pv = ps.value, model = ms.value;
    if (!pv || !model) { showToast('<span>Pick a product and a model to add.</span>'); return; }
    var parts = pv.split("§"), fam = parts[0], prod = parts[1], key = wireKey(fam, prod, model);
    if (ctx.assign[key]) { wireScrollToBucket(box, key); return; } // already a target
    ctx.assign[key] = [];
    ctx.targets.push({ key: key, fam: fam, prod: prod, model: model, own: (fam === ctx.fam && prod === ctx.prod) });
    renderWireBuckets(box, ctx);
    wireScrollToBucket(box, key);
    // Load that model's already-wired files so dropping onto it MERGES instead of overwriting.
    //
    // ctx.existing MUST get an entry either way. wireApprove reads it for two things: to refuse writing a
    // model whose current links it couldn't read (`exState.error`), and to work out how many links the save
    // intends to drop (`expectRemoved`). A model added here used to get no entry at all, which reads as
    // "no links" — so the error guard was structurally unreachable for a cross-product target, and a failed
    // or half-finished read let the dragged files be written as that model's COMPLETE set. Start from
    // "unreadable" and only clear it once the read actually lands.
    ctx.existing = ctx.existing || {};
    ctx.existing[model] = { files: [], hasFolder: false, folderUrl: "", error: true };
    if (window.PF_getToken) {
      window.PF_getToken().then(function (token) { return wireLoadExisting(model, ctx.resource, "Bearer " + token); })
        .then(function (ex) {
          // The wire tool can be rebuilt underneath this read (another product opened, the panel
          // re-rendered). Repainting a SUPERSEDED context into the fresh buckets showed a screen that no
          // longer matched what Approve would write.
          if (!document.body.contains(box)) return;
          if (box._wireCtx && box._wireCtx !== ctx) return;
          ctx.existing[model] = ex;
          if (!ctx.assign[key]) return;
          var seen = {}; ctx.assign[key].forEach(function (i) { seen[wireIdent(i)] = 1; });
          if (!ctx.assign[key].some(function (i) { return i.folder; }))
            wireFolderChips(ex).reverse().forEach(function (c) { ctx.assign[key].unshift(c); });
          ex.files.forEach(function (f) { var k = wireIdent(f); if (!seen[k]) { seen[k] = 1; ctx.assign[key].push(wireInst(f, true)); } });
          renderWireBuckets(box, ctx);
        }, function () { /* stays flagged unreadable — wireApprove will skip this model and say so */ });
    }
  }
  function wireDuplicate(box, ctx, id) {
    var src = null, bk = null;
    Object.keys(ctx.assign).forEach(function (mk) { ctx.assign[mk].forEach(function (x) { if (String(x.id) === String(id)) { src = x; bk = mk; } }); });
    if (!src || src.folder) return;
    ctx.assign[bk].push(wireInst(src, false)); // a fresh copy (drag it onto another model)
    renderWireBuckets(box, ctx);
  }
  // "Copy to every model" — fan one file out to ALL models of the SELECTED product at once (e.g. a
  // spec sheet identical across models), instead of dragging it onto each model one by one. Only the
  // selected product's own models (t.own); skips any model that already has this file.
  function wireToAllModels(box, ctx, id) {
    var src = null;
    Object.keys(ctx.assign).forEach(function (mk) { ctx.assign[mk].forEach(function (x) { if (String(x.id) === String(id)) src = x; }); });
    if (!src || src.folder) return;
    var ident = wireIdent(src), n = 0;
    (ctx.targets || []).forEach(function (t) {
      if (!t.own) return;
      var bucket = ctx.assign[t.key] || (ctx.assign[t.key] = []);
      if (bucket.some(function (i) { return wireIdent(i) === ident; })) return; // already on this model
      bucket.push(wireInst(src, false));
      n++;
    });
    renderWireBuckets(box, ctx);
    showToast(svg(ICONS.ext) + '<span>Copied “' + escapeHtml(src.name || "file") + '” to ' + n + ' model' + (n === 1 ? "" : "s") + '.</span>');
  }
  function wireApprove(ctx, btn) {
    var btnDone = busyButton(btn, "Saving…");
    // The bucket already holds the desired final set (existing + new, merged). De-dupe by file
    // identity so the same file can't be written twice. Explicit files win over a preserved folder;
    // a lone folder marker is written back untouched (auto-update kept); truly empty → "" (clear).
    function valueFor(key) {
      var insts = ctx.assign[key] || [];
      var files = insts.filter(function (i) { return !i.folder; });
      var seen = {}, urls = [];
      // A still-present "Current folder — auto-updating" marker is kept ALONGSIDE any explicit files —
      // a supported mixed folder+file resource (openMixedGallery), and what the Files-to-place writer
      // already does. Dropping it whenever a file was added silently replaced the auto-updating folder
      // (and every render it surfaced) with that one file, and slipped past both shrink guards because a
      // 1-URL folder collapsing to a 1-URL file nets to zero removed. A PM who means to REPLACE the
      // folder removes its chip first, so it's simply absent here.
      // Every folder chip still present, not just the first (F041).
      insts.filter(function (i) { return i.folder; }).forEach(function (f) { var k = "f:" + window.PF.canonFolderUrl(f.url || ""); if (f.url && !seen[k]) { seen[k] = 1; urls.push(f.url); } });
      files.forEach(function (i) { var k = wireIdent(i); if (!seen[k] && i.url) { seen[k] = 1; urls.push(i.url); } });
      return urls.join(" ");
    }
    // Group by (family, product). Selected product clears emptied models (reflects removals);
    // added products are only written when they have content, so they can never be cleared/clobbered.
    var groups = {}, skipped = [];
    ctx.targets.forEach(function (t) {
      var gk = t.fam + "§" + t.prod;
      if (!groups[gk]) groups[gk] = { fam: t.fam, prod: t.prod, own: t.own, assignments: {}, expectRemoved: 0, expectRemovedByModel: {}, base: {} };
      var exState = (ctx.existing && ctx.existing[t.model]) || null;
      var val = valueFor(t.key);
      // If we couldn't read this model's current links during preview, leave its cell completely
      // untouched — writing here would either clear the real link (val="") or clobber unread files.
      if (exState && exState.error) { if (val) skipped.push(t.model); return; }
      // Only a model whose links actually CHANGE is sent (F037). Every model of the product used to be — the
      // server re-wrote their rows and What's New showed all of them as new, and a model with no row for this
      // resource failed a save that had otherwise worked.
      if (exState && window.PF.sameUrlSet(wireExistingUrls(exState), window.PF.splitCellUrls(val))) return;
      if (t.own || val) {
        groups[gk].assignments[t.model] = val;
        // Compare-and-swap pre-image: only for models whose current links were actually read. A model
        // added from another product starts as { error: true } until its read lands, so an unread one is
        // left to the shrink guard rather than sending a false "" base.
        if (exState && !exState.error) {
          groups[gk].base[t.model] = wireExistingUrls(exState);
        }
        // Declare the links this save intends to drop on that model — a cleared cell, or chips the PM
        // removed from the bucket. The server refuses anything beyond this, so a preview that has gone
        // stale can't take links added since with it.
        var before = exState ? wireExistingUrls(exState).length : 0;
        // Count URLs the way the SERVER parses the cell (splitCellUrls), not a raw whitespace split — a URL
        // containing a raw space would otherwise over-count `after`, making a real removal look like it
        // dropped fewer links than it did, so the shrink guard refuses it (409) every time.
        var after = val ? (window.PF.splitCellUrls(val) || []).length : 0;
        if (before > after) { groups[gk].expectRemoved += before - after; groups[gk].expectRemovedByModel[t.model] = (groups[gk].expectRemovedByModel[t.model] || 0) + (before - after); }
      }
    });
    var list = Object.keys(groups).map(function (k) { return groups[k]; }).filter(function (g) { return Object.keys(g.assignments).length; });
    // Leftovers still in "Not on any model" get parked in the pile (so nothing is dropped).
    var noneItems = wireUnsortedItems(ctx);
    if (!list.length && !noneItems.length) {
      btnDone();
      if (skipped.length) { showError('<span>Couldn’t read the current links for ' + skipped.length + ' model' + (skipped.length === 1 ? '' : 's') + ' just now — nothing was saved, so their existing links are untouched. Please re-open the folder and try again in a moment.</span>', 9000); }
      else showToast('<span>Nothing to save yet — drag files onto a model, or use the Unsorted buttons below.</span>');
      return;
    }
    // Guard against silent link loss: if Approve would clear a model that currently has a link (the PM
    // removed its chips), confirm first — mirrors the section editor's delete confirmation.
    var clears = [];
    ctx.targets.forEach(function (t) {
      if (!t.own) return; // only the selected product's models can be cleared
      var ex = (ctx.existing && ctx.existing[t.model]) || null;
      if (!ex || ex.error) return;
      var had = (ex.files && ex.files.length) || ex.hasFolder;
      if (had && !valueFor(t.key)) clears.push(t.model);
    });
    pfConfirmIf(clears.length > 0, {
      title: "Remove existing links?",
      body: ["This will remove the existing " + ctx.resource + " link" + (clears.length === 1 ? "" : "s") + " from " + clears.length + " model" + (clears.length === 1 ? "" : "s") + ": " + clears.slice(0, 8).join(", ") + (clears.length > 8 ? ", …" : "") + ".",
             "The files stay in SharePoint — only the links are cleared."],
      confirmLabel: "Remove links", danger: true
    }).then(function (ok) {
    if (!ok) { btnDone(); return; }

    var wroteModels = {}; // models the SERVER confirmed it wrote — the only ones safe to clear health for
    Promise.resolve().then(function () {
      var updated = 0, prods = 0;
      return list.reduce(function (chain, g) {
        return chain.then(function () {
          // Compare-and-swap base only when this is a SINGLE product: a multi-product save's "Approve
          // again to finish the rest" retry must stay idempotent, and a committed product's base would go
          // stale and falsely conflict on the re-click. Multi-product falls back to the shrink guard.
          // postWire throws on a rejected link or a model with no resource row, so a 200 that did not
          // write everything can no longer read as success here.
          return postWire(g.fam, g.prod, ctx.resource, g.assignments, { expectRemoved: g.expectRemoved || 0, expectRemovedByModel: g.expectRemovedByModel || {}, base: (list.length === 1 ? (g.base || null) : null) })
            .then(function (res) {
              updated += (res.updated || 0); if (res.updated) prods++;
              res.wrote.forEach(function (m) { wroteModels[m] = 1; }); // clear health only for these
            });
        });
      }, Promise.resolve()).then(function () {
        if (!noneItems.length) return { updated: updated, prods: prods, pileAdded: 0 };
        // A failed park is REPORTED: these files are on no model, and the bucket holding them was about to be
        // re-rendered away under a "Saved." toast (F047/F042).
        return apiFetch("/api/unsorted", { method: "POST", body: { action: "add", items: noneItems } })
          .then(function (pile) { return { updated: updated, prods: prods, pileAdded: (pile && pile.added) || 0, pileFailed: false }; },
                function () { return { updated: updated, prods: prods, pileAdded: 0, pileFailed: true }; });
      });
    }).then(function (res) {
      if (res.pileFailed) {
        // Leave the panel up: the leftovers are still listed under "Not on any model", where Keep all →
        // Unsorted retries the park. Links that did save still refresh.
        btnDone();
        Object.keys(wroteModels).forEach(function (model) { clearHealthFor(model + "::" + ctx.resource); });
        showError('<span>' + (res.updated ? 'Saved ' + res.updated + ' ' + escapeHtml(ctx.resource) + ' row' + (res.updated === 1 ? '' : 's') + ', but c' : 'C') + 'ouldn’t park the ' + noneItems.length + ' file' + (noneItems.length === 1 ? '' : 's') + ' left in “Not on any model” in the Unsorted pile. They’re still listed below — use Keep all → Unsorted to try again.</span>', 12000);
        loadCatalog({ fresh: true }).then(refreshCovStats);
        return;
      }
      var msg = res.updated ? ('Saved — ' + res.updated + ' ' + escapeHtml(ctx.resource) + ' row' + (res.updated === 1 ? '' : 's') + (res.prods > 1 ? ' across ' + res.prods + ' products' : '') + ' updated.') : 'Saved.';
      if (res.pileAdded) msg += ' ' + res.pileAdded + ' unsorted file' + (res.pileAdded === 1 ? '' : 's') + ' parked.';
      if (skipped.length) msg += ' ⚠ ' + skipped.length + ' model' + (skipped.length === 1 ? '' : 's') + ' couldn’t be read and ' + (skipped.length === 1 ? 'was' : 'were') + ' left unchanged — re-run to save changes to ' + (skipped.length === 1 ? 'it' : 'them') + '.';
      toastSaved(msg, skipped.length ? 9000 : undefined); // hold the partial-save warning long enough to read
      // Drop any pre-fix health flag on the models the server CONFIRMED it wrote (survives the reload
      // re-seed). Clearing by intent rather than by outcome hid a still-broken link on a model whose row
      // the write never touched.
      Object.keys(wroteModels).forEach(function (model) { clearHealthFor(model + "::" + ctx.resource); });
      loadCatalog({ fresh: true }).then(function () { if (view === "progress") renderCoverage(); });
    }).catch(function (err) {
      btnDone();
      // A multi-product save is non-atomic: earlier products may have committed before this failure. Refresh
      // (cache-bypassing) so anything already saved is reflected, and invite a retry — re-sending an
      // already-saved product is idempotent, so Approve again finishes the rest without a full page reload.
      if (isConflict(err)) showError('<span>Another PM changed these links since you opened the folder, so nothing was saved. Click Preview again to see the current links, then sort.</span>', 12000);
      else showError('<span>Couldn’t finish saving (' + escapeHtml(String((err && err.message) || err)) + '). Some products may have saved — click “Approve &amp; save” again to finish the rest.</span>', 9000);
      // Refresh the DATA only. Re-rendering Progress here threw away the preview — and every file the PM
      // had sorted — that the message tells them to click again (F026).
      loadCatalog({ fresh: true }).then(refreshCovStats);
    });
    });
  }
  // Payloads for the pile built from the current "Not on any model" files (skip folder markers).
  function wireUnsortedItems(ctx) {
    return (ctx.assign.__none || []).filter(function (i) { return !i.folder && i.url; }).map(function (i) {
      return { name: i.name, url: i.url, fileId: i.fileId || "", resource: ctx.resource, folder: ctx.url };
    });
  }
  function wireKeepAllUnsorted(box, ctx, btn) {
    var items = wireUnsortedItems(ctx);
    if (!items.length) { ctx.assign.__none = []; renderWireBuckets(box, ctx); return; }
    if (btn) { btn.disabled = true; btn.textContent = "Saving…"; }
    apiFetch("/api/unsorted", { method: "POST", body: { action: "add", items: items } })
      .then(function (res) {
        ctx.assign.__none = []; renderWireBuckets(box, ctx);
        showToast(svg(ICONS.ext) + '<span>' + (res.added || 0) + ' file' + (res.added === 1 ? '' : 's') + ' saved to the pile.</span>');
      }).catch(function (err) {
        if (btn) { btn.disabled = false; btn.textContent = "Keep all → Unsorted"; }
        showError('<span>Couldn’t save to Unsorted (' + escapeHtml(String((err && err.message) || err)) + ').</span>', 7000);
      });
  }

  /* ============================================================
     PM · UNSORTED FILES PILE
     ============================================================ */
  // ---- Progress tool: the "Unsorted files" pile — park/sort stray files that weren't on any model ----
  function unsortedPanel(bare) {
    var box = document.createElement("div"); box.className = "wire unsorted";
    box.innerHTML =
      (bare ? '' : '<div class="wire-head" role="heading" aria-level="2">' + svg(ICONS.folder) + '<span>The pile</span><span class="uns-count" id="unsCount"></span></div>') +
      '<p class="wire-lead">Stray files — parked from the folder tool, or found sitting in a shared folder with a name that matches no model. Preview each, then sort it into a product &amp; model; it’s linked there (merged with anything already saved) and drops off this list.</p>' +
      '<div id="unsList"><div class="wire-msg">Loading…</div></div>';
    loadUnsorted(box);
    return box;
  }
  function loadUnsorted(box) {
    var list = box.querySelector("#unsList"); if (!list) return;
    if (!window.PF_getToken) { list.innerHTML = '<div class="wire-msg err">Please sign in again.</div>'; return; }
    // The pile is two sources rendered as one list: files a PM parked from the folder tool (/api/unsorted)
    // AND orphan files the nightly scan found loose in a shared folder (resolved client-side). Orphans are
    // best-effort — resolveOrphanItems never rejects, so only an /api/unsorted failure surfaces an error.
    // Generation guard: a Save/Remove triggers a reload while an earlier one may still be resolving; only the
    // NEWEST load may paint, or a slower earlier response could clobber the list with stale rows.
    var live = loadGen(box);
    Promise.all([
      apiFetch("/api/unsorted").then(function (data) { return data.items || []; }),
      resolveOrphanItems(),
      // Server-side dismissals: an orphan "Hide" is tombstoned via /api/placements (action "dismiss"), so it
      // holds for every PM/device — not just the browser that hid it. Best-effort: an outage → no server
      // skips, which never breaks the pile (the local tombstone still hides it for this browser).
      placementsStore.get().catch(function () { return { skip: [] }; })
    ]).then(function (arr) {
      if (!live()) return; // superseded by a newer reload
      var parked = arr[0] || [], orphans = arr[1] || [];
      // Reconcile orphans against live state client-side (the cron's list can be ~24h stale): drop any whose
      // file is already linked to a model (Saved since the last scan) or already shown as a parked row, and
      // dedup orphan-vs-orphan by canonical file url. Unresolved rows (no url) can't be deduped — keep them.
      var wired = (window.PF && PF.wiredUrlSet) ? PF.wiredUrlSet(LINKS) : {};
      var skip = {}; ((arr[2] && arr[2].skip) || []).forEach(function (u) { if (window.PF) skip[PF.canonFileUrl(u)] = 1; });
      var taken = {}; parked.forEach(function (p) { if (p.url && window.PF) taken[PF.canonFileUrl(p.url)] = 1; });
      orphans = orphans.filter(function (o) {
        if (!o.url || !window.PF) return true;
        var c = PF.canonFileUrl(o.url);
        if (wired[c] || taken[c] || skip[c]) return false; // wired since, already parked, or dismissed elsewhere (server tombstone)
        taken[c] = 1; return true;
      });
      renderUnsorted(box, parked.concat(orphans));
    })
      .catch(function (err) { if (!live()) return; list.innerHTML = '<div class="wire-msg err">Couldn’t load the pile (' + escapeHtml(String((err && err.message) || err)) + ').</div>'; if (box._onError) box._onError(); });
  }
  // --- Orphan folder files (from the nightly cron): files sitting in a SHARED auto-linked folder whose
  // names match no model, so they surface nowhere. Folded into the Sort-the-Pile list as actionable rows:
  // resolve each one's real file URL from its folder so it can be previewed + Saved onto a model, exactly
  // like a parked file. A handled orphan is tombstoned locally so it drops off at once; the nightly scan
  // then stops flagging it too (a Saved file is an explicit link — see computeFolderActivity). ---
  // Canonicalize the folder URL so a dismissed orphan stays dismissed even if the cron later reports the
  // same folder under a different spelling / row order.
  function orphanKey(o) { var fu = o.folderUrl || ""; return (window.PF && PF.canonFolderUrl ? PF.canonFolderUrl(fu) : fu) + "|" + (o.file || o.name || ""); }
  function orphanDismissed() { try { var a = JSON.parse(localStorage.getItem("pf-orphan-done") || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
  function orphanDismiss(key) { var a = orphanDismissed(); if (a.indexOf(key) === -1) a.unshift(key); try { localStorage.setItem("pf-orphan-done", JSON.stringify(a.slice(0, 500))); } catch (e) {} }
  function orphanItem(o, f) {
    return { name: o.file, file: o.file, resource: o.resource, folderUrl: o.folderUrl, detail: o.detail,
      orphan: true, url: f ? f.url : "", fileId: f ? f.fileId : null, unresolved: !(f && f.url) };
  }
  // Resolve the pending (non-dismissed) orphans into row items. Lists each shared folder ONCE (deduped) to
  // find each file's real webUrl + id. A folder that can't be listed yields "unresolved" rows that fall back
  // to a rename hint (no Save). Best-effort: never rejects, so it can't break the pile.
  // Memoized by the pending-orphan signature so a plain Progress re-render (same ORPHANS + dismissals) reuses
  // the last resolution instead of re-hitting Graph on every paint; a catalog refresh or a dismiss changes
  // the signature and re-resolves. A token failure is NOT cached, so it retries next render.
  var _orphanCache = null; // { sig, items }
  function resolveOrphanItems() {
    var done = orphanDismissed();
    var pending = (ORPHANS || []).filter(function (o) { return done.indexOf(orphanKey(o)) === -1; });
    var sig = pending.map(orphanKey).sort().join("\n");
    if (_orphanCache && _orphanCache.sig === sig) return Promise.resolve(_orphanCache.items.slice());
    if (!pending.length || !window.PF_getToken) { _orphanCache = { sig: sig, items: [] }; return Promise.resolve([]); }
    return window.PF_getToken().then(function (token) {
      var bearer = "Bearer " + token, byFolder = {};
      pending.forEach(function (o) { (byFolder[o.folderUrl] = byFolder[o.folderUrl] || []).push(o); });
      return Promise.all(Object.keys(byFolder).map(function (fu) {
        return fetchFolderChildren(fu, bearer).then(function (files) {
          var byName = {}; files.forEach(function (f) { byName[f.name] = f; });
          return byFolder[fu].map(function (o) { return orphanItem(o, byName[o.file]); });
        }).catch(function () { return byFolder[fu].map(function (o) { return orphanItem(o, null); }); });
      })).then(function (groups) { var out = groups.reduce(function (a, b) { return a.concat(b); }, []); _orphanCache = { sig: sig, items: out }; return out.slice(); });
    }).catch(function () { return []; });
  }
  // --- Multi-select model picker (checkbox dropdown) for the Unsorted "Sort to models" control ---
  // One menu open at a time. `selected` is mutated in place so the row keeps a live reference to it.
  var unsMenu = null;
  function unsMbtnLabel(models, selected) {
    if (!selected.length) return "Choose models…";
    if (selected.length === models.length) return "All " + models.length + " models";
    return selected.length + " model" + (selected.length === 1 ? "" : "s") + " selected";
  }
  function unsCloseMenu(focusTrigger) {
    if (!unsMenu) return;
    var m = unsMenu, trigger = m._trigger; unsMenu = null;
    document.removeEventListener("keydown", m._onKey, true);
    document.removeEventListener("mousedown", m._onDoc, true);
    window.removeEventListener("scroll", m._onScroll, true);
    window.removeEventListener("resize", m._onScroll, true);
    if (m.parentNode) m.parentNode.removeChild(m);
    if (trigger) { trigger.setAttribute("aria-expanded", "false"); trigger.removeAttribute("aria-controls"); if (focusTrigger && trigger.focus) trigger.focus(); }
  }
  var _unsMenuSeq = 0;
  function unsOpenCheckMenu(btn, models, selected, onChange) {
    if (unsMenu) { var same = unsMenu._trigger === btn; unsCloseMenu(false); if (same) return; } // toggle
    var menu = document.createElement("div");
    // role="dialog", not "group"/"menu": the popup holds checkboxes, so a screen reader announcing a
    // "menu button" would promise arrow-key menuitem semantics this popup doesn't have. The trigger's
    // aria-haspopup="dialog" matches, and aria-controls is set below now that the element exists.
    menu.id = "unsMenu" + (++_unsMenuSeq);
    menu.className = "wire-menu uns-menu"; menu.setAttribute("role", "dialog"); menu.setAttribute("aria-label", "Choose models to sort into");
    var allOn = models.length > 0 && selected.length === models.length;
    menu.innerHTML = '<div class="wire-menu-head">Send to which models?</div>' +
      '<label class="uns-check uns-check-all"><input type="checkbox" class="uns-ca"' + (allOn ? " checked" : "") + '> Select all</label>' +
      models.map(function (m) { return '<label class="uns-check"><input type="checkbox" class="uns-cm" value="' + escapeAttr(m) + '"' + (selected.indexOf(m) !== -1 ? " checked" : "") + '> <span>' + escapeHtml(m) + '</span></label>'; }).join("");
    document.body.appendChild(menu);
    // Keep the menu pinned to its trigger. Reposition on scroll/resize instead of closing — closing made the
    // menu vanish the instant a PM scrolled a long Unsorted/Files-to-place list to reach the row. Close only
    // when the trigger itself scrolls out of view.
    function place() {
      var r = btn.getBoundingClientRect();
      var vw = document.documentElement.clientWidth, vh = document.documentElement.clientHeight;
      if (r.bottom < 0 || r.top > vh) { unsCloseMenu(false); return; }
      var left = Math.min(r.left, vw - menu.offsetWidth - 8); if (left < 8) left = 8;
      var top = r.bottom + 4;
      if (top + menu.offsetHeight > vh - 8) { var up = r.top - menu.offsetHeight - 4; top = up >= 8 ? up : Math.max(8, vh - menu.offsetHeight - 8); }
      menu.style.left = left + "px"; menu.style.top = top + "px";
    }
    place();
    var allBox = menu.querySelector(".uns-ca"), boxes = [].slice.call(menu.querySelectorAll(".uns-cm"));
    function refreshAll() { allBox.checked = selected.length === models.length && models.length > 0; allBox.indeterminate = selected.length > 0 && selected.length < models.length; }
    function sync() { selected.length = 0; boxes.forEach(function (b) { if (b.checked) selected.push(b.value); }); refreshAll(); onChange(); }
    allBox.addEventListener("change", function () { boxes.forEach(function (b) { b.checked = allBox.checked; }); sync(); });
    boxes.forEach(function (b) { b.addEventListener("change", sync); });
    refreshAll();
    menu._trigger = btn;
    // The menu is appended to <body>, so it sits AFTER everything in DOM order: tabbing off its last
    // checkbox walked straight out of the document. Contain it — at either end, close and hand focus
    // back to the trigger, from which Tab continues normally. Esc does the same from anywhere inside.
    menu._onKey = function (e) {
      if (e.key === "Escape") { e.preventDefault(); unsCloseMenu(true); return; }
      if (e.key !== "Tab" || !menu.contains(document.activeElement)) return;
      var focusables = [allBox].concat(boxes);
      var i = focusables.indexOf(document.activeElement);
      if (i === -1) return;
      var atEnd = e.shiftKey ? i === 0 : i === focusables.length - 1;
      if (atEnd) { e.preventDefault(); unsCloseMenu(true); }
    };
    menu._onDoc = function (e) { if (!menu.contains(e.target) && !btn.contains(e.target)) unsCloseMenu(false); };
    menu._onScroll = function (e) { if (e && e.type !== "resize" && menu.contains(e.target)) return; place(); }; // reposition, don't close
    document.addEventListener("keydown", menu._onKey, true);
    document.addEventListener("mousedown", menu._onDoc, true);
    window.addEventListener("scroll", menu._onScroll, true);
    window.addEventListener("resize", menu._onScroll, true);
    unsMenu = menu;
    btn.setAttribute("aria-expanded", "true"); btn.setAttribute("aria-controls", menu.id);
    if (allBox && allBox.focus) allBox.focus();
  }
  function renderUnsorted(box, items) {
    var list = box.querySelector("#unsList");
    var cnt = box.querySelector("#unsCount"); if (cnt) cnt.textContent = items.length ? (" " + items.length) : "";
    if (box._onCount) box._onCount(items.length); // report to the To-do container (hide the box when the pile is empty)
    if (!items.length) { list.innerHTML = '<div class="uns-empty">Nothing to sort right now — parked leftovers and loose folder files show up here.</div>'; return; }
    // Oldest-first so the longest-parked file leads (the API returns newest-first). ISO dates compare
    // lexically; an unknown date sorts last so it never masquerades as the oldest.
    items = items.slice().sort(function (a, b) { return String(a.addedAt || "9999").localeCompare(String(b.addedAt || "9999")); });
    list.innerHTML = items.map(function (it, i) {
      // Each row's controls are otherwise byte-identical across all files ("Product"/"Model"/"Sort in"/
      // "Delete") — fold the filename into every accessible name so screen-reader and voice-control users
      // can tell which file a control acts on.
      var nm = escapeAttr(it.name || "file");
      // Orphan rows (found loose in a shared folder) look the same but carry a "shared folder" tag + an
      // open-folder link; an unresolved one (folder wouldn't list) shows a rename hint instead of the picker.
      if (it.orphan) {
        var fu = safeUrl(it.folderUrl); // block javascript:/data: — keep the URL invariant even though orphans only come from listable folders
        var folderLink = fu ? '<a class="uns-folder-link" href="' + escapeAttr(fu) + '" target="_blank" rel="noopener" aria-label="Open the SharePoint folder for ' + nm + '" title="Open the folder in SharePoint">open folder ↗</a>' : "";
        var fileCell = it.url
          ? '<button class="uns-file" type="button" title="Preview">' + svg(fileTypeIcon(it.name)) + '<span class="uns-name">' + escapeHtml(it.name || "file") + '</span></button>'
          : '<span class="uns-file uns-file-plain">' + svg(fileTypeIcon(it.name)) + '<span class="uns-name">' + escapeHtml(it.name || "file") + '</span></span>';
        var sortCell = it.unresolved
          ? '<span class="uns-sort uns-orphan-hint">Couldn’t read this folder — open it ↗ and rename the file to start with the model name.</span>'
          : '<span class="uns-sort">' +
              '<select class="wire-sel uns-prod" aria-label="Product for ' + nm + '"></select>' +
              '<button class="wire-sel uns-mbtn" type="button" aria-haspopup="dialog" aria-expanded="false" aria-label="Choose models for ' + nm + '" disabled>Model…</button>' +
              '<button class="wire-btn uns-assign" type="button" aria-label="Save ' + nm + ' onto the selected models">Save</button>' +
            '</span>';
        return '<div class="uns-row uns-orphan" data-i="' + i + '">' +
          '<span class="uns-where">' + fileCell +
            '<span class="uns-meta uns-orphan-tag" title="' + escapeAttr(it.detail || "In a shared folder; its name matches no model, so it shows nowhere") + '">shared folder</span>' + folderLink +
          '</span>' + sortCell +
          '<button class="wire-del uns-del" type="button" title="Not a stray file — hide it" aria-label="Hide ' + nm + '">' + DELICON + '</button>' +
        '</div>';
      }
      return '<div class="uns-row" data-i="' + i + '">' +
        '<span class="uns-where">' +
          '<button class="uns-file" type="button" title="Preview">' + svg(fileTypeIcon(it.name)) + '<span class="uns-name">' + escapeHtml(it.name || "file") + '</span></button>' +
          (it.addedBy ? '<span class="uns-meta">' + escapeHtml(prettyName(it.addedBy)) + '</span>' : "") + ageChip(it.addedAt, "", true) +
        '</span>' +
        '<span class="uns-sort">' +
          '<select class="wire-sel uns-prod" aria-label="Product for ' + nm + '"></select>' +
          '<button class="wire-sel uns-mbtn" type="button" aria-haspopup="dialog" aria-expanded="false" aria-label="Choose models for ' + nm + '" disabled>Model…</button>' +
          '<button class="wire-btn uns-assign" type="button" aria-label="Save ' + nm + ' onto the selected models">Save</button>' +
        '</span>' +
        '<button class="wire-del uns-del" type="button" title="Remove from pile" aria-label="Remove ' + nm + ' from pile">' + DELICON + '</button>' +
      '</div>';
    }).join("");
    [].forEach.call(list.querySelectorAll(".uns-row"), function (rowEl) {
      var it = items[+rowEl.getAttribute("data-i")];
      var prodSel = rowEl.querySelector(".uns-prod"), mbtn = rowEl.querySelector(".uns-mbtn");
      if (prodSel && mbtn) { // an unresolved orphan row has no picker
        var models = [], selected = []; // models of the chosen product; which ones are checked to sort into
        var repaintMbtn = function () { mbtn.textContent = models.length ? unsMbtnLabel(models, selected) : "Model…"; };
        prodSel.innerHTML = '<option value="">Choose a product…</option>';
        CATALOG.forEach(function (c) { c.families.forEach(function (f) { f.products.forEach(function (p) {
          var o = document.createElement("option"); o.value = f.name + "§" + p.name; o.textContent = f.name + " › " + p.name; prodSel.appendChild(o);
        }); }); });
        prodSel.addEventListener("change", function () {
          var v = prodSel.value; selected.length = 0; unsCloseMenu(false);
          var parts = v ? v.split("§") : null;
          models = parts ? wireModelsOf(parts[0], parts[1]) : [];
          mbtn.disabled = !models.length;
          repaintMbtn();
        });
        // Multi-select: open a checkbox dropdown to pick one, several, or all models to send this file to.
        mbtn.addEventListener("click", function () { if (models.length) unsOpenCheckMenu(mbtn, models, selected, repaintMbtn); });
        var assignBtn = rowEl.querySelector(".uns-assign");
        if (assignBtn) assignBtn.addEventListener("click", function () { unsortedAssign(box, it, prodSel.value, selected.slice(), this); });
      }
      var fileBtn = rowEl.querySelector("button.uns-file");
      if (fileBtn) fileBtn.addEventListener("click", function () { if (it.url) openPreview(it.url, it.name); });
      rowEl.querySelector(".uns-del").addEventListener("click", function () { unsortedDelete(box, it, this); });
    });
    // Keep the tile compact: show the first file, hide the rest behind a "Review all N files" toggle (same
    // pattern as Files to Place). Expanded state persists on the panel across re-renders.
    var unsRows = [].slice.call(list.querySelectorAll(".uns-row"));
    if (unsRows.length > 1 && !box._unsExpanded) {
      unsRows.forEach(function (li, i) { if (i >= 1) li.style.display = "none"; });
      var more = document.createElement("button"); more.className = "wire-btn ftp-showall"; more.type = "button";
      more.textContent = "Review all " + unsRows.length + " files";
      more.addEventListener("click", function () { box._unsExpanded = true; unsRows.forEach(function (li) { li.style.display = ""; }); more.remove(); });
      list.appendChild(more);
    }
    // Rebuilding #unsList dumped keyboard focus to <body>. If the user acted from WITHIN the pile (flag set
    // synchronously in the action, cleared here), put focus on a stable control so a keyboard user keeps their
    // place. Only after a user action — never a background/initial load — so it can never steal focus.
    if (_pileFocusPending) {
      _pileFocusPending = false;
      var refocus = list.querySelector(".uns-row button, .uns-row select");
      if (refocus && refocus.focus) { try { refocus.focus(); } catch (e) {} }
    }
  }
  var _pileFocusPending = false; // set by a pile action (Save/Remove/Hide) when it was invoked with focus inside the pile
  function unsortedDelete(box, it, btn) {
    var hadFocus = box.contains(document.activeElement); // capture NOW (synchronously), before any async reload moves it
    // An orphan has no Unsorted-sheet row — "hide" just tombstones it locally so it drops off this list.
    // (The nightly scan keeps its own list; once the file is linked or renamed it stops being flagged there.)
    if (it.orphan) { orphanDismiss(orphanKey(it)); if (it.url) ftpMark("dismiss", it).catch(function () {}); /* server tombstone so the hide holds for every PM/device (best-effort) */ showToast('<span>Hidden. It’ll stay off once it’s linked or renamed.</span>'); if (hadFocus) _pileFocusPending = true; loadUnsorted(box); return; }
    pfConfirm({
      title: "Remove from the pile?",
      body: ["Remove “" + (it.name || "this file") + "” from the pile?",
             "This only clears it from the pile — the file itself stays in SharePoint and is not deleted."],
      confirmLabel: "Remove", danger: true
    }).then(function (ok) {
    if (!ok) return;
    btn.disabled = true;
    apiFetch("/api/unsorted", { method: "POST", body: { action: "remove", rowIds: [it.id] } })
      .then(function () { showToast('<span>Removed from the pile.</span>'); if (hadFocus) _pileFocusPending = true; loadUnsorted(box); })
      .catch(function (err) { btn.disabled = false; showError('<span>Couldn’t remove (' + escapeHtml(String((err && err.message) || err)) + ').</span>', 7000); });
    });
  }
  function unsortedAssign(box, it, prodVal, targets, btn) {
    var hadFocus = box.contains(document.activeElement); // capture before the confirm/async reload moves focus
    if (!prodVal || !targets || !targets.length) { showToast('<span>Pick a product and at least one model.</span>'); return; }
    var resource = it.resource || "";
    if (!resource) { showToast('<span>This file has no resource type recorded — sort it via the folder tool instead.</span>'); return; }
    var parts = prodVal.split("§"), fam = parts[0], prod = parts[1];
    pfConfirmIf(targets.length > 1, {
      title: "Add to " + targets.length + " models?",
      body: ["Add “" + (it.name || "this file") + "” to " + targets.length + " models of " + prod + " as its " + resource + "?",
             "Existing links on those models are kept — this only adds."],
      confirmLabel: "Add"
    }).then(function (ok) {
    if (!ok) return;
    var btnDone = busyButton(btn, "Sorting…");
    // Serialize with the Files-to-place writer through ONE chain: /api/wire-links REPLACES a resource's link
    // set and every writer rebuilds that set from the client's current links, so an overlapping write on the
    // same model·resource reads a stale pre-image and silently drops a link (the 6→2 renders race). Running
    // the read+merge+write inside ftpSerialize makes each write plan from the state the previous one left.
    ftpSerialize(function () {
      return (window.PF_getToken ? window.PF_getToken() : Promise.reject(new Error("no token"))).then(function (token) {
        var bearer = "Bearer " + token;
        // wireLoadExisting resolves Graph driveItems (Graph token `bearer`); the /api writes use our API token.
        // Read each target's CURRENT links first and merge — see above.
        return lhRunPool(targets, function (m) { return wireLoadExisting(m, resource, bearer).then(function (ex) { return { model: m, ex: ex }; }); }, 5).then(function (reads) {
          var assignments = {}, skipped = [], base = {};
          reads.forEach(function (r) {
            // A transient Graph failure makes a model look empty; writing then would CLEAR its real links.
            // Skip it entirely and report, rather than guessing from incomplete information.
            if (!r || !r.ex || r.ex.error) { if (r && r.model) skipped.push(r.model); return; }
            var urls = [], seen = {};
            // Preserve an auto-updating FOLDER link the model already has — otherwise adding one explicit file
            // would REPLACE the folder (and every render it surfaced) with just this file. (wireFamSave does this.)
            r.ex.folderUrls.forEach(function (u) { urls.push(u); }); // every folder link (F297)
            r.ex.files.forEach(function (f) { var k = wireIdent(f); if (!seen[k] && f.url) { seen[k] = 1; urls.push(f.url); } });
            var mk = wireIdent({ fileId: it.fileId, url: it.url }); if (!seen[mk] && it.url) { seen[mk] = 1; urls.push(it.url); }
            assignments[r.model] = urls.join(" ");
            // The pre-image this list was built from. The shrink guard only COUNTS: [R1,R2] -> [R1,Y] is 2 -> 2,
            // so a render another PM added since this tab loaded was silently dropped (F003). With a base the
            // server refuses instead.
            base[r.model] = linksFor(r.model, resource).map(function (l) { return l.url; });
          });
          if (!Object.keys(assignments).length) throw new Error("couldn’t read the current links for the selected model" + (targets.length > 1 ? "s" : ""));
          return postWire(fam, prod, resource, assignments, { allowMissed: true, base: base })
            .then(function (res) {
              // Mirror what landed INSIDE the chain, so a back-to-back sort plans from the post-write state
              // instead of racing the background reload (F003).
              res.wrote.forEach(function (m) { setLinks(m + "::" + resource, window.PF.splitCellUrls(assignments[m] || "")); });
              // A 200 is not "written to every target". The server writes only models it found a resource
              // row for and lists them in `models`; dropping the pile row regardless would lose track of a
              // file that still needs sorting into the rest. SOME targets missed is a partial result the
              // caller reports; ALL of them missed is a failure.
              if (res.missed.length === Object.keys(assignments).length) throw new Error("no " + resource + " row in the sheet for " + res.missed.slice(0, 3).join(", "));
              return { wrote: res.wrote, missed: res.missed };
            })
            // A parked file is removed from the Unsorted sheet; an orphan has no sheet row — tombstone it
            // locally so it drops off at once (the nightly scan already won't re-flag a now-linked file).
            // Keep the row when some targets were missed: there's still work to do on this file.
            .then(function (out) {
              // Keep the row while ANY target is unfinished — models the server didn't write (`missed`)
              // OR models whose current links couldn't be read at all (`skipped`, which were excluded from
              // the write above). Clearing it here is what left "try those again" pointing at nothing.
              if (out.missed.length || skipped.length) { out.pileCleared = false; return out; }
              if (it.orphan) { orphanDismiss(orphanKey(it)); out.pileCleared = true; return out; }
              return apiFetch("/api/unsorted", { method: "POST", body: { action: "remove", rowIds: [it.id] } })
                .then(function () { out.pileCleared = true; return out; }, function () { out.pileCleared = false; return out; });
            })
            .then(function (out) { return { done: out.wrote, skipped: skipped, missed: out.missed, pileCleared: out.pileCleared }; });
        });
      });
    }).then(function (res) {
      var n = res.done.length;
      toastSaved('Sorted “' + escapeHtml(it.name || "file") + '” into ' + escapeHtml(prod) + ' · ' +
        (n > 1 ? n + ' models' : escapeHtml(res.done[0])) + '.' +
        (res.pileCleared === false ? ' It’s still in the pile.' : ''));
      if (res.missed && res.missed.length) showError('<span>' + res.missed.length + ' model' + (res.missed.length === 1 ? ' has' : 's have') + ' no ' + escapeHtml(resource) + ' row in the sheet (' + escapeHtml(res.missed.slice(0, 3).join(", ")) + ') — the file stays in the pile so you can finish those.</span>', 9000);
      if (res.skipped.length) showError('<span>Skipped ' + res.skipped.length + ' model' + (res.skipped.length === 1 ? '' : 's') + ' whose current links couldn’t be read (' + escapeHtml(res.skipped.join(", ")) + ') — nothing was overwritten. Try those again.</span>', 9000);
      res.done.forEach(function (m) { clearHealthFor(m + "::" + resource); }); // just populated — drop any pre-fix flag
      if (hadFocus) _pileFocusPending = true; // keyboard user acted from the pile — restore focus after the reload
      loadUnsorted(box);
      loadCatalog({ fresh: true });
    }).catch(function (err) {
      btnDone();
      if (isConflict(err)) { loadCatalog({ fresh: true }); showError('<span>Another PM changed ' + escapeHtml(resource) + ' on one of these models since this page loaded, so nothing was saved. The page now has the current links — click Save again.</span>', 12000); return; }
      showError('<span>Couldn’t sort (' + escapeHtml(String((err && err.message) || err)) + ').</span>', 7000);
    });
    });
  }

  /* ============================================================
     PM · COVERAGE EDITOR & LINK HEALTH
     ============================================================ */
  // Live per-model link status for the coverage matrix (recomputed so optimistic edits show at once).
  // Usable-only: a cell counts as "on" (and toward the count) only if its link resolves to files — a link the
  // last health run found EMPTY or BROKEN is treated as not-there, matching the Find view. `broken` flags a
  // wired-but-broken cell so the matrix can still surface the problem in red instead of silently hiding it.
  function covLinked(name) { var on = {}, num = {}, broken = {}, review = {}, na = {}, c = 0, applicable = 0; RESOURCE_TYPES.forEach(function (r) { if (naFor(name, r.name)) { na[r.name] = true; return; } applicable++; var u = usableLinked(name, r.name); on[r.name] = u; num[r.name] = linksFor(name, r.name).length; broken[r.name] = healthStatusFor(name, r.name) === "broken"; review[r.name] = u && !broken[r.name] && resIsStale(name, r.name); if (u) c++; }); return { on: on, num: num, broken: broken, review: review, na: na, count: c, applicable: applicable }; }
  function covLinkedCount(name) { return covLinked(name).count; }
  // expectRemoved = how many existing links this save intends to delete. The server refuses (409) a write
  // that would destroy more than that, so a payload built from a stale page can't quietly wipe links the
  // nightly run or another PM added. Callers that only ADD leave it at 0.
  // `base` (optional) = { model: [urls the caller based this edit on] }. Sent so the server can compare-
  // and-swap: it refuses (409 code:"conflict") if another PM changed one of these resources since this page
  // read it — catching a same-count swap the expectRemoved shrink guard can't. Absent -> legacy behavior.
  // opts: { expectRemoved, expectRemovedByModel, base, allowMissed }. Resolves with the server's reply plus
  // `wrote` / `missed` (models it did / did not find a resource row for); throws on a rejected link, and
  // on any missed model unless opts.allowMissed (the pile sorter reports a partial write itself).
  // A compare-and-swap refusal: another PM changed the resource after this tab's copy was taken. Nothing
  // was saved; the caller refreshes the copy so the very next click plans from the sheet as it is now.
  function isConflict(err) { return !!err && err.status === 409 && err.code === "conflict"; }
  function postWire(fam, prod, resource, assignments, opts) {
    opts = opts || {};
    var payload = { family: fam, product: prod, resource: resource, assignments: assignments, expectRemoved: opts.expectRemoved || 0 };
    // Only a map the caller actually filled. An empty one used to be sent by default, and the server read it
    // as "every model may lose 0 links" — overriding expectRemoved (F049/F089).
    if (opts.expectRemovedByModel && Object.keys(opts.expectRemovedByModel).length) payload.expectRemovedByModel = opts.expectRemovedByModel;
    if (opts.base) payload.base = opts.base;
    return apiFetch("/api/wire-links", { method: "POST", body: payload }).then(function (res) {
      noteServerWrite(res.writtenAt); // from now on, a catalog snapshot older than this save is ignored
      // A 200 here is NOT "everything saved". The server writes only the models it found a resource row
      // for and reports them back in `models`; anything else in `assignments` was dropped. It also counts
      // values its URL check threw out in `rejected`. Both were read nowhere once, so a model whose row is
      // missing from the sheet — or a malformed link — came back as a plain success and the caller went
      // on to update LINKS as though it had been written.
      // `rejected` first: a value the sanitizer threw out never reaches the sheet write either, so it is
      // ALSO absent from `models` — reporting that as a missing row would name the wrong cause.
      if (res.rejected) throw new Error(res.rejected + " link" + (res.rejected === 1 ? " wasn’t" : "s weren’t") + " a valid https:// address and " + (res.rejected === 1 ? "wasn’t" : "weren’t") + " saved");
      var wrote = {}; (res.models || []).forEach(function (m) { wrote[m] = 1; });
      var chosen = Object.keys(assignments || {});
      res.wrote = chosen.filter(function (m) { return wrote[m]; });
      res.missed = chosen.filter(function (m) { return !wrote[m]; });
      if (res.missed.length && !opts.allowMissed) {
        // The server DID write the other models — say which, so the caller can mirror them into LINKS
        // instead of leaving this tab behind the sheet (F121).
        var err = new Error("no " + resource + " row in the sheet for " + res.missed.slice(0, 4).join(", ") + (res.missed.length > 4 ? " +" + (res.missed.length - 4) + " more" : ""));
        err.wrote = res.wrote; err.missed = res.missed; throw err;
      }
      return res;
    });
  }
  // "Fix link": jump to Progress, find the model's product section in the coverage matrix, open its inline
  // editor, and scroll it into view — so a PM can repair a broken/empty link where they spotted it (the
  // link-health list), instead of being dropped on the read-only Find view.
  // Progress holds work that lives only in the page: a section editor's staged edits and an Add-files preview
  // the PM has sorted. A full renderCoverage() throws both away without a word, so a refresh that lands while
  // either is open only updates the numbers (F027).
  function progressHasOpenWork() { return !!(content.querySelector(".cov-grp-head.editing") || content.querySelector("#wireResult .wire-bucket")); }
  function repaintProgressSafely() { if (view !== "progress") return; if (progressHasOpenWork()) { refreshCovStats(); return; } renderCoverage(); }
  // Staged link edits nothing has saved: a section editor with changes, or Manage Document with ticks.
  function pmHasStagedEdits() { return !!(document.querySelector(".cov-group.has-staged") || (docPanel && !docPanel.hidden && docHasPending())); }
  window.addEventListener("beforeunload", function (e) { if (!pmHasStagedEdits()) return; e.preventDefault(); e.returnValue = ""; });
  // Leaving the view in-app (Back/Forward, a tab, the logo) re-renders and throws staged edits away just
  // like closing the page does — ask the same question first (F146, F027). Resolves true = go ahead.
  function confirmDiscardStaged() {
    if (!pmHasStagedEdits()) return Promise.resolve(true);
    return pfConfirm({ title: "Discard these changes?", body: ["You have unsaved link changes. Leaving now throws them away — nothing has been saved yet."], confirmLabel: "Discard", danger: true });
  }
  function openFixEditorFor(model, resource, fallbackPath) {
    var target = null;
    coverageStats().groups.forEach(function (g) { if (g.models.some(function (m) { return m.name === model; })) target = g; });
    if (!target) { if (fallbackPath) openTo(fallbackPath, true); return; } // shouldn't happen; degrade to navigation
    // Render only when the matrix isn't already on screen: re-rendering Progress from a To-do Fix wiped any
    // other section's open editor, staged edits and all (F027).
    if (view !== "progress" || !document.querySelector(".cov-group")) { view = "progress"; render(); } // renders the matrix synchronously so the group is in the DOM below
    var grp = null;
    [].forEach.call(document.querySelectorAll(".cov-group"), function (el) { if (el.dataset.prod === target.prod && el.dataset.fam === target.fam) grp = el; });
    if (!grp) return;
    var editBtn = grp.querySelector(".cov-grp-edit"); if (editBtn) editBtn.click(); // opens the inline editor synchronously
    // Land the PM ON the exact resource they clicked Fix for — scroll its model row into view and flash the
    // chip — instead of dropping them at the top of a multi-model editor with nothing marked.
    var chip = resource ? grp.querySelector('.cov-chip[data-m="' + cssEsc(model) + '"][data-r="' + cssEsc(resource) + '"]') : null;
    var scrollTo = (chip && chip.closest(".cov-ed-model")) || grp;
    if (chip) { chip.classList.add("flagfix"); setTimeout(function () { chip.classList.remove("flagfix"); }, 2500); }
    try { scrollTo.scrollIntoView({ behavior: RM.matches ? "auto" : "smooth", block: chip ? "center" : "start" }); } catch (e) { scrollTo.scrollIntoView(); }
  }
  // Normal (read-only) render of one product group in the coverage matrix, with a header Edit button.
  function renderCovGroup(grp, g, models, N) {
    var started = models.filter(function (m) { return covLinkedCount(m.name) > 0; }).length;
    var head = document.createElement("div"); head.className = "cov-grp-head";
    head.innerHTML = '<span class="cov-grp-title" role="heading" aria-level="3">' + escapeHtml(g.prod) + '<small>' + escapeHtml(g.fam) + ' · ' + started + '/' + models.length + ' models started</small></span>' +
      '<button class="cov-grp-edit" type="button" title="Edit this section’s links">' + svg(ICONS.edit) + '<span>Edit</span></button>';
    grp.innerHTML = ""; grp.classList.remove("has-staged"); grp.appendChild(head);
    // The matrix is a real ARIA table (role=table/row/columnheader/rowheader/cell) so a screen reader can
    // navigate it cell-by-cell and hear "<model>, <resource>: linked" — the old whole-row role="button"
    // announced only "button" and hid the grid. Layout is unchanged (roles carry no styling); the model
    // name is a real <button> so keyboard users can open the model, and clicking anywhere on the row still
    // opens it for mouse (the button's activation bubbles to the row handler).
    var tbl = document.createElement("div"); tbl.className = "cov-table"; tbl.setAttribute("role", "table"); tbl.setAttribute("aria-label", g.prod + " — resource coverage by model");
    tbl.style.setProperty("--cov-cols", RESOURCE_TYPES.length); // grid track count follows the taxonomy (app.css uses repeat(var(--cov-cols),…) instead of a hardcoded 9)
    var ch = document.createElement("div"); ch.className = "cov-colhead"; ch.setAttribute("role", "row");
    ch.innerHTML = '<span role="columnheader">Model</span>' + RESOURCE_TYPES.map(function (r) { return '<span class="cov-colicon" role="columnheader" title="' + escapeAttr(r.name) + '" aria-label="' + escapeAttr(r.name) + '">' + svg(ICONS[r.icon]) + '</span>'; }).join("") + '<span role="columnheader">Done</span>';
    tbl.appendChild(ch);
    models.forEach(function (m) {
      var lk = covLinked(m.name), full = lk.applicable > 0 && lk.count === lk.applicable;
      var row = document.createElement("div"); row.className = "cov-row"; row.setAttribute("role", "row");
      row.innerHTML = '<span class="cov-mname" role="rowheader"><button type="button" class="cov-mopen" aria-label="Open ' + escapeAttr(m.name) + '">' + escapeHtml(m.name) + '</button></span>' +
        RESOURCE_TYPES.map(function (r) { if (lk.na[r.name]) return '<span class="cov-cell na" role="cell" aria-label="' + escapeAttr(r.name + ": not applicable") + '" title="' + escapeAttr(r.name + ": not applicable") + '"></span>'; var on = lk.on[r.name], n = lk.num[r.name], bad = lk.broken[r.name], rev = lk.review[r.name]; var t = on ? ((n > 1 ? ": " + n + " links" : ": linked") + (rev ? " · needs review (" + (resUpdatedLabel(m.name, r.name).toLowerCase()) + ")" : "")) : (bad ? ": Link issue — can’t be opened" + healthWhen() : ": not linked"); return '<span class="cov-cell' + (on ? " on" + (rev ? " review" : "") : bad ? " broken" : "") + '" role="cell" aria-label="' + escapeAttr(r.name + t) + '" title="' + escapeAttr(r.name + t) + '">' + (on ? (n > 1 ? n : "✓") : bad ? "!" : "–") + '</span>'; }).join("") +
        '<span class="cov-cnt' + (full ? " full" : "") + '" role="cell">' + lk.count + '/' + lk.applicable + '</span>';
      row.addEventListener("click", function () { openTo(m.path, true); });
      tbl.appendChild(row);
    });
    grp.appendChild(tbl);
    head.querySelector(".cov-grp-edit").addEventListener("click", function () { paintCovGroupEditor(grp, g, models, N); });
  }
  // Staged link editor for one product section: delete (✕) + add (+ a link), commit on Save
  // (confirms before deleting), Cancel discards. Optimistic on success; sheet propagates ~1 min.
  function paintCovGroupEditor(grp, g, models, N) {
    var dels = [], adds = []; // dels: {model,resource}; adds: {model,resource,url}
    var naAdd = [], naUndo = []; // naAdd: mark {model,resource} not-applicable; naUndo: un-mark an existing N/A
    // "+ Add a link" stages only after an async Graph normalize. Save read `adds` while one was still in
    // flight and closed the editor without it (F045); a Cancelled add row resurrected it. Count them, and
    // let a Cancel orphan the pending result.
    var pendingAdds = 0, addSeq = 0;
    function isAdd(m, r) { return adds.some(function (a) { return a.model === m && a.resource === r; }); }
    // Staged removals are per FILE (url), not per resource type. A resource can hold several files, and
    // keying a removal by resource alone is what made deleting one file wipe every file under it.
    function isDel(m, r, u) { return dels.some(function (d) { return d.model === m && d.resource === r && d.url === u; }); }
    function isNaAdd(m, r) { return naAdd.some(function (x) { return x.model === m && x.resource === r; }); }
    function isNaUndo(m, r) { return naUndo.some(function (x) { return x.model === m && x.resource === r; }); }
    grp.innerHTML = "";
    var head = document.createElement("div"); head.className = "cov-grp-head editing";
    head.innerHTML = '<span class="cov-grp-title" role="heading" aria-level="3">' + escapeHtml(g.prod) + '<small>Editing links</small></span>' +
      '<span class="cov-ed-actions"><button class="wire-btn cov-ed-cancel" type="button">Cancel</button><button class="wire-btn primary cov-ed-save" type="button">Save</button></span>';
    grp.appendChild(head);
    var body = document.createElement("div"); body.className = "cov-ed-body"; grp.appendChild(body);

    // Resource types this model can still take a link for. Mirrors the "Mark N/A" picker's filter: a
    // resource already marked N/A isn't offered (un-mark it first — that stages an naUndo, and then it
    // appears), and neither is one staged for N/A in this very save, whose "N/A" write would otherwise
    // land on top of the link just added.
    function availResources(model) {
      return RESOURCE_TYPES.filter(function (r) {
        if (linksFor(model, r.name).length || isAdd(model, r.name)) return false;
        if (isNaAdd(model, r.name)) return false;
        return !(naFor(model, r.name) && !isNaUndo(model, r.name));
      }).map(function (r) { return r.name; });
    }
    function paint() {
      // Marks the section as holding unsaved edits, for the leave-page prompt (pmHasStagedEdits).
      grp.classList.toggle("has-staged", !!(dels.length || adds.length || naAdd.length || naUndo.length));
      body.innerHTML = models.map(function (m) {
        var chips = "";
        RESOURCE_TYPES.forEach(function (r) {
          var arr = linksFor(m.name, r.name);
          if (!arr.length) return;
          // ONE chip per file, labelled with the file's own name, so a resource holding several files shows
          // several chips — you can see which is which, preview it, and remove just that one.
          arr.forEach(function (l) {
            var removed = isDel(m.name, r.name, l.url);
            var label = l.label || r.name;
            chips += '<span class="cov-chip' + (removed ? " removed" : "") + '" data-m="' + escapeAttr(m.name) + '" data-r="' + escapeAttr(r.name) + '" data-u="' + escapeAttr(l.url) + '">' +
              '<span class="cov-chip-name cov-chip-prev" role="button" tabindex="0" title="Preview ' + escapeAttr(label) + '">' + escapeHtml(label) + '</span>' +
              '<span class="cov-chip-tag">' + escapeHtml(COV_SHORT[r.name] || r.name) + '</span>' +
              ((!removed && !looksLikeFolder(l.url)) ? '<button class="cov-chip-act manage" type="button" title="Manage this document across products" aria-label="Manage where ' + escapeAttr(label) + ' is used">' + svg('<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/>') + '</button>' : '') +
              (removed ? '<button class="cov-chip-act undo" type="button" aria-label="Undo removing ' + escapeAttr(label) + '">undo</button>' : '<button class="cov-chip-act del" type="button" aria-label="Remove ' + escapeAttr(label) + ' from ' + escapeAttr(r.name) + '">✕</button>') + '</span>';
          });
        });
        adds.filter(function (a) { return a.model === m.name; }).forEach(function (a) {
          chips += '<span class="cov-chip new" data-m="' + escapeAttr(m.name) + '" data-r="' + escapeAttr(a.resource) + '"><span class="cov-chip-name">' + escapeHtml(a.resource) + '</span><span class="cov-chip-tag">new</span><button class="cov-chip-act unadd" type="button" aria-label="' + escapeAttr("Undo adding " + a.resource) + '">✕</button></span>';
        });
        // N/A chips: a resource already marked not-applicable (unless staged to undo), or newly staged N/A.
        RESOURCE_TYPES.forEach(function (r) {
          if (!((naFor(m.name, r.name) && !isNaUndo(m.name, r.name)) || isNaAdd(m.name, r.name))) return;
          chips += '<span class="cov-chip na" data-m="' + escapeAttr(m.name) + '" data-r="' + escapeAttr(r.name) + '"><span class="cov-chip-name">' + escapeHtml(r.name) + '</span><span class="cov-chip-tag">N/A</span><button class="cov-chip-act naundo" type="button" aria-label="Not applicable — undo">undo</button></span>';
        });
        if (!chips) chips = '<span class="cov-ed-none">No links yet</span>';
        return '<div class="cov-ed-model"><div class="cov-ed-mname">' + escapeHtml(m.name) + '</div><div class="cov-ed-links">' + chips + '</div>' +
          '<div class="cov-ed-addwrap" data-m="' + escapeAttr(m.name) + '"><button class="cov-ed-add" type="button">+ Add a link</button><button class="cov-ed-add cov-ed-na" type="button" title="Mark a resource that doesn’t apply to this model, so it isn’t counted as missing">Mark N/A</button></div></div>';
      }).join("");
      wire();
    }
    function wire() {
      [].forEach.call(body.querySelectorAll(".cov-chip-act.manage"), function (b) {
        b.addEventListener("click", function () { var c = b.closest(".cov-chip"); openDocPanel(c.getAttribute("data-u"), c.getAttribute("data-r"), { mode: "manage" }); });
      });
      [].forEach.call(body.querySelectorAll(".cov-chip-act.del"), function (b) {
        b.addEventListener("click", function () { var c = b.closest(".cov-chip"); dels.push({ model: c.getAttribute("data-m"), resource: c.getAttribute("data-r"), url: c.getAttribute("data-u") }); paint(); });
      });
      [].forEach.call(body.querySelectorAll(".cov-chip-act.undo"), function (b) {
        b.addEventListener("click", function () { var c = b.closest(".cov-chip"); dels = dels.filter(function (d) { return !(d.model === c.getAttribute("data-m") && d.resource === c.getAttribute("data-r") && d.url === c.getAttribute("data-u")); }); paint(); });
      });
      // Preview the actual file before removing it — the whole point of showing one chip per file.
      [].forEach.call(body.querySelectorAll(".cov-chip-prev"), function (nameEl) {
        var open = function () {
          var c = nameEl.closest(".cov-chip");
          openResource([{ label: nameEl.textContent, url: c.getAttribute("data-u") }], c.getAttribute("data-m") + " · " + c.getAttribute("data-r"), "");
        };
        nameEl.addEventListener("click", open);
        nameEl.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); open(); } });
      });
      [].forEach.call(body.querySelectorAll(".cov-chip-act.unadd"), function (b) {
        b.addEventListener("click", function () { var c = b.closest(".cov-chip"); adds = adds.filter(function (a) { return !(a.model === c.getAttribute("data-m") && a.resource === c.getAttribute("data-r")); }); paint(); });
      });
      [].forEach.call(body.querySelectorAll(".cov-chip-act.naundo"), function (b) {
        b.addEventListener("click", function () {
          var c = b.closest(".cov-chip"), m = c.getAttribute("data-m"), r = c.getAttribute("data-r");
          if (isNaAdd(m, r)) naAdd = naAdd.filter(function (x) { return !(x.model === m && x.resource === r); }); // cancel a staged mark
          else naUndo.push({ model: m, resource: r }); // un-mark an already-saved N/A (becomes "not linked")
          paint();
        });
      });
      [].forEach.call(body.querySelectorAll(".cov-ed-na"), function (b) {
        b.addEventListener("click", function () {
          var wrap = b.closest(".cov-ed-addwrap"), model = wrap.getAttribute("data-m");
          var avail = RESOURCE_TYPES.filter(function (r) { return !linksFor(model, r.name).length && !isAdd(model, r.name) && !isNaAdd(model, r.name) && !(naFor(model, r.name) && !isNaUndo(model, r.name)); }).map(function (r) { return r.name; });
          if (!avail.length) { showToast('<span>Nothing to mark — every resource here is already linked, work in progress, or N/A.</span>'); return; }
          wrap.innerHTML = '<select class="wire-sel cov-na-res" aria-label="Resource type to mark not applicable">' + avail.map(function (n) { return '<option value="' + escapeAttr(n) + '">' + escapeHtml(n) + '</option>'; }).join("") + '</select>' +
            '<button class="wire-btn cov-na-ok" type="button">Mark N/A</button><button class="wire-btn cov-na-cancel" type="button">Cancel</button>';
          wrap.querySelector(".cov-na-cancel").addEventListener("click", function () { paint(); });
          wrap.querySelector(".cov-na-ok").addEventListener("click", function () { naAdd.push({ model: model, resource: wrap.querySelector(".cov-na-res").value }); paint(); });
        });
      });
      [].forEach.call(body.querySelectorAll(".cov-ed-add:not(.cov-ed-na)"), function (b) {
        b.addEventListener("click", function () {
          var wrap = b.closest(".cov-ed-addwrap"), model = wrap.getAttribute("data-m"), avail = availResources(model);
          if (!avail.length) { showToast('<span>All ' + RESOURCE_TYPES.length + ' resource types already have a link (or a pending one) on this model.</span>'); return; }
          wrap.innerHTML = '<select class="wire-sel cov-add-res" aria-label="Resource type">' + avail.map(function (n) { return '<option value="' + escapeAttr(n) + '">' + escapeHtml(n) + '</option>'; }).join("") + '</select>' +
            '<input class="cov-add-url" type="url" placeholder="Paste a SharePoint link" aria-label="SharePoint link" />' +
            '<button class="wire-btn cov-add-ok" type="button">Add</button><button class="wire-btn cov-add-cancel" type="button">Cancel</button>';
          var urlEl = wrap.querySelector(".cov-add-url"); urlEl.focus();
          wrap.querySelector(".cov-add-cancel").addEventListener("click", function () { addSeq++; paint(); });
          wrap.querySelector(".cov-add-ok").addEventListener("click", function () {
            var okB = wrap.querySelector(".cov-add-ok");
            var res = wrap.querySelector(".cov-add-res").value, url = urlEl.value.trim();
            if (!/^https?:\/\//i.test(url)) { showToast('<span>Paste a full https:// SharePoint link.</span>'); return; }
            okB.disabled = true; okB.textContent = "Adding…";
            // Normalize to the canonical webUrl+id before staging (identity-stable storage); never blocks.
            var mySeq = addSeq; pendingAdds++;
            normalizeLink(url).then(function (n) {
              pendingAdds--;
              // Cancelled meanwhile — the add row, or the whole editor (Cancel, Save, a re-render): a late
              // result must not stage into an editor that is no longer on screen (F045).
              if (mySeq !== addSeq || !head.isConnected) return;
              // Use the filename as the chip's label until the next catalog load re-derives it.
              var nm = "";
              try { nm = decodeURIComponent((String(n.url).split(/[?#]/)[0].split("/").pop() || "")).replace(/\.[a-z0-9]+$/i, ""); } catch (e) {}
              adds.push({ model: model, resource: res, url: n.url, name: nm || null });
              paint();
            });
          });
          urlEl.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); wrap.querySelector(".cov-add-ok").click(); } });
        });
      });
    }
    head.querySelector(".cov-ed-cancel").addEventListener("click", function () { addSeq++; renderCovGroup(grp, g, models, N); focusGroupEdit(); });
    head.querySelector(".cov-ed-save").addEventListener("click", function () {
      if (pendingAdds) { showToast('<span>Still adding that link — Save again in a moment.</span>'); return; }
      if (!dels.length && !adds.length && !naAdd.length && !naUndo.length) { renderCovGroup(grp, g, models, N); focusGroupEdit(); return; }
      pfConfirmIf(dels.length > 0, {
        title: "Remove file links?",
        body: ["Remove " + dels.length + " file link" + (dels.length > 1 ? "s" : "") + "?",
               "The files stay in SharePoint — only the links are cleared."],
        confirmLabel: "Remove", danger: true
      }).then(function (ok) {
      if (!ok) return;
      var byRes = {}, finalVal = {}; // finalVal[model::resource] = the urls that should remain
      // Removing a file means REWRITING that resource with the files that survive — not blanking it.
      // (Writing "" here is what wiped every file under a resource when only one chip was removed.)
      var delMap = {};
      dels.forEach(function (d) { var byModel = delMap[d.resource] = delMap[d.resource] || {}; (byModel[d.model] = byModel[d.model] || []).push(d.url); });
      Object.keys(delMap).forEach(function (res) {
        Object.keys(delMap[res]).forEach(function (model) {
          var gone = delMap[res][model];
          // Keep the {label,url} objects for the optimistic update; the WRITE still sends bare urls.
          var remaining = linksFor(model, res).filter(function (l) { return gone.indexOf(l.url) === -1; });
          (byRes[res] = byRes[res] || {})[model] = remaining.map(function (l) { return l.url; }).join(" ");
          finalVal[model + "::" + res] = remaining;
        });
      });
      adds.forEach(function (a) {
        var cur = (byRes[a.resource] = byRes[a.resource] || {})[a.model];
        byRes[a.resource][a.model] = cur ? (cur + " " + a.url).trim() : a.url;
        var k = a.model + "::" + a.resource;
        // Carry a label for the new file too, so its chip reads as a filename, not the resource name.
        finalVal[k] = (finalVal[k] || linksFor(a.model, a.resource)).concat([{ label: a.name || null, url: a.url }]);
      });
      naAdd.forEach(function (x) { (byRes[x.resource] = byRes[x.resource] || {})[x.model] = "N/A"; });   // mark not-applicable
      // Clear N/A -> "not linked". Deferential on purpose: un-marking N/A and adding a link to that same
      // resource is one natural edit ("this DOES apply, here it is"), and this line runs after `adds`, so
      // writing "" unconditionally threw the link away and saved an empty cell instead.
      naUndo.forEach(function (x) {
        var byModel = (byRes[x.resource] = byRes[x.resource] || {});
        if (byModel[x.model] == null) byModel[x.model] = "";
      });
      var saveDone = busyButton(head.querySelector(".cov-ed-save"), "Saving…");
      // Per resource, declare exactly what this save removes: each staged file delete, plus every link
      // wiped by a "Mark N/A". Anything beyond that is the catalog having moved under us — the server
      // refuses it rather than deleting links this editor never showed.
      function removalsFor(res) {
        var n = dels.filter(function (d) { return d.resource === res; }).length;
        naAdd.filter(function (x) { return x.resource === res; }).forEach(function (x) { n += linksFor(x.model, res).length; });
        return n;
      }
      // …and the SAME removals broken out per model, so the server's per-model shrink guard engages (a
      // pooled-only count let an unexpected shrink on model B hide behind the allowance meant for model A).
      function removalsByModelFor(res) {
        var m = {}, dm = delMap[res] || {};
        Object.keys(dm).forEach(function (model) { m[model] = (m[model] || 0) + dm[model].length; });
        naAdd.filter(function (x) { return x.resource === res; }).forEach(function (x) { m[x.model] = (m[x.model] || 0) + linksFor(x.model, res).length; });
        return m;
      }
      // ONE resource at a time (like wireApprove), each its own all-or-nothing write. So each can carry a
      // compare-and-swap base — the links THIS editor showed (LINKS is only updated as each resource commits,
      // so linksFor still returns the pre-edit set for the ones not yet written) — and a failure knows exactly
      // which resources landed. The old parallel Promise.all left no record: a partial failure read as total,
      // LINKS stayed behind the sheet, and the retry's base falsely conflicted (F043/F159).
      var resKeys = Object.keys(byRes), done = {};
      function baseFor(res) {
        var b = {};
        Object.keys(byRes[res] || {}).forEach(function (model) { b[model] = linksFor(model, res).map(function (l) { return l.url; }); });
        return b;
      }
      // Mirror what one resource's write committed: what's LEFT, not "gone".
      function applyRes(res) {
        var suffix = "::" + res;
        Object.keys(finalVal).forEach(function (k) { if (k.slice(-suffix.length) !== suffix) return; setLinks(k, finalVal[k]); clearHealthFor(k); });
        naAdd.forEach(function (x) { if (x.resource !== res) return; var k = x.model + suffix; setLinks(k, []); NA[k] = 1; clearHealthFor(k); });
        naUndo.forEach(function (x) { if (x.resource === res) delete NA[x.model + suffix]; });
      }
      resKeys.reduce(function (chain, res) {
        return chain.then(function () {
          return postWire(g.fam, g.prod, res, byRes[res], { expectRemoved: removalsFor(res), expectRemovedByModel: removalsByModelFor(res), base: baseFor(res) })
            .then(function () { applyRes(res); done[res] = 1; });
        });
      }, Promise.resolve()).then(function () {
        toastSaved('Saved changes for ' + escapeHtml(g.prod) + '.');
        renderCovGroup(grp, g, models, N);
        focusGroupEdit();
        refreshCovStats(); // keep the top summary %/bars in sync with the optimistic LINKS change
      }).catch(function (err) {
        saveDone();
        var landed = resKeys.filter(function (r) { return done[r]; });
        // Keep only the edits that did NOT land, so Save again re-sends exactly those.
        var keep = function (x) { return !done[x.resource]; };
        dels = dels.filter(keep); adds = adds.filter(keep); naAdd = naAdd.filter(keep); naUndo = naUndo.filter(keep);
        var lead = landed.length ? 'Saved ' + escapeHtml(landed.join(", ")) + ', but couldn’t save the rest' : 'Couldn’t save';
        if (isConflict(err)) showError('<span>' + lead + ' — another PM changed it since you opened the editor. Your changes are still staged over the current links: check them and Save again.</span>', 15000);
        else if (err && err.status === 409) showError('<span>' + (landed.length ? lead + '. ' : '') + escapeHtml(String(err.message)) + '</span>', 15000);
        else showError('<span>' + lead + ' (' + escapeHtml(String((err && err.message) || err)) + '). Your changes are still staged — Save again to finish.</span>', 9000);
        refreshCovStats();
        // Rebuild from the sheet as it is now, keeping the staged edits: the next Save's base is current.
        loadCatalog({ fresh: true }).then(function () { refreshCovStats(); paint(); });
      });
      });
    });
    // Keyboard a11y: closing the editor rebuilds the group and would otherwise drop focus to <body>.
    // Return focus to this section's Edit button. (Declaration hoists, so the Save/Cancel handlers above
    // can call it.)
    function focusGroupEdit() { var eb = grp.querySelector(".cov-grp-edit"); if (eb) eb.focus(); }
    paint();
    // Opening the editor: put focus on its Cancel button so a keyboard user lands inside the panel.
    var openCancel = head.querySelector(".cov-ed-cancel"); if (openCancel) openCancel.focus();
  }

  // ---- Progress tool: Link health — resolve every wired link/folder against SharePoint and
  // flag anything BROKEN (moved/deleted/no access) or EMPTY (a shared folder with no files for
  // that model — the "shows Open but opens to nothing" case). Uses the signed-in PM's Graph token.
  function healthStatusFor(model, rname) {
    var key = model + "::" + rname;
    if (healthCleared[key]) return null; // just re-wired this session — don't show a pre-fix flag
    // Prefer the newer COMPLETE snapshot. A PM's manual "Check all links" (healthMap/healthTs) and the
    // daily server-side cron (serverHealth/serverHealthTs) each check the same set; whichever ran more
    // recently is authoritative, so a stale local run can't override fresh server data or vice-versa.
    // Salespeople never run a manual check (healthTs 0) → they get the server flags.
    if (serverHealthTs >= healthTs) return serverHealth[key] || null;
    return healthMap[key] || null;
  }
  // The run healthStatusFor reads from — the newer of the local manual check and the server cron.
  function lastHealthTs() { return serverHealthTs >= healthTs ? serverHealthTs : healthTs;
  }
  function saveHealth(map, ts, okN, problems) {
    healthMap = map; healthTs = ts;
    healthCleared = {}; // a fresh full run is authoritative — drop this session's per-fix suppressions
    // okN/problems are still persisted for forward-compat with any future reader, but nothing reads them now.
    try { localStorage.setItem(HEALTH_KEY, JSON.stringify({ map: map, ts: ts, okN: okN, problems: problems })); } catch (e) {}
  }
  // Clear a link's health flag after the PM re-wires it: drop it from the local run map AND the
  // server map, and remember it in healthCleared so the post-save catalog reload (which re-seeds
  // serverHealth from the pre-fix cron data) doesn't resurrect the stale "Link issue".
  function clearHealthFor(key) { delete healthMap[key]; delete serverHealth[key]; healthCleared[key] = 1; }
  function healthAgo() {
    // Use whichever run healthStatusFor is actually reading from (a local manual run, or the daily
    // server cron) so the "last checked" note is right for everyone — incl. salespeople with no local run.
    var ts = lastHealthTs();
    if (!ts) return "";
    var mins = Math.round((Date.now() - ts) / 60000);
    if (mins < 1) return "just now"; if (mins < 60) return mins + " min ago";
    var h = Math.round(mins / 60); if (h < 24) return h + " hr ago"; return Math.round(h / 24) + " d ago";
  }
  // " (2 hr ago)" — or nothing at all when no run date is known. Callers used to inline the parentheses,
  // which read as "last checked () —" whenever the server health map arrived without a date.
  function healthWhen() { var a = healthAgo(); return a ? " (" + a + ")" : ""; }
  function collectWiredResources() {
    var out = [];
    CATALOG.forEach(function (c) { c.families.forEach(function (f) { f.products.forEach(function (p) { (p.models || []).forEach(function (m) {
      RESOURCE_TYPES.forEach(function (r) {
        var arr = linksFor(m, r.name);
        if (arr && arr.length) out.push({ cat: c.name, fam: f.name, prod: p.name, model: m, resource: r.name, arr: arr, path: [c.name, f.name, p.name, m] });
      });
    }); }); }); });
    return out;
  }
  // lhReason / lhTransient live in pf-shared.js (loaded via <script src>) so the browser and the server
  // cron classify link-health errors identically. Thin delegators keep every call site unchanged.
  function lhReason(e) { return window.PF.lhReason(e); }
  // Resolve one model+resource. Thin browser adapter over the SHARED classifier (window.PF.lhCheckResource,
  // from pf-shared.js) — injects the delegated-token client fetch layer + the folderRefCount sharing test,
  // so the manual "Check all links" and the daily cron classify links by ONE implementation (no drift).
  // cache memoizes folder listings + file lookups by URL so a folder shared across models is fetched once.
  function checkOneResource(it, bearer, cache) {
    return window.PF.lhCheckResource(it, {
      isShared: function (url) { return folderRefCount(url) > 1; },
      listFolder: function (url) { return fetchFolderChildren(url, bearer); },
      resolveOne: function (url) { return wireResolveItem(url, bearer); },
      canonUrl: window.PF.canonFolderUrl,
      cache: cache
    });
  }
  // Thin re-export of the shared concurrency pool; the 4th arg here is a progress callback (the live counter).
  function lhRunPool(items, worker, concurrency, onProgress) {
    return window.PF.lhRunPool(items, worker, concurrency, { onProgress: onProgress });
  }

  /* ============================================================
     PM · FILES TO PLACE
     ============================================================ */
  // ---- "Files to place": scan the folders your resources already come from, smart-match any file that
  // isn't linked yet to a model, and wire it on your confirm. Client-side — uses your sign-in to list
  // folders (same as "Check all links"), the shared matcher (window.PF.guessPlacement) to read names the
  // way a person would, and the existing /api/wire-links to save. Nothing is written without a click. ----
  function ftpDismissed() { try { var a = JSON.parse(localStorage.getItem("pf-place-done") || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
  function ftpDismiss(url) { var a = ftpDismissed(); if (a.indexOf(url) === -1) a.unshift(url); try { localStorage.setItem("pf-place-done", JSON.stringify(a.slice(0, 500))); } catch (e) {} }
  // ONE reader for /api/placements. Progress used to fetch it twice per render (the orphan pile and the
  // Files-to-place scan each pulled the same ~39 KB, 300-1200 ms each). Memoised for 30 s and invalidated
  // by every write below, so a Keep/Remove is visible on the very next read. Rejects on a non-2xx so each
  // caller keeps its own best-effort fallback; a rejection is never cached.
  var placementsStore = window.PF.memoPromise(function () { return apiFetch("/api/placements"); }, 30000);
  // Tell the server a human ruled on this file. The nightly run reads these verdicts, so a dismissal that
  // only lived in one browser's localStorage would be undone by tomorrow's cron — this is what makes
  // "not a fit" and "remove" stick for everyone. Best-effort: a failure here never blocks the UI action.
  function ftpMark(action, x, model) {
    if (!window.PF_getApiToken) return Promise.reject(new Error("Please sign in again."));
    placementsStore.invalidate(); // whatever the outcome, the next read must come from the server
    // This is the ONLY thing that makes a Keep / Remove / "not a fit" durable. If it fails and we
    // swallow it, the row disappears, the PM believes it saved, and tonight's run re-places exactly
    // the files they rejected. So a failure must surface (apiFetch rejects) and the row must stay put.
    return apiFetch("/api/placements", { method: "POST", body: { action: action, url: x.url, file: x.file, resource: x.resource, folderUrl: x.folderUrl || "", model: model || x.model || "" } });
  }
  // Every catalog write this panel makes goes through ONE promise chain.
  //
  // /api/wire-links has REPLACE semantics, and each caller rebuilds the resource's full url list from the
  // client's LINKS. So two writes to the same Model·Resource that OVERLAP both read the same pre-image —
  // the second one, reading a LINKS the first hasn't updated yet, quietly drops the first file. Two clicks
  // inside one round-trip is all it takes, which is not hypothetical: the bulk version of exactly this
  // shrank a model's Renders from 6 links to 2 in production while reporting success.
  //
  // The shrink guard does NOT catch this one — swapping one url for another keeps the count identical, so
  // nothing looks like a removal. Serializing is the actual fix: each write plans from the state the
  // previous write left behind. A failure must not wedge the queue, so the chain always continues.
  var ftpChain = Promise.resolve();
  function ftpSerialize(fn) {
    var run = ftpChain.then(function () { return fn(); }, function () { return fn(); });
    ftpChain = run.then(function () {}, function () {});
    return run;
  }
  // Reject a suggestion (or undo a still-live placement from the pre-propose-only cron): if the file is
  // actually wired, drop just this one from the model's resource (leaving other links intact); either way
  // tombstone the guess so it isn't re-suggested tonight.
  function ftpUnplace(x) {
    var path = findModelPath(x.model);
    if (!path) return Promise.reject(new Error("Couldn’t find " + x.model));
    return ftpSerialize(function () {
      // Read INSIDE the chain — outside it, this is the stale pre-image described above.
      var cur = linksFor(x.model, x.resource).map(function (l) { return l.url; });
      var isWired = cur.some(function (u) { return window.PF.canonFileUrl(u) === window.PF.canonFileUrl(x.url); });
      // Propose-only: a suggestion is NOT in the catalog until someone Keeps it, so removing one is purely a
      // tombstone — there's nothing to unwire. Issuing the wire-links removal anyway would remove 0 links
      // while telling the shrink guard to expect 1, and the write would (correctly) be rejected. Only unwire
      // when the file really is wired (a leftover row from the pre-propose-only cron, still live).
      if (!isWired) return ftpMark("remove", x);
      // Wired means LIVE: this Remove takes a link off the catalog, not just off the list, and the row
      // reads like the tombstone-only case above. Every other destructive write asks first; so does this.
      return pfConfirm({
        title: "Remove a live link?",
        body: ["\u201c" + x.file + "\u201d is already on the site as " + x.model + " \u00b7 " + x.resource + ". Removing it here takes it off the live catalog too.",
               "It comes back under \u201cNeeds your input\u201d so you can place it somewhere else."],
        confirmLabel: "Remove"
      }).then(function (ok) {
        if (!ok) return { cancelled: true };
        var rest = cur.filter(function (u) { return window.PF.canonFileUrl(u) !== window.PF.canonFileUrl(x.url); });
        var assign = {}; assign[x.model] = rest.join(" ");
        // Declare the removal per model, and the real count (the same file under two URL spellings is 2).
        var n = cur.length - rest.length, by = {}; by[x.model] = n;
        return postWire(path[1], path[2], x.resource, assign, { expectRemoved: n, expectRemovedByModel: by }).then(function () {
          var k = x.model + "::" + x.resource;
          setLinks(k, rest);
          clearHealthFor(k);
          return ftpMark("remove", x);
        });
      });
    });
  }
  function allModelsList() {
    var out = [];
    CATALOG.forEach(function (c) { c.families.forEach(function (f) { f.products.forEach(function (p) { (p.models || []).forEach(function (m) { out.push({ model: m, product: p.name }); }); }); }); });
    return out;
  }
  // Append one file's link to a model's resource (merge with what's already there, never clobber).
  // Deliberately just the one-item case of ftpPlaceAll rather than its own write: a second implementation
  // of "merge then replace" is a second place for the race above to come back, and the planner already
  // dedupes canonically (the same file reached by an Office/"Copy link" url must not be added twice).
  function ftpPlace(item, model) {
    var one = { file: item.file, url: item.url, resource: item.resource, folderUrl: item.folderUrl, model: model || item.model };
    return ftpPlaceAll([one]).then(function (res) {
      if (res.failed.length) throw new Error(res.failed[0]);
      return res;
    });
  }
  // Place MANY files at once. Critically this is NOT a loop over a per-file write: /api/wire-links REPLACES
  // a resource's link set, so placing 6 files into the same Model·Resource with 6 sequential writes makes
  // every write race the ones before it — and because each write rebuilds its list from the client's LINKS
  // (frozen for the life of the tab, and re-seedable from a 60s server cache that can still hold PRE-write
  // data), the later writes silently DROP the earlier files and any links that were already there. That's
  // what made "Place all" look like it did nothing while actually shrinking a resource from 6 renders to 2.
  //
  // Instead: compute the union ONCE, group by product+resource, and send a single merged write per group.
  // Failures are collected and reported — never swallowed, which is what made the damage invisible.
  function ftpPlaceAll(items) {
    return ftpSerialize(function () {
      var failed = [], enriched = [];
      (items || []).forEach(function (x) {
        var path = findModelPath(x.model);
        if (!path) { failed.push(x.file + " — couldn’t find " + x.model); return; }
        enriched.push({ fam: path[1], prod: path[2], resource: x.resource, model: x.model, url: x.url, file: x.file });
      });
      // Planned inside the serialized section, so it reads the LINKS the previous write just updated
      // rather than the snapshot that was current when the button was clicked.
      var writes = window.PF.planPlacementWrites(enriched, function (m, r) { return linksFor(m, r).map(function (l) { return l.url; }); });
      var placed = 0, conflict = false;
      return writes.reduce(function (chain, g) {
        return chain.then(function () {
          // g.base = the pre-image the planner built each list from: a stale copy now 409s instead of
          // silently dropping a link another PM added (F080). Failed groups re-plan from LINKS next time.
          return postWire(g.fam, g.prod, g.resource, g.assignments, { base: g.base }).then(function () {
            placed += g.files.length;
            Object.keys(g.assignments).forEach(function (m) {
              // Re-parse with the SAME splitter the server and the planner use. A plain split(" ") tore a
              // URL containing a raw space into two, and the next placement in this tab planned from that
              // truncated value and committed it to the sheet as a dead link.
              var arr = window.PF.splitCellUrls(g.assignments[m] || ""), k = m + "::" + g.resource;
              setLinks(k, arr);
              clearHealthFor(k);
            });
          }, function (e) {
            // Models the server wrote before postWire threw (another model in the group had no row) DID land —
            // mirror them, or the next Keep plans from the pre-write list and drops this one (F121).
            if (e && e.wrote) e.wrote.forEach(function (m) { var k = m + "::" + g.resource; setLinks(k, window.PF.splitCellUrls(g.assignments[m] || "")); clearHealthFor(k); });
            if (isConflict(e)) conflict = true;
            failed.push(g.prod + " · " + g.resource + " (" + ((e && e.message) || e) + ")");
          });
        });
      }, Promise.resolve()).then(function () {
        if (conflict) loadCatalog({ fresh: true }); // so Keep again plans from the sheet as it is now
        return { placed: placed, failed: failed };
      });
    });
  }
  // The red count badge next to "Files to place": how many suggestions await review. Hidden at 0 so a
  // clean inbox shows no symbol; visible (red) when there's work, so it's obvious to go check.
  function setFtpBadge(box, n) {
    var b = box && box.querySelector(".ftp-badge"); if (!b) return;
    if (n > 0) { b.textContent = String(n); b.hidden = false; b.setAttribute("aria-label", n + " file" + (n === 1 ? "" : "s") + " to review"); }
    else { b.textContent = ""; b.hidden = true; b.removeAttribute("aria-label"); }
    if (box._onCount) box._onCount(n); // report to the To-do container (hide the box + the whole inbox when empty)
  }
  function filesToPlacePanel() {
    var box = document.createElement("details"); box.className = "wire files-to-place";
    box.innerHTML =
      '<summary class="wire-head">' + svg(ICONS.clock) + '<span>New on SharePoint</span><span class="ftp-badge" hidden></span></summary>' +
      '<div class="ui-body"><p class="wire-lead">Finds files in your resource folders that aren’t linked to a product yet. <b>Every night this runs on its own</b> and <b>suggests</b> the confident matches — nothing goes on the live site until you click <b>Keep</b>. Anything it wasn’t sure about asks you to pick.</p>' +
      '<div class="ftp-bar"><button class="wire-btn ftp-rescan" type="button">Re-scan now</button>' +
      '<span class="ftp-bar-note">Runs the nightly scan on demand — it shows you what it found before anything is added.</span></div>' +
      '<div id="ftpResults"><div class="wire-msg">Scanning…</div></div></div>';
    // Open by default and scan on render so the badge shows the count without a click — the panel is the
    // one thing on Progress with work waiting that nobody asked for, so it should announce itself.
    box.open = true; box._loaded = true;
    box.addEventListener("toggle", function () { if (box.open && !box._loaded) { box._loaded = true; loadFilesToPlace(box); } });
    box.querySelector(".ftp-rescan").addEventListener("click", function () { ftpRescan(box, this); });
    loadFilesToPlace(box);
    return box;
  }
  // Manual trigger for the nightly run. Always does a DRY pass first and shows the plan, so a run you
  // asked for out-of-band still can't write to the catalog before you've seen what it intends to do.
  function ftpRescan(box, btn) {
    var host = box.querySelector("#ftpResults");
    if (!window.PF_getApiToken) { showError('<span>Please sign in again.</span>', 5000); return; }
    var reset = busyButton(btn, "Scanning…");
    var call = function (dry) {
      // The non-dry run adds suggestion rows, so it is NOT a safe method: send it as a POST. (A GET
      // that mutates can be triggered by anything that prefetches or retries a URL.) The dry run is
      // genuinely read-only and stays a GET. The scan has a 60s server budget — wait it out.
      return apiFetch("/api/place-cron" + (dry ? "?dry=1&cb=" : "?cb=") + Math.random(), { method: dry ? "GET" : "POST", timeout: 65000 });
    };
    call(true).then(function (plan) {
      reset();
      var n = plan.wouldSuggest || 0;
      if (!n) {
        // Say what the scan couldn't see. "Nothing new to suggest" reads as "you're all caught up",
        // which is wrong if folders failed or the run ran out of time.
        var caveats = [];
        if (plan.foldersFailed) caveats.push(plan.foldersFailed + " folder" + (plan.foldersFailed === 1 ? "" : "s") + " couldn’t be read");
        if (plan.foldersSkippedForTime) caveats.push(plan.foldersSkippedForTime + " skipped for time");
        showToast(svg(ICONS.sync) + '<span>Scanned ' + plan.listed + ' folder' + (plan.listed === 1 ? '' : 's') + ' — nothing new to suggest.' +
          (caveats.length ? ' ⚠ ' + escapeHtml(caveats.join(", ")) + ' — try again shortly.' : '') + '</span>', caveats.length ? 9000 : undefined);
        placementsStore.invalidate(); // the 30 s memo predates this scan (F194)
        box._loaded = true; loadFilesToPlace(box);
        return;
      }
      var capped = plan.wouldSuggestCapped || 0;
      var list = (plan.suggestions || []).map(function (p) { return '<li>' + escapeHtml(p) + '</li>'; }).join("");
      host.innerHTML = '<div class="ftp-h" role="heading" aria-level="3">Found by the scan <small>added to your review list, not the live site</small></div>' +
        '<ul class="ftp-plan">' + list + '</ul>' +
        (capped ? '<div class="wire-msg">' + capped + ' more held back by this run’s cap — they’ll come through on the next scan.</div>' : '') +
        (plan.foldersFailed || plan.foldersSkippedForTime
          ? '<div class="wire-msg err">⚠ ' + ((plan.foldersFailed || 0) + (plan.foldersSkippedForTime || 0)) + ' folder(s) weren’t read this run, so this list may be incomplete.</div>' : '') +
        '<div class="ftp-bar"><button class="wire-btn primary ftp-go" type="button">Add ' + n + ' to review</button>' +
        '<button class="wire-btn ftp-abort" type="button">Cancel</button></div>';
      host.querySelector(".ftp-abort").addEventListener("click", function () { loadFilesToPlace(box); });
      host.querySelector(".ftp-go").addEventListener("click", function () {
        var goDone = busyButton(this, "Adding…");
        call(false).then(function (res) {
          showToast(svg(ICONS.sync) + '<span>Added ' + res.suggested + ' suggestion' + (res.suggested === 1 ? '' : 's') + ' to the review list below. Nothing is live until you Keep them.</span>');
          placementsStore.invalidate(); // the new review rows are on the server, not in the memo (F194)
          loadFilesToPlace(box); refreshCovStats();
        }).catch(function (e) { goDone(); showError('<span>Couldn’t scan (' + escapeHtml(String((e && e.message) || e)) + ').</span>', 7000); });
      });
    }).catch(function (e) { reset(); showError('<span>Couldn’t scan (' + escapeHtml(String((e && e.message) || e)) + ').</span>', 7000); });
  }
  function loadFilesToPlace(box) {
    var host = box.querySelector("#ftpResults");
    // Generation guard, same as loadUnsorted: Place-all and Re-scan both reload this panel, and a slower
    // earlier scan finishing last used to paint its stale plan over the newer one.
    var live = loadGen(box);
    host.innerHTML = '<div class="wire-msg">Scanning folders…</div>';
    if (!window.PF_getToken) { host.innerHTML = '<div class="wire-msg err">Please sign in again to scan.</div>'; box._loaded = false; return; }
    var models = allModelsList();
    var watched = window.PF.deriveWatchedFolders(LINKS);
    var wiredSet = window.PF.wiredUrlSet(LINKS);
    // A resource wired to a whole folder already shows everything in that folder — don't offer to place
    // its files individually (it would add a redundant link and break the folder's auto-update).
    var folderWired = window.PF.folderWiredSet(LINKS);
    var dismissed = {}; ftpDismissed().forEach(function (u) { dismissed[window.PF.canonFileUrl(u)] = 1; });
    if (!watched.length) { host.innerHTML = '<div class="wire-msg">No file-linked folders to scan yet. (Once resources are linked to specific files, their folders get scanned here.)</div>'; return; }
    // What the nightly run has already done + what humans have already ruled out. Best-effort: if this
    // can't load we still scan (the panel worked without a server store before it existed).
    var storeP = placementsStore.get()
      .catch(function () { return { review: [], pending: [], skip: [], noAuto: [] }; });
    Promise.all([window.PF_getToken(), storeP]).then(function (both) {
      var bearer = "Bearer " + both[0], store = both[1], auto = [], pending = [];
      var noAuto = {}; (store.noAuto || []).forEach(function (u) { noAuto[window.PF.canonFileUrl(u)] = 1; });
      (store.skip || []).forEach(function (u) { dismissed[window.PF.canonFileUrl(u)] = 1; });
      // Three folders at a time (the nightly cron runs five against the same Graph limits). Strictly
      // sequential was ~7.5 s for 13 folders at 400-950 ms each before "Files to place" could fill.
      return lhRunPool(watched, function (w) {
          return fetchFolderChildren(w.folderUrl, bearer).then(function (files) {
            var plan = window.PF.planFolderPlacements(w.folderUrl, w.resource, files, models, wiredSet, folderWired);
            plan.autoPlace.forEach(function (x) {
              var c = window.PF.canonFileUrl(x.url);
              if (dismissed[c]) return;
              // A PM already removed this exact guess once — offer it, but don't call it confident again.
              if (noAuto[c]) pending.push({ file: x.file, url: x.url, folderUrl: x.folderUrl, resource: x.resource, candidates: [], reason: "removed-before" });
              else auto.push(x);
            });
            plan.pending.forEach(function (x) { if (!dismissed[window.PF.canonFileUrl(x.url)]) pending.push(x); });
          }, function () {}); // a folder that won't list is skipped, not fatal
      }, 3).then(function () {
        // Fold in anything the server knows about that this browser's scan didn't reach (a folder it
        // couldn't list, or a file the cron saw first). Client-scanned rows win — they carry live candidates.
        var seen = {}; pending.forEach(function (x) { seen[window.PF.canonFileUrl(x.url)] = 1; });
        (store.pending || []).forEach(function (x) {
          var c = window.PF.canonFileUrl(x.url);
          if (seen[c] || dismissed[c]) return; seen[c] = 1;
          pending.push({ file: x.file, url: x.url, folderUrl: x.folderUrl, resource: x.resource, candidates: [], reason: "from-nightly-scan", at: x.at });
        });
        if (!live()) return; // a newer scan started while this one ran — its results win
        renderFilesToPlace(box, host, auto, pending, store.review || []);
      });
    }).catch(function (e) { if (!live()) return; box._loaded = false; host.innerHTML = '<div class="wire-msg err">Couldn’t scan (' + escapeHtml(String((e && e.message) || e)) + ').</div>'; if (box._onError) box._onError(); });
  }
  function renderFilesToPlace(box, host, auto, pending, review) {
    review = review || []; auto = auto || []; pending = pending || [];
    // A nightly-scan suggestion names the model that was valid when the scan ran, but the catalog can change
    // during the day (a model renamed/removed). Keep/Place on a model the CURRENT catalog no longer has fails
    // with a confusing "couldn't find <model>". Re-route any suggestion whose model no longer resolves into
    // "Needs your input", where the human picks a valid model, instead of leaving a dead Keep/Place button.
    function ftpResolvable(name) { return !!findModelPath(name); }
    function ftpToPending(x) { return Object.assign({}, x, { candidates: (x.candidates || []).filter(ftpResolvable) }); }
    pending = pending.concat(
      review.filter(function (x) { return !ftpResolvable(x.model); }).map(ftpToPending),
      auto.filter(function (x) { return !ftpResolvable(x.model); }).map(ftpToPending)
    );
    review = review.filter(function (x) { return ftpResolvable(x.model); });
    auto = auto.filter(function (x) { return ftpResolvable(x.model); });
    // Oldest-waiting first — the longest-ignored suggestion leads each group. Rows with no recorded
    // suggested-date (found live this scan, not yet persisted) sort after the aged ones. `auto` is all
    // fresh-this-scan, so it keeps scan order.
    var byAge = function (a, b) { return (Date.parse(a.at) || Infinity) - (Date.parse(b.at) || Infinity); };
    review.sort(byAge); pending.sort(byAge);
    if (!auto.length && !pending.length && !review.length) { setFtpBadge(box, 0); host.innerHTML = '<div class="wire-msg">Every file in your linked folders is already placed.</div>'; return; }
    var html = "";
    // Found overnight and suggested — nothing is on the live site until a human clicks Add here.
    if (review.length) {
      html += '<div class="ftp-h" role="heading" aria-level="3">Suggested by the nightly scan — review before adding <small>nothing is live until you add it</small></div>' +
        '<ul class="wl-list ftp-review">' + review.map(function (x, i) {
          return '<li class="wl-row" data-i="' + i + '"><span class="lh-badge wanted">Suggested</span><span class="wl-where"><button class="ftp-fname" type="button" title="Preview">' + escapeHtml(x.file) + '</button> <span class="ftp-trail">→ ' + escapeHtml(x.model) + ' · ' + escapeHtml(x.resource) + '</span></span>' + ageChip(x.at, "waiting", true) +
            // Per-row accessible names: a list of 20 identical "Keep button" / "Remove button" tells a
            // screen-reader user nothing about WHICH file or model they are acting on.
            '<button class="wire-btn primary ftp-keep" type="button" aria-label="' + escapeAttr("Add " + x.file + " to " + x.model + " · " + x.resource) + '">Add</button>' +
            '<button class="ftp-reject ftp-undo" type="button" aria-label="' + escapeAttr("Don’t add " + x.file + " to " + x.model) + '">Remove</button></li>';
        }).join("") + '</ul>';
    }
    if (auto.length) {
      html += '<div class="ftp-h" role="heading" aria-level="3">Ready to place <small>confident matches</small></div>' +
        '<ul class="wl-list ftp-auto">' + auto.map(function (x, i) {
          return '<li class="wl-row" data-i="' + i + '"><span class="lh-badge wanted">Match</span><span class="wl-where"><button class="ftp-fname" type="button" title="Preview">' + escapeHtml(x.file) + '</button> <span class="ftp-trail">→ ' + escapeHtml(x.model) + ' · ' + escapeHtml(x.resource) + '</span></span>' +
            '<button class="wire-btn primary ftp-place" type="button" aria-label="' + escapeAttr("Add " + x.file + " to " + x.model + " · " + x.resource) + '">Add</button>' +
            '<button class="ftp-reject ftp-skip" type="button" title="Not a fit" aria-label="' + escapeAttr("Not a fit: " + x.file) + '">Remove</button></li>';
        }).join("") + '</ul>' +
        '<div class="ftp-allbar"><button class="wire-btn primary ftp-all" type="button">Add all ' + auto.length + '</button></div>';
    }
    if (pending.length) {
      html += '<div class="ftp-h" role="heading" aria-level="3">Needs your input <small>pick the right product</small></div>' +
        '<ul class="wl-list ftp-pend">' + pending.map(function (x, i) {
          // Start with nothing selected (placeholder button), NOT a pre-picked model. These rows exist
          // precisely because the matcher couldn't decide, so a pre-selection would be a wrong answer sitting
          // under a live Place button. Multi-select: pick one, several, or all models to place this file on.
          return '<li class="wl-row" data-i="' + i + '"><span class="lh-badge warn">Pick model</span><span class="wl-where"><button class="ftp-fname" type="button" title="Preview">' + escapeHtml(x.file) + '</button> <span class="ftp-trail">· ' + escapeHtml(x.resource) + '</span></span>' + ageChip(x.at, "waiting", true) +
            '<button class="wire-sel ftp-mbtn" type="button" aria-haspopup="dialog" aria-expanded="false" aria-label="Choose models for ' + escapeAttr(x.file) + '">Choose models…</button>' +
            '<button class="wire-btn primary ftp-place" type="button" disabled aria-label="' + escapeAttr("Add " + x.file + " to the models you pick") + '">Add</button>' +
            '<button class="ftp-reject ftp-skip" type="button" title="Not a fit" aria-label="' + escapeAttr("Not a fit: " + x.file) + '">Remove</button></li>';
        }).join("") + '</ul>';
    }
    host.innerHTML = html;
    // Mark the ITEM handled, not just its row. "Place all" reads the `auto` array, which the DOM removal
    // never touched — so a file the PM had just dismissed as "not a fit", or placed onto a different model
    // by hand, was replayed by the next "Place all" and wired anyway.
    function finishRow(li, x) {
      if (x) x._done = 1;
      // The handled row held focus; removing it would drop a keyboard user onto <body>. Remember a
      // neighbour BEFORE the removal, and announce what is left (the visual count is a badge).
      var neighbour = li.nextElementSibling || li.previousElementSibling;
      li.remove();
      var left = host.querySelectorAll(".ftp-auto .wl-row").length;
      var allBtn = host.querySelector(".ftp-all");
      if (allBtn) { if (left) allBtn.textContent = "Add all " + left; else { var bar = allBtn.closest(".ftp-allbar"); (bar || allBtn).remove(); } }
      setFtpBadge(box, host.querySelectorAll(".wl-row").length); // keep the header count in sync as rows are handled
      var sr = document.getElementById("srAnnounce"), n = host.querySelectorAll(".wl-row").length;
      if (sr) sr.textContent = n ? (n + " file" + (n === 1 ? "" : "s") + " left to review") : "All done here.";
      if (!n) { host.innerHTML = '<div class="wire-msg" tabindex="-1">All done here.</div>'; var done = host.querySelector(".wire-msg"); if (done) done.focus(); return; }
      applyFtpCollapse(); // the collapsed view was computed once, and drifted as rows were handled
      var visible = function (el) { return el && el.style.display !== "none"; };
      var row = visible(neighbour) ? neighbour : [].filter.call(host.querySelectorAll(".wl-row"), visible)[0];
      var tgt = row && row.querySelector("button:not([disabled])");
      if (tgt) tgt.focus();
    }
    // Show the first couple of rows so it's obvious there's work, and hide the rest behind a toggle.
    // Deliberately RE-RUNNABLE: handling a row changes which rows exist, so the count on the button, the
    // visibility of each section header and the Place-all footer all have to be recomputed. Computing it
    // once meant the two visible rows could be dealt with, leaving hidden rows behind a stale
    // "Review all N" button and headers that stayed hidden over rows that were now the only ones left.
    function applyFtpCollapse() {
      var old = host.querySelector(".ftp-showall"); if (old) old.remove();
      var rows = [].slice.call(host.querySelectorAll(".wl-row"));
      var collapsed = rows.length > 2 && !box._ftpExpanded;
      rows.forEach(function (li, i) { li.style.display = (collapsed && i >= 2) ? "none" : ""; });
      [].forEach.call(host.querySelectorAll(".ftp-h"), function (h) {
        var ul = h.nextElementSibling;
        var vis = ul && ul.querySelectorAll && [].some.call(ul.querySelectorAll(".wl-row"), function (li) { return li.style.display !== "none"; });
        h.style.display = vis ? "" : "none";
      });
      // Hide the "Place all" footer whenever no Ready-to-place row is actually showing.
      var autoVis = [].some.call(host.querySelectorAll(".ftp-auto .wl-row"), function (li) { return li.style.display !== "none"; });
      var bar = host.querySelector(".ftp-allbar"); if (bar) bar.style.display = autoVis ? "" : "none";
      if (!collapsed) return;
      var more = document.createElement("button");
      more.className = "wire-btn ftp-showall"; more.type = "button";
      more.textContent = "Review all " + rows.length + " files";
      more.addEventListener("click", function () { box._ftpExpanded = true; applyFtpCollapse(); });
      host.appendChild(more);
    }
    function wireRow(li, x, getModel, onPlace) {
      li.querySelector(".ftp-fname").addEventListener("click", function () { openResource([{ label: x.file, url: x.url }], x.file, ""); });
      li.querySelector(".ftp-skip").addEventListener("click", function () {
        var b = this; b.disabled = true;
        // Only drop the row once the server has the tombstone — otherwise tonight's run re-offers it.
        ftpMark("dismiss", x).then(function () { ftpDismiss(x.url); finishRow(li, x); })
          .catch(function (e) { b.disabled = false; showError('<span>Couldn’t save that — “' + escapeHtml(x.file) + '” is still listed. ' + escapeHtml(String((e && e.message) || e)) + '</span>', 7000); });
      });
      li.querySelector(".ftp-place").addEventListener("click", onPlace || function () {
        var b = this, mdl = getModel(); b.disabled = true; b.textContent = "…";
        // The placement itself is what matters here; recording it as reviewed is bookkeeping, so a
        // failure there is logged rather than surfaced — but it must not go unhandled.
        ftpPlace(x, mdl).then(function () { ftpMark("keep", x, mdl).catch(function (e) { console.warn("[pf] placement recorded but not marked reviewed", e); }); toastSaved('Added to ' + escapeHtml(mdl) + '.'); finishRow(li, x); refreshCovStats(); if (_todoIssuesRepaint) _todoIssuesRepaint(); })
          .catch(function (e) { b.disabled = false; b.textContent = "Add"; showError('<span>Couldn’t place (' + escapeHtml(String((e && e.message) || e)) + ').</span>', 6000); });
      });
    }
    [].forEach.call(host.querySelectorAll(".ftp-review .wl-row"), function (li) {
      var x = review[+li.getAttribute("data-i")];
      li.querySelector(".ftp-fname").addEventListener("click", function () { openResource([{ label: x.file, url: x.url }], x.file, ""); });
      li.querySelector(".ftp-keep").addEventListener("click", function () {
        var b = this; b.disabled = true; b.textContent = "…";
        // Keep is what puts the link on the live site now — the nightly scan only suggested it. Write the
        // catalog through the same proven merge path as every other placement, THEN record the verdict
        // (bookkeeping; a failure there is logged, not surfaced). A dedupe no-op if it was already wired.
        ftpPlace(x, x.model).then(function () {
          ftpMark("keep", x, x.model).catch(function (e) { console.warn("[pf] placement recorded but not marked reviewed", e); });
          toastSaved('Added to ' + escapeHtml(x.model) + '.'); finishRow(li, x); refreshCovStats(); if (_todoIssuesRepaint) _todoIssuesRepaint();
        }).catch(function (e) { b.disabled = false; b.textContent = "Add"; showError('<span>Couldn’t add that (' + escapeHtml(String((e && e.message) || e)) + ').</span>', 7000); });
      });
      li.querySelector(".ftp-undo").addEventListener("click", function () {
        var b = this; b.disabled = true; b.textContent = "…";
        ftpUnplace(x).then(function (res) {
          if (res && res.cancelled) { b.disabled = false; b.textContent = "Remove"; return; }
          showToast(svg(ICONS.ext) + '<span>Removed — it’ll come back under “Needs your input”.</span>'); finishRow(li, x); refreshCovStats();
        })
          .catch(function (e) { b.disabled = false; b.textContent = "Remove"; showError('<span>Couldn’t remove (' + escapeHtml(String((e && e.message) || e)) + ').</span>', 6000); });
      });
    });
    [].forEach.call(host.querySelectorAll(".ftp-auto .wl-row"), function (li) { var x = auto[+li.getAttribute("data-i")]; wireRow(li, x, function () { return x.model; }); });
    [].forEach.call(host.querySelectorAll(".ftp-pend .wl-row"), function (li) {
      var x = pending[+li.getAttribute("data-i")], mbtn = li.querySelector(".ftp-mbtn"), btn = li.querySelector(".ftp-place");
      var models = (x.candidates && x.candidates.length) ? x.candidates.slice() : allModelsList().map(function (m) { return m.model; });
      var selected = [];
      function repaint() { mbtn.textContent = unsMbtnLabel(models, selected); btn.disabled = !selected.length; }
      repaint();
      mbtn.addEventListener("click", function () { if (models.length) unsOpenCheckMenu(mbtn, models, selected, repaint); });
      // Place the file on ALL selected models in one merged batch (ftpPlaceAll handles the multi-write safely).
      wireRow(li, x, null, function () {
        var b = this; if (!selected.length) return; b.disabled = true; b.textContent = "…";
        var items = selected.map(function (m) { return { file: x.file, url: x.url, resource: x.resource, folderUrl: x.folderUrl, model: m }; });
        ftpPlaceAll(items).then(function (res) {
          if (res.failed.length) { b.disabled = false; b.textContent = "Add"; showError('<span>Couldn’t place on ' + res.failed.length + ' model' + (res.failed.length === 1 ? '' : 's') + ': ' + escapeHtml(res.failed.join("; ")) + '</span>', 9000); return; }
          selected.forEach(function (m) { ftpMark("keep", x, m).catch(function (e) { console.warn("[pf] placement recorded but not marked reviewed", e); }); });
          toastSaved('Added to ' + (selected.length > 1 ? selected.length + ' models' : escapeHtml(selected[0])) + '.');
          finishRow(li, x); refreshCovStats();
        }).catch(function (e) { b.disabled = false; b.textContent = "Add"; showError('<span>Couldn’t place (' + escapeHtml(String((e && e.message) || e)) + ').</span>', 6000); });
      });
    });
    var allBtn = host.querySelector(".ftp-all");
    if (allBtn) allBtn.addEventListener("click", function () {
      var b = this; b.disabled = true; b.textContent = "Placing…";
      // Only the rows still standing — never a file the PM already dismissed or placed by hand.
      var todo = auto.filter(function (x) { return !x._done; });
      if (!todo.length) { b.disabled = false; b.textContent = "Place all"; loadFilesToPlace(box); return; }
      ftpPlaceAll(todo).then(function (res) {
        todo.forEach(function (x) { x._done = 1; });
        if (res.failed.length) {
          showError('<span>Added ' + res.placed + ', but ' + res.failed.length + ' couldn’t be saved: ' + escapeHtml(res.failed.join("; ")) + '</span>', 10000);
        } else {
          // "Placed 0 files" reads as a failure when it actually means the planner de-duplicated —
          // every file was already linked, so there was nothing to write.
          if (res.placed) toastSaved('Added ' + res.placed + ' file' + (res.placed === 1 ? '' : 's') + '.');
          else showToast(svg(ICONS.ext) + '<span>Already linked — nothing to add.</span>');
        }
        refreshCovStats(); loadFilesToPlace(box);
      }).catch(function (err) {
        // Without this a hard rejection wedges the button on "Placing…" with an unhandled rejection.
        b.disabled = false; b.textContent = "Place all " + todo.length;
        showError('<span>Couldn’t place these (' + escapeHtml(String((err && err.message) || err)) + ').</span>', 8000);
      });
    });
    // Red count badge = everything awaiting a human decision.
    setFtpBadge(box, review.length + auto.length + pending.length);
    // Keep the panel a nudge, not a wall: show the first couple of rows so it's obvious there's work,
    // and hide the rest (plus any now-empty section header) behind a "Review all N" toggle.
    applyFtpCollapse(); // one implementation, re-run after every handled row (see finishRow)
  }

  // ("Linked but not opened" + unmatched-search signals moved to the dedicated Metrics tab; renderMetrics.)
  // (Orphan folder files are now actionable rows inside the To-Do "Sort the Pile" tile — see resolveOrphanItems.)

  /* ============================================================
     PM · TO-DO INBOX & ADD FILES
     ============================================================ */
  // Staff DEMAND, mined from search misses and folded into the worklist as "Wanted" rows. A miss is only
  // shown if the catalog STILL can't answer it (exact OR the new forgiving search) and it hasn't been
  // marked handled — so added-since products and plain typos drop off on their own.
  // "Handled" search misses. The server list (Search Aliases sheet, via /api/catalog) is what makes a
  // verdict hold for every PM and device; localStorage is the instant, offline-tolerant mirror for this
  // browser. Compared normalised (lower-case, single spaces) because the usage log keeps the typed case.
  function normMissQuery(q) { return String(q == null ? "" : q).toLowerCase().replace(/\s+/g, " ").trim(); }
  function demandLocalDone() { try { var a = JSON.parse(localStorage.getItem("pf-demand-done") || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
  function demandDismissed() { return demandLocalDone().map(normMissQuery).concat(RUNTIME_HANDLED); }
  function demandIsDone(q) { return demandDismissed().indexOf(normMissQuery(q)) !== -1; }
  function demandMarkDone(q) {
    var a = demandLocalDone(); if (a.indexOf(q) === -1) a.unshift(q);
    try { localStorage.setItem("pf-demand-done", JSON.stringify(a.slice(0, 200))); } catch (e) {}
    var n = normMissQuery(q); if (n && RUNTIME_HANDLED.indexOf(n) === -1) RUNTIME_HANDLED.push(n);
    // Share the verdict (POST /api/aliases {action:"dismiss-miss"}). The row is already gone locally; if
    // this fails the PM must know it is hidden here only, or a colleague re-does the same triage tomorrow.
    if (!window.PF_getApiToken) return;
    apiFetch("/api/aliases", { method: "POST", body: { action: "dismiss-miss", query: q } })
      .catch(function () { showError("<span>Marked handled in this browser only — couldn’t save it for the team. It will show again for others.</span>", 6000); });
  }
  function queryStillUnmet(q) {
    var query = String(q).trim().toLowerCase();
    if (!query) return false;
    if (window.PF && PF.expandSearchQuery) query = PF.expandSearchQuery(query, RUNTIME_ALIASES); // keep in lock-step with renderSearch
    var tokens = PF.tokenizeQuery(query);
    // Same predicate as renderSearch — if these two ever disagree, a query the search DOES answer keeps
    // being logged as a miss and shows up on the PM worklist as demand that doesn't exist.
    if (INDEX.some(function (n) { return PF.searchMatch(tokens, n.name, (n.trail || "") + " " + (n.extra || ""), n.model || n.name); })) return false;
    if (window.PF && PF.fuzzyPenalty && INDEX.some(function (n) { return PF.fuzzyPenalty(query, n.name + " " + (n.trail || "") + " " + (n.extra || ""), n.model || n.name) >= 0; })) return false;
    return true;
  }

  // "Do this next": one ranked queue of the highest-impact fixes, joining the client-side signals —
  // broken links (0), work-in-progress/empty (1), and linked-but-stale "review" (2) across all models,
  // plus an orphan-files pointer and (loaded async) the top staff searches that found nothing. Each fix
  // row has a "Fix" that opens that product's section editor; each "Wanted" row a "Request".
  // The fix-worklist signals: link issues (0), work-in-progress/empty (1), stale "review" (2) across all
  // models. Recomputed after a live link check so the rows reflect fresh statuses.
  function todoFixItems() {
    var items = [];
    coverageStats().groups.forEach(function (g) {
      g.models.forEach(function (mm) {
        RESOURCE_TYPES.forEach(function (r) {
          if (naFor(mm.name, r.name)) return;
          var hs = healthStatusFor(mm.name, r.name);
          if (hs === "broken") items.push({ prio: 0, tag: "Link issue", cls: "bad", model: mm.name, resource: r.name });
          else if (hs === "empty") items.push({ prio: 1, tag: "Work in progress", cls: "warn", model: mm.name, resource: r.name });
          else if (usableLinked(mm.name, r.name) && resIsStale(mm.name, r.name)) items.push({ prio: 2, tag: "Review", cls: "warn", model: mm.name, resource: r.name, note: resUpdatedLabel(mm.name, r.name).toLowerCase() });
        });
      });
    });
    items.sort(function (a, b) { return a.prio - b.prio; });
    return items;
  }
  // "To do" — the single PM inbox (one front door). Link issues / work-in-progress / review (with a live
  // "Check all links" that repaints them), the nightly "Files to place" suggestions (nested), and the top
  // staff searches that found nothing. Replaces the old scattered Do-this-next + Files-to-place + Link-health.
  var _todoIssuesRepaint = null; // set by todoPanel so the coverage-toolbar link check can repaint the issues box
  // A To-do result "tile": a collapsible section with a red count badge that hides when empty.
  // openDefault=false starts it collapsed (e.g. Sort the pile, which can be long).
  function todoBox(title, cls, iconHtml, subtitle, openDefault) {
    var d = document.createElement("details"); d.className = "todo-box " + cls; if (openDefault !== false) d.open = true; d.hidden = true;
    d.innerHTML = '<summary class="wire-head">' + (iconHtml || "") + '<span class="todo-box-title">' + escapeHtml(title) + (subtitle ? ' <small>' + escapeHtml(subtitle) + '</small>' : "") + '</span> <span class="ftp-badge" hidden></span></summary>' +
      '<div class="ui-body todo-body"></div>';
    return d;
  }
  // "To do" — the single PM inbox: ONE collapsible box containing collapsible result tiles (link issues,
  // files to place, sort the pile, searches with no result). Each tile carries a red count badge and hides
  // when empty; the whole box hides when there's nothing to do. The manual link check lives in the coverage
  // toolbar (always available).
  function todoPanel() {
    var box = document.createElement("details"); box.className = "wire todo-wrap"; box.open = true; box.hidden = true;
    box.innerHTML = '<summary class="wire-head todo-head">' + svg(ICONS.alert) + '<span role="heading" aria-level="2">To Do</span> <span class="ftp-badge todo-total" hidden></span></summary>' +
      '<div class="todo-boxes"></div>';
    var boxesEl = box.querySelector(".todo-boxes");
    var counts = { issues: 0, empty: 0, review: 0, place: 0, sort: 0, search: 0, reports: 0, fill: 0 }, errored = {};
    // Tiles that WILL report. The three token-gated ones are registered at their call sites, so a signed
    // -out/PM-less session doesn't wait forever for loaders that never run.
    var awaiting = { issues: 1, empty: 1, review: 1, place: 1, sort: 1 };
    function settled(k) { if (awaiting[k]) { delete awaiting[k]; reconcile(); } }
    setTimeout(function () { awaiting = {}; reconcile(); }, 12000); // a stuck loader must not strand the badge
    function counting() { for (var k in awaiting) if (awaiting.hasOwnProperty(k)) return true; return false; }
    function reconcile() {
      var total = counts.issues + counts.empty + counts.review + counts.place + counts.sort + counts.search + counts.reports + counts.fill;
      // Keep the box open when a tile FAILED to load, even at total 0 — otherwise a failed /api/unsorted,
      // /api/placements, or /api/usage-insights leaves the inbox looking empty while work is actually unknown.
      var anyErr = !!(errored.issues || errored.place || errored.sort || errored.search || errored.reports);
      box.hidden = total === 0 && !anyErr;
      var tb = box.querySelector(".todo-total");
      if (tb) {
        var busy = counting();
        tb.hidden = !busy && total === 0;
        tb.textContent = busy ? "…" : String(total);
        tb.setAttribute("aria-label", busy ? "Counting open items" : total + " open items");
      }
    }
    // 1. Fix work (sync). Split by urgency so the red count means "genuinely won't open":
    //    • Link Issues (red)   = broken + empty/WIP — shows "Open" but the file won't open / isn't there.
    //    • Needs review (amber) = linked, working, just stale — worth a refresh, not urgent.
    //    Both repaint after a live link check.
    // Broken and not-yet-filled are different jobs. Counting them together produced a red "14" that was
    // mostly work-in-progress folders, which teaches a PM to ignore the number that should mean "act now".
    var issues = todoBox("Link issues", "todo-issues", svg(ICONS.alert), "broken — these open nothing");
    issues._onCount = function (n) { counts.issues = n; settled("issues"); issues.hidden = n === 0; reconcile(); };
    var empties = todoBox("Empty folders", "todo-empty", svg(ICONS.folder), "linked, but nothing in them yet", false);
    empties._onCount = function (n) { counts.empty = n; settled("empty"); empties.hidden = n === 0; reconcile(); };
    var review = todoBox("Needs review", "todo-review", svg(ICONS.edit), null, false); // collapsed by default — non-urgent
    review._onCount = function (n) { counts.review = n; settled("review"); review.hidden = n === 0; reconcile(); };
    // Render one fix-work tile. `withRecheck` adds the "Re-check links" action — only on the health-based
    // Link Issues tile; the Needs-review tile is date-based, so a re-check can't clear its rows.
    function fillFixTile(boxEl, items, withRecheck) {
      var top = items.slice(0, 20), host = boxEl.querySelector(".todo-body"), recheck = "";
      if (withRecheck) {
        var ts = lastHealthTs();
        recheck = '<div class="lh-actions todo-recheck"><button class="wire-btn lh-recheck" type="button">' + svg(ICONS.ext) + '<span>Re-check links</span></button><span class="lh-status"></span><span class="lh-when">' + (ts ? "· last checked " + healthWhen() : "") + '</span></div>';
      }
      host.innerHTML = items.length
        ? '<ul class="wl-list">' + top.map(function (it, i) {
            return '<li class="wl-row" data-i="' + i + '"><span class="lh-badge ' + it.cls + '">' + it.tag + '</span>' +
              '<span class="wl-where"><b>' + escapeHtml(it.model) + '</b> · ' + escapeHtml(it.resource) + (it.note ? ' <span class="wl-note">' + escapeHtml(it.note) + '</span>' : '') + '</span>' +
              '<button class="wire-btn primary wl-fix" type="button" aria-label="Fix ' + escapeAttr(it.model + " " + it.resource) + '">Fix</button></li>';
          }).join("") + '</ul>' + (items.length > 20 ? '<div class="wire-msg">Showing 20 of ' + items.length + '.</div>' : '') + recheck
        : "";
      [].forEach.call(host.querySelectorAll(".wl-fix"), function (b) {
        var it = top[+b.closest(".wl-row").getAttribute("data-i")];
        b.addEventListener("click", function () { openFixEditorFor(it.model, it.resource); });
      });
      var rc = host.querySelector(".lh-recheck");
      if (rc) rc.addEventListener("click", function () { runLinkCheck(this, host.querySelector(".lh-status"), function () { paintFixWork(); }); });
      setFtpBadge(boxEl, items.length); // updates the badge AND fires _onCount (hide + reconcile)
    }
    function paintFixWork() {
      var all = todoFixItems();
      fillFixTile(issues, all.filter(function (it) { return it.prio === 0; }), true);   // broken → red, act now
      fillFixTile(empties, all.filter(function (it) { return it.prio === 1; }), false); // linked but empty → amber
      fillFixTile(review, all.filter(function (it) { return it.prio === 2; }), false);  // aging/stale → amber
    }
    _todoIssuesRepaint = paintFixWork;
    boxesEl.appendChild(issues); boxesEl.appendChild(empties); boxesEl.appendChild(review); paintFixWork();
    // 2. Files to place (async, existing panel) — hide when empty, report its count.
    var ftp = filesToPlacePanel(); ftp.hidden = true;
    ftp._onCount = function (n) { counts.place = n; settled("place"); ftp.hidden = n === 0; reconcile(); };
    ftp._onError = function () { errored.place = 1; ftp.hidden = false; reconcile(); };
    boxesEl.appendChild(ftp);
    // 3. Sort the pile (async) — the pile, moved here from the Add-files tool. Collapsed by default
    // (it can be long) so it doesn't dominate the inbox.
    var sort = todoBox("Sort the pile", "todo-sort", svg(ICONS.folder), null, false); // collapsed by default — can be long
    var uns = unsortedPanel(true); // bare (no own header — the box supplies it)
    uns._onCount = function (n) { counts.sort = n; settled("sort"); setFtpBadge(sort, n); sort.hidden = n === 0; reconcile(); };
    uns._onError = function () { errored.sort = 1; sort.hidden = false; reconcile(); }; // reveal the tile so its error shows
    sort.querySelector(".todo-body").appendChild(uns);
    boxesEl.appendChild(sort);
    // 4. Searches with no result (async) — staff searched and got nothing. Point one at a product, or request it.
    var search = todoBox("Searches with no result", "todo-search", svg(ICONS.ext), "last 60 days");
    search._onError = function () { errored.search = 1; search.hidden = false; reconcile(); };
    boxesEl.appendChild(search);
    if (window.PF_getApiToken) awaiting.search = 1, loadTodoDemand(search, function (n) { counts.search = n; settled("search"); setFtpBadge(search, n); search.hidden = n === 0; reconcile(); });
    // 5. Reports (async) — staff bug reports and flagged files landed only on Admin, a tab a PM has no
    // reason to open on a Monday, so the queue aged unseen. This tile is a COUNT and a doorway; triage
    // stays on Admin rather than becoming a second list with different actions.
    var reports = todoBox("Reports", "todo-reports", svg(ICONS.spec), "from staff");
    reports._onError = function () { errored.reports = 1; reports.hidden = false; reconcile(); };
    boxesEl.appendChild(reports);
    if (window.PF_getApiToken) awaiting.reports = 1, loadTodoReports(reports, function (n) { counts.reports = n; settled("reports"); setFtpBadge(reports, n); reports.hidden = n === 0; reconcile(); });
    // 6. Fill these next — slots that were NEVER linked, on models people actually open. The other fix
    // tiles are health-based (broken / empty / stale) and so cover only slots that already HAVE a link;
    // a never-linked resource on a busy model appeared nowhere in the inbox. Capped at 3 and collapsed:
    // this is a nudge, not the whole backlog (the full ranking lives on Metrics).
    var fill = todoBox("Fill these next", "todo-fill", svg(ICONS.edit), "unlinked, on models people open", false);
    boxesEl.appendChild(fill);
    if (window.PF_getApiToken) awaiting.fill = 1, loadTodoFillNext(fill, function (n) { counts.fill = n; settled("fill"); setFtpBadge(fill, n); fill.hidden = n === 0; reconcile(); });
    reconcile();
    return box;
  }
  // Rank never-linked slots by how much the model is actually opened, and offer the same Fix button as the
  // other To-do rows. Silent on any failure — a missing nudge is not worth an error row in the inbox.
  function loadTodoFillNext(box, onCount) {
    var wrap = box.querySelector(".todo-body");
    usageInsights().then(function (data) {
      var opensByModel = (data && data.opensByModel) || {};
      var rows = [];
      coverageStats().groups.forEach(function (g) {
        g.models.forEach(function (mm) {
          var opens = opensByModel[mm.name] || 0;
          if (!opens) return; // no demand signal → not "next"
          RESOURCE_TYPES.forEach(function (r) {
            if (naFor(mm.name, r.name)) return;
            if (usableLinked(mm.name, r.name)) return;      // already linked (health tiles own those)
            if (healthStatusFor(mm.name, r.name)) return;   // broken/empty → already in a fix tile
            rows.push({ model: mm.name, resource: r.name, prod: g.prod, opens: opens });
          });
        });
      });
      rows.sort(function (a, b) { return b.opens - a.opens || a.model.localeCompare(b.model) || a.resource.localeCompare(b.resource); });
      var top = rows.slice(0, 3);
      onCount(top.length);
      if (!top.length) { wrap.innerHTML = ""; return; }
      wrap.innerHTML = '<p class="wire-lead">Nothing is linked here yet, and staff open these models — the highest-value gaps to close first.</p>' +
        '<ul class="wl-list">' + top.map(function (it, i) {
          return '<li class="wl-row" data-i="' + i + '"><span class="lh-badge wanted">Not linked</span>' +
            '<span class="wl-where"><b>' + escapeHtml(it.model) + '</b> · ' + escapeHtml(it.resource) +
            ' <span class="wl-note">' + it.opens + ' open' + (it.opens === 1 ? '' : 's') + ' on this model</span></span>' +
            '<button class="wire-btn primary wl-fix" type="button" aria-label="Fix ' + escapeAttr(it.model + " " + it.resource) + '">Fix</button></li>';
        }).join("") + '</ul>' +
        (rows.length > top.length ? '<div class="wire-msg">Showing the top 3 of ' + rows.length + ' — the full ranking is on Metrics.</div>' : '');
      [].forEach.call(wrap.querySelectorAll(".wl-fix"), function (b) {
        var it = top[+b.closest(".wl-row").getAttribute("data-i")];
        b.addEventListener("click", function () { openFixEditorFor(it.model, it.resource); });
      });
    }).catch(function () { onCount(0); });
  }
  // "Is coverage moving?" — one number beside the big percentage, from the Coverage History sheet the
  // nightly cron already writes. Silent (stays hidden) until there are two snapshots to compare; a chart
  // for this lives on Metrics, and Progress only needs the direction.
  function fillCoverageDelta(el) {
    if (!el || !window.PF_getApiToken) return;
    usageInsights().then(function (data) {
      var ct = (data && data.coverageTrend) || [];
      if (ct.length < 2) return;
      var last = ct[ct.length - 1], prev = ct[ct.length - 2];
      var d = Math.round(Number(last.pct) - Number(prev.pct));
      if (!isFinite(d)) return;
      // Weeks are bucketed in UTC by the cron (weekMondayISO) — parse them the same way, deliberately,
      // so the label names the same week the snapshot belongs to. (Same exception as the weekly chart.)
      var since = ""; try { since = new Date(prev.week + "T00:00:00Z").toLocaleDateString(undefined, { month: "short", day: "numeric", timeZone: "UTC" }); } catch (e) {}
      var arrow = d > 0 ? "▲" : d < 0 ? "▼" : "—";
      var word = d > 0 ? "up" : d < 0 ? "down" : "level";
      el.className = "cov-delta" + (d > 0 ? " up" : d < 0 ? " down" : "");
      el.innerHTML = ' · <span aria-hidden="true">' + arrow + '</span> ' +
        (d === 0 ? 'level' : Math.abs(d) + ' pt' + (Math.abs(d) === 1 ? '' : 's')) +
        (since ? ' since ' + escapeHtml(since) : ' since the last snapshot') +
        '<span class="sr-only"> (' + word + ')</span>';
      el.hidden = false;
    }).catch(function () { /* no delta is fine — the number beside it is still true */ });
  }
  // Open bug reports / content flags for the To-do inbox. Counts only what still needs a decision
  // ("new" / "in progress") — a resolved report is history, and counting it would keep a badge red forever.
  function loadTodoReports(box, onCount) {
    var wrap = box.querySelector(".todo-body");
    apiFetch("/api/bug-reports", { fresh: true })
      .then(function (data) {
        var open = ((data && data.reports) || []).filter(function (x) { return x.status === "new" || x.status === "in_progress"; });
        onCount(open.length);
        if (!open.length) { wrap.innerHTML = ""; return; }
        // Oldest-first so the age chip describes the most-neglected one — the reason to go look.
        var oldest = open.slice().sort(function (a, b) { return String(a.created_at || "9999").localeCompare(String(b.created_at || "9999")); })[0];
        wrap.innerHTML = '<p class="wire-lead">' + open.length + ' report' + (open.length === 1 ? '' : 's') +
          ' from staff still need' + (open.length === 1 ? 's' : '') + ' a decision. Read and answer them on the Admin tab.</p>' +
          '<div class="lh-actions"><button class="wire-btn primary todo-reports-go" type="button">Open Admin</button>' +
          (oldest && oldest.created_at ? ageChip(oldest.created_at, "oldest", true) : '') + '</div>';
        var go = wrap.querySelector(".todo-reports-go");
        // Click the real tab rather than setting `view` — that path also writes pf-lasttab, syncs the hash
        // and resets scroll, which a direct assignment silently skipped.
        if (go) go.addEventListener("click", function () { var t = document.querySelector('.vtab[data-view="admin"]'); if (t) t.click(); });
      })
      .catch(function () { if (wrap) wrap.innerHTML = '<div class="wire-msg err">Couldn’t load staff reports right now.</div>'; if (box._onError) box._onError(); });
  }
  // Run the on-demand link check (moved to the coverage toolbar). Persists fresh statuses via saveHealth
  // then calls onDone (repaint the To-do issues box + refresh the "last checked" label). btn/statusEl are
  // the toolbar control's elements.
  function runLinkCheck(btn, statusEl, onDone) {
    if (!window.PF_getToken) { showError('<span>Please sign in again.</span>'); return; }
    var items = collectWiredResources();
    if (!items.length) { showToast('<span>No links yet — nothing to check.</span>'); return; }
    // Save/restore innerHTML, not textContent — these buttons hold an SVG icon + a text span, and
    // textContent would strip the icon permanently after the first run.
    var btnDone = busyButton(btn, "Checking…"); if (statusEl) statusEl.textContent = " 0 / " + items.length;
    var cache = {};
    window.PF_getToken().then(function (token) {
      var bearer = "Bearer " + token;
      return lhRunPool(items, function (it) { return checkOneResource(it, bearer, cache); }, 5, function (d, t) { if (statusEl) statusEl.textContent = " " + d + " / " + t; });
    }).then(function (results) {
      items.forEach(function (it, i) { it.result = results[i]; });
      btnDone(); if (statusEl) statusEl.textContent = "";
      var map = {}, okN = 0, probs = [];
      items.forEach(function (it) {
        var s = it.result && it.result.status;
        if (s === "broken" || s === "empty") { map[it.model + "::" + it.resource] = s; probs.push({ model: it.model, resource: it.resource, fam: it.fam, prod: it.prod, path: it.path, status: s, detail: it.result.detail, url: it.result.url }); }
        else if (s !== "unknown") okN++;
      });
      saveHealth(map, Date.now(), okN, probs);
      if (srAnnounce) srAnnounce.textContent = "Checked " + items.length + " link" + (items.length === 1 ? "" : "s") + " — " + (probs.length ? probs.length + " issue" + (probs.length === 1 ? "" : "s") : "no issues") + ".";
      if (onDone) onDone();
    }).catch(function (err) {
      btnDone(); if (statusEl) statusEl.textContent = "";
      showError('<span>Couldn’t run the check (' + escapeHtml(lhReason(err)) + ').</span>', 5000);
    });
  }
  // /api/usage-insights feeds several surfaces (To-do demand, the coverage delta on Progress, Metrics).
  // Memoize it briefly so opening Progress doesn't fetch the same payload twice. A failure is never
  // cached — the next caller retries.
  var _uiCache = null, _uiAt = 0;
  function usageInsights() {
    if (_uiCache && Date.now() - _uiAt < 60000) return _uiCache;
    _uiAt = Date.now();
    _uiCache = apiFetch("/api/usage-insights").catch(function (e) { _uiCache = null; throw e; });
    return _uiCache;
  }
  function loadTodoDemand(searchBox, onCount) {
    var wrap = searchBox.querySelector(".todo-body");
    var live = loadGen(searchBox); // a Re-check or a tab switch reloads this while a slow read is in flight
    usageInsights()
      .then(function (data) {
        if (!live()) return;
        var demand = (data.topMisses || []).filter(function (m) { return !demandIsDone(m.query) && queryStillUnmet(m.query); }).slice(0, 8);
        onCount(demand.length);
        if (!demand.length) { wrap.innerHTML = ""; return; }
        wrap.innerHTML = '<p class="wire-lead">Staff searched for these and found nothing — point one at a product (or a file type) so search finds it, or send a request with the query filled in. ' +
          'The full history, handled ones included, is on the <button class="linklike todo-go-metrics" type="button">Metrics tab</button>.</p><ul class="wl-list"></ul>';
        var goMet = wrap.querySelector(".todo-go-metrics");
        if (goMet) goMet.addEventListener("click", function () { var t = document.querySelector('.vtab[data-view="metrics"]'); if (t) t.click(); });
        var ul = wrap.querySelector(".wl-list");
        demand.forEach(function (m) { var li = document.createElement("li"); li.className = "wl-row wl-demand"; ul.appendChild(li); renderDemandRow(li, m, searchBox, onCount); });
      })
      .catch(function () { if (!live()) return; if (wrap) wrap.innerHTML = '<div class="wire-msg err">Couldn’t load staff searches right now.</div>'; if (searchBox._onError) searchBox._onError(); });
  }
  // A "searches with no result" row: + Alias (point it at a product), Request, or dismiss. Rebuildable so
  // cancelling the inline picker restores the row with working buttons.
  function renderDemandRow(li, m, searchBox, onCount) {
    li.innerHTML = '<span class="lh-badge wanted">No match</span>' +
      '<span class="wl-where"><b class="wl-q">“' + escapeHtml(m.query) + '”</b> · searched ' + m.count + '×</span>' +
      '<button class="wire-btn wl-alias" type="button" title="Point this search at a product">+ Alias</button>' +
      '<button class="wire-btn wl-req" type="button">Request</button>' +
      '<button class="wl-dismiss" type="button" aria-label="Mark “' + escapeAttr(m.query) + '” handled" title="Mark handled">×</button>';
    li.querySelector(".wl-alias").addEventListener("click", function () { openWlAlias(li, m, searchBox, onCount); });
    // Prefill with the query that missed — this row exists BECAUSE of that search, so an empty form just
    // made the PM retype it (or send it blank).
    li.querySelector(".wl-req").addEventListener("click", function () { openSearchRequest(m.query); });
    li.querySelector(".wl-dismiss").addEventListener("click", function () {
      demandMarkDone(m.query);
      var ul = li.parentNode; li.remove();
      onCount(ul ? ul.querySelectorAll(".wl-demand").length : 0); // updates the box badge + hides it when empty
    });
  }
  // Inline "point this search at a product" picker — a product dropdown (not free text). Reuses saveAlias
  // (POST /api/aliases) with the chosen product as the target; the query then finds it, so it drops off next render.
  function openWlAlias(li, m, searchBox, onCount) {
    // Two optgroups. A miss is often a VOCABULARY gap, not a missing product — someone searched "cert"
    // or "install guide" and meant a resource type. The runtime alias `to` has always accepted any
    // string (the built-ins map "compliance" → "compliance & testing"); only this picker refused to
    // offer one, so those misses had no fix short of editing the sheet.
    var prodOpts = "";
    CATALOG.forEach(function (c) { if (c.name === OTHER_COLLATERAL) return; c.families.forEach(function (f) { f.products.forEach(function (p) { if (p.models && p.models.length) prodOpts += '<option value="' + escapeAttr(p.name) + '">' + escapeHtml(f.name + " › " + p.name) + '</option>'; }); }); });
    var typeOpts = RESOURCE_TYPES.map(function (r) { return '<option value="' + escapeAttr(r.name) + '">' + escapeHtml(r.name) + '</option>'; }).join("");
    var opts = '<option value="">Choose what it should find…</option>' +
      '<optgroup label="Products">' + prodOpts + '</optgroup>' +
      '<optgroup label="Resource types">' + typeOpts + '</optgroup>';
    li.innerHTML = '<span class="lh-badge wanted">No match</span>' +
      '<span class="wl-where"><b class="wl-q">“' + escapeHtml(m.query) + '”</b></span>' +
      '<span class="met-alias-form wl-alias-form"><span class="met-arrow" aria-hidden="true">→</span>' +
      '<select class="met-alias-sel" aria-label="What “' + escapeAttr(m.query) + '” should find">' + opts + '</select>' +
      '<button class="met-alias-save" type="button">Save</button>' +
      '<button class="met-alias-cancel" type="button">Cancel</button></span>';
    var sel = li.querySelector(".met-alias-sel"); try { sel.focus(); } catch (e) {}
    li.querySelector(".met-alias-cancel").addEventListener("click", function () { renderDemandRow(li, m, searchBox, onCount); });
    li.querySelector(".met-alias-save").addEventListener("click", function () {
      if (!sel.value) { sel.focus(); sel.classList.add("met-alias-bad"); sel.setAttribute("aria-invalid", "true"); return; }
      saveAlias(li, m.query, sel.value, function () {
        // The query now resolves to a product, so it's handled: mark it done and drop it from the tile count
        // (leave the "→ alias added" confirmation visible, just stop counting it).
        demandMarkDone(m.query); li.classList.remove("wl-demand");
        onCount(searchBox.querySelectorAll(".wl-demand").length);
      });
    });
    sel.addEventListener("change", function () { sel.classList.remove("met-alias-bad"); sel.removeAttribute("aria-invalid"); });
  }
  // Unified "Add files" tool — one panel, a 2-way mode toggle swapping between the two upload flows.
  // Each mode is built lazily on first selection and cached (show/hide), so toggling preserves in-progress
  // work and never refetches. Modes: "doc" (one file → many products), "folder" (a folder → a model).
  // (Sorting the pile moved to the To-Do inbox.) The switcher is a proper ARIA tablist: roving tabindex,
  // arrow-key nav, and each tab wired to its role="tabpanel" via aria-controls / aria-labelledby.
  var _addFilesMode = "folder", _afSeq = 0;
  function addFilesPanel() {
    var box = document.createElement("div"); box.className = "wire addfiles"; box.setAttribute("data-anim", "addfiles");
    var uid = "af" + (++_afSeq); // unique tab/panel ids so multiple instances don't collide
    var modes = [
      { key: "doc", label: "One file → many products", build: function () { return docPlacePanel(true); } },
      { key: "folder", label: "A folder → a model", build: function () { return wirePanel(true); } }
    ];
    box.innerHTML =
      '<div class="wire-head" role="heading" aria-level="2">' + svg(ICONS.sync) + '<span>Add files</span></div>' +
      '<div class="af-seg" role="tablist" aria-label="Add-files mode"></div>' +
      '<div class="af-body"></div>';
    var seg = box.querySelector(".af-seg"), body = box.querySelector(".af-body"), built = {};
    function tabId(k) { return uid + "-tab-" + k; }
    function panelId(k) { return uid + "-panel-" + k; }
    function show(key, focusTab) {
      _addFilesMode = key;
      if (!built[key]) {
        var m = modes.filter(function (x) { return x.key === key; })[0]; if (!m) return;
        var node = m.build(); node.setAttribute("data-mode", key);
        node.id = panelId(key); node.setAttribute("role", "tabpanel"); node.setAttribute("aria-labelledby", tabId(key)); node.setAttribute("tabindex", "0");
        built[key] = node; body.appendChild(node);
        // aria-controls only once the panel it names actually exists — panels are built lazily, so
        // setting it up front left the unselected tab pointing at a nonexistent id (a broken IDREF).
        var tb = seg.querySelector('[data-mode="' + key + '"]'); if (tb) tb.setAttribute("aria-controls", panelId(key));
      }
      [].forEach.call(seg.children, function (b) {
        var on = b.getAttribute("data-mode") === key;
        b.classList.toggle("on", on); b.setAttribute("aria-selected", on ? "true" : "false"); b.tabIndex = on ? 0 : -1; // roving tabindex
      });
      [].forEach.call(body.children, function (c) { c.hidden = c.getAttribute("data-mode") !== key; });
      if (focusTab) { var t = seg.querySelector('[data-mode="' + key + '"]'); if (t) t.focus(); }
    }
    modes.forEach(function (m) {
      var b = document.createElement("button"); b.type = "button"; b.className = "af-seg-btn"; b.setAttribute("data-mode", m.key);
      b.setAttribute("role", "tab"); b.id = tabId(m.key); b.tabIndex = -1; // aria-controls is set in show(), when the panel exists
      b.textContent = m.label;
      b.addEventListener("click", function () { show(m.key); });
      seg.appendChild(b);
    });
    // Arrow-key navigation within the tablist (roving tabindex): Left/Right cycle, Home/End jump to the ends.
    seg.addEventListener("keydown", function (e) {
      if (["ArrowLeft", "ArrowRight", "Home", "End"].indexOf(e.key) === -1) return;
      var tabs = [].slice.call(seg.children), i = tabs.indexOf(document.activeElement); if (i === -1) return;
      var ni = e.key === "ArrowLeft" ? (i - 1 + tabs.length) % tabs.length : e.key === "ArrowRight" ? (i + 1) % tabs.length : e.key === "Home" ? 0 : tabs.length - 1;
      e.preventDefault(); show(tabs[ni].getAttribute("data-mode"), true);
    });
    show(modes.some(function (m) { return m.key === _addFilesMode; }) ? _addFilesMode : "folder");
    // (No _showMode hook: nothing ever called it. Re-add it the day something needs to jump to a mode.)
    return box;
  }

  /* ============================================================
     PM · PROGRESS VIEW, PM GUIDE & ALIAS FORM
     ============================================================ */
  function renderCoverage() {
    // Drop any repaint hook left over from the previous render of this view. Without this the coverage
    // toolbar's "Check links" callback repaints a DETACHED tree built by an earlier render.
    _todoIssuesRepaint = null;
    if (catalogError) { content.innerHTML = catalogUnavailable("Product team", "Progress"); syncViewTabs(); return; }
    var s = coverageStats(), N = RESOURCE_TYPES.length;
    var pct = s.totalSlots ? Math.floor(s.linkedSlots / s.totalSlots * 100) : 0; // floor so 199/200 reads 99%, not a misleading 100%
    content.innerHTML = head("Product team", "Progress", "Your coverage at a glance — what’s linked and which models still need work. You can also link folders and sort stray files. Nothing changes on the live site until you approve a preview; run a full link check from the toolbar above the grid.");
    var wrap = document.createElement("div"); wrap.className = "cov-view view";

    // "To do" leads the page — the single PM inbox (link issues, work-in-progress, the nightly file
    // suggestions nested inside, staff searches with no match, + a live "Check all links"). Everything that
    // needs a decision in one front door, so nothing is buried across separate panels/tabs.
    wrap.appendChild(todoPanel());
    var hyg = sheetHygienePanel(); if (hyg) wrap.appendChild(hyg); // what the sheet says that the app can't read as intended

    var sum = document.createElement("div"); sum.className = "cov-summary";
    sum.innerHTML =
      '<div><div class="cov-big">' + pct + '%<small>' + s.linkedSlots + ' of ' + s.totalSlots + ' linked</small></div>' +
        '<div class="cov-sub">' + s.totalModels + ' models · ' + s.totalSlots + ' applicable resources<span class="cov-delta" hidden></span></div></div>' +
      '<div class="cov-barwrap">' + covBar(pct) + '</div>';
    wrap.appendChild(sum);
    fillCoverageDelta(sum.querySelector(".cov-delta"));

    var types = document.createElement("div"); types.className = "cov-types";
    RESOURCE_TYPES.forEach(function (r) {
      var el = document.createElement("div"); el.className = "cov-type";
      el.innerHTML = covTypeHtml(r, s);
      types.appendChild(el);
    });
    wrap.appendChild(types);

    var toolbar = document.createElement("div"); toolbar.className = "cov-toolbar";
    var lbl = document.createElement("label"); lbl.className = "cov-check";
    lbl.innerHTML = '<input type="checkbox"' + (covIncompleteOnly ? " checked" : "") + '><span class="cov-check-box" aria-hidden="true">' + svg('<path d="M20 6 9 17l-5-5"/>') + '</span><span>Only models missing links</span>';
    lbl.querySelector("input").addEventListener("change", function () {
      // Refiltering rebuilds the matrix, and an open editor's staged edits live only in it (F200).
      if (document.querySelector(".cov-grp-head.editing")) { this.checked = covIncompleteOnly; showToast('<span>Save or cancel the section you’re editing first.</span>'); return; }
      covIncompleteOnly = this.checked; fillMatrix();
    });
    toolbar.appendChild(lbl);
    var legend = document.createElement("span"); legend.className = "cov-legend";
    // Legend mirrors the actual cell glyphs (✓ / number / – / !), not just colors, so the vocabulary is legible incl. colorblind.
    // Each glyph is wrapped WITH its label so the two stay together and the space falls BETWEEN items (not
    // between a glyph and its own word). The "not applicable" key only appears when a model actually has an
    // N/A — otherwise the glyph never shows in the grid and the key would just be noise.
    function legItem(cls, glyph, label) { return '<span class="cov-legend-item"><span class="cov-cell ' + cls + '" aria-hidden="true">' + glyph + '</span> ' + label + '</span>'; }
    var legendHtml =
      legItem("on", "✓", "linked") +
      legItem("on", "3", "= number of links") +
      legItem("", "–", "not linked") +
      legItem("on review", "✓", "review (over a year old)") +
      legItem("broken", "!", "link issue");
    if (Object.keys(NA).length) legendHtml += legItem("na", "", "not applicable");
    legend.innerHTML = legendHtml;
    toolbar.appendChild(legend);
    // Manual link check lives here (always available) — the To-do "Link issues" box only appears when there
    // ARE issues, so the check button no longer clutters the inbox when everything's clean. A run repaints
    // the To-do issues box in place.
    var lhTool = document.createElement("span"); lhTool.className = "cov-lh-tool";
    lhTool.innerHTML = '<button class="cov-lh-btn" type="button">' + svg(ICONS.ext) + '<span>Check links</span></button><span class="cov-lh-status"></span><span class="cov-lh-when"></span>';
    function setCovLhWhen() { var w = lhTool.querySelector(".cov-lh-when"); var ts = lastHealthTs(); if (w) w.textContent = ts ? ("· last checked " + healthWhen()) : ""; }
    lhTool.querySelector(".cov-lh-btn").addEventListener("click", function () {
      runLinkCheck(this, lhTool.querySelector(".cov-lh-status"), function () { setCovLhWhen(); if (_todoIssuesRepaint) _todoIssuesRepaint(); });
    });
    setCovLhWhen();
    toolbar.appendChild(lhTool);
    // "Scan SharePoint" belongs here for the same reason "Check links" does: its only other entry point is
    // the Re-scan button INSIDE the Files-to-place tile, and that tile hides itself when the count is zero
    // — which is exactly the moment a PM who just dropped files into a folder wants to run it.
    var scanTool = document.createElement("span"); scanTool.className = "cov-lh-tool";
    scanTool.innerHTML = '<button class="cov-lh-btn cov-scan-btn" type="button">' + svg(ICONS.sync) + '<span>Scan SharePoint</span></button>';
    scanTool.querySelector(".cov-scan-btn").addEventListener("click", function () {
      var box = document.querySelector(".files-to-place");
      if (!box || !box.querySelector("#ftpResults")) {
        showError('<span>Open the “New on SharePoint” box on this tab first, then scan.</span>', 6000);
        return;
      }
      if (!box.open) box.open = true; // make the results visible before they arrive
      ftpRescan(box, this); // dry-run first, exactly like the box's own button
    });
    toolbar.appendChild(scanTool);
    wrap.appendChild(toolbar);

    var matrix = document.createElement("div"); matrix.className = "cov-matrix";
    // Rebuild ONLY the matrix (recomputing coverage live) — used by the "incomplete only" toggle so it
    // no longer tears down the whole view (which re-created the wire tool and refetched /api/unsorted,
    // discarding any in-progress drag work).
    function fillMatrix(pre) { // `pre`: stats the caller already computed (the initial render); omitted → recompute live
      matrix.innerHTML = "";
      var st = pre || coverageStats(), anyShown = false;
      st.groups.forEach(function (g) {
        var models = covIncompleteOnly ? g.models.filter(function (m) { return m.count < m.applicable; }) : g.models;
        if (!models.length) return;
        anyShown = true;
        var grp = document.createElement("div"); grp.className = "cov-group";
        grp.dataset.prod = g.prod; grp.dataset.fam = g.fam; // so "Fix link" can locate + open this section's editor
        renderCovGroup(grp, g, models, N);
        matrix.appendChild(grp);
      });
      if (!anyShown) {
        var e = document.createElement("div"); e.className = "cov-empty";
        e.textContent = covIncompleteOnly ? "Every model with data is fully linked." : "No models with data yet.";
        matrix.appendChild(e);
      }
    }
    fillMatrix(s); // the stats computed at the top of this render — no second pass over the catalog
    wrap.appendChild(matrix);

    // Authoring tools live BELOW the coverage overview: a PM opening Progress to triage "what's missing"
    // sees status (summary + per-type + matrix) first, then scrolls to the tools to act on it.
    wrap.appendChild(addFilesPanel());

    content.appendChild(wrap);
  }

  // ---- PM guide: a compact, collapsible one-pager at the top of the Admin tab. Reference material
  // (how the app works + how to add a product in the Smartsheet), so it's collapsed by default and
  // doesn't crowd the daily tools below. Source of truth for products/models is the Smartsheet. ----
  var FAMILY_TREE_URL = "/demo/elsewhere.html";
  function pmGuidePanel() {
    var d = document.createElement("details"); d.className = "wire pm-guide";
    d.innerHTML =
      '<summary class="wire-head">' + svg(ICONS.book) + '<span>How this works — PM guide</span></summary>' +
      '<div class="pmg-body">' +
        '<p class="wire-lead">Product Hub is a friendly front door to every product’s resources. Every link points to the real file in SharePoint — nothing is copied here. The <b>Product Family Tree</b> Smartsheet is the source of truth: edit it and the site updates within about a minute.</p>' +

        '<h2 class="pmg-h">The four tabs</h2>' +
        '<ul class="pmg-list">' +
          '<li><b>Find</b> — what everyone sees: browse to a product and open its files.</li>' +
          '<li><b>Progress</b> — PMs only: see which resources are linked and which still need work, and add or remove links.</li>' +
          '<li><b>Metrics</b> — PMs only: how staff use the catalog — adoption, searches that found nothing, and linked resources no one opens. Charts are aggregate; the usage log behind them drops each person’s identity after ~60 days.</li>' +
          '<li><b>Admin</b> — PMs only: this page — this guide and staff bug reports.</li>' +
        '</ul>' +

        '<h2 class="pmg-h">Add links to a product</h2>' +
        '<p class="pmg-p">On the <b>Progress</b> tab, open <b>Add files</b> and pick <b>A folder&nbsp;→&nbsp;a model</b> to point a resource at a SharePoint folder (new files in that folder then show up automatically), or click <b>Edit</b> on any product section to add or remove individual links. Each resource stays <b>Work in progress</b> until it has a link that opens to real files.</p>' +
        '<p class="pmg-p">You’ll also see a <b>New on SharePoint</b> list in the To&nbsp;Do box at the top of the Progress tab. Every night the site scans your linked folders for new files it can match to a model and <b>suggests</b> them there. Review each one and <b>Add</b> it or <b>Remove</b> it — nothing goes on the live site until you add it.</p>' +

        '<h2 class="pmg-h">Manage a file across many products</h2>' +
        '<p class="pmg-p">When one file belongs to several products, open <b>Add files</b> and pick <b>One file&nbsp;→&nbsp;many products</b> on the Progress tab (or the <b>Manage</b> button on any file chip or in a file preview). It shows <b>everywhere the file is used</b> and lets you add it to more models at once, move it, <b>replace it everywhere</b> with a newer version, or remove it everywhere. Nothing saves until you press <b>Approve&nbsp;&amp;&nbsp;save</b>.</p>' +

        '<h2 class="pmg-h">Add a missing product to the family tree</h2>' +
        '<p class="pmg-p">New products and models come from the Smartsheet, not the app itself. To add one:</p>' +
        '<ol class="wire-help-steps pmg-steps">' +
          '<li><a href="' + FAMILY_TREE_URL + '" target="_blank" rel="noopener">Open the Product Family Tree Smartsheet</a>.</li>' +
          '<li>Every model needs <b>ten rows</b> — one for each resource type: Flyer, Product Overview, Spec Sheet, Cut Sheet, Renders, In-Store Images, Assembly Instructions, Compliance &amp; Testing, Warranty, Training.</li>' +
          '<li>Easiest way: find a model like the one you’re adding, select its ten rows, copy them, then insert them as new rows.</li>' +
          '<li>On <b>every</b> new row fill <b>Category</b>, <b>Family</b>, <b>Product</b>, <b>Models/Variants</b>, and <b>RSS</b> (the resource type). Repeat Category / Family / Product on all ten rows — don’t leave them blank.</li>' +
          '<li>Leave <b>SP Folder URL</b> empty for now — you’ll add the links from the Progress tab.</li>' +
          '<li>Save. The product appears on the site within about a minute, with every resource marked <b>Work in progress</b> until you link it.</li>' +
        '</ol>' +
        '<p class="pmg-note">Tips: spell each resource type exactly as above (Flyer, Product Overview, Spec Sheet, Cut Sheet, Renders, In-Store Images, Assembly Instructions, Compliance &amp; Testing, Warranty, Training) so the app recognizes it. A brand-new <b>Family</b> or <b>Category</b> just means typing a new name in that column. To add a single variant to an existing product, copy just that product’s ten-row block and rename the model.</p>' +

        '<h2 class="pmg-h">Check that every link works</h2>' +
        '<p class="pmg-p">On the <b>Progress</b> tab, click <b>Check links</b> (in the coverage toolbar, just above the grid) to test every link. A link that won’t open — moved, deleted, or no access — is flagged as a <b>Link issue</b> (the red flag salespeople see on the Find tab); a folder with no files yet stays <b>Work in progress</b>. Neither counts toward a model’s progress. This check also runs <b>automatically every day</b>, so those flags show for everyone even when no one runs it — run it from Progress when you want to check right now.</p>' +
        '<p class="pmg-p">Two more states in the grid: a linked resource turns <b>Review</b> once its newest file is over a year old — open it, confirm it’s still current, or replace it. Mark a resource <b>N/A</b> (from a product section’s <b>Edit</b> view) when it genuinely doesn’t apply to that model — it’s then left out of the model’s progress instead of counting as missing.</p>' +
      '</div>';
    return d;
  }

  // ---- Admin: bug report inbox (PM-only) ----
  var STATUS_LABELS = { new: "New", in_progress: "In progress", resolved: "Resolved", wont_fix: "Won’t fix" };
  var STATUS_ORDER = ["new", "in_progress", "resolved", "wont_fix"];
  // PM-only data-quality flag: Product Family Tree rows whose Resource type isn't one of the canonical
  // resource types (a typo in the sheet's Resource column). Such a row is silently ignored — it appears NOWHERE on
  // the site — so surface it on Admin with the fix. Reuses the orphan-panel styling.
  // "Sheet hygiene": what in the Product Family Tree the app could not read as intended — each row says
  // what to change in the sheet. Rendered on Progress (where PMs work), only when there is something.
  function sheetHygienePanel() {
    var h = SHEET_HYGIENE, n = h.unknownResourceTypes.length + h.invalidLinks.length + h.dupeModelNames.length + h.strayLinks.length + h.dupeResourceRows.length;
    if (!n) return null;
    var box = document.createElement("details"); box.className = "wire orphan-files"; box.open = true;
    function group(title, lead, items, render) {
      if (!items.length) return "";
      return '<div class="hyg-group"><div class="hyg-h" role="heading" aria-level="3">' + title + ' <span class="uns-count">' + items.length + '</span></div><p class="wire-lead">' + lead + '</p>' +
        '<ul class="ui-list">' + items.slice(0, 100).map(render).join("") + '</ul>' + (items.length > 100 ? '<div class="wire-msg">+ ' + (items.length - 100) + ' more</div>' : '') + '</div>';
    }
    box.innerHTML =
      '<summary class="wire-head wire-warn-head">' + svg(ICONS.alert) + '<span>Sheet hygiene</span> <span class="uns-count">' + n + '</span></summary>' +
      '<div class="ui-body">' +
      group("Unrecognized resource types", 'These rows have a <b>Resource</b> value that isn’t one of the recognized types, so they don’t appear anywhere on the site. Fix the spelling in the sheet.',
        h.unknownResourceTypes, function (u) { return '<li class="ui-orphan"><span class="ui-q">' + escapeHtml(u.model || "(no model)") + '</span><span class="ui-sub">unrecognized resource type: “' + escapeHtml(u.rss || "") + '”</span></li>'; }) +
      group("Text where a link should be", 'A link cell holding a note counts as nothing — it isn’t a link and it isn’t N/A, so the resource shows as missing. Paste the file link, or mark the resource N/A.',
        h.invalidLinks, function (u) { return '<li class="ui-orphan"><span class="ui-q">' + escapeHtml(u.model || "") + ' · ' + escapeHtml(u.rss || "") + '</span><span class="ui-sub">“' + escapeHtml(u.text || "") + '”</span></li>'; }) +
      group("One model name under two products", 'Links are keyed by model name, so a name shared by two products merges their files into one set. Give one of them a distinct name.',
        h.dupeModelNames, function (u) { return '<li class="ui-orphan"><span class="ui-q">' + escapeHtml(u.model || "") + '</span><span class="ui-sub">' + escapeHtml((u.products || []).join(" · ")) + '</span></li>'; }) +
      group("Links the site can show but not change", 'These links are in the old <b>SP Folder URL 2</b> column, or on a row that isn’t indented under its resource row. The site shows them, but saving from here can’t update or remove them. Move each into <b>SP Folder URL</b> on its own row, indented under the resource.',
        h.strayLinks, function (u) { return '<li class="ui-orphan"><span class="ui-q">' + escapeHtml(u.model || "") + ' · ' + escapeHtml(u.rss || "") + '</span><span class="ui-sub">' + (u.why === "column" ? 'in SP Folder URL 2: ' : 'not indented: ') + escapeHtml(String(u.url || "").split("/").pop() || u.url || "") + '</span></li>'; }) +
      group("The same resource listed twice", 'A model has two rows for the same resource, so saves to it are refused — the site can’t tell which row to update. Delete the extra row.',
        h.dupeResourceRows, function (u) { return '<li class="ui-orphan"><span class="ui-q">' + escapeHtml(u.model || "") + ' · ' + escapeHtml(u.rss || "") + '</span><span class="ui-sub">' + escapeHtml(u.product || "") + ' · ' + u.count + ' rows</span></li>'; }) +
      '</div>';
    return box;
  }
  // Turn a search miss into an alias, inline on the Metrics miss-list. Replaces the row with a small form
  // ("<miss> → [what should this find?]"); Save POSTs to /api/aliases and applies the alias immediately
  // (pushes it into RUNTIME_ALIASES) so the fix works for this session without a reload.
  function openAliasForm(row) {
    if (!row) return;
    var q = row.getAttribute("data-q") || "";
    row._orig = row.innerHTML; // restore on cancel
    row.innerHTML = '<span class="met-q">' + escapeHtml(q) + '</span>' +
      '<span class="met-alias-form"><span class="met-arrow" aria-hidden="true">→</span>' +
      '<input class="met-alias-in" type="text" maxlength="60" placeholder="what should this find? e.g. spec sheet" aria-label="Alias target for ' + escapeAttr(q) + '">' +
      '<button class="met-alias-save" type="button">Save</button>' +
      '<button class="met-alias-cancel" type="button">Cancel</button></span>';
    var input = row.querySelector(".met-alias-in");
    try { input.focus(); } catch (e) {}
    input.addEventListener("input", function () { input.classList.remove("met-alias-bad"); input.removeAttribute("aria-invalid"); }); // clear the invalid state once they start correcting it
    row.querySelector(".met-alias-cancel").addEventListener("click", function () {
      if (row._orig != null) { row.innerHTML = row._orig; var b = row.querySelector(".met-alias-btn"); if (b) b.addEventListener("click", function () { openAliasForm(row); }); }
    });
    row.querySelector(".met-alias-save").addEventListener("click", function () { saveAlias(row, q, input.value); });
    input.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); saveAlias(row, q, input.value); } else if (e.key === "Escape") { var c = row.querySelector(".met-alias-cancel"); if (c) c.click(); } });
  }
  function saveAlias(row, from, target, onSaved) {
    var to = String(target || "").trim();
    var input = row.querySelector(".met-alias-in"), saveBtn = row.querySelector(".met-alias-save");
    // The To-do version of this form uses a <select>, not a text input, so flagging only ".met-alias-in"
    // meant a rejected alias (over 60 chars, or a target identical to the query) did NOTHING AT ALL —
    // no error, no saved row. Flag whichever control exists, and always say why.
    if (!(window.PF && PF.normAlias && PF.normAlias(from, to))) {
      var ctl = input || row.querySelector(".met-alias-sel") || row.querySelector("select");
      if (ctl) { try { ctl.focus(); } catch (e) {} ctl.classList.add("met-alias-bad"); ctl.setAttribute("aria-invalid", "true"); }
      showError('<span>' + (to ? 'That can’t point “' + escapeHtml(from) + '” at “' + escapeHtml(to) + '” — pick a different product (and keep it under 60 characters).' : 'Pick a product to point “' + escapeHtml(from) + '” at.') + '</span>', 5000);
      return;
    }
    var saveDone = busyButton(saveBtn, "Saving…");
    apiFetch("/api/aliases", { method: "POST", body: { from: from, to: to } })
      .then(function (j) {
        var na = (j && j.alias) || PF.normAlias(from, to);
        if (na && !RUNTIME_ALIASES.some(function (a) { return a.from === na.from && a.to === na.to; })) RUNTIME_ALIASES.push(na); // apply now
        row.innerHTML = '<span class="met-q">' + escapeHtml(from) + '</span><span class="met-sub met-alias-done">→ ' + escapeHtml(to) + (j && j.added === false ? ' · already an alias' : ' · alias added') + '</span>';
        showToast(svg(ICONS.sync) + '<span>“' + escapeHtml(from) + '” now finds “' + escapeHtml(to) + '”.</span>');
        if (onSaved) onSaved(); // let the To-Do demand tile drop this from its count (the query now resolves)
      })
      .catch(function (e) {
        saveDone();
        var m = String((e && e.message) || e); if (/^api \d+$/.test(m)) m = "Couldn’t save the alias.";
        showError('<span>' + escapeHtml(m) + '</span>', 4000);
      });
  }

  /* ============================================================
     PM · #styleguide
     ============================================================ */
  // ---- #styleguide (PM-only, no tab) ----------------------------------------------------------
  // Every token and component state on one page, READ FROM THE LIVE CSS rather than retyped — a
  // hand-written swatch sheet is wrong the day after someone edits a token. Contrast ratios are
  // computed here too, so "is this legible" is a number on the page rather than a guess.
  //
  // NOTE: no inline style attributes anywhere below. This app's CSP is `style-src 'self'` with no
  // 'unsafe-inline', so an inline style attribute is silently dropped and the swatch renders blank —
  // which for a page whose whole job is showing colour would be worse than useless. The dynamic bits
  // carry data-sg-* attributes and are applied through the CSSOM in sgPaint(), which CSP does allow.
  // (tests/csp.test.js enforces this; it is what caught the first draft.)
  function sgLum(cs) {
    var m = String(cs).match(/[0-9.]+/g); if (!m) return 0;
    var f = [m[0], m[1], m[2]].map(function (v) { v = parseFloat(v) / 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); });
    return 0.2126 * f[0] + 0.7152 * f[1] + 0.0722 * f[2];
  }
  function sgContrast(fg, bg) {
    var a = sgLum(fg), b = sgLum(bg), hi = Math.max(a, b), lo = Math.min(a, b);
    return Math.round(((hi + 0.05) / (lo + 0.05)) * 100) / 100;
  }
  function sgToken(name) { return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }
  function sgResolve(value) {
    var p = document.createElement("span"); p.style.setProperty("color", value); p.style.setProperty("display", "none");
    document.body.appendChild(p); var out = getComputedStyle(p).color; p.remove(); return out;
  }
  // Apply every data-sg-* to its element via the CSSOM (see the CSP note above).
  function sgPaint(root) {
    var map = { sgBg: "background", sgFg: "color", sgFs: "font-size", sgRadius: "border-radius", sgLs: "letter-spacing", sgFont: "font-family" };
    Array.prototype.forEach.call(root.querySelectorAll("[data-sg-bg],[data-sg-fg],[data-sg-fs],[data-sg-radius],[data-sg-ls],[data-sg-font]"), function (el) {
      for (var k in map) {
        if (!map.hasOwnProperty(k)) continue;
        var tok = el.dataset[k];
        if (tok) el.style.setProperty(map[k], k === "sgFont" ? "var(" + tok + ")" : sgToken(tok));
      }
    });
  }
  function renderStyleguide() {
    content.innerHTML = head("Product team", "Style guide", "Every token and component state the app uses, read from the live CSS — so this page cannot drift from what actually ships. Contrast ratios are measured here, not assumed. It lives at #styleguide and has no tab on purpose.");
    var wrap = document.createElement("div"); wrap.className = "view cov-view";

    var FACES = [
      { token: "--font-label", label: "Barlow Semi Condensed", use: "Uppercase labels, view tabs, category headings, chips", test: '700 17px "Barlow Semi Condensed"' },
      { token: "--font-ui", label: "Inter", use: "Body text, product names, buttons, descriptions", test: "600 17px Inter" },
      { token: "--font-mono", label: "JetBrains Mono", use: "Model SKUs, counts, the preview title", test: '700 17px "JetBrains Mono"' }
    ];
    var faceRows = FACES.map(function (f) {
      var ok = !!(document.fonts && document.fonts.check(f.test));
      return '<div class="sg-face">' +
        '<div class="sg-spec" data-sg-font="' + f.token + '">Hoardmaster with Riser · HMX 400 IR · 7/9</div>' +
        '<div class="sg-meta"><b>' + escapeHtml(f.label) + '</b> <code>' + escapeHtml(f.token) + '</code> ' +
        '<span class="pill ' + (ok ? "ok" : "bad") + '">' + (ok ? "Self-hosted ✓" : "NOT loaded") + '</span> ' +
        '<span class="sg-use">' + escapeHtml(f.use) + '</span></div></div>';
    }).join("");

    var SIZES = ["--fs-9-5", "--fs-10", "--fs-10-5", "--fs-11", "--fs-12", "--fs-12-5", "--fs-13", "--fs-13-5", "--fs-14", "--fs-15", "--fs-17", "--fs-22", "--fs-30"];
    var sizeRows = SIZES.map(function (t) {
      var v = sgToken(t); if (!v) return "";
      return '<div class="sg-row"><code>' + t + '</code><span class="sg-val">' + v + '</span>' +
        '<span data-sg-fs="' + t + '">The quick brown fox</span></div>';
    }).join("");

    var surface = sgResolve(sgToken("--surface"));
    var INK = ["--ink", "--ink-soft", "--ink-faint", "--accent-ink", "--ok-ink", "--pending", "--focus-ring"];
    var FILL = ["--bg", "--surface", "--surface-2", "--line", "--line-strong", "--accent", "--accent-tint", "--ok", "--ok-tint", "--pending-fill", "--pending-tint", "--focus-tint", "--on-ok", "--on-accent"];
    var inkRows = INK.map(function (t) {
      var v = sgToken(t), r = sgContrast(sgResolve(v), surface);
      var lvl = r >= 4.5 ? "ok" : r >= 3 ? "pending" : "bad";
      return '<div class="sg-row"><span class="sg-chip" data-sg-bg="' + t + '"></span><code>' + t + '</code>' +
        '<span class="sg-val">' + escapeHtml(v) + '</span><span data-sg-fg="' + t + '">Sample text</span>' +
        '<span class="pill ' + lvl + '">' + r.toFixed(2) + ':1</span></div>';
    }).join("");
    var fillRows = FILL.map(function (t) {
      return '<div class="sg-row"><span class="sg-chip" data-sg-bg="' + t + '"></span><code>' + t + '</code>' +
        '<span class="sg-val">' + escapeHtml(sgToken(t)) + '</span></div>';
    }).join("");

    var RADII = ["--r-xs", "--r-chip", "--r-sm", "--r-ctl", "--r-card", "--r-lg", "--r-pill"];
    var radRows = RADII.map(function (t) {
      return '<div class="sg-row"><span class="sg-swatch" data-sg-radius="' + t + '"></span><code>' + t + '</code>' +
        '<span class="sg-val">' + escapeHtml(sgToken(t)) + '</span></div>';
    }).join("");
    var TRACK = ["--ls-label-sm", "--ls-label", "--ls-label-lg"];
    var trackRows = TRACK.map(function (t) {
      return '<div class="sg-row"><code>' + t + '</code><span class="sg-val">' + escapeHtml(sgToken(t)) + '</span>' +
        '<span class="sg-track" data-sg-ls="' + t + '">Assembly instructions</span></div>';
    }).join("");

    var comps =
      '<div class="sg-set"><span class="sg-lbl">Buttons</span>' +
        '<button class="wire-btn primary" type="button">Primary</button>' +
        '<button class="wire-btn" type="button">Secondary</button>' +
        '<button class="wire-btn danger" type="button">Destructive</button>' +
        '<button class="head-btn" type="button">Head button</button></div>' +
      '<div class="sg-set"><span class="sg-lbl">Compact actions</span>' +
        '<button class="lb-copy" type="button">Copy link</button>' +
        '<button class="share-pack" type="button">Copy links</button>' +
        '<button class="res-req" type="button">Request this</button></div>' +
      '<div class="sg-set"><span class="sg-lbl">Status</span>' +
        '<span class="pill ok">Preview</span><span class="pill pending">Not in the Hub yet</span><span class="pill bad">Link issue</span>' +
        '<span class="lh-badge ok">OK</span><span class="lh-badge warn">Stale</span><span class="lh-badge bad">Broken</span>' +
        '<span class="age-chip">2d</span><span class="age-chip warn">9d</span><span class="age-chip bad">30d</span>' +
        '<span class="ftp-badge">7</span></div>' +
      '<div class="sg-set"><span class="sg-lbl">Segmented</span>' +
        '<span class="ft-seg"><button class="ft-seg-btn on" type="button">All</button>' +
        '<button class="ft-seg-btn" type="button">Has files</button>' +
        '<button class="ft-seg-btn" type="button">Not yet</button></span>' +
        '<button class="ft-chip on" type="button">Chip on</button><button class="ft-chip" type="button">Chip off</button></div>' +
      '<div class="sg-set sg-set-col"><span class="sg-lbl">Rows</span>' +
        '<button class="col-item" type="button"><span class="ci-name">Not selected</span><span class="ci-meta">5/9</span></button>' +
        '<button class="col-item sel" type="button"><span class="ci-name">Selected — neutral fill, teal bar</span><span class="ci-meta">7/9</span></button></div>';

    function sgSection(title, body, lead) {
      return '<div class="met-section"><div class="met-h" role="heading" aria-level="2">' + title + '</div>' +
        (lead ? '<p class="met-lead">' + lead + '</p>' : "") + body + '</div>';
    }
    wrap.innerHTML =
      sgSection("Typefaces", '<div class="sg-faces">' + faceRows + '</div>',
        "All three are downloaded with the app rather than looked up on the machine, so Windows, iPhone and Android render the same thing. The badge is a live check, not a claim.") +
      sgSection("Type scale", '<div class="sg-list">' + sizeRows + '</div>') +
      sgSection("Label tracking", '<div class="sg-list">' + trackRows + '</div>', "Smaller uppercase needs more air, not less.") +
      sgSection("Ink", '<div class="sg-list">' + inkRows + '</div>',
        "Measured against --surface. 4.5:1 is the AA floor for body text, 3:1 for large text and UI edges.") +
      sgSection("Fills", '<div class="sg-list">' + fillRows + '</div>') +
      sgSection("Corner radius", '<div class="sg-list">' + radRows + '</div>', "An action takes --r-ctl; a state takes --r-pill.") +
      sgSection("Components", '<div class="sg-comps">' + comps + '</div>',
        "The real classes, so this mirrors the app instead of describing it. Tab through them — every focus ring is --focus-ring.");
    sgPaint(wrap);
    content.appendChild(wrap);
  }

  /* ============================================================
     PM · METRICS
     ============================================================ */
  // Metrics view (PM-only): how staff use the catalog — adoption, unmatched searches to act on, and linked-
  // but-unopened resources. Aggregate only (the server returns counts, never identities). Consolidates the
  // usage signals that were previously a collapsible buried in Progress.
  function renderMetrics() {
    content.innerHTML = head("Product team", "Metrics", "How staff use the catalog — so you can improve search, prune dead links, and track adoption. The charts show aggregate counts only; the usage log behind them records who opened what and drops the identity after ~60 days.");
    var wrap = document.createElement("div"); wrap.className = "view";
    wrap.innerHTML =
      '<div class="met-section met-impact"><div class="met-h" role="heading" aria-level="2">Resource impact <small>coverage × adoption</small></div>' +
        '<p class="met-lead">Where linking effort meets real use — what to finish next, and what’s built but going unused.</p>' +
        '<div class="impact-overtime">Loading…</div>' +
        '<div class="impact-cols">' +
          '<div class="impact-col"><div class="impact-h" role="heading" aria-level="3">Fill these next <small>most-opened with gaps</small></div><div class="impact-fill">Loading…</div></div>' +
          '<div class="impact-col"><div class="impact-h" role="heading" aria-level="3">Built but quiet <small>covered, rarely opened</small></div><div class="impact-quiet">Loading…</div></div>' +
        '</div>' +
        '<div class="impact-types"><div class="impact-h" role="heading" aria-level="3">By resource type <small>coverage vs. share of opens</small></div><div class="impact-typebars">Loading…</div></div>' +
      '</div>' +
      '<div class="met-section"><div class="met-h" role="heading" aria-level="2">Adoption <small>last 30 days</small></div><div class="met-tiles">' +
        '<div class="met-tile"><div class="met-num" data-tile="opens">…</div><div class="met-tlabel">Resource opens</div><div class="met-tsub">files &amp; renders opened</div></div>' +
        '<div class="met-tile"><div class="met-num" data-tile="people">…</div><div class="met-tlabel">Active people</div><div class="met-tsub">distinct staff</div></div>' +
        '<div class="met-tile"><div class="met-num" data-tile="signins">…</div><div class="met-tlabel">Sign-ins</div><div class="met-tsub">one per person per day</div></div>' +
      '</div></div>' +
      '<div class="met-section"><div class="met-h" role="heading" aria-level="2">Weekly activity <small>last 8 weeks</small></div><p class="met-lead">Resource opens and distinct people per week — is adoption moving?</p><div class="met-trend">Loading…</div></div>' +
      '<div class="met-section"><div class="met-h" role="heading" aria-level="2">Most opened <small>last 30 days</small></div><p class="met-lead">What staff actually use most — the top-opened resources.</p><div class="met-top">Loading…</div></div>' +
      '<div class="met-section"><div class="met-h met-miss-h" role="heading" aria-level="2">Searches with no result</div><p class="met-lead">Turn these into search aliases or new content — what staff looked for and didn’t find.</p><div class="met-miss">Loading…</div></div>' +
      '<div class="met-section"><div class="met-h" role="heading" aria-level="2">Clicked while empty <small>last 60 days</small></div><p class="met-lead">Staff reached these resources and found nothing yet — a direct request to add them (stronger than a search guess).</p><div class="met-wip">Loading…</div></div>' +
      '<div class="met-section"><div class="met-h" role="heading" aria-level="2">Preview problems <small>last 30 days</small></div><p class="met-lead">Linked files that failed to open for someone (no access, or moved) — the daily link check runs app-only, so it can’t catch these per-person failures.</p><div class="met-pf">Loading…</div></div>' +
      '<div class="met-section"><div class="met-h met-never-h" role="heading" aria-level="2">Linked but not opened</div><p class="met-lead">Linked resources no one has opened lately — maybe not needed, or hard to find. Click one to go see it.</p><div class="met-never">Loading…</div></div>';
    content.appendChild(wrap);
    // Clear EVERY async panel, not just the first one. The no-token early return below used to leave the
    // other ten sections reading "Loading…" forever, which looks like a hung page rather than a sign-in
    // problem. Shared with the failure path so the two can't drift.
    function metricsUnavailable(msgHtml) {
      ["opens", "people", "signins"].forEach(function (n) { var el = wrap.querySelector('[data-tile="' + n + '"]'); if (el) el.textContent = "—"; });
      var miss = wrap.querySelector(".met-miss"); if (miss) miss.innerHTML = '<div class="wire-msg err">' + msgHtml + '</div>';
      var never = wrap.querySelector(".met-never"); if (never) never.innerHTML = "";
      ["met-trend", "met-top", "met-wip", "met-pf", "impact-overtime", "impact-fill", "impact-quiet", "impact-typebars"].forEach(function (c) {
        var el = wrap.querySelector("." + c); if (el) el.innerHTML = "";
      });
    }
    if (!window.PF_getApiToken) { metricsUnavailable("Please sign in again."); return; }
    usageInsights()
      .then(function (data) {
        fillResourceImpact(wrap, data);
        function tile(name, v) { var el = wrap.querySelector('[data-tile="' + name + '"]'); if (el) el.textContent = (v == null ? "—" : v); }
        tile("opens", data.opens); tile("people", data.activePeople); tile("signins", data.signins);
        // Weekly adoption trend — two bars per week (opens, distinct people), scaled to the busiest week.
        var trend = data.trend || [];
        var maxV = Math.max(1, Math.max.apply(null, trend.map(function (w) { return Math.max(w.opens, w.people); })));
        function shortWk(iso) { try { return new Date(iso + "T00:00:00Z").toLocaleDateString(undefined, { month: "short", day: "numeric", timeZone: "UTC" }); } catch (e) { return iso; } }
        var mt = wrap.querySelector(".met-trend");
        if (mt) mt.innerHTML = trend.some(function (w) { return w.opens || w.people; })
          ? '<div class="met-bars">' + trend.map(function (w) {
              return '<div class="met-bcol" title="' + shortWk(w.week) + ': ' + w.opens + ' opens · ' + w.people + ' people">' +
                '<div class="met-bwrap"><div class="met-bar opens" data-barh="' + Math.round(w.opens / maxV * 110) + '"></div>' +
                '<div class="met-bar people" data-barh="' + Math.round(w.people / maxV * 110) + '"></div></div>' +
                '<div class="met-blabel">' + shortWk(w.week) + '</div></div>';
            }).join("") + '</div>' +
            '<div class="met-legend2"><span class="met-lk"><span class="met-sw opens"></span>Opens</span><span class="met-lk"><span class="met-sw people"></span>People</span></div>'
          : '<div class="wire-msg">No activity recorded yet.</div>';
        // Most-opened resources (open counts).
        // Shared renderer for a "resource · rest ...  N×" count list (most-opened, clicked-while-empty, preview problems).
        function countList(sel, arr, emptyMsg) {
          var el = wrap.querySelector(sel); if (!el) return;
          el.innerHTML = (arr && arr.length)
            ? '<ul class="met-list">' + arr.map(function (t) {
                var parts = String(t.resource).split(" · ");
                return '<li class="met-miss-row"><span class="met-q mono">' + escapeHtml(parts[0] || t.resource) + '</span>' +
                  (parts.length > 1 ? '<span class="met-sub"> · ' + escapeHtml(parts.slice(1).join(" · ")) + '</span>' : '') +
                  (t.lastDate ? '<span class="met-when">last seen ' + escapeHtml(shortWk(t.lastDate)) + '</span>' : '') +
                  '<span class="met-count">' + t.count + '×</span></li>';
              }).join("") + '</ul>'
            : '<div class="wire-msg">' + emptyMsg + '</div>';
        }
        countList(".met-top", data.topOpened, "No opens recorded yet.");
        countList(".met-wip", data.topWip, "No empty-slot clicks recorded yet.");
        countList(".met-pf", data.previewFails, "No preview failures recorded.");
        wrap.querySelector(".met-miss-h").innerHTML = 'Searches with no result <small>last ' + (data.missDays || 60) + ' days</small>';
        var misses = data.topMisses || [];
        var missEl = wrap.querySelector(".met-miss");
        // This is the SAME list the To-do inbox shows, and the two used to disagree: To-do filtered out
        // queries already dismissed or since answered by an alias, Metrics showed everything with an
        // identical "+ Alias" button. A PM aliasing here re-did work already done there. Settled as
        // To-do = act, Metrics = aggregate history: every miss stays listed (that's the history), but a
        // handled one is marked and loses its action, so only genuinely open rows offer one.
        missEl.innerHTML = misses.length
          ? '<ul class="met-list">' + misses.map(function (m) {
              var handled = demandIsDone(m.query) || !queryStillUnmet(m.query);
              return '<li class="met-miss-row' + (handled ? ' handled' : '') + '" data-q="' + escapeAttr(m.query) + '"><span class="met-q">' + escapeHtml(m.query) + '</span>' +
                '<span class="met-miss-actions">' +
                (handled
                  ? '<span class="met-handled">handled</span>'
                  : '<button class="met-alias-btn" type="button" title="Turn this into a search synonym">+ Alias</button>') +
                '<span class="met-count">' + (m.count || 1) + '×</span></span></li>';
            }).join("") + '</ul>' +
            '<p class="met-lead met-xlink">Work through the open ones in <b>To&nbsp;do</b> on the <button class="linklike met-go-progress" type="button">Progress tab</button>, where each row can also be requested or marked handled.</p>'
          : '<div class="wire-msg">No unmatched searches in the last ' + (data.missDays || 60) + ' days.</div>';
        var goProg = missEl.querySelector(".met-go-progress");
        if (goProg) goProg.addEventListener("click", function () { var t = document.querySelector('.vtab[data-view="progress"]'); if (t) t.click(); });
        [].forEach.call(missEl.querySelectorAll(".met-alias-btn"), function (btn) {
          btn.addEventListener("click", function () { openAliasForm(btn.closest(".met-miss-row")); });
        });
        var opened = {}; (data.openedRecent || []).forEach(function (k) { opened[k] = 1; });
        var never = [], linkedTotal = 0;
        CATALOG.forEach(function (c) { c.families.forEach(function (f) { f.products.forEach(function (p) { (p.models || []).forEach(function (m) {
          RESOURCE_TYPES.forEach(function (rt) {
            if (!usableLinked(m, rt.name)) return; linkedTotal++;
            if (!opened[m + " · " + rt.name]) never.push({ model: m, resource: rt.name, prod: p.name, path: [c.name, f.name, p.name, m] });
          });
        }); }); }); });
        never.sort(function (a, b) { return (a.model + a.resource).localeCompare(b.model + b.resource); });
        wrap.querySelector(".met-never-h").innerHTML = 'Linked but not opened <small>' + never.length + ' of ' + linkedTotal + ' · last ' + (data.openDays || 30) + ' days</small>';
        var shown = never.slice(0, 40);
        wrap.querySelector(".met-never").innerHTML = never.length
          ? '<ul class="met-list">' + shown.map(function (n, i) { return '<li><button class="met-jump" type="button" data-i="' + i + '"><span class="met-q mono">' + escapeHtml(n.model) + '</span><span class="met-sub"> · ' + escapeHtml(n.resource) + ' · ' + escapeHtml(n.prod) + '</span></button></li>'; }).join("") + '</ul>' + (never.length > 40 ? '<div class="wire-msg">+ ' + (never.length - 40) + ' more</div>' : '')
          : '<div class="wire-msg">Every linked resource has been opened recently.</div>';
        [].forEach.call(wrap.querySelectorAll(".met-jump"), function (b) { b.addEventListener("click", function () { openTo(shown[+b.getAttribute("data-i")].path, true); }); });
      })
      .catch(function (e) {
        metricsUnavailable("Couldn’t load metrics (" + escapeHtml(String((e && e.message) || e)) + ").");
      });
  }
  // Fill the Resource Impact section: join current coverage (coverageStats) with recent opens
  // (opensByModel/opensByType) + the weekly coverage trend. All aggregate; no per-person data.
  function fillResourceImpact(wrap, data) {
    var s = coverageStats();
    var opensByModel = data.opensByModel || {}, opensByType = data.opensByType || {};
    var models = [];
    s.groups.forEach(function (g) {
      g.models.forEach(function (m) {
        models.push({ name: m.name, prod: g.prod, path: m.path, count: m.count, applicable: m.applicable, opens: opensByModel[m.name] || 0 });
      });
    });
    // "Fill these next": an unmet applicable slot AND recent opens — demand meeting a gap. Rank opens × gap.
    var fill = models.filter(function (m) { return m.applicable > m.count && m.opens > 0; })
      .map(function (m) { var gap = m.applicable - m.count; return { m: m, score: m.opens * gap }; })
      .sort(function (a, b) { return b.score - a.score || b.m.opens - a.m.opens; }).slice(0, 8);
    // "Built but quiet": well covered (≥80%) but no recent opens. Biggest investment first.
    var quiet = models.filter(function (m) { return m.applicable > 0 && m.count / m.applicable >= 0.8 && m.opens === 0; })
      .sort(function (a, b) { return b.applicable - a.applicable || b.count - a.count || a.name.localeCompare(b.name); }).slice(0, 8);

    var jumpStore = [];
    function impactRow(m, right) {
      var i = jumpStore.push(m.path) - 1;
      return '<li><button class="impact-jump" type="button" data-i="' + i + '">' +
        '<span class="impact-nm"><span class="impact-name mono">' + escapeHtml(m.name) + '</span><span class="impact-sub"> · ' + escapeHtml(m.prod) + '</span></span>' +
        '<span class="impact-metric">' + right + '</span></button></li>';
    }
    var fillEl = wrap.querySelector(".impact-fill");
    if (fillEl) fillEl.innerHTML = fill.length
      ? '<ul class="impact-list">' + fill.map(function (x) { return impactRow(x.m, '<b>' + x.m.opens + '</b> open' + (x.m.opens === 1 ? '' : 's') + ' · <span class="impact-gap">' + x.m.count + '/' + x.m.applicable + '</span>'); }).join("") + '</ul>'
      : '<div class="wire-msg">Every model people open is fully linked.</div>';
    var quietEl = wrap.querySelector(".impact-quiet");
    if (quietEl) quietEl.innerHTML = quiet.length
      ? '<ul class="impact-list">' + quiet.map(function (m) { return impactRow(m, '<span class="impact-gap">' + m.count + '/' + m.applicable + '</span> · <span class="impact-quiet-tag">0 opens</span>'); }).join("") + '</ul>'
      : '<div class="wire-msg">Every well-covered model has recent opens.</div>';
    [].forEach.call(wrap.querySelectorAll(".impact-jump"), function (b) {
      b.addEventListener("click", function () { var p = jumpStore[+b.getAttribute("data-i")]; if (p) openTo(p, true); });
    });

    // By resource type: coverage % (catalog) vs share of opens (usage).
    var totalTypeOpens = 0; Object.keys(opensByType).forEach(function (k) { totalTypeOpens += opensByType[k]; });
    var tb = wrap.querySelector(".impact-typebars");
    if (tb) tb.innerHTML = '<div class="impact-tgrid">' + RESOURCE_TYPES.map(function (rt) {
      var den = s.perTypeApplicable[rt.name], cov = den ? Math.floor(s.perType[rt.name] / den * 100) : 0;
      var opens = opensByType[rt.name] || 0, share = totalTypeOpens ? Math.round(opens / totalTypeOpens * 100) : 0;
      return '<div class="impact-trow">' +
        '<div class="impact-tname">' + escapeHtml(rt.name) + '</div>' +
        '<div class="impact-tbars">' +
          '<div class="impact-tline"><div class="impact-tbar"><div class="impact-tbar-fill cov" data-barw="' + cov + '"></div></div><span class="impact-tval">' + cov + '% covered</span></div>' +
          '<div class="impact-tline"><div class="impact-tbar"><div class="impact-tbar-fill use" data-barw="' + share + '"></div></div><span class="impact-tval">' + share + '% of opens</span></div>' +
        '</div>' +
      '</div>';
    }).join("") + '</div>';

    // Over time: coverage % + weekly opens. Empty until snapshots accrue → an honest note (adoption trend
    // is still shown by the Weekly activity section below).
    var ov = wrap.querySelector(".impact-overtime");
    if (ov) {
      var ct = data.coverageTrend || [];
      ov.innerHTML = (ct.length >= 2)
        // Honest sample-size caption: a 2-point line is thin; a 12-point one is robust — say which.
        ? impactChartSVG(ct, data.opensByWeek || {}) + '<div class="impact-snapcount">Based on ' + ct.length + ' weekly snapshots.</div>'
        // Still collecting: show progress toward the 2 needed to draw the line, so it never looks stuck.
        : '<div class="impact-building">Coverage history is building — <b>' + ct.length + ' of 2</b> weekly snapshots so far. A new one is written after each nightly link check, and the line appears once there are two. Adoption over time is in <b>Weekly activity</b> below.</div>';
    }
  }
  // Two-line SVG: coverage % (left, 0–100) and weekly opens (right, scaled). Zero-dependency + theme-aware
  // via CSS classes. Only drawn once ≥2 weekly coverage snapshots exist.
  // Two STACKED panels sharing one weekly timeline — coverage % on top, weekly opens (as bars) below —
  // instead of a dual-axis overlay. Two independent y-scales on one plot implied a "usage tracks coverage"
  // correlation the data can't support; separate panels keep each trend legible and compare each line only
  // to itself. x is spaced by actual week DATE (not array index), so a skipped weekly snapshot shows as a
  // real gap rather than silently distorting the slope.
  function impactChartSVG(cov, openByWeek) {
    openByWeek = openByWeek || {}; // { weekISO: opens } spanning the full coverage window (see opensByWeek)
    function wkMs(iso) { var t = Date.parse(iso + "T00:00:00Z"); return isNaN(t) ? 0 : t; }
    var pts = cov.map(function (c) { return { week: c.week, ms: wkMs(c.week), pct: c.pct, opens: openByWeek[c.week] || 0 }; });
    var n = pts.length;
    var W = 680, padL = 40, padR = 14, padTop = 20, gap = 30, hTop = 78, hBot = 54, padBot = 22;
    var topY0 = padTop, topY1 = padTop + hTop, botY0 = topY1 + gap, botY1 = botY0 + hBot, H = botY1 + padBot;
    var iw = W - padL - padR;
    var minW = pts[0].ms, maxW = pts[n - 1].ms, span = maxW - minW;
    function x(i) { return (n === 1 || span <= 0) ? padL + iw / 2 : padL + iw * (pts[i].ms - minW) / span; }
    var maxOpens = Math.max(1, Math.max.apply(null, pts.map(function (p) { return p.opens; })));
    function yCov(v) { return topY0 + hTop * (1 - v / 100); }
    function shortWk(iso) { try { return new Date(iso + "T00:00:00Z").toLocaleDateString(undefined, { month: "short", day: "numeric", timeZone: "UTC" }); } catch (e) { return iso; } }
    var covLine = pts.map(function (p, i) { return x(i).toFixed(1) + ',' + yCov(p.pct).toFixed(1); }).join(" ");
    var covDots = pts.map(function (p, i) { return '<circle cx="' + x(i).toFixed(1) + '" cy="' + yCov(p.pct).toFixed(1) + '" r="2.5" class="ic-cov-dot"/>'; }).join("");
    var barW = Math.max(3, Math.min(20, iw / Math.max(n, 6) * 0.5));
    var bars = pts.map(function (p, i) { var bh = hBot * p.opens / maxOpens; return '<rect x="' + (x(i) - barW / 2).toFixed(1) + '" y="' + (botY1 - bh).toFixed(1) + '" width="' + barW.toFixed(1) + '" height="' + Math.max(0, bh).toFixed(1) + '" rx="1.5" class="ic-bar"/>'; }).join("");
    var labels = pts.map(function (p, i) { return '<text x="' + x(i).toFixed(1) + '" y="' + (H - 6) + '" class="ic-xl" text-anchor="middle">' + escapeHtml(shortWk(p.week)) + '</text>'; }).join("");
    var slot = n > 1 ? iw / (n - 1) : iw;
    var barVals = slot < 26 ? "" : pts.map(function (p, i) {
      if (!p.opens) return ""; // a "0" floating over no bar is noise
      var bh = hBot * p.opens / maxOpens;
      return '<text x="' + x(i).toFixed(1) + '" y="' + (botY1 - bh - 4).toFixed(1) + '" class="ic-val" text-anchor="middle">' + p.opens + '</text>';
    }).join("");
    var midY = yCov(50).toFixed(1);
    // One transparent hover column per week spanning BOTH panels, carrying the exact numbers.
    var colW = n > 1 ? iw / (n - 1) : iw;
    var hits = pts.map(function (p, i) {
      var t = shortWk(p.week) + ' · ' + p.pct + '% coverage · ' + p.opens + ' open' + (p.opens === 1 ? '' : 's');
      return '<rect x="' + (x(i) - colW / 2).toFixed(1) + '" y="' + topY0 + '" width="' + colW.toFixed(1) + '" height="' + (botY1 - topY0) + '" fill="transparent"><title>' + escapeHtml(t) + '</title></rect>';
    }).join("");
    return '<svg class="impact-chart" viewBox="0 0 ' + W + ' ' + H + '" preserveAspectRatio="xMidYMid meet" role="img" aria-label="Two stacked charts sharing a weekly timeline: catalog coverage percent (0 to 100) on top, and weekly opens (0 to ' + maxOpens + ') below.">' +
      '<text x="' + padL + '" y="' + (topY0 - 7) + '" class="ic-panel-title">COVERAGE %</text>' +
      '<line x1="' + padL + '" y1="' + topY0 + '" x2="' + padL + '" y2="' + topY1 + '" class="ic-axis"/>' +
      '<line x1="' + padL + '" y1="' + topY1 + '" x2="' + (W - padR) + '" y2="' + topY1 + '" class="ic-axis"/>' +
      '<text x="4" y="' + (topY0 + 5) + '" class="ic-yl">100</text><text x="4" y="' + topY1 + '" class="ic-yl">0</text>' +
      '<line x1="' + padL + '" y1="' + midY + '" x2="' + (W - padR) + '" y2="' + midY + '" class="ic-grid"/>' +
      '<text x="4" y="' + (parseFloat(midY) + 3.5).toFixed(1) + '" class="ic-yl">50</text>' +
      '<polyline points="' + covLine + '" class="ic-cov"/>' + covDots +
      '<text x="' + padL + '" y="' + (botY0 - 7) + '" class="ic-panel-title">WEEKLY OPENS</text>' +
      '<line x1="' + padL + '" y1="' + botY0 + '" x2="' + padL + '" y2="' + botY1 + '" class="ic-axis"/>' +
      '<line x1="' + padL + '" y1="' + botY1 + '" x2="' + (W - padR) + '" y2="' + botY1 + '" class="ic-axis"/>' +
      '<text x="4" y="' + (botY0 + 5) + '" class="ic-yl">' + maxOpens + '</text><text x="4" y="' + botY1 + '" class="ic-yl">0</text>' +
      bars + barVals + labels + hits +
      '</svg>' +
      // The per-week numbers exist only inside hover <title>s, which a screen-reader user can't reach on a
      // role="img" SVG. Same data as the hover columns, as a real table (visually hidden, in the a11y tree).
      '<table class="sr-only"><caption>Catalog coverage and weekly opens by week</caption>' +
        '<thead><tr><th scope="col">Week of</th><th scope="col">Coverage</th><th scope="col">Opens</th></tr></thead><tbody>' +
        pts.map(function (p) {
          return '<tr><th scope="row">' + escapeHtml(shortWk(p.week)) + '</th>' +
            '<td>' + p.pct + '%</td><td>' + p.opens + '</td></tr>';
        }).join("") +
      '</tbody></table>' +
      '<div class="impact-legend"><span class="il"><span class="il-sw cov"></span>Coverage %</span><span class="il"><span class="il-sw use"></span>Weekly opens</span></div>';
  }

  /* ============================================================
     PM · ADMIN
     ============================================================ */
  function renderAdmin() {
    content.innerHTML = head("Product team", "Admin", "Your guide and staff bug reports.");
    var wrap = document.createElement("div"); wrap.className = "view";
    (PF_ADMIN_ALERTS || []).forEach(function (a) {
      var el = document.createElement("div"); el.className = "adm-alert " + (a.level === "error" ? "err" : "warn"); el.setAttribute("role", "alert");
      el.innerHTML = svg(ICONS.alert) + '<span>' + escapeHtml(a.message || "") + '</span>';
      wrap.appendChild(el);
    });
    wrap.appendChild(pmGuidePanel());
    // Link health moved to the "To do" inbox on Progress (its findings + the "Check all links" run live there now).
    var bugHead = document.createElement("div"); bugHead.className = "wire-head adm-bughead"; bugHead.innerHTML = svg(ICONS.spec) + '<span>Bug reports</span>';
    var bar = document.createElement("div"); bar.className = "adm-toolbar"; bar.textContent = "Loading reports…";
    var list = document.createElement("div"); list.className = "adm-list";
    wrap.appendChild(bugHead); wrap.appendChild(bar); wrap.appendChild(list);
    wrap.appendChild(favErasePanel()); // privacy utility, bottom of Admin
    content.appendChild(wrap);

    if (!window.PF_getToken) { bar.textContent = "Please sign in again."; return; }
    apiFetch("/api/bug-reports", { fresh: true }).then(function (data) {
      var reports = (data && data.reports) || [];
      // Re-derive the count line from the CURRENT data whenever a card changes, rather than stamping it
      // once at load: marking a report resolved (or deleting one) left "N reports · M open" stale until a
      // reload — and that number is exactly what a PM reads to decide whether the queue is clear.
      var live = reports.slice();
      function paintCount() {
        var open = live.filter(function (x) { return x.status === "new" || x.status === "in_progress"; }).length;
        bar.textContent = live.length + " report" + (live.length === 1 ? "" : "s") + " · " + open + " open";
      }
      paintCount();
      list.innerHTML = "";
      if (!reports.length) { list.innerHTML = '<div class="adm-empty">No bug reports yet.</div>'; return; }
      // Triage order: OPEN reports (new / in progress) first, oldest-first so the most-neglected leads;
      // resolved/won't-fix drop below, newest-first (recent history). Turns the flat list into a queue.
      var isOpenStatus = function (s) { return s === "new" || s === "in_progress"; };
      reports.sort(function (a, b) {
        var ao = isOpenStatus(a.status) ? 0 : 1, bo = isOpenStatus(b.status) ? 0 : 1;
        if (ao !== bo) return ao - bo;
        var at = Date.parse(a.created_at) || 0, bt = Date.parse(b.created_at) || 0;
        return ao === 0 ? at - bt : bt - at; // open: oldest-first (triage) · closed: newest-first (history)
      });
      reports.forEach(function (rep) {
        // Each card reports back so the header count stays true after a status change or a delete.
        rep._onStatus = function (removed) {
          if (removed) live = live.filter(function (x) { return x !== rep; });
          paintCount();
        };
        list.appendChild(adminCard(rep));
      });
    }).catch(function (err) {
      bar.textContent = (err && err.status === 403)
        ? "You don’t have access to this page." : "Couldn’t load reports — refresh to try again.";
    });
  }
  // PM privacy utility: erase EVERY favorite for one email (departed employee / right-to-erasure). Calls the
  // PM-gated POST /api/me {action:"erase-favorites", email}. Two-step inline confirm (arm → confirm within
  // 4s), mirroring the bug-report Delete, so a stray click can't wipe someone's pins.
  function favErasePanel() {
    var box = document.createElement("div"); box.className = "adm-erase";
    var h = document.createElement("div"); h.className = "wire-head"; h.innerHTML = svg(ICONS.shield) + '<span>Erase a user’s favorites</span>';
    var help = document.createElement("p"); help.className = "adm-erase-help";
    help.textContent = "Permanently delete every pinned model for one person — e.g. a departed employee, or a right-to-erasure request. It affects only that person and can’t be undone.";
    var row = document.createElement("div"); row.className = "adm-erase-row";
    var input = document.createElement("input"); input.type = "email"; input.className = "adm-erase-input"; input.placeholder = "name@grumbleton.example"; input.setAttribute("aria-label", "Email of the user whose favorites to erase");
    var btn = document.createElement("button"); btn.type = "button"; btn.className = "adm-erase-btn"; btn.textContent = "Erase favorites";
    var msg = document.createElement("div"); msg.className = "adm-erase-msg"; msg.setAttribute("role", "status"); msg.setAttribute("aria-live", "polite");
    var armed = false, disarmT = null;
    function disarm() { armed = false; btn.textContent = "Erase favorites"; btn.classList.remove("armed"); if (disarmT) { clearTimeout(disarmT); disarmT = null; } }
    input.addEventListener("input", function () { if (armed) disarm(); }); // editing the target cancels the arm
    btn.addEventListener("click", function () {
      var email = input.value.trim();
      if (!email) { msg.textContent = "Enter an email first."; msg.className = "adm-erase-msg err"; input.focus(); return; }
      if (!armed) { armed = true; btn.textContent = "Confirm — erase " + email + "?"; btn.classList.add("armed"); disarmT = setTimeout(disarm, 4000); return; }
      disarm();
      if (!window.PF_getApiToken) { msg.textContent = "Please sign in again."; msg.className = "adm-erase-msg err"; return; }
      btn.disabled = true; msg.textContent = "Erasing…"; msg.className = "adm-erase-msg";
      apiFetch("/api/me", { method: "POST", body: { action: "erase-favorites", email: email } }).then(function (d) {
        var n = (d && typeof d.removed === "number") ? d.removed : 0;
        msg.textContent = n === 0 ? ("No favorites found for " + email + " — nothing to erase.") : ("Erased " + n + " favorite" + (n === 1 ? "" : "s") + " for " + email + ".");
        msg.className = "adm-erase-msg ok"; input.value = "";
      }).catch(function (err) {
        msg.textContent = (err && err.status === 403) ? "You don’t have access to this." : "Couldn’t erase — please try again.";
        msg.className = "adm-erase-msg err";
      }).finally(function () { btn.disabled = false; });
    });
    row.appendChild(input); row.appendChild(btn);
    box.appendChild(h); box.appendChild(help); box.appendChild(row); box.appendChild(msg);
    return box;
  }
  function adminCard(rep) {
    var card = document.createElement("div"); card.className = "adm-card";
    // An unparseable date yields the literal string "Invalid Date" rather than throwing, so the catch
    // below never protected anything. Check the parse, and fall back to the raw value (or nothing).
    var when = rep.created_at || "";
    try { var d = new Date(rep.created_at); if (!isNaN(d.getTime())) when = d.toLocaleString(undefined, { year: "numeric", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }); } catch (e) {}
    var top = document.createElement("div"); top.className = "adm-top";
    // Age chip glows amber/red only while the report is still OPEN — a months-old resolved bug shouldn't alarm.
    var repOpen = rep.status === "new" || rep.status === "in_progress";
    top.innerHTML =
      '<span class="adm-who">' + escapeHtml(rep.reporter_name || prettyName(rep.reporter_email) || "Someone") + '</span>' +
      '<span class="adm-when">' + escapeHtml(when) + '</span>' +
      ageChip(rep.created_at, "", repOpen) +
      '<span class="adm-badge ' + escapeAttr(rep.status || "new") + '">' + escapeHtml(STATUS_LABELS[rep.status] || rep.status || "New") + '</span>';
    var statusWrap = document.createElement("span"); statusWrap.className = "adm-status";
    var dd = document.createElement("select"); dd.setAttribute("aria-label", "Change report status");
    STATUS_ORDER.forEach(function (s) {
      var o = document.createElement("option"); o.value = s; o.textContent = STATUS_LABELS[s];
      if (s === rep.status) o.selected = true; dd.appendChild(o);
    });
    // Keep the control, the badge and the server in step. On a failed PATCH the <select> used to KEEP the
    // new value while the badge and the sheet still held the old one — and re-picking the same option
    // fires no "change", so the PM couldn't even retry.
    dd.addEventListener("change", function () {
      var prev = rep.status || "new", next = dd.value;
      dd.disabled = true;
      updateReportStatus(rep.id, next, card).then(function () {
        rep.status = next;            // the card's own copy, so a later failure reverts to the truth
        if (typeof rep._onStatus === "function") rep._onStatus();  // refresh the "N reports · M open" line
      }, function () {
        dd.value = prev;              // put the control back where the data actually is
      }).then(function () { dd.disabled = false; });
    });
    statusWrap.appendChild(dd);
    // PM delete — permanently removes the report + its reporter PII. Two-step inline confirm (no native
    // dialog): first click arms "Confirm delete?", a second within 4s deletes; otherwise it re-disarms.
    var del = document.createElement("button"); del.type = "button"; del.className = "adm-del"; del.textContent = "Delete";
    del.setAttribute("aria-label", "Delete this report");
    var armed = false, disarmT = null;
    del.addEventListener("click", function () {
      if (!armed) { armed = true; del.textContent = "Confirm delete?"; del.classList.add("armed"); disarmT = setTimeout(function () { armed = false; del.textContent = "Delete"; del.classList.remove("armed"); }, 4000); return; }
      clearTimeout(disarmT);
      del.disabled = true; del.textContent = "Deleting…";
      deleteReport(rep.id, card).then(function () {
        if (typeof rep._onStatus === "function") rep._onStatus(true); // drop it from the count line
      }, function () {
        // Disarm on failure: it used to sit on "Confirm delete?" forever, one click from trying again.
        armed = false; del.disabled = false; del.textContent = "Delete"; del.classList.remove("armed");
      });
    });
    statusWrap.appendChild(del);
    top.appendChild(statusWrap); card.appendChild(top);

    var desc = document.createElement("div"); desc.className = "adm-desc"; desc.textContent = rep.description || "";
    card.appendChild(desc);

    var meta = document.createElement("div"); meta.className = "adm-meta"; var bits = [];
    if (rep.reporter_email) bits.push(escapeHtml(rep.reporter_email));
    if (rep.context) bits.push("On: " + escapeHtml(rep.context));
    if (rep.page_url) {
      // Route to the reporter's exact view IN THIS already-signed-in tab. Opening the raw
      // page_url in a NEW tab lands on the sign-in card — MSAL caches per-tab (sessionStorage),
      // so a fresh tab isn't authenticated. Linking to just the hash lets the in-app router
      // (hashchange → applyHash) restore that view with no reload and no re-auth.
      var repHash = "#";
      try { repHash = new URL(rep.page_url).hash || "#"; }
      catch (e) { var hi = rep.page_url.indexOf("#"); repHash = hi >= 0 ? rep.page_url.slice(hi) : "#"; }
      bits.push('<a href="' + escapeAttr(repHash) + '">the page they were on</a>');
    }
    meta.innerHTML = bits.join(" &nbsp;·&nbsp; ");
    card.appendChild(meta);
    return card;
  }
  // Both of these RETURN their promise: the caller reverts the <select> / re-enables the button on
  // failure and refreshes the header count on success, so swallowing the outcome here left the control
  // lying about the data — and returning nothing threw "Cannot read properties of undefined".
  function updateReportStatus(id, status, card) {
    if (!window.PF_getToken) return Promise.reject(new Error("Please sign in again."));
    return apiFetch("/api/bug-reports", { method: "PATCH", body: { id: id, status: status } }).then(function () {
      var badge = card.querySelector(".adm-badge");
      if (badge) { badge.className = "adm-badge " + status; badge.textContent = STATUS_LABELS[status] || status; }
      showToast(svg(ICONS.ext) + '<span>Updated to <b>' + (STATUS_LABELS[status] || status) + '</b></span>');
    }).catch(function (e) {
      showError(svg(ICONS.ext) + '<span>Couldn\'t update — try again.</span>');
      throw e; // the caller puts the <select> back where the data actually is
    });
  }
  function deleteReport(id, card) {
    if (!window.PF_getApiToken) return Promise.reject(new Error("Please sign in again."));
    return apiFetch("/api/bug-reports", { method: "DELETE", body: { id: id } }).then(function () {
      card.style.transition = "opacity .2s"; card.style.opacity = "0";
      setTimeout(function () { card.remove(); }, 200);
      showToast(svg(ICONS.ext) + '<span>Report deleted.</span>');
    }).catch(function (e) {
      showError('<span>Couldn\'t delete — try again.</span>');
      throw e; // the caller re-enables and disarms the button
    });
  }

  /* ============================================================
     SEARCH INDEX, ROUTER, DEEP LINKS & SHARE PACK
     ============================================================ */
  // ---- search ----
  var INDEX = [];
  function buildIndex() {
    INDEX.length = 0;
    CATALOG.forEach(function (c) {
      c.families.forEach(function (f) {
        INDEX.push({ lvl: "Family", name: f.name, mono: false, path: [c.name, f.name], trail: c.name });
        f.products.forEach(function (p) {
          INDEX.push({ lvl: "Product", name: p.name, mono: false, path: [c.name, f.name, p.name], trail: c.name + " › " + f.name });
          p.models.forEach(function (m) {
            // Include the DECODED model-code words (e.g. "HMX 400 T" → "Tall") in the searchable text, so
            // staff searching the attribute the app itself teaches ("tall", "integrated", "wide", "pullout")
            // actually find the model instead of a false search_miss.
            var dec = decodeModelName(m), extra = dec ? dec.map(function (d) { return d.label; }).join(" ") : "";
            // Also index each APPLICABLE resource-type name on the model row, so "<model> spec sheet" finds
            // the model even when that resource isn't LINKED yet — otherwise the model is hidden and the
            // searcher hits a dead "No match" for a product that plainly exists. Landing on the model shows
            // that resource as Work-in-progress with a Request button. This does NOT lose the PM demand
            // signal: an empty slot is already on the "Do this next" worklist (search_miss uniquely covers
            // things that AREN'T catalog slots). Both renderSearch and queryStillUnmet read `extra`, so
            // they stay in lock-step. N/A resources are excluded so a model isn't surfaced for a slot it
            // explicitly doesn't have.
            var rtypes = RESOURCE_TYPES.filter(function (r) { return !naFor(m, r.name); }).map(function (r) { return r.name; }).join(" ");
            extra = (extra + " " + rtypes).trim();
            INDEX.push({ lvl: "Model", name: m, mono: true, path: [c.name, f.name, p.name, m], trail: c.name + " › " + f.name + " › " + p.name, extra: extra });
            RESOURCE_TYPES.forEach(function (r) {
              var arr = linksFor(m, r.name);
              // Only index resources that actually open (skip empty/broken) — matches the What's New strip,
              // so search never advertises a resource that browse would flag as a Link issue. Lead the row
              // label with the MODEL (the distinguishing token), keeping the resource type for matching.
              if (arr.length && usableLinked(m, r.name)) INDEX.push({ lvl: "Live", name: m + " — " + r.name, mono: false, path: [c.name, f.name, p.name, m], trail: c.name + " › " + f.name + " › " + p.name, live: true, url: arr[0].url, model: m, resource: r.name });
            });
          });
        });
      });
    });
  }
  buildIndex();

  // `focus` = a user-initiated jump (search result, breadcrumb, coverage row, link-health "Go to"), so
  // move keyboard focus to the landed node after the innerHTML rebuild. The router path (applyHash: cold
  // load / deep link / back-forward) calls openTo WITHOUT focus, so a page load never steals focus to a
  // deep node — preserving the deliberate earlier decision not to grab focus on restore.
  function openTo(p, focus) {
    _preSearchView = null; // navigating into the catalog abandons any search-return memory
    if (p.length >= 2) openFams[keyOf([p[0], p[1]])] = true; // expand the target family (keep others open)
    sel.prod = p.length >= 3 ? keyOf([p[0], p[1], p[2]]) : null;
    sel.model = p.length >= 4 ? keyOf(p) : null;
    flashKey = keyOf(p); if (focus) pendingFocusKey = keyOf(p); searchInput.value = "";
    view = "catalog";
    render(); syncHash();
  }

  // Breadcrumb of the current selection (Category › Family › Product › Model). Category is context;
  // Family/Product are clickable (jump up via openTo); the deepest segment is the current page.
  function breadcrumbBar() {
    var key = sel.model || sel.prod;
    if (!key) return null;
    var parts = key.split("§");
    var bc = document.createElement("nav"); bc.className = "breadcrumb"; bc.setAttribute("aria-label", "Breadcrumb");
    parts.forEach(function (name, i) {
      if (i) { var sep = document.createElement("span"); sep.className = "bc-sep"; sep.setAttribute("aria-hidden", "true"); sep.textContent = "›"; bc.appendChild(sep); }
      var last = i === parts.length - 1;
      if (i === 0 || last) {
        var s = document.createElement("span"); s.className = "bc-cur" + (i === 0 ? " bc-cat" : ""); s.textContent = name;
        if (last) s.setAttribute("aria-current", "page");
        bc.appendChild(s);
      } else {
        var b = document.createElement("button"); b.type = "button"; b.className = "bc-link"; b.textContent = name;
        b.addEventListener("click", function () { openTo(parts.slice(0, i + 1), true); });
        bc.appendChild(b);
      }
    });
    return bc;
  }

  // ---- shareable deep links (hash router) ----
  // Every family / product / model gets its own bookmarkable URL, e.g.
  //   #/Hoardmaster/Hanging Hoardmaster/HMX 400I   (encoded)
  // The Progress view keeps its existing #progress hash.
  function currentPath() {
    var k = sel.model || sel.prod;
    return k ? k.split("§") : [];
  }
  function hashFor() {
    if (view === "progress") return "#progress";
    if (view === "metrics") return "#metrics";
    if (view === "admin") return "#admin";
    var p = currentPath();
    return p.length ? "#/" + p.map(encodeURIComponent).join("/") : "";
  }
  // Full [category, family, product, model] path for a model name (for resource share links).
  function findModelPath(modelName) {
    for (var ci = 0; ci < CATALOG.length; ci++) {
      var c = CATALOG[ci];
      for (var fi = 0; fi < c.families.length; fi++) {
        var f = c.families[fi];
        for (var pi = 0; pi < f.products.length; pi++) {
          var p = f.products[pi];
          if (p.models.indexOf(modelName) !== -1) return [c.name, f.name, p.name, modelName];
        }
      }
    }
    return null;
  }
  // Absolute URL that deep-links to a resource and auto-opens its preview on load.
  function resourceShareUrl(model, resource) {
    var path = findModelPath(model);
    if (!path) return location.href;
    if (resource) path = path.concat([resource]);
    return location.origin + location.pathname + location.search + "#/" + path.map(encodeURIComponent).join("/");
  }
  // ---- "Share pack": copy every in-app link for a model (or a whole product) to the clipboard, as BOTH
  // rich HTML (clickable labels in Outlook/Teams) and plain text ("Name: url"). Hub deep-links, not raw
  // SharePoint URLs, so they survive a file being re-linked/moved (each resolves to whatever's wired now).
  var PACKICON = DUPICON; // the same "two overlapping sheets" glyph: a copied link pack, like a duplicated chip, is "a copy of this"
  function modelResourceLinks(m) {
    // Only resources that actually open — skip N/A and anything the last health run found empty/broken.
    return RESOURCE_TYPES.filter(function (r) { return !naFor(m, r.name) && usableLinked(m, r.name); })
      .map(function (r) { return { name: r.name, url: resourceShareUrl(m, r.name) }; });
  }
  function modelSharePack(m) {
    var res = modelResourceLinks(m);
    return { kind: "model", title: m, resources: res, count: res.length, empty: res.length === 0 };
  }
  function sharePackContent(pack) {
    var suffix = " — Product Hub", link = function (r) { return '<a href="' + escapeAttr(r.url) + '">' + escapeHtml(r.name) + '</a>'; };
    // One shape only: a title + a flat list of {name, url} (a model's resources, or a gallery's picked files).
    return {
      html: "<b>" + escapeHtml(pack.title) + "</b>" + suffix + "<br>" + pack.resources.map(link).join(" &middot; "),
      text: pack.title + suffix + "\n" + pack.resources.map(function (r) { return r.name + ": " + r.url; }).join("\n")
    };
  }
  // Write BOTH text/html and text/plain: a paste into Outlook/Teams gets clickable labels, a paste into a
  // plain box gets "Name: url". Degrades gracefully: rich write → plain writeText → execCommand fallback.
  function copyRichText(html, text) {
    return new Promise(function (resolve, reject) {
      var plain = function () {
        if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(resolve, function () { if (fallbackCopy(text)) resolve(); else reject(); });
        else if (fallbackCopy(text)) resolve(); else reject();
      };
      try {
        if (navigator.clipboard && window.ClipboardItem) {
          var item = new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([text], { type: "text/plain" }) });
          navigator.clipboard.write([item]).then(resolve, plain);
        } else plain();
      } catch (e) { plain(); }
    });
  }
  function sharePackBtn(cls, label, aria, getPack) {
    var b = document.createElement("button"); b.type = "button"; b.className = "share-pack " + cls;
    b.innerHTML = PACKICON + '<span>' + escapeHtml(label) + '</span>'; b.setAttribute("aria-label", aria); b.title = aria;
    b.addEventListener("click", function (e) {
      e.stopPropagation(); // header/column shouldn't react to the copy click
      var pack = getPack(); if (!pack || pack.empty) { showToast('<span>No links to copy yet.</span>'); return; }
      var out = sharePackContent(pack);
      copyRichText(out.html, out.text).then(function () {
        b.classList.add("copied"); setTimeout(function () { b.classList.remove("copied"); }, 1400);
        // Same internal-only caveat as the single Copy-link hint (COPY_HINT_MSG): these are Grumbleton-gated
        // deep links, so a rep must NOT send them to a customer — flag it so they don't paste a dead link.
        showToast(PACKICON + '<span><b>' + pack.count + ' internal link' + (pack.count === 1 ? '' : 's') + ' copied</b> — they open only for Grumbleton sign-ins. Paste into an email or chat to a colleague.</span>', 6000);
      }, function () { showToast('<span>Couldn’t copy — try again.</span>'); });
    });
    return b;
  }
  var _internalHash = false; // guard so our own hash writes don't re-trigger a restore
  var _restoring = false;    // true while applyHash reflects existing state — normalize via replaceState, never push
  function syncHash() {
    var target = hashFor(), cur = location.hash || "";
    if (cur === target) return;
    try {
      // During a restore (cold load / back-forward / pasted link) we're REFLECTING state, not navigating —
      // REPLACE the entry, never push. Assigning location.hash pushed a 2nd entry, so normalizing a shared
      // resource link (#/…/M/R, which hashFor can't represent) down to #/…/M created a bounce: Back returned
      // to #/…/M/R, re-ran applyHash, re-opened the preview, and pushed again — trapping the user.
      if (_restoring) { history.replaceState(null, "", location.pathname + location.search + (target || "")); }
      // Belt-and-braces reset. The listener clears the guard when hashchange arrives, but hashchange only
      // fires if the hash ACTUALLY changed — and a browser that normalizes our value back to the current
      // one fires nothing, leaving the flag set to swallow the next GENUINE back/forward. A 0ms timer
      // would race the event and could clear the guard first, letting our own write re-trigger a restore
      // (the loop this flag exists to prevent); 50ms is comfortably after delivery and still far below
      // any real user gesture.
      else if (target) { _internalHash = true; location.hash = target; setTimeout(function () { _internalHash = false; }, 50); }
      else { history.replaceState(null, "", location.pathname + location.search); }
    } catch (e) { _internalHash = false; }
  }
  // Restore state from the current hash (initial load, back/forward, pasted link).
  // Does a deep-link path [cat, fam, prod?, model?] still exist in the catalog? Guards against a
  // shared link to something since renamed/removed (e.g. after a model rename) rendering empty columns.
  function catalogHasPath(p) {
    var c = CATALOG.filter(function (x) { return x.name === p[0]; })[0]; if (!c) return false;
    if (p.length < 2) return true;
    var f = (c.families || []).filter(function (x) { return x.name === p[1]; })[0]; if (!f) return false;
    if (p.length < 3) return true;
    var pr = (f.products || []).filter(function (x) { return x.name === p[2]; })[0]; if (!pr) return false;
    if (p.length < 4) return true;
    return (pr.models || []).indexOf(p[3]) !== -1;
  }
  function applyHash() {
    // Back/Forward with a preview open used to re-render the page UNDERNEATH the modal and leave it up —
    // closing it then revealed a view the user never chose. Close it before navigating.
    if (lb && !lb.hidden) closePreview();
    // Same for the other layers: a confirm is answered No, Manage Document closes WITHOUT reopening its
    // preview over the new view, and the bug modal closes (keeping its draft) (F334).
    if (pfConfirmEl && !pfConfirmEl.hidden) pfConfirmEl._done(false);
    if (docPanel && !docPanel.hidden) { if (docState) docState.onClose = null; closeDocPanel(); }
    if (bugModal && !bugModal.hidden) closeBug();
    _preSearchView = null; // any hash-driven navigation abandons the search-return memory
    _restoring = true; // reflect state via replaceState (see syncHash) so a shared resource link doesn't trap Back
    try {
    var h = location.hash || "";
    if (h === "#progress" || h === "#admin" || h === "#metrics" || h === "#styleguide") {
      if (!isPM()) { // a non-PM opened a shared PM link — don't just silently bounce them
        view = "catalog"; searchInput.value = ""; render(); resetScroll();
        syncHash(); // drop the rejected hash (replaceState while _restoring) so Back can't re-fire this
        showToast('<span>That’s a product-team page — showing the catalog instead.</span>');
        return;
      }
      view = h === "#admin" ? "admin" : h === "#metrics" ? "metrics" : h === "#styleguide" ? "styleguide" : "progress";
      sel = { prod: null, model: null }; // keep openFams so returning to Find keeps expansion
      searchInput.value = ""; render(); resetScroll(); return;
    }
    if (h.indexOf("#/") === 0) {
      var parts = h.slice(2).split("/").filter(Boolean).map(function (s) {
        try { return decodeURIComponent(s); } catch (e) { return s; }
      });
      if (parts.length) {
        if (!catalogHasPath(parts.slice(0, 4))) { // renamed/removed since the link was shared
          view = "catalog"; sel = { prod: null, model: null }; searchInput.value = ""; render(); resetScroll();
          syncHash(); // don't leave the dead deep link in the address bar to be copied or revisited
          showToast('<span>That link points to something no longer in the catalog — showing the catalog instead.</span>');
          return;
        }
        openTo(parts.slice(0, 4));           // restore the family/product/model selection
        if (parts.length >= 5) {             // 5th segment = a resource → auto-open its preview
          var model = parts[3], rname = parts[4], arr = linksFor(model, rname);
          var known = RESOURCE_TYPES.some(function (r) { return r.name === rname; });
          if (arr.length) openResource(arr, model + " · " + rname, model);
          // An UNKNOWN 5th segment is a typo'd or outdated link, not an unlinked resource. Saying
          // "isn't linked yet" sent the reader hunting for a missing file instead of checking the URL.
          else if (!known) showToast('<span>That link names a resource type that doesn’t exist (“' + escapeHtml(rname) + '”) — showing the model instead.</span>', 5000);
          else showToast('<span>That resource isn’t linked yet.</span>');
        }
        return;
      }
    }
    // root / unrecognized hash → catalog home
    view = "catalog"; sel = { prod: null, model: null }; openFams = {};
    searchInput.value = ""; render(); resetScroll();
    } finally { _restoring = false; }
  }

  // ---- lean usage telemetry (best-effort; backend = api/event.js + lib/usagesheet.js) ----
  var _missTimer = null, _loggedMisses = {};
  window.PF_logEvent = function (event, detail) {
    try {
      if (!window.PF_getApiToken) return;
      apiFetch("/api/event", { method: "POST", keepalive: true, body: { event: event, detail: detail || "" } }).catch(function () {});
    } catch (e) {}
  };

  /* ============================================================
     SEARCH RESULTS, HTML HELPERS & TOASTS
     ============================================================ */
  function renderSearch(q) {
    setDocTitle(String(q || "").trim()); // the tab should say what's being searched (WCAG 2.4.2)
    // The catalog (and its search INDEX) hasn't loaded yet: searching an empty index would render a false
    // "not in the catalog" AND log a bogus search_miss. Show a loading note instead; __startPF re-runs the
    // query for real once the index exists. (catalogError = the cold load truly failed → fall through.)
    if (!catalogLoaded && !catalogError) {
      var lw = document.createElement("div"); lw.className = "view";
      lw.innerHTML = '<div class="sr-head">Loading the catalog… your search will run in a moment.</div>';
      content.innerHTML = ""; content.appendChild(lw);
      return;
    }
    // The cold load truly failed — the INDEX is empty, so searching would render a false "not in the catalog"
    // for real products AND log a bogus search_miss that pollutes the demand list. Say so instead of searching.
    if (catalogError || !INDEX.length) {
      var ew = document.createElement("div"); ew.className = "view";
      ew.innerHTML = '<div class="sr-head">Couldn’t load the catalog right now — search is unavailable. Please refresh in a moment.</div>';
      content.innerHTML = ""; content.appendChild(ew);
      return;
    }
    var query = q.trim().toLowerCase();
    if (window.PF && PF.expandSearchQuery) query = PF.expandSearchQuery(query, RUNTIME_ALIASES); // fr→Hoardmaster, tobacco wall→goblet rack, + PM runtime aliases
    // Match every whitespace-separated token against the item name AND its path/trail, so
    // "Hoardmaster 400" or "hanging 400i" finds the model even though the name alone is "HMX 400I".
    var tokens = PF.tokenizeQuery(query);
    var hits = INDEX.filter(function (n) { return PF.searchMatch(tokens, n.name, (n.trail || "") + " " + (n.extra || ""), n.model || n.name); });
    hits.sort(function (a, b) {
      var rank = { Live: 0, Model: 1, Product: 2, Family: 2 };
      var am = rank[a.lvl], bm = rank[b.lvl];
      if (am !== bm) return am - bm;
      return a.name.length - b.name.length;
    });
    // Forgiving fallback: only when an EXACT match found nothing, retry typo-tolerantly ("assemby",
    // "prpusher", "hpp 40i"). Ranks by fewest typos, then the same level/length order. Keeps the fast
    // exact path unchanged for the common case, and means a rescued typo is NOT logged as a false miss.
    var fuzzy = false;
    if (hits.length === 0 && query.length >= 3 && window.PF && PF.fuzzyPenalty) {
      var scored = [];
      INDEX.forEach(function (n) {
        var pen = PF.fuzzyPenalty(query, n.name + " " + (n.trail || "") + " " + (n.extra || ""), n.model || n.name);
        if (pen >= 0) scored.push({ n: n, pen: pen });
      });
      if (scored.length) {
        var rank2 = { Live: 0, Model: 1, Product: 2, Family: 2 };
        scored.sort(function (a, b) {
          if (a.pen !== b.pen) return a.pen - b.pen;
          if (rank2[a.n.lvl] !== rank2[b.n.lvl]) return rank2[a.n.lvl] - rank2[b.n.lvl];
          return a.n.name.length - b.n.name.length;
        });
        hits = scored.map(function (s) { return s.n; });
        fuzzy = true;
      }
    }
    // Log a SETTLED search miss (query stuck ~1.5s with 0 results — even after the forgiving pass), once
    // per session — the highest-value "what should we add next" signal. Debounced so prefixes typed on the
    // way don't spam the log.
    // Don't record keyboard-mash / nonsense ("zzzqqxnotathing") as demand — it isn't a real "add this next"
    // signal and just clutters the PM worklist + Metrics list (also filtered server-side on read).
    if (hits.length === 0 && query.length >= 3 && !(window.PF && PF.looksLikeGibberish && PF.looksLikeGibberish(query))) {
      clearTimeout(_missTimer);
      _missTimer = setTimeout(function () { if (!_loggedMisses[query]) { _loggedMisses[query] = 1; window.PF_logEvent("search_miss", query); } }, 1500);
    } else { clearTimeout(_missTimer); }
    var wrap = document.createElement("div"); wrap.className = "view";
    var srN = hits.length; // count announced to screen readers (post-filter when a facet is active)
    if (hits.length === 0) {
      wrap.innerHTML = '<div class="sr-head">No match for “' + escapeHtml(q) + '” yet. Try the product’s name or one word from it (e.g. “hoard”, “boulder”).</div>' +
        '<div class="sr-empty-cta"><span class="sr-empty-hint">Can’t find it? Send a request — we’ll include what you searched for.</span>' +
        '<button class="head-btn" id="srReqBtn" type="button">' + REQICON + '<span>Request it</span></button></div>';
    } else {
      // Filter/sort toolbar over the model-bearing rows (Model + Live). Family/Product rows are
      // navigational — a facet like "Work in progress" isn't meaningful for them — so when any filter is
      // active we narrow to model rows; with no filter we keep every row and only (optionally) sort.
      var pool = hits.filter(function (n) { return n.lvl === "Model" || n.live; });
      var availCodes = PF.availableAttrs(pool, function (n) { return modelAttrs(n.model || n.name); }, MODEL_CODE_ORDER);
      searchFilter.attrs = searchFilter.attrs.filter(function (c) { return availCodes.indexOf(c) !== -1; }); // drop codes not in this result set
      var fActive = filterActive(searchFilter);
      var rows = PF.filterSortModels(fActive ? pool.slice() : hits.slice(), searchFilter, hitAccessors());
      srN = rows.length;
      // Header reports the TRUE count and flags truncation, so a model ranked 41st doesn't look absent.
      // On a forgiving (typo-tolerant) result set, say so — the user typed something we didn't match exactly.
      var head = document.createElement("div"); head.className = "sr-head";
      head.innerHTML = (fuzzy
        ? 'No exact match for “' + escapeHtml(q) + '” — showing the closest ' + (rows.length > 40 ? '40 of ' + rows.length : rows.length === 1 ? 'match' : rows.length + ' matches') + '.'
        : rows.length > 40
        ? 'Showing 40 of ' + rows.length + ' results for “' + escapeHtml(q) + '” — refine to narrow.'
        : fActive
        ? rows.length + ' result' + (rows.length === 1 ? '' : 's') + ' match “' + escapeHtml(q) + '”.'
        : rows.length + ' result' + (rows.length === 1 ? '' : 's') + ' for “' + escapeHtml(q) + '” — pick one to open it.');
      wrap.appendChild(head);
      if (pool.length >= TOOLBAR_MIN || fActive || searchFilter.sort) {
        wrap.appendChild(findToolbar(searchFilter, availCodes, { shown: fActive ? rows.length : pool.length, total: pool.length, noun: "results" }, "Best match", function () { renderSearch(q); }));
      }
      if (rows.length === 0) {
        var ne = document.createElement("div"); ne.className = "sr-empty-filter"; ne.textContent = "No results match these filters.";
        wrap.appendChild(ne);
      } else {
        var list = document.createElement("div"); list.className = "sr-list";
        rows.slice(0, 40).forEach(function (n, i) {
          var b = document.createElement("button"); b.type = "button"; b.className = "sr"; b.style.setProperty("--i", Math.min(i, 8));
          b.innerHTML = '<span class="sr-lvl">' + (n.live ? "File" : n.lvl) + '</span><span class="sr-main">' +
            '<span class="sr-name' + (n.mono ? ' mono' : '') + '">' + highlight(n.name, query) + decodeSuffixHtml(n.model || n.name) + '</span>' +
            (n.trail ? '<span class="sr-path">' + escapeHtml(n.trail) + '</span>' : '') + '</span>' +
            '<span class="ci-chev sr-chev' + (n.live ? ' live' : '') + '">' + (n.live ? svg(ICONS.ext) : CHEV) + '</span>';
          b.addEventListener("click", function () {
            // A "Live" hit now opens the SAME in-app gallery/preview as the resource column
            // (was window.open on the raw SharePoint URL — and only the first file of a multi-file resource).
            if (n.live && n.model) { openTo(n.path, true); openResource(linksFor(n.model, n.resource), n.model + " · " + n.resource, n.model); }
            else openTo(n.path, true);
          });
          list.appendChild(b);
        });
        // Keyboard nav: ↓/↑ move between results; ↑ from the first (or Esc) returns to the search box.
        list.addEventListener("keydown", function (e) {
          var items = Array.prototype.slice.call(list.querySelectorAll(".sr"));
          var i = items.indexOf(document.activeElement);
          if (e.key === "ArrowDown") { e.preventDefault(); (items[i + 1] || items[0]).focus(); }
          else if (e.key === "ArrowUp") { e.preventDefault(); if (i <= 0) searchInput.focus(); else items[i - 1].focus(); }
          else if (e.key === "Escape") { e.preventDefault(); searchInput.focus(); }
        });
        wrap.appendChild(list);
      }
    }
    content.innerHTML = "";
    content.appendChild(wrap);
    var srReqBtn = document.getElementById("srReqBtn");
    if (srReqBtn) srReqBtn.addEventListener("click", function () { openSearchRequest(q); });
    srAnnounce.textContent = srN + " result" + (srN === 1 ? "" : "s") + " for " + q;
  }

  // Highlight EACH token, not the whole query. Matching only the full string meant a perfectly good
  // two-word search ("Hoardmaster riser") highlighted nothing, because no single result contains the
  // words adjacently in that order. Tokens are matched longest-first so a longer one can't be broken up
  // by a shorter one that is a substring of it, and overlaps are merged before any markup is emitted.
  function highlight(text, q) {
    var raw = String(text == null ? "" : text);
    var toks = String(q || "").toLowerCase().split(/\s+/).filter(Boolean);
    if (!toks.length) return escapeHtml(raw);
    toks.sort(function (a, b) { return b.length - a.length; });
    var lower = raw.toLowerCase(), spans = [];
    toks.forEach(function (t) {
      for (var i = lower.indexOf(t); i !== -1; i = lower.indexOf(t, i + 1)) spans.push([i, i + t.length]);
    });
    if (!spans.length) return escapeHtml(raw);
    spans.sort(function (a, b) { return a[0] - b[0] || a[1] - b[1]; });
    var merged = [spans[0]];
    for (var j = 1; j < spans.length; j++) {
      var last = merged[merged.length - 1];
      if (spans[j][0] <= last[1]) last[1] = Math.max(last[1], spans[j][1]);
      else merged.push(spans[j]);
    }
    var out = "", at = 0;
    merged.forEach(function (sp) {
      out += escapeHtml(raw.slice(at, sp[0])) + '<mark class="hl">' + escapeHtml(raw.slice(sp[0], sp[1])) + '</mark>';
      at = sp[1];
    });
    return out + escapeHtml(raw.slice(at));
  }
  function escapeHtml(s) { return String(s).replace(/[&<>"']/g, function (c) { return ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]; }); }
  // ---- Age / triage helpers (bug inbox, placement backlog, unsorted pile) ----
  // Whole days since an ISO date; -1 if missing/unparseable so callers can skip the chip and sort it last.
  function ageDays(iso) { var t = iso ? Date.parse(iso) : NaN; return isNaN(t) ? -1 : Math.floor((Date.now() - t) / 86400000); }
  // Compact "how long ago" — glanceable, not precise: today / 5d / 3w / 2mo.
  function relAge(iso) {
    var d = ageDays(iso); if (d < 0) return "";
    if (d === 0) return "today";
    if (d < 14) return d + "d";
    if (d < 60) return Math.round(d / 7) + "w";
    return Math.round(d / 30) + "mo";
  }
  // A ready-to-inject age chip. prefix ("waiting"/"added") is optional; stale (default true) applies the
  // amber/red thresholds — pass false for closed/resolved rows that shouldn't glow.
  function ageChip(iso, prefix, stale) {
    var a = relAge(iso); if (!a) return "";
    var label = (prefix ? prefix + " " : "") + a;
    var cls = "", note = "";
    if (stale !== false) {
      var d = ageDays(iso);
      // The amber/red tint is the ONLY signal that a row is aging, which fails "not by colour alone".
      // Say it in text for AT; the tint stays the at-a-glance cue for everyone else.
      if (d >= 21) { cls = " bad"; note = ", needs attention"; }
      else if (d >= 7) { cls = " warn"; note = ", over a week old"; }
    }
    // title = the exact date (additive). Echoing the visible label here just double-announces it.
    var exact = ""; try { exact = new Date(iso).toLocaleDateString(); } catch (e) { exact = ""; }
    return '<span class="age-chip' + cls + '"' + (exact ? ' title="' + escapeAttr(exact) + '"' : '') + '>' +
      escapeHtml(label) + (note ? '<span class="sr-only">' + escapeHtml(note) + '</span>' : '') + '</span>';
  }
  function escapeAttr(s) { return escapeHtml(s); }
  // Only web/blob URLs may reach an href or window.open — blocks javascript:/data: from the sheet.
  function safeUrl(u) { var s = String(u == null ? "" : u).trim(); return /^(https?:|blob:)/i.test(s) ? s : ""; }
  // "first.last@grumbleton.example" → "First Last" (split the first.last local part, capitalize each).
  function prettyName(email) {
    var local = String(email || "").split("@")[0];
    if (!local) return "";
    return local.split(/[._-]+/).filter(Boolean).map(function (w) { return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase(); }).join(" ");
  }
  function cssEsc(s) { return s.replace(/["\\]/g, "\\$&"); }

  // ---- toast ----
  // How long a toast stays: long enough to READ it. A flat 2.6 s default gave a 35-word "nothing was saved"
  // message 2.6 s (F152). ~350 ms a word over a 1.2 s base, 2.6–15 s. An explicit `ms` is a floor, so a
  // caller can ask for longer (the 1-hour version nudge) but can't cut a long message short.
  function toastMs(t, ms) {
    var words = String(t.textContent || "").trim().split(/\s+/).filter(Boolean).length;
    return Math.max(ms || 0, Math.min(15000, Math.max(2600, 1200 + words * 350)));
  }
  // action (optional): { label, onClick } renders a button in the toast.
  // The toast-wrap is pointer-events:none so its empty space never blocks the page; the toast itself takes
  // clicks (app.css), so its links work and a click on it dismisses it early (F055/F149).
  function showToast(html, ms, action, assertive) {
    var t = document.createElement("div"); t.className = "toast";
    // Politeness comes from WHICH region we append into (#toastAlerts vs #toastStatus), not from a
    // role on the toast itself — a role="alert" nested in a role="status" wrap is ambiguous to AT.
    var region = (assertive ? toastAlerts : toastStatus) || toastWrap;
    t.innerHTML = html;
    var timer = null, left = 0, due = 0, gone = false;
    function dismiss() { if (gone) return; gone = true; clearTimeout(timer); timer = null; t.style.transition = "opacity .3s, transform .3s"; t.style.opacity = "0"; t.style.transform = "translateY(8px)"; setTimeout(function () { t.remove(); }, 320); }
    function run(n) { clearTimeout(timer); left = n; due = Date.now() + n; timer = setTimeout(dismiss, n); }
    // EVERY toast holds while the pointer or keyboard focus is on it (WCAG 2.2.1), then resumes with the time
    // it had LEFT. It used to be only toasts with a button, restarting at a flat 1.5 s — so a pointer merely
    // crossing the 1-hour "new version" nudge dismissed it for the hour (F062/F133).
    function pause() { if (gone || timer === null) return; clearTimeout(timer); timer = null; left = Math.max(0, due - Date.now()); }
    function resume() { if (gone || timer !== null) return; run(Math.max(1500, left)); }
    t.addEventListener("click", function (e) {
      // Not the action button (its own handler) and not a link: a failed "Download all" hands the files
      // back as SharePoint links here, and opening one must not take the list away.
      if (e.target && e.target.closest && e.target.closest(".toast-act, a")) return;
      dismiss();
    });
    t.addEventListener("mouseenter", pause); t.addEventListener("mouseleave", resume);
    t.addEventListener("focusin", pause); t.addEventListener("focusout", resume);
    if (action) {
      var btn = document.createElement("button"); btn.type = "button"; btn.className = "toast-act";
      btn.innerHTML = action.label;
      btn.addEventListener("click", function () { action.onClick(); dismiss(); });
      t.appendChild(btn);
    }
    region.appendChild(t);
    run(toastMs(t, ms));
  }
  function showError(html, ms) { showToast(html, ms, null, true); } // assertive (role=alert) for failures
  // The one "it saved" toast: every catalog write lands in Smartsheet at once and reaches every other
  // browser on its next catalog refresh — hence the fixed tail.
  function toastSaved(html, ms) { showToast(svg(ICONS.ext) + '<span>' + html + ' Demo — resets when you reload.</span>', ms); }

  /* ============================================================
     PREVIEW LIGHTBOX & GALLERIES
     ============================================================ */
  // ---- file preview (Microsoft Graph, as the signed-in user) ----
  // UTF-8 → base64url (Graph's documented sharing-URL encoding). Plain btoa() throws on any codepoint
  // > U+00FF, so a SharePoint URL with a non-ASCII path segment (or a stray smart-quote pasted into the
  // wire tool) would crash every preview/folder call; encodeURIComponent+unescape makes it byte-safe.
  function b64url(s) { return btoa(unescape(encodeURIComponent(s))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-"); }
  function shareId(url) { return "u!" + b64url(url); }
  // Parent-folder URL of a file link, so "Open in SharePoint" lands in the folder (its real
  // location in the SharePoint structure) rather than opening the bare file.
  function folderUrlOf(fileUrl) {
    try {
      var base = String(fileUrl).split("#")[0].split("?")[0].replace(/\/+$/, "");
      var i = base.lastIndexOf("/");
      return i > 8 ? base.slice(0, i) : String(fileUrl);
    } catch (e) { return fileUrl; }
  }
  // Microsoft embeddable preview URL for a document (PDF/Office) — used to show docs inline.
  function docPreviewUrl(fileUrl) {
    if (!window.PF_getToken) return Promise.reject(new Error("no token"));
    return window.PF_getToken().then(function (token) {
      return fetch("https://graph.microsoft.com/v1.0/shares/" + shareId(fileUrl) + "/driveItem/preview", {
        method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" }, body: "{}"
      });
    }).then(function (r) { if (!r.ok) throw new Error("preview " + r.status); return r.json(); })
      .then(function (p) { if (!p.getUrl) throw new Error("no preview url"); return p.getUrl; });
  }
  // A PDF blob for ANY document, so it can show in the browser's native viewer (no OneDrive chrome):
  // real PDFs stream from the pre-authenticated downloadUrl; Office docs (docx/pptx/xlsx) are
  // converted to PDF by Graph via ?format=pdf. Rejects → caller falls back to the MS embed viewer.
  // `signal` (optional AbortSignal) lets the caller cancel the fetch when the lightbox closes or pages
  // away — an Office→PDF conversion (?format=pdf) can take many seconds, so a closed preview shouldn't
  // keep it (and its bandwidth) running to completion.
  function docPdfBlob(fileUrl, name, downloadUrl, signal) {
    var isPdf = /\.pdf(\?|#|$)/i.test(name || "") || /\.pdf(\?|#|$)/i.test(fileUrl || "");
    if (isPdf && downloadUrl) return fetch(downloadUrl, { signal: signal }).then(function (r) { if (!r.ok) throw new Error("pdf " + r.status); return r.blob(); });
    if (!window.PF_getToken) return Promise.reject(new Error("no token"));
    return window.PF_getToken().then(function (token) {
      return fetch("https://graph.microsoft.com/v1.0/shares/" + shareId(fileUrl) + "/driveItem/content?format=pdf", { headers: { Authorization: "Bearer " + token }, signal: signal });
    }).then(function (r) { if (!r.ok) throw new Error("pdfconv " + r.status); return r.blob(); });
  }
  var lb = document.getElementById("lightbox");
  var galNav = null; // {prev, next} while a render gallery is open (for arrow keys)
  var lbShare = null; // {model, resource} the open lightbox is showing (for its Copy-link button)
  // Gallery select mode (T7-17): keys = item index in strip order. Download all / Share / Copy link read it.
  var lbSel = window.PF.makeSelection();
  var lbGalKeys = null, lbGalEntries = null; // the open gallery's indices, and {url, name} per item
  var lbSelCapNoted = false;                 // the "Share takes 9 at most" note, once per gallery
  var lbFile = null;  // {name, downloadUrl, fileUrl} of the file currently shown (for the Download button)
  // PDF open-parameter for the native viewer iframe (desktop path — phones use the MS viewer). FitH = fit to
  // WIDTH: the page fills the viewer width and scrolls vertically, so a dense doc is readable at a real size
  // instead of a whole-page "Fit" that shrinks it to an unreadable thumbnail on a short/wide window.
  function pdfViewHash() { return "#toolbar=0&view=FitH"; }
  var lbBlobUrl = null; // object URL for a PDF shown in the native viewer (revoke on close/switch)
  function lbClearBlob() { if (lbBlobUrl) { try { URL.revokeObjectURL(lbBlobUrl); } catch (e) {} lbBlobUrl = null; } }
  // Cancel an in-flight preview fetch (notably the slow Office→PDF conversion) when the lightbox closes
  // or pages to another view — so it doesn't keep running against a preview no one is looking at.
  var lbAbort = null;
  function lbAbortInflight() { if (lbAbort) { try { lbAbort.abort(); } catch (e) {} lbAbort = null; } }
  function lbFreshSignal() { lbAbortInflight(); lbAbort = (typeof AbortController !== "undefined") ? new AbortController() : null; return lbAbort ? lbAbort.signal : undefined; }
  (function () {
    if (!lb) return;
    lb.querySelector(".lb-close").addEventListener("click", closePreview);
    lb.querySelector(".lb-backdrop").addEventListener("click", closePreview);
    document.addEventListener("keydown", function (e) {
      if (lb.hidden) return;
      if (e.key === "Escape") closePreview();
      else if (galNav && e.key === "ArrowLeft") galNav.prev();
      else if (galNav && e.key === "ArrowRight") galNav.next();
    });
  })();
  // Accessible modal focus: trap Tab inside the open dialog and restore focus to the trigger on close.
  // A STACK, not one slot: the confirm dialog opens on top of Manage Document. With one slot the cached,
  // hidden confirm had already been made inert by the panel's open, so its Save/Cancel ignored every
  // click; and answering it tore down the panel's trap while the panel was still up (F002/F017). Each
  // layer remembers its own root, the focus to restore, and the nodes IT made inert.
  var a11yModal = (function () {
    var frames = [], active = null, onKey = null;
    function items(root) {
      return Array.prototype.slice.call(root.querySelectorAll(
        'a[href],button:not([disabled]),textarea:not([disabled]),input:not([disabled]),select:not([disabled]),iframe,[tabindex]:not([tabindex="-1"])'
      )).filter(function (el) {
        if (el.classList && el.classList.contains("focus-guard")) return false; // sentinels aren't content
        return el.offsetWidth || el.offsetHeight || el === document.activeElement;
      });
    }
    // Focus sentinels. The keydown trap below CANNOT contain a cross-origin iframe: once focus is inside
    // the Microsoft document viewer, Tab keydowns are delivered to ITS document, so our capture-phase
    // handler never runs and focus walks straight out of the dialog into the page (everything there is
    // inert, so it lands on <body> — nothing reachable, but the user is silently outside the dialog).
    // A sentinel is caught by a FOCUS event instead, which does fire in this document, so it works even
    // for the iframe case. No aria-hidden: a focusable aria-hidden element is itself an axe violation.
    var guardTop = null, guardEnd = null;
    function makeGuard(toEnd) {
      var g = document.createElement("div");
      g.className = "focus-guard"; g.tabIndex = 0;
      g.addEventListener("focus", function () {
        var list = items(active || document.body);
        if (!list.length) return;
        try { (toEnd ? list[list.length - 1] : list[0]).focus(); } catch (x) {}
      });
      return g;
    }
    function addGuards(panel) {
      if (!guardTop) { guardTop = makeGuard(true); guardEnd = makeGuard(false); }
      // Same nodes every time — insertBefore/appendChild MOVE them, so re-opening can't duplicate them.
      panel.insertBefore(guardTop, panel.firstChild);
      panel.appendChild(guardEnd);
    }
    function removeGuards() {
      [guardTop, guardEnd].forEach(function (g) { if (g && g.parentNode) g.parentNode.removeChild(g); });
    }
    // Make everything outside the open modal inert (unfocusable + hidden from assistive tech). This
    // also contains a cross-origin preview iframe: even if focus enters it, nothing behind is reachable.
    // Exempt: the toast stacks (#toastWrap, and #toastTopWrap, which holds the Copy-link "Copied" hint
    // that only ever shows inside the preview — F064) and #srAnnounce, which the gallery announces
    // "view 2 of 5" through. Nodes a lower layer already made inert are that layer's to clear.
    function setInert(frame) {
      frame.inerted = [];
      Array.prototype.forEach.call(document.body.children, function (c) {
        if (c === frame.root || c.id === "toastWrap" || c.id === "toastTopWrap" || c.id === "srAnnounce" || c.hasAttribute("inert")) return;
        c.setAttribute("inert", ""); frame.inerted.push(c);
      });
    }
    function modalRootOf(panel) { var n = panel; while (n && n.parentElement && n.parentElement !== document.body) n = n.parentElement; return n; }
    function indexOfRoot(root) { for (var i = frames.length - 1; i >= 0; i--) if (frames[i].root === root) return i; return -1; }
    function trap(e) {
      if (e.key !== "Tab" || !active) return;
      var list = items(active); if (!list.length) { e.preventDefault(); return; }
      var first = list[0], last = list[list.length - 1];
      if (!active.contains(document.activeElement)) { e.preventDefault(); first.focus(); }
      else if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
    }
    function seat(panel) {
      active = panel;
      addGuards(panel); // re-seat them each open: a resource switch rebuilds the panel's body
      if (!onKey) { onKey = trap; document.addEventListener("keydown", onKey, true); }
    }
    function open(panel, initial) {
      if (!panel) return;
      var root = modalRootOf(panel), i = indexOfRoot(root);
      if (i === -1) {
        // A new layer. Its own root may have been made inert by the layer beneath (the confirm is a cached
        // body child like any other) — lift that, and put it back when this layer closes.
        var f = { root: root, panel: panel, prevFocus: document.activeElement, inerted: [], wasInert: root.hasAttribute("inert") };
        if (f.wasInert) root.removeAttribute("inert");
        setInert(f);
        frames.push(f);
        seat(panel);
      } else {
        frames[i].panel = panel; // the same dialog re-opening (a resource switch) — keep its state
        if (i === frames.length - 1) seat(panel);
      }
      setTimeout(function () { try { (initial || items(panel)[0] || panel).focus(); } catch (x) {} }, 0);
    }
    // close(panel) closes THAT dialog; a close for one that isn't open is a no-op, so a late close (a save
    // landing after the PM already dismissed its panel) can't strip another dialog's trap (F185).
    function close(panel) {
      var i = panel ? indexOfRoot(modalRootOf(panel)) : frames.length - 1;
      if (i < 0) return;
      var f = frames[i], above = frames[i + 1];
      frames.splice(i, 1);
      if (above) {
        // Closing a layer that is NOT on top (a route change drops the panel under its confirm): the page
        // must stay inert while the layer above is open, so that layer inherits this one's inert set —
        // plus this root, now hidden — and its focus-return falls through to where this layer began.
        f.inerted.forEach(function (c) { if (c !== above.root) above.inerted.push(c); });
        if (above.wasInert) above.wasInert = false; // the layer that inerted it is gone
        if (!f.root.hasAttribute("inert")) { f.root.setAttribute("inert", ""); above.inerted.push(f.root); }
        if (!above.prevFocus || f.root.contains(above.prevFocus)) above.prevFocus = f.prevFocus;
        return;
      }
      f.inerted.forEach(function (c) { c.removeAttribute("inert"); });
      if (f.wasInert) f.root.setAttribute("inert", ""); // back to how the layer beneath left it
      removeGuards();
      if (frames.length) seat(frames[frames.length - 1].panel);
      else { if (onKey) document.removeEventListener("keydown", onKey, true); onKey = null; active = null; }
      if (f.prevFocus && f.prevFocus.focus) { try { f.prevFocus.focus(); } catch (x) {} }
    }
    return { open: open, close: close };
  })();
  // ---- confirm dialog ----
  // Replaces window.confirm() for the destructive wire-tool actions. Native confirm is unstyled, ignores
  // the theme, cannot show structure (these messages carry a list of affected models), blocks the whole
  // renderer, and cannot be driven by the e2e suite. This is the same shell as the bug modal, so it
  // inherits the focus trap, the Escape handling and the inert-background treatment from a11yModal.
  //
  // Returns a Promise<boolean>. It NEVER rejects: a caller's `if (!ok) return;` is the only branch it has
  // to handle, exactly like the `if (!confirm(...)) return;` it replaces.
  var pfConfirmEl = null, pfConfirmResolve = null;
  function pfConfirm(opts) {
    opts = opts || {};
    return new Promise(function (resolve) {
      // Only one at a time: a second call while one is open answers `false` for the first rather than
      // stranding its promise forever (an un-resolved promise here would silently drop a PM's action).
      if (pfConfirmResolve) { var prev = pfConfirmResolve; pfConfirmResolve = null; prev(false); }
      if (!pfConfirmEl) {
        pfConfirmEl = document.createElement("div");
        pfConfirmEl.className = "bugm pfc";
        pfConfirmEl.innerHTML =
          '<div class="bugm-backdrop pfc-backdrop"></div>' +
          '<div class="bugm-panel pfc-panel" role="dialog" aria-modal="true" aria-labelledby="pfcTitle" aria-describedby="pfcBody">' +
            '<div class="bugm-head"><span class="bugm-title" id="pfcTitle"></span>' +
              '<button class="bugm-close pfc-x" type="button" aria-label="Cancel">\u2715</button></div>' +
            '<div class="bugm-body"><div class="pfc-lead" id="pfcBody"></div>' +
              '<div class="bugm-actions"><button class="bugm-cancel pfc-no" type="button">Cancel</button>' +
              '<button class="bugm-submit pfc-yes" type="button"></button></div></div>' +
          '</div>';
        document.body.appendChild(pfConfirmEl);
        var done = function (ok) {
          var r = pfConfirmResolve; pfConfirmResolve = null;
          pfConfirmEl.hidden = true; a11yModal.close(pfConfirmEl.querySelector(".pfc-panel"));
          if (r) r(ok);
        };
        pfConfirmEl.querySelector(".pfc-yes").addEventListener("click", function () { done(true); });
        pfConfirmEl.querySelector(".pfc-no").addEventListener("click", function () { done(false); });
        pfConfirmEl.querySelector(".pfc-x").addEventListener("click", function () { done(false); });
        pfConfirmEl.querySelector(".pfc-backdrop").addEventListener("click", function () { done(false); });
        // Escape cancels, like the native dialog. a11yModal traps Tab; it does not close for us.
        pfConfirmEl.addEventListener("keydown", function (e) { if (e.key === "Escape") { e.stopPropagation(); done(false); } });
        pfConfirmEl._done = done;
      }
      pfConfirmResolve = resolve;
      pfConfirmEl.querySelector("#pfcTitle").textContent = opts.title || "Are you sure?";
      // `body` may be an array of paragraphs. Text only \u2014 never innerHTML: every one of these messages
      // interpolates a file or model name that came from the sheet.
      var bodyEl = pfConfirmEl.querySelector("#pfcBody");
      bodyEl.innerHTML = "";
      [].concat(opts.body || []).forEach(function (para) {
        var pEl = document.createElement("p"); pEl.className = "pfc-p"; pEl.textContent = String(para); bodyEl.appendChild(pEl);
      });
      var yes = pfConfirmEl.querySelector(".pfc-yes");
      yes.textContent = opts.confirmLabel || "Continue";
      yes.classList.toggle("pfc-danger", !!opts.danger);
      pfConfirmEl.hidden = false;
      // Focus the SAFE choice by default. A destructive action should never be one stray Enter away.
      a11yModal.open(pfConfirmEl.querySelector(".pfc-panel"), pfConfirmEl.querySelector(".pfc-no"));
    });
  }
  // `cond ? ask : straight through` \u2014 mirrors the `if (cond && !confirm(...)) return;` shape these replaced,
  // so a call site keeps ONE code path instead of branching around the dialog.
  function pfConfirmIf(cond, opts) { return cond ? pfConfirm(opts) : Promise.resolve(true); }

  function closePreview() { if (lb) { lb.hidden = true; lbSel.stop(); lbAbortInflight(); setDlBusy(false); setShareBusy(false); /* a hung download/share must not leave its button disabled into the NEXT preview */ var b = lb.querySelector(".lb-body"); b.innerHTML = ""; b.classList.remove("doc"); lbClearBlob(); galNav = null; setDownloadAll(null); clearCopyHint(); lbMoreClose(false); lbShare = null; lbFile = null; /* the next preview must not inherit them (F059) */ a11yModal.close(lb.querySelector(".lb-panel")); } }
  // Turn a technical fetch error (e.g. "Graph 503", "preview 500") into plain language for the staff-facing
  // preview UI. The raw message still goes to console.error for debugging — staff never see a bare HTTP code.
  function friendlyErr(err) {
    var m = (err && err.message) ? String(err.message) : String(err || "");
    var code = (m.match(/\b(\d{3})\b/) || [])[1];
    if (code === "401" || /sign in/i.test(m)) return "Please sign in again to open this.";
    if (code === "403") return "You may not have permission to open this file.";
    if (code === "404") return "This file couldn’t be found — it may have been moved, renamed, or deleted.";
    if (code === "429" || (code && code.charAt(0) === "5")) return "The file service is busy right now — please try again in a moment.";
    if (/timed out|timeout/i.test(m)) return "This is taking too long to load — try again, or open it in SharePoint.";
    return "Something went wrong opening this. Try again, or open it in SharePoint.";
  }
  function lbError(url, msg) {
    if (!lb) return;
    lb.querySelector(".lb-body").innerHTML = '<div class="lb-err"><p>' + escapeHtml(msg) + '</p><a class="lb-openbtn" href="' + escapeAttr(url) + '" target="_blank" rel="noopener">Open in SharePoint →</a></div>';
  }
  // A per-user 403/404 opening a WIRED file is invisible to the daily link-health cron (which checks
  // app-only): the file exists and the cron can read it, but THIS user can't (no SharePoint access) or it
  // moved for them. Log those as demand/diagnostic so PMs can see access gaps the cron can't. Only 403/404
  // (a real access/not-found), keyed to the resource the lightbox is showing.
  function logPreviewFail(err) {
    var code = (String((err && err.message) || err).match(/\b(40[34])\b/) || [])[1];
    if (code && lbShare && lbShare.model) window.PF_logEvent && window.PF_logEvent("preview_fail", lbShare.model + (lbShare.resource ? " · " + lbShare.resource : "") + " · " + code);
  }
  // Show the resource's "Updated …" currency in the preview header bar. The lightbox title is
  // "Model · Resource", so derive the model+resource from it and look up its added/updated date.
  function setLbUpdated(title) {
    var el = document.getElementById("lbUpdated"); if (!el) return;
    var parts = String(title || "").split(" · "), upd = "", stale = false;
    if (parts.length >= 2) { var r = parts.pop(), m = parts.join(" · "); upd = resUpdatedLabel(m, r); stale = !!(upd && resIsStale(m, r)); }
    el.textContent = upd ? (upd + (stale ? " · review" : "")) : "";
    el.className = "lb-updated" + (stale ? " stale" : "");
  }
  // Common prologue of every lightbox opener: cancel the previous view's in-flight work, set the title and
  // its "updated" note, point Open-in-SharePoint at `openHref`, reset the body (incl. the doc-mode class a
  // PDF leaves behind) to `loadingHtml`, show the dialog with its focus trap, and bump the sequence
  // number so a superseded open's async resolves are ignored. Returns that sequence number. Four openers
  // each carried their own copy, and two of them had drifted (T1-30, T1-34).
  // 15 s watchdog for a stalled open: if the opener's INITIAL loader is still up, show the error with Open in
  // SharePoint. Only that loader — lbBegin marks it .lb-loading-init. Matching any .lb-loading also caught the
  // gallery's per-view loader and the Microsoft viewer's overlay, and wiped a working gallery (picks and all)
  // when the PM happened to page at the 15 s mark (F005). A newer open or a close bumps/hides lb, so a stale
  // watchdog no-ops.
  function lbWatchdog(_seq, url) {
    setTimeout(function () { if (lb._seq === _seq && !lb.hidden && lb.querySelector(".lb-body .lb-loading-init")) lbError(url, "This is taking a while to load. Try opening it in SharePoint."); }, 15000);
  }
  function lbBegin(title, fallbackTitle, openHref, loadingHtml) {
    lbSel.stop(); lbSelCapNoted = false; lbGalKeys = null; lbGalEntries = null; // no picks carry into the next preview
    lbAbortInflight(); lbClearBlob(); // cancel a prior view's in-flight Office→PDF conversion + revoke its blob
    lb.querySelector(".lb-title").textContent = title || fallbackTitle; setLbUpdated(title);
    lb.querySelector(".lb-open").href = openHref;
    var body = lb.querySelector(".lb-body"); body.classList.remove("doc"); body.innerHTML = loadingHtml;
    var first = body.firstElementChild; if (first && first.classList.contains("lb-loading")) first.classList.add("lb-loading-init");
    lbFile = null; // until this open resolves a file, Download must not grab the previous one (F059)
    lb.hidden = false; a11yModal.open(lb.querySelector(".lb-panel"), lb.querySelector(".lb-close"));
    lb._seq = (lb._seq || 0) + 1; return lb._seq;
  }
  function openPreview(url, title) {
    if (!lb) { window.open(url, "_blank", "noopener"); return; }
    setDownloadAll(null); // single-file preview — no "Download all"
    // Open-in-SharePoint goes to the containing folder, not the bare file.
    var _seq = lbBegin(title, "Preview", folderUrlOf(url), '<div class="lb-loading"><span class="lb-skel" aria-hidden="true"></span><span>Loading preview…</span></div>');
    // Watchdog: if a Graph fetch stalls (slow VPN/proxy, a file that never streams), the spinner would sit
    // forever. After 15s, if this same open is still showing "Loading…", fall back to the error UI (which
    // offers "Open in SharePoint"). A newer open or a close bumps/hides lb, so a stale watchdog no-ops.
    var lbSig = lbFreshSignal(); // cancels any prior in-flight conversion; aborted on close/switch
    lbWatchdog(_seq, url);
    if (!window.PF_getToken) { lbError(url, "Please sign in again to preview files."); return; }
    lbFile = { name: null, downloadUrl: null, fileUrl: url };
    var bearer;
    var isImgUrl = /\.(png|jpe?g|gif|webp|bmp|svg)(\?|#|$)/i.test(url);
    window.PF_getToken().then(function (token) {
      bearer = "Bearer " + token;
      return graphJson("https://graph.microsoft.com/v1.0/shares/" + shareId(url) + "/driveItem", bearer);
    }).then(function (item) {
      if (lb.hidden || lb._seq !== _seq) return; // closed, or a newer open superseded this response
      var mime = (item.file && item.file.mimeType) || "";
      var dl = item["@microsoft.graph.downloadUrl"] || item["@content.downloadUrl"];
      lbFile = { name: item.name || null, downloadUrl: dl || null, fileUrl: url };
      var isImg = /^image\//.test(mime) || isImgUrl;
      var bodyEl = lb.querySelector(".lb-body");
      if (isImg && dl) {
        // images render cleanly straight from the download URL (no download-disposition problem)
        // Keep "Loading preview…" up until the (often multi-MB) image actually paints, then swap it in —
        // otherwise the stage flashes blank on every open. The doc branch already does this.
        bodyEl.innerHTML = '<div class="lb-loading lb-loading-init"><span class="lb-skel" aria-hidden="true"></span><span>Loading preview…</span></div>';
        var img = document.createElement("img"); img.className = "lb-img"; img.alt = item.name || ""; img.style.display = "none";
        img.onload = function () {
          if (lb._seq !== _seq || lb.hidden) return;
          // The 15s watchdog may have swapped in the error UI for a slow-but-fine image, detaching this img —
          // rebuild the stage so a successful late load still shows instead of staying stuck behind the error.
          if (!bodyEl.contains(img)) { bodyEl.innerHTML = ""; bodyEl.appendChild(img); }
          var l = bodyEl.querySelector(".lb-loading"); if (l) l.remove(); img.style.display = "";
        };
        img.onerror = function () { if (lb._seq !== _seq) return; lbError(url, "The image couldn’t be displayed here."); };
        img.src = dl; bodyEl.appendChild(img);
        return;
      }
      // Office docs (docx/pptx/xlsx) → Microsoft's embeddable viewer (its OneDrive bar is cropped off).
      function showEmbed() {
        return fetch("https://graph.microsoft.com/v1.0/shares/" + shareId(url) + "/driveItem/preview", {
          method: "POST", headers: { Authorization: bearer, "Content-Type": "application/json" }, body: "{}"
        }).then(function (r) { if (!r.ok) throw new Error("preview " + r.status); return r.json(); })
          .then(function (p) {
            if (lb.hidden || lb._seq !== _seq) return;
            if (!p.getUrl) throw new Error("no preview url");
            bodyEl.classList.add("doc");
            // Microsoft's viewer can take several seconds to paint the document (heavier than a raw PDF), so
            // hold a "Loading preview…" overlay OVER the frame until it fires `load` (or a 20s safety cap) —
            // otherwise a blank white frame shows in the meantime.
            bodyEl.innerHTML = '<div class="lb-framewrap"><div class="lb-loading lb-frameload">Loading preview…</div><iframe class="lb-frame" title="Document preview" src="' + escapeAttr(p.getUrl) + '" allowfullscreen></iframe></div>';
            var ifr = bodyEl.querySelector(".lb-frame"), dropLoad = function () { var l = bodyEl.querySelector(".lb-frameload"); if (l) l.remove(); };
            ifr.addEventListener("load", dropLoad); setTimeout(dropLoad, 20000);
          });
      }
      // On a PHONE, iOS Safari's native PDF-in-iframe ignores fit-to-width and won't zoom, so a doc wider
      // than the screen gets cut off with no way to see it in full. Microsoft's embeddable viewer IS
      // responsive (fits the width, pinch-to-zoom), so use it on mobile. Desktop keeps the native PDF viewer
      // below (fast, no OneDrive chrome, real zoom/paging) — where it already fits fine.
      if (window.matchMedia && window.matchMedia("(max-width: 620px)").matches) return showEmbed();
      // Any document → the browser's native PDF viewer via a blob URL (no OneDrive chrome, full-size,
      // real zoom/paging). PDFs stream directly; Office docs are converted to PDF. Falls back to MS embed.
      return docPdfBlob(url, item.name, dl, lbSig).then(function (blob) {
        if (lb.hidden || lb._seq !== _seq) return;
        lbClearBlob(); lbBlobUrl = URL.createObjectURL(blob);
        bodyEl.classList.add("doc");
        // FitH (fit to WIDTH) everywhere: a doc slightly wider than the viewer is never cut off on the right,
        // and on a short/wide desktop window whole-page "Fit" shrank it to an unreadable thumbnail.
        bodyEl.innerHTML = '<iframe class="lb-pdf" src="' + lbBlobUrl + pdfViewHash() + '" title="Document preview"></iframe>';
      }).catch(function () { if ((lbSig && lbSig.aborted) || lb._seq !== _seq) return; return showEmbed(); });
    }).catch(function (err) {
      if (lb.hidden || lb._seq !== _seq) return; // superseded/closed — don't paint a stale error over the current preview
      console.error("preview error", err);
      logPreviewFail(err);
      lbError(url, friendlyErr(err));
    });
  }
  function galResolve(url, bearer) {
    // fetchTimeout (not a bare fetch): a hung Graph driveItem resolve must not leave the gallery spinner up
    // forever — browser fetch has no default timeout, so a blackholed connection would never settle Promise.all.
    return graphJson("https://graph.microsoft.com/v1.0/shares/" + shareId(url) + "/driveItem?$expand=thumbnails($select=small,medium)", bearer)
      .then(function (item) {
        var full = item["@microsoft.graph.downloadUrl"] || item["@content.downloadUrl"] || null;
        var ts = (item.thumbnails && item.thumbnails[0]) || {};
        var img = isImageFile(item.name, (item.file && item.file.mimeType) || "");
        var thumb = window.PF.pickThumb(ts, 176) || (img ? full : null); // 96×72 gallery strip tiles at 2× → `medium`
        return { full: full, thumb: thumb, name: item.name, isImage: img };
      });
  }
  function openGallery(items, title) {
    if (!lb) { window.open(items[0].url, "_blank", "noopener"); return; }
    var _seq = lbBegin(title, "Renders", folderUrlOf(items[0].url), '<div class="lb-loading">Loading renders…</div>'); // Open → the folder that holds the views
    // Loading watchdog (mirrors openPreview): if it's still spinning after 15s, surface a recovery path
    // rather than an eternal "Loading renders…" — field sales hit exactly the flaky networks this guards.
    lbWatchdog(_seq, items[0].url);
    if (!window.PF_getToken) { lbError(items[0].url, "Please sign in again to preview files."); return; }
    var resolved = new Array(items.length);
    window.PF_getToken().then(function (token) {
      var bearer = "Bearer " + token;
      return Promise.all(items.map(function (it, i) {
        return galResolve(it.url, bearer).then(function (r) { resolved[i] = r; }).catch(function () { resolved[i] = { full: null, thumb: null, name: it.label }; });
      }));
    }).then(function () {
      if (lb.hidden || lb._seq !== _seq) return;
      buildGallery(items, resolved);
    }).catch(function (err) {
      if (lb.hidden || lb._seq !== _seq) return;
      console.error("gallery error", err);
      lbError(items[0].url, friendlyErr(err));
    });
  }
  function buildGallery(items, resolved) {
    var body = lb.querySelector(".lb-body");
    var LCHEV = svg('<path d="m15 18-6-6 6-6"/>'), RCHEV = svg('<path d="m9 18 6-6-6-6"/>');
    body.innerHTML =
      '<div class="gal">' +
        '<div class="gal-strip"></div>' + // list of documents on the LEFT (named), the viewer to its right
        '<div class="gal-main">' +
          '<button class="gal-nav gal-prev" type="button" aria-label="Previous view">' + LCHEV + '</button>' +
          '<button class="gal-nav gal-next" type="button" aria-label="Next view">' + RCHEV + '</button>' +
        '</div>' +
      '</div>';
    var main = body.querySelector(".gal-main");
    var strip = body.querySelector(".gal-strip");
    // ---- Select mode: pick tiles, then Download all / Share / Copy link act on the picks (T7-17) ----
    var bar = document.createElement("div"); bar.className = "gal-modebar";
    bar.innerHTML = '<button class="gal-sel-toggle" type="button" aria-pressed="false">Select</button>' +
      '<button class="gal-sel-all" type="button" hidden>Select all</button>' +
      '<button class="gal-sel-clear" type="button" hidden>Clear</button>' +
      '<span class="gal-sel-count" aria-live="polite"></span>';
    strip.appendChild(bar);
    lbGalKeys = items.map(function (_, i) { return i; });
    lbGalEntries = items.map(function (it, i) { var rr = resolved[i]; return { url: it.url, name: (rr && rr.name) || it.label || null }; });
    var tiles = [];
    function paintSel() {
      var on = lbSel.isOn(), n = lbSel.count();
      strip.classList.toggle("selecting", on);
      var tg = bar.querySelector(".gal-sel-toggle"); tg.textContent = on ? "Done" : "Select"; tg.setAttribute("aria-pressed", on ? "true" : "false");
      bar.querySelector(".gal-sel-all").hidden = !on; bar.querySelector(".gal-sel-clear").hidden = !on || !n;
      bar.querySelector(".gal-sel-count").textContent = on ? (n + " of " + items.length + " selected") : "";
      tiles.forEach(function (t, i) {
        var p = on && lbSel.has(i);
        t.classList.toggle("picked", p);
        if (on) { t.setAttribute("role", "checkbox"); t.setAttribute("aria-checked", p ? "true" : "false"); }
        else { t.removeAttribute("role"); t.removeAttribute("aria-checked"); }
      });
      // The share sheet takes MODEL_SHARE_CAP at most: say so once when the picks cross it, not a silent trim.
      if (on && n > MODEL_SHARE_CAP && !lbSelCapNoted && lbShareBtn && !lbShareBtn.hidden) { lbSelCapNoted = true; showToast('<span>Share can attach up to ' + MODEL_SHARE_CAP + ' files — Download and Copy links take all ' + n + '.</span>', 6000); }
      syncLbSelLabels();
    }
    bar.querySelector(".gal-sel-toggle").addEventListener("click", function () { if (lbSel.isOn()) lbSel.stop(); else lbSel.start(); paintSel(); });
    bar.querySelector(".gal-sel-all").addEventListener("click", function () { lbSel.all(lbGalKeys); paintSel(); });
    bar.querySelector(".gal-sel-clear").addEventListener("click", function () { lbSel.clear(); paintSel(); });
    // Swap only the image/error node, leaving the (absolute) arrows in place.
    function setStage(node) { var old = main.querySelector("img, .lb-err, .gal-doc, .lb-loading"); if (old) old.remove(); main.appendChild(node); }
    function errNode(url, msg) {
      var d = document.createElement("div"); d.className = "lb-err";
      d.innerHTML = '<p>' + escapeHtml(msg) + '</p><a class="lb-openbtn" href="' + escapeAttr(url) + '" target="_blank" rel="noopener">Open in SharePoint →</a>';
      return d;
    }
    var prevBtn = body.querySelector(".gal-prev");
    var nextBtn = body.querySelector(".gal-next");
    // Nothing to navigate with a single view → hide the arrows and the thumbnail strip entirely.
    var single = items.length < 2;
    if (single) { strip.style.display = "none"; prevBtn.style.display = "none"; nextBtn.style.display = "none"; }
    var idx = 0;
    // Extension-stripped display name so each item reads as "what it actually is" (not "….pdf").
    // Strip a real file extension only. /\.[a-z0-9]{1,6}$/ also ate a version tail, so "Shelf v1.2"
    // displayed as "Shelf v1" — an extension has at least one letter, a version segment is all digits.
    function galName(i) { var rr = resolved[i], nm = (rr && rr.name) || items[i].label || ("View " + (i + 1)); return String(nm).replace(/\.(?=[a-z0-9]{1,6}$)(?=[a-z0-9]*[a-z])[a-z0-9]+$/i, "") || nm; }
    items.forEach(function (it, i) {
      var nm = galName(i);
      var b = document.createElement("button"); b.type = "button"; b.className = "gal-thumb"; b.title = nm;
      var r = resolved[i];
      var isImg = (r && typeof r.isImage === "boolean") ? r.isImage : (it.isImage !== false);
      var fig = (r && r.thumb) ? '<img src="' + escapeAttr(r.thumb) + '" alt="" loading="lazy" decoding="async">' : '<span class="gal-tico">' + svg(isImg ? ICONS.render : ICONS.spec) + '</span>';
      b.innerHTML = '<span class="gal-tick" aria-hidden="true">' + TICKICON + '</span><span class="gal-thumb-fig">' + fig + '</span><span class="gal-thumb-name">' + escapeHtml(nm) + '</span>';
      b.addEventListener("click", function () { if (lbSel.isOn()) { lbSel.toggle(i); paintSel(); return; } show(i); }); // a pick, not a page
      strip.appendChild(b); tiles.push(b);
    });
    function galEmbed(it, want) { // document → Microsoft's responsive embed viewer
      docPreviewUrl(it.url).then(function (u) {
        if (idx !== want || lb.hidden) return;
        var wrap = document.createElement("div"); wrap.className = "gal-doc embed";
        // Hold a loading overlay until the (slower) MS viewer paints, so it doesn't flash blank.
        wrap.innerHTML = '<div class="lb-loading lb-frameload"><span class="lb-skel" aria-hidden="true"></span><span>Loading preview…</span></div><iframe src="' + escapeAttr(u) + '" title="Document preview" allowfullscreen></iframe>';
        var ifr = wrap.querySelector("iframe"), drop = function () { var l = wrap.querySelector(".lb-frameload"); if (l) l.remove(); };
        ifr.addEventListener("load", drop); setTimeout(drop, 20000);
        setStage(wrap);
      }).catch(function () { if (idx === want && !lb.hidden) setStage(errNode(it.url, "Preview isn’t available here.")); });
    }
    function show(i) {
      idx = i; lbClearBlob();
      var sig = lbFreshSignal(); // paging away cancels the previous view's in-flight conversion
      // NOTE: "Open in SharePoint" stays pointed at the containing FOLDER (set by the caller),
      // not the individual view — so people land in the real folder location. Download still
      // grabs the specific file shown (lbFile below).
      var r = resolved[i], it = items[i];
      lbFile = { name: (r && r.name) || it.label || null, downloadUrl: (r && r.full) || null, fileUrl: it.url };
      var isImg = (r && typeof r.isImage === "boolean") ? r.isImage : (it.isImage !== false);
      if (isImg) { // image
        if (r && r.full) {
          var img = document.createElement("img"); img.alt = it.label || (r && r.name) || "Product view"; // main content image — never leave alt empty
          // Guard like the doc branch: a late error from a PREVIOUS view's image must not paint over the view
          // the user has since paged to.
          img.onerror = function () { if (idx !== i || lb.hidden) return; setStage(errNode(it.url, "This view couldn’t be displayed.")); };
          img.src = r.full; setStage(img);
        } else { setStage(errNode(it.url, "This view couldn’t be loaded.")); }
      } else { // document
        var load = document.createElement("div"); load.className = "lb-loading";
        load.innerHTML = '<span class="lb-skel" aria-hidden="true"></span><span>Loading preview…</span>';
        setStage(load);
        var want = i;
        // On a PHONE, iOS's native PDF-in-iframe ignores fit-to-width so gallery docs get cut off on the
        // right — route them through the responsive MS viewer too (same as the single-doc preview). Desktop
        // keeps the fast native PDF viewer (no OneDrive chrome, full-size), falling back to MS embed on error.
        if (window.matchMedia && window.matchMedia("(max-width: 620px)").matches) { galEmbed(it, want); }
        else docPdfBlob(it.url, (r && r.name), (r && r.full), sig).then(function (blob) {
          if (idx !== want || lb.hidden) return;
          lbClearBlob(); lbBlobUrl = URL.createObjectURL(blob);
          var wrap = document.createElement("div"); wrap.className = "gal-doc";
          wrap.innerHTML = '<iframe src="' + lbBlobUrl + pdfViewHash() + '" title="Document preview"></iframe>';
          setStage(wrap);
        }).catch(function () { if (sig && sig.aborted) return; galEmbed(it, want); });
      }
      // Walk the TILES, not the strip's child nodes: the Select bar is the strip's first child, and counting it put
      // the highlight one tile behind the file on stage (hands-on L2).
      tiles.forEach(function (tile, t) { var onT = t === idx; tile.classList.toggle("on", onT); if (onT) tile.setAttribute("aria-current", "true"); else tile.removeAttribute("aria-current"); });
      // Announce the new view for screen readers (the stage swap is otherwise silent).
      if (srAnnounce && items.length > 1) srAnnounce.textContent = (it.label || ("View " + (i + 1))) + " — view " + (i + 1) + " of " + items.length;
      // Don't leave focus on a button we're about to disable — the browser drops it to <body>, escaping
      // the modal's focus context. Move focus to the other arrow first, so paging to a boundary keeps
      // the user inside the dialog.
      var offP = idx <= 0, offN = idx >= items.length - 1;
      if (offP && !offN && document.activeElement === prevBtn) nextBtn.focus();
      else if (offN && !offP && document.activeElement === nextBtn) prevBtn.focus();
      prevBtn.disabled = offP; nextBtn.disabled = offN;
    }
    prevBtn.addEventListener("click", function () { if (idx > 0) show(idx - 1); });
    nextBtn.addEventListener("click", function () { if (idx < items.length - 1) show(idx + 1); });
    galNav = single ? null : { prev: function () { if (idx > 0) show(idx - 1); }, next: function () { if (idx < items.length - 1) show(idx + 1); } };
    // Offer "Download all" for a 2+ file set; the list mirrors what show() puts in lbFile per item.
    setDownloadAll(items.length >= 2 ? items.map(function (it, i) { return { fileUrl: it.url, downloadUrl: (resolved[i] && resolved[i].full) || null, name: (resolved[i] && resolved[i].name) || it.label || null }; }) : null);
    paintSel();
    show(0);
  }
  // ---- folder-backed renders (SP Folder URL points at a folder; we list it live) ----
  // A link with no file extension in its last path segment is treated as a folder.
  // Model-name matching lives in pf-shared.js (loaded via <script src>) so the browser and the server
  // cron use ONE implementation of the HMX/GTX/SNX naming rules. Thin delegators keep call sites intact.
  function looksLikeFolder(url) { return window.PF.looksLikeFolder(url); }
  function modelRe(model) { return window.PF.modelRe(model); }
  function normModel(s) { return window.PF.normModel(s); }
  function matchesModel(name, model) { return window.PF.matchesModel(name, model); }
  // Turn a filename into a short view label (drop model code, date stamp, extension).
  function viewLabel(name, model) {
    var base = String(name).replace(/\.[a-z0-9]+$/i, "");
    var re = modelRe(model); if (re) base = base.replace(re, "");
    base = base.replace(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s*\d{0,4}\b/ig, "");
    base = base.replace(/[-_]+/g, " ").replace(/\s+/g, " ").trim();
    return base ? base.charAt(0).toUpperCase() + base.slice(1) : "View";
  }
  function isImageFile(name, mime) {
    return /^image\//.test(mime || "") || /\.(png|jpe?g|gif|webp|bmp|svg)(\?|#|$)/i.test(name || "");
  }
  function mapFolderItem(it) {
    var ts = (it.thumbnails && it.thumbnails[0]) || {};
    return {
      name: it.name,
      url: it.webUrl,
      full: it["@microsoft.graph.downloadUrl"] || it["@content.downloadUrl"] || null,
      thumb: window.PF.pickThumb(ts, 176), // feeds both the wire tiles and the gallery strip → `medium`
      isImage: isImageFile(it.name, (it.file && it.file.mimeType) || ""),
      fileId: it.id || null
    };
  }
  // Raw list of every file in a folder (no model filter). Follows @odata.nextLink so folders with
  // >200 files aren't truncated (returns a Promise, so .then callers are unaffected).
  // Browser fetch() has NO default timeout, so a hung Graph call left "Loading renders…" spinning forever.
  // Abort after `ms` and reject with a "timed out" error the callers' .catch already turns into a friendly
  // message (see friendlyErr). Each page of the listing gets its own timeout.
  // One Graph JSON read: bounded (15s), bearer set, non-2xx thrown with `status` so callers can branch on
  // it (lhTransient maps "timed out" to unknown). Six call sites hand-rolled this before and drifted.
  function graphJson(url, bearer, ms) {
    return fetchTimeout(url, { headers: { Authorization: bearer } }, ms || 15000)
      .then(function (r) { if (!r.ok) { var e = new Error("Graph " + r.status); e.status = r.status; throw e; } return r.json(); });
  }
  function fetchTimeout(url, opts, ms) {
    var ctrl = new AbortController();
    var t = setTimeout(function () { ctrl.abort(); }, ms || 15000);
    return fetch(url, Object.assign({}, opts || {}, { signal: ctrl.signal }))
      .catch(function (e) { if (e && e.name === "AbortError") throw new Error("timed out"); throw e; })
      .finally(function () { clearTimeout(t); });
  }
  async function fetchFolderChildren(folderUrl, bearer) {
    var url = "https://graph.microsoft.com/v1.0/shares/" + shareId(folderUrl) + "/driveItem/children?$expand=thumbnails($select=small,medium)&$top=200";
    var out = [];
    for (var guard = 0; guard < 25 && url; guard++) { // guard caps at ~5000 files
      var j = await graphJson(url, bearer);
      (j.value || []).forEach(function (it) { if (it.file) out.push(mapFolderItem(it)); });
      url = j["@odata.nextLink"] || null;
    }
    // Truncated at the page cap → throw rather than return a short list, so link-health reads "unknown"
    // (preserved) instead of a false "empty" (mirrors the server-side fetchFolderChildren in graphApp.js).
    if (url) throw new Error("folder listing exceeded the page cap (truncated)");
    return out;
  }
  // How many catalog resources point at this exact folder URL — >1 means it's a SHARED folder
  // (filter files by model name); ==1 means a DEDICATED folder (show every file, no naming rules).
  function folderRefCount(url) {
    var n = 0, cu = window.PF.canonFolderUrl(url); // compare canonically so two spellings of one folder count as one
    Object.keys(LINKS).forEach(function (k) { var a = normLinks(LINKS[k]); if (a.length === 1 && window.PF.canonFolderUrl(a[0].url) === cu) n++; });
    return n;
  }
  function listFolderRenders(folderUrl, model, bearer) {
    var shared = folderRefCount(folderUrl) > 1;
    return fetchFolderChildren(folderUrl, bearer).then(function (files) {
      var matched = shared ? files.filter(function (f) { return matchesModel(f.name, model); }) : files;
      matched.sort(function (a, b) { return String(a.name).localeCompare(String(b.name)); });
      return matched.map(function (f) {
        return { label: viewLabel(f.name, model), name: f.name, url: f.url, full: f.full, thumb: f.thumb, isImage: f.isImage };
      });
    });
  }
  function openRenders(folderUrl, model, title) {
    if (!lb) { window.open(folderUrl, "_blank", "noopener"); return; }
    var _seq = lbBegin(title, "Renders", folderUrl, '<div class="lb-loading">Loading renders…</div>');
    if (!window.PF_getToken) { lbError(folderUrl, "Please sign in again to preview files."); return; }
    // The same 15s watchdog every other opener has. Without it a stalled token or Graph call left
    // "Loading renders…" on screen indefinitely, with no way out but closing the modal.
    lbWatchdog(_seq, folderUrl);
    window.PF_getToken().then(function (token) {
      return listFolderRenders(folderUrl, model, "Bearer " + token);
    }).then(function (items) {
      if (lb.hidden || lb._seq !== _seq) return;
      if (!items.length) { lbError(folderUrl, "No renders for " + model + " in this folder yet."); return; }
      var resolved = items.map(function (it) { return { full: it.full, thumb: it.thumb, name: it.name || it.label, isImage: it.isImage }; });
      buildGallery(items, resolved);
    }).catch(function (err) {
      if (lb.hidden || lb._seq !== _seq) return;
      console.error("renders error", err);
      logPreviewFail(err);
      lbError(folderUrl, friendlyErr(err));
    });
  }
  // Mixed resource: a cell holding BOTH a folder url AND one or more file urls (e.g. a folder-wired model
  // that later had a single file pinned — Tier-1 fix #2 can create these). Expand every folder to its files
  // and flatten all of it into ONE gallery, so a folder url no longer lands as a broken "file" tile.
  // Folder listings already carry the resolved fields; direct files are resolved via galResolve.
  function openMixedGallery(arr, title, model) {
    if (!lb) { window.open(arr[0].url, "_blank", "noopener"); return; }
    var folderHref = (arr.filter(function (a) { return looksLikeFolder(a.url); })[0] || {}).url || folderUrlOf(arr[0].url);
    var _seq = lbBegin(title, "Renders", folderHref, '<div class="lb-loading">Loading renders…</div>');
    // Loading watchdog (mirrors openPreview/openGallery): recover from a hung resolve instead of spinning forever.
    lbWatchdog(_seq, arr[0].url);
    if (!window.PF_getToken) { lbError(arr[0].url, "Please sign in again to preview files."); return; }
    window.PF_getToken().then(function (token) {
      var bearer = "Bearer " + token;
      return Promise.all(arr.map(function (a) {
        if (looksLikeFolder(a.url)) {
          return listFolderRenders(a.url, model, bearer).then(function (fs) {
            return fs.map(function (it) { return { item: { label: it.label, url: it.url, isImage: it.isImage }, resolved: { full: it.full, thumb: it.thumb, name: it.name, isImage: it.isImage } }; });
          }).catch(function () { return []; }); // a folder that fails to list drops out; the pinned files still show
        }
        return galResolve(a.url, bearer).then(function (r) {
          return [{ item: { label: a.label, url: a.url, isImage: r.isImage }, resolved: r }];
        }).catch(function () { return [{ item: { label: a.label, url: a.url }, resolved: { full: null, thumb: null, name: a.label } }]; });
      }));
    }).then(function (groups) {
      if (lb.hidden || lb._seq !== _seq) return;
      var items = [], resolved = [], seen = {};
      groups.forEach(function (g) { g.forEach(function (x) {
        var key = x.item.url ? window.PF.canonFileUrl(x.item.url) : ("#" + items.length);
        if (seen[key]) return; seen[key] = 1; // a pinned file that also lives in the folder shows once, not twice
        items.push(x.item); resolved.push(x.resolved);
      }); });
      if (!items.length) { lbError(folderHref, "No files to preview here yet."); return; }
      buildGallery(items, resolved);
    }).catch(function (err) {
      if (lb.hidden || lb._seq !== _seq) return;
      console.error("mixed gallery error", err); logPreviewFail(err); lbError(arr[0].url, friendlyErr(err));
    });
  }
  // One entry point for opening a resource: folder → live render gallery, many files → gallery, one file → preview.
  // A mix of folder(s) + file(s) routes to openMixedGallery, which expands the folder(s) into the same gallery.
  function openResource(arr, title, model) {
    if (!arr || !arr.length) return;
    // remember what the lightbox is showing so its "Copy link" button can build a deep link
    var rname = (title && title.indexOf("·") !== -1) ? title.split("·").pop().trim() : null;
    lbShare = { model: model, resource: rname };
    window.PF_logEvent && window.PF_logEvent("resource_open", model ? (model + (rname ? " · " + rname : "")) : (title || ""));
    if (arr.length === 1 && looksLikeFolder(arr[0].url)) openRenders(arr[0].url, model, title);
    else if (arr.length > 1) { if (arr.some(function (a) { return looksLikeFolder(a.url); })) openMixedGallery(arr, title, model); else openGallery(arr, title); }
    else openPreview(arr[0].url, title);
  }
  function resourceClick(e, arr, title, model) {
    if (e.ctrlKey || e.metaKey || e.shiftKey) return; // modified click → let the browser open SharePoint in a new tab
    e.preventDefault(); e.stopPropagation();
    openResource(arr, title, model);
  }

  /* ============================================================
     GLOBAL EVENTS, SHORTCUTS & THEME
     ============================================================ */
  // ---- events ----
  var brandEl = document.getElementById("brand");
  if (brandEl) brandEl.addEventListener("click", function () {
    var mark = this.querySelector(".brand-mark");
    if (mark) { mark.classList.remove("tap"); void mark.offsetWidth; mark.classList.add("tap"); }
    confirmDiscardStaged().then(function (ok) { if (ok) goHome(); });
  });
  var searchTimer = null;
  searchInput.addEventListener("input", function () {
    var q = searchInput.value;
    if (searchTimer) { clearTimeout(searchTimer); searchTimer = null; }
    if (q.trim().length === 0) {
      // Clearing restores the tab the viewer was on before searching (a PM on Progress/Metrics isn't
      // stranded on Find), or the catalog otherwise. Also drop any active result facets so the next
      // search starts clean rather than inheriting a filter from the last one.
      view = _preSearchView || "catalog"; _preSearchView = null; searchFilter = newFilterState();
      render(); syncViewTabs(); return;
    }
    if (_preSearchView === null && view !== "catalog") _preSearchView = view; // remember the PM tab we're leaving to search
    view = "catalog";
    searchTimer = setTimeout(function () { renderSearch(q); syncViewTabs(); }, 120); // debounce keystrokes
  });
  // ArrowDown from the search box jumps into the results list (paired with the list's roving nav).
  searchInput.addEventListener("keydown", function (e) {
    if (e.key === "ArrowDown") {
      var first = content.querySelector(".sr-list .sr");
      if (first) { e.preventDefault(); first.focus(); }
      return;
    }
    // Esc clears the query and returns to the tab you were on — reuses the input handler's restore path.
    if (e.key === "Escape" && searchInput.value) {
      e.preventDefault(); searchInput.value = ""; searchInput.dispatchEvent(new Event("input"));
    }
  });
  // MOBILE: collapse the header chrome — the search bar AND the view tabs — while scrolling DOWN, and
  // bring both back on scroll UP. rAF-throttled; flips one body class that the CSS collapses on phones
  // only. Always shown near the top of the page, and never hidden while focus is anywhere in the header,
  // so neither a search in progress nor a tab a keyboard user is on can vanish from under them.
  (function () {
    var lastY = 0, ticking = false, topbar = document.querySelector(".topbar");
    function show() { document.body.classList.remove("nav-hidden"); }
    function update() {
      ticking = false;
      var y = window.scrollY || document.documentElement.scrollTop || 0;
      var mobile = !window.matchMedia || window.matchMedia("(max-width: 620px)").matches;
      var focusInHeader = topbar && topbar.contains(document.activeElement);
      if (!mobile || y < 80 || focusInHeader) { show(); lastY = y; return; }
      if (y > lastY + 4) document.body.classList.add("nav-hidden");         // scrolling down → hide
      else if (y < lastY - 4) show();                                       // scrolling up → show
      lastY = y;
    }
    window.addEventListener("scroll", function () { if (!ticking) { ticking = true; window.requestAnimationFrame(update); } }, { passive: true });
    // Collapsed controls stay focusable on purpose: Tabbing into one reveals the header rather than
    // landing the user on something they cannot see. (visibility:hidden would remove them from the tab
    // order entirely, which would make the tabs unreachable by keyboard while scrolled down.)
    if (topbar) topbar.addEventListener("focusin", show);
  })();
  Array.prototype.forEach.call(document.querySelectorAll(".vtab"), function (btn) {
    btn.addEventListener("click", function () {
      var v = btn.getAttribute("data-view");
      if (v === view) return;
      confirmDiscardStaged().then(function (ok) { if (ok) switchView(v); });
    });
  });
  function switchView(v) {
    view = v; _preSearchView = null; // an explicit tab choice supersedes any search-return memory
    // Clear the query too: text left in the box rendered the catalog home underneath a live search, and
    // made the tab-return refetch treat the page as permanently "busy". (applyHash/goHome/openTo do this.)
    if (searchInput && searchInput.value) { searchInput.value = ""; searchFilter = newFilterState(); }
    try { localStorage.setItem("pf-lasttab", v); } catch (e) {} // reopen this tab on a later bare visit
    render(); syncHash(); resetScroll();
  }
  // ARIA tabs keyboard pattern: arrow/Home/End move focus between VISIBLE tabs and activate.
  var tablistEl = document.querySelector(".viewtabs");
  if (tablistEl) tablistEl.addEventListener("keydown", function (e) {
    if (["ArrowRight", "ArrowLeft", "Home", "End"].indexOf(e.key) === -1) return;
    var vis = Array.prototype.filter.call(document.querySelectorAll(".vtab"), function (t) { return !t.classList.contains("vtab-pm-hidden"); });
    var idx = vis.indexOf(document.activeElement);
    if (idx === -1) return;
    e.preventDefault();
    var next = e.key === "Home" ? 0 : e.key === "End" ? vis.length - 1 : e.key === "ArrowRight" ? (idx + 1) % vis.length : (idx - 1 + vis.length) % vis.length;
    vis[next].focus(); vis[next].click();
  });
  document.addEventListener("keydown", function (e) {
    if (e.key !== "/") return;
    // WCAG 2.1.4 (Level A): a single printable-character shortcut must be turn-off-able, because
    // speech-input users emit stray characters constantly. The field/modal guards below stop it eating
    // TYPED text; they don't help someone dictating. `pf-shortcuts=off` disables it (and hides the hint).
    if (shortcutsOff()) return;
    var a = document.activeElement;
    // Don't hijack "/" while the user is typing in a field (it would eat the character)…
    if (a && (a === searchInput || /^(input|textarea|select)$/i.test(a.tagName) || a.isContentEditable)) return;
    // …or while a modal is open (search sits behind it).
    // Reuse the module-level `lb` rather than re-querying into a shadowing local (which read as a
    // different element to anyone scanning this file).
    var bm = document.getElementById("bugModal");
    if ((lb && !lb.hidden) || (bm && !bm.hidden)) return;
    e.preventDefault(); searchInput.focus();
  });

  // Keyboard-shortcut preference (see the "/" handler above). Remembered per device, like the theme.
  function shortcutsOff() { try { return localStorage.getItem("pf-shortcuts") === "off"; } catch (e) { return false; } }
  // The hint IS the switch: a screen-reader or speech-input user needs an on-page control (WCAG 2.1.4),
  // not a console hook. Off, it stays visible (struck through) so there is a way back on.
  function applyShortcutHint() {
    var kbd = document.querySelector(".search-kbd"); if (!kbd) return;
    var off = shortcutsOff();
    kbd.classList.toggle("off", off);
    kbd.setAttribute("aria-pressed", off ? "false" : "true");
    kbd.setAttribute("aria-label", off ? "Keyboard shortcut off: the / key no longer jumps to search. Turn it on" : "Keyboard shortcut: press / to jump to search. Turn it off");
    kbd.setAttribute("title", off ? "Shortcut off \u00b7 click to turn on" : "Press / to search \u00b7 click to turn off");
  }
  applyShortcutHint();
  var kbdHint = document.querySelector(".search-kbd");
  if (kbdHint) kbdHint.addEventListener("click", function () {
    var turnOn = shortcutsOff();
    window.PF_setShortcuts(turnOn);
    showToast('<span>' + (turnOn ? 'Keyboard shortcut on \u2014 press / to search.' : 'Keyboard shortcut off \u2014 / types normally now.') + '</span>');
  });
  window.PF_setShortcuts = function (on) {
    try { localStorage.setItem("pf-shortcuts", on ? "on" : "off"); } catch (e) {}
    applyShortcutHint();
  };

  // ---- theme toggle ----
  var SUN = svg('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>');
  var MOON = svg('<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/>');
  // Tri-state theme: System → Light → Dark → System. "System" clears the override so the OS
  // preference (prefers-color-scheme) wins again — previously the toggle was stuck at light/dark
  // with no way back to following the OS.
  var AUTO = svg('<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M8 21h8M12 17v4"/>');
  var themeToggle = document.getElementById("themeToggle");
  var THEME_ORDER = ["system", "light", "dark"];
  function themeMode() { try { var t = localStorage.getItem("pf-theme"); return (t === "light" || t === "dark") ? t : "system"; } catch (e) { return "system"; } }
  // The two <meta name="theme-color"> tags in the head are keyed to prefers-color-scheme, so they follow
  // the OS and NOTHING else — the in-app toggle never moved them. Installed to a Home Screen the app runs
  // standalone, where iOS paints the status bar from this value: an explicit dark toggle on a phone set to
  // light left a white strip above a dark app. Read the resolved --surface rather than repeating the hex,
  // so the status bar can never drift from the theme it is meant to match.
  var themeColorMeta = null;
  function syncThemeColor() {
    if (!themeColorMeta) {
      var old = document.querySelectorAll('meta[name="theme-color"]');
      for (var i = 0; i < old.length; i++) old[i].parentNode.removeChild(old[i]); // media-keyed ones would win over ours
      themeColorMeta = document.createElement("meta");
      themeColorMeta.setAttribute("name", "theme-color");
      document.head.appendChild(themeColorMeta);
    }
    var surface = getComputedStyle(document.documentElement).getPropertyValue("--surface").trim();
    themeColorMeta.setAttribute("content", surface || "#ffffff");
  }
  function applyTheme(mode) {
    if (mode === "system") { document.documentElement.removeAttribute("data-theme"); try { localStorage.removeItem("pf-theme"); } catch (e) {} }
    else { document.documentElement.setAttribute("data-theme", mode); try { localStorage.setItem("pf-theme", mode); } catch (e) {} }
    syncThemeColor();
  }
  function syncToggle() {
    if (!themeToggle) return;
    var m = themeMode(), name = m.charAt(0).toUpperCase() + m.slice(1);
    themeToggle.innerHTML = m === "system" ? AUTO : (m === "dark" ? MOON : SUN);
    themeToggle.setAttribute("title", "Theme: " + name + " (click to change)");
    themeToggle.setAttribute("aria-label", "Theme: " + name + ". Click to change.");
  }
  if (themeToggle) themeToggle.addEventListener("click", function () {
    applyTheme(THEME_ORDER[(THEME_ORDER.indexOf(themeMode()) + 1) % THEME_ORDER.length]);
    syncToggle();
  });
  syncToggle();
  syncThemeColor();
  // While following the OS, track it changing (sunset/sunrise, or Control Centre) so the status bar keeps up.
  if (window.matchMedia) {
    var darkQ = window.matchMedia("(prefers-color-scheme: dark)");
    var onSchemeChange = function () { if (themeMode() === "system") syncThemeColor(); };
    if (darkQ.addEventListener) darkQ.addEventListener("change", onSchemeChange);
    else if (darkQ.addListener) darkQ.addListener(onSchemeChange);
  }

  /* ============================================================
     LIGHTBOX ACTIONS — copy link, overflow, download, share
     ============================================================ */
  // ---- copy shareable link (lives inside the resource preview, next to "Open in SharePoint") ----
  var LINKICON = svg('<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>');
  var lbCopyBtn = document.getElementById("lbCopyBtn");
  function fallbackCopy(text) {
    try {
      var ta = document.createElement("textarea"); ta.value = text;
      ta.style.position = "fixed"; ta.style.top = "-1000px"; ta.style.opacity = "0";
      document.body.appendChild(ta); ta.focus(); ta.select();
      var ok = document.execCommand("copy"); ta.remove(); return ok;
    } catch (e) { return false; }
  }
  // Copy-link hint popup: the copied link is a HUB deep-link, gated to Grumbleton sign-ins — useless to an
  // external customer. Surface that (and point them to Download) on hover/focus AND on copy, as a toast-styled
  // popup anchored at the top by the preview header. One toggled element (not a fresh toast per mousemove) so
  // it can persist while the button is hovered/focused instead of flickering.
  var INFOICON = svg('<circle cx="12" cy="12" r="9"/><path d="M12 11v5"/><path d="M12 8h.01"/>');
  var COPY_HINT_MSG = INFOICON + '<span><b>Internal link</b> — opens only for Grumbleton sign-ins. To share with a customer, use <b>Download</b> and send them the file.</span>';
  var lbCopyHint = null, copyHintPinned = false, copyHintTimer;
  function copyHintEl() {
    if (!lbCopyHint && toastTopWrap) {
      lbCopyHint = document.createElement("div"); lbCopyHint.className = "copy-hint"; lbCopyHint.id = "lbCopyHint";
      lbCopyHint.innerHTML = COPY_HINT_MSG; // present from the start so aria-describedby resolves
      toastTopWrap.appendChild(lbCopyHint);
      if (lbCopyBtn) lbCopyBtn.setAttribute("aria-describedby", "lbCopyHint");
    }
    return lbCopyHint;
  }
  // Only rewrite when the text actually differs. The hint lives in a polite live region AND is the
  // button's aria-describedby target, so re-assigning identical markup on every hover/focus made a
  // screen reader announce the description a second time for no reason.
  function setCopyHint(h, html) { if (h && h.innerHTML !== html) h.innerHTML = html; }
  function showCopyHint() { var h = copyHintEl(); if (!h) return; if (!copyHintPinned) setCopyHint(h, COPY_HINT_MSG); h.classList.add("show"); }
  function hideCopyHint() { if (lbCopyHint && !copyHintPinned) lbCopyHint.classList.remove("show"); }
  function clearCopyHint() { copyHintPinned = false; clearTimeout(copyHintTimer); if (lbCopyHint) { lbCopyHint.classList.remove("show"); setCopyHint(lbCopyHint, COPY_HINT_MSG); } }
  function lbCopyDone() {
    var h = copyHintEl();
    if (h) {
      setCopyHint(h, LINKICON + '<span><b>Copied</b> — internal link (Grumbleton sign-in only). To share with a customer, <b>Download</b> the file and send it.</span>');
      h.classList.add("show"); copyHintPinned = true;
      // Restore the base wording when the confirmation expires — this element is also the Copy button's
      // aria-describedby target, so leaving "Copied" here made the button describe itself as already copied
      // for the rest of the session.
      clearTimeout(copyHintTimer); copyHintTimer = setTimeout(function () { copyHintPinned = false; if (lbCopyHint) { lbCopyHint.classList.remove("show"); setCopyHint(lbCopyHint, COPY_HINT_MSG); } }, 4200);
    }
    if (lbCopyBtn) { lbCopyBtn.classList.add("copied"); setTimeout(function () { lbCopyBtn.classList.remove("copied"); }, 1400); }
  }
  function lbCopyLabel() {
    if (lbCopyBtn) lbCopyBtn.innerHTML = LINKICON + '<span>' + window.PF.selectionLabel("Copy link", "Copy selected links", lbSel) + '</span>';
  }
  if (lbCopyBtn) {
    lbCopyLabel();
    copyHintEl(); // build it now so #lbCopyHint exists for aria-describedby, not only after a hover
    lbCopyBtn.addEventListener("mouseenter", showCopyHint);
    lbCopyBtn.addEventListener("focus", showCopyHint);
    lbCopyBtn.addEventListener("mouseleave", hideCopyHint);
    lbCopyBtn.addEventListener("blur", hideCopyHint);
    // WCAG 1.4.13: hover/focus content must be dismissible without moving the pointer or focus.
    lbCopyBtn.addEventListener("keydown", function (e) { if (e.key === "Escape" && lbCopyHint && lbCopyHint.classList.contains("show")) { e.stopPropagation(); clearCopyHint(); } });
    lbCopyBtn.addEventListener("click", function () {
      if (lbSel.isOn()) { // Copy selected links: the picked files' SharePoint links as one rich pack
        var picks = lbSel.pick(lbGalKeys || []).map(function (i) { return lbGalEntries[i]; }).filter(function (e) { return e && e.url; });
        if (!picks.length) return;
        var out = sharePackContent({ kind: "model", title: lb.querySelector(".lb-title").textContent, resources: picks.map(function (e) { return { name: e.name || "file", url: e.url }; }) });
        copyRichText(out.html, out.text).then(function () {
          lbCopyDone();
          showToast(PACKICON + '<span><b>' + picks.length + ' link' + (picks.length === 1 ? '' : 's') + ' copied</b> — they open only for people with access to that SharePoint folder.</span>', 6000);
        }, function () { showToast('<span>Couldn’t copy — try again.</span>'); });
        return;
      }
      var url = (lbShare && lbShare.model) ? resourceShareUrl(lbShare.model, lbShare.resource) : location.href;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(url).then(lbCopyDone, function () { if (fallbackCopy(url)) lbCopyDone(); });
      } else if (fallbackCopy(url)) { lbCopyDone(); }
    });
  }

  // ---- "⋯" overflow: Download all / Copy link / Flag / Manage (+ Open in SharePoint on a device
  // that can Share). The header had EIGHT equal-weight pills, so nothing said what to do first and the
  // row wrapped to two lines at 390px.
  //
  // A DISCLOSURE (aria-expanded + aria-controls), not a menu and not a dialog: it promises no arrow-key
  // menuitem semantics, and it deliberately does NOT move focus on open — Tab walks into it and back
  // out to ✕ because it lives INLINE after its trigger, and the lightbox's own trap contains it.
  // Focusing the first item WOULD have been wrong here for a concrete reason: that item is Copy link,
  // whose focus shows its 1.4.13-dismissible hint, whose Escape handler stopPropagation()s — so the
  // first Escape dismissed the hint and the menu stayed open. Layered as-is, the hint takes the first
  // Escape only when it is actually showing, and the menu takes the next. ----
  var FLAGICON = svg('<path d="M4.5 21V4"/><path d="M4.5 4.6h12l-2 3.9 2 3.9h-12z"/>');
  var MANAGEICON = svg('<rect x="8.5" y="8.5" width="11" height="11" rx="2"/><path d="M4.5 15.5V6a2 2 0 0 1 2-2H16"/>');
  var MOREICON = svg('<circle cx="5" cy="12" r="1.15" fill="currentColor" stroke="none"/><circle cx="12" cy="12" r="1.15" fill="currentColor" stroke="none"/><circle cx="19" cy="12" r="1.15" fill="currentColor" stroke="none"/>');
  var lbMoreBtn = document.getElementById("lbMoreBtn"), lbMoreMenu = document.getElementById("lbMoreMenu");
  function lbMoreOpen() { return !!(lbMoreMenu && !lbMoreMenu.hidden); }
  function lbMoreClose(refocus) {
    if (!lbMoreMenu || lbMoreMenu.hidden) return;
    lbMoreMenu.hidden = true;
    if (lbMoreBtn) { lbMoreBtn.setAttribute("aria-expanded", "false"); if (refocus) lbMoreBtn.focus(); }
  }
  var lbOpenEl = document.getElementById("lbOpenLink");
  if (lbOpenEl) lbOpenEl.innerHTML = svg(ICONS.ext) + '<span>Open in SharePoint</span>';
  if (lbMoreBtn && lbMoreMenu) {
    lbMoreBtn.innerHTML = MOREICON;
    lbMoreBtn.addEventListener("click", function () {
      if (lbMoreOpen()) { lbMoreClose(false); return; }
      lbMoreMenu.hidden = false;
      lbMoreBtn.setAttribute("aria-expanded", "true");
    });
    // Escape closes the MENU, not the whole preview. The listener sits on the wrapper so
    // stopPropagation actually keeps the event off document (where the preview's own handler lives) —
    // two listeners on document could not be ordered this way.
    lbMoreBtn.parentNode.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && lbMoreOpen()) { e.stopPropagation(); lbMoreClose(true); }
    });
    // Any action inside the menu is a decision; close behind it. (Most of them close the preview too.)
    lbMoreMenu.addEventListener("click", function (e) { if (e.target.closest("button, a")) lbMoreClose(false); });
    document.addEventListener("click", function (e) {
      if (!lbMoreOpen()) return;
      if (!e.target.closest(".lb-more-wrap")) lbMoreClose(false);
    });
  }

  // ---- download the actual file shown in the preview (image / spec sheet / whatever it is) ----
  var DLICON = svg('<path d="M12 3v12"/><path d="m8 11 4 4 4-4"/><path d="M5 21h14"/>');
  var lbDownloadBtn = document.getElementById("lbDownloadBtn");
  function saveBlob(blob, name) {
    var u = URL.createObjectURL(blob), a = document.createElement("a");
    a.href = u; a.download = name || "download"; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(u); a.remove(); }, 1500);
  }
  function setDlBusy(on) {
    if (!lbDownloadBtn) return;
    lbDownloadBtn.disabled = !!on;
    lbDownloadBtn.innerHTML = DLICON + '<span>' + (on ? "Saving…" : "Download") + '</span>';
  }
  // Re-resolve a fresh pre-authenticated downloadUrl from the stable SharePoint URL. The downloadUrl captured
  // at preview-open expires (~1h), so Download on a long-open preview 404s — and the OLD fallback handed that
  // same expired URL to an <a>, so it broke too. This yields a live URL to retry with.
  function freshDownloadUrl(fileUrl) {
    if (!fileUrl || !window.PF_getToken) return Promise.reject(new Error("no url"));
    return window.PF_getToken()
      .then(function (token) { return graphJson("https://graph.microsoft.com/v1.0/shares/" + shareId(fileUrl) + "/driveItem", "Bearer " + token); })
      .then(function (item) { return { url: item["@microsoft.graph.downloadUrl"] || item["@content.downloadUrl"] || null, name: item.name || null }; });
  }
  function downloadCurrent() {
    var f = lbFile;
    if (!f || (!f.downloadUrl && !f.fileUrl)) return;
    setDlBusy(true);
    function saveOk(b, name) {
      saveBlob(b, name || f.name);
      showToast(DLICON + '<span>Downloading <span class="tmono">' + escapeHtml(name || f.name || "file") + '</span>…</span>');
      setDlBusy(false);
    }
    // Hand a KNOWN-FRESH pre-authed URL to the browser (docs download via content-disposition; images open in
    // a tab) — the read-blocked-but-valid path.
    function handToBrowser(dl, name) {
      var a = document.createElement("a"); a.href = dl; if (name) a.download = name;
      a.target = "_blank"; a.rel = "noopener"; document.body.appendChild(a); a.click(); a.remove();
      setDlBusy(false);
    }
    function openInSharePoint() { setDlBusy(false); if (f.fileUrl) window.open(f.fileUrl, "_blank", "noopener"); }
    // Cached downloadUrl missing or stale → re-resolve a fresh one, retry the fetch, and only then fall back
    // (fresh-URL <a>, else the always-valid SharePoint page). Never reuse the expired URL.
    function recover() {
      freshDownloadUrl(f.fileUrl).then(function (r) {
        if (!r.url) throw new Error("no fresh url");
        return fetchTimeout(r.url, {}, 30000)
          .then(function (res) { if (!res.ok) throw new Error(res.status); return res.blob(); })
          .then(function (b) { saveOk(b, r.name); })
          .catch(function () { handToBrowser(r.url, r.name); });
      }).catch(openInSharePoint);
    }
    if (!f.downloadUrl) { recover(); return; } // no cached URL (e.g. a gallery item that failed to resolve)
    fetchTimeout(f.downloadUrl, {}, 30000)
      .then(function (r) { if (!r.ok) throw new Error(r.status); return r.blob(); })
      .then(function (b) { saveOk(b, f.name); })
      .catch(recover); // stale/expired/blocked → re-resolve fresh and retry
  }
  if (lbDownloadBtn) {
    setDlBusy(false);
    lbDownloadBtn.addEventListener("click", downloadCurrent);
  }

  // ---- "Share to customer": the native phone share sheet with the actual file attached ----
  // A rep's real job in the preview is to SEND the file to a customer. Copy link gives an internal-only
  // (dead-for-customers) link; Download is a clunky save on mobile. The Web Share API (Level 2, files) hands
  // the file straight to Mail/Messages/WhatsApp/etc. Shown ONLY when the device can actually share a file
  // (mobile Safari/Chrome, Windows share sheet) — on unsupported desktops it stays hidden and Download leads.
  var SHAREICON = svg('<circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><path d="M8.6 13.5l6.8 4M15.4 6.5l-6.8 4"/>');
  var lbShareBtn = document.getElementById("lbShareBtn");
  var lbShareBusy = false;
  function setShareBusy(on) {
    if (!lbShareBtn) return;
    lbShareBusy = !!on;
    lbShareBtn.disabled = !!on || (lbSel.isOn() && !lbSel.count());
    lbShareBtn.innerHTML = SHAREICON + '<span>' + (on ? "Preparing…" : window.PF.selectionLabel("Share to customer", "Share selected", lbSel, MODEL_SHARE_CAP)) + '</span>';
  }
  function shareCurrent() {
    if (lbSel.isOn()) { // Share selected: the picked files through the same core as the whole-model share
      var picks = lbSel.pick(lbGalKeys || []).map(function (i) { return lbGalEntries[i]; });
      if (!picks.length) return;
      var plan = window.PF.selectModelShareFiles(picks, { cap: MODEL_SHARE_CAP, renderResources: [] });
      shareFileEntries(plan.files, lb.querySelector(".lb-title").textContent, function (on) { setShareBusy(on); },
        plan.setAside > 0 ? ' Only the first ' + MODEL_SHARE_CAP + ' were attached — use Download selected for the rest.' : '');
      return;
    }
    var f = lbFile;
    if (!f || (!f.downloadUrl && !f.fileUrl)) return;
    if (!(navigator.canShare && navigator.share)) { downloadCurrent(); return; } // no Web Share → the download path
    setShareBusy(true);
    function got(blob, name) {
      var file = null;
      try { file = new File([blob], name || f.name || "file", { type: blob.type || "application/octet-stream" }); } catch (e) {}
      setShareBusy(false);
      if (file && navigator.canShare({ files: [file] })) {
        navigator.share({ files: [file], title: name || f.name || "" }).catch(function (err) { shareRejected(err, f.name || ""); });
      } else { saveBlob(blob, name || f.name); } // device can't share files → just save it
    }
    function fail() { setShareBusy(false); downloadCurrent(); } // blob unreadable → fall back to the robust download path
    function viaFresh() {
      freshDownloadUrl(f.fileUrl).then(function (r) {
        if (!r.url) return fail();
        fetchTimeout(r.url, {}, 30000).then(function (res) { if (!res.ok) throw new Error("download " + res.status); return res.blob(); }).then(function (b) { got(b, r.name); }).catch(fail);
      }).catch(fail);
    }
    if (!f.downloadUrl) { viaFresh(); return; }
    fetchTimeout(f.downloadUrl, {}, 30000).then(function (r) { if (!r.ok) throw new Error("download " + r.status); return r.blob(); }).then(function (b) { got(b, f.name); }).catch(viaFresh);
  }
  if (lbShareBtn && navigator.canShare && navigator.share) {
    lbShareBtn.hidden = false; // this device can share files — reveal the primary action
    setShareBusy(false);
    lbShareBtn.addEventListener("click", shareCurrent);
    var lbOpenLink = document.getElementById("lbOpenLink");
    var touch = window.matchMedia && window.matchMedia("(pointer: coarse)").matches;
    if (touch && lbOpenLink && lbMoreMenu) lbMoreMenu.insertBefore(lbOpenLink, lbMoreMenu.firstChild);
  }

  // ---- "Share to customer" for a WHOLE model — send every (document) file for it in one share sheet ----
  // A phone-safe number of the model's actual files: documents first, then explicit render images, capped
  // so the iOS share sheet isn't overloaded (auto-updating render FOLDERS are set aside — the toast points
  // the rep to Download all for those). Built only on devices that can share files (see buildResources).
  var MODEL_SHARE_CAP = 9;
  function gatherModelShareEntries(m) {
    var entries = [];
    RESOURCE_TYPES.forEach(function (r) {
      if (naFor(m, r.name)) return; // skip resource types marked not-applicable for this model
      linksFor(m, r.name).forEach(function (l) { if (l && l.url) entries.push({ url: l.url, name: l.label || null, resource: r.name }); });
    });
    return entries;
  }
  function setModelShareBusy(btn, on, n) {
    if (!btn) return; btn.disabled = !!on || (modelSel.isOn() && !modelSel.count());
    btn.innerHTML = SHAREICON + '<span>' + (on ? ("Preparing" + (n ? " " + n + " file" + (n === 1 ? "" : "s") : "") + "…") : window.PF.selectionLabel("Share to customer", "Share selected", modelSel)) + '</span>';
  }
  function modelShareCustomerBtn(m) {
    var b = document.createElement("button"); b.type = "button"; b.className = "share-pack res-share-cust";
    b.title = "Send this model's files to a customer"; b.setAttribute("aria-label", "Share " + m + "'s files to a customer");
    b.innerHTML = SHAREICON + '<span>Share to customer</span>';
    b.addEventListener("click", function () { shareModelToCustomer(m, b); });
    return b;
  }
  // The share sheet needs a live user gesture. Preparing the files takes seconds (resolve + download each
  // one), and on a phone network that outlives the browser's transient activation — so share() rejects with
  // NotAllowedError. Both call sites used to swallow every rejection as "the user tapped cancel", so the
  // button simply reset and reps concluded the feature was broken. AbortError IS the real cancel; anything
  // else has to say what happened and what to do about it.
  function shareRejected(err, label) {
    if (err && err.name === "AbortError") return; // the person closed the share sheet — nothing to report
    showError('<span>Your phone closed the share sheet while ' + (label ? escapeHtml(label) + "’s files were" : "the file was") +
      ' being prepared. Tap Share to customer again — it’s quicker the second time — or use Download.</span>', 8000);
  }
  // The share core both surfaces use: resolve each entry to a fresh downloadUrl, fetch it as a File, hand
  // the set to the share sheet (dropping trailing files if the OS refuses the whole set). `setBusy(on, n)`
  // paints whichever button started it; `note` is appended to the success toast.
  function shareFileEntries(files, title, setBusy, note) {
    if (!(navigator.canShare && navigator.share)) return Promise.resolve();
    if (!files.length) { showToast('<span>No files to share for ' + escapeHtml(title) + ' yet.</span>'); return Promise.resolve(); }
    if (!window.PF_getToken) { showToast('<span>Please sign in again to share.</span>'); return Promise.resolve(); }
    setBusy(true, files.length);
    return window.PF_getToken().then(function (token) {
      var bearer = "Bearer " + token;
      // Resolve each file to a fresh downloadUrl (galResolve) then fetch it as a File. A file that won't
      // resolve is dropped (null) rather than failing the whole share.
      return Promise.all(files.map(function (e) {
        return galResolve(e.url, bearer).then(function (r) {
          if (!r || !r.full) return null;
          return fetchTimeout(r.full, {}, 30000).then(function (res) { if (!res.ok) throw new Error("share " + res.status); return res.blob(); })
            .then(function (b) { try { return new File([b], r.name || e.name || "file", { type: b.type || "application/octet-stream" }); } catch (e2) { return null; } });
        }).catch(function () { return null; });
      }));
    }).then(function (objs) {
      setBusy(false);
      var fl = objs.filter(Boolean), prepared = fl.length;
      if (!fl.length) { showError('<span>Couldn’t prepare these files to share. Open a resource and use Download instead.</span>', 7000); return; }
      // If the OS rejects the whole set (too many/large), drop trailing files until it accepts.
      while (fl.length && !navigator.canShare({ files: fl })) fl.pop();
      if (!fl.length) { showError('<span>These files are too large to share at once — send them a few at a time with Download.</span>', 7000); return; }
      var trimmed = fl.length < prepared;
      navigator.share({ files: fl, title: title }).then(function () {
        var extra = note || (trimmed ? ' Some files were too many to attach — use Download for the rest.' : '');
        showToast(svg(ICONS.ext) + '<span>Shared ' + fl.length + ' file' + (fl.length === 1 ? '' : 's') + ' for ' + escapeHtml(title) + '.' + extra + '</span>', extra ? 9000 : undefined);
      }).catch(function (err) { shareRejected(err, title); }); // a user cancel is fine; a lapsed gesture is not
    }).catch(function () {
      setBusy(false);
      showError('<span>Couldn’t share these files just now. Open a resource and use Download instead.</span>', 7000);
    });
  }
  function shareModelToCustomer(m, btn) {
    if (!(navigator.canShare && navigator.share)) return;
    var entries = gatherModelShareEntries(m);
    if (modelSel.isOn()) entries = entries.filter(function (e) { return modelSel.has(e.resource); }); // Share selected: picked rows only
    var plan = window.PF.selectModelShareFiles(entries, { cap: MODEL_SHARE_CAP, renderResources: ["Renders"] });
    var note = plan.setAside > 0 ? ' Renders weren’t included — open Renders and use “Download all” to send those images.' : '';
    shareFileEntries(plan.files, m, function (on, n) { setModelShareBusy(btn, on, n); }, note);
  }

  // ---- download EVERY file in a multi-file preview (renders / multi-doc) so a rep can attach them to an email ----
  var DLALLICON = svg('<path d="M12 3v9"/><path d="m8 9 4 3 4-3"/><path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/>');
  var lbDownloadAllBtn = document.getElementById("lbDownloadAllBtn");
  var lbAllFiles = null; // set by buildGallery for a 2+ file set; null for a single-file preview
  function setDownloadAllLabel(txt, busy) {
    if (!lbDownloadAllBtn) return;
    lbDownloadAllBtn.disabled = !!busy;
    lbDownloadAllBtn.innerHTML = DLALLICON + '<span>' + txt + '</span>';
  }
  // Show the button only for a 2+ file set, and store its file list — each { fileUrl (for a fresh re-resolve),
  // downloadUrl (fast path), name (saved filename) }, mirroring what show() puts in lbFile for a single item.
  function setDownloadAll(files) {
    lbAllFiles = (files && files.length >= 2) ? files : null;
    if (!lbDownloadAllBtn) return;
    lbDownloadAllBtn.hidden = !lbAllFiles;
    if (lbAllFiles) setDownloadAllLabel("Download all (" + lbAllFiles.length + ")", false);
  }
  // Paint the three header actions for the current selection state (labels + the empty-pick disable).
  function syncLbSelLabels() {
    var on = lbSel.isOn(), n = lbSel.count();
    if (lbAllFiles) setDownloadAllLabel(window.PF.selectionLabel("Download all (" + lbAllFiles.length + ")", "Download selected", lbSel), on && !n);
    if (!lbShareBusy) setShareBusy(false);
    if (lbCopyBtn) { lbCopyBtn.innerHTML = LINKICON + '<span>' + window.PF.selectionLabel("Copy link", "Copy selected links", lbSel) + '</span>'; lbCopyBtn.disabled = on && !n; }
  }
  function downloadAll() {
    if (!lbAllFiles || !lbAllFiles.length) return;
    // Download selected: the picked tiles (strip order); otherwise every file, as before.
    var files = lbSel.isOn() ? lbSel.pick(lbGalKeys || []).map(function (i) { return lbAllFiles[i]; }).filter(Boolean) : lbAllFiles.slice();
    if (!files.length) return;
    var total = files.length, saved = 0, failed = 0, unsaved = [];
    // Bound every download: without a timeout a stalled Graph URL left the button stuck on "Saving…"
    // and that state leaked into the next preview.
    function fetchBlob(url) { return fetchTimeout(url, {}, 30000).then(function (r) { if (!r.ok) throw new Error("download " + r.status); return r.blob(); }); }
    // Same self-healing chain as the single Download: cached URL → re-resolve a fresh one → hand the fresh URL
    // to the browser; a total failure opens that file's SharePoint page so nothing is silently lost.
    function one(f) {
      var attempt = f.downloadUrl ? fetchBlob(f.downloadUrl) : Promise.reject(new Error("no cached"));
      return attempt.then(function (b) { saveBlob(b, f.name); }, function () {
        return freshDownloadUrl(f.fileUrl).then(function (r) {
          if (!r.url) throw new Error("no fresh url");
          return fetchBlob(r.url).then(function (b) { saveBlob(b, r.name || f.name); }, function () {
            var a = document.createElement("a"); a.href = r.url; if (r.name || f.name) a.download = r.name || f.name; a.target = "_blank"; a.rel = "noopener"; document.body.appendChild(a); a.click(); a.remove();
          });
        });
      }).then(function () { saved++; }, function () { failed++; if (f.fileUrl) unsaved.push(f); });
    }
    // Sequential with a small gap so a rapid burst of saves doesn't trip the browser's download throttle.
    files.reduce(function (chain, f, i) {
      return chain.then(function () { setDownloadAllLabel("Downloading " + (i + 1) + "/" + total + "…", true); return one(f); })
        .then(function () { return new Promise(function (res) { setTimeout(res, 250); }); });
    }, Promise.resolve()).then(function () {
      syncLbSelLabels();
      if (!failed) {
        // First multi-file save on this device: the browser's own "allow multiple downloads?" prompt is
        // easy to miss; say it once so a silent second file doesn't read as a failure.
        var firstMulti = false;
        if (total > 1) { try { firstMulti = !localStorage.getItem("pf-multi-dl-noted"); if (firstMulti) localStorage.setItem("pf-multi-dl-noted", "1"); } catch (e) {} }
        showToast(DLALLICON + '<span>Saved ' + saved + ' of ' + total + ' file' + (total === 1 ? "" : "s") + '.' + (firstMulti ? ' Your browser may ask once to allow multiple downloads.' : '') + '</span>', firstMulti ? 8000 : undefined);
        return;
      }
      // Hand the failures back as real LINKS instead of firing window.open from here. By this point the
      // user's click is long spent, so the browser blocks these as popups — and the old toast reported
      // them as "opened in SharePoint" when nothing had opened at all. A link the user clicks carries its
      // own activation, so it always works.
      var links = unsaved.slice(0, 8).map(function (f) {
        return '<a href="' + escapeAttr(safeUrl(f.fileUrl) || "#") + '" target="_blank" rel="noopener">' + escapeHtml(f.name || "file") + '</a>';
      }).join(", ");
      showError(DLALLICON + '<span>Saved ' + saved + ' of ' + total + '. Couldn’t download ' + failed +
        ' — open ' + (failed === 1 ? "it" : "them") + ' in SharePoint: ' + links +
        (unsaved.length > 8 ? ' and ' + (unsaved.length - 8) + ' more' : '') + '.</span>', 15000);
    });
  }
  if (lbDownloadAllBtn) lbDownloadAllBtn.addEventListener("click", downloadAll);

  // React to back/forward navigation and externally pasted links.
  window.addEventListener("hashchange", function (e) {
    if (_internalHash) { _internalHash = false; return; }
    if (!pmHasStagedEdits()) { applyHash(); return; }
    // Staged edits on screen: put the address back on the view still showing while we ask (replaceState
    // fires no hashchange), then go where the user was heading only on "Discard" (F146).
    var target = location.href;
    history.replaceState(null, "", e.oldURL);
    confirmDiscardStaged().then(function (ok) {
      if (!ok) return;
      history.replaceState(null, "", target);
      applyHash();
    });
  });

  // Refresh the catalog when the user returns to a tab they left open. A viewer who never saves
  // otherwise sees the morning's links/coverage/health all day — the server's 60s SWR cache is wasted
  // client-side. Throttled to once/60s, and it only repaints the passive Find view with nothing open:
  // never yank a PM mid-edit on Progress/Admin, and never interrupt an open preview or an active search
  // (those read the freshly-loaded globals on their next action anyway).
  var _lastRevisitRefetch = 0;
  document.addEventListener("visibilitychange", function () {
    // Also retry after a FAILED cold load: gating on catalogLoaded alone meant the one state that most
    // needs a retry — the app sitting on "couldn't load, please refresh" — was the one that never retried.
    if (document.visibilityState !== "visible" || !(catalogLoaded || catalogError)) return;
    var now = Date.now();
    if (now - _lastRevisitRefetch < 60000) return;
    _lastRevisitRefetch = now;
    var wasError = catalogError;
    loadCatalog().then(function () {
      // Never repaint under someone. A modal, a typed query — or the keyboard focus sitting inside the
      // content we're about to replace: render() rebuilds #content, which drops focus to <body> and loses
      // a keyboard user's place mid-task (WCAG 2.4.3).
      var focusInside = content && document.activeElement && document.activeElement !== document.body && content.contains(document.activeElement);
      var busy = (lb && !lb.hidden) || (searchInput && searchInput.value.trim()) || focusInside;
      if (busy) return;
      if (wasError || view === "catalog") render(); // recovered from the error screen, or a normal refresh
    }).catch(function () {});
  });

  /* ============================================================
     VERSION NUDGE & BUG REPORT
     ============================================================ */
  // ---- "New version available" nudge ----
  // The Hub is a long-lived SPA: a coworker who leaves it open won't pick up a fresh Vercel deploy until the
  // page reloads (the service worker is network-first, so a plain reload is enough — no hard refresh). Watch
  // the deployed front-end assets' ETags; when the live version differs from the one present at load, the
  // running UI is stale — offer a one-click refresh. We check app.css AND app.js because many deploys are
  // CSS-only. HEAD + cache:no-store reads the live edge version (never a cached copy); the SW ignores non-GET
  // requests so these bypass it. One nudge per session; if any HEAD fails we skip (never a false alarm).
  (function versionWatch() {
    var baseline = null, nudged = false, inflight = false, lastCheck = 0;
    function etag(r) { return (r && r.ok) ? (r.headers.get("etag") || r.headers.get("last-modified") || null) : null; }
    function liveSig() {
      return Promise.all(["/app.css", "/app.js"].map(function (u) {
        return fetch(u, { method: "HEAD", cache: "no-store" }).then(etag).catch(function () { return null; });
      })).then(function (tags) { return tags.every(Boolean) ? tags.join("|") : null; }); // null → a probe failed, skip this round
    }
    function nudge() {
      if (nudged) return; nudged = true;
      showToast('<span><b>A new version is available.</b> Refresh to get the latest.</span>',
        3600000, { label: "Refresh", onClick: function () { location.reload(); } });
      // Re-arm after the toast's lifetime. `nudged` was a one-shot, so dismissing it (a click anywhere on
      // the toast dismisses) meant the tab never mentioned the new version again — and the whole point is
      // that a stale tab keeps running old code. Re-arming re-checks and nudges again if it is STILL stale.
      setTimeout(function () { nudged = false; }, 3600000 + 5000);
    }
    function check() {
      if (nudged || inflight) return;
      var now = Date.now(); if (now - lastCheck < 30000) return; lastCheck = now; inflight = true;
      liveSig().then(function (sig) {
        inflight = false;
        if (!sig) return;
        if (baseline === null) { baseline = sig; return; }   // first successful probe = the version we're running
        if (sig !== baseline) nudge();
      });
    }
    check();                            // establish the baseline shortly after load
    setInterval(check, 5 * 60 * 1000);  // then poll every 5 minutes
    document.addEventListener("visibilitychange", function () { if (document.visibilityState === "visible") check(); });
  })();

  // ---- Report a bug (floating button + modal) ----
  var BUGICON = svg('<path d="M21 14a2 2 0 0 1-2 2H8l-5 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M12 7.2v3.4"/><path d="M12 13.2h.01"/>');
  var bugFab = document.getElementById("bugFab");
  var bugModal = document.getElementById("bugModal");
  var bugText = document.getElementById("bugText");
  var bugSubmit = document.getElementById("bugSubmit");
  function previewingResource() { return (lb && !lb.hidden && lbShare && lbShare.model) ? lbShare : null; }
  function bugContext() {
    var pr = previewingResource();
    if (pr) return "Previewing " + pr.resource + " · " + pr.model;
    if (view === "progress") return "Progress view";
    if (view === "metrics") return "Metrics view";
    if (view === "admin") return "Admin view";
    if (sel.model) { var p = sel.model.split("§"); return "Model " + p[p.length - 1] + " (" + p[1] + " › " + p[2] + ")"; }
    if (sel.prod) { var q = sel.prod.split("§"); return "Product " + q[2] + " (" + q[1] + ")"; }
    if (searchInput.value.trim()) return "Search view"; // don't capture the raw query — it can contain a customer/deal name (same reason search_miss is anonymized)
    return "Home / catalog";
  }
  var bugResourceUrl = null; // set by "Flag" so the report links the exact previewed file (preview is closed before submit)
  function openBug(prefill, resourceUrl) {
    if (!bugModal) return;
    bugResourceUrl = resourceUrl || null;
    // Restore an unsent draft rather than wiping it. A backdrop mis-click used to discard a long report
    // silently, and reopening started from blank. An explicit prefill (the "report this file" path) wins.
    bugText.value = prefill || bugDraft || "";
    var ctx = document.getElementById("bugCtx");
    if (ctx) ctx.textContent = "We’ll automatically include where you are (" + bugContext() + ") and your name and email so we can follow up.";
    bugSubmit.disabled = false; bugSubmit.textContent = "Send report";
    bugModal.hidden = false;
    a11yModal.open(bugModal.querySelector(".bugm-panel"), bugText);
    if (prefill) { try { bugText.setSelectionRange(prefill.length, prefill.length); } catch (e) {} } // cursor after the marker
  }
  // "Flag" in the resource preview: report an outdated/wrong file. Reuses the bug pipeline (endpoint,
  // sheet, admin inbox) with a [Content flag] marker + the model·resource prefilled, so sales staff can
  // surface stale content without new infrastructure.
  var lbFlagBtn = document.getElementById("lbFlagBtn");
  if (lbFlagBtn) lbFlagBtn.innerHTML = FLAGICON + '<span>Flag as outdated</span>';
  if (lbFlagBtn) lbFlagBtn.addEventListener("click", function () {
    var pr = previewingResource(); // capture BEFORE closing the preview — previewingResource() reads the open lightbox
    var title = ((lb && lb.querySelector(".lb-title")) || {}).textContent || "this resource";
    var rurl = pr ? resourceShareUrl(pr.model, pr.resource) : null;
    closePreview();
    openBug("[Content flag] " + title + " — ", rurl);
  });
  // PM-only "Manage placements": open the document panel for the file currently in the preview. Revealed
  // for PMs once /api/me confirms status (see below). Derive the file + resource DIRECTLY from the live
  // lightbox (lbFile is set on every open path; the resource comes from lbShare when present, else parsed
  // from the "<model> · <resource>" title) — don't depend on previewingResource(), which only resolves
  // when the preview was opened via openResource (search/What's-New/gallery/deep-link paths don't set it).
  var lbManageBtn = document.getElementById("lbManageBtn");
  if (lbManageBtn) lbManageBtn.innerHTML = MANAGEICON + '<span>Manage placements</span>';
  if (lbManageBtn) lbManageBtn.addEventListener("click", function () {
    var fileUrl = (lbFile && lbFile.fileUrl) || null;
    var titleEl = lb && lb.querySelector(".lb-title");
    var titleTxt = (titleEl && titleEl.textContent) || "";
    var resource = (lbShare && lbShare.resource) || (titleTxt.indexOf("·") !== -1 ? titleTxt.split("·").pop().trim() : "");
    if (!fileUrl || !resource) { showToast('<span>Open a file to manage where it’s used.</span>'); return; }
    var keepShare = lbShare; // closePreview clears it; the reopened preview needs it back for Copy link
    closePreview();
    // Reopen this file's preview when the panel closes, so Cancel/Save returns the PM to what they were viewing.
    openDocPanel(fileUrl, resource, { mode: "manage", onClose: function () { lbShare = keepShare; openPreview(fileUrl, titleTxt); } });
  });
  // Whatever is in the box when the modal closes, unless it was sent. Cleared on a successful submit.
  var bugDraft = "";
  function closeBug() { if (bugModal) { bugDraft = bugText ? bugText.value : ""; bugModal.hidden = true; a11yModal.close(bugModal.querySelector(".bugm-panel")); } }
  if (bugFab) { bugFab.innerHTML = BUGICON; bugFab.addEventListener("click", function () { openBug(); }); }
  if (bugModal) {
    document.getElementById("bugClose").addEventListener("click", closeBug);
    document.getElementById("bugCancel").addEventListener("click", closeBug);
    document.getElementById("bugBackdrop").addEventListener("click", closeBug);
    var bugErr = document.getElementById("bugErr");
    // Clear the invalid state as soon as they start typing, so the error can't linger over a now-valid field.
    bugText.addEventListener("input", function () { if (bugText.value.trim()) { bugText.removeAttribute("aria-invalid"); if (bugErr) bugErr.textContent = ""; } });
    bugSubmit.addEventListener("click", function () {
      var desc = bugText.value.trim();
      if (!desc) {
        // Was silent: focus just returned to the textarea with no message, so a screen-reader user
        // believed the report sent. role="alert" announces it; aria-invalid marks the field.
        bugText.setAttribute("aria-invalid", "true");
        if (bugErr) bugErr.textContent = "Please describe what went wrong before sending.";
        bugText.focus();
        return;
      }
      bugText.removeAttribute("aria-invalid"); if (bugErr) bugErr.textContent = "";
      if (!window.PF_getToken) { showError('<span>Please sign in again.</span>'); return; }
      bugSubmit.disabled = true; bugSubmit.textContent = "Sending…";
      // While a file preview is open the address bar shows only the model (the hash reflects the
      // persistent selection, not the transient lightbox) — send the resource-scoped deep link so the
      // admin's "the page they were on" reopens the EXACT file the reporter was looking at.
      var pr = previewingResource();
      var pageUrl = pr ? resourceShareUrl(pr.model, pr.resource) : (bugResourceUrl || location.href);
      var payload = { description: desc, page_url: pageUrl, context: bugContext() };
      apiFetch("/api/report-bug", { method: "POST", body: payload }).then(function () {
        if (bugText) bugText.value = ""; // sent — drop the draft so the next report starts clean
        closeBug();
        showToast(BUGICON + '<span>Thanks! Your report went to the product team.</span>');
      }).catch(function () {
        bugSubmit.disabled = false; bugSubmit.textContent = "Send report";
        showError('<span>Couldn\'t send just now — please try again.</span>');
      });
    });
    document.addEventListener("keydown", function (e) { if (e.key === "Escape" && !bugModal.hidden) closeBug(); });
  }

  // Load the catalog (structure + SP Folder URL links) from the auth-gated live endpoint.
  // No public fallback anymore (that would leak the data) — on failure the app shows a
  // "couldn't load, please refresh" message instead of stale/baked-in data.
  var catalogStale = false; // true when the last load was served from a stale cache (server couldn't refresh)
  var catalogLoaded = false; // true once a catalog load has succeeded — a later refresh failure keeps last-good data instead of wiping

  /* ============================================================
     CATALOG LOAD & BOOTSTRAP
     ============================================================ */
  // The parser's hygiene findings, normalised so the panel can always count them. Falls back to the older
  // top-level `unknownResourceTypes` field during a deploy where the server is still the previous build.
  function hygieneFrom(data) {
    var d = (data && data.diagnostics) || {};
    var arr = function (v) { return Array.isArray(v) ? v : []; };
    return {
      unknownResourceTypes: Array.isArray(d.unknownResourceTypes) ? d.unknownResourceTypes : arr(data && data.unknownResourceTypes),
      invalidLinks: arr(d.invalidLinks),
      dupeModelNames: arr(d.dupeModelNames),
      strayLinks: arr(d.strayLinks),
      dupeResourceRows: arr(d.dupeResourceRows)
    };
  }
  // The outage panel's Retry. Delegated: the panel is re-rendered by every view, so a per-render binding
  // would have to be repeated at each call site (and was missed at one). A failed retry re-renders the
  // panel with a fresh, enabled button, so there is nothing to reset here.
  content.addEventListener("click", function (e) {
    var b = e.target && e.target.closest ? e.target.closest("#catRetry") : null;
    if (!b) return;
    b.disabled = true; b.textContent = "Loading…";
    loadCatalog();
  });
  // Newest catalog snapshot this tab applied, and the server time of its last successful save (both server
  // ISO times). /api/catalog caches per warm instance for up to 5 min, so a tab-revisit reload could hand
  // back a copy from BEFORE this PM's save; applying it reverted LINKS, and the next add-only save then
  // silently dropped the link just saved (F010/F019). An older response landing late did the same (F216).
  var catAppliedAt = "", catWriteFloor = "";
  function noteServerWrite(at) { if (at && at > catWriteFloor) catWriteFloor = at; }
  function loadCatalog(opts) {
    // Catalog structure + SharePoint links are INTERNAL. They are fetched only from the
    // auth-gated /api/catalog, authenticated with the signed-in user's Microsoft token.
    // Nothing is baked into this public page, so an external visitor can obtain nothing.
    // `fresh` (passed after a PM write) bypasses the server's 60s cache so the just-saved change is
    // read back instead of a stale cache reverting the optimistic update.
    // A fresh read follows a WRITE, so the resolved-driveItem memo may now be describing links that no
    // longer exist. Drop it and let the next panel re-resolve.
    if (opts && opts.fresh) wireResolveMemoClear();
    // apiFetch carries the status on the error: a 403 means this ACCOUNT isn't authorized (a B2B guest, or
    // an address outside the allowed domains), which no amount of refreshing fixes — the generic outage
    // screen sent those users into a refresh loop. Anything else genuinely is "try again".
    return apiFetch("/api/catalog" + (opts && opts.fresh ? "?fresh=1" : ""), { fresh: true }).then(function (data) {
      // Older than what this tab already has, or than its own last save? Keep the current (optimistic)
      // state — the next load brings the newer copy.
      if (catalogLoaded && data && window.PF.catalogSnapshotStale(data.generatedAt, catAppliedAt, catWriteFloor)) return;
      if (data && data.generatedAt) catAppliedAt = data.generatedAt;
      if (data && Array.isArray(data.catalog)) CATALOG = data.catalog;
      if (data && data.links && typeof data.links === "object") LINKS = data.links;
      if (data && data.added && typeof data.added === "object") ADDED = data.added; // drives What's New (incl. new folder files, merged server-side)
      ORPHANS = (data && Array.isArray(data.orphans)) ? data.orphans : []; // files in a shared folder matching no model (Progress flag)
      SHEET_HYGIENE = hygieneFrom(data);
      NA = {}; if (data && Array.isArray(data.na)) data.na.forEach(function (k) { NA[k] = 1; }); // not-applicable resource types
      RUNTIME_ALIASES = (data && Array.isArray(data.aliases)) ? data.aliases : []; // PM-editable search synonyms
      RUNTIME_HANDLED = (data && Array.isArray(data.handledMisses)) ? data.handledMisses : [];
      // Link health from the daily cron — visible to every user (see healthStatusFor).
      serverHealth = (data && data.health && typeof data.health === "object") ? data.health : {};
      // healthAt is the cron's last-run DATE (e.g. "2026-08-25"). Parse a bare date as LOCAL midnight — NOT
      // Date.parse, which reads "2026-08-25" as UTC midnight and in Central rendered as the PREVIOUS evening
      // ("8/24 7:00 PM"). serverHealthDateOnly tells healthWhen to show the day, not a fabricated clock time.
      serverHealthDateOnly = false;
      if (data && data.healthAt) {
        var _hm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(data.healthAt));
        if (_hm) { serverHealthTs = new Date(+_hm[1], +_hm[2] - 1, +_hm[3]).getTime(); serverHealthDateOnly = true; }
        else { serverHealthTs = Date.parse(data.healthAt) || 0; }
      } else { serverHealthTs = 0; }
      Object.keys(healthCleared).forEach(function (k) { delete serverHealth[k]; }); // keep just-fixed links clear
      var nowStale = !!(data && data.stale);
      // The server served a cached copy because its refresh failed — tell the user once (on the transition
      // into stale) that link/coverage data may be a little behind; it self-heals on the next good refresh.
      if (nowStale && !catalogStale) showToast(svg(ICONS.alert) + '<span>Showing recently-cached data — the last refresh didn’t go through. It’ll update on its own shortly.</span>', 6000, null, true);
      catalogStale = nowStale;
      catalogError = false; catalogDenied = false; // a later good load clears the not-authorized screen
      catalogLoaded = true;
      buildIndex();
    }).catch(function (err) {
      console.warn("Product Hub: catalog failed to load", err);
      // A refresh that fails AFTER a good load must NOT blank the app — a PM save triggers this reload, and
      // wiping here reads as data loss right after "Saved". Keep the last-good catalog and mark it stale; it
      // self-heals on the next good refresh. Only a cold first load (nothing to fall back to) shows the hard
      // "couldn't load" screen.
      if (catalogLoaded) {
        if (!catalogStale) showToast(svg(ICONS.alert) + '<span>Couldn’t refresh just now — showing the last loaded data. It’ll update on its own shortly.</span>', 6000, null, true);
        catalogStale = true;
      } else {
        catalogError = true; catalogDenied = (err && err.status === 403); // 403 = this account, not an outage
        CATALOG = []; LINKS = {}; NA = {}; ORPHANS = []; SHEET_HYGIENE = hygieneFrom(null); RUNTIME_ALIASES = []; serverHealth = {}; serverHealthTs = 0; buildIndex();
      }
    });
  }
  // Everything under "pf-" is PER-USER state (favorites + their tombstones, recently viewed, last visit,
  // dismissals), but it lives in a localStorage that belongs to the BROWSER. On a shared machine — a
  // showroom PC, a borrowed laptop — the next person to sign in inherited it, and favSyncFromServer then
  // treated the previous person's data as this account's: their un-pin tombstones were re-sent as
  // /api/me deletes against the NEW user's server favorites, and the previous user's local-only pins were
  // pushed up into it. So: stamp who the data belongs to and wipe it on a mismatch. Sign-out alone can't
  // do this — an expired session never passes through the sign-out handler. Theme is a device preference,
  // not personal data, so it survives.
  function claimLocalStateForCurrentUser() {
    try {
      var who = String(window.PF_user || "").toLowerCase();
      if (!who) return; // identity not known yet — never wipe on a guess
      if (localStorage.getItem("pf-owner") === who) return;
      Object.keys(localStorage).forEach(function (k) {
        if (k.indexOf("pf-") === 0 && k !== "pf-theme" && k !== "pf-owner") localStorage.removeItem(k);
      });
      localStorage.setItem("pf-owner", who);
    } catch (e) {} // storage disabled/full — the server remains the source of truth
  }

  window.__startPF = function () {
    claimLocalStateForCurrentUser(); // BEFORE anything reads pf-favs / pf-recent or syncs with the server
    renderLoading(); // show a loading state instead of a blank void while the catalog fetches
    // Ask the server whether this user is a PM (single source of truth — no client allowlist). The
    // Progress/Admin tabs start hidden in the markup; reveal them only if the server says so. Runs in
    // parallel with the catalog load; applyHash waits for both so the default view knows isPM.
    function fetchMe() {
      // Any non-2xx means we FAILED to determine PM status — it is not a "no". /api/me answers 200 or
      // 502, so a 401/403 blip used to silently strip a PM of their tools until they reloaded. apiFetch
      // rejects, which routes it into the retry below and, if that also fails, into the "couldn't confirm"
      // notice.
      return apiFetch("/api/me");
    }
    var meFailed = false;
    var meP = fetchMe()
      .catch(function () { return fetchMe(); }) // one retry — a transient blip shouldn't hide a PM's tools
      .then(function (me) { window.PF_isPM = !!(me && me.isPM); PF_ADMIN_ALERTS = (me && Array.isArray(me.alerts)) ? me.alerts : []; PF_serverFavs = (me && Array.isArray(me.favorites)) ? me.favorites : null; })
      .catch(function () { window.PF_isPM = false; PF_ADMIN_ALERTS = []; meFailed = true; })
      .then(function () {
        if (isPM()) {
          var pmTabs = document.querySelectorAll('.vtab[data-view="progress"], .vtab[data-view="metrics"], .vtab[data-view="admin"]');
          for (var i = 0; i < pmTabs.length; i++) pmTabs[i].classList.remove("vtab-pm-hidden");
          var mgBtn = document.getElementById("lbManageBtn"); if (mgBtn) mgBtn.hidden = false; // reveal "Manage placements" in the preview
        } else if (meFailed) {
          // Couldn't DETERMINE PM status (not a confirmed "no") — don't leave a real PM silently locked out
          // of their workspace; make it recoverable.
          try { showToast("Couldn’t confirm your access. If you manage products, refresh to load your tools.", 12000, { label: "Refresh", onClick: function () { location.reload(); } }, true); } catch (e) {}
        }
      });
    // First-run nudge: once, if the viewer has no favorites yet, hint that models can be pinned.
    function maybeFavNudge() {
      try {
        if (localStorage.getItem("pf-fav-nudged")) return;
        localStorage.setItem("pf-fav-nudged", "1");
        if (favGet().length) return; // already a favorites user — no nudge needed
        showToast('<span>Tip: pin models you use often with <b>☆ Add to favorites</b> — they’ll wait at the top of Find.</span>', 8000);
      } catch (e) {}
    }
    // Read the landing hash BEFORE applyHash runs. applyHash can CLEAR it — a deep link to something that
    // has since been renamed or removed falls back to the catalog and calls syncHash(), which replaceStates
    // the hash away. The remember-last-tab check below then saw an empty hash, decided this was a bare
    // landing, and jumped to Progress — directly contradicting the "showing the catalog instead" toast the
    // user had just been shown.
    var bootHash = location.hash;
    Promise.all([loadCatalog(), meP]).then(applyHash).then(function () {
      // A search typed during load showed a "loading" placeholder (not a false miss) — resolve it now.
      if (searchInput && searchInput.value.trim()) { view = "catalog"; renderSearch(searchInput.value); syncViewTabs(); return; }
      // Remember-last-tab: on a bare landing (no deep-link hash), reopen the tab last used, if it's available
      // (a deep link like #/model or #progress takes precedence and is left untouched).
      if (!bootHash || bootHash === "#") {
        var last = null; try { last = localStorage.getItem("pf-lasttab"); } catch (e) {}
        var tabEl = last && document.getElementById("tab-" + last);
        // PM tabs are hidden via the .vtab-pm-hidden CLASS (not an inline style), so a style.display
        // check never blocked them — a non-PM on a shared profile, or a PM whose /api/me failed, could
        // restore Progress/Admin/Metrics. Gate on the real hidden state AND role (render() has no isPM check).
        if (last && last !== view && tabEl && !tabEl.classList.contains("vtab-pm-hidden") && (last === "catalog" || isPM())) { view = last; render(); syncViewTabs(); syncHash(); }
      }
    }).then(favSyncFromServer).then(maybeFavNudge)
      .catch(function (e) {
        // Last-resort recovery: a throw anywhere in this startup chain used to leave the app on its
        // "Loading…" skeleton forever (and skip the favorites sync entirely). Land on a usable Find view.
        console.error("[pf] startup", (e && e.stack) || e);
        view = "catalog"; sel = { prod: null, model: null };
        try { render(); syncViewTabs(); } catch (e2) { console.error("[pf] startup render", e2); }
      });
  };
})();
