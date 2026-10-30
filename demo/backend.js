/* Demo backend: answers every /api/* and Microsoft Graph call in the browser, from demo/seed.js, so the
   public demo runs with no server and no sign-in. Response shapes mirror the real endpoints (contract in
   docs/design/specs/2026-09-22-demo-mode-design.md). Writes change this tab's copy only — a reload resets.

   createBackend(seed, {now}) is pure (Node-testable): handle(method, url, body) -> {status, json} or
   {status, file} (a seed file to stream). install(backend, fileUrlFor) wraps window.fetch in the browser. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.PF_DEMO = factory();
})(this, function () {
  "use strict";
  var RESOURCES = ["Flyer", "Product Overview", "Spec Sheet", "Cut Sheet", "Renders", "In-Store Images", "Assembly Instructions", "Graphics Dielines", "Compliance & Testing", "Warranty", "Training"];
  var DAY = 86400000;

  function clone(x) { return JSON.parse(JSON.stringify(x)); }
  function ymd(ms) { return new Date(ms).toISOString().slice(0, 10); }
  function mondayOf(ms) { var d = new Date(ymd(ms) + "T00:00:00Z"); var back = (d.getUTCDay() + 6) % 7; return d.getTime() - back * DAY; }
  function urlsOf(v) { if (!v) return []; return Array.isArray(v) ? v.map(function (x) { return x.url; }) : [v]; }
  function distinct(urls) { var s = {}; urls.forEach(function (u) { s[u] = 1; }); return Object.keys(s).length; }
  function sameSet(a, b) { if (a.length !== b.length) return false; var s = {}; a.forEach(function (x) { s[x] = 1; }); return b.every(function (x) { return s[x]; }); }
  function decodeShare(id) {
    var b = id.slice(2).replace(/-/g, "+").replace(/_/g, "/"); b += "===".slice((b.length + 3) % 4);
    if (typeof atob === "function") return decodeURIComponent(escape(atob(b)));
    return Buffer.from(b, "base64").toString("utf8");
  }
  function ok(json) { return { status: 200, json: json }; }
  function err(status, json) { return { status: status, json: json }; }

  function createBackend(seed, opts) {
    var now = (opts && opts.now) || function () { return Date.now(); };
    var t0 = now();
    var s = clone(seed);
    var models = {};
    s.catalog.forEach(function (c) { c.families.forEach(function (f) { f.products.forEach(function (p) { p.models.forEach(function (m) { models[m] = { family: f.name, product: p.name }; }); }); }); });
    // Dates relative to "now", so the demo always looks recently used.
    s.added = {}; s.added["Hoardmaster 3000::Flyer"] = ymd(t0 - DAY); s.added["Gemtumbler 2::Renders"] = ymd(t0 - 2 * DAY);
    s.placements.review.concat(s.placements.pending).forEach(function (p, i) { p.at = new Date(t0 - (i + 1) * 5 * 3600000).toISOString(); });
    s.bugReports.forEach(function (r, i) { r.created_at = new Date(t0 - (i + 1) * 2 * DAY).toISOString(); });
    s.favorites = []; s.lastWrite = 0; s.nextId = 10000;
    var ids = {}, idN = 0;
    function idFor(url) { if (!ids[url]) ids[url] = "demo-" + (++idN); return ids[url]; }

    function coveragePct() {
      var linked = 0, total = 0, na = {};
      s.na.forEach(function (k) { na[k] = 1; });
      Object.keys(models).forEach(function (m) {
        RESOURCES.forEach(function (r) { var k = m + "::" + r; if (na[k]) return; total++; if (s.links[k] && s.health[k] !== "broken") linked++; });
      });
      return { pct: Math.round((linked / total) * 1000) / 10, linked: linked, total: total };
    }

    function catalog() {
      return ok({
        catalog: s.catalog, links: s.links, added: s.added, na: s.na, health: s.health, healthAt: new Date(t0 - 4 * 3600000).toISOString(),
        healthMoved: s.healthMoved || {}, healthRepaired: s.healthRepaired || {},
        aliases: s.aliases, handledMisses: s.handledMisses, orphans: s.orphans,
        diagnostics: { unknownResourceTypes: [], invalidLinks: [], strayLinks: [], dupeResourceRows: [], dupeModelNames: [] },
        generatedAt: new Date(Math.max(now(), s.lastWrite)).toISOString(), source: "Demo data (fictional)"
      });
    }

    // The same refusals as the real /api/wire-links (F080): a bad request, a value that is not a link (all or nothing),
    // no model of that product, a compare-and-swap conflict, and removals the request did not declare.
    function splitUrls(v) { return String(v).split(/\s+(?=https?:\/\/)/).map(function (x) { return x.trim(); }).filter(Boolean); }
    function wire(body) {
      var a = (body && body.assignments) || null, res = body && body.resource, base = body && body.base;
      var fam = String((body && body.family) || "").trim(), prod = String((body && body.product) || "").trim();
      if (!fam || !prod || RESOURCES.indexOf(res) === -1 || !a || typeof a !== "object") return err(400, { error: "Need family, product, a valid resource type, and per-model assignments." });
      var bad = Object.keys(a).filter(function (m) { var v = String(a[m] || "").trim(); return !(v === "" || v === "N/A" || splitUrls(v).every(function (u) { return /^https?:\/\//i.test(u); })); });
      if (bad.length) return err(400, { error: "Not a link — each value must be a full https:// address (" + bad.slice(0, 4).join(", ") + "). Nothing was saved." });
      var names = Object.keys(a).filter(function (m) { return models[m] && models[m].family === fam && models[m].product === prod; });
      if (!names.length) return err(404, { error: "No matching rows found for that product/resource." });
      var conflicted = base ? names.filter(function (m) { return base[m] && !sameSet(urlsOf(s.links[m + "::" + res]), base[m]); }) : [];
      if (conflicted.length) return err(409, { error: "Another PM changed this since you opened it — reload and try again.", code: "conflict", models: conflicted });
      // Shrink guard: distinct links each model would lose, against what the request declared (pooled or per model).
      var lost = {}, removed = 0;
      names.forEach(function (m) { var v = String(a[m] || "").trim(), after = (v === "" || v === "N/A") ? [] : splitUrls(v), before = urlsOf(s.links[m + "::" + res]); var n = Math.max(0, distinct(before) - distinct(after)); if (n) { lost[m] = n; removed += n; } }); // net, as planWireOps counts
      var byModel = body && body.expectRemovedByModel, allowed = Math.max(0, Number(body && body.expectRemoved) || 0);
      if (byModel && typeof byModel === "object") {
        allowed = Object.keys(byModel).reduce(function (t, k) { return t + Math.max(0, Number(byModel[k]) || 0); }, 0);
        var over = Object.keys(lost).filter(function (m) { return lost[m] > Math.max(0, Number(byModel[m]) || 0); });
        if (over.length) return err(409, { error: "This would remove links that weren't part of your change. Nothing was saved. Reload and try again.", code: "would_shrink", removed: removed, declared: allowed, models: over });
      }
      if (removed > allowed) return err(409, { error: "This would remove " + removed + " link" + (removed === 1 ? "" : "s") + " that weren't part of your change — most likely the catalog changed since this page loaded. Nothing was saved. Reload and try again.", code: "would_shrink", removed: removed, declared: allowed, models: Object.keys(lost) });
      var today = ymd(now());
      names.forEach(function (m) {
        var k = m + "::" + res, v = String(a[m] || "").trim(), naAt = s.na.indexOf(k);
        if (naAt !== -1) s.na.splice(naAt, 1);
        delete s.health[k];
        if (!v) { delete s.links[k]; return; }
        if (v === "N/A") { delete s.links[k]; s.na.push(k); return; }
        var list = splitUrls(v);
        var before = urlsOf(s.links[k]);
        s.links[k] = list.length === 1 ? list[0] : list.map(function (x) { return { label: null, url: x }; });
        if (list.some(function (x) { return before.indexOf(x) === -1; })) s.added[k] = today;
      });
      s.lastWrite = now();
      return ok({ rejected: 0, updated: names.length, models: names, writtenAt: new Date(s.lastWrite).toISOString() });
    }

    function usage() {
      var mon = mondayOf(now()), trend = [], opensByWeek = {}, coverageTrend = [], cur = coveragePct();
      for (var i = 11; i >= 0; i--) {
        var wk = ymd(mon - i * 7 * DAY), opens = 18 + (11 - i) * 4 + ((i * 7) % 5);
        opensByWeek[wk] = opens;
        if (i < 8) trend.push({ week: wk, opens: opens, people: 5 + Math.round((11 - i) * 0.9) });
        coverageTrend.push({ week: wk, date: ymd(mon - i * 7 * DAY + 2 * DAY), pct: Math.round(cur.pct * (0.55 + 0.45 * ((11 - i) / 11)) * 10) / 10, linked: 0, total: cur.total });
      }
      return ok({
        topMisses: [{ query: "gold polish", count: 6 }, { query: "knight repellent spray", count: 3 }, { query: "hot rock", count: 3 }, { query: "egg warmer", count: 2 }, { query: "dragon door", count: 2 }, { query: "rug", count: 2 }],
        openedRecent: ["Hoardmaster 3000 · Spec Sheet", "Hoardmaster 3000 · Renders", "Snugrock King · Warranty", "Vaultvane 100 · Assembly Instructions", "Pebblechime 5 · Cut Sheet"],
        topOpened: [{ resource: "Hoardmaster 3000 · Spec Sheet", count: 41 }, { resource: "Hoardmaster 3000 · Renders", count: 33 }, { resource: "Vaultvane 100 · Assembly Instructions", count: 24 }, { resource: "Snugrock Classic · Spec Sheet", count: 19 }, { resource: "Pebblechime 5 · Cut Sheet", count: 16 }, { resource: "Decoy Princess Tower · Renders", count: 12 }, { resource: "Ashweave Hearth · Spec Sheet", count: 9 }],
        topWip: [{ resource: "Cinderpuff Bolster · Spec Sheet", count: 7 }, { resource: "Trophy Rack Six · Training", count: 6 }, { resource: "Goblet Rack Twelve · Renders", count: 5 }, { resource: "Hoardmaster 5000 D · Assembly Instructions", count: 4 }],
        previewFails: [{ resource: "Moat-in-a-Box 20 ft · Spec Sheet · 404", count: 5, lastDate: ymd(now() - DAY) }, { resource: "Clawpost Classic · Warranty · 403", count: 1, lastDate: ymd(now() - 6 * DAY) }],
        topFuzzy: [{ query: "hordmaster", closest: "Hoardmaster 3000", count: 2 }],
        shares: { count: 27, files: 84 },
        shareFails: [{ what: "Hoardmaster 3000 · Renders · NotAllowedError", count: 2, lastDate: ymd(now() - 2 * DAY) }, { what: "Vaultvane 100 · 3 files · download", count: 1, lastDate: ymd(now() - 9 * DAY) }],
        clientErrors: [],
        signinFails: [{ what: "AADSTS50105", count: 1, lastDate: ymd(now() - 4 * DAY) }],
        failTotals: { share: 3, client: 0, signin: 1 },
        totalMiss: 23, missDays: 60, openDays: 30, opens: 412, activePeople: 21, signins: 96,
        trend: trend, trendWeeks: 8, opensByWeek: opensByWeek,
        opensByModel: { "Hoardmaster 3000": 96, "Hoardmaster 3000 X": 31, "Vaultvane 100": 44, "Snugrock Classic": 29, "Snugrock King": 21, "Pebblechime 5": 24, "Ashweave Hearth": 18, "Decoy Princess Tower": 17, "Gemtumbler 2": 14, "Clawpost Classic": 11, "Trophy Rack Six": 9, "Moat-in-a-Box 40 ft": 7, "Crownpolisher 1": 6 },
        opensByType: { "Spec Sheet": 148, "Renders": 86, "Cut Sheet": 44, "Assembly Instructions": 38, "Warranty": 31, "Product Overview": 24, "Flyer": 19, "Training": 12, "In-Store Images": 10 },
        coverageTrend: coverageTrend
      });
    }

    function driveItem(url) {
      var f = s.files[url];
      if (f) {
        var dl = "demo-file:" + f.name, isImg = /^image\//.test(f.mime);
        return {
          id: idFor(url), name: f.name, webUrl: url, file: { mimeType: f.mime }, size: 1024,
          createdDateTime: new Date(t0 - 3 * DAY).toISOString(), lastModifiedDateTime: new Date(t0 - 3 * DAY).toISOString(),
          "@microsoft.graph.downloadUrl": dl,
          thumbnails: isImg ? [{ small: { url: dl, width: 96, height: 72 }, medium: { url: dl, width: 176, height: 132 } }] : []
        };
      }
      if (s.folders[url]) return { id: idFor(url), name: decodeURIComponent(url.split("/").pop()), webUrl: url, folder: { childCount: s.folders[url].length } };
      return null;
    }

    function graph(method, url) {
      var m = /\/v1\.0\/shares\/(u![^/?]+)\/driveItem(\/children|\/content|\/preview)?/.exec(url);
      if (!m) return err(400, { error: { code: "invalidRequest" } });
      var target = decodeShare(m[1]), sub = m[2] || "";
      var notFound = err(404, { error: { code: "itemNotFound", message: "The resource could not be found." } });
      // Phones use Microsoft's embed viewer (iOS ignores fit-to-width in a PDF iframe); the demo has no such
      // viewer, so it hands back the file itself for the iframe.
      if (sub === "/preview") return s.files[target] ? ok({ getUrl: "demo-file:" + s.files[target].name }) : notFound;
      if (sub === "/children") {
        if (s.folders[target]) return ok({ value: s.folders[target].map(driveItem) });
        // Any other folder that holds demo files (Flyers, Spec Sheets…) lists them, as SharePoint would: answering 404
        // for those left a permanent "⚠ 11 folders couldn't be read" on the demo's Progress (F050).
        var base = String(target).replace(/\/+$/, "") + "/";
        var inside = Object.keys(s.files).filter(function (u) { return u.indexOf(base) === 0 && u.slice(base.length).indexOf("/") === -1; });
        // A folder inside the demo library with no demo files in it (the parent of the deliberately broken link) lists
        // as empty: the broken FILE still fails its own lookup, and the scan shows no error a visitor can't clear.
        if (inside.length || /\/Shared%20Documents\//.test(base)) return ok({ value: inside.map(driveItem) });
        return notFound;
      }
      if (sub === "/content") return s.files[target] ? { status: 200, file: s.files[target].name } : notFound;
      var item = driveItem(target);
      return item ? ok(item) : notFound;
    }

    function handle(method, url, body) {
      method = String(method || "GET").toUpperCase();
      if (/^https:\/\/graph\.microsoft\.com\//.test(url)) return graph(method, url);
      var p = new URL(url, "https://demo.local"), path = p.pathname;
      if (path === "/api/catalog") return catalog();
      if (path === "/api/me") {
        if (method === "GET") return ok({ email: s.user.email, name: s.user.name, isPM: true, commit: "demo", favorites: s.favorites.slice(), alerts: [] });
        var act = body && body.action, model = body && body.model;
        if (act === "add") { if (s.favorites.indexOf(model) !== -1) return ok({ ok: true, added: false, reason: "duplicate" }); s.favorites.push(model); return ok({ ok: true, added: true }); }
        if (act === "remove") { var n0 = s.favorites.length; s.favorites = s.favorites.filter(function (x) { return x !== model; }); return ok({ ok: true, removed: n0 - s.favorites.length }); }
        // Only the named person's favorites: the demo has one person, so anyone else has none to erase (F080).
        if (act === "erase-favorites") { if (String(body.email || "").trim().toLowerCase() !== String(s.user.email).toLowerCase()) return ok({ ok: true, email: body.email, removed: 0 }); var n1 = s.favorites.length; s.favorites = []; return ok({ ok: true, email: body.email, removed: n1 }); }
        return err(400, { error: "Unknown action" });
      }
      if (path === "/api/wire-links") return wire(body);
      if (path === "/api/usage-insights") return usage();
      if (path === "/api/placements") {
        if (method === "GET") return ok({ review: s.placements.review, pending: s.placements.pending, skip: s.placements.skip, noAuto: s.placements.noAuto });
        var pu = body && body.url, pa = body && body.action;
        if (!pu || ["keep", "remove", "dismiss", "undismiss"].indexOf(pa) === -1) return err(400, { error: "Bad request" });
        if (pa === "undismiss") s.placements.skip = s.placements.skip.filter(function (u) { return u !== pu; });
        var drop = function (x) { return x.url !== pu; };
        s.placements.review = s.placements.review.filter(drop);
        if (pa !== "keep") s.placements.pending = s.placements.pending.filter(drop);
        if (pa === "remove") s.placements.noAuto.push(pu);
        if (pa === "dismiss") s.placements.skip.push(pu);
        return ok({ ok: true, action: pa, status: pa === "keep" ? "kept" : pa === "remove" ? "removed" : pa === "undismiss" ? "open" : "dismissed", existed: true });
      }
      if (path === "/api/place-cron") {
        var sugg = s.placements.review.map(function (r) { return r.file + " → " + r.model + " · " + r.resource; });
        if (/(^|&)dry=1(&|$)/.test(p.search.slice(1))) return ok({ ok: true, dryRun: true, at: new Date(now()).toISOString(), watched: Object.keys(s.folders).length, listed: Object.keys(s.folders).length, foldersFailed: 0, foldersSkippedForTime: 0, wouldSuggest: sugg.length, wouldSuggestCapped: 0, suggestions: sugg, pending: s.placements.pending.length, pendingFiles: [], alreadyRuledOut: { dismissed: s.placements.skip.length, noAuto: s.placements.noAuto.length } });
        return ok({ ok: true, at: new Date(now()).toISOString(), degraded: [], suggested: 0, suggestions: [] });
      }
      if (path === "/api/bug-reports") {
        if (method === "GET") return ok({ reports: s.bugReports });
        var bid = body && String(body.id);
        if (method === "PATCH") { s.bugReports.forEach(function (r) { if (r.id === bid) r.status = body.status; }); return ok({ ok: true }); }
        if (method === "DELETE") { s.bugReports = s.bugReports.filter(function (r) { return r.id !== bid; }); return ok({ ok: true }); }
        return err(405, { error: "Method not allowed" });
      }
      if (path === "/api/report-bug") {
        if (!body || !String(body.description || "").trim()) return err(400, { error: "Please describe the problem." });
        s.bugReports.unshift({ id: String(++s.nextId), created_at: new Date(now()).toISOString(), description: body.description, reporter_name: s.user.name, reporter_email: s.user.email, context: body.context || "", page_url: body.page_url || "", status: "new" });
        return ok({ ok: true });
      }
      if (path === "/api/aliases") {
        if (method === "GET") return ok({ aliases: s.aliases });
        if (body && body.action === "dismiss-miss") { var q = String(body.query || "").trim(); if (!q) return err(400, { error: "Missing query" }); var had = s.handledMisses.indexOf(q) !== -1; if (!had) s.handledMisses.push(q); return ok({ handled: true, added: !had, query: q }); }
        if (body && body.action === "undismiss-miss") { var uq = String(body.query || "").trim(); if (!uq) return err(400, { error: "Nothing to undo." }); var uBefore = s.handledMisses.length; s.handledMisses = s.handledMisses.filter(function (x) { return x !== uq; }); return ok({ handled: false, removed: uBefore - s.handledMisses.length, query: uq }); }
        var from = String((body && body.from) || "").trim().toLowerCase(), to = String((body && body.to) || "").trim().toLowerCase();
        if (!from || !to) return err(400, { error: "Both words are needed." });
        var alias = { from: from, to: to };
        if (s.aliases.some(function (x) { return x.from === from && x.to === to; })) return ok({ added: false, reason: "duplicate", alias: alias });
        s.aliases.push(alias); return ok({ added: true, alias: alias });
      }
      if (path === "/api/event") return ok({ ok: true });
      return err(404, { error: "Not part of the demo" });
    }

    return { handle: handle, state: s };
  }

  // Browser only: route /api/* and Graph through the backend; everything else (the app's own files,
  // fonts, the demo files) goes to the network as usual. `demo-file:<name>` markers in a JSON answer
  // become real paths before app.js sees them, so <img src>/<a href> work directly.
  function install(backend, fileUrlFor) {
    var realFetch = window.fetch.bind(window);
    window.fetch = function (input, init) {
      var url = typeof input === "string" ? input : input.url;
      var method = String((init && init.method) || (input && input.method) || "GET").toUpperCase();
      if (url.indexOf("demo-file:") === 0) return realFetch(fileUrlFor(url.slice(10)), init);
      var isApi = /^\/api\//.test(url) || url.indexOf(location.origin + "/api/") === 0;
      var isGraph = url.indexOf("https://graph.microsoft.com/") === 0;
      if (!isApi && !isGraph) return realFetch(input, init);
      var body; try { body = init && init.body ? JSON.parse(init.body) : undefined; } catch (e) { body = undefined; }
      var out;
      try { out = backend.handle(method, isApi ? url.replace(location.origin, "") : url, body); }
      catch (e) { out = { status: 500, json: { error: "Demo backend error: " + (e && e.message) } }; }
      if (out.file) return realFetch(fileUrlFor(out.file));
      var text = JSON.stringify(out.json === undefined ? {} : out.json).replace(/demo-file:([^"]+)/g, function (_, n) { return fileUrlFor(n); });
      // A short delay keeps loading states honest (and visible) instead of flashing.
      return new Promise(function (resolve) {
        setTimeout(function () { resolve(new Response(text, { status: out.status, headers: { "Content-Type": "application/json" } })); }, 60);
      });
    };
  }

  return { createBackend: createBackend, install: install };
});
