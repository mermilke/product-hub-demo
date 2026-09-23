// Shared, SIDE-EFFECT-FREE helpers used by BOTH the browser SPA (index.html, via <script src>) and the
// serverless link-health code (lib/*, via require). Pure logic only: model-name matching (the company-
// wide HMX/GTX/SNX naming convention, which grows as product lines change) and link-health error
// classification. Keeping ONE copy stops the browser and server from drifting — they once did, which
// made the manual "Check all links" flag transient 429s as "broken" while the cron correctly said
// "unknown".
//
// Dual target: module.exports for Node (require), window.PF for the browser (<script src>). Contains NO
// secrets or PII, so it is intentionally served publicly at /pf-shared.js (unlike lib/*, which is now
// blocked from static serving).
(function (factory) {
  var api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  if (typeof window !== "undefined") window.PF = api;
})(function () {
  "use strict";

  // Canonical resource-type names (RSS values), in display order. ONE shared copy — the browser reads
  // window.PF.RESOURCES, the serverless handlers require() it — so the vocabulary can't drift across the
  // places it used to be hand-copied (the two API handlers each had their own "must match" array).
  var RESOURCES = ["Flyer", "Product Overview", "Spec Sheet", "Cut Sheet", "Renders", "In-Store Images", "Assembly Instructions", "Compliance & Testing", "Warranty", "Training"];

  // Matcher for a coded model (e.g. "HMX 400 I"): "<prefix> <number> <letters>". The negative lookahead
  // on a letter keeps "400 I" from matching "400 IR"/"400 IPO" (and "600 I" from "600 IC"), while still
  // allowing view suffixes like "IPO-1". "Mini" models don't fit the clean letter-run shape → return
  // null so matchesModel falls back to the name-prefix match.
  // Build a case-INSENSITIVE literal without the /i flag (so it can be combined with case-SENSITIVE
  // parts in one pattern): each letter becomes [Aa], other chars are regex-escaped.
  function ciLiteral(s) {
    return String(s).split("").map(function (ch) {
      return /[a-z]/i.test(ch) ? "[" + ch.toUpperCase() + ch.toLowerCase() + "]" : ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }).join("");
  }
  function modelRe(model) {
    var code = String(model).replace(/[^A-Za-z0-9]/g, "").toUpperCase();
    var mm = code.match(/^(HMX|GTX|SNX)(\d+)([A-Z]*)$/);
    if (!mm || /MINI/.test(mm[3])) return null;
    if (mm[3]) return new RegExp(mm[1] + "\\s*" + mm[2] + "[\\s\\-]*" + mm[3] + "(?![A-Za-z])", "i");
    // Base model with NO letter suffix (e.g. "GTX 600"). The old pattern ended in [\s\-]*(?![A-Za-z]),
    // and since [\s\-]* can match zero chars the lookahead passed at the space after the number — so
    // "GTX 600" wrongly matched a suffixed sibling's file ("GTX 600 I ...", "VPP600-I-open"). Instead,
    // reject a following variant CODE = 1-4 UPPERCASE letters standing alone (I, D, R, PO, IR, IDPO, ...).
    // Case-SENSITIVE (no /i flag; the prefix is made case-insensitive via ciLiteral) so a lowercase
    // descriptive word ("render", "open") is NOT mistaken for a code. Trailing (?![A-Za-z0-9]) keeps
    // "600" from matching "6000"/"600A".
    // …and reject a following DECODED variant word too. The app spells suffixes out for people ("GTX 600 I"
    // is shown as "Integrated"), and PMs name files the same way — so "GTX 600 Integrated.png" was landing
    // on the plain GTX 600 as a confident match. The uppercase-code rule above can't catch it: "I" is
    // followed by more letters, so the code alternative fails and the lookahead passes.
    var decoded = ["integrated", "deep", "reduced", "pullout", "wide", "center", "centre", "left", "mini", "riser"]
      .map(ciLiteral).join("|");
    return new RegExp(ciLiteral(mm[1]) + "\\s*" + mm[2] + "(?![\\s\\-]*(?:[A-Z]{1,4}(?![A-Za-z])|(?:" + decoded + ")(?![A-Za-z])))(?![A-Za-z0-9])");
  }
  function normModel(s) { return String(s).toLowerCase().replace(/\.[a-z0-9]+$/i, "").replace(/[^a-z0-9]/g, ""); }
  function matchesModel(name, model) {
    var re = modelRe(model);
    if (re) return re.test(name); // coded part numbers (strict — keeps 400I out of 400IR etc.)
    var m = normModel(model);      // descriptive models: normalized filename must START with the model name
    return m.length > 0 && normModel(name).indexOf(m) === 0;
  }
  // Normalize a sheet cell's text. End-trimming is not enough: a name pasted from Word or Excel can carry
  // a NON-BREAKING space (or a doubled space) INSIDE it. It looks identical in Smartsheet, but the app keys
  // every link on the model NAME, so one model silently became two with its links split between them. The
  // reader and the writer must both call this or a wire 404s on a name the parser normalized. Lives here,
  // in the shared pure module, because both sides depend on it agreeing exactly.
  function normCell(v) {
    return String(v == null ? "" : v).replace(/[   ]/g, " ").replace(/\s+/g, " ").trim();
  }
  // ---- SharePoint URL intake: ONE parser for every shape the app meets --------------------------------
  //   short share link   …/:f:/s/Site/<id>                       letter f = folder; b/w/x/p/i/v/u = a file
  //   folder listing     …/Forms/AllItems.aspx?id=<folder path>  an address-bar copy while looking at a folder
  //   Office web link    …/_layouts/15/Doc.aspx?sourcedoc={GUID}&file=Name.pptx   identity lives in the QUERY
  //   plain path         …/Renders/HMX 200.png  or  …/Renders     the last segment's extension decides
  // The four helpers below are views over this one parse, so a new shape is taught in one place.
  function parseSpUrl(url) {
    var raw = String(url || "");
    var noHash = raw.split("#")[0], qi = noHash.indexOf("?");
    var path = (qi === -1 ? noHash : noHash.slice(0, qi)).replace(/\/+$/, "");
    var params = {};
    (qi === -1 ? "" : noHash.slice(qi + 1)).split("&").forEach(function (kv) {
      var i = kv.indexOf("="); if (i > 0) params[kv.slice(0, i).toLowerCase()] = kv.slice(i + 1);
    });
    var origin = raw.split("/").slice(0, 3).join("/");
    // SharePoint's own "Copy link" makes a SHORT share link whose last segment is an opaque id with no
    // extension — the letter IS the answer, so read it before the extension test (which would call every
    // one a folder: that skipped Graph resolution, listed files as folders in link health, and opened a
    // single file as a folder gallery).
    var share = /\/:([a-z]):\//i.exec(path);
    // ".../Forms/AllItems.aspx?id=%2Fsites%2FX%2FDocs%2FRenders&viewid=..." is what copying the address bar
    // gives you while looking at a folder: the folder it shows is in the query.
    var listed = "";
    if (/\/Forms\/AllItems\.aspx$/i.test(path) && params.id) {
      var id = params.id; try { id = decodeURIComponent(id); } catch (e) {}
      if (id) listed = id.charAt(0) === "/" ? (origin + id) : id;
    }
    var last = path.substring(path.lastIndexOf("/") + 1), ext = last.match(/\.([A-Za-z0-9]{1,5})$/);
    var kind;
    if (share) kind = share[1].toLowerCase() === "f" ? "folder" : "file";
    else if (listed) kind = "list";
    else if (/\.aspx$/i.test(path)) kind = (params.sourcedoc || params.file) ? "office" : "page";
    // A FILE needs a real-looking extension: a dot + 1–5 chars with at least one LETTER. A folder named
    // "Renders v2.0" has a dotted last segment, but ".0" is digits-only, so it is still a folder.
    else kind = (ext && /[A-Za-z]/.test(ext[1])) ? "file" : "folder";
    return { raw: raw, path: path, params: params, origin: origin, host: (origin.split("//")[1] || "").toLowerCase(), shareLetter: share ? share[1].toLowerCase() : "", listed: listed, kind: kind };
  }
  function looksLikeFolder(url) { var k = parseSpUrl(url).kind; return k === "folder" || k === "list"; }
  function folderFromAllItems(url) { return parseSpUrl(url).listed; }
  // Canonical KEY for a folder URL (never used to fetch — only as a map/cache key). PMs paste the same
  // folder in slightly different spellings (trailing slash, ?web=1/&csf query junk, %20-vs-space, host
  // casing); without canonicalizing the key, two spellings look like two folders → each is counted as
  // "referenced once" (DEDICATED not SHARED → orphan detection skipped + every file mis-attributed) and
  // the folder is listed from Graph twice. A folder listing canonicalizes to the folder it shows.
  function canonFolderUrl(url) {
    var p = parseSpUrl(url);
    var s = p.listed ? p.listed.split(/[?#]/)[0].replace(/\/+$/, "") : p.path;
    // Fold "+" to a space and decode TWICE: SharePoint hands out the same folder as "Spec%20Sheets",
    // "Spec+Sheets" and (when a link is re-encoded on its way through Office) "Spec%2520Sheets". Each
    // spelling used to be its own key, which is what makes a shared folder look "referenced once".
    s = s.replace(/\+/g, " ");
    try { s = decodeURIComponent(s); } catch (e) {}
    try { if (/%[0-9a-f]{2}/i.test(s)) s = decodeURIComponent(s); } catch (e) {}
    return s.toLowerCase();
  }
  // Canonical identity for a FILE url. Dropping the query (what canonFolderUrl does, correctly, for
  // folders) is WRONG for a file: a SharePoint Office web link puts the file's identity *in* the query —
  // …/_layouts/15/Doc.aspx?sourcedoc={GUID}&file=Name.pptx — so every Office doc in a site collapses to the
  // same …/doc.aspx key. That silently merged four distinct .pptx/.xlsx files into one entry, and because
  // the same key drives the dismiss/removed tombstones, rejecting one of them would have rejected them all.
  // So: keep the identifying query part when there is one, otherwise fall back to the path.
  function canonFileUrl(url) {
    var p = parseSpUrl(url), base = canonFolderUrl(url), got = {};
    ["sourcedoc", "id", "file"].forEach(function (k) {
      if (p.params[k] == null) return;
      var v = p.params[k]; try { v = decodeURIComponent(v); } catch (e) {}
      // sourcedoc is a GUID in braces whose case/brace style varies between links to the same file.
      got[k] = v.replace(/[{}]/g, "").toLowerCase();
    });
    if (got.sourcedoc) return base + "?sourcedoc=" + got.sourcedoc;
    if (got.id) return base + "?id=" + got.id;
    if (got.file) return base + "?file=" + got.file;
    return base;
  }
  // The public summary: { kind, canon, host, path, parent }. `canon` is the identity key the app stores
  // and compares (folder key for a folder/listing, file identity otherwise); `parent` is the containing
  // folder when the URL alone can tell ("" for an Office link — resolve those through Graph).
  function normalizeUrl(url) {
    var p = parseSpUrl(url);
    var isFolder = p.kind === "folder" || p.kind === "list";
    return { kind: p.kind, canon: isFolder ? canonFolderUrl(url) : canonFileUrl(url), host: p.host, path: p.path, parent: parentFolderOf(url) };
  }

  // Two links point at the same file if their Graph ids match (when both known),
  // else if their canonical URLs match. Defeats "Copy link" URL variance.
  function sameFile(a, b) {
    if (!a || !b) return false;
    if (a.id && b.id) return String(a.id) === String(b.id);
    return canonFileUrl(a.url) === canonFileUrl(b.url);
  }
  // Split a "SP Folder URL" cell into its individual link URLs. The legacy fat-cell format packs several
  // URLs into one cell separated by whitespace, so the naive split is /\s+/ — but that ALSO tears a single
  // URL apart at a raw (un-encoded) space, e.g. a hand-pasted ".../My File.pdf", turning one good link into
  // two dead ones that the health checker then flags broken. Splitting only at whitespace that PRECEDES an
  // http(s):// keeps a real multi-URL cell separated while leaving an internal space intact. The one
  // authority for this, shared by the reader (sheetToCatalog) and the writer/guard (catalogsheet), so a
  // link counted one way on read is counted the same way when the shrink guard measures a removal.
  function splitCellUrls(v) {
    var s = String(v == null ? "" : v).trim();
    if (!s) return [];
    return s.split(/\s+(?=https?:\/\/)/i).map(function (u) { return u.trim(); }).filter(Boolean);
  }
  function lhReason(e) {
    var s = String((e && e.message) || e);
    if (/\b404\b/.test(s)) return "file not found — moved or deleted";
    if (/\b403\b/.test(s)) return "no permission to this file";
    if (/\b410\b/.test(s)) return "file no longer exists";
    if (/\b429\b/.test(s)) return "SharePoint rate-limited the check — try again in a moment";
    if (/timed out|timeout|abort/i.test(s)) return "the check timed out — will retry";
    // Everything else is a failure-to-check (network/DNS/TLS/parse), not a diagnosis. Keep the raw code
    // out of the PM sheet — a friendly "couldn't check" reads as "unknown", which is what it is.
    return "couldn’t check right now (network or service issue)";
  }
  // "Broken" must be PROVEN, never assumed. Only a per-file 403 / 404 / 410 (forbidden / not-found / gone)
  // is a genuinely dead link. This is a POLARITY choice, not a completeness one: rather than enumerate
  // every "transient" error string (and silently mark any un-listed one — an undici "fetch failed" from a
  // DNS/TLS/reset blip, or a JSON parse error on a malformed body — as "broken"), we enumerate the small,
  // stable set of codes that PROVE brokenness and treat everything else as "unknown" (preserve, retry
  // tomorrow). On a cry-wolf-sensitive path that already has heartbeat + mass-broken + high-unknown
  // backstops, a rare false "unknown" beats a routine false "broken" that emails the PM team and trains
  // them to ignore the sheet. Fixes both the nightly cron AND the browser "Check all links" button, which
  // share this classifier. A 401 (systemic auth) is not per-file → stays "unknown", as before.
  function lhBroken(e) {
    // Prefer a real status code. Matching the message text is only a fallback for callers that build
    // "Graph " + status — it would also fire on a URL or filename that happens to contain "404".
    var st = e && typeof e.status === "number" ? e.status : null;
    if (st !== null) return st === 403 || st === 404 || st === 410;
    return /\b(403|404|410)\b/.test(String((e && e.message) || e));
  }
  function lhTransient(e) { return !lhBroken(e); }

  // ---- Forgiving (typo-tolerant) search matching ----
  // Optimal String Alignment distance = Levenshtein PLUS adjacent transposition (so "assmebly" is 1 edit
  // from "assembly", not 2). Early-exits with max+1 once a whole row exceeds `max`, keeping the fuzzy
  // pass cheap over the few-hundred-row search index. Pure + string-only, so it lives with the other
  // shared helpers and is unit-tested offline.
  function osaWithin(a, b, max) {
    a = String(a); b = String(b);
    var la = a.length, lb = b.length;
    if (Math.abs(la - lb) > max) return max + 1;
    if (la === 0) return lb <= max ? lb : max + 1;
    if (lb === 0) return la <= max ? la : max + 1;
    var prev2 = [], prev = [], cur = [], i, j;
    for (j = 0; j <= lb; j++) prev[j] = j;
    for (i = 1; i <= la; i++) {
      cur[0] = i;
      var rowMin = cur[0];
      for (j = 1; j <= lb; j++) {
        var cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
        var v = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
        if (i > 1 && j > 1 && a.charAt(i - 1) === b.charAt(j - 2) && a.charAt(i - 2) === b.charAt(j - 1)) v = Math.min(v, prev2[j - 2] + 1);
        cur[j] = v;
        if (v < rowMin) rowMin = v;
      }
      if (rowMin > max) return max + 1;
      prev2 = prev.slice(); prev = cur.slice();
    }
    return prev[lb] <= max ? prev[lb] : max + 1;
  }
  // Edits allowed for a token, scaled by its length: too-short tokens (<=2) get NONE (a 1-edit match on a
  // 2-char token would hit almost anything) — they must land as a clean substring instead.
  function fuzzyThreshold(len) { return len <= 2 ? 0 : len <= 4 ? 1 : len <= 7 ? 2 : 3; }

  // Does the single character `ch` occur in `hay` where a NEW code starts? That means: at the very start,
  // straight after a separator, or across a digit<->letter change — so "400i" and "400 i" both contain the
  // code "i", while "shelving" and "compliance" do not.
  //
  // This exists because the model suffixes the app itself teaches (I, D, R, W, C, L, T) are single letters,
  // and plain substring matching makes every one of them match nearly any text. "HMX 400 I spec sheet"
  // matched HMX 400 *D*'s spec sheet — the "i" came from "shelving" in its breadcrumb — and, being the only
  // hit, was presented as the confident answer. That is the wrong technical document going to a customer,
  // so a one-character token has to land on a real code boundary or not at all.
  function codeTokenIn(hay, ch) {
    hay = String(hay || "").toLowerCase(); ch = String(ch || "").toLowerCase();
    if (!ch) return false;
    var chIsDigit = ch >= "0" && ch <= "9";
    for (var i = hay.indexOf(ch); i !== -1; i = hay.indexOf(ch, i + 1)) {
      if (i === 0) return true;
      var p = hay.charAt(i - 1);
      var pIsDigit = p >= "0" && p <= "9", pIsAlpha = p >= "a" && p <= "z";
      if (!pIsDigit && !pIsAlpha) return true;   // after a space, "—", "›", "-", "."
      if (pIsDigit !== chIsDigit) return true;   // 400|I  or  L|5
    }
    return false;
  }

  // Split a query into match tokens: on whitespace AND at digit↔letter boundaries, so a glued "400i" /
  // "vpp600ipo" separates the number from the model code. Without this, "hpp 400i" leaves "400i" as one
  // blob the exact path can't find (model names carry the space, "HMX 400 I"); the fuzzy fallback then
  // rescues it only by matching "400" and DROPPING the "i", surfacing 400 I / D / R at equal rank — exactly
  // the wrong-suffix mixing the strict path works to prevent. Splitting routes the code fragment through
  // codeTokenIn in BOTH passes. Used by the live search and the demand-list unmet check so they stay aligned.
  // Split on ANY run of non-alphanumerics, not just whitespace. Punctuation is never part of a model code,
  // and a stray one-character token is read as a code below — so "compliance & testing" (a real resource
  // name AND an alias target) left "&" to be looked up in the model name and matched NOTHING, taking the
  // whole compliance/certification alias family down with it and logging each as unmet demand. "hpp-400"
  // failed the same way. Splitting here also removes the lookbehind this function used to need, which was
  // the only one in the browser bundle (a parse-time SyntaxError, i.e. a blank app, on iOS < 16.4).
  function tokenizeQuery(q) {
    var out = [];
    String(q == null ? "" : q).toLowerCase().split(/[^a-z0-9]+/).forEach(function (t) {
      if (!t) return;
      // …then at digit↔letter boundaries, so a glued "400i" separates into "400" + "i".
      t.replace(/([0-9])(?=[a-z])|([a-z])(?=[0-9])/g, "$1$2 ").split(" ").forEach(function (p) { if (p) out.push(p); });
    });
    return out;
  }
  // The catalog search predicate, shared by the live search and the "is this query still unmet?" check that
  // feeds the PM demand list — they MUST agree, or a query the search answers is still logged as a miss.
  //
  // Tokens of 2+ characters match anywhere in name+rest (so "Hoardmaster" finds "Hanging Hoardmaster" and
  // "integrated" finds it through the decoded suffix labels). A ONE-character token is a model code, so it
  // is matched with codeTokenIn against `codeSource` — the MODEL NAME, which is the only place codes live.
  // Everything else in a row is prose that happens to contain letters: the breadcrumb ("Shelving"), and the
  // resource half of a linked row's label, which is why "SNX 200 C" used to pull up
  // "SNX 200 L — Compliance & Testing". Defaults to `name` for rows that are just a name.
  //
  // Conversational filler people type around the words that carry meaning. Dropped only when some other
  // token matched (see searchMatch), so "the thing" on its own still finds nothing.
  var FILLER = { the: 1, a: 1, an: 1, thing: 1, things: 1, stuff: 1, one: 1, some: 1, any: 1, please: 1,
                 that: 1, this: 1, and: 1, or: 1, of: 1, to: 1, my: 1, our: 1, need: 1, want: 1,
                 find: 1, show: 1, me: 1, get: 1 };
  function searchMatch(tokens, name, rest, codeSource) {
    var nm = String(name || "").toLowerCase();
    var hay = (nm + " " + String(rest || "")).toLowerCase();
    var code = String(codeSource == null ? name : codeSource || "").toLowerCase();
    var matched = 0, missedFiller = 0;
    for (var i = 0; i < (tokens || []).length; i++) {
      var t = String(tokens[i] || "").toLowerCase();
      if (!t) continue;
      var ok = (t.length > 1) ? (hay.indexOf(t) !== -1) : codeTokenIn(code, t);
      if (ok) { matched++; continue; }
      // A miss on a MEANINGFUL word still rules the row out — "Hoardmaster riser" must not match the
      // Gondola row. A miss on FILLER is forgiven, but only when something else matched, so a rep typing
      // the way they speak ("the roller shelf thing") isn't told the catalog has nothing.
      if (FILLER[t]) { missedFiller++; continue; }
      return false;
    }
    if (missedFiller && !matched) return false; // filler alone is not a search
    return true;
  }
  // Penalty for matching `query` against `haystack`, tolerating typos. Every whitespace token of the query
  // must match — either as a plain substring of the haystack (penalty 0) OR within its edit-distance
  // threshold of some haystack word. Returns the SUMMED penalty (0 = all clean substring hits, higher =
  // more/bigger typos) so callers can rank best-first, or -1 if any token can't be matched at all.
  // `haystack` should be the combined searchable text (name + trail + any decoded extras); `name` is that
  // row's own name, used to hold one-character tokens to the same code-boundary rule the exact search uses.
  // Without it the fallback re-opens the hole the exact path just closed: the fallback only runs when exact
  // found nothing, which is precisely when a stricter exact match has just rejected the wrong-suffix row.
  function fuzzyPenalty(query, haystack, name) {
    var hay = String(haystack).toLowerCase();
    var nm = name == null ? hay : String(name).toLowerCase();
    var words = hay.split(/[^a-z0-9]+/).filter(Boolean);
    var tokens = tokenizeQuery(query); // same digit↔letter split as the exact path
    if (!tokens.length) return -1;
    var total = 0;
    for (var t = 0; t < tokens.length; t++) {
      var tok = tokens[t];
      if (tok.length === 1) { if (codeTokenIn(nm, tok)) continue; return -1; } // a code matches exactly or not at all
      // A purely numeric token is a model NUMBER — exact substring or nothing. Fuzzy-matching it would let
      // "500" match "400" (one edit), surfacing the wrong product as a confident "closest match" AND
      // suppressing the genuine demand-miss log for a model that really doesn't exist.
      if (/^[0-9]+$/.test(tok)) { if (hay.indexOf(tok) !== -1) continue; return -1; }
      if (hay.indexOf(tok) !== -1) continue; // clean substring hit — no penalty
      var max = fuzzyThreshold(tok.length);
      if (max === 0) return -1; // too short to fuzzy-match safely and it wasn't a substring
      var best = max + 1;
      for (var w = 0; w < words.length && best > 0; w++) {
        var d = osaWithin(tok, words[w], max);
        if (d < best) best = d;
      }
      if (best > max) return -1; // no word near this token → the whole query can't match this row
      total += best;
    }
    return total;
  }

  // Search SYNONYMS: sales / shop-floor vocabulary that doesn't textually match the catalog. Each entry
  // rewrites a whole-word (or whole-phrase) occurrence in the query to the canonical term the catalog
  // actually contains — so "fr" finds Hoardmaster and "tobacco wall" finds goblet rack. The acronyms
  // HMX/GTX/SNX need NO alias: model names start with them, so they already match. Keep multi-word phrases
  // BEFORE single words so a phrase alias is tried before its parts. Add rows as more shop vocabulary
  // surfaces — that's the whole maintenance story ("we can always add more later").
  var SEARCH_ALIASES = [
    { from: "coin shover", to: "coin pushers" },
    { from: "hot rock", to: "heated boulders" },
    // Resource-type colloquialisms — so searching what a resource is COMMONLY called surfaces that resource
    // type across the whole catalog (e.g. "installation sheet" → every model's Assembly Instructions, the way
    // "spec sheet" already surfaces every Spec Sheet). Multi-word phrases first so they're tried before their
    // parts. The target must be the EXACT resource name (that's what the search index stores). Add rows as more
    // shop vocabulary surfaces — the search-miss list on Metrics shows what people actually type.
    { from: "installation sheet", to: "assembly instructions" },
    { from: "installation instructions", to: "assembly instructions" },
    { from: "installation guide", to: "assembly instructions" },
    { from: "install guide", to: "assembly instructions" },
    { from: "install instructions", to: "assembly instructions" },
    { from: "assembly sheet", to: "assembly instructions" },
    { from: "assembly guide", to: "assembly instructions" },
    { from: "sell sheet", to: "flyer" },
    { from: "sales sheet", to: "flyer" },
    { from: "data sheet", to: "spec sheet" },
    { from: "test report", to: "compliance & testing" },
    { from: "product images", to: "renders" },
    { from: "installation", to: "assembly instructions" },
    { from: "install", to: "assembly instructions" },
    { from: "specs", to: "spec sheet" },
    { from: "specification", to: "spec sheet" },
    { from: "render", to: "renders" },
    { from: "renderings", to: "renders" },
    { from: "compliance", to: "compliance & testing" },
    { from: "certification", to: "compliance & testing" },
    { from: "overview", to: "product overview" }
  ];
  // Normalize + validate one PM-entered alias into {from,to} (lowercased, whitespace-collapsed) or null if
  // unusable — empty, identical (a no-op that would just churn), or too long. ONE authority used by the
  // write endpoint (reject bad input before it reaches the sheet) and the client (optimistic add). Keeps a
  // runtime alias shaped exactly like a hardcoded SEARCH_ALIASES row so expandSearchQuery treats them alike.
  function normAlias(from, to) {
    var f = String(from == null ? "" : from).toLowerCase().replace(/\s+/g, " ").trim();
    var t = String(to == null ? "" : to).toLowerCase().replace(/\s+/g, " ").trim();
    if (!f || !t || f === t || f.length > 60 || t.length > 60) return null;
    return { from: f, to: t };
  }
  // Rewrite known aliases in a (lowercased) query. Whole token / whole phrase only — the needle is
  // space-padded, so "fr" hits the standalone token "fr" but never the "fr" inside "front" or "Hoardmaster".
  // Returns the rewritten query, or the original string unchanged when nothing matched. Applied by BOTH the
  // live search and the demand-list "still unmet?" check so an aliased query the search now answers is not
  // logged as a miss. `extra` = runtime PM-added aliases ([{from,to}]) from the Search Aliases sheet, applied
  // AFTER the built-in ones.
  function expandSearchQuery(q, extra) {
    var orig = String(q == null ? "" : q);
    var base = orig.toLowerCase().replace(/\s+/g, " ").trim();
    if (!base) return orig;
    var list = (extra && extra.length) ? SEARCH_ALIASES.concat(extra) : SEARCH_ALIASES;
    var s = " " + base + " ", changed = false;
    for (var i = 0; i < list.length; i++) {
      if (!list[i] || !list[i].from) continue;
      var needle = " " + list[i].from + " ";
      // Skip a rule whose target is ALREADY in the query, or "compliance & testing" expands its own
      // "compliance" half into "compliance & testing & testing".
      if (s.indexOf(" " + list[i].to + " ") !== -1) continue;
      if (s.indexOf(needle) !== -1) { s = s.split(needle).join(" " + list[i].to + " "); changed = true; }
    }
    return changed ? s.replace(/\s+/g, " ").trim() : orig;
  }

  // Union a browser's local favorites with the server-synced set. Server is the cross-device source of
  // truth (newest-first), so it leads; any local-only pins not yet uploaded append after. Deduped,
  // order-stable. Used on load to reconcile localStorage with /api/favorites (and to migrate local-only
  // pins up). Pure so it can be unit-tested; the caller caps the length.
  function mergeFavorites(local, server, removed) {
    // `removed` = models the user un-pinned whose delete-write may not have reached the server yet. Without
    // subtracting them, this union (server ∪ local) can never represent a removal, so a failed/slow delete
    // resurrects the pin on the next load. Dropping tombstoned models here makes removes as durable as adds.
    var drop = {}; (removed || []).forEach(function (m) { if (m) drop[m] = 1; });
    var out = [], seen = {};
    (server || []).concat(local || []).forEach(function (m) { if (!m || seen[m] || drop[m]) return; seen[m] = 1; out.push(m); });
    return out;
  }

  // ---- Smart PLACEMENT guessing + scan planning (the "Files to place" inbox) ----
  // Lives here (browser + node) so the Progress panel and the daily scan share one matcher. Reads a
  // filename the way a person would ("Hoardmaster w riser" == "Hoardmaster with Riser") instead of the
  // dumb exact-prefix match the folder-link filter uses.
  function expandName(fileName) {
    var s = " " + String(fileName).toLowerCase() + " ";
    s = s.replace(/\.[a-z0-9]{1,5}\s*$/, " ");
    s = s.replace(/_+/g, " "); // "HPP_400_I_Spec_Sheet.pdf" — underscores are word separators, not letters
    // "w" is shorthand for "with" only BETWEEN words ("Hoardmaster w riser") or written as "w/". A trailing
    // standalone W is the Wide variant suffix (HMX 400 W) — rewriting that one proposed the file under the
    // BASE model at high confidence, i.e. a wrong suggestion wearing a confident badge.
    s = s.replace(/\bw\/+/g, " with ");
    s = s.replace(/([a-z])\s+w\s+(?=[a-z])/g, "$1 with ");
    s = s.replace(/&/g, " and ");
    s = s.replace(/\b(product\s+)?specs?[\s-]*sheet\b/g, " ");
    s = s.replace(/\bcut\s*sheet\b/g, " ");
    s = s.replace(/\b(assembly|instructions?|manual|guide|warranty|training)\b/g, " ");
    s = s.replace(/\b(low|hi|high)[\s-]*res\b/g, " ");
    s = s.replace(/\b(compressed|final|draft|copy|new|updated?)\b/g, " ");
    s = s.replace(/\bv\d+(\.\d+)?\b/g, " ");
    s = s.replace(/\b\d{1,2}[.\-/]\d{2,4}\b/g, " ");
    s = s.replace(/\b\d{6,8}\b/g, " ");
    s = s.replace(/\s+\d+\s*$/, " ");
    return s.replace(/\s+/g, " ").replace(/^[^a-z0-9]+|[^a-z0-9]+$/g, "").trim();
  }
  // Read the RESOURCE TYPE a filename explicitly names, BEFORE expandName strips those words away. The scan
  // decides a file's resource type from the FOLDER it sits in; this is the cross-check that catches a file
  // whose own name says otherwise (a "Cut Sheet" dropped into the Spec Sheet folder). Only fires on
  // unambiguous phrases — Flyer/Product Overview/Renders/Compliance are too weakly named to trust here, so
  // they return null and fall back to folder context. "cut sheet" is checked before "spec sheet" because
  // both contain "sheet"; a filename with both words is a genuine ambiguity best left to a human anyway.
  function resourceFromName(fileName) {
    var s = " " + String(fileName || "").toLowerCase().replace(/[_\-]+/g, " ") + " ";
    if (/\bcut\s*sheet\b/.test(s)) return "Cut Sheet";
    if (/\b(product\s+)?spec(ification)?s?\s*sheet\b/.test(s)) return "Spec Sheet";
    if (/\b(assembly|installation|instructions?)\b/.test(s)) return "Assembly Instructions";
    if (/\bwarranty\b/.test(s)) return "Warranty";
    if (/\btraining\b/.test(s)) return "Training";
    if (/\bin\s?store\b/.test(s)) return "In-Store Images"; // "in-store"/"in store"/"instore" photos
    return null;
  }
  // models: [{ model, product, … }]. Returns { model, resource, confidence:"high"|"low", candidates, reason }.
  function guessPlacement(fileName, resource, models) {
    var expanded = expandName(fileName);
    var core = normModel(expanded);
    var out = { model: null, resource: resource, confidence: "low", candidates: [], reason: "no-match" };
    if (!core) return out;
    var a = models.filter(function (m) { return matchesModel(expanded, m.model); });
    if (a.length) {
      var maxLen = a.reduce(function (n, m) { return Math.max(n, normModel(m.model).length); }, 0);
      var best = a.filter(function (m) { return normModel(m.model).length === maxLen; });
      if (best.length === 1) return { model: best[0], resource: resource, confidence: "high", candidates: [best[0].model], reason: "name-match" };
      return { model: null, resource: resource, confidence: "low", candidates: best.map(function (m) { return m.model; }), reason: "ambiguous" };
    }
    if (core.length >= 3) {
      var b = models.filter(function (m) { return normModel(m.model).indexOf(core) === 0; });
      if (b.length === 1) return { model: b[0], resource: resource, confidence: "high", candidates: [b[0].model], reason: "unique-prefix" };
      if (b.length > 1) return { model: null, resource: resource, confidence: "low", candidates: b.map(function (m) { return m.model; }), reason: "generic" };
    }
    var fz = models.map(function (m) {
      var mc = normModel(m.model), cap = Math.min(3, Math.floor(mc.length / 3));
      return { m: m, d: osaWithin(core, mc, cap), cap: cap };
    }).filter(function (x) { return x.d <= x.cap; }).sort(function (x, y) { return x.d - y.d; });
    if (fz.length) return { model: null, resource: resource, confidence: "low", candidates: fz.slice(0, 3).map(function (x) { return x.m.model; }), reason: "fuzzy" };
    return out;
  }
  function parentFolderOf(fileUrl) {
    var p = parseSpUrl(fileUrl);
    // A folder listing's parent is the folder above the one it shows.
    if (p.listed) { var j = p.listed.lastIndexOf("/"); return j > 10 ? p.listed.slice(0, j) : ""; }
    // Any other .aspx (an Office web link, a DispForm) has no folder derivable from the URL alone; the
    // caller resolves those through Graph instead.
    if (/\.aspx$/i.test(p.path) || /\/Forms\/DispForm/i.test(p.path)) return "";
    var i = p.path.lastIndexOf("/");
    return i > 10 ? p.path.slice(0, i) : "";
  }
  function wiredUrlSet(links) {
    var set = {};
    Object.keys(links || {}).forEach(function (key) {
      var v = links[key];
      (Array.isArray(v) ? v : [v]).forEach(function (l) { var u = l && l.url ? l.url : l; if (u) set[canonFileUrl(u)] = 1; });
    });
    return set;
  }
  // Folders worth scanning = parent folders of FILE-wired resources (a lone folder link auto-updates
  // already and is left to folder-activity). Tags each with the resource its files serve (majority wins).
  function deriveWatchedFolders(links) {
    var byFolder = {};
    Object.keys(links || {}).forEach(function (key) {
      var resource = key.split("::")[1];
      if (!resource) return;
      var v = links[key];
      var arr = (Array.isArray(v) ? v : [v]).map(function (l) { return l && l.url ? l.url : l; }).filter(Boolean);
      if (arr.length === 1 && looksLikeFolder(arr[0])) return;
      arr.forEach(function (url) {
        if (looksLikeFolder(url)) return;
        var folder = parentFolderOf(url);
        if (!folder) return;
        var canon = canonFolderUrl(folder);
        var e = byFolder[canon] || (byFolder[canon] = { folderUrl: folder, resources: {} });
        e.resources[resource] = (e.resources[resource] || 0) + 1;
      });
    });
    return Object.keys(byFolder).map(function (c) {
      var e = byFolder[c];
      var resource = Object.keys(e.resources).sort(function (a, b) { return e.resources[b] - e.resources[a] || (a < b ? -1 : 1); })[0];
      return { folderUrl: e.folderUrl, resource: resource };
    });
  }
  // "model::resource" keys that are wired to a SINGLE FOLDER. Those already auto-update — whatever is in
  // the folder shows in the app — so placing an individual file link into one is worse than doing nothing:
  // it adds a redundant second link and turns a self-maintaining folder link into a mixed folder+file list.
  function folderWiredSet(links) {
    var set = {};
    Object.keys(links || {}).forEach(function (key) {
      var v = links[key];
      var arr = (Array.isArray(v) ? v : [v]).map(function (l) { return l && l.url ? l.url : l; }).filter(Boolean);
      if (!arr.length) return;
      // Folder-wired = the resource points at ONE folder and nothing else. A doubled folder link (the same
      // folder listed on two rows) is still folder-wired — but the raw array has length 2, so the naive
      // `length === 1` check said "not folder-wired", quietly turning off the auto-update semantics and
      // inviting the placement scan to drop individual files into a folder that already shows everything.
      // Collapse by canonical folder identity first: still folder-wired iff every entry is the SAME folder.
      if (!arr.every(looksLikeFolder)) return;
      var distinct = {};
      arr.forEach(function (u) { distinct[canonFolderUrl(u)] = 1; });
      if (Object.keys(distinct).length === 1) set[key] = 1;
    });
    return set;
  }
  // files: [{ name, url }]. wiredSet from wiredUrlSet(); folderWired from folderWiredSet() (optional).
  // -> { autoPlace:[{model,resource,url,file,folderUrl}], pending:[{file,url,folderUrl,resource,candidates,reason}] }.
  function planFolderPlacements(folderUrl, resource, files, models, wiredSet, folderWired) {
    var autoPlace = [], pending = [];
    (files || []).forEach(function (f) {
      var url = f && (f.url || f.webUrl);
      var name = f && f.name;
      if (!url || !name) return;
      if (wiredSet && wiredSet[canonFileUrl(url)]) return;
      var g = guessPlacement(name, resource, models);
      // Resource-type cross-check. The scan decides a file's resource type from its FOLDER; this catches a
      // file whose own name says otherwise (a "Cut Sheet" dropped into the Spec Sheet folder). It applies
      // whatever the model-match confidence: a confident match is WITHHELD rather than auto-misfiled, and an
      // ambiguous one is offered under its REAL type too — otherwise a human placing it just re-creates the
      // misfile by hand. Only a confident match with NO conflict is placed automatically.
      var named = resourceFromName(name);
      var conflict = !!(named && named !== resource);
      if (g.confidence === "high" && g.model && !conflict) {
        // Already served by a folder link -> nothing to place, and nothing worth asking a human about.
        if (folderWired && folderWired[g.model.model + "::" + resource]) return;
        autoPlace.push({ model: g.model.model, resource: resource, url: url, file: name, folderUrl: folderUrl });
        return;
      }
      pending.push({
        file: name, url: url, folderUrl: folderUrl,
        resource: conflict ? named : resource,
        candidates: g.model ? [g.model.model] : (g.candidates || []),
        reason: conflict ? "resource-conflict" : g.reason
      });
    });
    return { autoPlace: autoPlace, pending: pending };
  }

  // Plan the writes for placing MANY files at once.
  //
  // The wire endpoint REPLACES a resource's whole link set. So placing N files into the same
  // Model·Resource with N sequential writes makes every write race the ones before it, and each write
  // rebuilds its list from a client-side cache the periodic catalog refresh can reset mid-run — the later
  // writes then silently DROP the earlier files AND whatever was already wired. That is not theoretical:
  // it shrank a model's Renders from 6 links to 2 in production while reporting success.
  //
  // So: merge everything up front and emit ONE write per (family, product, resource), carrying every
  // affected model's complete url list. items: [{fam,prod,resource,model,url,file}];
  // existing(model,resource) -> current urls. Returns [{fam,prod,resource,assignments,files}].
  // Should the browser IGNORE a catalog snapshot? Yes when it was generated before the newest snapshot this
  // tab already applied (an older response landing late — F216), or before this tab's last successful
  // save (a pre-write copy served by another warm instance's cache — F010/F019). Applying either reverted
  // the optimistic LINKS, and the next add-only save then silently dropped the link just saved. All three
  // are server ISO timestamps (same format, so they compare as strings). No stamp = can't tell = apply.
  // Do two link lists hold the same files (canonical, order- and duplicate-blind)?
  function sameUrlSet(a, b) {
    var A = Object.create(null), B = Object.create(null), na = 0, nb = 0;
    (a || []).forEach(function (u) { var k = canonFileUrl(u); if (k && !A[k]) { A[k] = 1; na++; } });
    (b || []).forEach(function (u) { var k = canonFileUrl(u); if (k && !B[k]) { B[k] = 1; nb++; } });
    return na === nb && Object.keys(A).every(function (k) { return B[k]; });
  }

  function catalogSnapshotStale(generatedAt, appliedAt, writeFloor) {
    if (!generatedAt) return false;
    return (!!appliedAt && generatedAt < appliedAt) || (!!writeFloor && generatedAt < writeFloor);
  }

  function planPlacementWrites(items, existing) {
    var groups = {}, order = [], urlsByGroup = {};
    (items || []).forEach(function (x) {
      if (!x || !x.fam || !x.prod || !x.resource || !x.model || !x.url) return;
      var key = x.fam + "§" + x.prod + "§" + x.resource;
      if (!groups[key]) { groups[key] = { fam: x.fam, prod: x.prod, resource: x.resource, assignments: {}, files: [] }; urlsByGroup[key] = {}; order.push(key); }
      var byModel = urlsByGroup[key];
      if (!byModel[x.model]) {
        var pre = ((existing && existing(x.model, x.resource)) || []).slice();
        byModel[x.model] = { base: pre.slice(), urls: pre, added: false };
      }
      var cur = byModel[x.model];
      // Canonical compare so the same file via a differently-spelled link isn't added twice.
      if (cur.urls.map(canonFileUrl).indexOf(canonFileUrl(x.url)) === -1) { cur.urls.push(x.url); cur.added = true; groups[key].files.push(x.file || x.url); }
    });
    // Each group also carries `base` — the pre-image each model's list was built from — so the write can
    // be compare-and-swapped: the shrink guard only counts, and a stale copy [A,B] + D over a sheet that
    // now holds [A,B,C] is 3 → 3, so C was silently dropped (F080). A model (or a whole group) that gains
    // nothing is left out: rewriting it from a possibly-stale copy could only undo someone else's change.
    return order.map(function (k) {
      var g = groups[k], byModel = urlsByGroup[k];
      g.base = {};
      Object.keys(byModel).forEach(function (m) {
        if (!byModel[m].added) return;
        g.assignments[m] = byModel[m].urls.join(" ");
        g.base[m] = byModel[m].base;
      });
      return g;
    }).filter(function (g) { return Object.keys(g.assignments).length > 0; });
  }

  // Compile per-model add/remove intents into one write group per (fam,prod,resource),
  // each carrying every affected model's COMPLETE desired list + declared removal counts
  // so the /api/wire-links shrink guard passes. Removal & dedupe are canonical.
  // intents: [{fam,prod,model,addUrl?,removeUrl?}] (both = "replace on this model").
  // existing(model,resource) -> current urls. resource = the single resource type string.
  function planDocumentOps(intents, resource, existing) {
    var groups = {}, order = [];
    function grp(fam, prod) {
      var key = fam + "§" + prod;
      if (!groups[key]) { groups[key] = { fam: fam, prod: prod, resource: resource, desired: {}, original: {} }; order.push(key); }
      return groups[key];
    }
    function cur(g, model) {
      if (!g.desired[model]) { g.original[model] = ((existing && existing(model, resource)) || []).slice(); g.desired[model] = g.original[model].slice(); }
      return g.desired[model];
    }
    (intents || []).forEach(function (it) {
      if (!it || !it.fam || !it.prod || !it.model) return;
      var g = grp(it.fam, it.prod), list = cur(g, it.model);
      if (it.removeUrl) {
        var kept = list.filter(function (u) { return canonFileUrl(u) !== canonFileUrl(it.removeUrl); });
        g.desired[it.model] = kept; list = kept;
      }
      if (it.addUrl && list.map(canonFileUrl).indexOf(canonFileUrl(it.addUrl)) === -1) list.push(it.addUrl);
    });
    // Per model: the NET drop in distinct files. A replace removes the old file and adds the new one — net
    // 0 — and declaring the gross 1 handed the count-based shrink guard an allowance that hid a link lost
    // from a stale copy on exactly those models (F038). `base` is the pre-image each list was built from,
    // for the compare-and-swap. A model already in the wanted state is left out (and an empty group
    // dropped), so a retry after a partial failure — re-planned from a fresh read — sends only what
    // didn't land.
    function distinct(list) { var seen = {}, n = 0; list.forEach(function (u) { var k = canonFileUrl(u); if (k && !seen[k]) { seen[k] = 1; n++; } }); return n; }
    function sameSet(a, b) { var A = {}, B = {}; a.forEach(function (u) { A[canonFileUrl(u)] = 1; }); b.forEach(function (u) { B[canonFileUrl(u)] = 1; }); var ka = Object.keys(A), kb = Object.keys(B); return ka.length === kb.length && ka.every(function (k) { return B[k]; }); }
    return order.map(function (key) {
      var g = groups[key], assignments = {}, byModel = {}, finalByKey = {}, base = {}, total = 0;
      Object.keys(g.desired).forEach(function (m) {
        if (sameSet(g.desired[m], g.original[m])) return;
        assignments[m] = g.desired[m].join(" ");
        finalByKey[m + "::" + resource] = g.desired[m].slice();
        base[m] = g.original[m].slice();
        var drop = Math.max(0, distinct(g.original[m]) - distinct(g.desired[m]));
        if (drop) { byModel[m] = drop; total += drop; }
      });
      return { fam: g.fam, prod: g.prod, resource: resource, assignments: assignments, expectRemoved: total, expectRemovedByModel: byModel, finalByKey: finalByKey, base: base };
    }).filter(function (g) { return Object.keys(g.assignments).length > 0; });
  }

  // ---- Link-health CLASSIFIER (shared by the browser "Check all links" and the daily cron) ----
  // The ONLY difference between the two callers is HOW they reach SharePoint (delegated token + client
  // fetch vs. app-only token + server fetch) and how they count folder sharing. So the fetch layer is
  // injected via `deps`; the empty/broken/ok/unknown decision + wording lives here, once.
  //   deps.isShared(url)     -> bool: is this single folder referenced by >1 model (filter by name) ?
  //   deps.listFolder(url)   -> Promise<[{name,id}]>: files in a folder
  //   deps.resolveOne(url)   -> Promise (resolves if the file opens, rejects otherwise)
  //   deps.cache             -> per-run memo object (a folder shared across models is fetched once)
  // `it` = { model, resource, arr:[{url}] }. Returns Promise<{ status, detail, url }>.
  function lhCheckResource(it, deps) {
    var arr = it.arr, cache = deps.cache || (deps.cache = {});
    if (arr.length === 1 && looksLikeFolder(arr[0].url)) {
      var url = arr[0].url, shared = deps.isShared(url);
      var ck = deps.canonUrl ? deps.canonUrl(url) : url; // dedupe folder listings across URL spellings
      var p = cache[ck] || (cache[ck] = deps.listFolder(url));
      return p.then(function (files) {
        var matched = shared ? files.filter(function (fl) { return matchesModel(fl.name, it.model); }) : files;
        if (!matched.length) return { status: "empty", detail: shared ? ("shared folder has no files matching “" + it.model + "”") : "this folder is empty", url: url };
        return { status: "ok", detail: matched.length + " file" + (matched.length === 1 ? "" : "s") + " in folder", url: url };
      }, function (e) { return { status: lhTransient(e) ? "unknown" : "broken", detail: lhReason(e), url: url }; });
    }
    var jobs = arr.map(function (l) {
      var key = "f:" + l.url;
      return cache[key] || (cache[key] = deps.resolveOne(l.url).then(function () { return { ok: true }; }, function (e) { return { ok: false, transient: lhTransient(e) }; }));
    });
    return Promise.all(jobs).then(function (res) {
      var u = arr[0].url;
      var okN = res.filter(function (x) { return x.ok; }).length;
      var hardBad = res.filter(function (x) { return !x.ok && !x.transient; }).length;
      // A genuine 404/403 on any file => broken. If we resolved nothing but every failure was transient
      // (all throttled/timed out), report "unknown" rather than falsely "broken".
      // Report the FAILING file's url, not arr[0]'s. On a multi-file resource the Link Health sheet's
      // "Link" column was pointing the PM at a file that opens perfectly while a different one was broken.
      if (hardBad > 0) {
        var badIdx = -1;
        for (var bi = 0; bi < res.length; bi++) if (!res[bi].ok && !res[bi].transient) { badIdx = bi; break; }
        var badUrl = (badIdx >= 0 && arr[badIdx]) ? arr[badIdx].url : u;
        return { status: "broken", detail: (okN === 0 && res.length === 1) ? "the file can’t be opened — moved, deleted, or no access" : (hardBad + " of " + res.length + " file" + (res.length === 1 ? "" : "s") + " can’t be opened"), url: badUrl };
      }
      if (okN === 0) return { status: "unknown", detail: "couldn’t check right now (rate-limited or timed out)", url: u };
      return { status: "ok", detail: okN + " file" + (okN === 1 ? "" : "s"), url: u };
    });
  }

  // Concurrency pool so we don't fire hundreds of Graph calls at once (429s). opts.deadline (ms epoch):
  // once passed, every not-yet-started item is marked "unknown" (preserved on sync, not resolved) so a
  // slow daily run still writes its heartbeat instead of being killed. opts.onProgress(done,total) is an
  // optional progress callback (the browser's live counter).
  function lhRunPool(items, worker, concurrency, opts) {
    opts = opts || {};
    var deadline = opts.deadline || 0, onProgress = opts.onProgress || null;
    return new Promise(function (resolve) {
      if (!items.length) { resolve([]); return; }
      var i = 0, done = 0, active = 0, results = new Array(items.length);
      function pump() {
        if (deadline && Date.now() > deadline) {
          while (i < items.length) { results[i] = { status: "unknown", detail: "not checked — daily run reached its time limit" }; i++; done++; if (onProgress) onProgress(done, items.length); }
          if (done === items.length) resolve(results);
          return; // let any in-flight workers settle; the last one to finish resolves
        }
        while (active < concurrency && i < items.length) {
          (function (idx) {
            active++;
            // `new Promise(res => res(worker(...)))`, not `Promise.resolve(worker(...))`: if the worker
            // throws SYNCHRONOUSLY the latter never gets to build a promise, so the throw escapes pump()
            // — the pool then never settles (the whole check hangs) and it surfaces as an unhandled
            // rejection. Wrapping turns a sync throw into a normal rejection this handler already covers.
            new Promise(function (res) { res(worker(items[idx])); }).then(function (r) { results[idx] = r; }, function (e) { results[idx] = { status: lhTransient(e) ? "unknown" : "broken", detail: lhReason(e) }; })
              .then(function () { active--; done++; if (onProgress) onProgress(done, items.length); if (done === items.length) resolve(results); else pump(); });
          })(i++);
        }
      }
      pump();
    });
  }

  // Data-integrity check: the app keys every link by MODEL NAME alone (LINKS["<model>::<resource>"]), so
  // two DIFFERENT products that each define a model with the same name silently SHARE one link set — open
  // either and you get the other's files (the "Goblet Rack" incident). Scan the catalog tree for a model name
  // that appears under 2+ distinct products and report each collision so a PM can rename one in the sheet.
  // Product identity is the product NAME, so the same product legitimately listed under multiple families
  // is NOT a false positive. Grouping key is the exact (trimmed) model string — the same string sheetToCatalog
  // uses to build the LINKS key — so a reported collision is exactly one that shares links. Returns
  // [{ model, products: [name, ...] }] sorted by model, empty when clean.
  function dupeModelNames(catalog) {
    var byModel = Object.create(null); // model string -> { display, products: {name: true} }
    (catalog || []).forEach(function (c) {
      (c.families || []).forEach(function (f) {
        (f.products || []).forEach(function (p) {
          var pname = String((p && p.name) || "").trim();
          (p.models || []).forEach(function (m) {
            var key = String(m == null ? "" : m).trim();
            if (!key) return;
            var e = byModel[key] || (byModel[key] = { display: key, products: Object.create(null) });
            if (pname) e.products[pname] = true;
          });
        });
      });
    });
    var out = [];
    Object.keys(byModel).forEach(function (k) {
      var prods = Object.keys(byModel[k].products);
      if (prods.length >= 2) out.push({ model: byModel[k].display, products: prods.sort(function (a, b) { return a.localeCompare(b); }) });
    });
    out.sort(function (a, b) { return a.model.toLowerCase() < b.model.toLowerCase() ? -1 : a.model.toLowerCase() > b.model.toLowerCase() ? 1 : 0; });
    return out;
  }

  // ---- Find filter/sort core (used by the model column AND search results) ----
  // Pure logic only: the browser supplies `accessors` built from its own closure helpers
  // (usableCount, decodeModelName, ADDED, …) so this file stays side-effect-free and testable.
  //
  // state = { status: "any"|"has"|"wip", attrs: [codes], recent: bool, sort: ""|"name"|"complete-desc"|"complete-asc"|"newest" }
  // accessors = { hasFiles(m)->bool, attrs(m)->[codes], addedAt(m)->ms|null, completeness(m)->number, name(m)->string, recentCutoff: ms }
  function filterSortModels(items, state, accessors) {
    state = state || {};
    var status = state.status || "any";
    var wantAttrs = state.attrs || [];
    var recent = !!state.recent;
    var sort = state.sort || "";
    var cutoff = accessors.recentCutoff;
    // Filter — groups combine with AND; attributes are OR within their group.
    var kept = items.filter(function (m) {
      if (status === "has" && !accessors.hasFiles(m)) return false;
      if (status === "wip" && accessors.hasFiles(m)) return false;
      if (wantAttrs.length) {
        var a = accessors.attrs(m) || [];
        var hit = wantAttrs.some(function (code) { return a.indexOf(code) !== -1; });
        if (!hit) return false;
      }
      if (recent) {
        var t = accessors.addedAt(m);
        if (t == null || t < cutoff) return false;
      }
      return true;
    });
    // Sort — decorate with the original index so every comparator is a STABLE sort (ties keep
    // input order; on search results the input order is the relevance rank we must not scramble).
    var dec = kept.map(function (m, i) { return { m: m, i: i }; });
    var cmp = null;
    if (sort === "name") {
      cmp = function (a, b) {
        var an = String(accessors.name(a.m)).toLowerCase(), bn = String(accessors.name(b.m)).toLowerCase();
        return an < bn ? -1 : an > bn ? 1 : a.i - b.i;
      };
    } else if (sort === "complete-desc") {
      cmp = function (a, b) { return (accessors.completeness(b.m) - accessors.completeness(a.m)) || (a.i - b.i); };
    } else if (sort === "complete-asc") {
      cmp = function (a, b) { return (accessors.completeness(a.m) - accessors.completeness(b.m)) || (a.i - b.i); };
    } else if (sort === "newest") {
      cmp = function (a, b) {
        var at = accessors.addedAt(a.m), bt = accessors.addedAt(b.m);
        if (at == null && bt == null) return a.i - b.i; // both undated → keep input order
        if (at == null) return 1;  // undated sinks below any dated model
        if (bt == null) return -1;
        return (bt - at) || (a.i - b.i);
      };
    }
    if (cmp) dec.sort(cmp); // unknown/"" sort → leave input order untouched
    return dec.map(function (d) { return d.m; });
  }
  // Which attribute chips to offer for a list: only codes actually present on some model, returned
  // in the caller's canonical code order (so the chips read I · PO · W, not the order first seen).
  function availableAttrs(items, attrsOf, codeOrder) {
    var present = {};
    items.forEach(function (m) { (attrsOf(m) || []).forEach(function (c) { present[c] = true; }); });
    return (codeOrder || []).filter(function (c) { return present[c]; });
  }

  // ---- Coverage computation (Resource Impact snapshot) ----
  // Pure, server-usable: mirrors the client Progress matrix. A slot (model × resource-type) counts toward
  // the denominator unless it's N/A; it counts as covered when it has a link that this run did NOT find
  // broken/empty. `naKeys` and `problemKeys` are "model::resource" strings (array or Set).
  function computeCoverage(catalog, links, naKeys, problemKeys, resourceTypes) {
    var types = resourceTypes || RESOURCES;
    var na = (naKeys instanceof Set) ? naKeys : new Set(naKeys || []);
    var problems = (problemKeys instanceof Set) ? problemKeys : new Set(problemKeys || []);
    links = links || {};
    var linked = 0, total = 0, perType = {};
    types.forEach(function (t) { perType[t] = { linked: 0, total: 0 }; });
    (catalog || []).forEach(function (c) {
      (c.families || []).forEach(function (f) {
        (f.products || []).forEach(function (p) {
          (p.models || []).forEach(function (m) {
            types.forEach(function (t) {
              var key = m + "::" + t;
              if (na.has(key)) return; // not applicable — out of the denominator
              total++; perType[t].total++;
              var arr = links[key];
              if (arr && arr.length && !problems.has(key)) { linked++; perType[t].linked++; }
            });
          });
        });
      });
    });
    return { linked: linked, total: total, pct: total ? Math.floor(linked / total * 100) : 0, perType: perType };
  }

  // ---- Gibberish search-miss filter ----
  // Keeps keyboard-mash queries ("zzzqqxnotathing") off the PM "wanted by staff" worklist and the Metrics
  // list, without dropping real demand. High-precision (few false positives): a query with any digit is
  // assumed real; otherwise it's gibberish only if some token (≥5 letters) has a 3+ same-letter run, a 5+
  // consonant run, or no vowel at all. Real product/model/sales terms don't hit any of those.
  function looksLikeGibberish(q) {
    var str = String(q == null ? "" : q).toLowerCase();
    if (/\d/.test(str)) return false; // a number → likely a real model/size query, keep it
    var toks = str.split(/[^a-z]+/).filter(Boolean);
    return toks.some(function (t) {
      if (/(.)\1\1/.test(t)) return true;                    // 3+ identical letters in a row (zzz, aaaa) — any length
      if (t.length < 5) return false;                        // other rules only judge longer tokens
      if (/[bcdfghjklmnpqrstvwxz]{5,}/.test(t)) return true; // 5+ consonants in a row (y treated as a vowel)
      if (!/[aeiouy]/.test(t)) return true;                  // no vowel at all — y counts, as the rule above says
                                                             // (without it "rhythm"/"glyph" were called gibberish)
      return false;
    });
  }

  // "Share to customer" (whole model): choose which of a model's linked files to hand the phone's share
  // sheet. DOCUMENTS go first (every resource type except the render/image ones), then render files, capped
  // at opts.cap so an iPhone share sheet isn't overloaded (it only accepts a handful reliably). Auto-updating
  // render FOLDER links are set aside — they expand to many images that can't ride one share; the caller
  // tells the rep to use "Download all" for those. Pure + order-stable so it can be unit-tested offline.
  //   entries: [{ url, name, resource }]   opts: { cap=9, renderResources=["Renders"] }
  //   returns { files:[{url,name,resource}], setAside }  (setAside = folder links + files past the cap)
  function selectModelShareFiles(entries, opts) {
    opts = opts || {};
    var cap = opts.cap > 0 ? opts.cap : 9;
    var renderSet = {}; (opts.renderResources || ["Renders"]).forEach(function (r) { renderSet[r] = 1; });
    var direct = [], folders = 0;
    (entries || []).forEach(function (e) {
      if (!e || !e.url) return;
      if (looksLikeFolder(e.url)) { folders++; return; } // an auto-updating folder can't ride one share sheet
      direct.push({ url: e.url, name: e.name || null, resource: e.resource || null });
    });
    var docs = direct.filter(function (e) { return !renderSet[e.resource]; });
    var renders = direct.filter(function (e) { return renderSet[e.resource]; });
    var ordered = docs.concat(renders);                 // documents first, then renders
    var files = ordered.slice(0, cap);
    return { files: files, setAside: folders + (ordered.length - files.length) };
  }

  // ---- Select mode (gallery tiles keyed by index, model-panel rows keyed by resource name) ----
  // One state machine for both surfaces: OFF = the actions take everything (as before); ON = they take
  // the picks, and an EMPTY pick is refused ("Pick files first") rather than silently meaning "all".
  function makeSelection() {
    var on = false, picked = Object.create(null);
    function clear() { picked = Object.create(null); }
    return {
      isOn: function () { return on; },
      start: function () { on = true; },
      stop: function () { on = false; clear(); },
      toggle: function (k) { if (picked[k]) delete picked[k]; else picked[k] = 1; return !!picked[k]; },
      has: function (k) { return !!picked[k]; },
      all: function (keys) { (keys || []).forEach(function (k) { picked[k] = 1; }); },
      clear: clear,
      count: function () { return Object.keys(picked).length; },
      pick: function (keys) { return (keys || []).filter(function (k) { return !!picked[k]; }); }
    };
  }
  // The label an action button shows. `cap` = the most the action can take (Share's 9): over it the
  // label says "9 of 11" so the limit is visible instead of a silent trim.
  function selectionLabel(base, selectedWord, sel, cap) {
    if (!sel || !sel.isOn()) return base;
    var n = sel.count();
    if (!n) return "Pick files first";
    if (cap > 0 && n > cap) return selectedWord + " (" + cap + " of " + n + ")";
    return selectedWord + " (" + n + ")";
  }

  // Graph returns three thumbnail sizes per item (small 96px, medium 176px, large 800px). A strip tile
  // or a 40px chip must not pull the 800px one — a 30-file render folder was 1-3 MB of thumbnails for
  // 36px tiles. Pick the SMALLEST size that still covers `minPx` (pass the box's CSS size × device
  // pixel ratio); fall back to the largest available when none does; null when there are none.
  function pickThumb(ts, minPx) {
    if (!ts) return null;
    var order = ["small", "medium", "large"], last = null;
    for (var i = 0; i < order.length; i++) {
      var t = ts[order[i]];
      if (!t || !t.url) continue;
      last = t.url;
      var px = Math.max(t.width || 0, t.height || 0);
      if (px && px >= (minPx || 0)) return t.url;
    }
    return last;
  }

  // Promise memo with a TTL: `get()` shares one in-flight/recent call, `invalidate()` forces the next
  // get() to refetch (call it after any write to the same resource). A rejection is never cached — the
  // next caller retries. `now` is injectable for tests.
  function memoPromise(fn, ttlMs, now) {
    now = now || function () { return Date.now(); };
    var p = null, at = 0;
    return {
      get: function () {
        if (p && now() - at < ttlMs) return p;
        at = now();
        var mine = p = new Promise(function (res) { res(fn()); });
        mine.catch(function () { if (p === mine) p = null; });
        return mine;
      },
      invalidate: function () { p = null; at = 0; }
    };
  }

  return {
    RESOURCES: RESOURCES, selectModelShareFiles: selectModelShareFiles,
    pickThumb: pickThumb, memoPromise: memoPromise, makeSelection: makeSelection, selectionLabel: selectionLabel,
    filterSortModels: filterSortModels, availableAttrs: availableAttrs,
    computeCoverage: computeCoverage, looksLikeGibberish: looksLikeGibberish, normCell: normCell,
    dupeModelNames: dupeModelNames,
    modelRe: modelRe, normModel: normModel, matchesModel: matchesModel,
    looksLikeFolder: looksLikeFolder, canonFolderUrl: canonFolderUrl, canonFileUrl: canonFileUrl, parseSpUrl: parseSpUrl, normalizeUrl: normalizeUrl, sameFile: sameFile, splitCellUrls: splitCellUrls, lhReason: lhReason, lhTransient: lhTransient, lhBroken: lhBroken,
    lhCheckResource: lhCheckResource, lhRunPool: lhRunPool,
    osaWithin: osaWithin, fuzzyThreshold: fuzzyThreshold, fuzzyPenalty: fuzzyPenalty,
    codeTokenIn: codeTokenIn, searchMatch: searchMatch, tokenizeQuery: tokenizeQuery, expandSearchQuery: expandSearchQuery, normAlias: normAlias, SEARCH_ALIASES: SEARCH_ALIASES,
    expandName: expandName, resourceFromName: resourceFromName, guessPlacement: guessPlacement, parentFolderOf: parentFolderOf,
    wiredUrlSet: wiredUrlSet, deriveWatchedFolders: deriveWatchedFolders, planFolderPlacements: planFolderPlacements,
    folderWiredSet: folderWiredSet, planPlacementWrites: planPlacementWrites, planDocumentOps: planDocumentOps, catalogSnapshotStale: catalogSnapshotStale, sameUrlSet: sameUrlSet,
    mergeFavorites: mergeFavorites
  };
});
