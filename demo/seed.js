/* Demo seed — Grumbleton & Sons, Fine Goods for Dragons. Every product, person and file here is fictional.
   Shapes match what the real /api/* endpoints return (see docs/design/specs/2026-09-22-demo-mode-design.md);
   demo/backend.js serves them. Links are SharePoint-shaped URLs on a fictional host so every URL helper in
   pf-shared.js (file vs folder, canonical keys, share ids) works unchanged. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.PF_DEMO_SEED = factory();
})(this, function () {
  "use strict";
  var BASE = "https://grumbleton.example/sites/Lair/Shared%20Documents/";
  function u(p) { return BASE + p.split("/").map(encodeURIComponent).join("/"); }

  var catalog = [
    { name: "Hoard Management", families: [
      { name: "Coin Pushers", products: [{ name: "Hoardmaster", models: ["Hoardmaster 3000", "Hoardmaster 3000 X", "Hoardmaster 3000 XF", "Hoardmaster 5000 D"] }] },
      { name: "Treasure Sorters", products: [{ name: "Gemtumbler", models: ["Gemtumbler 2", "Gemtumbler 2 G"] }, { name: "Crownpolisher", models: ["Crownpolisher 1"] }] },
      { name: "Hoard Doors", products: [{ name: "Vaultvane", models: ["Vaultvane 100", "Vaultvane 200 F"] }] },
      { name: "Goblet Racks", products: [{ name: "Goblet Rack", models: ["Goblet Rack Twelve"] }] }
    ] },
    { name: "Lair Comfort", families: [
      { name: "Heated Boulders", products: [{ name: "Snugrock", models: ["Snugrock Classic", "Snugrock King"] }] },
      { name: "Fireproof Throw Pillows", products: [{ name: "Cinderpuff", models: ["Cinderpuff Pillow", "Cinderpuff Bolster"] }] },
      { name: "Cave Mood Lighting", products: [{ name: "Glowmoss", models: ["Glowmoss Lantern"] }] },
      { name: "Cave Rugs", products: [{ name: "Ashweave", models: ["Ashweave Hearth", "Ashweave Great Hall"] }] },
      { name: "Scratching Posts", products: [{ name: "Clawpost", models: ["Clawpost Classic"] }] }
    ] },
    { name: "Knight Deterrence", families: [
      { name: "Moat Kits", products: [{ name: "Moat-in-a-Box", models: ["Moat-in-a-Box 20 ft", "Moat-in-a-Box 40 ft"] }] },
      { name: "Signage", products: [{ name: "Beware of Me", models: ["Beware of Me Sign"] }] },
      { name: "Decoy Towers", products: [{ name: "Decoy Princess Tower", models: ["Decoy Princess Tower"] }] },
      { name: "Alarms", products: [{ name: "Pebblechime", models: ["Pebblechime 5"] }] },
      { name: "Armour Racks", products: [{ name: "Trophy Rack", models: ["Trophy Rack Six"] }] }
    ] }
  ];

  var files = {}, links = {}, folders = {};
  // Register a file; `body` tells scripts/make-demo-files.mjs what to draw/write.
  function file(path, body) {
    var url = u(path), name = path.split("/").pop();
    files[url] = { name: name, mime: /\.jpe?g$/i.test(name) ? "image/jpeg" : /\.svg$/i.test(name) ? "image/svg+xml" : "application/pdf", body: body };
    return url;
  }
  function folder(path, entries) { var url = u(path); folders[url] = entries.map(function (e) { return file(path + "/" + e[0], e[1]); }); return url; }
  function doc(model, kind, label) { return file(label + "/" + model + " " + kind + ".pdf", kind.toLowerCase().split(" ")[0] + ":" + model); }

  // --- Hoard Management ---
  ["Hoardmaster 3000", "Hoardmaster 3000 X", "Hoardmaster 3000 XF", "Hoardmaster 5000 D"].forEach(function (m) {
    links[m + "::Spec Sheet"] = doc(m, "Spec Sheet", "Spec Sheets");
    links[m + "::Warranty"] = doc(m, "Warranty", "Warranty");
  });
  links["Hoardmaster 3000::Assembly Instructions"] = doc("Hoardmaster 3000", "Assembly Instructions", "Assembly");
  links["Hoardmaster 3000 X::Assembly Instructions"] = doc("Hoardmaster 3000 X", "Assembly Instructions", "Assembly");
  // A model with its own renders folder (the folder link keeps itself up to date) …
  links["Hoardmaster 3000::Renders"] = folder("Renders/Hoardmaster 3000", [
    ["Hoardmaster 3000 front.jpg", "photo:hoardmaster-3000"], ["Hoardmaster 3000 straight on.jpg", "photo:hoardmaster-3000-front"],
    ["Hoardmaster 3000 top.jpg", "photo:hoardmaster-3000-top"], ["Hoardmaster 3000 in the lair.jpg", "photo:hoardmaster-3000-lair"]
  ]);
  links["Hoardmaster 3000 X::Renders"] = folder("Renders/Hoardmaster 3000 X", [
    ["Hoardmaster 3000 X loaded.jpg", "photo:hoardmaster-3000-x"], ["Hoardmaster 3000 X on display.jpg", "photo:store-shelf"]
  ]);
  // … and a two-file flyer (a multi-file resource opens as a gallery).
  links["Hoardmaster 3000::Flyer"] = [
    { label: "Flyer (page 1)", url: file("Flyers/Hoardmaster 3000 Flyer p1.pdf", "flyer:Hoardmaster 3000:1") },
    { label: "Flyer (page 2)", url: file("Flyers/Hoardmaster 3000 Flyer p2.pdf", "flyer:Hoardmaster 3000:2") }
  ];
  links["Hoardmaster 3000::Training"] = doc("Hoardmaster 3000", "Training", "Training");
  links["Gemtumbler 2::Spec Sheet"] = doc("Gemtumbler 2", "Spec Sheet", "Spec Sheets");
  links["Gemtumbler 2 G::Spec Sheet"] = doc("Gemtumbler 2 G", "Spec Sheet", "Spec Sheets");
  links["Gemtumbler 2::Renders"] = [
    { label: null, url: file("Renders/Gemtumbler 2 front.jpg", "photo:gemtumbler") },
    { label: null, url: file("Renders/Gemtumbler 2 on display.jpg", "photo:store-shelf") }
  ];
  links["Goblet Rack Twelve::Spec Sheet"] = doc("Goblet Rack Twelve", "Spec Sheet", "Spec Sheets");

  // --- Lair Comfort --- (one renders folder shared by both Snugrock models; it also holds a stray file)
  var snugFolder = folder("Renders/Snugrock", [
    ["Snugrock Classic studio.jpg", "photo:snugrock"], ["Snugrock King in the cave.jpg", "photo:snugrock-cave"],
    ["Mystery Egg Warmer.jpg", "photo:egg-warmer"]
  ]);
  ["Snugrock Classic", "Snugrock King"].forEach(function (m) {
    links[m + "::Spec Sheet"] = doc(m, "Spec Sheet", "Spec Sheets");
    links[m + "::Renders"] = snugFolder;
  });
  links["Snugrock King::Warranty"] = doc("Snugrock King", "Warranty", "Warranty");
  links["Snugrock Classic::Training"] = doc("Snugrock Classic", "Training", "Training");
  links["Cinderpuff Pillow::Spec Sheet"] = doc("Cinderpuff Pillow", "Spec Sheet", "Spec Sheets");
  links["Cinderpuff Pillow::Compliance & Testing"] = doc("Cinderpuff Pillow", "Compliance Certificate", "Compliance");
  links["Glowmoss Lantern::Product Overview"] = doc("Glowmoss Lantern", "Product Overview", "Overviews");

  // --- Knight Deterrence --- (one link deliberately points at a file that was moved: the broken link)
  var BROKEN_URL = u("Spec Sheets/Old/Moat-in-a-Box 20 ft Spec Sheet (OLD).pdf");
  links["Moat-in-a-Box 20 ft::Spec Sheet"] = BROKEN_URL;
  links["Moat-in-a-Box 40 ft::Spec Sheet"] = doc("Moat-in-a-Box 40 ft", "Spec Sheet", "Spec Sheets");
  links["Moat-in-a-Box 40 ft::Assembly Instructions"] = doc("Moat-in-a-Box 40 ft", "Assembly Instructions", "Assembly");
  links["Beware of Me Sign::Spec Sheet"] = doc("Beware of Me Sign", "Spec Sheet", "Spec Sheets");
  links["Decoy Princess Tower::Renders"] = [{ label: null, url: file("Renders/Decoy Princess Tower front.jpg", "photo:decoy-tower") }];
  links["Decoy Princess Tower::Warranty"] = doc("Decoy Princess Tower", "Warranty", "Warranty");

  // Product illustrations for the new lines.
  links["Vaultvane 100::Renders"] = [{ label: null, url: file("Renders/Vaultvane 100.jpg", "photo:vaultvane") }];
  links["Crownpolisher 1::Renders"] = [{ label: null, url: file("Renders/Crownpolisher 1.jpg", "photo:crownpolisher") }];
  links["Ashweave Hearth::Renders"] = [{ label: null, url: file("Renders/Ashweave rug.jpg", "photo:ashweave") }];
  links["Clawpost Classic::Renders"] = [{ label: null, url: file("Renders/Clawpost Classic.jpg", "photo:clawpost") }];
  links["Pebblechime 5::Renders"] = [{ label: null, url: file("Renders/Pebblechime 5.jpg", "photo:pebblechime") }];
  links["Trophy Rack Six::Renders"] = [{ label: null, url: file("Renders/Trophy Rack Six.jpg", "photo:trophy-rack") }];

  // Product illustrations for the rest of the range.
  links["Hoardmaster 5000 D::Renders"] = [{ label: null, url: file("Renders/Hoardmaster 5000 D front.jpg", "photo:hoardmaster-5000-d") }];
  links["Cinderpuff Pillow::Renders"] = [{ label: null, url: file("Renders/Cinderpuff Pillow and Bolster.jpg", "photo:cinderpuff") }];
  links["Glowmoss Lantern::Renders"] = [
    { label: null, url: file("Renders/Glowmoss Lantern.jpg", "photo:glowmoss") },
    { label: null, url: file("Renders/Glowmoss Lantern at night.jpg", "photo:glowmoss-night") }
  ];
  links["Moat-in-a-Box 40 ft::Renders"] = [{ label: null, url: file("Renders/Moat-in-a-Box kit.jpg", "photo:moat-box") }];
  links["Beware of Me Sign::Renders"] = [{ label: null, url: file("Renders/Beware of Me Sign.jpg", "photo:beware-sign") }];

  var na = ["Beware of Me Sign::Assembly Instructions", "Glowmoss Lantern::Compliance & Testing"];

  // Fill the everyday resources across the range so the catalog is about half covered — gaps stay where a
  // real catalog has them (cut sheets, training, the brand-new Cinderpuff Bolster, the moved moat sheet).
  // Which generated photo shows each model in a shop (or, for lair goods, in use).
  function inStorePhoto(m) { return /^Snugrock/.test(m) ? "snugrock-cave" : /^Glowmoss/.test(m) ? "glowmoss-night" : "store-shelf"; }
  var everyModel = [];
  catalog.forEach(function (c) { c.families.forEach(function (f) { f.products.forEach(function (p) { p.models.forEach(function (m) { everyModel.push(m); }); }); }); });
  everyModel.filter(function (m) { return m !== "Cinderpuff Bolster"; }).forEach(function (m) {
    // Slots deliberately left empty: the moat sheet that moved (a broken link), and the two resources the
    // to-do inbox is waiting to place (Hoardmaster 5000 D assembly, Trophy Rack Six training).
    var WAITING = { "Moat-in-a-Box 20 ft::Spec Sheet": 1, "Hoardmaster 5000 D::Assembly Instructions": 1, "Trophy Rack Six::Training": 1, "Goblet Rack Twelve::Renders": 1 };
    function fill(res, make) { var k = m + "::" + res; if (!links[k] && na.indexOf(k) === -1 && !WAITING[k]) links[k] = make(); }
    fill("Spec Sheet", function () { return doc(m, "Spec Sheet", "Spec Sheets"); });
    fill("Product Overview", function () { return doc(m, "Product Overview", "Overviews"); });
    fill("Warranty", function () { return doc(m, "Warranty", "Warranty"); });
    fill("Compliance & Testing", function () { return doc(m, "Compliance Certificate", "Compliance"); });
    fill("Cut Sheet", function () { return doc(m, "Cut Sheet", "Cut Sheets"); });
    // Assembly for everything that is actually assembled (rugs, pillows and signage are not).
    if (!/^Ashweave|^Cinderpuff|^Beware/.test(m)) fill("Assembly Instructions", function () { return doc(m, "Assembly Instructions", "Assembly"); });
    // Training and flyers exist for the lines the reps are pushing this season.
    if (/^Hoardmaster|^Snugrock|^Vaultvane|^Gemtumbler/.test(m)) fill("Training", function () { return doc(m, "Training", "Training"); });
    if (/^Hoardmaster 3000$|^Snugrock Classic$|^Vaultvane 100$|^Pebblechime/.test(m)) fill("Flyer", function () { return doc(m, "Flyer", "Flyers"); });
    if (!/^Moat|^Beware|^Decoy/.test(m)) fill("In-Store Images", function () { return file("In-Store/" + m + " in store.jpg", "photo:" + inStorePhoto(m)); });
  });
  var health = { "Moat-in-a-Box 20 ft::Spec Sheet": "broken" };

  // Files-to-place inbox: two unmatched files and one confident suggestion waiting for a PM.
  var suggestion = file("Incoming/Hoardmaster 5000 D Assembly Instructions.pdf", "assembly:Hoardmaster 5000 D");
  var placements = {
    review: [
      { file: "Hoardmaster 5000 D Assembly Instructions.pdf", url: suggestion, model: "Hoardmaster 5000 D", resource: "Assembly Instructions", folderUrl: u("Incoming"), detail: "suggested — confident name match", at: null },
      { file: "Trophy Rack Six Training.pdf", url: file("Incoming/Trophy Rack Six Training.pdf", "training:Trophy Rack Six"), model: "Trophy Rack Six", resource: "Training", folderUrl: u("Incoming"), detail: "suggested — confident name match", at: null },
      { file: "Goblet Rack Twelve front.jpg", url: file("Incoming/Goblet Rack Twelve front.jpg", "photo:goblet-rack"), model: "Goblet Rack Twelve", resource: "Renders", folderUrl: u("Incoming"), detail: "suggested — confident name match", at: null }
    ],
    pending: [
      { file: "Hoard Insurance Brochure.pdf", url: file("Incoming/Hoard Insurance Brochure.pdf", "flyer:Hoard Insurance:1"), model: "", resource: "Flyer", folderUrl: u("Incoming"), detail: "no match (no model in the name)", at: null },
      { file: "Glowmoss Lantern night mode.jpg", url: file("Incoming/Glowmoss Lantern night mode.jpg", "photo:glowmoss-night"), model: "", resource: "Renders", folderUrl: u("Incoming"), detail: "maybe: Glowmoss Lantern", at: null },
      { file: "Vaultvane hinge detail.jpg", url: file("Incoming/Vaultvane hinge detail.jpg", "photo:vaultvane"), model: "", resource: "In-Store Images", folderUrl: u("Incoming"), detail: "maybe: Vaultvane 100, Vaultvane 200 F", at: null },
      { file: "Spring catalogue cover.jpg", url: file("Incoming/Spring catalogue cover.jpg", "photo:store-shelf"), model: "", resource: "Flyer", folderUrl: u("Incoming"), detail: "no match (no model in the name)", at: null },
      { file: "Clawpost sleeve fitting.pdf", url: file("Incoming/Clawpost sleeve fitting.pdf", "assembly:Clawpost Classic"), model: "", resource: "Assembly Instructions", folderUrl: u("Incoming"), detail: "maybe: Clawpost Classic", at: null }
    ],
    skip: [], noAuto: []
  };
  var orphans = [{ resource: "Renders", file: "Mystery Egg Warmer.jpg", folderUrl: snugFolder, detail: "in a shared folder, matches no model" }];
  var unsorted = [
    { id: "9001", name: "Goblet Rack Twelve lifestyle.jpg", url: file("Uploads/Goblet Rack Twelve lifestyle.jpg", "photo:goblet-rack"), fileId: "demo-9001", resource: "Renders", folder: u("Uploads"), addedBy: "wren@grumbleton.example", addedAt: null },
    { id: "9002", name: "Ashweave rug in the hall.jpg", url: file("Uploads/Ashweave rug in the hall.jpg", "photo:ashweave"), fileId: "demo-9002", resource: "In-Store Images", folder: u("Uploads"), addedBy: "bramble@grumbleton.example", addedAt: null },
    { id: "9003", name: "Trophy Rack Six Cut Sheet v2.pdf", url: file("Uploads/Trophy Rack Six Cut Sheet v2.pdf", "cut:Trophy Rack Six"), fileId: "demo-9003", resource: "Cut Sheet", folder: u("Uploads"), addedBy: "wren@grumbleton.example", addedAt: null }
  ];

  var bugReports = [
    { id: "501", created_at: null, description: "The Hoardmaster spec sheet says “not rated for gems”. My gems are upset.", reporter_name: "Sir Reginald (retired knight)", reporter_email: "reginald@grumbleton.example", context: "Hoardmaster 3000 · Spec Sheet", page_url: "/#/Hoard%20Management/Coin%20Pushers/Hoardmaster/Hoardmaster%203000/Spec%20Sheet", status: "new" },
    { id: "502", created_at: null, description: "Moat-in-a-Box 20 ft spec sheet link goes nowhere. Moat also goes nowhere. Related?", reporter_name: "Bramble (castle intern)", reporter_email: "bramble@grumbleton.example", context: "Moat-in-a-Box 20 ft · Spec Sheet", page_url: "/#/Knight%20Deterrence/Moat%20Kits/Moat-in-a-Box/Moat-in-a-Box%2020%20ft", status: "in_progress" }
  ];
  bugReports.push(
    { id: "503", created_at: null, description: "The Vaultvane assembly instructions say “this is a two-dragon step”. I live alone.", reporter_name: "Pyrella the Restless", reporter_email: "pyrella@grumbleton.example", context: "Vaultvane 100 · Assembly Instructions", page_url: "/#/Hoard%20Management/Hoard%20Doors/Vaultvane/Vaultvane%20100", status: "new" },
    { id: "504", created_at: null, description: "Searched for “hot rock” and found nothing. It is called a Snugrock apparently.", reporter_name: "Bramble (castle intern)", reporter_email: "bramble@grumbleton.example", context: "Search", page_url: "/", status: "resolved" },
    { id: "505", created_at: null, description: "Clawpost sleeves arrived already scratched. Suspect the goblins.", reporter_name: "Sir Reginald (retired knight)", reporter_email: "reginald@grumbleton.example", context: "Clawpost Classic", page_url: "/#/Lair%20Comfort/Scratching%20Posts/Clawpost/Clawpost%20Classic", status: "in_progress" }
  );

  return {
    catalog: catalog, links: links, na: na, health: health, files: files, folders: folders, BROKEN_URL: BROKEN_URL,
    placements: placements, orphans: orphans, unsorted: unsorted, bugReports: bugReports,
    aliases: [{ from: "coin shover", to: "coin pusher" }], handledMisses: [],
    user: { email: "wren@grumbleton.example", name: "Wren Grumbleton" }
  };
});
