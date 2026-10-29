# Grumbleton & Sons Hub — a Product Hub demo

**Live demo:** https://product-hub-demo-olive.vercel.app

This is a working demo of **Product Hub**, an internal web app that gives a sales and product team one place
to find every product's spec sheets, renders, assembly instructions, warranties and training — and gives the
product managers the tools to keep that catalog complete.

**Every product, person and file here is fictional.** Grumbleton & Sons sells fine goods for dragons. The
product illustrations and documents (spec sheets, assembly instructions, warranties, certificates, training
guides, flyers) were made for this demo.

## What to try

- **Find:** browse categories → products → models. A product shows two cards: the top one holds the files
  every model shares (the Hoardmaster warranty and flyer), the one below holds only what differs for the
  model you picked. Open a spec sheet or a render gallery in the preview.
- **Search:** type "hoard", "boulder", or a typo like "hoardmastr".
- **Favorites and Select mode:** pin a model; pick several files and download or share them together.
- **Progress** (product-manager view):
  - **Add files › Link a product folder**: the box already holds a sample folder (`02 Gemtumbler`). Press
    Preview. The files already linked start ticked where they sit, and the new flyer is offered for both
    models. Approve & save to link it.
  - **To do**: a broken link, files waiting to be placed, a file that **moved** to another folder (press
    **Repair** to follow it), and a rename the nightly check **repaired** by itself.
  - Coverage by product, and the section editor: add, remove or mark links N/A, then Save.
- **Metrics:** adoption, what people open, what they search for and don't find, coverage over time.
- **Admin:** bug reports from the (fictional) team.

Changes work exactly as they do in the real app, but they only live in your browser tab: a reload resets
the demo.

## How the demo works

The real app is a static front end plus serverless functions backed by Smartsheet (structure and links) and
Microsoft Graph / SharePoint (the files), behind Microsoft sign-in. The demo runs the **same front-end code**
with two pieces swapped:

- `demo/demo-signin.js` stands in for Microsoft sign-in and hands the app a pretend product-manager session.
- `demo/backend.js` wraps `window.fetch` and answers every `/api/*` and Graph call from `demo/seed.js`,
  with the same response shapes the real endpoints return — including compare-and-swap conflicts on saves.

There is no server, no database and no third-party script. The Content-Security-Policy stays strict
(`script-src 'self'; style-src 'self'`).

## License

MIT — see `LICENSE`.
