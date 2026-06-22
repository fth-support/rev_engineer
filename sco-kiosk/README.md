# Foodland SCO — Self-Checkout Kiosk (Interactive Demo)

An interactive, end-to-end self-checkout (SCO) prototype for Foodland Supermarket.
Tap through the full loop on a portrait kiosk: **Idle → Scan → Member → Payment → Receipt**,
plus item overlays (weight / price / age approval / unknown), help, offline, timeout and a
PIN-protected staff settings screen.

Bilingual (ไทย / English / 中文) UI, Foodland branding, rotating promo ads.

---

## Run it

It is plain HTML — **no build step, no install.**

- **Locally:** open `index.html` in a browser. (For the images to load, open it through a
  local web server rather than `file://` — e.g. `npx serve .` or VS Code "Live Server".)
- **GitHub Pages:** push this folder to a repo, then
  *Settings → Pages → Branch: `main` / root → Save*.
  It will be served at `https://<username>.github.io/<repo>/`.

---

## Project structure

```
sco-kiosk/
├── index.html        ← the kiosk (edit THIS — markup + logic live here)
├── support.js        ← runtime that renders index.html (do not edit)
├── .nojekyll         ← tells GitHub Pages to serve files as-is
├── README.md
└── assets/
    ├── foodland_logo.png
    ├── ads1_snack.png  ads2_milk.png  ads3_sauce.png   ← idle full-screen promo carousel
    └── ad1.png  ad2.png  ad3.png  ad4.png              ← scan-screen promo strip
```

---

## How it's built

`index.html` is a single self-contained component with two parts:

1. **Template** — the markup between `<x-dc> … </x-dc>` (all screens, inline-styled).
2. **Logic** — a `class Component` in the `<script data-dc-script>` block at the bottom:
   state (`screen`, `cart`, `member`, `overlay`, …), navigation, the product catalog,
   pricing/discount maths, and the prototype-control panel handlers.

`support.js` reads those two parts and renders the live UI. You only ever edit `index.html`.

### Common edits
- **Products / prices** — edit the `catalog` array in the logic class.
- **Promo images** — replace the files in `assets/` (keep the same names, or update the
  `src="assets/…"` references in the template).
- **Copy / labels** — edit the text directly in the template markup.
- **Flow / behaviour** — edit the handler methods in the logic class.

---

## Notes
- Demo only — no real payment, scale, scanner or backend; all data is mocked in `index.html`.
- The left **Prototype controls** panel is a reviewer aid (jump to any screen / overlay,
  toggle online-offline, force a payment failure). Hide it for a clean kiosk view by removing
  the control-panel block at the top of the template.
