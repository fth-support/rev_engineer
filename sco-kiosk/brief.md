# Foodland SCO — Project Brief & Handoff

> Handoff document for continued development (e.g. with Claude Code).
> It captures what exists today, how it's built, and what is real vs. mocked.

---

## 1. What this is

An **interactive self-checkout (SCO) prototype** for **Foodland Supermarket (Thailand)**,
built as a single self-contained HTML component. It demonstrates the complete shopper loop
on a **portrait 1080×1920 kiosk**, plus staff and edge-case flows. It is a **design /
UX prototype** — not production software. All data and integrations are mocked in-file.

Context: Fujitsu is proposing a **New SCO** on the **Fujitsu POS API Center** for Foodland
(reference: *TMG × Fujitsu POS API Center & SCO kickoff, 22 Feb 2024*). This repo is the
front-end demo of that proposed experience.

Two deliverables exist in the wider project:
1. **The kiosk demo** — this repo (`index.html`).
2. **A function-specification slide deck** (separate file `Foodland SCO Pitch v2.dc.html`),
   exported to PPTX — not part of this repo unless asked.

---

## 2. Tech / architecture of the prototype

- **One file:** `index.html`. No framework install, no build, no bundler.
- It is a **"Design Component" (DC)**: the file has a `<x-dc>` template (markup) and a
  `<script data-dc-script>` block containing `class Component extends DCLogic { … }` (state +
  logic). `support.js` is the runtime that renders it. **Only `index.html` is edited.**
- **Styling:** inline styles throughout (no external CSS). Fonts: Rubik + Nunito Sans +
  Noto Sans Thai (Google Fonts). Brand red `#E2231A` / `#C2161A`, dark `#211B1A`,
  cream `#F6F1EE`, gold `#F5B800`, green `#1F8A4C`.
- **Scaling:** the kiosk is authored at 1080×1920 and scaled to fit via a computed transform
  in the logic (`calcScale()`), with a left-hand **Prototype controls** panel for reviewers.

### If re-platforming
The logic is plain JS and portable. A production build would more naturally be React/Vue with
the same state shape (below). The screen markup can be lifted largely as-is.

---

## 3. The full flow (state machine)

State lives in `this.state` in the logic class. Key field: `screen`.

```
idle ──START──▶ memberfirst ──(apply | skip)──▶ scan
scan ──Pay──▶ (member overlay if not yet identified) ──▶ payment
payment ──pick method──▶ paying ──success──▶ receipt ──done/timeout──▶ idle
                              └─fail──▶ paying(failed) ──retry/other──▶ payment
```

Screens (each is a `data-screen-label` section, shown/hidden by `screen`):
| screen value | Purpose |
|---|---|
| `idle` | Attract: Touch to start, promo ad carousel, OPEN 24H, lang TH/EN/CN, staff + help |
| `memberfirst` | After START: enter phone / scan member QR / skip |
| `scan` | Core cart: scan strip, line items, qty ±, delete, subtotal/total, member banner |
| `member` | Member & coupon overlay/screen reachable from scan/pay |
| `payment` | Method grid: card, QR PromptPay/QRCS, Alipay, WeChat |
| `paying` | Processing (mirrors EDC) → success or failed |
| `receipt` | Success, receipt options (none / print ABB / e-receipt), auto-reset countdown |
| `staff` | PIN-gated store-management menu |
| `idleDown` | "Out of service" screen |

**Transitions / loading:** `loadThen(msg, sub, fn, delay)` shows a branded full-screen
loader before key navigations (verify membership, validate coupon, prepare payment, start txn).

**Overlays** (`this.state.overlay`): `weight` (scale-priced produce), `price` (price-required
item), `attendant` (age/staff approval, e.g. alcohol), `unknown` (barcode not found),
`help` (call staff), `keypad` (manual barcode), `timeout` ("Are you still here?"),
`coupon` (enter code).

---

## 4. Data model (all mocked in `index.html`)

- **`catalog`** — array of products: `{ id, name, nameTh, price, barcode, type }` where
  `type` ∈ regular / weighed / age-restricted / price-required. Edit here to change demo items.
- **`cart`** — array of line items with qty; totals/discounts computed in the logic.
- **`member`** — `{ phone (PDPA-masked e.g. 092629xxxx), name, points }` or null.
- Discount/total maths and the points calc are in the logic class (not a backend).

---

## 5. Screens & functions (per the function spec)

1. **Startup** — touch/scan to start · Need Help (red lane LED) · language ×5 + voice · volume.
2. **Scan product** — scan barcode/QR · item line (name, unit price, qty, line total at
   regular price) · running item count · qty ± · delete · group repeats · cancel item/order.
   Safeguards: inactivity timeout ("Are you still here?"), confirmed cancel-order.
3. **Member & coupon** — prompt on Pay · phone/member no. · app QR · co-brand-card QR ·
   scan coupons · PDPA mask · skip (guest).
4. **Payment** — order summary + discounts (member coupon, member card, auto-trandis,
   deal & free) · card · QR PromptPay/QRCS · Alipay · WeChat — all via EDC.
5. **Processing & receipt** — EDC-mirrored guidance · processing state · retry on failure
   (no money taken) · receipt none / print ABB / e-receipt (e-receipt → Foodland landing page).
6. **Store management (staff)** — PIN login · reprint ABB · close shift · close program ·
   out of service · closing report · about SCO.
7. **Edge cases** — weighed produce (scale) · age-restricted (staff approval) · unknown
   barcode · call staff.

---

## 6. Integration scope (from kickoff doc — NOT implemented, for production)

**Outbound → SCO:** master data (catalog & images, POS/SAP) · promotions (BBY, POS) ·
loyalty profile & points (Member) · coupon validation (App).
**Inbound → POS:** finalized transaction · EDC link (BBL) · member earn-point · printing.

**Architecture (target):**
`Customer ERP ──sync schedule──▶ POS Server HQ ──realtime──▶ POS Store Server ──realtime──▶
POS API ──realtime──▶ SCO`. The **SCO is treated as one POS terminal** on the store network.
Cloud: SCO client → AWS API Gateway → Fujitsu API Engine (cache / DB / CICD) → Foodland
loyalty, POS and member servers. Fujitsu provides the lane + API center; Foodland keeps its
POS/loyalty/member systems of record.

**Payments:** EDC (BBL) for credit/debit, QRCS, PromptPay, Alipay, WeChat; EDC prints the slip.

**Thailand fiscal/localization:** registered POS Tax ID on every receipt · abbreviated tax
invoice (ABB) issued by SCO · closing report · full tax invoice & Sale VAT report via Fujitsu
API. ABB must carry: branch & tax ID, transaction no., date/time, items & VAT type, tender,
member ID, points, reference barcode, footer message.

**Promotions (Fujitsu BBY codes):** self-serve at SCO — P001 outright/time, P002 consign,
P011 lower price, P003 auto-trandis, P010 member price, P012 SKU point, D001–D014 deal,
F001–F006 free. Operation-controlled / API-ready — P005 coupon redeem, P006/P007 cooperate,
P009 on-top, P013 transaction discount, P015 line markdown, M001–M006 manual, I001 installment.
*(Codes are from the Fujitsu BBY master in the kickoff doc; confirm final scope & naming with
Foodland. Anything labelled "The Mall" was removed as not confirmed for Foodland.)*

---

## 7. Known stubs / not-yet-real (priority backlog for production)

- **No backend** — catalog, members, prices, discounts, points all hardcoded in `index.html`.
- **No real hardware** — scanner, scale, EDC terminal, receipt printer, camera feed are all
  simulated. (Scan advances through a scripted product list.)
- **"Reach mode (a11y)"** — there is a toggle stub in the controls panel but it does NOT yet
  shift/enlarge the UI. Either implement (drop action zone toward the bottom + larger targets)
  or remove the toggle. Flagged as not-functional.
- **Voice guidance & 5-language** — UI shows TH/EN/CN; voice and JP/AR are spec, not built.
- **Payments** — `paying` is a timed simulation; the controls panel can force success/failure.
- **Receipt / ABB** — visual only; no real fiscal document generation.

---

## 8. Repo files

```
sco-kiosk/
├── index.html        ← the kiosk: template + logic (edit this)
├── support.js        ← DC runtime (do not edit)
├── .nojekyll         ← GitHub Pages: serve files as-is
├── README.md         ← run / structure / edit guide
├── brief.md          ← this document
└── assets/           ← foodland_logo, foodland_open24, ads1–3_*, ad1–4
```

Run: open `index.html` via a local server (`npx serve .`) — images 404 on `file://`.
Deploy: push to GitHub → Settings → Pages → `main` / root.

---

## 9. Suggested next steps for Claude Code

1. **Decide platform** — keep as single DC HTML, or port to React/TS with the same state shape.
2. **Extract data** — move `catalog` / member / pricing into JSON or an API layer behind a
   small service interface, so the UI no longer hardcodes them.
3. **Define the SCO↔POS API contract** — item lookup, promotion/quota, coupon validate,
   member lookup & earn-point, transaction submit, ABB/closing print (map to Fujitsu API #s
   in the kickoff doc: display authorize #0, customer/item/promotion/coupon #1·#2·#7,
   send transaction & ABB #10·#11, logout #13, closing #14, EDC interface).
4. **Hardware abstraction** — interfaces for scanner, scale, EDC, printer, camera; mock impls
   for the demo, real impls for the lane.
5. **Finish or cut "reach mode"**; build voice + remaining languages if in scope.
6. **Harden edge cases** — partial-payment, void/refund, network loss mid-payment, receipt
   reprint, shift reconciliation.
