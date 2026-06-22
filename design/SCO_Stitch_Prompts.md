# Self-Checkout (SCO) — UI generation prompts for Stitch (Google) / Claude

Copy these into **Stitch by Google** or **Claude** to generate the screen templates, then
bring the result back (Figma link / exported HTML / screenshots + which framework) and I'll
turn it into a working, POS-connected, offline-first app.

## Fidelity legend — read this first
This brief mixes two things; keep them distinct:

- 🟢 **POS-verified** — mirrors the real Foodland/ISS3000 system (DB `POSSDB` + the user
  manual). Flows, item types, tender list, discount types, points, VAT receipt, offline
  behavior. **Do not change these** — they must match `legacy_source/POS_Transaction_Model.md`.
- 🟡 **Placeholder (swap later)** — the **self-checkout interaction itself** (the existing
  POSFront is *cashier-operated*, so the unattended kiosk UX is designed by us to SCO best
  practice) and the **visual brand** (colors, fonts, logo are generic retail, **NOT** Foodland's
  real brand). Swap with Foodland's brand guide when available.

> **How to use**
> 1. Paste **BLOCK 0 (System / Style)** first so the tool locks the look.
> 2. Then paste each **SCREEN** block (one at a time gives the cleanest result).
> 3. Generate **portrait 1080×1920** first (primary kiosk), then ask for **1920×1080 landscape**.
> 4. The 🟢 "POS data" note under each screen is for ME (implementation) — you can include it
>    in the prompt for realism or leave it out.

---

## BLOCK 0 — System / Style (paste once, keep for every screen)

```
You are designing a SELF-CHECKOUT (SCO) touchscreen kiosk app for a Thai grocery/retail
store (Foodland-style supermarket). Customers operate it ALONE, standing, no keyboard.
Design must be calm, glanceable, forgiving, and fast.

DEVICE: large vertical touchscreen, 1080 x 1920 (portrait). Everything finger-operated.
AUDIENCE: general public incl. elderly; bilingual Thai + English on every label.

VISUAL STYLE: clean modern retail, friendly and high-trust (not playful/childish). Soft
depth (subtle shadows, rounded 16–24px cards), generous whitespace, big clear type.

>>> BRAND COLORS ARE PLACEHOLDERS — to be replaced by Foodland's real brand. Use: <<<
- Primary (brand/buttons): #2563EB   Primary-dark: #1D4ED8
- Accent / main CTA (Pay, Start): #F97316
- Success (paid/added): #16A34A   Warning/attendant: #F59E0B   Danger/void/error: #DC2626
- Background: #F8FAFC   Surface/cards: #FFFFFF   Text: #1E293B   Muted text: #475569
- Promo "yellow tag": #FACC15 bg / #713F12 text   Regular "blue tag": #DBEAFE bg / #1E40AF text

TYPOGRAPHY (placeholder, swap to Foodland's): Headings = Rubik (600/700), Body = Nunito Sans
(400/600). Base body ≥ 22px, prices ≥ 32px, primary total ≥ 48px. Never small text.

TOUCH & ACCESSIBILITY (mandatory, do NOT relax):
- All tap targets ≥ 72px tall; ≥ 12px gap between targets.
- Contrast ≥ 4.5:1; clear pressed/active states; 200–300ms transitions only; respect reduced-motion.
- Primary action is always the LARGEST element, anchored bottom.
- Line icons (Lucide/Heroicons), never emoji.
- Persistent header on shopping screens: store logo + bilingual title + ONLINE/OFFLINE status
  dot + language toggle (TH/EN).
- Persistent "Need help / เรียกพนักงาน" button reachable on every shopping screen.
```

---

## SCREEN 1 — Idle / Attract  🟡 (SCO pattern)

```
Design the IDLE / ATTRACT screen for the SCO kiosk (use the System/Style above).
- Full-bleed welcoming hero, large friendly headline: "แตะเพื่อเริ่ม • Touch to start".
- A big pulsing primary START button (accent), centered, ≥ 30% of height.
- Top: store logo placeholder + bilingual "Self Checkout / ชำระเงินด้วยตนเอง".
- Bottom strip: language toggle (TH | EN), accessibility icon, small "Need help" link.
- Optional muted promo carousel band (3 promo cards) below the CTA.
```
🟢 **POS data:** promo cards come from the store's current promotions (deal/promotion tables).

## SCREEN 2 — Scan & Cart (MAIN screen)  🟡 layout · 🟢 data

```
Design the MAIN SCAN & CART screen for the SCO kiosk (use the System/Style above). Layout:
- HEADER: bilingual "Scan your items / สแกนสินค้า", online status dot, language toggle.
- CART LIST (scrollable): each row = thumbnail, name (Thai line + English line), unit price,
  qty stepper (− value +), line total, small remove (trash). Promo rows show a YELLOW "PROMO"
  tag with struck-through original price and the lower price in green; member-price rows show a
  blue "MEMBER" tag. Plain rows show a subtle BLUE "REG" tag.
- EMPTY STATE: big barcode icon + "สแกนสินค้าชิ้นแรก / Scan your first item".
- STICKY SUMMARY (bottom): Subtotal, "You saved ฿x / ประหยัด ฿x" in green, big TOTAL (≥48px),
  item count, and a huge accent button "ชำระเงิน • Pay ฿1,234".
- Secondary row: "Member / สมาชิก", "Coupon / คูปอง", "Help / เรียกพนักงาน".
- A "manual entry / ใส่บาร์โค้ดเอง" affordance opening a numeric keypad.
Show 4 sample grocery lines: 1 normal, 1 yellow-tag promo (strike price), 1 member-price,
1 weighted item (e.g. "ผักกาดหอม 0.452 kg").
```
🟢 **POS data:** lines = `TPSTSalDT` (`FTSkuCode`, `FTSdtBarCode`, `FTSkuAbbName` TH/EN, `FCSdtQty`,
`FCSdtRegPrice`, `FCSdtSalePrice`). Yellow tag = `FCSdtYellowPrice`/`FCShdSubPro`; blue = regular;
member = `FCSdtMemberPrice`. "You saved" = sum of line discounts (`TPSTSalCD`). Whole bill is a
**Sale (transaction type 03)**.

## SCREEN 3 — Item overlays / restricted states  🟢 data

```
Design 4 modal/overlay states for the scan screen (use the System/Style above), each a centered
card over a dimmed cart:
A) WEIGHT ITEM: product name, "Place on scale / วางบนตาชั่ง", live weight (0.452 kg), computed
   price, Confirm (accent) / Cancel.
B) PRICE PROMPT (price-required item): product name, big numeric keypad to enter price, OK/Cancel.
C) ATTENDANT NEEDED (age-restricted e.g. alcohol, or void/override): warning style (#F59E0B),
   "Waiting for staff / รอพนักงานยืนยัน", spinner, reason text, Cancel.
D) UNKNOWN BARCODE: "Item not found / ไม่พบสินค้า", barcode shown, "Try again" + "Call staff".
```
🟢 **POS data:** weight = `FCSdtWeight`/`FTSdtWeight`; price-prompt = `FTSkuStaPpt4Price`; override =
`FTSdtPriceOverideFlag`+`FTEmpCodeOveride` (staff approval); restricted = age/hours rule (alcohol
sale hours). Unknown barcode = not in `TCNMSKUBarcode`/`TCNMSKU`.

## SCREEN 4 — Member & Discounts  🟢 data

```
Design the MEMBER & DISCOUNT screen (use the System/Style above).
- Two tabs: "Member card / บัตรสมาชิก" and "Coupon / คูปอง".
- MEMBER tab: large numeric keypad to enter member or phone number, plus a "Scan card / รูดบัตร"
  affordance. After lookup show member name (masked), current points, and "Member price applied".
- COUPON tab: big "Scan coupon / สแกนคูปอง" target + manual code entry; a list of applied
  discounts each labelled by its type and amount.
- Keep the running TOTAL visible at top. Bottom: primary "Done / เสร็จสิ้น".
```
🟢 **POS data:** member lookup by card OR phone, **must work offline** (`TCNMMallCard` /
`TCNMCstTelMtn`, cached). Applied discounts → `TPSTSalCD.FTScdDiscType` using real labels from
`TCNMDiscountType`: Member Card, Employee Card, Cooperate Coupon, Cooperate Brand, Discount
Coupon, Transaction Discount. Points preview → `TPSTSalePoint`.

## SCREEN 5 — Payment method  🟢 data

```
Design the PAYMENT METHOD screen (use the System/Style above).
- Top: "Amount due / ยอดที่ต้องชำระ ฿1,234.00" (≥48px).
- Grid of large payment tiles (≥160px), each line-icon + bilingual label:
  Cash / เงินสด · Credit Card / บัตรเครดิต · QR PromptPay · TrueMoney Wallet · Alipay ·
  WeChat Pay · Gift Voucher / บัตรกำนัล · Voucher · Member Points / คะแนนสะสม.
- Tiles have subtle shadow + pressed state. Back link top-left; "Help" bottom.
- Note line: "Follow the payment terminal / ทำตามเครื่องชำระเงิน".
```
🟢 **POS data:** tiles = real tenders from `TCNMTenderMtn` (T001 Cash, T002 Credit, T004 Gift
Voucher, T007 True Point, T009 True Wallet, T010 Voucher, T016 Alipay, T019 QR PromptPay, T021
WeChat…). Each tender selection → a `TPSTSalRC` row (`FTTdmCode`, `FCSrcAmt`). Split tender allowed
(multiple `FNSrcSeqNo`).

## SCREEN 6 — Payment in progress  🟡

```
Design the PAYMENT-IN-PROGRESS screen (use the System/Style above):
- Centered animated indicator; big bilingual instruction that adapts to the method:
  card → "Follow the terminal / ทำตามเครื่องชำระเงิน"; QR → "Scan this QR / สแกน QR นี้" with a
  PromptPay QR placeholder. Show the amount and a Cancel button. Reassuring, uncluttered.
- Also design a PAYMENT FAILED variant: red icon, "Payment failed / ชำระเงินไม่สำเร็จ",
  "Try again" + "Choose another method" + "Call staff".
```

## SCREEN 7 — Receipt / Thank you  🟢 data (real Foodland slip)

```
Design the THANK YOU / RECEIPT screen (use the System/Style above).
- Big green success check + "Payment complete / ชำระเงินสำเร็จ".
- Left/main: a realistic Thai ABBREVIATED TAX INVOICE preview on thermal-receipt styling:
    Foodland                         Tel.02-1905-111
    Branch ID                                   00xx
    TAX INVOICE (ABB.) / ใบกำกับภาษีอย่างย่อ
    Terminal 00xx   Cashier SCO   dd/mm/yyyy hh:mm
    -------------------------------------------------
    <items: name … qty x price … amount, with * marking promo/yellow-tag lines>
    -------------------------------------------------
    Sub Reg / Sub Pro / Total / Discount / VAT 7% / NET
    Tender (e.g. QR PromptPay)            change
    Points earned: xx
- Right/bottom: choices "Print receipt / พิมพ์ใบเสร็จ" and "E-receipt (QR) / ใบเสร็จออนไลน์"
  with a QR placeholder; plus a smaller "Request full tax invoice / ขอใบกำกับภาษีเต็มรูป".
- "Please take your items / กรุณารับสินค้า" reminder; auto-reset countdown chip + "Done".
```
🟢 **POS data:** it's an *abbreviated tax invoice* (ใบกำกับภาษีอย่างย่อ) — `*` marks yellow-tag/promo
lines; totals = `FCShdSubReg`/`FCShdSubPro`/`FCShdTotal`/`FCShdDis`/`FCShdVat`/`FCShdGrand`; tender from
`TPSTSalRC`; points from `TPSTSalePoint`. "Full tax invoice" captures name/address/Tax ID →
`TPSTVatTransAddess`/`TPSTSalVatHD`.

## SCREEN 8 — Help / Offline / system states  🟢 behavior

```
Design these small overlays/banners (use the System/Style above):
A) CALL ATTENDANT: friendly card "Staff are on the way / พนักงานกำลังมา", animated dots, reason
   chips (price check, void item, bagging help, age check), Close.
B) OFFLINE BANNER: non-blocking top banner "Working offline — your sale will sync automatically /
   ออฟไลน์อยู่ ระบบจะซิงค์ให้อัตโนมัติ" with an amber dot. Checkout MUST stay fully usable offline.
C) SESSION TIMEOUT: gentle "Are you still there? / ยังอยู่ไหมคะ" with Continue / Start over.
```
🟢 **POS data:** offline = the real producer→consumer pattern: capture the sale locally with each
child row `FTStaSentOnOff='0'`, queue it, and let the sync push to POS/HQ when back online (same
handshake as POSFront → ServiceTransfer). Never block a sale on connectivity.

---

## What to send back to me

1. **The template** — Figma link, exported HTML/CSS, or screenshots of every screen.
2. **Foodland brand assets** (so I replace the 🟡 placeholders): logo, real brand colors, fonts,
   and ideally a photo of a real Foodland receipt for exact wording/layout.
3. **Target framework** (else I'll recommend): React/Next + Tailwind (web/PWA) or React Native /
   Flutter (native). **Web/PWA is the easiest to make offline-capable** for a kiosk.
4. **Kiosk hardware**: screen size/orientation, barcode scanner, scale, payment terminal/EDC, OS.

I'll then wire the screens to the **real POS data contract** (SKU/promo/tender/discount/points from
`legacy_source/POS_Transaction_Model.md`), implement the **offline cache + local sale queue**, and
the **write-back to POS** (`TPSTSal*` rows stamped `FTStaSentOnOff='0'` → trigger → ServiceTransfer).
