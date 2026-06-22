# POS Transaction Model — Verified Reference (POSFront / ISS3000)

**Purpose.** A build-ready, *verified* model of how the POS records a transaction, so a
rebuild in any language saves data that matches what the existing **reports** and
**ServiceTransfer** sync expect. Every fact below is cross-checked against four sources:

| Source | What it confirmed |
|---|---|
| VB6 source (`POSFront/`, `ServiceTranfer`) | save/sync logic, flags, triggers |
| **User manual** (`FL-POS_ML-IP-NewUI-User Manual.docx`) | business meaning of each screen/field |
| **Live schema** (`pos_export/_columns.csv`, `_keys.csv`) | real columns + data types |
| **Live data + metadata** (`TSysTransType`, `TCNMTenderMtn`, `TCNMDiscountType`, row counts) | actual codes in use |

> Field naming (Hungarian): `F` + type + entity-prefix. Type: **T**=text, **C**=currency/float,
> **D**=date, **N**=int. Entity: `Shd`=Sale HeaDer, `Sdt`=Sale DeTail, `Src`=Sale ReCeipt(tender),
> `Scd`=Sale disCount Detail, `Spo`=Sale POint. Every transaction row carries audit columns
> `FDDateIns/FTTimeIns/FTWhoIns` + `…Upd`.

---

## 1. Transaction types — `TSysTransType` (`FTSttTranCode`)

Each bill's `FTShdTransType` is one of these. **Group** drives sign (SALE = +, RETURN = −).

| Code | Description | Group | Code | Description | Group |
|----|----|----|----|----|----|
| 01 | Open Terminal | | 22 | Other Receive | SALE |
| 02 | Start Bank | | 23 | Other Disbursement | RETURN |
| **03** | **Sale Transaction** | SALE | 24 | Deposit Money | |
| **04** | **Return Transaction** | RETURN | 25 | Training Mode | |
| **05** | **IPV (Void)** | RETURN | 26 | IPV VIP Service Sale | RETURN |
| **06** | **Deposit Create** (← Big Lot) | | 27 | IPV Tax Except Sale | RETURN |
| 07 | Redeem | | 28 | IPV Duty Free Sale | |
| 08 | Close Terminal | | **29** | **Suspend** (hold bill) | |
| 09 | Print Closing Report | | 30 | Create VIP | |
| 10 | VIP Service Sale | SALE | 31 | Total Report | |
| 11 | Tax Except Sale | SALE | 32 | IPV Other Receive | |
| 12 | Pick Up (cash out) | | 33 | IPV Other Disbursement | |
| 13 | Duty Free Sale | | 34 | Copy Transaction to Device | |
| 14 | IPV Return | SALE | 35 | Copy Transaction From Device | |
| 15 | Voucher | | 36 | Update Master | |
| 16 | Cancel BigLot | | 40 | Bill Payment | |
| 17 | IPV Redeem | | 45 | Tender Declaration | |
| 18 | No Sale | | | | |
| 19 | Add Deposit | | | | |
| 20 | Add Deposit Cancel | | | | |
| 21 | Cancel Voucher | RETURN | | | |

> **Manual ↔ code confirmed:** Big Lot is stored as **Deposit (06)** with a Big-Lot flag (full
> deposit); IPV (05) voids same-day Sale/Return; Suspend/Recall = type 29.

---

## 2. The sale transaction — 5 linked tables

One completed bill spans up to five tables, keyed by **`FTTmnNum` (terminal) + `FTShdTransNo` (bill no) + `FDShdTransDate`**:

```
TPSTSalHD  (header, 1 row/bill)   — 105,105 bills
   │ 1:N
   ├── TPSTSalDT   (line items)   — 828,888 rows   FNSdtSeqNo
   ├── TPSTSalRC   (tenders/pay)  — 124,534 rows   FNSrcSeqNo
   ├── TPSTSalCD   (line discounts)— 238,866 rows  FNSdtSeqNo+FNScdSeqNo
   └── TPSTSalePoint (member pts)  — 1,264,796 rows
```

Every child row carries **`FTStaSentOnOff`** (`'0'`=pending → `'1'`=synced) — the flag
**ServiceTransfer** polls to push to HQ. HD-first rule: HD is only sent after all children = `'1'`.

### 2.1 `TPSTSalHD` — header (key fields)
| Field | Meaning (from manual) |
|---|---|
| `FTStmCode` / `FTTmnNum` | store / terminal number |
| `FTShdTransNo` · `FTShdTransType` · `FDShdTransDate` · `FTShdTransTime` | bill no · type (§1) · date · time |
| `FTEmpCode` · `FTSpnCode` | cashier · salesperson |
| `FTShdTrainMode` | training-mode flag (type 25) |
| `FTCstCardCode` / `FTCstCardType` / `FTCstCode` / `FTCstVIPCode` | member card / customer / VVIP |
| `FTEcdCode` / `FTEcdCardType` | employee card (staff discount) |
| `FCShdDepositAmt` `FDShdDepositDueDate` `FCShdDepositBal` `FTShdDepReTransNo` | deposit / Big-Lot (type 06) |
| `FCShdCoCpnRate` / `FCShdCoBrandRate` | Cooperate Coupon / Brand rate |
| `FTShdSexCode` `FTShdNatCode` `FTShdAgeCode` | customer demographics |
| `FCShdVATRate` `FTVatCode` `FCShdVat` `FCShdVatable` `FCShdNonVat` | VAT (abbreviated tax invoice) |
| **`FCShdSubReg`** / **`FCShdSubPro`** | **blue-tag** (regular) / **yellow-tag** (promo) subtotal |
| `FCShdTotal` `FCShdDis` `FCShdCpnDis` `FCShdAftDisChg` `FCShdGrand` `FCShdRnd` | Total · discount · coupon disc · net · grand total · rounding |
| `FTShdStaCancel` | cancelled/void status |

> **Totals logic (manual):** `Total` = items − Deal/Promo/PromoTime/Cooperate/Line disc/markdown;
> `Total Discount` = the discount-screen items (Member, Employee, Coop Coupon/Brand, Promo Coupon,
> Transaction Disc/Mkdn); `Net Sale = Total − Total Discount`; then VAT, rounding → `Grand`.

### 2.2 `TPSTSalDT` — line items (108 cols; key fields)
`FNSdtSeqNo` (line seq) · `FTSkuCode` / `FTSdtBarCode` (UPC) · `FTSkuAbbName` (name) · `FCSdtQty` ·
`FCSdtRegPrice` (Reg.Price) · `FCSdtSalePrice` (Sale.Price) · `FCSdtSaleAmt` (SubTotal) · `FCSdtNet` ·
`FCSdtVat` `FCSdtVatable` `FTTaxCode` `FCSdtTax` · `FCSdtBSubReg` / `FCSdtBSubPro` (blue/yellow base) ·
`FCSdtBluePrice` / `FCSdtYellowPrice` / `FCSdtMemberPrice` · `FCSdtMemberPoint` (points earned) ·
`FCSdtWeight` `FTSdtWeight` (weight items) · `FTSdtPriceOverideFlag` `FCSdtPriceOveride` `FTEmpCodeOveride` ·
`FTSdtConsignFlag` (consignment) · `FCSdtReturnAmt` `FCSdtReturnQty` · `FCSdtCostIn`/`CostEx` · `FTStaSentOnOff`.

> **Item types (manual):** *Prompt-for-Sale* (price 0 → cashier enters), *Weight* (barcode embeds
> weight+price), *Normal*. `*` / **yellow tag** = promo item (`Sub Pro`); **blue tag** = regular (`Sub Reg`).

### 2.3 `TPSTSalRC` — tenders / payment (56 cols; key fields)
`FNSrcSeqNo` · **`FTTdmCode`** (tender code → §3) · `FTTdmType` · `FCSrcAmt` `FCSrcNet` `FCSrcChg`
(change) · `FCSrcFAmt` (foreign-currency amount) · `FTRteCode` `FCSrcRteFac` (FX rate) ·
`FTSrcCardNo` `FTSrcCardExp` `FTApprCode` (credit card) · `FTEDCBank` `FTEDCModel` `FTEDCCreditType`
(EDC online) · `FTSrcBBYNo` (Big-Lot bonus-buy) · `FCSrcRetAmt` (return) · `FTStaSentOnOff`.

### 2.4 `TPSTSalCD` — line discounts (70 cols; key fields)
`FNSdtSeqNo`+`FNScdSeqNo` (which line, which discount) · **`FTScdDiscType`** (→ §4) · `FNDctNo`
`FTDctName` · `FCScdDiscPer` / `FCScdDiscAmt` / `FCScdDiscNew` (% / baht / new-price) ·
`FTScdCardID` `FDScdExpDate` `FTScdCardType` (member/employee card) · `FTScdCstID` ·
`FCScdBSubReg`/`FCScdBSubPro` · `FTScdTierID` · `FTScdBBYNo` · `FTStaSentOnOff`.

### 2.5 `TPSTSalePoint` — member points (25 cols)
`FTSpoMemID` (member id — **tokenized via SafeNet** on sync) · `FCSpoPoint` (± points) ·
`FCSpoAmt` (spending) · `FTSpoType` · `FTEventID` · `FTSpoBBYNo` · `FTStaSentOnOff`.
> Sign flips negative for Return/Void/Cancel (per `W_UPDxUpdatePoint`).

---

## 3. Tender / payment types — `TCNMTenderMtn` (`FTTdmCode`)

| Code | Name | Code | Name | Code | Name |
|---|---|---|---|---|---|
| T001 | Cash | T009 | True Wallet | T017 | Online Shopping |
| T002 | Credit Card | T010 | Voucher | T018 | Self Service Disc. |
| T003 | Installment Credit | T011 | AirPay | T019 | QR PromptPay |
| T004 | Gift Voucher | T012 | Foreign Currency | T021 | WeChat |
| T005 | E-Wallet | T013 | Cash Card | T024 | QR Code |
| T006 | Vendor Coupon | T014 | Dolfin Wallet | T027 | M Gift Card |
| T007 | True Point | T015 | ON Top | T028 | M Special Gift Card |
| T008 | Foodland Acc (Store Credit) | T016 | Alipay | T035–T037 | QR/Alipay/WeChat (BBL) |

> Manual lists Cash, Credit (online EDC / offline / installment), Gift Voucher, E-Wallet, Vendor
> Coupon, True Point, Voucher, Foreign Currency, Foodland Account / Employee Credit, True Wallet,
> Alipay, KBank Online — all present here. Newer wallets/QR were added over time (T013+).

---

## 4. Discount types — `TCNMDiscountType` (`FTScdDiscType`)

| # | Name | # | Name | # | Name |
|---|---|---|---|---|---|
| 1 | Deal & Free Item | 9 | Cooperate Coupon | 17 | Ontop Installment |
| 2 | Promotion Timesale | 10 | Cooperate Brand | 21 | Member Price |
| 3 | Promotion | 11 | Member Card | 22 | Member Price + Consignment |
| 4 | Cooperate Discount | 12 | Employee Card | 24 | Tran Mark Cash Coupon |
| 5 | Line Discount | 13 | Transaction Disc. | 25 | Tran Mark Member Point |
| 6 | Line Markdown | 14 | Transaction Markdown | 26 | Tran.Mark AIS/TRUE |
| 7 | Auto Tran Disc. | 15 | Ontop Discount | 27 | Tran.Mark Bank |
| 8 | Discount Coupon | | | 28 | MCardRedeem MCpn |

> Matches the manual's Discount screen exactly (Member/Employee Card, Cooperate Coupon/Brand,
> Discount/Promotion Coupon, Transaction Disc/Markdown, Manual Discount/VVIP, Line Disc/Markdown).

---

## 5. Key workflows → tables/types

| Workflow (manual) | TransType | Writes |
|---|---|---|
| **Sale** (scan → discount → tender → receipt) | 03 (10/11/13 variants) | HD + DT + RC (+ CD, Point) |
| **IPV / Void** (same-day, by terminal+bill no) | 05 (+ 14/26/27) | reverses referenced bill |
| **Return** (no-ref or Refer Sale Return) | 04 | HD + DT + RC (− amounts) |
| **Big Lot** (Bonus-Buy from HQ) | 06 Deposit + BigLot flag | HD(deposit) + DT + RC(`FTSrcBBYNo`) |
| **Suspend / Recall** (hold bill) | 29 | `TPSTHoldHD/DT/RC/CD` |
| **Pickup Money** (cash out, rounds 14/17/20/22h) | 12 | cash declaration |
| **Tender Declaration** (end of round) | 45 | denomination + tender totals |
| **Deposit Money** (cash in) | 24 | |

All sale/return/deposit rows land with `FTStaSentOnOff='0'` → the SQL trigger promotes completed
working-table bills into `TPSTSal*` → **ServiceTransfer** pushes to HQ and flips the flag to `'1'`.

---

## 6. Report dependency (do-not-break list)

Reports are defined in **`TSysReport`** (`FTRptNameSql` = .rpt file, `FTRptTable` = driving table,
`FTRptField` = filter keys). 30 reports bind to `.rpt` files; nearly all drive off **`TPSTSalHD`**
(filter `FTShdTransNo, FDShdTransDate, FTShdTransTime, FTEmpCode`), a few off `TPSTHoldHD`
(Suspend/Recall) or `TCNMTerminalMtn` (Location/Audit). **Implication for the rebuild:** keep the
`TPSTSal*` / `TCNM*` table & column names and the transaction-type / tender / discount codes intact,
or the Crystal reports will fail to bind.

*(Full per-report list → see the planned `Report_Field_Map.md`.)*

---

## 7. Real data volumes (top tables)

`TPSTSalePoint` 1.26M · `TPSTSalDT` 829K · `TPSTSalCD` 239K · `TPSTSalRC` 125K · **`TPSTSalHD` 105K bills** ·
`TCNMSKU` 92K products · `TCNMCstMtn` 34K customers. Confirms a live, production-scale dataset.
