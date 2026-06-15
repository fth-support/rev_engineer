# POSSDB — Local DB Table Inventory (Schema Reference)

> สร้างจากการสแกน `POSSDB.bak` (Local SQL Server Express ของ POSFront) — ตัวไฟล์ `.bak` (~1.2GB) ไม่ถูก commit ขึ้น repo ด้วยเหตุผลด้านขนาดและข้อมูลจริง จึงใช้รายการตารางนี้แทน  ·  ปรับปรุง 2026-06-15

พบตารางทั้งหมด **235** ตาราง (กรอง catalog/replication noise แล้ว). คำอธิบายของตารางหลักยืนยันกับ source code; ที่เหลืออนุมานจากชื่อแบบ Hungarian.

> หมายเหตุ: ตารางทำงานต่อเครื่อง `TPSHD<Tmn>` / `TPSDT` / `TPSRC` / `TPSCD` (เช่น `TPSHD01`) ถูกสร้าง/ลบแบบ dynamic ต่อ Terminal จึงอาจไม่ปรากฏในรายการนี้ — ดู Data Dictionary หัวข้อ 3.0

## POS Transaction Tables  (46)

*ตารางธุรกรรมการขาย/ฝาก/พัก/Voucher/คะแนน*

| Table | Description |
| --- | --- |
| `TPSTCV` | Credit voucher. |
| `TPSTCancelItems` | Cancelled line items log. |
| `TPSTClose` | End-of-day close records. |
| `TPSTDeal` | Promotion deal transactions. |
| `TPSTDealAct` | POS transaction |
| `TPSTDealActPreDay` | pre-day (next operation date) staging, POS transaction |
| `TPSTDealSKU` | Deal-to-SKU mapping. |
| `TPSTDepositCD` | Deposit card detail. |
| `TPSTDepositDT` | Deposit detail. |
| `TPSTDepositHD` | Deposit / layaway header. |
| `TPSTDepositRC` | Deposit receipt. |
| `TPSTHandFreeDT` | POS transaction, detail |
| `TPSTHandFreeHD` | POS transaction, header |
| `TPSTHol` | POS transaction |
| `TPSTHold` | POS transaction |
| `TPSTHoldCD` | Held bill card detail. |
| `TPSTHoldDT` | Held bill detail. |
| `TPSTHoldHD` | Held (suspended) bill header. |
| `TPSTHoldRC` | Held bill receipt. |
| `TPSTLastTransNo` | POS transaction |
| `TPSTLogDT` | POS transaction, detail |
| `TPSTLogIn` | Cashier login log. |
| `TPSTLogInAct` | POS transaction |
| `TPSTLogMissing` | POS transaction |
| `TPSTLogWallet` | POS transaction |
| `TPSTPriceOveride` | POS transaction |
| `TPSTSALDT` | POS transaction, detail |
| `TPSTSALHD` | POS transaction, header |
| `TPSTSalCD` | Sales card/discount detail (FTScdCardID token, FNDctNo). |
| `TPSTSalDT` | Sales detail — line items per bill. |
| `TPSTSalHD` | Sales header — one row per bill (PK FTTmnNum/FTShdTransNo/FDShdTransDate). Carries FTShdStaDoc + FTStaSentOnOff. |
| `TPSTSalPmtCpn` | POS transaction, coupon |
| `TPSTSalRC` | Sales receipt/tender — payments per bill (FTTdmCode, FTSrcCardNo token). |
| `TPSTSalVatDT` | VAT breakdown detail. |
| `TPSTSalVatHD` | VAT breakdown header. |
| `TPSTSalVatRC` | VAT breakdown receipt. |
| `TPSTSalePoint` | Member points accrued per transaction (FTSpoMemID token, FCSpoPoint). |
| `TPSTSerDownload` | POS transaction |
| `TPSTServerPOS` | POS transaction |
| `TPSTTmpVatDT` | POS transaction, detail |
| `TPSTTmpVatHD` | POS transaction, header |
| `TPSTTokenLst` | Token mapping: real value <-> SafeNet token (FTToken/FTValue/FTTblName/FTFldName). |
| `TPSTUNSENTCV` | POS transaction |
| `TPSTVarSQL` | POS transaction |
| `TPSTVoucherHD` | Voucher transaction header. |
| `TPSTVoucherRC` | Voucher receipt. |

## Center Master Tables  (136)

*ตาราง Master ที่ดูแลที่ส่วนกลางและ Replicate ลงสาขา*

| Table | Description |
| --- | --- |
| `TCNMAutoTranDep` | center master / maintenance, deposit |
| `TCNMAutoTranDepdbo` | center master / maintenance, deposit |
| `TCNMAutoTranDis` | center master / maintenance |
| `TCNMAutoTranDisdbo` | center master / maintenance |
| `TCNMAutoTranID` | center master / maintenance |
| `TCNMAutoTranIDTier` | center master / maintenance, tier, detail |
| `TCNMAutoTranIDdbo` | center master / maintenance |
| `TCNMAutoTranSKU` | center master / maintenance |
| `TCNMAutoTranSKUdbo` | center master / maintenance |
| `TCNMBanknote` | Banknote denomination master. |
| `TCNMBanknotedbo` | center master / maintenance |
| `TCNMCancelDpt` | center master / maintenance |
| `TCNMCancelDptdbo` | center master / maintenance |
| `TCNMCoopBrnMtn` | center master / maintenance, co-operate |
| `TCNMCoopBrnMtndbo` | center master / maintenance, co-operate |
| `TCNMCoopCpMtn` | center master / maintenance, co-operate |
| `TCNMCoopCpMtndbo` | center master / maintenance, co-operate |
| `TCNMCreditCard` | Credit-card type/BIN master. |
| `TCNMCreditCardDep` | center master / maintenance, deposit |
| `TCNMCreditCardOnTopAlwPOS` | center master / maintenance |
| `TCNMCreditCardRange` | center master / maintenance |
| `TCNMCreditCardSKU` | center master / maintenance |
| `TCNMCreditCardTierOnTop` | center master / maintenance, tier |
| `TCNMCreditCardTierPoint` | center master / maintenance, tier |
| `TCNMCstAtt` | center master / maintenance |
| `TCNMCstMtn` | Customer master. |
| `TCNMCstTelMtn` | center master / maintenance |
| `TCNMCstTelMtndbo` | center master / maintenance |
| `TCNMDiscountType` | Discount-type master. |
| `TCNMDoubleDiscPolicy` | center master / maintenance |
| `TCNMEVENTMTN` | center master / maintenance |
| `TCNMEmpCardDep` | center master / maintenance, employee, deposit |
| `TCNMEmpCardDepdbo` | center master / maintenance, employee, deposit |
| `TCNMEmpCardMtn` | Employee card master. |
| `TCNMEmpCardMtndbo` | center master / maintenance, employee |
| `TCNMEmpMtn` | Employee master. |
| `TCNMEmpMtnDC` | center master / maintenance, employee |
| `TCNMEventMtn` | center master / maintenance |
| `TCNMEventMtndbo` | center master / maintenance |
| `TCNMEventPrc` | center master / maintenance |
| `TCNMEventPrcdbo` | center master / maintenance |
| `TCNMForCurrency` | Foreign-currency master. |
| `TCNMForCurrencydbo` | center master / maintenance |
| `TCNMFreeGift` | Free-gift master. |
| `TCNMFreeGiftdbo` | center master / maintenance |
| `TCNMGiftVoucher` | Gift-voucher master. |
| `TCNMGiftVoucherdbo` | center master / maintenance |
| `TCNMKeyAscii` | center master / maintenance |
| `TCNMLevel` | center master / maintenance |
| `TCNMMallCardDep` | center master / maintenance, mall member, deposit |
| `TCNMMallCardDepdbo` | center master / maintenance, mall member, deposit |
| `TCNMMallCardMtn` | Mall member card master. |
| `TCNMMallCardMtndbo` | center master / maintenance, mall member |
| `TCNMMallCardRange` | center master / maintenance, mall member |
| `TCNMMallCardSKU` | Member-card SKU rules. |
| `TCNMMallCardSKUdbo` | center master / maintenance, mall member |
| `TCNMMallCardTier` | Member tier definitions. |
| `TCNMMallCpn` | center master / maintenance, mall member, coupon |
| `TCNMMemberCooMtn` | center master / maintenance |
| `TCNMMemberCoopAlw` | center master / maintenance, co-operate |
| `TCNMMemberGroup` | center master / maintenance |
| `TCNMMemberGroupdbo` | center master / maintenance |
| `TCNMMemberPrice` | center master / maintenance |
| `TCNMMemberPricePreDay` | pre-day (next operation date) staging, center master / maintenance |
| `TCNMMemberPricedbo` | center master / maintenance |
| `TCNMMemberSectAlw` | center master / maintenance |
| `TCNMMemberSectMtn` | center master / maintenance |
| `TCNMMessage` | center master / maintenance |
| `TCNMOthDisburse` | center master / maintenance |
| `TCNMOthDisbursedbo` | center master / maintenance |
| `TCNMOthReceive` | center master / maintenance |
| `TCNMOthReceivedbo` | center master / maintenance |
| `TCNMPENDPRICE` | center master / maintenance |
| `TCNMPendBPDeptBrand` | replication-pending staging, center master / maintenance, deposit |
| `TCNMPendBPSKU` | replication-pending staging, center master / maintenance |
| `TCNMPendBPTier` | replication-pending staging, center master / maintenance, tier |
| `TCNMPendBarcode` | replication-pending staging, center master / maintenance |
| `TCNMPendCoBrand` | replication-pending staging, center master / maintenance |
| `TCNMPendDisc` | replication-pending staging, center master / maintenance |
| `TCNMPendMember` | replication-pending staging, center master / maintenance |
| `TCNMPendOthReceive` | replication-pending staging, center master / maintenance |
| `TCNMPendPoint` | replication-pending staging, center master / maintenance |
| `TCNMPendPrice` | replication-pending staging, center master / maintenance |
| `TCNMPendPromo` | replication-pending staging, center master / maintenance, promotion |
| `TCNMPendPromoTime` | replication-pending staging, center master / maintenance, promotion |
| `TCNMPendSKU` | replication-pending staging, center master / maintenance |
| `TCNMPendSKUInfo` | replication-pending staging, center master / maintenance |
| `TCNMPmtCpnJN` | center master / maintenance, coupon |
| `TCNMPmtCpnJNTier` | center master / maintenance, coupon, tier |
| `TCNMPmtCpnJNdbo` | center master / maintenance, coupon |
| `TCNMPointCard` | center master / maintenance |
| `TCNMPointCardPOS` | center master / maintenance |
| `TCNMPreDayMtn` | pre-day (next operation date) staging, center master / maintenance |
| `TCNMPreDayMtndbo` | pre-day (next operation date) staging, center master / maintenance |
| `TCNMPromo` | Promotion master. |
| `TCNMPromoPreDay` | pre-day (next operation date) staging, center master / maintenance, promotion |
| `TCNMPromoPreDaydbo` | pre-day (next operation date) staging, center master / maintenance, promotion |
| `TCNMPromoTime` | Time-sale promotion master. |
| `TCNMPromoTimePreDay` | pre-day (next operation date) staging, center master / maintenance, promotion |
| `TCNMPromoTimedbo` | center master / maintenance, promotion |
| `TCNMReason` | center master / maintenance |
| `TCNMReceipt` | Receipt layout master. |
| `TCNMReceiptJN` | center master / maintenance |
| `TCNMReceiptJNdbo` | center master / maintenance |
| `TCNMSKU` | SKU / product master. |
| `TCNMSKUBarcode` | SKU barcode master. |
| `TCNMSKUBarcodePreDay` | pre-day (next operation date) staging, center master / maintenance |
| `TCNMSKUBarcodedbo` | center master / maintenance |
| `TCNMSKUPreDay` | pre-day (next operation date) staging, center master / maintenance |
| `TCNMSKUPreDaydbo` | pre-day (next operation date) staging, center master / maintenance |
| `TCNMSkuBarCon` | center master / maintenance, SKU/product |
| `TCNMSkuBarCondbo` | center master / maintenance, SKU/product |
| `TCNMSkuDepCls` | SKU department/class. |
| `TCNMSkuDepClsdbo` | center master / maintenance, SKU/product, deposit |
| `TCNMSkuDiv` | center master / maintenance, SKU/product |
| `TCNMSkuFlagMtn` | center master / maintenance, SKU/product |
| `TCNMSkuFlagMtndbo` | center master / maintenance, SKU/product |
| `TCNMSkuPoint` | SKU point-earning rules. |
| `TCNMSkuPointPreDay` | pre-day (next operation date) staging, center master / maintenance, SKU/product |
| `TCNMSkuPointdbo` | center master / maintenance, SKU/product |
| `TCNMSlipMsg` | Slip message master. |
| `TCNMStoreMtn` | Store master. |
| `TCNMStoreMtndbo` | center master / maintenance |
| `TCNMStorePara` | Store parameters (incl. token-enable FTStpCode=29). |
| `TCNMStoreParadbo` | center master / maintenance |
| `TCNMTaxMtn` | Tax-rate master. |
| `TCNMTenderMember` | center master / maintenance, tender |
| `TCNMTenderMtn` | Tender (payment type) master. |
| `TCNMTenderMtndbo` | center master / maintenance, tender |
| `TCNMTerminalMtn` | POS terminal master (health-check source). |
| `TCNMUnLockTerminal` | center master / maintenance |
| `TCNMUsr` | User master. |
| `TCNMUsrDep` | center master / maintenance, deposit |
| `TCNMUsrTle` | center master / maintenance |
| `TCNMVndCoupon` | Vendor-coupon master. |
| `TCNMVndCoupondbo` | center master / maintenance, vendor |

## System / Configuration Tables  (53)

*ตารางตั้งค่าระบบและคุมการทำงาน*

| Table | Description |
| --- | --- |
| `TSysAuto` | system / configuration |
| `TSysAutoRe` | system / configuration |
| `TSysAutoRep` | Auto-replication config (master Server->branch). |
| `TSysBrwCol` | system / configuration |
| `TSysCardType` | system / configuration |
| `TSysCardTypedbo` | system / configuration |
| `TSysChgDateTime` | Operation/sale date (FDCdtDate). |
| `TSysChgDateTimedbo` | system / configuration |
| `TSysConfig` | System config (incl. tokenizer creds, code TKToken). |
| `TSysConfigAlw` | system / configuration |
| `TSysDCConfig` | system / configuration |
| `TSysDcSchedule` | system / configuration |
| `TSysEvent` | system / configuration |
| `TSysExport` | system / configuration |
| `TSysKeyboard` | On-screen keyboard layout. |
| `TSysKeyboarddbo` | system / configuration |
| `TSysLang` | Language resource strings. |
| `TSysLevMenu` | system / configuration |
| `TSysLog` | system / configuration |
| `TSysLogPrc` | system / configuration |
| `TSysLogTableUpdate` | system / configuration |
| `TSysMeaning` | system / configuration |
| `TSysMenu` | Menu definitions. |
| `TSysMenuDC` | system / configuration |
| `TSysObjColor` | system / configuration |
| `TSysObjColordbo` | system / configuration |
| `TSysOptDateMtn` | system / configuration, maintenance |
| `TSysOptDateMtndbo` | system / configuration, maintenance |
| `TSysPortPrn` | system / configuration |
| `TSysPosConfig` | system / configuration |
| `TSysPosMenu` | system / configuration |
| `TSysRegLic` | system / configuration |
| `TSysReplication` | Replication config. |
| `TSysReport` | Report definitions. |
| `TSysReportStd` | system / configuration |
| `TSysRptCfg` | system / configuration |
| `TSysRptLang` | system / configuration |
| `TSysRptOpt` | system / configuration |
| `TSysRptSch` | system / configuration |
| `TSysSQLProc` | system / configuration |
| `TSysSync` | Sync configuration — which tables ServiceTransfer pushes (FTSscStaActive). |
| `TSysTable` | system / configuration |
| `TSysTblSync` | system / configuration |
| `TSysTerminalSta` | Terminal status (store-open flag FTTmsOpnStore). |
| `TSysTerminalStadbo` | system / configuration |
| `TSysTransType` | Transaction-type codes. |
| `TSysTransTypedbo` | system / configuration |
| `TSysUser` | System users. |
| `TSysUserMenu` | system / configuration |
| `TSysUserMenuDC` | system / configuration |
| `TSysUserMenudbo` | system / configuration |
| `TSysUserOpt` | system / configuration |
| `TSysVersion` | Schema/app version. |
