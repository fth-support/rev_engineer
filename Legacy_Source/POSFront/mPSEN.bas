Attribute VB_Name = "mPSEN"
Option Explicit

'*Ao 51-12-19 Modules นี้มีโปรแกรมที่ใช้ร่วมกัน คือ POSFront และ ManageEJ

Public Enum eEN_DbMode
    nEN_DbOnline = 1
    nEN_DbOffline = 2
End Enum

Public Enum eEN_DocMode     '*Tee 49-03-04
  nEN_Type01SaleS = 1
  nEN_Type01SaleNoVatN = 1
  nEN_Type09SaleR = 9
  nEN_Type09SaleNoVatU = 9
  nEN_Type70InComeI = 70
  nEN_Type71InComeJ = 71
  nEN_Type80ExpensesX = 80
  nEN_Type81ExpensesY = 81
End Enum

'for EJ     '*Tee 47-02-13   25 Line
Public Enum EJ_TypePicture
    EJ_0Wait = 0
    EJ_1Complete = 1    'พิมพ์กำกับอย่างย่อ
    EJ_2Cancel = 2         'พิมพ์ใบยกเลิก
    EJ_3ShwRCV = 3      'พิมพ์ใบสรุปยอดเงินที่นับได้
    EJ_4ShwSUM = 4      'พิมพ์ใบสรุปยอดขายกันกับยอดบันทึก
    EJ_5ShwFULL = 5     'พิมพ์ใบกำกับภาษีเต็มรูปอย่างย่อ
    EJ_6Delete = 6          'ลบไฟล์ข้อมูลที่ไม่ต้องการออก
    EJ_7Error = 7
    EJ_8ShwBankNote = 8     'ธนบัตร
    EJ_9Coupon = 9      'พิมพ์ใบคูปอง
    EJ_10ShwMnyIn = 10 'นำเงินเข้า
    EJ_11ShwMnyOut = 11 'นำเงินออก
End Enum
'folder
Public Enum eEN_FldManage
    nEN_Fld1Exist = 1               'check if exist
    nEN_Fld2Delete = 2           'delete folder
    nEN_Fld3Create = 3           'create folder
End Enum

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Public Enum eEN_RowAction        '*Tee 47-12-03
    nEN_Row1Del = 1
    nEN_Row2Remove = 2
End Enum

Public Enum eEN_UsrAutho        '*Tee 47-12-03
    nEN_Usr0AthAll = 0
    nEN_Usr1AthFull = 1
    nEN_Usr2AthRead = 2
    nEN_Usr3AthAdd = 3
    nEN_Usr4AthEdit = 4
    nEN_Usr5AthDelete = 5
    nEN_Usr6AthCancel = 6
    nEN_Usr7AthAppv = 7
End Enum

'common dialouge style
Public Enum eEN_OcdStyle        '*Tee 47-12-03
    nEN_Ocd1Open = 1           'show open
    nEN_Ocd2Save = 2           'show save
    nEN_Ocd3Color = 3           'show color
    nEN_Ocd4Font = 4            'show font
    nEN_Ocd5Printer = 5        'show printer
End Enum

'ชนิดของฐานข้อมูลที่ใช้ในโปรแกรมนี้ chk. sub main()
Public Enum eEN_DBMS
    nEN_Dbs1ACS = 1
    nEN_Dbs2SQL = 2
End Enum

'for pos printer
Public Enum Position
    PrintHead = 0
    PrintPrice = 1
    PrintForward = 2
    PrintBackward = 3
    PrintCut = 4
    PrintFoot = 5
End Enum

'file
Public Enum eEN_FleManage
    nEN_Fle1Exist = 1               'check if exist
    nEN_Fle2Delete = 2           'delete file
End Enum

Public Enum eEN_KeyIsBar
    nEN_Key1Bar1OrCode = 1
    nEN_Key2Bar2 = 2
    nEN_Key3Bar3 = 3
End Enum

'*Tee 48-03-24
Public Enum eEN_UsrOpt
    nEN_Usr0OptAll = 0
    nEN_Usr1OptEditPdt = 1
    nEN_Usr2OptSeePdt = 2
    nEN_Usr3OptSeeCost = 3
    nEN_Usr4OptSeeQty = 4
    nEN_Usr5OptBSgnOut = 5
    nEN_Usr6OptCheBill = 6
    nEN_Usr7OptDummy = 7
End Enum

'*TW 50-02-17
Public Enum eEN_Period
  nEN_Year = 1
  nEN_Month = 2
  nEN_Day = 3
  nEN_None = 0
End Enum

'*Em 50-10-22
Public Enum eEN_FedLang
    nEN_FedLang1 = 1
    nEN_FedLang2 = 2
    nEN_FedGrp = 3
    nEN_ColFed = 4
    nEN_ColWidth = 5
End Enum

'DB CursorLocation
Public Enum eEN_DBCursorLocation
    nEN_UseServer = 2
    nEN_UseClient = 3
End Enum

'DB CursorType
Public Enum eEN_DBCursorType
    nEN_OpenForwardOnly = 0
    nEN_OpenKeyset = 1
    nEN_OpenDynamic = 2
    nEN_OpenStatic = 3
End Enum

'DB LockType
Public Enum eEN_DBLockType
    nEN_LockReadOnly = 1
    nEN_LockPessimistic = 2
    nEN_LockOptimistic = 3
    nEN_LockBatchOptimistic = 4
End Enum

'cut caption
Public Enum eEN_CutCaption
    nEN_Cut1Msg = 1
    nEN_Cut2Tag = 2
    nEN_Cut3Const = 3
End Enum

'version
Public Enum eEN_Version
    nEN_Ver1Demo = 1
    nEN_Ver2Std = 2
    nEN_Ver3Pro = 3
    nEN_Ver4Ent = 4
End Enum

'Discount Type '*TW 51-02-18
Public Enum eEN_DiscType
    nEN_Disc01Deal = 1 ' Deal & Free Item
    nEN_Disc02PmtTime = 2 'Promotion Timesale
    nEN_Disc03Pmt = 3 'Promotion
    nEN_Disc04CoopDisc = 4 'Cooperate Discount
    nEN_Disc05LnDisc = 5 'Line Discount
    nEN_Disc06LnMDown = 6 ' Line Markdown
    nEN_Disc07AutoTran = 7 'Auto Tran Disc.
    nEN_Disc08MallCpn = 8 'Mall Coupon
    nEN_Disc09CoopCpn = 9 'Cooperate Coupon
    nEN_Disc10CoopBrand = 10 'Cooperate Brand
    nEN_Disc11MallCrd = 11 'Mall Card
    nEN_Disc12EmpCrd = 12 'Employee Card
    nEN_Disc13TranDisc = 13 'Transaction Disc.
    nEN_Disc14TranMDown = 14 ' Transaction Mkdn.
    nEN_Disc15Ontop = 15 'On top
    nEN_Disc16FreeGift = 16 'Free Gift (Special)
    nEN_Disc20Special = 20 'Free Gift (Special)
    nEN_Disc21MPrice = 21 'Member Price + Non Consignment
    nEN_Disc22MDisc = 22 'Member Price + Consignment '*KT 52-09-07 Member Price + Consignment
    nEN_Disc23TranMDown1 = 23 'Transaction Mkdn.1
    nEN_Disc24TranMDown2 = 24 'Transaction Mkdn.2
    nEN_Disc25TranMDown3 = 25 'Transaction Mkdn.3
    nEN_Disc26TranMDown4 = 26 'Transaction Mkdn.4
    nEN_Disc27TranMDown5 = 27 'Transaction Mkdn.5
End Enum

'Line Item Discount Type '*TW 51-02-18
Public Enum eEN_LineDiscType
    nEN_LnDis0NoDis = 0 'No Disc
    nEN_LnDis1Per = 1 'Percent Discount
    nEN_LnDis2Amt = 2 'Amount Discount
    nEN_LnDis3NewPrice = 3 'New Price
    nEN_LnDis4Free = 4 'Free
End Enum

Public Enum eEN_TransMode
    tEN_TrnOpen = 1
    tEN_TrnStartBank = 2
    tEN_TrnSale = 3
    tEN_TrnReturn = 4
    tEN_TrnIPV = 5
    tEN_TrnDeposit = 6
    tEN_TrnRedeem = 7
    tEN_TrnClose = 8
    tEN_TrnPrintClose = 9
    tEN_TrnVIPSale = 10
    tEN_TrnTaxExceptSale = 11
    tEN_TrnPickUp = 12
    tEN_TrnDutyFreeSale = 13
    tEN_TrnIPVReturn = 14
    tEN_TrnVoucher = 15
    tEN_TrnDepositCancel = 16
    tEN_TrnIPVRedeem = 17
    tEN_TrnNoSale = 18
    tEN_TrnAddDeposit = 19
    tEN_TrnAddDepositCancel = 20
    tEN_TrnCancelVoucher = 21
    tEN_TrnOtherReceive = 22
    tEN_TrnOtherDisbursement = 23
    tEN_TrnDepositMoney = 24
    tEN_TrnTraining = 25
    tEN_TrnIPVVIPService = 26
    tEN_TrnIPVTaxExceptSale = 27
    tEN_TrnIPVDutyFreeSale = 28
    tEN_TrnSuspended = 29
    tEN_TrnCreateVIP = 30
    tEN_TrnTotalReport = 31
    tEN_TrnIPVOtherReceive = 32
    tEN_TrnIPVOtherDisbursement = 33
    tEN_TrnCopyTrans = 34
    tEN_TrnCopyFromDev = 35
    tEN_TrnCopyMaster = 36
    tEN_TrnOverShortReport = 37
    
    tEN_TrnSaleTaxInc = 38
    tEN_TrnReturnTaxInc = 39
    tEN_TrnTotalReportTaxInc = 40
    tEN_TrnOverShortReportTaxInc = 41
    tEN_TrnClosingReportTaxInc = 42
    tEN_TrnOpenTaxInc = 48 '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
    
    tEN_TrnSaleManualTax = 43
    tEN_TrnReturnManualTax = 44
    tEN_TrnTotalReportManualTax = 45
    tEN_TrnOverShortReportManualTax = 46
    tEN_TrnClosingReportManualTax = 47
    tEN_TrnOpenManualTax = 49 '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
   
End Enum

Public Enum eEN_NameAbbMode
    tEN_AbbOpen = 1
    tEN_AbbStartBank = 2
    tEN_AbbSale = 3
    tEN_AbbDeposit = 4
    tEN_AbbRedeem = 5
    tEN_AbbDepositCancel = 6
    tEN_AbbAddDeposit = 7
    tEN_AbbAddDepositCancel = 8
    tEN_AbbCreateVIP = 9
    tEN_AbbVIPSale = 10
    tEN_AbbVoucher = 11
    tEN_AbbCancelVoucher = 12
    tEN_AbbPickUp = 13
    tEN_AbbTaxExceptSale = 14
    tEN_AbbIPV = 15
    tEN_AbbReturn = 16
    tEN_AbbRefReturn = 17
    tEN_AbbOtherReceive = 18
    tEN_AbbOtherDisbursement = 19
    tEN_AbbDepositMoney = 20
    tEN_AbbDutyFreeSale = 21
    tEN_AbbNoSale = 22
    tEN_AbbTotalReport = 23
    tEN_AbbTrainingOn = 24
    tEN_AbbTrainingOff = 25
    tEN_AbbClose = 26
    tEN_AbbPrintClose = 27
    tEN_AbbRePrintClose = 28
    tEN_AbbIPVReturn = 29
    tEN_AbbIPVRedeem = 30
    tEN_AbbIPVVIPService = 31
    tEN_AbbIPVTaxExceptSale = 32
    tEN_AbbIPVDutyFreeSale = 33
    tEN_AbbSuspended = 34
    tEN_AbbIPVOtherReceive = 35
    tEN_AbbIPVOtherDisbursement = 36
    tEN_AbbCopyTrans = 37
    tEN_AbbCopyFromDev = 38
    tEN_AbbCopyMaster = 39
    tEN_AbbOverShort = 40
    
    tEN_AbbSaleTaxInc = 41
    tEN_AbbReturnTaxInc = 42
    tEN_AbbTotalReportTaxInc = 43
    tEN_AbbOverShortTaxInc = 44
    tEN_AbbClosingReportTaxInc = 45
    tEN_AbbOpenTaxInc = 51 '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
    
    tEN_AbbSaleManualTax = 46
    tEN_AbbReturnManualTax = 47
    tEN_AbbTotalReportManualTax = 48
    tEN_AbbOverShortManualTax = 49
    tEN_AbbClosingReportManualTax = 50
    tEN_AbbOpenManualTax = 52 '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
    
End Enum

Enum eEN_TypeLogDT
    eCancel = 0 'Cancel
    eQtyChg = 1 'Qty Change
    eNagativeScn = 2 'Nagative Scan
End Enum

'Const tW_ConRnd = "|||Don't Round"

Enum eEN_TypeRnd 'ประเภทการ Rouding
    eRndUp = 0 'Round Up
    eRndDn = 1 'Round Down
    eRndMeaning = 2 'Meaning
    eRndDon = 3 'Don't Round
End Enum

'OPOS
'------------------------------------
Enum eTrans
    PTR_TP_TRANSACTION = 11&
    PTR_TP_NORMAL = 12&
End Enum

Enum eStation
    PTR_S_JOURNAL = 1&
    PTR_S_RECEIPT = 2&
    PTR_S_SLIP = 4&
End Enum

Enum eSymbology
    PTR_BCS_UPCA = 101&         ' // Digits
    PTR_BCS_UPCE = 102&         ' // Digits
    PTR_BCS_JAN8 = 103&        ' // = EAN 8
    PTR_BCS_EAN8 = 103&        ' // = JAN 8 (added in 1.2)
    PTR_BCS_JAN13 = 104&       ' // = EAN 13
    PTR_BCS_EAN13 = 104&       ' // = JAN 13 (added in 1.2)
    PTR_BCS_TF = 105&          ' // (Discrete 2 of 5) Digits
    PTR_BCS_ITF = 106&         ' // (Interleaved 2 of 5) Digits
    PTR_BCS_Codabar = 107&     ' // Digits, -, $, :, /, ., +;
                                                     ' //   4 start/stop characters
                                                     ' //   (a, b, c, d)
    PTR_BCS_Code39 = 108&      ' // Alpha, Digits, Space, -, .,
                                                     ' //   $, /, +, %; start/stop (*)
                                                     ' // Also has Full ASCII feature
    PTR_BCS_Code93 = 109&       ' // Same characters as Code 39
    PTR_BCS_Code128 = 110&     ' // 128 data characters
    
    PTR_BCS_UPCA_S = 111&         ' // UPC-A with supplemental
                                                       ' //   barcode
    PTR_BCS_UPCE_S = 112&         ' // UPC-E with supplemental
                                                       ' //   barcode
    PTR_BCS_UPCD1 = 113&         ' // UPC-D1
    PTR_BCS_UPCD2 = 114&          ' // UPC-D2
    PTR_BCS_UPCD3 = 115&         ' // UPC-D3
    PTR_BCS_UPCD4 = 116&         ' // UPC-D4
    PTR_BCS_UPCD5 = 117&         ' // UPC-D5
    PTR_BCS_EAN8_S = 118&         ' // EAN 8 with supplemental
                                                       ' //   barcode
    PTR_BCS_EAN13_S = 119&        ' // EAN 13 with supplemental
                                                       ' //   barcode
    PTR_BCS_EAN128 = 120&         ' // EAN 128
    PTR_BCS_OCRA = 121&          ' // OCR "A"
    PTR_BCS_OCRB = 122&          ' // OCR "B"
    
    ' //  - One dimensional symbologies (added in 1.8)
    PTR_BCS_RSS14 = 131&        ' // Reduced Space Symbology - 14 digit GTIN
    PTR_BCS_RSS_EXPANDED = 132&   ' // RSS - 14 digit GTIN plus additional fields
    
    ' //  - Composite Symbologies (added in 1.8)
    PTR_BCS_CCA = 151&            ' // Composite Component A.
    PTR_BCS_CCB = 152&            ' // Composite Component B.
    PTR_BCS_CCC = 153&            ' // Composite Component C.
    
    ' // Two dimensional symbologies
    PTR_BCS_PDF417 = 201&
    PTR_BCS_MAXICODE = 202&
    
    ' // Special cases
    PTR_BCS_OTHER = 501&        ' // Start of Scanner-Specific bar
                                                      ' //   code symbologies
    PTR_BCS_UNKNOWN = 0&         ' // Cannot determine the barcode
End Enum

Enum eAlignBarCode
    DISP_BM_LEFT = -1&
    DISP_BM_CENTER = -2&
    DISP_BM_RIGHT = -3&
End Enum

Enum eAlign
    eLeft = 0
    eCenter = 1
    eRight = 2
End Enum

Enum eAlignColum
    eLeftColum = 0
    eRightColum = 2
End Enum

Enum eDISP
    DISP_DT_NORMAL = 0&
    DISP_DT_BLINK = 1&
    DISP_DT_REVERSE = 2&
    DISP_DT_BLINK_REVERSE = 3&
End Enum

Enum eAlignTextPosition
    PTR_BC_TEXT_NONE = -11&
    PTR_BC_TEXT_ABOVE = -12&
    PTR_BC_TEXT_BELOW = -13&
End Enum

Enum eMapMode
    PCRW_MM_DOTS = 1&
    PCRW_MM_TWIPS = 2&
    PCRW_MM_ENGLISH = 3&
    PCRW_MM_METRIC = 4&
End Enum

Enum eLacationLogo
    PTR_L_TOP = 1&
    PTR_L_BOTTOM = 2&
End Enum

'--------------------------------------

'*TW 51-05-09 For Element of Array {Column}
Public Enum eEN_DealDT
    nEN_DealDT00SeqNo = 0 ' DT SeqNo
    nEN_DealDT01QtyAll = 1 ' DT Qty All
    nEN_DealDT02QtyUsed = 2 ' DT Qty Used
    nEN_DealDT03QtyLeft = 3 ' DT Qty Left
    nEN_Deal04CoopDisc = 4 'Cooperate Discount
End Enum

Public Enum eEN_DealByDeal
    nEN_DealByDeal00Key = 0 ' DealNo + # + EventID
    nEN_DealByDeal01DealNo = 1 ' DealNo
    nEN_DealByDeal02EventID = 2 ' EventID
    nEN_DealByDeal03DealType = 3 ' DealType
    nEN_DealByDeal04DTSeqBuy = 4 ' DT SeqNo Buy
    nEN_DealByDeal05DTQtyBuy = 5 ' DT Qty Buy
    nEN_DealByDeal06DTSeqGet = 6 ' DT SeqNo Get
    nEN_DealByDeal07DTQtyGet = 7 ' DT Qty Get
End Enum

Public Enum eEN_DealSum
    nEN_DealSum00Key = 0 ' DealNo + # + EventID
    nEN_DealSum01SeqDeal = 1 ' DealNo + # + EventID
    nEN_DealSum02Sum = 2 ' DealNo + # + EventID
End Enum
'----------------------------------------------------------
'*TW 51-06-20 DealType Recordset
Public Enum eEN_DealTypeRec
    nEN_Deal01Single = 1 ' DealNo + # + EventID
    nEN_Deal02Multiple = 2 ' DealNo + # + EventID
End Enum

Public Enum eEN_VerifyDiscType
  nEN_Type1ByItem = 1
  nEN_Type2ByDcn = 2
  nEN_Type3None = 3
End Enum

Public Enum eEN_PointType
  nEN_Type0DblLowBP = 0 '*Tao 60-04-20 Double Low BP (ใช้ 10 ไม่ได้ เพราะ FTSpoType = Varchar(1))
  nEN_Type1SKUPoint = 1
  nEN_Type2CooPoint = 2
  nEN_Type3CooAddPoint = 3
  nEN_Type4SecPoint = 4
  nEN_Type5SecAddPoint = 5
  nEN_Type6MemPoint = 6
  nEN_Type7CrdPoint = 7
  nEN_Type8AllPoint = 8
  nEN_Type9LowBP = 9 '*Tao 60-04-05 Low BP
End Enum

Public Const PTR_BM_ASIS = -11

Public Const PTR_BM_LEFT = -1
Public Const PTR_BM_CENTER = -2
Public Const PTR_BM_RIGHT = -3

Public Type eEN_ReIssue
        tEN_ReIssueExp As String
        tEN_ReIssueNo As String
        cEN_ReIssueAmt  As Double
        tEN_ReIssueTel As String '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
End Type

Public Enum eEN_NETSType      '*Eaw 56-04-29 NETS Port
    eEN_NETSTender = 0
    eEN_BCA = 1
    eEN_GiftCard = 2
    eEN_NetsQR = 3 '*Tao 62-07-19 NETS QR
    eEN_CashCard = 4
    eEN_FlashPay = 5
    eEN_BeforeGC = 6   '*Eaw 56-06-15
    eEN_AfterGC = 7   '*Eaw 56-06-15
    eEN_PrePaid = 8 '*Tao 62-07-19 เปลี่ยนจาก 3 เป็น 8
End Enum

