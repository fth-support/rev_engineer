Attribute VB_Name = "mCS"
Option Explicit

'for compile version
Public Enum CS_Compile
    nCS_ComStd = 2
    nCS_ComPro = 3
    nCS_ComEnt = 4
End Enum

'API
Public Const nCS_MaxFileLen = 256
Public Const nCS_RptPgZoom = 120

'authorize
Public Const tCS_CNPwdNoUsr = "ADASOFT"
Public Const tCS_CNPwdAdaUsrC = "ADAADMIN"
Public Const tCS_CNPwdAdaUsrP = "SUPPORT"
Public Const tCS_PwdEncHD = "NwNk@b3g"
Public Const tCS_PwdEncHDKey = "adaXSOFT"

'module for register
Public Const tCS_ModulePOS = "01"   'module point of sale
Public Const tCS_ModuleWHS = "02"   'module whole sale


Public Const tCS_DefaultExtMDB = "MDB"
Public Const tCS_FileName = "Company1"
Public Const tCS_Filter = "Microsoft Access(*.MDB)|*.MDB|All File(*.*)|*.*"
Public Const tCS_FilterBitmap = "Bitmap Files(*.BMP)|*.BMP"
Public Const tCS_FilterGif = "Gif Image Files(*.GIF)|*.GIF"
Public Const tCS_FilterMeta = "Meta Files(*.WMF)|*.WMF"
Public Const tCS_FilterIcons = "Icons Files(*.ICO)|*.ICO"
Public Const tCS_FilterJpeg = "Jpeg Files(*.JPG)|*.JPG"
Public Const tCS_FilterText = "Text Files(*.TXT)|*.TXT"
Public Const tCS_FilterAll = "All Files(*.*)|*.*"
Public Const tCS_FilterRpt = "Report Files(*.Rpt)|*.Rpt"

'set button msgbox style
Public Const nCS_Error = vbOKOnly + vbCritical
Public Const nCS_Warning = vbOKOnly + vbExclamation
Public Const nCS_Information = vbOKOnly + vbInformation
Public Const nCS_OkCancel = vbOKCancel + vbQuestion
Public Const nCS_OkCancelDefaultButton1 = vbOKCancel + vbQuestion + vbDefaultButton1

'set button msgbox style
Public Const tCS_CNMsgErr = vbOKOnly + vbCritical
Public Const tCS_CNMsgWarn = vbOKOnly + vbExclamation
Public Const tCS_CNMsgInfo = vbOKOnly + vbInformation
Public Const tCS_CNMsgYesNo = vbYesNo + vbQuestion
Public Const tCS_CNMsgOkCancel = vbOKCancel + vbQuestion
Public Const tCS_CNMsgYesNoCancel = vbYesNoCancel + vbQuestion
Public Const tCS_CNMsgYesNoDefNo = vbYesNo + vbQuestion + vbDefaultButton2
Public Const tCS_CNMsgYesNoDefYes = vbYesNo + vbQuestion + vbDefaultButton1

'set for form
Public Const iCS_Left2 = 2
Public Const iCS_Top4 = 4

Public Const bCS_ShowErr = True
Public Const bCS_UnShowErr = False

'for tranfer product
Public Const tCS_ColNo = "ลำดับ;Seq"
Public Const tCS_ColChk = "เช็ค;Chk"
Public Const tCS_ColCode = "รหัสสินค้า;Product Code"
Public Const tCS_ColBarCode = "รหัสบาร์โค้ด;Barcode"
Public Const tCS_ColProductName = "ชื่อสินค้า;Item Name"
Public Const tCS_ColUnit = "หน่วย;Unit"
Public Const tCS_ColFactor = "อัตราส่วน;Factor"
Public Const tCS_ColCost = "ราคาต้นทุน;Cost"
Public Const tCS_ColQty = "จำนวน;Qty"
Public Const tCS_ColSubTotal = "รวมราคา;SubTotal"
Public Const tCS_ColRegPrice = "ราคาปกติ;Reg.Price"
Public Const tCS_ColSalePrice = "ราคาขาย;Sale.Price"
Public Const tCS_ColCoupon = "หมายเลขคูปอง;Coupon No."
Public Const tCS_ColSaleman = "พนักงานขาย;Sale Man"

'for lock color
Public Const tCS_ItemQty = " หน่วย; Units"
Public Const tCS_HoldBill = "!!! พักบิล !!!;!!! Hold Bill !!!"
Public Const tCS_BillCpn = "บิลนี้คุณได้รับคูปอง;This bill you get coupon"   '*TW 50-02-12 จำนวนคูปองทีได้โปรโมชั่น
Public Const tCS_Coupon = "ใบ;Coupon(s)" '*TW 50-02-12 จำนวนคูปองทีได้โปรโมชั่น

Public Const tCS_GetCoupon = "คุณได้รับคูปองส่วนลด;You get discount coupon" '*TW 50-02-12
'for Brows Bill no.
 '*NR 49-03-03 SFM  for new  wCNBrwPdt
Public Const tCS_CNHBThaPdtCode = "รหัสสินค้า;ชื่อสินค้าที่;กลุ่มสินค้า;รหัสผู้จำหน่าย;ราคาซื้อ;จำนวนคงเหลือ"
Public Const tCS_CNHBEngPdtCode = "Code;First Name;Group;Supplier Code;Buying Price;Quantity"
Public Const tCS_CNHBFldPdtCode = "TCNMPdt.FTPdtCode,FTPdtName,FTPgpChain,FTSplCode,FCPdtQtyWhs"

 '*NR 49-03-03 SFM  for new  wCNBrwPdt
Public Const tCS_CNHBThaPdtBar = "ชื่อสินค้าที่;รหัสสินค้า;กลุ่มสินค้า;รหัสผู้จำหน่าย;ราคาซื้อ;จำนวนคงเหลือ"
Public Const tCS_CNHBEngPdtBar = "First Name;Code;Group;Supplier Code;Buying Price;Quantity"
Public Const tCS_CNHBFldPdtBar = "FTPdtName,FTPdtCode,FTPgpChain,FTSplCode,FCPdtQtyWhs"

 '*NR 49-03-03 SFM  for new  wCNBrwPdt
Public Const tCS_CNHBFixThaPdtCode = "รหัสสินค้า;ชื่อสินค้าที่;"
Public Const tCS_CNHBFixEngPdtCode = "Code;First Name;"
Public Const tCS_CNHBFixFldPdtCode = "TCNMPdt.FTPdtCode,FTPdtName,"

 '*NR 49-03-03 SFM  for new  wCNBrwPdt
Public Const tCS_CNHBFixThaPdtBar = "ชื่อสินค้าที่;รหัสสินค้า;"
Public Const tCS_CNHBFixEngPdtBar = "First Name;Code;"
Public Const tCS_CNHBFixFldPdtBar = "FTPdtName,FTPdtCode,"

Public Const tCS_HBThaHoldBill = "เลขที่พักบิล;เลขที่เอกสาร;วันที่เอกสาร;เวลาที่เกิดเอกสาร;รหัสแคชเชียร์;รหัสเครื่อง;ข้อมูลอ้างอิง;เลขที่เข้าระบบ"
Public Const tCS_HBEngHoldBill = "Hold Bill No.;Bill No.;Doc Date;Doc Time;User Code;Pos Code;Refference;Sign In No."
Public Const tCS_CNLastFedIns = "FDDateUpd,FTTimeUpd,FTWhoUpd,FDDateIns,FTTimeIns,FTWhoIns)"
'for color
Public Const nCS_ColorRGrid = &HFFFF80          'light blue

'*Tee 49-06-09
Public Const nCS_CNTimeOutLevel1 = 90
Public Const nCS_CNTimeOutLevel2 = 120
Public Const nCS_CNTimeOutLevel3 = 180
'Public Const nCS_CNTimeOutLevel4 = 300


Public Const CS_Theme_Blue = "&H00FCD19D&"
Public Const CS_Theme_Green = "&H009BFC9B&"
Public Const CS_Theme_Marine = "&H00C9B49D&"
Public Const CS_Theme_Orange = "&H0099E0FB&"

'For Registration number
Public tCS_RegModule As String
Public tCS_RegCompanyLen As String
Public tCS_RegCompanySum As String
Public tCS_RegDateExp As String
Public tCS_RegSerial As String
Public tCS_RegLicenseBack As String
Public tCS_RegLicenseBackChk As String
Public tCS_RegLicenseFront As String
Public tCS_RegLicenseFrontChk As String


'*Em 50-10-22
Public Const tCS_CNLangPri = "FTUsrLangPri"
Public Const tCS_CNLangSec = "FTUsrLangSec"


'*NEW
Public Const tCS_CNBlank = ""
Public Const nCS_DisplayCol = 20
Public Const tCS_DisplayCol = "BAL"
'*AO51-03-21
Public Const tCS_FormatDateM = "dd/MM/yyyy"
Public Const tCS_FormatDateSec = "yyyy/MM/dd"

Public Const cCS_CNMaxReceive = 99999999.99
'Public Const cCS_CNMaxQty = 999999  '*BG 57-04-18
Public Const cCS_CNMaxQty = 9999

'OPOS
Public Const OPOS_SUCCESS As Long = 0&
Public Const OPOS_E_CLOSED As Long = 101&
Public Const OPOS_E_CLAIMED As Long = 102&
Public Const OPOS_E_NOTCLAIMED As Long = 103&
Public Const OPOS_E_NOSERVICE As Long = 104&
Public Const OPOS_E_DISABLED As Long = 105&
Public Const OPOS_E_ILLEGAL As Long = 106&
Public Const OPOS_E_NOHARDWARE As Long = 107&
Public Const OPOS_E_OFFLINE As Long = 108&
Public Const OPOS_E_NOEXIST As Long = 109&
Public Const OPOS_E_EXISTS As Long = 110&
Public Const OPOS_E_FAILURE As Long = 111&
Public Const OPOS_E_TIMEOUT As Long = 112&
Public Const OPOS_E_BUSY As Long = 113&
Public Const OPOS_E_EXTENDED As Long = 114&

Public Const OPOS_EPTR_COVER_OPEN As Long = 201&           ' // (Several)
Public Const OPOS_EPTR_JRN_EMPTY As Long = 202&            ' // (Several)
Public Const OPOS_EPTR_REC_EMPTY As Long = 203&            ' // (Several)
Public Const OPOS_EPTR_SLP_EMPTY As Long = 204&            ' // (Several)
Public Const OPOS_EPTR_SLP_FORM As Long = 205&             ' // EndRemoval
Public Const OPOS_EPTR_TOOBIG As Long = 206&               ' // PrintBitmap
Public Const OPOS_EPTR_BADFORMAT As Long = 207&            ' // PrintBitmap
Public Const OPOS_EPTR_JRN_CARTRIDGE_REMOVED As Long = 208&     ' // (Several) (added in 1.5)
Public Const OPOS_EPTR_JRN_CARTRIDGE_EMPTY As Long = 209&       ' // (Several) (added in 1.5)
Public Const OPOS_EPTR_JRN_HEAD_CLEANING As Long = 210&         ' // (Several) (added in 1.5)
Public Const OPOS_EPTR_REC_CARTRIDGE_REMOVED As Long = 211&     ' // (Several) (added in 1.5)
Public Const OPOS_EPTR_REC_CARTRIDGE_EMPTY As Long = 212&       ' // (Several) (added in 1.5)
Public Const OPOS_EPTR_REC_HEAD_CLEANING As Long = 213&         ' // (Several) (added in 1.5)
Public Const OPOS_EPTR_SLP_CARTRIDGE_REMOVED As Long = 214&     ' // (Several) (added in 1.5)
Public Const OPOS_EPTR_SLP_CARTRIDGE_EMPTY As Long = 215&       ' // (Several) (added in 1.5)
Public Const OPOS_EPTR_SLP_HEAD_CLEANING As Long = 216&         ' // (Several) (added in 1.5)

'filter file
Public Const tCS_CNFilterALL = "All Files(*.*)|*.*"
Public Const tCS_CNFilterMDB = "Microsoft Access(*.MDB)|*.MDB"
Public Const tCS_CNFilterBMP = "Bitmap Files(*.BMP)|*.BMP"
Public Const tCS_CNFilterGIF = "Gif Files(*.GIF)|*.GIF"
Public Const tCS_CNFilterWMF = "Meta Files(*.WMF)|*.WMF"
Public Const tCS_CNFilterJPG = "Jpeg Files(*.JPG)|*.JPG"
Public Const tCS_CNFilterICO = "Icons Files(*.ICO)|*.ICO"
Public Const tCS_CNFilterPNG = "Png Files(*.PNG)|*.PNG"
Public Const tCS_CNFilterTXT = "Text Files(*.TXT)|*.TXT"
Public Const tCS_CNFilterRPT = "Report Files(*.RPT)|*.RPT"

Public Const cCS_CNMaxSale = 99999999.99

'Not Allow Ontop
Public Const tCS_CNNotAllowTenderMall = "01#02#04#05#08#09#12#14#15#16#17#18#20#21#22#23#24#26#27#28#29#31#32#33"

Public Const cCS_Top800x600Vendor = 2400
Public Const cCS_Top680x480Vendor = 2100
Public Const cCS_Left800x600Vendor = 1000
Public Const cCS_Left680x480Vendor = 1000

'*Ao 2010/11/04
Public Const tCS_ColAmount = "มูลค่า;Amount "
Public Const tCS_ColTAmount = "รวมมูลค่า;Total Amount"
'--------------------------------------------------------------------
'*KT 51-12-08 Pre Day
'Public Const tCS_CNTblMTSKU = "TCNMSku"
'Public Const tCS_CNTblMTPromo = "TCNMPromo"
'Public Const tCS_CNTblMTPromoT = "TCNMPromoTime"
'Public Const tCS_CNTblMTPromoT = "TPSTDealAct"
'-------------------------------------------------------------------

'*KT 55-07-06 Result check status printer
Public Const PTR_SUE_COVER_OPEN As Long = 11&
Public Const PTR_SUE_COVER_OK As Long = 12&
Public Const PTR_SUE_JRN_COVER_OPEN As Long = 60&            ' // (added in 1.8)
Public Const PTR_SUE_JRN_COVER_OK As Long = 61&              ' // (added in 1.8)
Public Const PTR_SUE_REC_COVER_OPEN As Long = 62&            ' // (added in 1.8)
Public Const PTR_SUE_REC_COVER_OK As Long = 63&              ' // (added in 1.8)
Public Const PTR_SUE_SLP_COVER_OPEN As Long = 64&            ' // (added in 1.8)
Public Const PTR_SUE_SLP_COVER_OK As Long = 65&              ' // (added in 1.8)

Public Const PTR_SUE_JRN_EMPTY As Long = 21&
Public Const PTR_SUE_JRN_NEAREMPTY As Long = 22&
Public Const PTR_SUE_JRN_PAPEROK As Long = 23&

Public Const PTR_SUE_REC_EMPTY As Long = 24&
Public Const PTR_SUE_REC_NEAREMPTY As Long = 25&
Public Const PTR_SUE_REC_PAPEROK As Long = 26&

Public Const PTR_SUE_SLP_EMPTY As Long = 27&
Public Const PTR_SUE_SLP_NEAREMPTY As Long = 28&
Public Const PTR_SUE_SLP_PAPEROK As Long = 29&

Public Const PTR_SUE_JRN_CARTRIDGE_EMPTY As Long = 41&       ' // (added in 1.5)
Public Const PTR_SUE_JRN_CARTRIDGE_NEAREMPTY As Long = 42&      ' // (added in 1.5)
Public Const PTR_SUE_JRN_HEAD_CLEANING As Long = 43&            ' // (added in 1.5)
Public Const PTR_SUE_JRN_CARTRIDGE_OK As Long = 44&             ' // (added in 1.5)

Public Const PTR_SUE_REC_CARTRIDGE_EMPTY As Long = 45&       ' // (added in 1.5)
Public Const PTR_SUE_REC_CARTRIDGE_NEAREMPTY As Long = 46&      ' // (added in 1.5)
Public Const PTR_SUE_REC_HEAD_CLEANING As Long = 47&            ' // (added in 1.5)
Public Const PTR_SUE_REC_CARTRIDGE_OK As Long = 48&             ' // (added in 1.5)

Public Const PTR_SUE_SLP_CARTRIDGE_EMPTY As Long = 49&       ' // (added in 1.5)
Public Const PTR_SUE_SLP_CARTRIDGE_NEAREMPTY As Long = 50&      ' // (added in 1.5)
Public Const PTR_SUE_SLP_HEAD_CLEANING As Long = 51&            ' // (added in 1.5)
Public Const PTR_SUE_SLP_CARTRIDGE_OK As Long = 52&             ' // (added in 1.5)

Public Const PTR_SUE_IDLE As Long = 1001&

Public Const tSNR_TOUCH_SCANNER = "Fujitsu Touch Scanner" '*Bump 56-02-19   scanner ของ TAKASHIMAYA ที่เป็น Touch Scanner
'* Bump 56-02-19 เอาไว้เช็ค oScanner.DeviceName = tSNR_TOUCH_SCANNER
'* Bump 56-05-03 เพิ่ม จำนวน Max Lin เพื่อแก้ปัญหา Return Trx. CommSheet TK-ISS3000-078 00.00.06 FSBT NG.xls
'Public Const nCS_MaxLineSlipSale = 21          '*Ao 2013/07/25
Public Const nCS_MaxLineSlipSale = 22
'Public Const nCS_MaxLineSlipReturn = 24 '*Bump 56-04-07 เปลี่ยนจาก 23  เป็น 24 CommSheet TK-ISS3000-078, CommSheet TK-ISS3000-167-IT2  สินค้า 5 รายการ(ไม่มีส่วนลด Line Item)พิมพ์ พอดีในหน้าเดียว
'Public Const nCS_MaxLineSlipLayaway = 34
Public Const nCS_MaxLineSlipReturn = 22 '*BG 56-06-04  TK-ISS3000-078 ,TK -ISS3000 - 216 - ST,TK-ISS3000-181-ST
Public Const nCS_MaxLineSlipLayaway = 31 '*BG 56-06-04 TK-ISS3000-078 ,TK -ISS3000 - 216 - ST,TK-ISS3000-181-ST

'*****************************
Public Const OPOS_CH_INTERNAL As Long = 1&
Public Const OPOS_CH_EXTERNAL As Long = 2&
Public Const OPOS_CH_INTERACTIVE As Long = 3&
Public Const tCS_AutoDisableUSB As String = "Auto-Disable-USB"
