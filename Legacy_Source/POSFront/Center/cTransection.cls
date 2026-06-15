VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTransection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตังต้น 3.0.19

Public nW_StateCon As Long, nW_StatePrn As Long, nW_StateDoc As Long, nW_StateSuccess As Long
'nW_StateSuccess : 0 : ไม่มีข้อมูล HD,1 : มีข้อมูล HD '*KT 53-03-02 PH3-ML-SF-051  No EJ, Dat File 1 kb, No data on Database but it's Sale Transaction
Private ocnAll As ADODB.Connection
Private oPrn As New cPrinter

Private oW_Font(1 To 3) As New StdFont  'font style

Private aW_DctNo() As String
Private aW_DctNo2() As String
Private aW_CooCpnID() As String
Private aW_CooCpnPer() As String

Private tW_TblHD As String
Private tW_TblDT As String
Private tW_TblRC As String
Private tW_TblCD As String

Private tW_TblDpsHD As String
Private tW_TblDpsDT As String
Private tW_TblDpsRC As String
Private tW_TblDpsCD As String
Private tW_WhereCD As String  '*BG 57-01-02

Private Type eEN_Vocabulary
        nEN_Seq  As Long
        tEN_SpmID As String
        tEN_Name As String
End Type
Private aW_VocabularyPrn() As eEN_Vocabulary

Private Type eEN_TransPrimary
        dEN_TransDate  As Date
        tEN_TmnNum As String
        tEN_TransNo As String
End Type

Private Type eEN_XPos
        cEN_Left  As Double
        cEN_HalfLeft As Double
        cEN_Center As Double
        cEN_HalfRight As Double
        cEN_Right As Double
        cEN_HalfColumn As Double
End Type

Private Type eEN_Deposit
        tEN_DepositNo As String
        cEN_PevPayment  As Double
        cEN_Deposit As Double
        cEN_PrevBalDue As Double
        cEN_BalDue As Double
        cEN_BalDueDate As Date
        eEN_Parrent As eEN_TransPrimary
        tEN_PhoneNum As String '*Neung 55-07-23 Layaway TK
        cEN_RndAmt As Double '*Eaw 56-07-20 ใช้เก็บ Rounding ของตาราง TPSTDepositHD
        cEN_Reprint As Double
End Type

Private Enum eEN_TransMode
    nEN_TrnOpen = 1
    nEN_TrnStartBank = 2
    nEN_TrnSale = 3
    nEN_TrnReturn = 4
    nEN_TrnIPV = 5
    nEN_TrnDeposit = 6
    nEN_TrnRedeem = 7
    nEN_TrnClose = 8
    nEN_TrnPrintClose = 9
    nEN_TrnVIPSale = 10
    nEN_TrnTaxExceptSale = 11
    nEN_TrnPickUp = 12
    nEN_TrnDutyFreeSale = 13
    nEN_TrnIPVReturn = 14
    nEN_TrnVoucher = 15
    nEN_TrnDepositCancel = 16
    nEN_TrnIPVRedeem = 17
    nEN_TrnNoSale = 18
    nEN_TrnAddDeposit = 19
    nEN_TrnAddDepositCancel = 20
    nEN_TrnCancelVoucher = 21
    nEN_TrnOtherReceive = 22
    nEN_TrnOtherDisbursement = 23
    nEN_TrnDepositMoney = 24
    nEN_TrnTraining = 25
    nEN_TrnIPVVIPService = 26
    nEN_TrnIPVTaxExceptSale = 27
    nEN_TrnIPVDutyFreeSale = 28
    nEN_TrnSuspended = 29
    nEN_TrnCreateVIP = 30
    nEN_TrnTotalReport = 31
    nEN_TrnIPVOtherReceive = 32
    nEN_TrnIPVOtherDisbursement = 33
    nEN_TrnCopyTrans = 34
    nEN_TrnCopyFromDev = 35
    nEN_TrnCopyMaster = 36
    nEN_TrnOverShortReport = 37
    nEN_TrnSaleTaxInc = 38
    nEN_TrnReturnTaxInc = 39
    nEN_TrnTotalReportTaxInc = 40
    nEN_TrnOverShortReportTaxInc = 41
    nEN_TrnClosingReportTaxInc = 42
    nEN_TrnSaleManualTax = 43
    nEN_TrnReturnManualTax = 44
    nEN_TrnTotalReportManualTax = 45
    nEN_TrnOverShortReportManualTax = 46
    nEN_TrnClosingReportManualTax = 47
     nEN_TrnReturn7 = 48 'Amorn 65-08-02
End Enum

''Discount Type '*TW 51-02-18
'Private Enum eEN_DiscType
'    nEN_Disc01Deal = 1 ' Deal & Free Item
'    nEN_Disc02PmtTime = 2 'Promotion Timesale
'    nEN_Disc03Pmt = 3 'Promotion
'    nEN_Disc04CoopDisc = 4 'Cooperate Discount
'    nEN_Disc05LnDisc = 5 'Line Discount
'    nEN_Disc06LnMDown = 6 ' Line Markdown
'    nEN_Disc07AutoTran = 7 'Auto Tran Disc.
'    nEN_Disc08MallCpn = 8 'Mall Coupon
'    nEN_Disc09CoopCpn = 9 'Cooperate Coupon
'    nEN_Disc10CoopBrand = 10 'Cooperate Brand
'    nEN_Disc11MallCrd = 11 'Mall Card
'    nEN_Disc12EmpCrd = 12 'Employee Card
'    nEN_Disc13TranDisc = 13 'Transaction Disc.
'    nEN_Disc14TranMDown = 14 ' Transaction Mkdn.
'    nEN_Disc15Ontop = 15 'On top
'    nEN_Disc16FreeGift = 16 'Free Gift (Special)
'
'End Enum

'property  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  -

Private dW_TransDate As Date
Private dW_EJDate As Date            '*Ao 2013/10/01       ReEntry  ให้เก็บ EJ ตาม Operation Date
Private tW_TmnNum As String
Private tW_PhyTmnNum As String
Private tW_TransNo As String
Private tW_DepID As String
Private tW_StoreID As String

Private dW_SysDate As Date
Private tW_SysTime As String

Private tW_TransType As String  'type หลัก
Private tW_TransTypeSub As String  'type ที่เป็นรูปแบบการพิมพ์ เช่น type หลัก เป็น Traninng mode 25 แต่ type รอง เป็น sale 3
Private tW_TransName As String
Private tW_TransNameABB As String
Private tW_TransTime As String
Public nW_PrintCnt As Long
Public bW_NoPrintCreditVoucher As Boolean '*Tao 62-09-12 ไม่พิมพ์ Credit Voucher

Private bW_TrainMode As Boolean
Private bW_RecallMode As Boolean
Private bW_RePrintMode As Boolean
Private bW_DepositMode As Boolean

Private tW_StoreNum As String
Private tW_StoreNameABB As String
Private tW_StoreNameTax As String
Private tW_StoreTel As String
Private tW_TaxInvoice As String

Private tW_EmpCdoe As String
Private tW_CustomerID As String
Private tW_CustomerVIPID As String
Private tW_SalesManID As String
Private tW_CashierName As String

Private tW_TaxInvoiceABB As String
Private tW_SlipCodeH As String
Private tW_SlipCodeT As String
Private aW_Header(1 To 7) As String
Private aW_Trailer(1 To 7) As String
'Private tW_LogoPath As String
Private oW_Logo As IPictureDisp

Private cW_SubAll As Double
Private cW_SubReg As Double
Private cW_SubTotal As Double
Private cW_Total As Double
Private cW_Rnd As Double
Private cW_Chg As Double
Private cW_MiscTnd As Double   '*Eaw 56-01-04 แก้ไขการปัด Rounding

Private cW_Vat As Double


'Private cW_DiscTotalTran As Double

'Private cW_DiscTotalTran2 As Double

Private bW_EnableEJ  As Boolean
Private bW_EnableRePrint As Boolean
Private bW_EnableRePrintNormal As Boolean 'CV
Private bW_EnableRePrintCV1 As Boolean 'CV
Private bW_EnableRePrintCV2 As Boolean 'Sale Endorsement

'deposit
Private cW_DepBalDue As Double
Private cW_DepPayment As Double
Private cW_DepPrevBalDue  As Double
Private cW_DepAmt  As Double
Private tW_DepNo  As String

Private tW_VoicherNo As String

Private eW_Deposit As eEN_Deposit

Private eW_PhoneNum As eEN_Deposit '*Neung 55-07-23 Layaway TK
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  -

Private orsHD As ADODB.Recordset
Private orsDT As ADODB.Recordset
Private orsRC As ADODB.Recordset
Private orsCD As ADODB.Recordset
Private orsCV As ADODB.Recordset
Private orsPmtCpn As ADODB.Recordset

Private orsPrintSeq  As ADODB.Recordset

Private cW_PageWidth As Double
Private cW_MrakOffSet As Double
Private uW_X  As eEN_XPos
Private cXL As Double, cXHL As Double, cXC As Double, cXHR As Double, cXR As Double
Private cXHC As Double, cXMK  As Double '  อาจไม่ถูก standard แต่ต้องการให้สั้น
Private cX5 As Double, cX2 As Double, cX3 As Double, cX4 As Double, cX15 As Double, cX8 As Double, cX8x As Double
Private tW_TrMk As String
Private nW_Items As Long
Private Const tW_CurFmt = "#,##0.00"
Private Const tTkMdTrainingMrak = "Tr"

Private tW_TransWhere As String

Private bW_PrnTraining As Boolean
Private bW_PrnVoid As Boolean
Private tW_DateFmt As String

Private tW_OptAlw As String

Private tW_CstCardCode As String
Private nW_LPoint As Long
Private tW_NamePoint As String '*KT 52-02-23 member point
Private nW_ToDayPoint As Long
Private tW_CstCardCode2 As String
Private nW_LPoint2 As Long
Private tW_NamePoint2  As String '*KT 52-02-23 member point
Private nW_ToDayPoint2 As Long
Private tW_CstCardCode3 As String
Private nW_LPoint3 As Long
Private tW_NamePoint3 As String '*KT 52-02-23 member point
Private nW_ToDayPoint3 As Long
Private tW_PrnPoint As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Bak
Private tW_ExportFolder As String
Private tW_EJFolder As String
'Private bW_SaveField As Boolean

Public Enum eEN_BackupFormat
        nEN_Sales = 1
End Enum

'DMO
Private oW_Server As New SQLDMO.SQLServer
Private oW_Database As SQLDMO.Database
Private oW_Table As SQLDMO.Table
Private oW_BulkCp As New SQLDMO.BulkCopy

Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Private Const LOCALE_SSHORTDATE = &H1F

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 255
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'

Private bW_ShwSubTotal As Boolean

Private bC_PrnByOPOS As Boolean

Private cW_EmpBalDis As Double

Private cW_OnTopDps  As Double

Private tW_StatusOffline As String '*KT 51-12-15 Phase 2 ข้อ 31 Communication Sheet PH2-ML-SF-031 Print Messageon Receipt  when OFFLINE.xls ถ้าเป็น 1: Offline นอกนั้นเป็น Online

Private bW_PRMGenEJ As Boolean '*KT 51-12-16 Flag ที่บอกว่ามีการเรียกใช้งานจาก Tools Generate EJ

Private tW_GETFileEJName As String  '*KT 51-12-22

Private aW_CrossDeal() As String

Private tW_MemberPriceID As String '*KT 52-04-10 กรณีมีการใช้ Member Price หรือ Member Point ให้พิมพ์ Member ID:  และหากเป็น member ที่ได้ส่วนลดด้วย ก็พิมพ์ แค่ครั้งเดียว ไม่จำเป็นต้องพิมพ์ Customer Id: อีก หากเป็น Card ที่มีส่วนลดอย่างเดียว ก็พิมพ์คำว่า Customer ID: ตามเอกสาร phase 1.8 UAT status on 08-04-2009

Private aW_SalemanID() As String '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
Private tW_SalesManIDDT As String '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย

Private bW_StaPrnPmtCpn As Boolean '*KT 54-07-25 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon

Private tW_TypePrinter As String '0:Slip Printer 1:Receipt Printer

Private tW_TransDate As String
Private tW_StaCancel As String

Private bW_StaPrnCV2 As Boolean
Private bW_StaPrnCV3 As Boolean

'*Thirapong 2021-03-27
'*CommSheet TTK-ISS3000-719 TK ISSUE - Zero Credit Voucher_Rev01.xlsx
'*CommSheet TTK-ISS3000-767 TK ISSUE - Credit Voucher Balance Zero Amount Shown In Re-issue on v5.2002.1.xlsx
'*CommSheet TTK-ISS3000-774 TK ISSUE - Credit Voucher Balance Zero Amount Shown In Re-issue on V. 5.2010.2-00.01.00.xlsx
'*CommSheet TTK-ISS3000-782 TK ISSUE - Recurred error on Zero Amount on Re-issued CV.xlsx
Private bW_StaPrnCVReIssue As Boolean
'******************************************************************************************
'Private tW_ReIssueExp As String, tW_ReIssueNo As String, cW_ReIssueAmt As Double
Private Type eEN_ReIssue
        tEN_ReIssueExp As String
        tEN_ReIssueNo As String
        cEN_ReIssueAmt  As Double
        tEN_ReIssueTel As String '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
End Type
Private Type eEN_CV
        tEN_CVExp As String
        tEN_CVNo As String
        cEN_CVAmt  As Double
        tEN_CVTel As String '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
End Type
Private oW_ReIssue As eEN_ReIssue
Private oW_CV As eEN_CV

Private bW_ChkMark As Boolean
Private bW_ChkMEM As Boolean
Private bW_ChkEmp As Boolean
Private bW_PrintMem As Boolean

Private tW_RefDepNo As String    '*Eaw 56-06-03 CommSheet TK-ISS3000-198-ST 00 00 01.xls
Private tW_RefDepNum As String
Private bW_UseServer As String '*Eaw 56-07-24
Private tW_AccTypeReEntry As String '*Eaw 56-08-29
Private tW_PrintType As String '*Bump 57-04-05  CommSheet TK-ISS3000-461-00.00.04.xls
Private tW_VirTerminal As String '*BG 56-09-23 Comsheet-356 print EJ tenant sale
Private nW_LineNo As Integer '*Eaw 56-10-08
Private nW_LineMax As Integer '*Eaw 56-10-08
Private tW_SlipFooterStr As String '*Eaw 56-10-08
Private tW_Replicated As String
Private bW_ChkRep As Boolean
Private cW_RegAmt As Double '*Tao 60-04-10
Private cW_LowBPAmt As Double '*Tao 60-04-10
Private nW_TokenTransType  As Long '*Bmp 61-04-05
Private cW_HDRate As Double '*KT 66-08-03
Private cW_TaxRate As Double '*KT 66-08-03


 '*KT 51-12-22
Property Get tGETFileEJName() As String
    tGETFileEJName = tW_GETFileEJName
End Property

 '*KT 51-12-16 Flag ที่บอกว่ามีการเรียกใช้งานจาก Tools Generate EJ
Property Let bPRMGenEJ(New_bPRMGenEJ As Boolean)
    bW_PRMGenEJ = New_bPRMGenEJ
End Property

 '*KT 51-12-16 Flag ที่บอกว่ามีการเรียกใช้งานจาก Tools Generate EJ
Property Get bPRMGenEJ() As Boolean
    bPRMGenEJ = bW_PRMGenEJ
End Property

Property Let cEmpBalDis(New_cEmpBalDis As Double)
    cW_EmpBalDis = New_cEmpBalDis
End Property

Property Let nXCPic(New_nXCPic As Double)
    cXC = New_nXCPic
End Property

Property Let bPrnByOPOS(New_bPrnByOPOS As Boolean)
    bC_PrnByOPOS = New_bPrnByOPOS
End Property

Property Get bPrnByOPOS() As Boolean
    bPrnByOPOS = bC_PrnByOPOS
End Property

Public Property Set ActivePrinter(poPrinterObj As Printer)
        nW_StatePrn = IIf(poPrinterObj Is Nothing, 0, 1)
End Property

Public Property Set ActiveConnection(poConnection As ADODB.Connection)
        nW_StateCon = IIf(poConnection Is Nothing, 0, 1)
        Set ocnAll = poConnection
        
        'get paramiter
        'tW_DateFmt = W_GETtDateTimeFmt
        
End Property

Public Property Get TherminalNo() As String
        TherminalNo = tW_TmnNum
End Property

Public Property Get TransactionDate() As String
        TransactionDate = tW_TmnNum
End Property

Public Property Get TransactionNo() As String
        TransactionNo = tW_TransNo
End Property

Public Property Set Logo(poPicture As IPictureDisp)
        Set oW_Logo = poPicture
End Property

Public Property Get Logo() As IPictureDisp
        Set Logo = oW_Logo
End Property

Private Sub W_SETxReplicated(ByVal pnTransType As Long)
Dim tSql As String, tPrintRep As String
Dim orsTmp As ADODB.Recordset

On Error GoTo ErrHandler
    
    bW_ChkRep = False
    
    tSql = "SELECT FTSttPrintRep FROM TSysTransType"
    tSql = tSql & vbCrLf & "WHERE FTSttTranCode = '" & Format(pnTransType, "00") & "'"
    If SP_SQLvExecute(tSql, orsTmp, oDB.ocnOnLine, , , , , True, True) = 0 Then
        If Not (orsTmp.BOF And orsTmp.EOF) Then
            tPrintRep = Trim(SP_FEDtChkString(orsTmp, "FTSttPrintRep"))
            If UCase(tPrintRep) = UCase("Y") Then
                bW_ChkRep = True
            Else
                bW_ChkRep = False
            End If
        End If
    End If
    
    Call SP_DATxRsNothing(orsTmp)
Exit Sub
ErrHandler:
'    MsgBox tSql
    bW_ChkRep = False
End Sub

Private Sub W_PRNxTicketReplicated()

On Error GoTo ErrHandler

'    tRepMsg = GetSetting("AdaPosMall", "POSFront", "RepMsg")
'    MsgBox "tW_Replicated = " & tW_Replicated
    If tW_Replicated <> "" Then
'        Call oPrn.PrintText(tW_Replicated, cXL, , vbLeftJustify, True)
        tW_SlipFooterStr = tW_SlipFooterStr & "1" & tW_Replicated & ";"
    End If
Exit Sub
ErrHandler:
End Sub

Private Sub W_PRNxTicketCashTrDeno(ByVal ptOrsDT As ADODB.Recordset)
'-----------------------------------------------------------
'   Parameter :
'   Cmt :
'   File Refer :
'   Developer : BUNG
'   Date Create : 29/04/2014 :  13:18:34
'-----------------------------------------------------------
    Dim nSeq As Long
    Dim nTmp As Long, tTmp As String, cTmp As String
    Dim cWidthLimit As Double
    Dim tPdtBar As String, tPdtName As String, tPdtName2 As String
    Dim tPdtDpt As String, tPdtDptCut As String
    Dim cQty As Double, cRegPrice As Double, cSalePrice As Double, cNet As Double
    Dim tTaxType As String, cVatRate As Double, cVatAmt As Double
    Dim tPdtType As String
    Dim tCancelFlag  As String
    Dim tTxMk As String, tTrMk As String, tNdMk As String
    Dim tMallCardFlag As String
    Dim tEmpCardFlag As String
    Dim tBPFlag As String
    Dim cWeight As String
    Dim cDiscVal As Double
    Dim cDutyFreeDeposit As Double
    Dim bHasLineDisc As Boolean
    Dim bHasPromo As Boolean
    Dim cHasPromo As Double
    'Dim nIsPromo As Boolean '3 : Promo,2 : Promo Time
    Dim nIsPromo As Long '3 : Promo,2 : Promo Time    '*Ao 2011/02/15 CommSheet ML-PRV-040-EJ Shows Promo Price wrong-.xls
    Dim bHasDeal As Boolean
    Dim nDealGrp As Long
    Dim tDealNum  As String
    Dim tPdtSkuCode As String 'Bump 56-03-02 TK -ISS3000 - 55 :copy จาก  Foodland, CommSheet ST-FL-047-SF- Print UPC or SKU on Receipt and EJ.xls
    Dim orsTmp As ADODB.Recordset
    Dim orsTmp2 As ADODB.Recordset
    Dim orsNew1 As ADODB.Recordset
    Dim orsNew2 As ADODB.Recordset

    Dim tSql As String
    
    Dim tLogType As String
    Dim cLogQty As String
    Dim nLogDTSeq As Long
    
    Dim tPriceOver As String
    
    Dim nDealGroup  As Long
    
    Dim cTaxExceptAmt As Double, cDiscPer As Double
    Dim tMallEmpFlag As String
    
    'Line_Chars
    'Const tTkWeight = "Weight"
    Dim tTkWeight$
    tTkWeight = W_GETxVocabularyTicketFromArray("1942")
    'Const tTkDsRegPri = "Reg.Price"
    Dim tTkDsRegPri$
    tTkDsRegPri = Mid(W_GETxVocabularyTicketFromArray("1900"), 1, 7)

    'Const tTkDsDis = "Coop DIS."
    Dim tTkDsCoDis$
     tTkDsCoDis = W_GETxVocabularyTicketFromArray("1908")
    
    'Const tTkDsDis = "DIS."
    Dim tTkDsDis$
     tTkDsDis = W_GETxVocabularyTicketFromArray("190601")
      'Const tTkDsDisAmt = "DISA."
    Dim tTkDsDisAmt$
    tTkDsDisAmt = W_GETxVocabularyTicketFromArray("190602")
     'Const tTkDsDisAmt = "DIS."
    Dim tTkDsDisNew$
    tTkDsDisNew = W_GETxVocabularyTicketFromArray("190603")
    
     
     'Const tTkDsMkDn = "MKDN."
    Dim tTkDsMkDn$
    tTkDsMkDn = W_GETxVocabularyTicketFromArray("190701")
    Dim tTkDsMkDnAmt$
    tTkDsMkDnAmt = W_GETxVocabularyTicketFromArray("190702")
     'Const tTkDsMkDn = "MKDN."
    Dim tTkDsMkDnNew$
    tTkDsMkDnNew = W_GETxVocabularyTicketFromArray("190703")
     
    'Const tTkDsPromo = "Promo"
    Dim tTkDsPromo$
    tTkDsPromo = Mid(W_GETxVocabularyTicketFromArray("1903"), 1, 7)
    'Const tTkDsPromo = "Promo"
    Dim tTkDsPromoTime$
    tTkDsPromoTime = Mid(W_GETxVocabularyTicketFromArray("1904"), 1, 7)

    'Const tTkDsCoCpn = "CoCpn."
    Dim tTkDsCoCpn$
    tTkDsCoCpn = W_GETxVocabularyTicketFromArray("1917")
    'Const tTkDsCoBnd = "CoBnd."
    Dim tTkDsCoBnd$
    tTkDsCoBnd = W_GETxVocabularyTicketFromArray("1918")
    
    'Const tTkDsDeal = "Deal"
    Dim tTkDsDeal$
    tTkDsDeal = W_GETxVocabularyTicketFromArray("1905")
    
    'Const tTkPost = "POST#"
    Dim tTkPost$
    tTkPost = W_GETxVocabularyTicketFromArray("1927")
    'Const tTkTran = "TRAN#"
    Dim tTkTran$
    tTkTran = W_GETxVocabularyTicketFromArray("1928")
    'Const tTkAmt = "AMOUNT"
    Dim tTkAmt$
    tTkAmt = W_GETxVocabularyTicketFromArray("1929")
    
    'Const tTkOrgTotalUpd = "Original Post Total Updated"
    Dim tTkOrgTotalUpd$
    tTkOrgTotalUpd = W_GETxVocabularyTicketFromArray("1930")
    
    'Const tTkDepNo = "Deposit No.: "
    Dim tTkDepNo$
    tTkDepNo = W_GETxVocabularyTicketFromArray("1931")
    
    '*Neung 55-07-23 Layaway TK -------------------------------------------------
    'Const tTkPhoneNum = "Phone Number: "
    Dim tTkPhoneNum$
    tTkPhoneNum = W_GETxVocabularyTicketFromArray("1997")
    '----------------------------------------------------------------------------------------------

    'Const tTkCurrency = "Currency"
    Dim tTkCurrency$
    tTkCurrency = W_GETxVocabularyTicketFromArray("1932")
    Dim tTkCoin$
    tTkCoin = W_GETxVocabularyTicketFromArray("5")
    Dim tTkGV$
    tTkGV = W_GETxVocabularyTicketFromArray("6")
    Dim tTkForeign$
    tTkForeign = W_GETxVocabularyTicketFromArray("7")
    
    
    'Const tTkTotal = "TOTAL"
    Dim tTkTotal$
    tTkTotal = W_GETxVocabularyTicketFromArray("1933")
    
    'Const tTkDepAmt = "Deposit Amt. "
    Dim tTkDepAmt$
    tTkDepAmt = W_GETxVocabularyTicketFromArray("1934")
    
    'Const tTkVat = "Vat"
    Dim tTkVat$
    tTkVat = W_GETxVocabularyTicketFromArray("1935")
    
    
    'Const tTkVoucherNo = "Voucher No.:"
    Dim tTkVoucherNo$
    tTkVoucherNo = W_GETxVocabularyTicketFromArray("1936")
    
    'Const tTkLine = "- - LINE "
    Dim tTkLine$
    tTkLine = W_GETxVocabularyTicketFromArray("1937")
    'Const tTkLineCancel = "CANCELLED "
    Dim tTkLineCancel$
    tTkLineCancel = W_GETxVocabularyTicketFromArray("1938")
    'Const tTkLineQtyChange = "Qty Changed TO "
    Dim tTkLineQtyChange$
    tTkLineQtyChange = W_GETxVocabularyTicketFromArray("1939")
    'Const tTkLineNagativeScan = "Cancelled "
    Dim tTkLineNagativeScan$
    tTkLineNagativeScan = W_GETxVocabularyTicketFromArray("1940")
    'Const tTkLineQuantity = "Quantity "
    Dim tTkLineQuantity$
    tTkLineQuantity = W_GETxVocabularyTicketFromArray("1941")
    
    'Const tTkDsRegPri = "Reg.Price"
    Dim tTkDsMBPri$
    tTkDsMBPri = Mid(W_GETxVocabularyTicketFromArray("1989"), 1, 7)
    
    'Const tTkDsRegPri = "Reg.Price"
    Dim tTkDsMPDis$
    tTkDsMPDis = Mid(W_GETxVocabularyTicketFromArray("1990"), 1, 7)
    
    'Const tTkLnVendorCpn = "Line Vendor Coupon"
    Dim tTkLnVendorCpn$
    tTkLnVendorCpn = W_GETxVocabularyTicketFromArray("1991") '*KT 53-10-16 Line Vendor Coupon
    Dim tIndexLnVendorCpn As String '*KT 53-10-16 Line Vendor Coupon
    Dim tQtyLnVendorCpn As String '*KT 53-10-16 Line Vendor Coupon
    Dim nQtyLnVendorCpn As Double '*KT 53-10-16 Line Vendor Coupon
    Dim cLnVendorCpnNet As Double '*KT 53-10-16 Line Vendor Coupon
    Dim cLnVendorCpnPerQty  As Double '*KT 53-10-16 Line Vendor Coupon
    
    Dim bPass As Boolean
    Dim nI As Long
    
    Dim tTkCooCpn As String
    tTkCooCpn = "Cooperate Coupon "
    
    Dim tTkCooBrn As String
    tTkCooBrn = "Cooperate Brand "

    Dim tLen As String
    
    Dim bDealTypeFree  As Boolean
    
    Dim tTmpBlank$
    
    Dim nPosNo As Long
    Dim tTrnNo As String
    Dim cNetBDiscAll As Double
    Dim cSumDiscOnLn As Double '*KT 51-10-01  แก้ไข net amount
    Dim cB4Dis As Double '*KT 51-01-24 เก็บค่า Net Amount (qty*price)ก่อนหักส่วนลดใดๆๆ
    
    'Deal
    Dim bPntDeal As Boolean
    Dim bPntDealST As Boolean
    Dim bPntDealEnd As Boolean
    Dim bCross As Boolean
    Dim nOldPosition As Long
    Dim bIsKeep As Boolean
    Dim cHasDeal As Double
        
    Dim tInStoreBar As String
    Dim tSdtStaRfd As String
    Dim cMPrice  As Double
    Dim bHasMPrice As Boolean
    Dim bHasMPDisc As Boolean
    Dim nLenTha As Long
    Dim tTempName1 As String
    Dim tTempName2 As String
    
    Dim tPrintWeight As String
    Dim cHasMember21  As Double
    
    Dim tSKUCode As String  '*Eaw 56-02-20
    Dim orsOverP As ADODB.Recordset '*Eaw 56-03-13
    Dim cPriceOver As Double '*Eaw 56-03-13
    Dim tForFac As String  '*Eaw 56-04-11
    '--------------------------------------
    '*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
    'Dim tDealPromo As String
    'Dim tHaveDealPromo As String
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '*KT 55-08-07 Takashimaya
    Dim nLoop As Long
    Dim tType As String
    Dim nRecordCount As Integer, nCount As Integer
    Dim tSpace As String
    Dim cTotalCashOverShort As Double
    
    On Error GoTo ErrHandle
        With ptOrsDT
            nLoop = 0
             Do While Not ptOrsDT.EOF
                nLoop = nLoop + 1
               Select Case nLoop
                    Case 1
                        'coin
                        If SP_FEDvChkNumeric(ptOrsDT, "FCSdtSaleAmt") <> 0 Then
                            Call oPrn.PrintText(tTkCoin, cXL, , vbLeftJustify, False) 'tTkCurrency
                            Call oPrn.PrintText(Format(SP_FEDvChkNumeric(ptOrsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                            Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                            cTotalCashOverShort = cTotalCashOverShort + SP_FEDvChkNumeric(ptOrsDT, "FCSdtSaleAmt")
                        End If
                    Case 2
                        'currency
                        If SP_FEDvChkNumeric(ptOrsDT, "FCSdtSaleAmt") <> 0 Then
                            Call oPrn.PrintText(tTkCurrency, cXL, , vbLeftJustify, False) 'tTkCurrency
                            Call oPrn.PrintText(Format(SP_FEDvChkNumeric(ptOrsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                            Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                            cTotalCashOverShort = cTotalCashOverShort + SP_FEDvChkNumeric(ptOrsDT, "FCSdtSaleAmt")
                        End If
                    Case 3
                        'GV
                        If SP_FEDvChkNumeric(ptOrsDT, "FCSdtSaleAmt") <> 0 Then
                            Call oPrn.PrintText(tTkGV, cXL, , vbLeftJustify, False) 'tTkGV
                            Call oPrn.PrintText(Format(SP_FEDvChkNumeric(ptOrsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                            Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                        End If
                    Case 4
                        'Foreign(Base)
                        If SP_FEDvChkNumeric(ptOrsDT, "FCSdtSaleAmt") <> 0 Then
                            Call oPrn.PrintText(tTkForeign, cXL, , vbLeftJustify, False) 'tTkForeign
                            Call oPrn.PrintText(Format(SP_FEDvChkNumeric(ptOrsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                            Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                        End If
                End Select
                
                If nLoop = 2 And (Val(tW_TransTypeSub) = 37 Or Val(tW_TransTypeSub) = 41 Or Val(tW_TransTypeSub) = 46) Then
                    'Call oPrn.PrintText(tTkTotal, cXL, , vbLeftJustify, False) 'tTkTotal
                    Call oPrn.PrintText("TOTAL CASH", cXL, , vbLeftJustify, False)
                    Call oPrn.PrintText(Format(cTotalCashOverShort, tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                    Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                End If
                ptOrsDT.MoveNext
            Loop
            'total
            If Val(tW_TransTypeSub) <> 37 And Val(tW_TransTypeSub) <> 41 And Val(tW_TransTypeSub) <> 46 Then
                Call oPrn.PrintText(tTkTotal, cXL, , vbLeftJustify, False) 'tTkTotal
                'Call oPrn.PrintText("TOTAL CASH", cXL, , vbLeftJustify, False)
                Call oPrn.PrintText(Format(cW_Total, tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
            End If
            'blank
            ptOrsDT.Filter = adFilterNone
            'If Val(tW_TransTypeSub) = 37 Then
            Select Case Val(Format(tW_TransTypeSub))
                Case 37, 41, 46
                    Call W_PRNbTicketOverShort
            End Select
            'End If
            Call oPrn.PrintText("", cXL, , vbLeftJustify, True)
        End With
        
    Exit Sub
ErrHandle:
        
End Sub

Private Sub W_PRNxTicketCashTrNO(ByVal porsDT As ADODB.Recordset)
'-------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 2013/03/07
'-------------------------------------------------------------------
Dim nLoop As Integer
Dim tTmp As String, tPdtName As String

Dim cRateCurrency As Double, cQty As Double, cNet As Double, cTotalAmt As Double
Dim nSpace As Integer, nSpaceEx As Integer
Dim orsFor As ADODB.Recordset
Dim tSql As String
Dim cTotalGVchAmt As Double
On Error GoTo ErrHandler

nSpace = 25
cTotalGVchAmt = 0
tSql = "SELECT FTFcrCode,FTFcrSign,FTFcrType,FCFcrExchang FROM TCNMForCurrency"
Call SP_SQLvExecute(tSql, orsFor, oDB.ocnOnLine, , , , , True, True) 'read

For nLoop = 0 To 3
    With porsDT
        .Filter = "FTSdtStaPdt='" & CStr(nLoop) & "' AND FCSdtQty > 0"
        .Sort = "FNSdtSeqNo ASC"
         
        If Not (porsDT.BOF And porsDT.EOF) Then
            Do While Not porsDT.EOF
                tPdtName = SP_FEDtChkString(porsDT, "FTSkuAbbName")
                tTmp = ""
                If nLoop = 3 Then
                    cRateCurrency = SP_FEDvChkNumeric(porsDT, "FCSdtStkFac")
                    tTmp = Format(cRateCurrency, "#0.00000")
                Else
                    cQty = SP_FEDvChkNumeric(porsDT, "FCSdtQty")
                    If UCase(tPdtName) <> UCase("Other Cash") And UCase(tPdtName) <> UCase("Other GVch") Then
                        tTmp = Format(cQty, "###0") & "X"
                    End If
                End If
                tPdtName = Mid(tPdtName, 1, 20) 'SKU Name  Abb Eng
                If nSpace > (Len(tTmp) + Len(tPdtName)) Then
                    nSpaceEx = nSpace - (Len(tTmp) + Len(tPdtName))
                Else
                    nSpaceEx = 0
                End If
                If nLoop = 3 Then
                    Call oPrn.PrintText(tPdtName & Space(nSpaceEx) & tTmp, cXL, , vbLeftJustify, False)
                Else
                    Call oPrn.PrintText(tTmp & Space(nSpaceEx) & tPdtName, cXL, , vbLeftJustify, False)
                End If
                
                If nLoop = 3 Then
                    cNet = SP_FEDvChkNumeric(porsDT, "FCSdtRegPrice")
                    '----------------------------------------------------------------------------------------------
                    tTmp = SP_FEDtChkString(porsDT, "FTSkuCode")
                    tTmp = Trim(tTmp)
                    orsFor.Filter = "FTFcrCode = '" & tTmp & "'"
                    If Not (orsFor.BOF And orsFor.EOF) Then
                        tTmp = SP_FEDtChkString(orsFor, "FTFcrSign")
                        tTmp = Trim(tTmp)
                    Else
                        tTmp = ""
                    End If
                    orsFor.Filter = adFilterNone
                    '----------------------------------------------------------------------------------------------
                Else
                    cNet = SP_FEDvChkNumeric(porsDT, "FCSdtSaleAmt")
                    tTmp = ""
                End If
                
                Call oPrn.PrintText(tTmp & Format(cNet, tW_CurFmt), cXMK, , vbRightJustify, True)
                If nLoop = 0 Or nLoop = 1 Then 'Coin + Curreny = Cash Total
                    cTotalAmt = cTotalAmt + cNet
                End If
                If nLoop = 2 Then
                    cTotalGVchAmt = cTotalGVchAmt + cNet
                End If
                .MoveNext
            Loop
        End If
        
        If nLoop = 1 Then
            If cTotalAmt > 0 Then
                tTmp = "TOTAL CASH"
                nSpaceEx = nSpace - Len(tTmp)
                tTmp = Space(nSpaceEx) & tTmp
                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                Call oPrn.PrintText(Format(cTotalAmt, tW_CurFmt), cXMK, , vbRightJustify, True)
            End If
        End If

        If nLoop = 2 Then
            If cTotalGVchAmt > 0 Then
                tTmp = "TOTAL GVch"
                nSpaceEx = nSpace - Len(tTmp)
                tTmp = Space(nSpaceEx) & tTmp
                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                Call oPrn.PrintText(Format(cTotalGVchAmt, tW_CurFmt), cXMK, , vbRightJustify, True)
            End If
        End If

    End With
    'ขยาย เคป้ะ
    
Next nLoop
Exit Sub
ErrHandler:

End Sub


Private Function W_GETbGetTBCV(ByRef porsTmp As ADODB.Recordset, ByVal ptTbName As String, ByVal ptWhere As String) As Boolean
'_________________________________________________
'   Call:
'   Cmt:    '*Eaw 56-10-18
'_________________________________________________
Dim tSql As String

On Error GoTo ErrHandler
    
    tSql = "SELECT * FROM " & ptTbName
    tSql = tSql & vbCrLf & " WHERE " & ptWhere
    If SP_SQLvExecute(tSql, porsTmp, oDB.ocnOnLine) = 0 Then
        'PASS มี CV ใน DB Online
    Else
'        Call SP_DATxRsNothing(porsTmp)
        '*Eaw 56-11-18 กรณีที่ใช้งาน แบบ Offline
        If SP_SQLvExecute(tSql, porsTmp, oDB.ocnOffLine) = 0 Then
            'PASS มี CV ใน DB offline
        Else
            Call SP_DATxRsNothing(porsTmp)
        End If
    End If
    W_GETbGetTBCV = True
    Exit Function
ErrHandler:
    Call SP_DATxRsNothing(porsTmp)
    W_GETbGetTBCV = True
End Function

Public Sub W_PRNxPrintSlip(ByVal ptStr As String, ByVal ptTypePrinter As String, ByVal pnLine As Integer, ByVal pnLineMax As Integer)
Dim arrStr() As String
Dim nLoop As Integer, nCount As Integer
Dim tTmp As String, tPrint As String, tCase As String
Dim bChk As Boolean

On Error GoTo ErrHandle
    arrStr = Split(ptStr, ";")
    
    nCount = UBound(arrStr)
    
    If ptTypePrinter = "0" Then
        If pnLine + nCount <= pnLineMax + 2 Then
            bVB_LastLine = False
        Else
            bVB_LastLine = True
        End If
    Else
        bVB_LastLine = False
    End If
    
    For nLoop = 0 To UBound(arrStr)
        tTmp = arrStr(nLoop)
        Do While tTmp <> ""
'            tPrint = cUT.UT_STRtToken(tTmp, ",", 1)
            tPrint = cUT.UT_STRtToken(tTmp, "{", 1)
            tCase = Left(tPrint, 1)
            tPrint = Right(tPrint, Len(tPrint) - 1)
            Select Case tCase
                Case "1"
                    Call oPrn.PrintText(tPrint, cXL, , vbLeftJustify, True)
                Case "2"
                    Call oPrn.PrintText(tPrint, cXMK, , vbRightJustify, False)
                Case "3"
                    Call oPrn.PrintText(tPrint, cXMK, , vbLeftJustify, True)
                Case "4"
                    Set oPrn.Font = oW_Font(2)
                    Call oPrn.PrintText(tPrint, cXC, , vbCenter, , , , , True)
                    Set oPrn.Font = oW_Font(1)
                Case "5"
                    Call oPrn.PrintText(tPrint, cXC - cXHC, , vbLeftJustify, False)
                Case "6"
                    Call oPrn.PrintText(tPrint, cXHR - cX2, , vbLeftJustify, True)
                Case "7"
                    Call oPrn.PrintText(tPrint, cXR, , vbRightJustify, True)
                Case "8"
                    Call oPrn.PrintText(tPrint, cXL, , vbLeftJustify, False)
                Case "9"
                    Call oPrn.PrintText(tPrint, cXC, , vbRightJustify, False)
                Case "$"
                    Call oPrn.PrintText(tPrint, cXMK, , vbRightJustify, True)
                Case "#"
                    Call oPrn.PrintText(tPrint, cX5, , vbLeftJustify, False)
                Case "0"
                    Call oPrn.PrintText(tPrint, cX2, , vbLeftJustify, True)
                Case "*"
                    Call oPrn.PrintText(tPrint, cXC, , vbCenter, True, nEN_PrinterOnly)
            End Select
        Loop
    Next nLoop
    Exit Sub
ErrHandle:
End Sub

Private Sub W_PRNbTicketPointRe()
'_________________________________________________
'   Call:
'   Cmt:    '*Ao 2013/11/30 CommSheet ML-PH4-R3-051 Bonus point missing .xls
'_________________________________________________
Dim tSql As String
Dim orsTmpSP As ADODB.Recordset
Dim orsTmpHD As ADODB.Recordset
Dim orsMem As ADODB.Recordset
Dim tID As String, tName As String, tPoint As String, tMemID As String
Dim nPoint As Long, nBalance As Long
Dim tSaveID As String
Dim tSaveName As String
Dim tSavePoint As String
Dim tRefMin As String, tRefMax As String, tRemark As String
Dim tTmp As String, tPointEx As String
Dim tPointBefore As String, tBonusPoints As String, tTotalBonus As String
Dim tBP1 As String, tBP2 As String, tExpDate As String, tExpiry As String

On Error GoTo ErrHandler
Dim tTkPointL1$
Dim tTkPointL2$
Dim tTkPointL3$
Dim tTkPointL4$
Dim tTkPointL5$
Dim tTkPointL6$
Dim tTkPointL7$
Dim tTkPointL8$
   
    tTkPointL1 = W_GETxVocabularyTicketFromArray("1980")
    tTkPointL2 = W_GETxVocabularyTicketFromArray("1981")
    tTkPointL3 = W_GETxVocabularyTicketFromArray("1982")
    tTkPointL4 = W_GETxVocabularyTicketFromArray("1983")
    tTkPointL5 = W_GETxVocabularyTicketFromArray("1984")
    tTkPointL6 = W_GETxVocabularyTicketFromArray("1985")
    tTkPointL7 = W_GETxVocabularyTicketFromArray("1986")
    tTkPointL8 = W_GETxVocabularyTicketFromArray("1987")
    
    tBP1 = W_GETxVocabularyTicketFromArray("22")
    tBP2 = W_GETxVocabularyTicketFromArray("23")
    tExpiry = W_GETxVocabularyTicketFromArray("24")
       
    'tSql = "SELECT (FTCstPointName+FTShdDocRefMin) AS FTCstPointName,FTCstPointName2,FTCstPointName3 FROM " & tW_TblHD '*Tao 60-06-07 ของเดิม
    tSql = "SELECT FTCstPointName,FTCstPointName2,FTCstPointName3,FTShdDocRefMin,FTShdDocRefMax,FTRemark FROM " & tW_TblHD '*Tao 60-06-07 ของเดิม
    tSql = tSql & vbCrLf & "WHERE " & tW_TransWhere
    'Call SP_SQLvExecute(tSql, orsTmpHD) '*Tao 60-06-07 ของเดิม
    Call SP_SQLvExecute(tSql, orsTmpHD, IIf(bW_UseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(bW_UseServer, True, False)) '*Tao 60-06-07 Reprint หลังจาก Closed Teminal
    If Not (orsTmpHD.BOF And orsTmpHD.EOF) Then
        orsTmpHD.MoveFirst
        tID = SP_FEDtChkString(orsTmpHD, "FTCstPointName")
        tName = SP_FEDtChkString(orsTmpHD, "FTCstPointName2")
        tPoint = SP_FEDtChkString(orsTmpHD, "FTCstPointName3")
        tRefMin = SP_FEDtChkString(orsTmpHD, "FTShdDocRefMin")
        tRefMax = SP_FEDtChkString(orsTmpHD, "FTShdDocRefMax")
        tRemark = SP_FEDtChkString(orsTmpHD, "FTRemark")
        
        '*Tao 60-06-07 เช็คข้อมูลจาก Version เดิม ว่าจะเอา Column ไหนมารวมกัน
        If Left(tRemark, 2) = "V." Then tRemark = ""
        Select Case Len(tID)
        Case 45
            'ของเวอร์ชั่นเดิม
            tID = tID & tRefMin
        Case Is > 45
            'ของเวอร์ชั่นใหม่
            tID = tID & tRemark
            tName = tName & tRefMin
            tPoint = tPoint & tRefMax
        End Select
        '=====================================================
        
        If tID <> "" Then '*Tao 60-06-07 ถ้าใช้ Member เกิน 2 ครั้ง จะบันทึกลง HD ไม่ได้
            'บันทึกได้
            Do While tID <> ""
                tTmp = ""
                tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                tExpDate = cUT.UT_STRtToken(tID, ";", 1)
                tTmp = cUT.UT_STRtToken(tExpDate, ",", 1)                                                   '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
             
                tTmp = tTkPointL1 & tTmp
                tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                tTmp = cUT.UT_STRtToken(tName, ";", 1)
                tTmp = tTkPointL2 & " " & tTmp
                tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                
                tPointEx = cUT.UT_STRtToken(tPoint, ";", 1)
                tPointBefore = cUT.UT_STRtToken(tPointEx, ",", 1, True)
                tBonusPoints = cUT.UT_STRtToken(tPointEx, ",", 2, True)
                tTotalBonus = cUT.UT_STRtToken(tPointEx, ",", 3, True)
    
                If Mid(tW_OptAlw, 2, 1) = "1" Then
                    tTmp = tTkPointL3
                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                    tTmp = tPointBefore
                    tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                End If
                If Mid(tW_OptAlw, 3, 1) = "1" Then
                    tTmp = tTkPointL5
                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                    tTmp = tBonusPoints
                    tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                End If
                If Mid(tW_OptAlw, 2, 1) = "1" Or Mid(tW_OptAlw, 3, 1) = "1" Then
                    tTmp = tTkPointL7
                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                    tTmp = tTotalBonus
                    tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                End If
                
                'tW_SlipFooterStr = tW_SlipFooterStr & "1" & tExpiry & " " & tExpDate & ";"           '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
                If Trim(tExpiry) <> "" Then tW_SlipFooterStr = tW_SlipFooterStr & "1" & tExpiry & " " & tExpDate & ";"           '*Bump 58-02-12 CommSheet TK-ISS3000-578.xls
            Loop
        Else
            '*Tao 60-06-07 ถ้าบันทึกไม่ได้ ให้ไปเอาจาก TPSTSalePoint
            tSql = "SELECT SUM(FCSpoPoint) AS FCSpoPoint,MIN(CONVERT(INT,FTSpoID)) AS FTSpoID,FTSpoMemID,FDDateIns"
            tSql = tSql & vbCrLf & "FROM TPSTSalePoint"
            tSql = tSql & vbCrLf & "WHERE " & tW_TransWhere
            tSql = tSql & vbCrLf & "GROUP BY FTSpoMemID,FDDateIns"
            tSql = tSql & vbCrLf & "ORDER BY FTSpoID"
            'Call SP_SQLvExecute(tSql, orsTmpSP) '*Tao 60-06-07 ของเดิม
            Call SP_SQLvExecute(tSql, orsTmpSP, IIf(bW_UseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(bW_UseServer, True, False)) '*Tao 60-06-07 Reprint หลังจาก Closed Teminal
            If bW_UseServer Then
                '*Tao 61-09-18 ถ้าหาที่ POSServer ไม่เจอ ให้ไปหาที่ Local DB ก่อน (เผื่อยัง Transfer ไปไม่ได้)
                If (orsTmpSP.BOF And orsTmpSP.EOF) Then Call SP_SQLvExecute(tSql, orsTmpSP, oDB.ocnOffLine)
            Else
                '*Tao 61-09-18 ถ้าหาที่ Local DB ไม่เจอ ให้ไปหาที่ POSServer ก่อน
                If (orsTmpSP.BOF And orsTmpSP.EOF) Then Call SP_SQLvExecute(tSql, orsTmpSP, oDB.ocnOnLine)
            End If
            If (orsTmpSP.BOF And orsTmpSP.EOF) Then GoTo ErrHandler
            orsTmpSP.MoveFirst
            Set oPrn.Font = oW_Font(1)
            Do While Not orsTmpSP.EOF
                
                tMemID = SP_FEDtChkString(orsTmpSP, "FTSpoMemID")
                nPoint = SP_FEDvChkNumeric(orsTmpSP, "FCSpoPoint")
                tExpDate = SP_DTEtStrSTD(SP_FEDdChkDate(orsTmpSP, "FDDateIns"))
                       
                If (nPoint > 0) Then  '*Tao 60-04-28 ถ้า Point = 0 ไม่ต้องพิมพ์
        '            tMemID = SP_FEDtChkString(orsTemp, "FTSpoMemID")
        '            nPoint = SP_FEDvChkNumeric(orsTemp, "FCSpoPoint")
                    tExpDate = Format(Month(CDate(tExpDate)), "00") & "/" & Right(Format(Year(CDate(tExpDate)), "0000"), 2)
                    tTmp = ""
                    'tName = ""
                    tName = "New Member" '*Bump 57-09-12 15.57pm. CommSheet TK-ISS3000-552.xlsx
                    nBalance = 0
                    Select Case Val(Format(tW_TransTypeSub))
                        Case tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                            tPointBefore = nBalance
                            tBonusPoints = (-1) * nPoint
                            tTotalBonus = nBalance
                        Case Else
                            tPointBefore = nBalance
                            tBonusPoints = nPoint
                            tTotalBonus = nPoint
                    End Select
                    
        '            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                    tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
        '            tMemID = W_GETtCreditNo(tMemID)                                     '*Ao 2013/09/23
        '            tTmp = tTkPointL1 & tMemID
                    tTmp = W_CHKtChangeCardID(tMemID, "0") '*Eaw 56-08-17 CommSheet TK-ISS3000-269-TR_Req.xls  ถ้าเป็นบัตร Memder จะส่ง "0" ไปแทน
                    
                    If tMemID = SP_PRCtGetMemPointID Then tTmp = tTmp & tVB_MallCardIDRef '*Tao 60-04-21 ถ้าได้ MallCard แบบ Manual ให้ใส่ * ต่อท้าย (TK-ISS3000-632)
                    
                    tSaveID = tSaveID & tTmp & "," & tExpDate & ";"                            '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
                    tTmp = tTkPointL1 & tTmp
        '            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                    tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                    tTmp = tTkPointL2 & " " & tName
        '            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                    tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                    '-------------------------------------------------------------------------------------
                    tSaveName = tSaveName & tName & ";"                              '*Ao 2013/11/30             ABCD;XYZ
                    tSavePoint = tSavePoint & tPointBefore & "," & tBonusPoints & "," & tTotalBonus & ";"
                    '-------------------------------------------------------------------------------------
                    If Mid(tW_OptAlw, 2, 1) = "1" Then
                        tTmp = tTkPointL3
        '                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                        tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                        tTmp = tPointBefore
        '                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                        tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                    End If
                    If Mid(tW_OptAlw, 3, 1) = "1" Then
                        tTmp = tTkPointL5
        '                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                        tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                        tTmp = tBonusPoints
        '                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                        tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                    End If
                    If Mid(tW_OptAlw, 2, 1) = "1" Or Mid(tW_OptAlw, 3, 1) = "1" Then
                        tTmp = tTkPointL7
        '                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                        tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                        tTmp = tTotalBonus
        '                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                        tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                    End If
                  
        '            tW_SlipFooterStr = tW_SlipFooterStr & "1" & tExpiry & " " & tExpDate & ";"          '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
                    If Trim(tExpiry) <> "" Then tW_SlipFooterStr = tW_SlipFooterStr & "1" & tExpiry & " " & tExpDate & ";" '*Bump 58-02-12 CommSheet TK-ISS3000-578.xls         version 4.2.165 12/02/2015
                End If
            orsTmpSP.MoveNext
            Loop
            '====================================
        End If
    End If
    
'    tW_SlipFooterStr = tW_SlipFooterStr & "1" & "   " & ";"
'    tW_SlipFooterStr = tW_SlipFooterStr & "*" & tBP1 & ";"          '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
'    tW_SlipFooterStr = tW_SlipFooterStr & "*" & tBP2 & ";"
    
    '*Bump 58-02-12 CommSheet TK-ISS3000-578.xls
    If Trim(tBP1) = "" And Trim(tBP2) = "" Then
        '* ไม่ต้องพิมพ์
    Else
            tW_SlipFooterStr = tW_SlipFooterStr & "1" & "   " & ";"
            If Trim(tBP1) <> "" Then tW_SlipFooterStr = tW_SlipFooterStr & "*" & tBP1 & ";"
            If Trim(tBP2) <> "" Then tW_SlipFooterStr = tW_SlipFooterStr & "*" & tBP2 & ";"
    End If
Call SP_DATxRsNothing(orsTmpSP)
Call SP_DATxRsNothing(orsTmpHD)
Exit Sub
ErrHandler:
Call SP_DATxRsNothing(orsTmpSP)
Call SP_DATxRsNothing(orsTmpHD)
End Sub

Private Function W_GETbGetTBSalePoint(ByRef porsTmp As ADODB.Recordset, ByVal ptTbName As String, ByVal ptWhere As String) As Boolean
'_________________________________________________
'   Call:
'   Cmt:    '*Ao 2013/07/04  CommSheet ML-PH4-R3-051 Bonus point missing .xls
'_________________________________________________
Dim tSql As String

On Error GoTo ErrHandler
    
    tSql = "SELECT * FROM " & ptTbName
    tSql = tSql & vbCrLf & "WHERE " & ptWhere
    If SP_SQLvExecute(tSql, porsTmp) = 0 Then
        'PASS
    Else
        Call SP_DATxRsNothing(porsTmp)
    End If
    W_GETbGetTBSalePoint = True
    Exit Function
ErrHandler:
    Call SP_DATxRsNothing(porsTmp)
    W_GETbGetTBSalePoint = True
End Function

Private Sub W_GETxBalRefDeposit()
'-------------------------------------------------------------------
'   Call:
'   Cmt: '*Eaw 56-06-03 CommSheet TK-ISS3000-198-ST 00 00 01.xls
'-------------------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim eTmp As eEN_Deposit
Dim cRndAmt As Double '*Eaw 56-07-20
On Error GoTo ErrorHandles
    Select Case SP_VALnText2Double(tW_TransTypeSub)  '*Eaw 56-06-18 แยกเคสกัน โดย Add Deposit ,AddDepositCancel ใช้ค่าใหม่ตาม CommSheet TK-ISS3000-231-ST 00 00 03.xls
        Case tEN_TrnAddDeposit
'            tSql = " WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "'"
'            tSql = tSql & vbCrLf & "AND FTShdStaCancel<> 'Y' AND FTShdTransType <> '20' AND FTShdTransNo <> '" & tW_TransNo & "'" '*Eaw 56-06-27
'            eW_Deposit.cEN_PevPayment = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCShdDepositAmt", "TPSTDepositHD", tSql, oDB.ocnOnLine, True))

            '*Ao 2014/07/21   V 4.2.138  TK-ISS3000PH2-532
            If nW_PrintCnt > 0 Or bW_PRMGenEJ Then
'                tSql = "WHERE FTCdpDepositNo = '" & eW_Deposit.tEN_DepositNo & "' AND FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tW_TransNo & "'"
'                eW_Deposit.cEN_PevPayment = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCCstCardAmt3", "TPSTDepositHD", tSql, oDB.ocnOnLine, True))
                eW_Deposit.cEN_PevPayment = eW_Deposit.cEN_Reprint
            Else
                tSql = " WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "'"
                tSql = tSql & vbCrLf & "AND FTShdStaCancel <> 'Y' AND FTShdTransType <> '20' AND FTShdTransNo <> '" & tW_TransNo & "'" '*Eaw 56-06-27
                eW_Deposit.cEN_PevPayment = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCShdDepositAmt", "TPSTDepositHD", tSql, oDB.ocnOnLine, True))
                '---------------------------------------------------------
                tSql = "UPDATE TPSTDepositHD"
                tSql = tSql & vbCrLf & "SET FCCstCardAmt3 = " & eW_Deposit.cEN_PevPayment               '*Ao 2014/07/21   V 4.2.138  TK-ISS3000PH2-532
                tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tW_TransNo & "'"
                Call SP_SQLvExecute(tSql)
                tSql = "UPDATE TPSTSalHD"
                tSql = tSql & vbCrLf & "SET FCCstCardAmt3 = " & eW_Deposit.cEN_PevPayment               '*Ao 2014/07/21   V 4.2.138  TK-ISS3000PH2-532
                tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tW_TransNo & "'"
                Call SP_SQLvExecute(tSql)
                tSql = "UPDATE " & tW_TblHD
                tSql = tSql & vbCrLf & "SET FCCstCardAmt3 = " & eW_Deposit.cEN_PevPayment               '*Ao 2014/07/21   V 4.2.138  TK-ISS3000PH2-532
                tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tW_TransNo & "'"
                Call SP_SQLvExecute(tSql)
            End If
            '---------------------------------------------------------
            tSql = "SELECT FCShdDepositBPay,FCShdDepositPrvBal,*  FROM TPSTDepositHD"
            tSql = tSql & vbCrLf & " WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tW_TransNo & "'"  '*Eaw 56-06-27
            Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                eW_Deposit.cEN_PrevBalDue = SP_FEDvChkNumeric(orsTemp, "FCShdDepositPrvBal") 'เงินที่ยังขาดอยู่ก่อนหน้า
                eW_Deposit.cEN_BalDue = SP_FEDvChkNumeric(orsTemp, "FCShdDepositBal") 'เงินมัดจำที่ยังขาดอยู่
            Else
                Call SP_SQLvExecute(tSql, orsTemp)
                eW_Deposit.cEN_PrevBalDue = SP_FEDvChkNumeric(orsTemp, "FCShdDepositPrvBal") 'เงินที่ยังขาดอยู่ก่อนหน้า
                eW_Deposit.cEN_BalDue = SP_FEDvChkNumeric(orsTemp, "FCShdDepositBal") 'เงินมัดจำที่ยังขาดอยู่
            End If
        Case tEN_TrnAddDepositCancel
            tSql = " WHERE (FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTShdTransType <> '20' AND FTShdStaCancel <> 'Y')"
            eW_Deposit.cEN_PevPayment = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCShdDepositAmt", "TPSTDepositHD", tSql, oDB.ocnOnLine, True))
'            tSql = " WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTShdTransNo = '" & tVB_Running & "'"
            tSql = " WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTShdTransNo = '" & tW_TransNo & "'"   '*Eaw 56-06-27
            eW_Deposit.cEN_PevPayment = eW_Deposit.cEN_PevPayment + SP_VALcText2Double(SP_SQLtFunction("SUM", "FCShdDepositAmt", "TPSTDepositHD", tSql))
'            tSql = " WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tVB_Running & "'"
            tSql = " WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tW_TransNo & "'"  '*Eaw 56-06-27
            eW_Deposit.cEN_PrevBalDue = SP_VALcText2Double(SP_SQLtFunction("", "FCShdDepositBal", "TPSTDepositHD", tSql))
            tSql = " WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTShdTransType = '06' "
            eW_Deposit.cEN_BalDue = SP_VALcText2Double(SP_SQLtFunction("", "FCShdDepositBal", "TPSTDepositHD", tSql, oDB.ocnOnLine, True))
        Case nEN_TrnDeposit
            If nW_PrintCnt = 0 Then
                tSql = "UPDATE TPSTDepositHD"
                tSql = tSql & vbCrLf & "SET FCCstCardAmt3 = " & eW_Deposit.cEN_BalDue               '*Ao 2014/07/21   V 4.2.138  TK-ISS3000PH2-532
                tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tW_TransNo & "'"
                Call SP_SQLvExecute(tSql)
                tSql = "UPDATE TPSTSalHD"
                tSql = tSql & vbCrLf & "SET FCCstCardAmt3 = " & eW_Deposit.cEN_BalDue
                tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tW_TransNo & "'"
                Call SP_SQLvExecute(tSql)
                tSql = "UPDATE " & tW_TblHD
                tSql = tSql & vbCrLf & "SET FCCstCardAmt3 = " & eW_Deposit.cEN_BalDue
                tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo= '" & eW_Deposit.tEN_DepositNo & "' AND FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tW_TransNo & "'"
                Call SP_SQLvExecute(tSql)
            End If
        Case tEN_TrnRedeem '*Eaw 56-07-20 CommSheet TK-ISS3000-287 TR-FAPL-067.xls
'            eW_Deposit.cEN_PevPayment = eW_Deposit.cEN_PevPayment - eW_Deposit.cEN_RndAmt  '*Eaw 56-07-26 เอามาผิดฟิลด์
    End Select
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrorHandles:
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGShowError
End Sub

Private Function W_CHKbDisCount(ByVal pnSeq As Integer) As Integer
'-------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 2013/03/07
'-------------------------------------------------------------------
Dim orsTemp As ADODB.Recordset
Dim tSql  As String
Dim nCount As Integer

On Error GoTo ErrorHandles
    
    tSql = "SELECT * FROM " & tW_TblCD
'    tSql = tSql & vbCrLf & "WHERE " & tW_TransWhere         '*Ao 2014/02/07
    tSql = tSql & vbCrLf & "WHERE " & tW_WhereCD
    tSql = tSql & vbCrLf & "AND FNDctNo IN (1,2,3,4,5,6,9,10,22)"
    tSql = tSql & vbCrLf & "AND FNSdtSeqNo = " & pnSeq
'    Call SP_SQLvExecute(tSql, orsTemp)
    Call SP_SQLvExecute(tSql, orsTemp, IIf(bW_UseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(bW_UseServer, True, False)) '*Eaw 56-07-24
    If orsTemp.BOF And orsTemp.EOF Then
        nCount = 0
    Else
        nCount = orsTemp.RecordCount
    End If
    W_CHKbDisCount = nCount
    Exit Function
ErrorHandles:
    W_CHKbDisCount = 0
End Function

Private Function W_CHKbCheckMemEmp(ByVal pbUseServer As Boolean) As Boolean
'-------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 2013/03/07
'   Cmt: '*Eaw 56-07-24 เพิ่ม parameter pbUseServer
'-------------------------------------------------------------------
Dim bChk As Boolean
Dim orsTemp As ADODB.Recordset
Dim tSql  As String

On Error GoTo ErrorHandles

    tSql = "SELECT * FROM " & tW_TblCD
'    tSql = tSql & vbCrLf & "WHERE " & tW_TransWhere tW_WhereCD '*BG 57-01-02
    tSql = tSql & vbCrLf & "WHERE " & tW_WhereCD
    tSql = tSql & vbCrLf & "AND FNDctNo IN (11,12)"     'Member Card,Employee Card
    tSql = tSql & vbCrLf & "AND FCScdAmt > 0"
'    Call SP_SQLvExecute(tSql, orsTemp)
    Call SP_SQLvExecute(tSql, orsTemp, IIf(pbUseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(pbUseServer, True, False)) '*Eaw 56-07-24
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        bChk = True
    Else
        bChk = False
    End If
    W_CHKbCheckMemEmp = bChk
    Exit Function
ErrorHandles:
    Call SP_MSGxShowError(Err.Number, "W_CHKxCheckMemEmp:" & Err.Description)
    W_CHKbCheckMemEmp = True
End Function

Private Sub W_PRNxTicketCashTr(ByVal porsDT As ADODB.Recordset)
'-------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 2013/03/07
'-------------------------------------------------------------------
Dim nLoop As Integer
Dim tTmp As String, tPdtName As String

Dim cRateCurrency As Double, cQty As Double, cNet As Double, cTotalAmt As Double
Dim nSpace As Integer, nSpaceEx As Integer
Dim orsFor As ADODB.Recordset
Dim tSql As String
Dim cTotalGVchAmt As Double
'*Tao 62-08-22  Vendor Voucher
Dim orsTemp As ADODB.Recordset '*Tao 62-08-22 Clone DT
Dim cTotalVVchAmt As Double
On Error GoTo ErrHandler
'*Tao 62-08-22  Vendor Voucher
If Not (porsDT.BOF And porsDT.EOF) Then Set orsTemp = porsDT.Clone

nSpace = 25
cTotalGVchAmt = 0   '*BG 56-09-03
tSql = "SELECT FTFcrCode,FTFcrSign,FTFcrType,FCFcrExchang FROM TCNMForCurrency" '*Ao 2013/05/03 CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
Call SP_SQLvExecute(tSql, orsFor, oDB.ocnOnLine, , , , , True, True) 'read

For nLoop = 0 To 3
    With porsDT
        .Filter = "FTSdtStaPdt='" & CStr(nLoop) & "'"
        .Sort = "FNSdtSeqNo ASC"
         
        If Not (porsDT.BOF And porsDT.EOF) Then
            Do While Not porsDT.EOF
                tPdtName = SP_FEDtChkString(porsDT, "FTSkuAbbName")  '*BG 56-09-02
                tTmp = ""
                If nLoop = 3 Then
                    cRateCurrency = SP_FEDvChkNumeric(porsDT, "FCSdtStkFac")
'                    Call oPrn.PrintText(Format(cRateCurrency, "#0.00000"), cXL, , vbLeftJustify, False)
                    tTmp = Format(cRateCurrency, "#0.00000")
                Else
'                    If nLoop <> 2 Then
'                        cQty = SP_FEDvChkNumeric(porsDT, "FCSdtQty")
'                        If Val(tW_TransTypeSub) = 37 Then
'                            tTmp = Format(cQty, "###0") & "X"
'                        Else
'                            tTmp = Format(cQty, "###0")
'                        End If
'                    End If
'*BG 56-09-02
                    cQty = SP_FEDvChkNumeric(porsDT, "FCSdtQty")                       '*Ao 2013/09/02
                    If Val(tW_TransTypeSub) = 37 Or Val(tW_TransTypeSub) = 41 Or Val(tW_TransTypeSub) = 46 Or Val(tW_TransTypeSub) = 12 Or Val(tW_TransTypeSub) = 24 Then
                        If UCase(tPdtName) <> UCase("Other Cash") And UCase(tPdtName) <> UCase("Other GVch") Then
                            tTmp = Format(cQty, "###0") & "X"
                        End If
                    Else
                        If UCase(tPdtName) <> UCase("Other Cash") And UCase(tPdtName) <> UCase("Other GVch") Then            '*Ao 2013/09/02
                            tTmp = Format(cQty, "###0")
                        End If
'                        tTmp = Format(cQty, "###0")
                    End If

                End If
'                tPdtName = SP_FEDtChkString(porsDT, "FTSkuAbbName") 'SKU Name  Abb Eng
'                     tTempName1 = tPdtName
                tPdtName = Mid(tPdtName, 1, 20) 'SKU Name  Abb Eng
'                     tPdtName2 = SP_FEDtChkString(porsDT, "FTSkuAbbNameSnd") 'SKU Name  Abb Thai
'                     tTempName2 = tPdtName2
'                     tPdtName2 = Mid(tPdtName2, 1, 20) 'SKU Name  Abb Thai
'                     Call oPrn.PrintText(tPdtName, cXC, , vbRightJustify, False)
                If nSpace > (Len(tTmp) + Len(tPdtName)) Then
                    nSpaceEx = nSpace - (Len(tTmp) + Len(tPdtName))
                Else
                    nSpaceEx = 0
                End If
                If nLoop = 3 Then
                    Call oPrn.PrintText(tPdtName & Space(nSpaceEx) & tTmp, cXL, , vbLeftJustify, False)     '*Ao 2013/05/03  CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
                Else
                    Call oPrn.PrintText(tTmp & Space(nSpaceEx) & tPdtName, cXL, , vbLeftJustify, False)
                End If
'                Call oPrn.PrintText(tTmp & Space(nSpaceEx) & tPdtName, cXL, , vbLeftJustify, False)
                
                If nLoop = 3 Then
                    cNet = SP_FEDvChkNumeric(porsDT, "FCSdtRegPrice")
                    '----------------------------------------------------------------------------------------------
                    tTmp = SP_FEDtChkString(porsDT, "FTSkuCode")        '*Ao 2013/05/03  CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
                    tTmp = Trim(tTmp)
                    orsFor.Filter = "FTFcrCode='" & tTmp & "'"
                    If Not (orsFor.BOF And orsFor.EOF) Then
                        tTmp = SP_FEDtChkString(orsFor, "FTFcrSign")
                        tTmp = Trim(tTmp)
                    Else
                        tTmp = ""
                    End If
                    orsFor.Filter = adFilterNone
                    '----------------------------------------------------------------------------------------------
                Else
                    cNet = SP_FEDvChkNumeric(porsDT, "FCSdtSaleAmt")
                    tTmp = ""
                End If
                
                Call oPrn.PrintText(tTmp & Format(cNet, tW_CurFmt), cXMK, , vbRightJustify, True)
'                Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXMK, , vbRightJustify, True)
'                Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                If nLoop = 0 Or nLoop = 1 Then 'Coin + Curreny = Cash Total
                    cTotalAmt = cTotalAmt + cNet
                End If
                '*BG 56-09-02
                If nLoop = 2 Then
                    cTotalGVchAmt = cTotalGVchAmt + cNet
                End If
                '****
                .MoveNext
            Loop
        End If
        
        If nLoop = 1 Then
            tTmp = "TOTAL CASH"
            nSpaceEx = nSpace - Len(tTmp)
            tTmp = Space(nSpaceEx) & tTmp
            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
            Call oPrn.PrintText(Format(cTotalAmt, tW_CurFmt), cXMK, , vbRightJustify, True)
'            Call oPrn.PrintText(Format(cTotalAmt, tW_CurFmt), cXR, , vbRightJustify, True)
        End If
'*BG 56-09-02
        If nLoop = 2 Then                                                          '*Ao 2013/09/02
            tTmp = "TOTAL GVch"
            nSpaceEx = nSpace - Len(tTmp)
            tTmp = Space(nSpaceEx) & tTmp
            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
            Call oPrn.PrintText(Format(cTotalGVchAmt, tW_CurFmt), cXMK, , vbRightJustify, True)
            
            '*Tao 62-08-22  Print Vendor Voucher ต่อท้าย TOTAL GVch
            If Not (orsTemp Is Nothing) Then
                orsTemp.Filter = "FTSkuCode='T006' AND FTSdtStaPdt='6'" '6=ที่ป้อนข้อมูลไว้
                orsTemp.Sort = "FNSdtSeqNo ASC"
                If Not (orsTemp.BOF And orsTemp.EOF) Then
                    tTmp = "TOTAL VVch"
                    nSpaceEx = nSpace - Len(tTmp)
                    tTmp = Space(nSpaceEx) & tTmp
                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                    cTotalVVchAmt = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt")
                    Call oPrn.PrintText(Format(cTotalVVchAmt, tW_CurFmt), cXMK, , vbRightJustify, True)
                End If
                Call SP_DATxRsNothing(orsTemp)
            End If
            '================================================
            
        End If
'****
    End With
Next nLoop

Exit Sub
ErrHandler:

End Sub

Private Function W_CHKbCheckCopy(ByVal ptSource As String, ByVal ptDestination As String) As Boolean
'-------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 2013/03/07
'-------------------------------------------------------------------
Dim oFSO As New FileSystemObject
Dim bChk As Boolean

On Error GoTo ErrHandler
    bChk = False
    oFSO.CopyFile ptSource, ptDestination
    bChk = True
    Set oFSO = Nothing
    W_CHKbCheckCopy = bChk
    Exit Function
ErrHandler:
    W_CHKbCheckCopy = False
    Set oFSO = Nothing
End Function

Private Function W_GETtSkuSign(ByVal ptStaMall As String, ByVal ptStaEmp As String, ByVal ptStaMallEmp As String) As String
'-------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 2013/03/07
'-------------------------------------------------------------------------------------------------------------------------
Dim tSql As String, tSign As String
Dim orsTemp As ADODB.Recordset

On Error GoTo ErrHandler

    If bW_ChkMark Then
        tSql = "SELECT TOP 1.* FROM " & tW_TblCD
'        tSql = tSql & vbCrLf & "WHERE " & tW_TransWhere        '*Ao 2014/02/07
        tSql = tSql & vbCrLf & "WHERE " & tW_WhereCD
        tSql = tSql & vbCrLf & "AND FNDctNo = 11"               'Member Discount
        If SP_SQLvExecute(tSql, orsTemp) = 0 Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                bW_ChkMEM = True
            Else
                bW_ChkMEM = False
            End If
        End If
        If bW_ChkMEM Then
        
        Else
            tSql = "SELECT TOP 1.* FROM " & tW_TblCD
'            tSql = tSql & vbCrLf & "WHERE " & tW_TransWhere        '*Ao 2014/02/07
            tSql = tSql & vbCrLf & "WHERE " & tW_WhereCD
            tSql = tSql & vbCrLf & "AND FNDctNo = 12"           'Employee Discount
            If SP_SQLvExecute(tSql, orsTemp) = 0 Then
                If Not (orsTemp.BOF And orsTemp.EOF) Then
                    bW_ChkEmp = True
                Else
                    bW_ChkEmp = False
                End If
            End If
        End If
        '-------------------------------------------------
        bW_ChkMark = False
        '-------------------------------------------------
    End If
    
    If bW_ChkMEM Then
        If UCase(ptStaMallEmp) = UCase("Y") Then
            tSign = ""
        Else
            tSign = "*"                                                                         'ไม่ให้ส่วนลดบัตรสมาชิก( Member Discount)
        End If
    Else
        If bW_ChkEmp Then
            If UCase(ptStaMallEmp) = UCase("Y") Then
                tSign = ""
            Else
                tSign = "#"                                                                     'ไม่ให้ส่วนลดพนักงาน (Employee Discount)
            End If
        Else
            If UCase(ptStaMall) = UCase("N") Then
                tSign = "*"
            Else
                If UCase(ptStaEmp) = UCase("N") Then
                    tSign = "#"
                Else
                    tSign = ""
                End If
            End If
        End If
    End If
    
    W_GETtSkuSign = tSign
    Exit Function
ErrHandler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    W_GETtSkuSign = ""
End Function

'Private Function W_PRNtSalesManID(ByVal ptSalesManID As String) As String
''-------------------------------------------------------------------------------------------------------------------------
''   Call:
''   Cmt:    '*Ao 2013/03/06
''-------------------------------------------------------------------------------------------------------------------------
'Dim tSalesManID As String
'
'On Error GoTo ErrHandler
'
'    Select Case Len(ptSalesManID)
'        Case 1 To 5
'            tSalesManID = ptSalesManID
'        Case 6 To 10
'            tSalesManID = Left(ptSalesManID, 5) & "/" & Right(ptSalesManID, Len(ptSalesManID) - 5)
'        Case 11 To 15
'            tSalesManID = Left(ptSalesManID, 5) & "/" & Mid(ptSalesManID, 6, 5) & "/" & Right(ptSalesManID, Len(ptSalesManID) - 10)
'        Case 16 To 20
'            tSalesManID = Left(ptSalesManID, 5) & "/" & Mid(ptSalesManID, 6, 5) & "/" & Mid(ptSalesManID, 11, 5) & "/" & Right(ptSalesManID, Len(ptSalesManID) - 15)
'        Case Else
'            tSalesManID = ptSalesManID
'    End Select
'
'    W_PRNtSalesManID = tSalesManID
'    Exit Function
'ErrHandler:
'    W_PRNtSalesManID = ptSalesManID
'End Function

Private Function W_PRNbTicketLOGO() As Boolean       '*Ao 2013/01/24   Print FRONTTECH
Dim nLoop As Integer

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
    
    If bVB_ImageLogo Then
'        MsgBox "bVB_ImageLogo =True"
        Call oPrn.PrintLogo
'    Else
'        MsgBox "bVB_ImageLogo = False"
'        Call oPrn.PrintLogo
    End If
    W_PRNbTicketLOGO = True
Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
    W_PRNbTicketLOGO = True
End Function

Private Function W_PRNbTicketFooterBlank() As Boolean       '*Ao 2013/01/24   Print FRONTTECH
Dim nLoop As Integer
Dim tTmp As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler

    For nLoop = 0 To 3
'        Call oPrn.PrintText("", cXC, , vbCenter, True, nEN_PrinterOnly)
        tTmp = " "
        Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True, nEN_PrinterOnly)
    Next nLoop
    W_PRNbTicketFooterBlank = True
Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
    W_PRNbTicketFooterBlank = True
End Function

Private Function W_GETbCVIssue(ByVal ptTransType As String, ByVal ptTmnNo As String, ByVal ptTransNo As String) As Boolean
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
On Error GoTo ErrHandle
    W_GETbCVIssue = False
    tSql = "SELECT FTCVNum,FCCVAmt,FCCVBalAmt,FDCVExpDate,FTCustContact" '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls เพิ่ม FTCustContact
    tSql = tSql & vbCrLf & " FROM TPSTCV"
    tSql = tSql & vbCrLf & " WHERE FTTransType ='" & ptTransType & "'"
    tSql = tSql & vbCrLf & " AND FTPOSNo ='" & ptTmnNo & "'"
    tSql = tSql & vbCrLf & " AND FTTransNo='" & ptTransNo & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            W_GETbCVIssue = True
            'TK-ISS3000-064
            If bW_EnableRePrint Then
                oW_ReIssue.cEN_ReIssueAmt = SP_FEDvChkNumeric(orsTemp, "FCCVAmt")
                oW_ReIssue.tEN_ReIssueExp = CStr(Format(SP_FEDdChkDate(orsTemp, "FDCVExpDate"), "MM/yy"))
                oW_ReIssue.tEN_ReIssueNo = SP_FEDtChkString(orsTemp, "FTCVNum")
                oW_ReIssue.tEN_ReIssueTel = SP_FEDtChkString(orsTemp, "FTCustContact") '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
            End If
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function W_GETbCVIssueEx(ByVal ptCVNum As String) As Boolean
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
On Error GoTo ErrHandle
    W_GETbCVIssueEx = False
    tSql = "SELECT FTCVNum,FCCVAmt,FCCVBalAmt,FDCVExpDate,FTCustContact" '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls เพิ่ม FTCustContact
    tSql = tSql & vbCrLf & " FROM TPSTCV"
    tSql = tSql & vbCrLf & " WHERE FTCVNum ='" & ptCVNum & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then
'     If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then  '*Eaw 56-06-20 ให้ออก CV ขณะที่ระบบ offline ได้
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            W_GETbCVIssueEx = True
            'TK-ISS3000-064
                oW_CV.cEN_CVAmt = SP_FEDvChkNumeric(orsTemp, "FCCVAmt")
                oW_CV.tEN_CVExp = CStr(Format(SP_FEDdChkDate(orsTemp, "FDCVExpDate"), "MM/yy"))
                oW_CV.tEN_CVNo = SP_FEDtChkString(orsTemp, "FTCVNum")
                oW_CV.tEN_CVTel = SP_FEDtChkString(orsTemp, "FTCustContact") '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Sub W_GETxReprintCV(ByVal ptTransType As String, ByVal ptTmnNo As String, ByVal ptTransNo As String)
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
On Error GoTo ErrHandle
    oW_ReIssue.cEN_ReIssueAmt = 0
    oW_ReIssue.tEN_ReIssueExp = ""
    oW_ReIssue.tEN_ReIssueNo = ""
    oW_ReIssue.tEN_ReIssueTel = "" '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
    tSql = "SELECT FCCVAmt,FDCVExpDate,FTCVNum ,FTCustContact " '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls เพิ่ม FTCustContact
    tSql = tSql & vbCrLf & " FROM TPSTCV"
    tSql = tSql & vbCrLf & " WHERE FTTransType ='" & ptTransType & "'"
    tSql = tSql & vbCrLf & " AND FTPOSNo ='" & ptTmnNo & "'"
    tSql = tSql & vbCrLf & " AND FTTransNo='" & ptTransNo & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then
'    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then  '*Eaw 56-06-20 ให้ออก CV ขณะที่ระบบ offline ได้
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            oW_ReIssue.cEN_ReIssueAmt = SP_FEDvChkNumeric(orsTemp, "FCCVAmt")
            oW_ReIssue.tEN_ReIssueExp = CStr(Format(SP_FEDdChkDate(orsTemp, "FDCVExpDate"), "MM/yy"))
            oW_ReIssue.tEN_ReIssueNo = SP_FEDtChkString(orsTemp, "FTCVNum")
            oW_ReIssue.tEN_ReIssueTel = SP_FEDtChkString(orsTemp, "FTCustContact") '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
End Sub

Public Function W_PRNbTicketCVSlp(ByVal pnType As Long, ByVal porsRC As ADODB.Recordset)
    Dim tTmp As String
    Dim tCardExp As String
    Dim tCardNo As String
    Dim cRcvAmt As Double
    Dim tRcvTel As String '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
On Error GoTo ErrHandler
    'TK-ISS3000-069
    Select Case pnType
        Case 2
            'orsRC.Filter = "FTTdmCode='T018' AND FTSrcCardNo = '" & ptCardNo & "'"
            If Not (porsRC.BOF And porsRC.EOF) Then
                cRcvAmt = SP_FEDtChkString(porsRC, "FCSrcNet")
                tCardNo = SP_FEDtChkString(porsRC, "FTSrcCardNo")
                tRcvTel = SP_FEDtChkString(porsRC, "FTSrcRetDocRef") '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
                Call oPrn.PrintText("", cXL, , vbLeftJustify, True) '*BG 56-06-05 '*BG 56-06-05 CommSheet TK-ISS3000-221-ST 00 00 01.xls
                Call oPrn.PrintText("", cXL, , vbLeftJustify, True) '*BG 56-06-05 '*BG 56-06-05 CommSheet TK-ISS3000-221-ST 00 00 01.xls
                tTmp = "Cr Vchr AMT"
                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                tTmp = Format(cRcvAmt, tW_CurFmt)
                Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True)
                'Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                Call oPrn.PrintText("", cXL, , vbLeftJustify, True)
                tTmp = "CREDIT VCR"
                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                tTmp = "ACCOUNT #  " & "     " & tCardNo '*Eaw 56-08-07 ขยับช่องว่างเพิ่มอีก 3 ช่อง
                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                tCardExp = SP_FEDtChkString(porsRC, "FTSrcCardExp")
                tTmp = "EXPIRY DATE" & "     " & tCardExp '*Eaw 56-08-07 ขยับช่องว่างเพิ่มอีก 3 ช่อง
                 Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
'                tTmp = "CONTACT NUMBER" & "  " & tRcvTel
'                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
            End If
            'orsRC.Filter = adFilterNone
        Case Else
            If tW_TransTypeSub = "16" And oW_ReIssue.tEN_ReIssueNo <> "" And oW_ReIssue.cEN_ReIssueAmt > 0 Then bW_StaPrnCVReIssue = True '*Thirapong 2021-04-30 CommSheet TTK-ISS3000-787 TK ISSUE -  Credit Voucher is printed with BLANK content.xlsx
            If bW_StaPrnCVReIssue And oW_ReIssue.tEN_ReIssueNo <> "" Then '*Thirapong 2021-03-27 bW_StaPrnCVReIssue
                cRcvAmt = oW_ReIssue.cEN_ReIssueAmt
                tCardNo = oW_ReIssue.tEN_ReIssueNo
                tRcvTel = oW_ReIssue.tEN_ReIssueTel  '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
                Call oPrn.PrintText("", cXL, , vbLeftJustify, True) '*BG 56-06-05 CommSheet TK-ISS3000-221-ST 00 00 01.xls
                Call oPrn.PrintText("", cXL, , vbLeftJustify, True) '*BG 56-06-05 CommSheet TK-ISS3000-221-ST 00 00 01.xls
                tTmp = "Cr Vchr AMT"
                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                tTmp = Format(cRcvAmt, tW_CurFmt)
                Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True)
                'Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                Call oPrn.PrintText("", cXL, , vbLeftJustify, True)
                tTmp = "CREDIT VCR"
                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                tTmp = "ACCOUNT #  " & "     " & tCardNo '*Eaw 56-08-07 ขยับช่องว่างเพิ่มอีก 3 ช่อง
                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                tCardExp = oW_ReIssue.tEN_ReIssueExp
                tTmp = "EXPIRY DATE" & "     " & tCardExp '*Eaw 56-08-07 ขยับช่องว่างเพิ่มอีก 3 ช่อง
                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
            End If
            Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ":W_PRNbTicketCVSlp :tEN_ReIssueNo :" & oW_ReIssue.tEN_ReIssueNo & ",cEN_ReIssueAmt :" & oW_ReIssue.cEN_ReIssueAmt & ",bW_StaPrnCV2:" & bW_StaPrnCV2)
            '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
'            tTmp = "CONTACT NUMBER" & "  " & tRcvTel
'            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
    End Select
    W_PRNbTicketCVSlp = True
    Exit Function
ErrHandler:
    W_PRNbTicketCVSlp = False
End Function

Public Function W_PRNbTicketForeignCurrency() As Boolean
'______________________________________________________
'   Call:
'   Cmt: Aoo 2012/07/25
'______________________________________________________
Dim tSql As String
Dim tFcrType As String
Dim tExchang As String
Dim orsTmp As ADODB.Recordset
        
    On Error GoTo ErrHandler
        W_PRNbTicketForeignCurrency = True
        Select Case Val(tW_TransType)
            Case 1 'OPEN
                tSql = "SELECT * FROM TCNMForCurrency"
                If SP_SQLvExecute(tSql, orsTmp, oDB.ocnOnLine, , , , , True, True) = 0 Then
                    If Not (orsTmp.BOF And orsTmp.EOF) Then
                        Set oPrn.Font = oW_Font(1)
                        Do While Not orsTmp.EOF
                            tFcrType = Trim(SP_FEDtChkString(orsTmp, "FTFcrType"))                      '*Ao 2013/03/14  TK-ISS3000-094
                            tExchang = SP_FMTtNum2ShwCurrency(SP_FEDvChkNumeric(orsTmp, "FCFcrExchang"))
                            
                            Call oPrn.PrintText(tFcrType, cXL, , vbLeftJustify, False)                  'US
                            Call oPrn.PrintText(tExchang, cXMK, , vbRightJustify, True)            '30
                            orsTmp.MoveNext
                        Loop
                                      
                    End If
                End If
        End Select
        Exit Function
ErrHandler:
                     
End Function

Public Function OpenTransaction(ByVal pnTransType As Long, pdTransDate As Date, ptTransTmn As String, ptTransNo As String, Optional ByVal pbUseTmpTbl As Boolean, Optional ByVal ptTransPrint As String, Optional ByVal pbUseServer As Boolean = False) As Boolean
'*KT 51-11-03
        Dim tSql As String
        Dim nTmp As Long
        Dim tTmp As String, tDate As String
        Dim orsTmp As ADODB.Recordset
        Dim nError As Long
        Static tLastTmn As String
        Dim tChkAdaDB As String  '*BG 57-01-02
On Error GoTo ErrHandler
        
        nW_StateSuccess = 0
        nW_StateDoc = 0
        
        Select Case pnTransType
        Case nEN_TrnDeposit, nEN_TrnAddDeposit
                If pbUseTmpTbl Then
                        tW_TblHD = "TPSHD" & ptTransTmn
                        tW_TblDT = "TPSDT" & ptTransTmn
                        tW_TblCD = "TPSCD" & ptTransTmn
                        tW_TblRC = "TPSRC" & ptTransTmn
                Else
                        tW_TblHD = "TPSTSALHD"
                        tW_TblDT = "TPSTSALDT"
                        tW_TblCD = "TPSTSALCD"
                        tW_TblRC = "TPSTSALRC"
                End If
        Case nEN_TrnVoucher
                If pbUseTmpTbl Then
                        tW_TblHD = "TPSHD" & ptTransTmn
                        tW_TblDT = "TPSDT" & ptTransTmn
                        tW_TblCD = "TPSCD" & ptTransTmn
                        tW_TblRC = "TPSRC" & ptTransTmn
                Else
                        tW_TblHD = "TPSTSALHD"
                        tW_TblDT = "TPSTSALDT"
                        tW_TblCD = "TPSTSALCD"
                        tW_TblRC = "TPSTSALRC"
                End If
        Case Else
                If pbUseTmpTbl Then
                        tW_TblHD = "TPSHD" & ptTransTmn
                        tW_TblDT = "TPSDT" & ptTransTmn
                        tW_TblCD = "TPSCD" & ptTransTmn
                        tW_TblRC = "TPSRC" & ptTransTmn
                Else
                        tW_TblHD = "TPSTSALHD"
                        tW_TblDT = "TPSTSALDT"
                        tW_TblCD = "TPSTSALCD"
                        tW_TblRC = "TPSTSALRC"
                End If
        End Select
    
        If nW_StateCon <> 1 Then Exit Function
        
        tDate = SP_DTEtStrSTD(pdTransDate)
        tW_TransDate = tDate
        '*BG 57-01-02
'        tW_TransWhere = "FDShdTransDate = '" & tDate & "' AND FTTmnNum ='" & ptTransTmn & "' AND  FTShdTransNo='" & ptTransNo & "'"
'        tW_TransWhere = tW_TransWhere & " AND FTShdTransType = '" & Format(pnTransType, "00") & "'"
        Select Case Val(pnTransType)
            Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax, tEN_TrnTotalReportManualTax, tEN_TrnOverShortReportManualTax, tEN_TrnClosingReportManualTax, tEN_TrnOpenManualTax
                If UCase(SP_CHKtAdaConfig) = UCase("1") Then
                    tW_TransWhere = "FDShdTransDate = '" & tDate & "' AND FTOrgPhyPosNo ='" & ptTransTmn & "' AND  FTShdTransNo='" & ptTransNo & "'"
                    tW_TransWhere = tW_TransWhere & " AND FTShdTransType = '" & Format(pnTransType, "00") & "'"
                Else
                    tW_TransWhere = "FDShdTransDate = '" & tDate & "' AND FTTmnNum ='" & ptTransTmn & "' AND  FTShdTransNo='" & ptTransNo & "'"
                    tW_TransWhere = tW_TransWhere & " AND FTShdTransType = '" & Format(pnTransType, "00") & "'"
                End If
            Case Else
                tW_TransWhere = "FDShdTransDate = '" & tDate & "' AND FTTmnNum ='" & ptTransTmn & "' AND  FTShdTransNo='" & ptTransNo & "'"
                tW_TransWhere = tW_TransWhere & " AND FTShdTransType = '" & Format(pnTransType, "00") & "'"
        End Select
        tW_WhereCD = "FDShdTransDate = '" & tDate & "' AND FTTmnNum ='" & ptTransTmn & "' AND  FTShdTransNo='" & ptTransNo & "'"
        tW_WhereCD = tW_WhereCD & " AND FTShdTransType = '" & Format(pnTransType, "00") & "'"
        '****
        tTmp = tW_TransWhere
        bW_UseServer = pbUseServer '*Eaw 56-07-24
' -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -
        Set orsHD = New ADODB.Recordset
        Set orsDT = New ADODB.Recordset
        Set orsRC = New ADODB.Recordset
        Set orsCD = New ADODB.Recordset
        Set orsPmtCpn = New ADODB.Recordset
        Set orsCV = New ADODB.Recordset
   
'        tSql = "SELECT *  FROM " & tW_TblHD & " WHERE " & tW_TransWhere & " AND FTShdStaDoc = '1'"
        tSql = "SELECT *  FROM " & tW_TblHD & " WITH (NOLOCK) WHERE " & tW_TransWhere         '*Ao 2013/04/30 CommSheet TK-ISS3000-163-IT2 00.00.01.xls
        nError = SP_SQLvExecute(tSql, orsTmp, IIf(pbUseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(pbUseServer, True, False))
        If nError = 0 Then 'Local Only
            If orsTmp.BOF And orsTmp.EOF Then
                '*KT 53-01-20 ทำการเก็บ Log กรณี PH3-ML-SF-051 กรณีที่ Select แล้วไม่มีข้อมูล
'                Call SP_CRTxLogPrg("ErrorNumber:" & nError & " ErrorDesc:No data" & " Form:" & "cTransaction" & " Events:OpenTransaction[1]")
'                Call SP_CRTxLogPrg("Query on " & IIf(pbUseServer, "server", "local") & " : " & tSql)
                Call SP_CRTxWriteLogTerminated("ErrorNumber:" & nError & " ErrorDesc:No data" & " Form:" & "cTransaction" & " Events:OpenTransaction[1]")
                Call SP_CRTxWriteLogTerminated("Query on " & IIf(pbUseServer, "server", "local") & " : " & tSql)
                Exit Function
            Else
                Set orsHD = orsTmp
            End If
        Else
              '*KT 53-01-20 ทำการเก็บ Log กรณี PH3-ML-SF-051กรณีที่ SP_SQLvExecute เกิด Error เช่น Time Out
'            Call SP_CRTxLogPrg("ErrorNumber:" & nError & " ErrorDesc:Execute Error" & " Form:" & "cTransaction" & " Events:OpenTransaction[2]")
'            Call SP_CRTxLogPrg("Query on " & IIf(pbUseServer, "server", "local") & " : " & tSql)
            Call SP_CRTxWriteLogTerminated("ErrorNumber:" & nError & " ErrorDesc:Execute Error" & " Form:" & "cTransaction" & " Events:OpenTransaction[2]")
            Call SP_CRTxWriteLogTerminated("Query on " & IIf(pbUseServer, "server", "local") & " : " & tSql)
            Exit Function
        End If
        
        nW_StateSuccess = 1 '*KT 53-03-02 PH3-ML-SF-051  No EJ, Dat File 1 kb, No data on Database but it's Sale Transaction
        
        If tW_TblDT <> vbNullString Then
                tSql = "SELECT * FROM " & tW_TblDT & " WITH (NOLOCK) WHERE " & tW_TransWhere & " ORDER BY FNSdtSeqNo ASC"
                Call SP_SQLvExecute(tSql, orsDT, IIf(pbUseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(pbUseServer, True, False)) 'Local Only
        End If

        If tW_TblRC <> vbNullString Then
                tSql = "SELECT * FROM " & tW_TblRC & " WITH (NOLOCK) WHERE " & tW_TransWhere & " ORDER BY FNSrcSeqNo ASC"
                Call SP_SQLvExecute(tSql, orsRC, IIf(pbUseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(pbUseServer, True, False)) 'Local Only
        End If
        
        If tW_TblCD <> vbNullString Then
'                tSql = "SELECT * FROM " & tW_TblCD & " WHERE " & tW_TransWhere & " ORDER BY FNScdSeqNo ASC" '*BG 57-01-02
                tSql = "SELECT * FROM " & tW_TblCD & " WITH (NOLOCK) WHERE " & tW_WhereCD & " ORDER BY FNScdSeqNo ASC"
                Call SP_SQLvExecute(tSql, orsCD, IIf(pbUseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(pbUseServer, True, False)) 'Local Only
        End If
        '*KT 54-10-06 V.4.0.1Req.6 Phase4.0(Member Point) Promotion Coupon ขอทำการแก้ไขตาม  DCR_5410001
'        tSql = "SELECT * FROM TPSTSalPmtCpn WHERE " & tW_TransWhere & " ORDER BY FNPcnSeqNo ASC" '*BG 57-02-03
        tSql = "SELECT * FROM TPSTSalPmtCpn WHERE " & tW_WhereCD & " ORDER BY FNPcnSeqNo ASC"
        Call SP_SQLvExecute(tSql, orsPmtCpn, IIf(pbUseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(pbUseServer, True, False)) 'Local Only
  
         tSql = "SELECT * FROM TPSTCV WITH(NOLOCK)"
         tSql = tSql & vbCrLf & " WHERE  FTCVNum='" & SP_FEDtChkString(orsHD, "FTCVNum") & "'"
         Call SP_SQLvExecute(tSql, orsCV, IIf(pbUseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(pbUseServer, True, False)) 'Local Only
'get transaction property -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -
            
            'key
            dW_TransDate = pdTransDate
            'dW_EJDate = pdTransDate             '*Ao 2013/10/01       ReEntry  ให้เก็บ EJ ตาม Operation Date
            '*start TTK-ISS3000-762
            '*Thirapong 63-08-03 CommSheet TTK-ISS3000-762 TK ISSUE - ManageEJ.exe printed wrong credit card number.xlsx
            '* เพิ่มการเช็ค If bW_PRMGenEJ Then
            'dW_EJDate = tVB_EJDate '*Tao 59-04-19 ReEntry  ให้เก็บ EJ ตาม Sale Date (CommSheet TK-ISS3000-610)
             If bW_PRMGenEJ Then
                Select Case Val(tW_TransType)
                    Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax, tEN_TrnTotalReportManualTax, tEN_TrnOverShortReportManualTax, tEN_TrnClosingReportManualTax, tEN_TrnOpenManualTax
                        If SP_FEDtChkString(orsHD, "FTCstCardCode3") = "" Then
                            tVB_EJDate = UT_tStringDate(pdTransDate)
                        Else
                            tVB_EJDate = SP_FEDtChkString(orsHD, "FTCstCardCode3")
                        End If
                        dW_EJDate = tVB_EJDate
                Case Else
                End Select
            End If
            'end TTK-ISS3000-762
'            tW_TmnNum = ptTransTmn
            Select Case Val(pnTransType)               '*Ao 2014/03/19   CommSheet TK-ISS3000-461-00.00.03.xls
                Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax, tEN_TrnTotalReportManualTax, tEN_TrnOverShortReportManualTax, tEN_TrnClosingReportManualTax, tEN_TrnOpenManualTax
                    tChkAdaDB = SP_CHKtAdaConfig  '43,44,45,46,47,49
                    If UCase(tChkAdaDB) = UCase("1") Then                                       ' 1= Cluster ,   0 = POS Server
                        tW_TmnNum = SP_FEDtChkString(orsHD, "FTOrgPhyPosNo")
                    Else
                        tW_TmnNum = ptTransTmn
                    End If
                Case Else
                    tW_TmnNum = ptTransTmn
            End Select
            
            tW_TransNo = ptTransNo
            tW_Replicated = SP_FEDtChkString(orsHD, "FTShdDisChgTxt")            '*Ao 2014/05/06
            'trans
            tW_DepID = SP_FEDtChkString(orsHD, "FTShdDepID")
            tW_StoreID = SP_FEDtChkString(orsHD, "FTStmCode")
            tW_TransType = SP_FEDtChkString(orsHD, "FTShdTransType")
            tW_TransTypeSub = SP_FEDtChkString(orsHD, "FTShdTransTypeSub")
            If ptTransPrint = "" Then
                tW_TransNameABB = SP_FEDtChkString(orsHD, "FTShdTransNameAbb")
            'cVB_CNGSTRate
                cW_HDRate = SP_FEDvChkNumeric(orsHD, "FCShdVATRate")
                '*KT 66-08-03 TTK-ISS3000-829
               ' If Mid(tW_TransNameABB, 8, 1) = cVB_OldGstRate Then '*Amorn 65-07-12 Reprint 7%
                If tW_TransTypeSub = "04" And InStr(1, tW_TransNameABB, CStr(cW_HDRate)) > 0 Then
'               If Mid(tW_TransNameABB, 8, 1) <> cVB_CNGSTRate Then '*Amorn 2023-02-20 TTK-ISS3000-818 TK ISSUE -Pickup layaway Vat 9_ GST show wording wrong 7_ GST
                    bVB_ChkRetGST7 = True
                    cVB_OldGstRate = cW_HDRate
                ElseIf tW_TransTypeSub = "07" Or tW_TransTypeSub = "16" Then
                    cVB_OldGstRate = cW_HDRate
                End If
            Else
                tW_TransNameABB = ptTransPrint
            End If
            tW_TransName = W_DATtGetDetail(tW_TransType, "FTSttTranCode", "TSysTransType", "FTSttDesc")
            tW_TransTime = SP_FEDtChkString(orsHD, "FTShdTransTime")
            dW_SysDate = SP_FEDdChkDateTime(orsHD, "FDShdSysDate")
            tW_SysTime = SP_FEDtChkString(orsHD, "FTShdSysTime")
            
            'mode
            bW_TrainMode = (SP_FEDtChkString(orsHD, "FTShdTrainMode") = "1")
            bW_RecallMode = (SP_FEDtChkString(orsHD, "FTShdRecallMode") = "1")
            tW_AccTypeReEntry = SP_FEDtChkString(orsHD, "FTShdStaJob") '*Eaw 56-08-29 Acc Mode for Re Entry
            tW_PrintType = SP_FEDvChkNumeric(orsHD, "FNLogStaExport") '*Bump 57-04-05 CommSheet TK-ISS3000-461-00.00.04.xls
            If tW_PrintType = "" Then tW_PrintType = "1"
'            เก็บค่า Print Type ที่บิลนั้นพิมพ์ออกมา  '0' คือ Slip printer                           '1' คือ Thermal printer
            '*******************
            'stroe
            tW_StoreNum = SP_FEDtChkString(orsHD, "FTStmCode")
'สาเหตุที่ต้องอ่านข้อมุลทุกครั้งเพราะว่า Generate EJ ใช้งาน Procedure นี้ด้วย
'            tW_StoreNameABB = W_DATtGetDetail(tW_StoreNum, "FTStmCode", "TCNMStoreMtn", "FTStmNameAbb")
'            tW_StoreNameTax = W_DATtGetDetail(tW_StoreNum, "FTStmCode", "TCNMStoreMtn", "FTStmNameTax")
'            tW_TaxInvoice = W_DATtGetDetail(tW_StoreNum, "FTStmCode", "TCNMStoreMtn", "FTStmTaxID")
'            tW_StoreTel = W_DATtGetDetail(tW_StoreNum, "FTStmCode", "TCNMStoreMtn", "FTStmTel")
            tSql = "SELECT FTStmNameAbb,FTStmNameTax,FTStmTaxID,FTStmTel"
            tSql = tSql & vbCrLf & "FROM TCNMStoreMtn WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTStmCode = '" & tW_StoreNum & "'"
            If SP_SQLvExecute(tSql, orsTmp, oDB.ocnOnLine, , , , , True, True) = 0 Then
                If Not (orsTmp.BOF And orsTmp.EOF) Then
                    tW_StoreNameABB = SP_FEDtChkString(orsTmp, "FTStmNameAbb")
                    tW_StoreNameTax = SP_FEDtChkString(orsTmp, "FTStmNameTax")
                    tW_TaxInvoice = SP_FEDtChkString(orsTmp, "FTStmTaxID")
                    tW_StoreTel = SP_FEDtChkString(orsTmp, "FTStmTel")
                Else
                    tW_StoreNameABB = ""
                    tW_StoreNameTax = ""
                    tW_TaxInvoice = ""
                    tW_StoreTel = ""
                End If
            Else
                tW_StoreNameABB = ""
                tW_StoreNameTax = ""
                tW_TaxInvoice = ""
                tW_StoreTel = ""
            End If
            
            'people
            tW_CustomerID = SP_FEDtChkString(orsHD, "FTCstCode")
            
            tW_CustomerVIPID = SP_FEDtChkString(orsHD, "FTCstVIPCode")
            tW_SalesManID = SP_FEDtChkString(orsHD, "FTSpnCode")
            tW_EmpCdoe = SP_FEDtChkString(orsHD, "FTEmpCode")
            tW_CashierName = W_DATtGetDetail(tW_EmpCdoe, "FTEmpCode", "TCNMEmpMtn", "FTEmpFName")
            If Trim(tW_CashierName) = "" Then
                tW_CashierName = SP_FEDtChkString(orsHD, "FTWhoUpd")
            End If
            'Point
            'Cst 1
            tW_CstCardCode = SP_FEDtChkString(orsHD, "FTCstCardCode")
            nW_LPoint = SP_FEDtChkString(orsHD, "FNCstCardLPoint")
            nW_ToDayPoint = SP_FEDtChkString(orsHD, "FNCstCardPoint")
            tW_NamePoint = SP_FEDtChkString(orsHD, "FTCstPointName")
'            'Cst 2
'            tW_CstCardCode2 = SP_FEDtChkString(orsHD, "FTCstCardCode2")
'            nW_LPoint2 = SP_FEDtChkString(orsHD, "FNCstCardLPoint2")
'            nW_ToDayPoint2 = SP_FEDtChkString(orsHD, "FNCstCardPoint2")
'             tW_NamePoint2 = SP_FEDtChkString(orsHD, "FTCstPointName2")
'            'Cst 3
'            tW_CstCardCode3 = SP_FEDtChkString(orsHD, "FTCstCardCode3")
'            nW_LPoint3 = SP_FEDtChkString(orsHD, "FNCstCardLPoint3")
'            nW_ToDayPoint3 = SP_FEDtChkString(orsHD, "FNCstCardPoint3")
'             tW_NamePoint3 = SP_FEDtChkString(orsHD, "FTCstPointName3")
             tW_PrnPoint = SP_FEDtChkString(orsHD, "FTShdStaPrcCst")

            tW_MemberPriceID = SP_FEDtChkString(orsHD, "FTMemberPriceID")
            
            'OptAlw *KT 51-04-08
            tW_OptAlw = SP_FEDtChkString(orsHD, "FTShdOptAlw")
            tW_DateFmt = W_GETtDateTimeFmt
            
           ' tW_StatusOffline = SP_FEDtChkString(orsHD, "FTShdStaPrcGL") '*KT 51-12-15 Phase 2 ข้อ 31 Communication Sheet PH2-ML-SF-031 Print Messageon Receipt  when OFFLINE.xls ถ้าเป็น 1: Offline นอกนั้นเป็น Online
            tW_StatusOffline = SP_FEDtChkString(orsHD, "FTShdStaType") '*KT 52-06-25 เปลี่ยนจากการขายแบบ Local
           
            'therminal
            If tLastTmn <> tW_TmnNum Then
                    tW_TaxInvoiceABB = W_DATtGetDetail(tW_TmnNum, "FTTmnNum", "TCNMTerminalMtn", "FTTmnRegNum")
                    
                    tSql = "SELECT a.FORjfLogo,a.FTRjfLogo ,a.FTRjfSlpHead,a.FTRjfSlpTrail " & vbCrLf
                    tSql = tSql & "FROM TCNMReceiptJN a,TCNMTerminalMtn b " & vbCrLf
                    tSql = tSql & "WHERE a.FTRjfCode =b.FTRjfCode AND b.FTTmnNum ='" & tW_TmnNum & "'"
                    'Set orsTmp = ocnAll.Execute(tSql)
                    If SP_SQLvExecute(tSql, orsTmp, oDB.ocnOnLine, , , , , True, True) = 0 Then
                        If Not (orsTmp.BOF And orsTmp.EOF) Then
                               'picture
            '                    tW_LogoPath = SP_FEDtChkString(orsTmp, "FTRjfLogo")
                                Set oW_Logo = SP_FEDoChkImage(orsTmp, "FORjfLogo")
            
                                tW_SlipCodeH = SP_FEDtChkString(orsTmp, "FTRjfSlpHead")
                                tW_SlipCodeT = SP_FEDtChkString(orsTmp, "FTRjfSlpTrail")
                                tSql = "SELECT * FROM TCNMSlipMsg "
                                If SP_SQLvExecute(tSql, orsTmp, oDB.ocnOnLine, , , , , True, True) = 0 Then
                                       'slip H
                                    orsTmp.Filter = "FTSlpCode = '" & tW_SlipCodeH & "'"
                                    If Not (orsTmp.BOF And orsTmp.EOF) Then
                                                For nTmp = 1 To 7
                                                        aW_Header(nTmp) = SP_FEDtChkString(orsTmp, "FTSlpDescLine" & nTmp)
                                                Next
                                    End If
                                    orsTmp.Filter = adFilterNone
                                    'slip H
                                    orsTmp.Filter = "FTSlpCode = '" & tW_SlipCodeT & "'"
                                     If Not (orsTmp.BOF And orsTmp.EOF) Then
                                                 For nTmp = 1 To 7
                                                         aW_Trailer(nTmp) = SP_FEDtChkString(orsTmp, "FTSlpDescLine" & nTmp)
                                                 Next
                                     End If
                                     orsTmp.Filter = adFilterNone
                                End If
                         End If
                    End If
            End If
            tLastTmn = tW_TmnNum
            
            'total
            '*KT 51-04-11 Comsheet  ML-SF-208/ML-SF-209 กรณีที่มีส่วนลด Mall Card จะใช้  FCShdTotal เพราะยอดที่แสเงจะมีส่วนลด Mall Card รวมอยู่แต่ใน FCShdTotal หัดส่วนลดทุกชนิดแล้ว
            cW_SubAll = SP_FEDvChkNumeric(orsHD, "FCShdSubReg") + SP_FEDvChkNumeric(orsHD, "FCShdSubPro") 'SP_FEDvChkNumeric(orsHD, "FCShdTotal")
'            cW_SubReg = SP_FEDvChkNumeric(orsHD, "FCShdSubReg")
            cW_SubReg = SP_FEDvChkNumeric(orsHD, "FCShdWpTax")                 '*Ao 2013/03/08  CommSheet TK-ISS3000-026-00 00 04.xls
            cW_SubTotal = SP_FEDvChkNumeric(orsHD, "FCShdTotal")
            cW_Total = SP_FEDvChkNumeric(orsHD, "FCShdGrand")
            cW_Rnd = SP_FEDvChkNumeric(orsHD, "FCShdRnd")
            cW_Chg = SP_FEDvChkNumeric(orsHD, "FCShdChn")
            cW_Vat = SP_FEDvChkNumeric(orsHD, "FCShdVat")
            tW_StaCancel = SP_FEDtChkString(orsHD, "FTShdStaCancel")
            cW_MiscTnd = SP_FEDvChkNumeric(orsHD, "FCShdRndMnyChg") '*Eaw 56-01-04 แก้ไขการปัด Rounding
            tW_RefDepNum = SP_FEDtChkString(orsHD, "FTShdTmnDepositCancel") '*Eaw 56-06-03 CommSheet TK-ISS3000-198-ST 00 00 01.xls
            tW_RefDepNo = SP_FEDtChkString(orsHD, "FTShdTrnDepositCancel")  '*Eaw 56-06-03 CommSheet TK-ISS3000-198-ST 00 00 01.xls
            '*BG 57-01-22 CommSheet TK-ISS3000-446
'            tW_VirTerminal = SP_FEDtChkString(orsHD, "FTShdPosCN")  '*BG 56-09-23 CommSheet TK-ISS3000-356-ST 00 00 01.xls
            Select Case Val(tW_TransType)
                Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax, tEN_TrnTotalReportManualTax, tEN_TrnOverShortReportManualTax, tEN_TrnClosingReportManualTax, tEN_TrnOpenManualTax
                    tChkAdaDB = SP_CHKtAdaConfig  '43,44,45,46,47,49
                    If UCase(tChkAdaDB) = UCase("1") Then ' "1" เครื่อง TK   "0" เครื่อง เอด้า
                        tW_VirTerminal = SP_FEDtChkString(orsHD, "FTTmnNum")
                    Else
                        tW_VirTerminal = SP_FEDtChkString(orsHD, "FTShdPosCN")
                    End If
                Case Else
                    tW_VirTerminal = SP_FEDtChkString(orsHD, "FTShdPosCN")
            End Select
            '****
            If bW_PRMGenEJ Then                 '*Ao 2014/02/03   CommSheet TK-ISS3000-458-FAPL-180-03_Aoo
                '*Bump 58-04-08 CommSheet TK-ISS3000-584.xls
                Select Case Val(tW_TransType)
                    Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax, tEN_TrnTotalReportManualTax, tEN_TrnOverShortReportManualTax, tEN_TrnClosingReportManualTax, tEN_TrnOpenManualTax
                        If SP_FEDtChkString(orsHD, "FTCstCardCode3") = "" Then
                            tVB_EJDate = UT_tStringDate(pdTransDate)
                        Else
                            tVB_EJDate = SP_FEDtChkString(orsHD, "FTCstCardCode3")
                        End If
                Case Else
                    'tVB_EJDate = UT_tStringDate(pdTransDate) '*Tao 59-04-19 ไม่ต้องเปลี่ยนตาม Operation Date (CommSheet TK-ISS3000-610)
                End Select
'                 If SP_FEDtChkString(orsHD, "FTCstCardCode3") = "" Then
'                     tVB_EJDate = UT_tStringDate(pdTransDate)
'                 Else
'                     tVB_EJDate = SP_FEDtChkString(orsHD, "FTCstCardCode3")
'                 End If
            End If
            '*Eaw 56-07-31CommSheet TK-ISS3000-283 TR FAPL-063.xls แยกคิด ManualTax ออกจากค่าอื่น
            Select Case Val(tW_TransTypeSub)
                Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax
                    cW_SubAll = SP_FEDvChkNumeric(orsHD, "FCShdVatable")
            End Select
'            If W_CHKbCheckMemEmp Then
            If W_CHKbCheckMemEmp(pbUseServer) Then '*Eaw 56-07-24 เพิ่ม parameter
                bW_PrintMem = True
            Else
                bW_PrintMem = False
            End If
'            cW_SubReg = W_GETcRegPrice(cW_SubReg, pnTransType, tDate, ptTransTmn, ptTransNo)  '*Eaw 56-02-20  New Req Discountable
            'voucher
            Select Case Val(tW_TransTypeSub)
            Case nEN_TrnVoucher, nEN_TrnCancelVoucher
                    tW_VoicherNo = SP_FEDtChkString(orsHD, "FTShdVoucherNo")
            End Select
            
            'deposit
            Select Case Val(tW_TransTypeSub)
            Case nEN_TrnDeposit, nEN_TrnRedeem, nEN_TrnDepositCancel, nEN_TrnAddDeposit, nEN_TrnAddDepositCancel 'deposit
                eW_Deposit = W_GETeDeposit(pdTransDate, ptTransTmn, ptTransNo)
                bW_DepositMode = True
            Case Else
                bW_DepositMode = False
            End Select
                
            'print
            nW_PrintCnt = SP_FEDvChkNumeric(orsHD, "FNShdDocPrint")
            Call W_SETxReplicated(pnTransType)              '*Ao 2014/05/06
            
            cW_TaxRate = SP_SQLtFunction("", "FCTaxRate", "TCNMTaxMtn", " WHERE FTTaxCode = '" & "01" & "'", oDB.ocnOnLine, True, True) 'read master
                                
' -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -   -  -   -   -   -
           Set orsTmp = Nothing
            '*Amorn 65-08-26 STK-POS ComSheet CR803 - 004_20220824 เคลียค่า cW_RegAmt
'                cW_RegAmt = "0.00"
'                cW_LowBPAmt = "0.00"
'            '*Tao 60-04-10 Low BP
'            If tW_CstCardCode <> "" Then '*Thirapong 64-03-18 performance tuning
                cW_RegAmt = SP_PRCxRegAmount("2", tW_TblDT, tW_TransWhere, pbUseServer) '*Tao 60-04-10 หาจำนวนเงินจากสินค้า Normal BP
                cW_LowBPAmt = SP_PRCxRegAmount("1", tW_TblDT, tW_TransWhere, pbUseServer)  '*Tao 60-04-10 หาจำนวนเงินจากสินค้า Low BP
'            End If
            'return
            OpenTransaction = True
            nW_StateDoc = 1
            
            Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
            nW_StateDoc = 0
             '*KT 53-01-20 ทำการเก็บ Log กรณี PH3-ML-SF-051กรณีที่เกิด OnError
'            Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & "cTransaction" & " Events:OpenTransaction[3]")
'            Call SP_CRTxLogPrg("Last query on " & IIf(pbUseServer, "server", "local") & " : " & tSql)
            Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & "cTransaction" & " Events:OpenTransaction[3]")
            Call SP_CRTxWriteLogTerminated("Last query on " & IIf(pbUseServer, "server", "local") & " : " & tSql)
End Function

Public Function PrintOut() As Boolean

'PrintOut = True
'Exit Function

Dim oFSO As New FileSystemObject
Dim oFile As File
Dim cW As Double
Dim nTnsType  As Long
Dim tSql As String
Dim tTmp As String
Dim tFileEj As String
Dim bFlagEJ As Boolean
Dim bFlagEJWrite As Boolean
Dim tSta As String
Dim bFileOpen As Boolean '*KT 53-11-08 สำหรับ File ที่เกิด Error Runtime 55
Dim bFlagPrn As Boolean
Dim tFolderEJLocal  As String
Dim bCopy As Boolean
Dim tTemp As String
Dim bFlagPrn2 As Boolean
Dim bFlagPrn3 As Boolean
Dim tStaCancel As String  '*Eaw 56-03-04  TK-ISS3000-082
Dim tStaClose As String     '*Eaw 56-03-04  TK-ISS3000-082
Dim nLevel  As Long
Dim nTmp As Integer

On Error GoTo ErrHandler
    oPrn.LineNo = 0
    oPrn.PageNo = 1
    bW_StaPrnCV2 = False
    bW_StaPrnCV3 = False
    bVB_LastLine = True     '*Ao 2013/07/25    ไม่ได้ Login แล้วทำ Layaway  เลยจะทำให้ พิมพ์ Slip ผิด
    bFlagEJ = False
    bFlagPrn = False
    
    bW_StaPrnCVReIssue = False '*Thirapong 2021-03-27
    
       '*KT 53-03-02 PH3-ML-SF-051  No EJ, Dat File 1 kb, No data on Database but it's Sale Transaction
    tSta = ""
    tSta = IIf(nW_StateCon = 0, "0", "1")
    tSta = tSta & IIf(nW_StatePrn = 0, "0", "1")
    tSta = tSta & IIf(nW_StateDoc = 0, "0", "1")
    If (nW_StateCon = 0) Or (nW_StatePrn = 0) Or (nW_StateDoc = 0) Then
         If bW_EnableRePrint = False Then
             ' MsgBox "กรุณาทำการแจ้งเตือนข้อความนี้กับผู้ดูแลระบบ ก่อนทำงานต่อไป" & vbCrLf & "nW_StateCon = " & nW_StateCon & vbCrLf & "nW_StatePrn = " & nW_StatePrn & vbCrLf & "nW_StateDoc = " & nW_StateDoc, vbCritical
              Select Case tSta
                  Case "110"
                      MsgBox "Transaction incomplete, Please Re-Entry Transaction or contact Administrator !", vbCritical 'tms_CN005240
                  Case Else
                      MsgBox "กรุณาทำการแจ้งเตือนข้อความนี้กับผู้ดูแลระบบ ก่อนทำงานต่อไป !" & vbCrLf & "nW_StateCon = " & nW_StateCon & vbCrLf & "nW_StatePrn = " & nW_StatePrn & vbCrLf & "nW_StateDoc = " & nW_StateDoc, vbCritical
              End Select
              Exit Function
        Else
             If nW_StateSuccess = 0 Then
                 ' MsgBox "กรุณาทำการแจ้งเตือนข้อความนี้กับผู้ดูแลระบบ ก่อนทำงานต่อไป" & vbCrLf & "nW_StateCon = " & nW_StateCon & vbCrLf & "nW_StatePrn = " & nW_StatePrn & vbCrLf & "nW_StateDoc = " & nW_StateDoc, vbCritical
                 Select Case tSta
                     Case "110"
                         MsgBox "Transaction incomplete, Please Re-Entry Transaction or contact Administrator !", vbCritical 'tms_CN005240
                     Case Else
                         MsgBox "กรุณาทำการแจ้งเตือนข้อความนี้กับผู้ดูแลระบบ ก่อนทำงานต่อไป !" & vbCrLf & "nW_StateCon = " & nW_StateCon & vbCrLf & "nW_StatePrn = " & nW_StatePrn & vbCrLf & "nW_StateDoc = " & nW_StateDoc, vbCritical
                 End Select
                 Exit Function
             Else
                 bW_EnableEJ = True
             End If
        End If
    End If
       '*Eaw 56-03-04   TK-ISS3000-082
    tStaCancel = SP_FEDtChkString(orsHD, "FTShdStaCancel")
    tStaClose = SP_FEDtChkString(orsHD, "FTCstStaClose")
    If UCase(tStaCancel) = "Y" And tStaClose = "1" Then
        tW_TypePrinter = "1"
    End If
    '-----------------------------------------------------------------------------------------
    tSql = "UPDATE TPSTSalHD SET "                      '*Ao 2015/01/20   CommSheet TK-ISS3000-571
    Select Case tW_TypePrinter
        Case "0"
            tSql = tSql & vbCrLf & "FNLogStaExport = '0'" '*TON 58-10-08 แก้ไขเพิ่ม' ' จาก 0 ให้เป็น '0' ป้องกัน Overflow CommSheet TK-ISS3000-600
        Case "1", "2"
            tSql = tSql & vbCrLf & "FNLogStaExport = '1'" '*TON 58-10-08 แก้ไขเพิ่ม' ' จาก 1 ให้เป็น '1' ป้องกัน Overflow CommSheet TK-ISS3000-600
    End Select
    tSql = tSql & vbCrLf & " WHERE " & tW_TransWhere
    Call SP_SQLvExecute(tSql)
    tSql = "UPDATE " & "TPSHD" & tW_TmnNum & " SET "
    Select Case tW_TypePrinter
        Case "0"
            tSql = tSql & vbCrLf & "FNLogStaExport = '0'" '*TON 58-10-08 แก้ไขเพิ่ม' ' จาก 0 ให้เป็น '0' ป้องกัน Overflow CommSheet TK-ISS3000-600
        Case "1", "2"
            tSql = tSql & vbCrLf & "FNLogStaExport = '1'" '*TON 58-10-08 แก้ไขเพิ่ม' ' จาก 1 ให้เป็น '1' ป้องกัน Overflow CommSheet TK-ISS3000-600
    End Select
    tSql = tSql & vbCrLf & " WHERE " & tW_TransWhere
    Call SP_SQLvExecute(tSql)
    '-----------------------------------------------------------------------------------------
    bW_PrnTraining = False
    bW_PrnVoid = False
    'x pos
    ' - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -
    cW = cW_PageWidth
    uW_X.cEN_Left = 0
    uW_X.cEN_HalfLeft = (cW / 4)
    uW_X.cEN_Center = (cW / 2)
    uW_X.cEN_HalfRight = cW - (cW / 4)
    uW_X.cEN_Right = cW
    uW_X.cEN_HalfColumn = (cW / 8)
    
    cXL = uW_X.cEN_Left
    cXHL = uW_X.cEN_HalfLeft
    cXC = uW_X.cEN_Center
    cXHR = uW_X.cEN_HalfRight
    cXR = uW_X.cEN_Right
    cXHC = uW_X.cEN_HalfColumn
    cXMK = cXR - cW_MrakOffSet
    
    cX5 = cW / 5
    cX4 = 0.5
    cX2 = 0.25
    cX3 = 0.425
    cX15 = 2.275
    cX8 = 1.08 'เป็นค่าสมมุติแทน cX8 สำหรับ OPOS
    cX8x = cX5 + cX2 'สำหรับ Driver
'       'font  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -

    tTmp = "DilleniaUPC"
    
    Set oW_Font(1) = New StdFont
    oW_Font(1).Name = tTmp
    oW_Font(1).bOld = False
    oW_Font(1).Size = 7 + 3.5
    oW_Font(1).Charset = 222 'thai
    
    Set oW_Font(2) = New StdFont
    oW_Font(2).Name = tTmp
    oW_Font(2).bOld = True
    oW_Font(2).Size = 9 + 3.5
    oW_Font(2).Charset = 222 'thai
    
    Set oW_Font(3) = New StdFont  'Barcode
    oW_Font(3).Name = "3 of 9 Barcode"
    oW_Font(3).bOld = False
    oW_Font(3).Size = 17

    oPrn.bPrnByOPOS = bC_PrnByOPOS
    oPrn.tTmnNum = tW_TmnNum
    oPrn.LineOffset = -0.12
    
    oPrn.tTransTypeSub = tW_TransTypeSub                '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
    '*KT 51-12-16 Flag ที่บอกว่ามีการเรียกใช้งานจาก Tools Generate EJ
    oPrn.bPRMGenEJ = bW_PRMGenEJ
    oPrn.PrinterType = tW_TypePrinter
        
        
        '*Amorn 65-08-09 PRJ10001TK-SE-External Specification CR Increase GST Rate from current 7% to 8%-00.01.00.docx
        Select Case Val(tW_TransTypeSub)
        Case nEN_TrnRedeem, nEN_TrnDepositCancel, nEN_TrnAddDeposit, nEN_TrnAddDepositCancel
                tVB_DepositNoGST7 = eW_Deposit.tEN_DepositNo
        Case Else
                tVB_DepositNoGST7 = ""
        End Select
         
        
    If Not bW_PRMGenEJ Then                     '*Ao 2013/08/03  Check print EJ
        If bC_PrnByOPOS Then
            If tW_TypePrinter = "0" Then
                nLevel = oOposSlip.nCheckHealth(OPOS_CH_INTERNAL)       '*Ao 2013/07/26 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
                If nLevel = OPOS_E_ILLEGAL Then
                '*Bump 57-09-30 10.15am. CommSheet TK-ISS3000-555 Add log เพื่อ Monitor
                Call SP_CRTxLogPrg("ErrorNumber:" & OPOS_E_ILLEGAL & " ErrorDesc: Slip Printer is ILLEGAL (Program retry open OPOS Slip Printer" & "cTransection" & " Function:PrintOut: " & ",TerminalNum:" & tVB_CNTerminalNum & " TransNo.:" & tVB_Running & " TransType:" & tVB_Transaction)
                 wCNMsgSlip.Show vbModal
                End If

                If (bW_EnableRePrintNormal And bW_EnableRePrint) Or bW_EnableRePrint = False Then          '*Ao 2013/10/01 CommSheet TK-ISS3000-367   V 4.0.73
                    wCNMsgErrSlip.tW_Msg = ""                                                                   'Please insert slip form
                    wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
                    wCNMsgErrSlip.bW_Printing = False
                    wCNMsgErrSlip.bW_Wait = False
                    wCNMsgErrSlip.nW_Line = 0
                    wCNMsgErrSlip.Show vbModal
                    
                    Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ",Please insert slip form")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                    
                End If
            Else
                If oOpos.oPrinterState = 4 Then
                    wCNMsgErr.Show vbModal
                Else
                    If oOpos.CoverOpen Then
                        wCNMsgErr.Show vbModal
                        End If
                End If
            End If
            If (bW_EnableRePrintNormal And bW_EnableRePrint) Or bW_EnableRePrint = False Then
                If tW_TypePrinter <> "0" Then
                    oOpos.TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'Initial start  print by OPOS
                    oOpos.RecLineSpacing = 10
                Else
                    oOposSlip.TransactionPrint PTR_S_SLIP, PTR_TP_TRANSACTION 'Initial start  print by OPOS
                    Call SP_CRTxLogPrg("SLIP PRINTER-TransType:" & tW_TransTypeSub & ",Initial start OPOS")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                    'oOposSlip.RecLineSpacing = 10
                    Select Case Val(tW_TransTypeSub)
'                        Case 3 'Sale
                        Case 3, 38, 43 '3 = Sale , 38 = Tax Include Sale ,43 = Tax Manual Sale '*Eaw 56-06-27
                            oPrn.LineMax = nCS_MaxLineSlipSale  '*Bump 56-05-03 เพิ่ม Case 3 'Sale
'                        Case 4 'Return
                        Case 4, 39, 44 '4 = Return , 39 = Tax Include Return ,44 = Tax Manual Return '*Eaw 56-06-27
                            'oPrn.LineMax = 22 'TK-ISS3000-078 Counting number of lines to be printed at Slip printer is incorrected
                            'oPrn.LineMax = 21 '*Bump 56-05-03 Comment บรรทัดนี้
                            oPrn.LineMax = nCS_MaxLineSlipReturn '*Bump 56-05-03
                        Case 6, 7, 16, 17, 19, 20 'Layaway,Layaway Pickup.Layaway Cancel ,IPV Layaway Pickup,Add Layaway,Cancel Add Layaway
                           'oPrn.LineMax = 34 '*Bump 56-05-03 Comment บรรทัดนี้
                            oPrn.LineMax = nCS_MaxLineSlipLayaway '*Bump 56-05-03
                        Case Else
                            'oPrn.LineMax = 18 'TK-ComSheet-003   SALE , RETURN Transaction ที่พิมพ์ออก SLIP Printer  ขอให้พิมพ์ประมาณ 22 บรรทัดครับ
                            'oPrn.LineMax = 21 '*Bump 56-05-03 Comment บรรทัดนี้
                            oPrn.LineMax = nCS_MaxLineSlipSale '*Bump 56-05-03
                    End Select
                End If
            End If
        End If
    End If
        'EJ - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -
                'new ej
    If bW_EnableEJ Then
        oPrn.PrintEJ = True
        Call oPrn.EJNewCapture(cW_PageWidth, 64)
    End If
        'other - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -
    tW_TrMk = IIf(bW_TrainMode, tTkMdTrainingMrak, vbNullString)
    nW_Items = 0
        ' - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -
        
    nTnsType = Val(tW_TransTypeSub)
    nW_TokenTransType = nTnsType '*Bmp 61-04-05 ใช้ detoken
    Call W_GETxVocabularyTicketToArray
                    
    If Not bW_PRMGenEJ Then                     '*Ao 2013/08/03  Check print EJ
        If Not bC_PrnByOPOS Then
            If (bW_EnableRePrintNormal And bW_EnableRePrint) Or bW_EnableRePrint = False Then
                oPrn.PrintPicture oW_Logo, 2.5, 0, vbCenter
            End If
        End If
    End If
    
    If (bW_EnableRePrintNormal And bW_EnableRePrint) Or bW_EnableRePrint = False Then
        
        If Not bW_PRMGenEJ Then                     '*Ao 2013/08/03  Check print EJ
            If tW_TypePrinter <> "0" Then       '*Ao 2013/01/24   Print FRONTTECH
                If Not W_PRNbTicketLOGO Then GoTo ErrHandler
            End If
        End If
        
        If tW_TypePrinter <> "0" Then
            If Not W_PRNbTicketHeaderMsg Then GoTo ErrHandler
        End If
        
        If tW_TypePrinter <> "0" Then
            If Not W_PRNbTicketHeader Then GoTo ErrHandler
        Else
            If Not W_PRNbTicketHeaderSlp Then GoTo ErrHandler
        End If
    
        If Not W_PRNbTicketModeStart Then GoTo ErrHandler
                    
        If Not W_PRNbTicketForeignCurrency Then GoTo ErrHandler              '*Ao 2012/07/25  PRJ10001TK-External Specification Open Trx-00.01.00.doc
                    
        Select Case nTnsType
            Case 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49 '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls เพิ่ม 48,49
                If Not W_PRNbTicketTransReEntry Then GoTo ErrHandler
        End Select
                    
        If Not W_PRNbTicketDetail Then GoTo ErrHandler
        
        Select Case nTnsType
            Case 3, 4, 6, 7, 10, 11, 13, 16, 19, 20, 22, 23, 30, 38, 39, 43, 44
                If Not W_PRNbTicketTransDisc Then GoTo ErrHandler
        End Select
        
        Select Case nTnsType
            Case 3, 4, 6, 7, 10, 11, 13, 15, 16, 19, 20, 21, 22, 23, 29, 30, 38, 39, 43, 44
                    If Not W_PRNbTicketTotal Then GoTo ErrHandler
        End Select
         
         Select Case nTnsType
            Case 3, 4, 10, 13, 11, 22, 23, 26, 30, 38, 39, 43, 44
                    If Not W_PRNbTicketFreeGift Then GoTo ErrHandler
        End Select
         
        tW_SlipFooterStr = "" '*Eaw 56-10-08
        Select Case nTnsType
            Case 3, 4, 6, 7, 10, 13, 11, 15, 16, 19, 20, 21, 22, 23, 30, 38, 39, 43, 44
                If Not W_PRNbTicketTender Then
                    If Not bW_PRMGenEJ Then GoTo ErrHandler
                End If
        End Select
        
        If Not W_PRNbTicketModeEnd Then GoTo ErrHandler
        
        '*Eaw 56-10-08
        nW_LineNo = oPrn.LineNo
        nW_LineMax = oPrn.LineMax

        
        If Not W_PRNbTicketPeople Then
            If Not bW_PRMGenEJ Then GoTo ErrHandler 'GoTo ErrHandler
        End If
'        MsgBox "bW_ChkRep = " & bW_ChkRep
        If bW_ChkRep Then
            Call W_PRNxTicketReplicated         '*Ao 2014/05/06
        End If
        
        Call W_PRNbTicketShowOffline '*KT 51-12-15 Phase 2  ข้อ 31 Communication Sheet PH2-ML-SF-031 Print Messageon Receipt  when OFFLINE.xls
        
        Select Case nTnsType
            Case 7, 16, 19, 20 '*Eaw 56-08-03 '*Eaw 56-08-03 CommSheet TK-ISS3000-304-UAT_Req.xls เพิ่ม Type 7,16,20
                If Not W_PRNbTicketDepInfo Then GoTo ErrHandler
        End Select
        
        nW_TokenTransType = nTnsType '*Bmp 61-04-05
        
        If nW_PrintCnt > 0 Or bW_PRMGenEJ Then              'Reprint OR ManageEJ
            Call W_PRNbTicketPointRe                '*Ao 2013/11/30  CommSheet TK-ISS3000-FAPL-160
        Else
            Call W_PRNbTicketPointEx '*KT 52-02-23 member point
        End If
'        Call W_PRNbTicketPointEx '*KT 52-02-23 member point
        
'        Call W_PRNxPrintSlip(tW_SlipFooterStr, tW_TypePrinter, nW_LineNo, nW_LineMax) '*Eaw 56-10-08

        If tW_TypePrinter <> "0" Then               '*Ao 2014/06/20
            nW_LineMax = 500
            Call W_PRNxPrintSlip(tW_SlipFooterStr, tW_TypePrinter, nW_LineNo, nW_LineMax)
            If Not bW_PRMGenEJ Then                     '*Ao 2014/04/21  Check print EJ
'                If tW_TypePrinter <> "0" Then
'                    If Not W_PRNbTicketBarcode Then GoTo ErrHandler         '*Ao 2013/01/24
'                End If
                If Not W_PRNbTicketBarcode Then GoTo ErrHandler         '*Ao 2013/01/24
                If Not W_PRNbTicketFooterMsg Then GoTo ErrHandler
                If Not W_PRNbTicketFooterBlank Then GoTo ErrHandler
            End If
'            Call oPrn.PrintText("", cXC, , vbCenter, , nEN_PrinterOnly)                 '*Ao 2014/07/19  TK-ISS3000PH2-525   V 4.2.138
'            If Not W_PRNbTicketFooterMsg Then GoTo ErrHandler
'            If Not W_PRNbTicketFooterBlank Then GoTo ErrHandler
        Else        '0 = Slip Printer
            If nTnsType <> 12 Then
                tTmp = " "
                tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"                  '*Ao 2014/07/19  TK-ISS3000PH2-525   V 4.2.138
                Set oPrn.Font = oW_Font(1)
                For nTmp = 1 To 7
                    If aW_Trailer(nTmp) <> vbNullString Then
                        tTmp = aW_Trailer(nTmp)
                        tW_SlipFooterStr = tW_SlipFooterStr & "*" & tTmp & ";"
                    End If
                Next
            End If
            Call W_PRNxPrintSlip(tW_SlipFooterStr, tW_TypePrinter, nW_LineNo, nW_LineMax)
        End If
        
'        If Not bW_PRMGenEJ Then                     '*Ao 2014/04/21  Check print EJ
'            If tW_TypePrinter <> "0" Then
'                If Not W_PRNbTicketBarcode Then GoTo ErrHandler         '*Ao 2013/01/24
'            End If
'        End If
        'Bump Note  57-04-22
        'PRJ10001TK-PH2-External Spec-Slip Printing 00.02.04 FSBT.docx
        'การแก้ไข ไม่ต้องพิมพ์ข้อความท้ายบิล ของ Cash transfer, Credit vocuher
        '-------------------------
'        If tW_TypePrinter <> "0" Then  '*BG 57-04-28 PRJ10001TK-PH2-External Spec-Slip Printing 00.02.04 FSBT.docx
'         If Not W_PRNbTicketFooterMsg Then GoTo ErrHandler   '*Ao 2013/01/24
'        End If
        
'        If Not W_PRNbTicketFooterMsg Then GoTo ErrHandler
'        If tW_TypePrinter <> "0" Then       '*Ao 2013/01/24   Print FRONTTECH
'            If Not W_PRNbTicketFooterBlank Then GoTo ErrHandler
'        End If
        
    End If
    
    bFlagEJ = True
    
    If bC_PrnByOPOS Then                        '*Ao 2014/06/21
        If tW_TypePrinter = "0" Then
            If oOposSlip.SlpEmpty = True Then
                wCNMsgErrSlip.tW_Msg = ""                                                                   'Press Enter to print
                wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub           '*Ao 2014/04/23  PRJ10001TK-PH2-External Spec-Auto Release Paper at Slip Printer-00.02.03
                wCNMsgErrSlip.bW_Printing = False
                wCNMsgErrSlip.bW_Wait = False
                wCNMsgErrSlip.nW_Line = 0
                wCNMsgErrSlip.Show vbModal
                Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ",Press Enter to print")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
            End If
        End If
    End If
    
    If Not bW_PRMGenEJ Then '*KT 51-12-16 Flag นี้แสดงว่ามีการเรียกใช้งานจาก Tools Generate EJ ซึ่งจะไม่มีการพิมพ์ Slip
        If (bW_EnableRePrintNormal And bW_EnableRePrint) Or bW_EnableRePrint = False Then   '*Ao 2013/10/02  CommSheet TK-ISS3000-367
            Select Case nTnsType
                Case 1, 5, 8, 9, 14, 17, 25, 26, 27, 28, 29, 30, 31, 32, 34, 35, 36, 37, 40, 42, 45, 47, 48, 49 '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls เพิ่ม 48,49
                    bFlagPrn = oPrn.EndDoc(oW_Logo, False)
                Case Else
                         '*KT 54-07-25 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon
                         '*KT 54-11-14 V.4.0.4 DCR_5411001  ต้องการให้ Slip กับ Promotion Coupon เป็นแผ่นเดียว เนื่องจากกลัวว่าลูกค้าจะทำ Promotion Coupon หาย
                     Select Case nTnsType
                        Case 3, 6, 10, 11, 13, 38, 43
                            If Not orsPmtCpn Is Nothing Then
                                If orsPmtCpn.State = adStateOpen Then
                                    If Not (orsPmtCpn.BOF And orsPmtCpn.EOF) Then
                                        Call W_PRNbTicketPmtCpn
                                    End If
                                End If
                            End If
                    End Select

                    If bVB_OPOSCashDrawerEnable Then
                        bFlagPrn = oPrn.EndDoc(oW_Logo, oOpos.DrawerOpen, True)
                    Else
                        bFlagPrn = oPrn.EndDoc(oW_Logo, False, True)
                    End If
            End Select

            If bC_PrnByOPOS Then
                If tW_TypePrinter = "0" Then
                    wCNMsgErrSlip.tW_Msg = ""
                    wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
                    wCNMsgErrSlip.bW_Printing = True
                    wCNMsgErrSlip.bW_Wait = True
                    wCNMsgErrSlip.nW_Line = oPrn.LineNoEx
                    wCNMsgErrSlip.Show vbModal
                    Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ",Slip printing in progress(1)")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                End If
            End If
        Else
            bFlagPrn = True             '*Ao 2013/10/02  CommSheet TK-ISS3000-367
        End If
    End If
    
        'update
        '*KT 51-12-25 ตรวจสอบเงื่อนไขว่าใช้ Tools Generate EJ หรือไม่
        If Not bW_PRMGenEJ Then
            nW_PrintCnt = nW_PrintCnt + 1
            'tSql = "UPDATE " & tW_TblHD & " SET FNShdDocPrint = " & nW_PrintCnt & " WHERE " & tW_TransWhere
            '*KT 52-04-09 Update จำนวนการพิมพ์ Receipt ทั้ง ตาราง TPSTSALHD และ TPSHDxxxxx
                '*Eaw 56-10-22 TK-CommSheet 382
    '            tSql = "UPDATE TPSTSALHD SET FNShdDocPrint = " & nW_PrintCnt & " WHERE " & tW_TransWhere
                tSql = "UPDATE TPSTSALHD SET FNShdDocPrint = " & nW_PrintCnt
                Select Case tW_TypePrinter
                    Case "0"
                        tSql = tSql & vbCrLf & " , FNLogStaExport = 0"
                    Case "1", "2"
                        tSql = tSql & vbCrLf & " , FNLogStaExport = 1"
                End Select
                tSql = tSql & vbCrLf & " WHERE " & tW_TransWhere
                '----------------------------------------
            Call SP_SQLvExecute(tSql)  'Call ocnAll.Execute(tSql)
            If SP_TBLbChkExist("TPSHD" & tW_TmnNum, ocnAll) Then
                '*Eaw 56-10-22 TK-CommSheet 382
    '            tSql = "UPDATE " & "TPSHD" & tW_TmnNum & " SET FNShdDocPrint = " & nW_PrintCnt & " WHERE " & tW_TransWhere
                tSql = "UPDATE " & "TPSHD" & tW_TmnNum & " SET FNShdDocPrint = " & nW_PrintCnt
                Select Case tW_TypePrinter
                    Case "0"
                        tSql = tSql & vbCrLf & " , FNLogStaExport = 0"
                    Case "1", "2"
                        tSql = tSql & vbCrLf & " , FNLogStaExport = 1"
                End Select
                tSql = tSql & vbCrLf & " WHERE " & tW_TransWhere
                '----------------------------------------
                Call SP_SQLvExecute(tSql)  'Call ocnAll.Execute(tSql)
            End If
        End If
        
'        Call SP_CRTxLogFile("Begin:" & Format(tW_DepID, "00") & Format(tW_TmnNum, "00000") & Format(tW_TransNo, "00000") & tW_StoreID)  '*Ao 2013/12/16   เก็บ LOG สร้าง EJ
        'EJ - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -
        If bVB_ReEntryMode Then                '*Ao 2013/10/01       ReEntry  ให้เก็บ EJ ตาม Operation Date
            If tVB_EJDate = "" Then
                tVB_EJDate = UT_tStringDate(SP_GETdSaleDate)
            End If
            dW_TransDate = Format(CDate(tVB_EJDate), "dd/MM/yyyy")
'            Call SP_CRTxLogFile("Ej ReEntry date:" & Format(CDate(tVB_EJDate), "dd/MM/yyyy"))           '*Ao 2013/12/16   เก็บ LOG สร้าง EJ
        End If
        If bW_PRMGenEJ Then                 '*Ao 2014/02/03   CommSheet TK-ISS3000-458-FAPL-180-03_Aoo
            If tVB_EJDate = "" Then
'                dW_TransDate = dW_TransDate
            Else
                dW_TransDate = Format(CDate(tVB_EJDate), "dd/MM/yyyy")
            End If
        End If
        If bW_EnableEJ Then
    'PH 2.0 ML-SF-019 EJ missing  บางครั้ง EJ หาย โดยที่ไม่มีทั้งบน Server และที่ เครื่อง POS (Mall5 3/06/09 POS no. 4410 /5196  Mall2 2/06/09 POS 1107/30551,POS 3305/37061)
            bFlagEJWrite = False
            'create  folder by date
            'สร้างข้อมูลลง EJ ที่ Local ก่อน
            tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\" 'หา Path ของการสร้าง ej ที่ Loacal (folder = "POSEJ")
            If Not SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then  'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                tTmp = App.Path & "\POSEJ\" 'หา Path ของการสร้าง ej ที่ Loacal ใหม่
                Call SP_FLDbManFolder(tTmp, nEN_Fld3Create) 'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
            End If
            tTmp = tTmp & Format(dW_TransDate, "yyMMdd") & "\" 'หา Path ของการสร้าง ej ที่ Loacal (folder = "POSEJ") ที่เป็น folder ตามวันที่
            tFolderEJLocal = tTmp
            Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)  'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
            'get and write ej
            tFileEj = tTmp & Format(tW_DepID, "00") & Format(tW_TmnNum, "00000") & Format(tW_TransNo, "00000") & tW_StoreID & ".EJ" 'หา Path ของ file ej
            If SP_FLEbManFile(tFileEj, nEN_Fle1Exist) Then 'ตรวจสอบว่า file ej มีหรือไม่
                Call SP_FLEbManFile(tFileEj, nEN_Fld2Delete) 'พบ file ej ให้ทำการลบทิ้ง
            End If
            tW_GETFileEJName = tFileEj
            tTmp = oPrn.EJText
            Open tFileEj For Output As #1
            bFileOpen = True
            Print #1, tTmp
            Close #1
            bFileOpen = False
            '*Bmp 59-08-05
            If SP_FLEbManFile(tFolderEJLocal & ".EJ", nEN_Fle2Delete) Then
                '*เจอปัญหา ไฟล์ EJ เป็น ".EJ" แทนที่จะเป็น ชื่อเต็ม เช่น "D1000650031101.EJ"
                'ทำให้ Merge EJ ไม่ได้
                'Write Log ไฟล์ที่เป็น ".EJ"
                Call SP_CRTxLogFile("TransNo.:" & Format(tW_TransNo, "00000") & " Trans.Type:" & tW_TransTypeSub & ", .EJ is problem")
            End If
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            If SP_FLEbManFile(tFileEj, nEN_Fle1Exist) Then
                Call SP_CRTxLogFile("Success create:" & Format(tW_DepID, "00") & Format(tW_TmnNum, "00000") & Format(tW_TransNo, "00000") & tW_StoreID & ".EJ")        '*Ao 2013/12/16   เก็บ LOG สร้าง EJ
            Else
                Call SP_CRTxLogFile("Can not create:" & Format(tW_DepID, "00") & Format(tW_TmnNum, "00000") & Format(tW_TransNo, "00000") & tW_StoreID & ".EJ")        '*Ao 2013/12/16   เก็บ LOG สร้าง EJ
            End If

            tTmp = tW_EJFolder & Format(dW_TransDate, "yyMMdd") & "\" 'หา Path ของการสร้าง ej ที่ server (folder = "POSEJ") ที่เป็น folder ตามวันที่
                    '*KT 54-09-20 V.4.0.1 ป้องกันกรณีที่ Path EJ Local กับ Server ที่เดียวกัน
            If UCase(tFolderEJLocal) <> UCase(tTmp) Then
                    'C:\Program Files\ISS3000\POSEJ  กับ
                    '\\Ada54a\ISS3000\POSEJ
                    If UCase("\\" & cUT.UT_SYStGetComName) = UCase(Mid(tTmp, 1, 2 + Len(cUT.UT_SYStGetComName))) Then
                        tTemp = Replace(UCase(tTmp), UCase("\\" & cUT.UT_SYStGetComName), "")
                        If UCase(Right(tFolderEJLocal, Len(tTemp))) = UCase(Right(tTmp, Len(tTemp))) Then
                            'bCopy = False
                            'Call SP_CRTxLogFile("UCase(Right(tFolderEJLocal, Len(tTemp))) = UCase(Right(tTmp, Len(tTemp)))")
                            'Call SP_CRTxLogFile("tFolderEJLocal:" & tFolderEJLocal & " tTmp:" & tTmp)
                            '**********************************************************************************************************************
                            '*Bump 57-07-16 เพิ่มการเช็ค ป้องกับ Gen EJ แล้ว ไม่ Copy EJ เนื่องจากกำหนด Path POSEJ อยู่ในเครื่องเดียวกัน แต่คนละที่
                            '* เช่น Comname  = Cluster ,Path EJ ปกติ               C:\Program Files (x86)\ISS3000\POSEJ
                            '*  กำหนด Path EJ ไว้ที่     E:\POSEJ    shared folder เป็น \\Cluster\POSEJ
                            If Not bW_PRMGenEJ Then
                                '* การทำงานจาก POSFront.exe ให้เช็คแบบเดิม
                                bCopy = False
                                Call SP_CRTxLogFile("UCase(Right(tFolderEJLocal, Len(tTemp))) = UCase(Right(tTmp, Len(tTemp)))")
                                Call SP_CRTxLogFile("tFolderEJLocal:" & tFolderEJLocal & " tTmp:" & tTmp)
                            Else
                                '* การทำงานจาก ManageEJ.exe  ให้เช็คก่อนว่า EJ ที่ Folder ปลายทางมีหรือไม่ ถ้าไม่มี ให้ bCopy = True เริ่มแก้ไขใน ManageEJ V.4.0.16 16/07/2014
                                '* เนื่องจากปัญหา พบใน ManageEJ.exe
                                Set oFile = oFSO.GetFile(tFileEj)
                                If SP_FLEbManFile(tTmp & "\" & oFile.Name, nEN_Fle1Exist) Then
                                    bCopy = False
                                Else
                                    bCopy = True
                                End If
                            End If
                            '**********************************************************************************************************************
                        Else
                            bCopy = True
                        End If
                    Else
                        bCopy = True
                    End If
            Else
                bCopy = False
                Call SP_CRTxLogFile("tFolderEJLocal:" & tFolderEJLocal & " tTmp:" & tTmp)
            End If
                    '*KT 54-09-20 V.4.0.1 ป้องกันกรณีที่ Path EJ Local กับ Server ที่เดียวกัน
            If bCopy Then
                If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then                        'เช็ค ON LINE หรือไม่        '*Ao 2011/03/19   CommSheet ML-PRV-046- POS Front  Terminate  Reboot and Not responding.xls
                    If SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then 'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                            Set oFile = oFSO.GetFile(tFileEj)
                            If oFSO.GetParentFolderName(oFile.Path) <> tTmp Then  'ตรวจสอบว่า Path Local กับ Server เป็น Path เดียวกันหรือไม่ กรณีที่เป็น path เดียวกันจะไม่ทำการลบ file
                                Call SP_FLEbManFile(tTmp & oFile.Name, nEN_Fld2Delete) 'พบ file ej ให้ทำการลบทิ้ง
                            End If
                            If SP_FLEbManFile(oFile.Path, nEN_Fle1Exist) Then 'ตรวจสอบว่า file ej ที่ local มีหรือไม่
                                oFile.Attributes = Normal 'set file ไม่ให้ read only
                                oFSO.CopyFile oFile.Path, tTmp 'copy ไปที่ server
                            End If
                            If SP_FLEbManFile(tTmp & "\" & oFile.Name, nEN_Fle1Exist) Then                          '*Ao 2014/03/04 CommSheet TK-ISS3000-468
                                Call SP_CRTxLogFile("Send completed:" & oFile.Name)
                            Else
                                Call SP_CRTxLogFile("Send not completed:" & oFile.Name)
                            End If
                            'PH 2.0 ML-SF-018 Backup EJ File ไว้ที่ ....\\POSEJ\Backup\YYMMDD
                            tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\Backup\" 'หา Path ของการสร้าง ej ที่ local (folder = "POSEJ") ที่เป็น folder Backup
                            If Not SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then 'ไม่เจอ folder backup  'ตรวจสอบว่า folder Backup มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                               tTmp = App.Path & "\POSEJ\Backup\" 'หา Path ของการสร้าง ej ที่ Loacal Backup ใหม่
                               Call SP_FLDbManFolder(tTmp, nEN_Fld3Create) 'สร้าง folder backup 'ตรวจสอบว่า folder backup มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                            End If
                            tTmp = tTmp & Format(dW_TransDate, "yyMMdd") & "\"  'หา Path ของการสร้าง ej ที่ Loacal Backup(folder = "POSEJ") ที่เป็น folder ตามวันที่
                            Call SP_FLDbManFolder(tTmp, nEN_Fld3Create) 'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                            If SP_FLEbManFile(oFile.Path, nEN_Fle1Exist) Then 'เจอ file ที่ Temp 'ตรวจสอบว่า file ej ใน local มีหรือไม่
                                oFile.Attributes = Normal 'set file ไม่ให้ read only
                                Call SP_FLEbManFile(tTmp & oFile.Name, nEN_Fld2Delete) 'ลบ file ที่เหมือนกันที่ Backup 'ตรวจสอบว่า file ej ใน local Backup มีหรือไม่ ถ้ามีให้ลบทิ้ง
                                oFSO.CopyFile oFile.Path, tTmp 'copy จาก Temp ไป backup
                            End If
                            'ลบ File EJ ที่เป็น Temp ทิ้ง
                            tW_GETFileEJName = tTmp & oFile.Name
                            If SP_FLEbManFile(tTmp & oFile.Name, nEN_Fle1Exist) Then 'เจอ file ที่ Backup แล้ว
                                oFile.Attributes = Normal 'set file ไม่ให้ read only
                                oFSO.DeleteFile oFile.Path 'ลบข้อมูลที่ Temp ของ Local
                            End If
                    Else 'ส่ง File EJ ไปที่ Server ไม่สำเร็จ
                            'PH 2.0 ML-SF-018 Backup EJ File ไว้ที่ ....\\POSEJ\Backup\YYMMDD
                            Set oFile = oFSO.GetFile(tFileEj)
                            
                            Call SP_CRTxLogFile("Send not completed  " & oFile.Name)            '*Ao 2014/03/04 CommSheet TK-ISS3000-468
                            Call SP_CRTxLogFile("Send not completed because can not create " & tTmp)            '*Ao 2014/03/04 CommSheet TK-ISS3000-468
                            
                            tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\Backup\"
                            If Not SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then
                               tTmp = App.Path & "\POSEJ\Backup\"
                               Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)
                            End If
                            tTmp = tTmp & Format(dW_TransDate, "yyMMdd") & "\"
                            Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)
                            oFile.Attributes = Normal
                            Call SP_FLEbManFile(tTmp & oFile.Name, nEN_Fld2Delete)
                            oFSO.CopyFile oFile.Path, tTmp
                    End If
                    If bVB_ReEntryMode Then                '*Ao 2013/10/01       ReEntry  ให้เก็บ EJ ตาม Operation Date
                        Call SP_MOVxEJToOnlineRE(Format(dW_TransDate, "yyMMdd"))
                    End If
                Else
                    'OFF LINE ไม่ต้องย้าย EJ ไป server
                    'PH 2.0 ML-SF-018 Backup EJ File ไว้ที่ ....\\POSEJ\Backup\YYMMDD
                    Set oFile = oFSO.GetFile(tFileEj)
                    tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\Backup\"
                    If Not SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then
                       tTmp = App.Path & "\POSEJ\Backup\"
                       Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)
                    End If
                    tTmp = tTmp & Format(dW_TransDate, "yyMMdd") & "\"
                    Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)
                    oFile.Attributes = Normal
                    Call SP_FLEbManFile(tTmp & oFile.Name, nEN_Fld2Delete)
                    oFSO.CopyFile oFile.Path, tTmp
                    Call SP_CRTxLogFile("Send not Completed:OFF LINE:" & oFile.Name)            '*Ao 2014/03/04 CommSheet TK-ISS3000-468
                End If
            Else
                Set oFile = oFSO.GetFile(tFileEj)
                tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\Backup\"
                If Not SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then
                   tTmp = App.Path & "\POSEJ\Backup\"
                   Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)
                End If
                tTmp = tTmp & Format(dW_TransDate, "yyMMdd") & "\"
                Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)
                oFile.Attributes = Normal
                Call SP_FLEbManFile(tTmp & oFile.Name, nEN_Fld2Delete)
                oFSO.CopyFile oFile.Path, tTmp
            End If
            Set oFSO = Nothing
            bFlagEJWrite = True
        End If
        If bVB_ReEntryMode Then                '*Ao 2013/10/01       ReEntry  ให้เก็บ EJ ตาม Operation Date
            dW_TransDate = dW_EJDate
        End If
        If bW_PRMGenEJ Then                         '*Ao 2014/02/03   CommSheet TK-ISS3000-458-FAPL-180-03_Aoo
            dW_TransDate = dW_EJDate
        End If
        ' - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -
        bFlagEJ = False
        


'return
        PrintOut = True
         '*KT 54-04-28 ป้องกันการพิมพ์ Slip ซ้ำต่อเนื่องกัน ถ้าคำสั่ง EndDoc ไม่สำเร็จให้ทำการ ClearInput
         '*KT 54-05-30 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย ***แสดงข้อความเฉพาะกรณีที่เป็น POSFront.exe เท่านั้น
        If Not bFlagPrn And Not bW_PRMGenEJ Then
            MsgBox "Print not success.Goto function Reprint Receipt"
            oPrn.ClearBufferOPOSPrn
        End If
        oPrn.KillDoc
        
        If bW_EnableEJ And Not bFlagEJWrite And Not bW_PRMGenEJ Then
           Call SP_CRTxLogFile("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Events:PrintOut" & " Type:Write EJ Ticket : Terminal No. " & tW_TmnNum & " Trans No. " & tW_TransNo) '*KT 52-12-21 ทำการเก้บ Log การเกิด Error ที่เกี่ยวข้องการการทำ Dat File และ EJ
        End If
        
         'TK-ISS3000-066
        '*KT 55-12-11 Print Credit Voucher with Refunded amount, print to Slip printer.
        'Check Slip
        'SALS ENDORSEMENT = 2
        'If Not bW_PRMGenEJ Then      '*Ao 2013/04/30 CommSheet TK-ISS3000-163-IT2 00.00.01.xls
        If Not bW_PRMGenEJ And Not bW_NoPrintCreditVoucher Then '*Tao 62-09-12 TTK-ISS3000-747 พิมพ์ซ้ำคูปอง ไม่ต้องพิมพ์ Credit Voucher
            If (bW_EnableRePrintCV2 And bW_EnableRePrint) Or bW_EnableRePrint = False Then
                If bW_EnableRePrintCV2 Then Call W_GETxDataToReprint
                If bW_StaPrnCV3 Then
                    If bC_PrnByOPOS Then
                        nLevel = oOposSlip.nCheckHealth(OPOS_CH_INTERNAL)       '*Ao 2013/07/26 CommSheet TK-ISS3000-347-UAT        V 4.0.62
                        If nLevel = OPOS_E_ILLEGAL Then
                            wCNMsgSlip.Show vbModal
                        End If
                    End If
                    'If tW_TypePrinter = 0 Then
                        'If bVB_CNTestSlipPrn Then
                            'แก้ไข Comsheet  TK-ISS3000-066
                    orsRC.Filter = adFilterNone
                    orsRC.Filter = "FTTdmCode='T018'"
                    If Not (orsRC.BOF And orsRC.EOF) Then
                        Do While Not orsRC.EOF
'                                    If tW_TypePrinter = "0" Then
'                                        wCNMsgErrSlip.bW_Wait = True
'                                    Else
'                                        wCNMsgErrSlip.bW_Wait = False
'                                    End If
'                                    wCNMsgErrSlip.bW_Printing = False
'                                    wCNMsgErrSlip.nW_Line = 0
'                                    cVB_CVAmtEx = SP_FEDvChkNumeric(orsRC, "FCSrcAmt") '*Eaw 56-07-04 CommSheet TK-ISS3000-257.xls
'                                    wCNMsgErrSlip.tW_Msg = "PRINT SALES ENDORSEMENT"
'                                    wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
'                                    wCNMsgErrSlip.Show vbModal
'                                    If (Not bVB_CNCheckCV) And (tW_TypePrinter = "0") Then
'                                        wCNMsgErrSlip.bW_Wait = False
'                                        wCNMsgErrSlip.tW_Msg = "PRINT SALES ENDORSEMENT"
'                                        wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
'                                        wCNMsgErrSlip.Show vbModal
'                                    End If
                            cVB_CVAmtEx = SP_FEDvChkNumeric(orsRC, "FCSrcAmt") '*Eaw 56-07-04 CommSheet TK-ISS3000-257.xls
                            If bC_PrnByOPOS Then
                                wCNMsgErrSlip.tW_Msg = "PRINT SALES ENDORSEMENT"
                                wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
                                wCNMsgErrSlip.bW_Printing = False
                                wCNMsgErrSlip.bW_Wait = False
                                wCNMsgErrSlip.nW_Line = 0
                                wCNMsgErrSlip.Show vbModal
                                Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ",PRINT SALES ENDORSEMENT(bW_Printing:False)")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                            End If
                            oPrn.LineNo = 0
                            If bC_PrnByOPOS Then
                                oOposSlip.TransactionPrint PTR_S_SLIP, PTR_TP_TRANSACTION 'Initial start  print by OPOS
                                'oOposSlip.RecLineSpacing = 10
                                oPrn.PrinterType = "0"                          '*Ao 2013/03/13  CommSheet TK-ISS3000-038 00.00.01.xlsx
                                'If bVB_CNCheckCV Then
                                '    oPrn.PrinterType = "1"
                                'End If
                            Else
                                oPrn.PrinterType = "1"
                            End If
                            
                            If Not W_PRNbTicketHeaderSlp2 Then GoTo ErrHandler3
                            If Not W_PRNbTicketCVSlp(2, orsRC) Then GoTo ErrHandler3
                            If oPrn.PrinterType <> "0" Then
                                Call W_PRNbTicketFooterBlank                '*Ao 2013/03/13
                            End If
                            oPrn.LineNo = 36
                            bFlagPrn3 = oPrn.EndDoc(oW_Logo, False)
                            If bC_PrnByOPOS Then
                                wCNMsgErrSlip.tW_Msg = ""
                                wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
                                wCNMsgErrSlip.bW_Printing = True
                                wCNMsgErrSlip.bW_Wait = True
                                wCNMsgErrSlip.nW_Line = oPrn.LineNoEx
                                wCNMsgErrSlip.Show vbModal
                                Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ",PRINT SALES ENDORSEMENT(bW_Printing:True)")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                            End If
                            
                            orsRC.MoveNext
                        Loop
                    End If
                    orsRC.Filter = adFilterNone
                        'End If
                    'Else
                    '    If bC_PrnByOPOS Then
                    '        oOpos.TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'Initial start  print by OPOS
                    '        oOpos.RecLineSpacing = 10
                    '    End If
                    '    If Not W_PRNbTicketHeaderSlp2 Then GoTo ErrHandler3
                    '    If Not W_PRNbTicketCVSlp(2) Then GoTo ErrHandler3
                    '    oPrn.LineNo = 36
                    '    bFlagPrn3 = oPrn.EndDoc(oW_Logo, False)
                    'End If
                End If
            End If
        End If
        'Credit Voucher(Return) = 0
        'Credit Voucher(Re-Issue) = 1
        'Select Case nTnsType
        '    Case 16, 17
        '        'none
        '    Case Else
        'If Not bW_PRMGenEJ Then     '*Ao 2013/04/30 CommSheet TK-ISS3000-163-IT2 00.00.01.xls
        If Not bW_PRMGenEJ And Not bW_NoPrintCreditVoucher Then '*Tao 62-09-12 TTK-ISS3000-747 พิมพ์ซ้ำคูปอง ไม่ต้องพิมพ์ Credit Voucher
            If (bW_EnableRePrintCV1 And bW_EnableRePrint) Or bW_EnableRePrint = False Then
                If bW_EnableRePrintCV1 Then Call W_GETxDataToReprint
                If bW_StaPrnCV2 Then
                    If bC_PrnByOPOS Then
                        nLevel = oOposSlip.nCheckHealth(OPOS_CH_INTERNAL)      '*Ao 2013/07/26 CommSheet TK-ISS3000-347-UAT        V 4.0.62
                        If nLevel = OPOS_E_ILLEGAL Then
                            wCNMsgSlip.Show vbModal
                        End If
                    End If
                    'If tW_TypePrinter = 0 Then
                        'If bVB_CNTestSlipPrn Then
                            'แก้ไข Comsheet TK-ISS3000-066
                            If bC_PrnByOPOS Then
'                                If tW_TypePrinter = "0" Or bW_StaPrnCV3 Then                   '*Ao 2013/10/03
'                                    wCNMsgErrSlip.bW_Wait = True
'                                Else
'                                    wCNMsgErrSlip.bW_Wait = False
'                                End If
'                                wCNMsgErrSlip.bW_Printing = False
'                                wCNMsgErrSlip.nW_Line = 0
'                                wCNMsgErrSlip.tW_Msg = "PRINT CREDIT VOUCHER"
'                                wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
'                                wCNMsgErrSlip.Show vbModal
                                If (Not bVB_CNCheckCV) And (tW_TypePrinter = "0" Or bW_StaPrnCV3) Then
                                    Select Case SP_VALnText2Double(tW_TransTypeSub)
    '                                    Case nEN_TrnReturn, nEN_TrnReturnTaxInc, nEN_TrnReturnManualTax                    '*Ao 2013/09/09  CommSheet TK-ISS3000-315-UAT-FAPL-092 00.00.02
                                        Case nEN_TrnReturn, nEN_TrnReturnTaxInc, nEN_TrnReturnManualTax, nEN_TrnDepositCancel, nEN_TrnAddDepositCancel, nEN_TrnCancelVoucher
                                            wCNMsgErrSlip.tW_Msg = "PRINT CREDIT VOUCHER"
                                            wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
                                            wCNMsgErrSlip.bW_Printing = False
                                            wCNMsgErrSlip.bW_Wait = False
                                            wCNMsgErrSlip.nW_Line = 0
                                            wCNMsgErrSlip.Show vbModal
                                            Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ",PRINT CREDIT VOUCHER(bW_Printing:False)")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                                        Case Else
                                            If bW_StaPrnCVReIssue Then '*Thirapong 2021-03-27
                                                wCNMsgErrSlip.tW_Msg = "PRINT RE-ISSUE"
                                                wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
                                                wCNMsgErrSlip.bW_Printing = False
                                                wCNMsgErrSlip.bW_Wait = False
                                                wCNMsgErrSlip.nW_Line = 0
                                                wCNMsgErrSlip.Show vbModal
                                                Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ",PRINT RE-ISSUE(bW_Printing:False)")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                                            End If
                                            Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ",EN_ReIssueNo :" & oW_ReIssue.tEN_ReIssueNo & ",cEN_ReIssueAmt :" & oW_ReIssue.cEN_ReIssueAmt & ",bW_StaPrnCV2:" & bW_StaPrnCV2 & ",bW_StaPrnCVReIssue:" & bW_StaPrnCVReIssue)
                                    End Select
                                End If
                            End If
                            oPrn.LineNo = 0
                            If bC_PrnByOPOS Then
                                oOposSlip.TransactionPrint PTR_S_SLIP, PTR_TP_TRANSACTION 'Initial start  print by OPOS

                                Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ",Initial start OPOS(2))")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง

                                'oOposSlip.RecLineSpacing = 10
                                oPrn.PrinterType = "0"                          '*Ao 2013/03/13  CommSheet TK-ISS3000-038 00.00.01.xlsx
                                'If bVB_CNCheckCV Then
                                '    oPrn.PrinterType = "1"
                                'End If
                            Else
                                oPrn.PrinterType = "1"
                            End If
                            
                            If Not W_PRNbTicketHeaderSlp Then GoTo ErrHandler3
                            If Not W_PRNbTicketCVSlp(0, orsRC) Then GoTo ErrHandler3
                            If oPrn.PrinterType <> "0" Then
                                Call W_PRNbTicketFooterBlank            '*Ao 2013/03/13
                            End If
                            oPrn.LineNo = 36
                            bFlagPrn2 = oPrn.EndDoc(oW_Logo, False)
                            SleepEx 1000, 0 'Msgbox "PrintCV"  '*BG 57-04-17 CommSheet TK-ISS3000-480-FAPL-196_Bung.xlsx  กระตุ้นโปรแกรมให้ทำงาน
                            If bC_PrnByOPOS Then
                                '*Thirapong 2021-03-27
'                                If cVB_CVAmtEx > 0 Then 'TU 2020-07-24 TK ISSUE - Credit Voucher Balance Zero Amount Shown In Re-issue on v5.2002.1'เมื่อ Printer Error แล้ว ห้าม พิมพ์ ยอด RE-ISSUE ที่ <= 0
                                    wCNMsgErrSlip.tW_Msg = ""
                                    wCNMsgErrSlip.tW_TransTypeSub = tW_TransTypeSub            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
                                    wCNMsgErrSlip.bW_Printing = True
                                    wCNMsgErrSlip.bW_Wait = True
                                    wCNMsgErrSlip.nW_Line = oPrn.LineNoEx
                                    wCNMsgErrSlip.Show vbModal
                                    Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & "Slip printing in progress(2)")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
'                                End If
                            End If
                        'End If
                    'Else
                    '    If bC_PrnByOPOS Then
                    '        oOpos.TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'Initial start  print by OPOS
                    '        oOpos.RecLineSpacing = 10
                    '    End If
                    '    If Not W_PRNbTicketHeaderSlp Then GoTo ErrHandler3
                    '    If Not W_PRNbTicketCVSlp(0) Then GoTo ErrHandler3
                    '    oPrn.LineNo = 36
                    '    bFlagPrn2 = oPrn.EndDoc(oW_Logo, False)
                    'End If
                End If
            End If
        End If
        
        Select Case SP_VALnText2Double(tW_TransTypeSub)         '*Ao 2013/05/28  CommSheet TK-ISS3000-190-ST 00 00 01.xls
            Case nEN_TrnReturn, nEN_TrnSale, nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnAddDepositCancel, nEN_TrnDepositCancel, nEN_TrnRedeem, nEN_TrnTaxExceptSale, nEN_TrnSaleTaxInc, nEN_TrnSaleManualTax, nEN_TrnReturnTaxInc, nEN_TrnReturnManualTax '*Eaw 56-07-16 CommSheet TK-ISS3000-276 TR FAPL-055.xls เพิ่มจากอันเดิมที่มีแค่ nEN_TrnReturn เท่านั้น,'*BG 56-08-15 CommSheet TK-ISS3000-311-UAT-FAPL-088.xls
                If Not oOpos Is Nothing Then
                    If bVB_OPOSCashDrawerEnable Then
                        If oOpos.DrawerOpen Then
                            wMsgDrawer.Show vbModal
                            DoEvents
                        End If
                    End If
                End If
        End Select
        'End Select
        
        Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

ErrHandler:
        '***KT 51-11-22 เมื่อมี Error ควรพิมพ์ Slip ออกมาไม่นั้นจะไม่ทราบคิดว่ามีการกระโดดของบิล
        'ให้อยู่หลังการทำ EJ
        '--------------------------------------------------------------------------------------------------------------------------------------
        '*KT 54-04-28 ป้องกันการพิมพ์ Slip ซ้ำต่อเนื่องกัน
        If Not bFlagPrn Then
            Select Case nTnsType
                 Case 1, 5, 8, 9, 14, 17, 25, 26, 27, 28, 29, 30, 31, 32, 34, 35, 36, 37, 40, 41, 42, 45, 46, 47, 48, 49 '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls เพิ่ม 48,49
                         bFlagPrn = oPrn.EndDoc(oW_Logo, False)
                 Case Else
                         '*KT 54-07-25 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon
                         '*KT 54-11-14 V.4.0.4 DCR_5411001  ต้องการให้ Slip กับ Promotion Coupon เป็นแผ่นเดียว เนื่องจากกลัวว่าลูกค้าจะทำ Promotion Coupon หาย
                         If Not bW_StaPrnPmtCpn Then
                             Select Case nTnsType
                                Case 3, 6, 10, 11, 13, 38
                                    If Not orsPmtCpn Is Nothing Then
                                        If orsPmtCpn.State = adStateOpen Then
                                            If Not (orsPmtCpn.BOF And orsPmtCpn.EOF) Then
                                                Call W_PRNbTicketPmtCpn
                                            End If
                                        End If
                                    End If
                            End Select
                        End If
'                        If Not orsPmtCpn Is Nothing Then
'                            If Not (orsPmtCpn.BOF And orsPmtCpn.EOF) Then
'                                bFlagPrn = oPrn.EndDoc(oW_Logo, oOpos.DrawerOpen, False)
'                            Else
'                                 bFlagPrn = oPrn.EndDoc(oW_Logo, oOpos.DrawerOpen, True)
'                            End If
'                        Else
                            bFlagPrn = oPrn.EndDoc(oW_Logo, oOpos.DrawerOpen, True)
'                        End If
             End Select
        End If
        
        '*KT 54-07-25 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon
'        If Not bW_StaPrnPmtCpn Then
'             Select Case nTnsType
'                Case 3, 6, 10, 11, 13, 38
'                    'If Not bW_EnableRePrint Then
'                        If Not orsPmtCpn Is Nothing Then
'                            If Not (orsPmtCpn.BOF And orsPmtCpn.EOF) Then
'                                Call W_PRNbTicketPmtCpn
'                            End If
'                        End If
'                    'End If
'            End Select
'        End If
        
        '--------------------------------------------------------------------------------------------------------------------------------------
        If bFlagEJ Then
        'EJ - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -
                If bW_EnableEJ Then
On Error GoTo ErrHandler2
'PH 2.0 ML-SF-019 EJ missing  บางครั้ง EJ หาย โดยที่ไม่มีทั้งบน Server และที่ เครื่อง POS (Mall5 3/06/09 POS no. 4410 /5196  Mall2 2/06/09 POS 1107/30551,POS 3305/37061)
                        'create  folder by date
                        'สร้างข้อมูลลง EJ ที่ Local ก่อน
                        If bFileOpen Then
                            Close #1
                            bFileOpen = False
                        End If
                         tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\" 'หา Path ของการสร้าง ej ที่ Loacal (folder = "POSEJ")
                         If Not SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then 'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                                tTmp = App.Path & "\POSEJ\" 'หา Path ของการสร้าง ej ที่ Loacal ใหม่
                                Call SP_FLDbManFolder(tTmp, nEN_Fld3Create) 'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                         End If
                        tTmp = tTmp & Format(dW_TransDate, "yyMMdd") & "\"  'หา Path ของการสร้าง ej ที่ Loacal (folder = "POSEJ") ที่เป็น folder ตามวันที่
                        Call SP_FLDbManFolder(tTmp, nEN_Fld3Create) 'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                         'get and write ej
                        tFileEj = tTmp & Format(tW_DepID, "00") & Format(tW_TmnNum, "00000") & Format(tW_TransNo, "00000") & tW_StoreID & ".EJ" 'หา Path ของ file ej
                        If SP_FLEbManFile(tFileEj, nEN_Fle1Exist) Then 'ตรวจสอบว่า file ej มีหรือไม่
                            Call SP_FLEbManFile(tFileEj, nEN_Fld2Delete) 'พบ file ej ให้ทำการลบทิ้ง
                        End If
                        tW_GETFileEJName = tFileEj
                        tTmp = oPrn.EJText
                        Open tFileEj For Output As #1
                        Print #1, tTmp
                        Close #1
                        
                        tTmp = tW_EJFolder & Format(dW_TransDate, "yyMMdd") & "\" 'หา Path ของการสร้าง ej ที่ server (folder = "POSEJ") ที่เป็น folder ตามวันที่
                        If SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then 'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                                Set oFile = oFSO.GetFile(tFileEj)           '*Ao 53-03-10 Pae บอกใช้คำสั่งผิดให้แก้เป็น GetFile
'                                Set oFile = oFSO.OpenTextFile(tFileEJ)
                                 If oFSO.GetParentFolderName(oFile.Path) <> tTmp Then 'ตรวจสอบว่า Path Local กับ Server เป็น Path เดียวกันหรือไม่ กรณีที่เป็น path เดียวกันจะไม่ทำการลบ file
                                    Call SP_FLEbManFile(tTmp & oFile.Name, nEN_Fld2Delete) 'พบ file ej ให้ทำการลบทิ้ง
                                End If
                                If SP_FLEbManFile(oFile.Path, nEN_Fle1Exist) Then 'ตรวจสอบว่า file ej ที่ local มีหรือไม่
                                    oFile.Attributes = Normal 'set file ไม่ให้ read only
'                                    oFSO.CopyFile oFile.Path, tTmp 'copy ไปที่ server
                                    If W_CHKbCheckCopy(oFile.Path, tTmp) Then                   '*Ao 2013/03/07   ERROR แล้วไม่เข้า ErrHandler2
                                        
                                    Else
                                        Call SP_CRTxLogFile("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Events:PrintOut" & " Type:Write Print Ticket : Terminal No. " & tW_TmnNum & " Trans No. " & tW_TransNo & "Copy to Server Not Success") '*KT 52-12-21 ทำการเก้บ Log การเกิด Error ที่เกี่ยวข้องการการทำ Dat File และ EJ
                                    End If
                                End If
                                'PH 2.0 ML-SF-018 Backup EJ File ไว้ที่ ....\\POSEJ\Backup\YYMMDD
                                tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\Backup\" 'หา Path ของการสร้าง ej ที่ local (folder = "POSEJ") ที่เป็น folder Backup
                                If Not SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then  'ตรวจสอบว่า folder Backup มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                                   tTmp = App.Path & "\POSEJ\Backup\" 'หา Path ของการสร้าง ej ที่ Loacal Backup ใหม่
                                   Call SP_FLDbManFolder(tTmp, nEN_Fld3Create) 'ตรวจสอบว่า folder backup มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                                End If
                                tTmp = tTmp & Format(dW_TransDate, "yyMMdd") & "\" 'หา Path ของการสร้าง ej ที่ Loacal Backup(folder = "POSEJ") ที่เป็น folder ตามวันที่
                                Call SP_FLDbManFolder(tTmp, nEN_Fld3Create) 'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
                                If SP_FLEbManFile(oFile.Path, nEN_Fle1Exist) Then 'ตรวจสอบว่า file ej ใน local มีหรือไม่
                                    oFile.Attributes = Normal 'set file ไม่ให้ read only
                                    Call SP_FLEbManFile(tTmp & oFile.Name, nEN_Fld2Delete) 'ตรวจสอบว่า file ej ใน local Backup มีหรือไม่ ถ้ามีให้ลบทิ้ง
                                    oFSO.CopyFile oFile.Path, tTmp 'copy ไปที่ local backup
                                End If
                                'ลบ File EJ ที่เป็น Temp ทิ้ง
                                tW_GETFileEJName = tTmp & oFile.Name
                                If SP_FLEbManFile(tTmp & oFile.Name, nEN_Fle1Exist) Then 'ตรวจสอบว่า file ej ใน local มีหรือไม่
                                    oFile.Attributes = Normal 'set file ไม่ให้ read only
                                    oFSO.DeleteFile oFile.Path 'ลบทิ้ง
                                End If
                        Else 'ส่ง File EJ ไปที่ Server ไม่สำเร็จ
                                Set oFile = oFSO.GetFile(tFileEj)
                                'PH 2.0 ML-SF-018 Backup EJ File ไว้ที่ ....\\POSEJ\Backup\YYMMDD
                                tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\Backup\"
                                If Not SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then
                                   tTmp = App.Path & "\POSEJ\Backup\"
                                   Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)
                                End If
                                tTmp = tTmp & Format(dW_TransDate, "yyMMdd") & "\"
                                Call SP_FLDbManFolder(tTmp, nEN_Fld3Create)
                                oFile.Attributes = Normal
                                Call SP_FLEbManFile(tTmp & oFile.Name, nEN_Fld2Delete)
                                oFSO.CopyFile oFile.Path, tTmp
                        End If
                        Set oFSO = Nothing
                End If
        ' - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -  - - - - - - - -
        End If
        
        '*KT 54-04-28 ป้องกันการพิมพ์ Slip ซ้ำต่อเนื่องกัน ถ้าคำสั่ง EndDoc ไม่สำเร็จให้ทำการ ClearInput
        If Not bFlagPrn Then
            MsgBox "Print not success.Goto function Reprint Receipt"
            oPrn.ClearBufferOPOSPrn
        End If
       
        oPrn.KillDoc
        Call SP_CRTxLogFile("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Events:PrintOut" & " Type:Write Print Ticket : Terminal No. " & tW_TmnNum & " Trans No. " & tW_TransNo) '*KT 52-12-21 ทำการเก้บ Log การเกิด Error ที่เกี่ยวข้องการการทำ Dat File และ EJ
        If bFlagEJ And bW_EnableEJ Then Call SP_CRTxLogFile("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Events:PrintOut" & " Type:Write EJ Ticket : Terminal No. " & tW_TmnNum & " Trans No. " & tW_TransNo) '*KT 52-12-21 ทำการเก้บ Log การเกิด Error ที่เกี่ยวข้องการการทำ Dat File และ EJ
        Exit Function
ErrHandler2:
        '*KT 54-04-28 ป้องกันการพิมพ์ Slip ซ้ำต่อเนื่องกัน ถ้าคำสั่ง EndDoc ไม่สำเร็จให้ทำการ ClearInput
        If Not bFlagPrn Then
             MsgBox "Print not success.Goto function Reprint Receipt"
            oPrn.ClearBufferOPOSPrn
        End If
        oPrn.KillDoc
        Call SP_CRTxLogFile("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Events:PrintOut" & " Type:Write Print Ticket : Terminal No. " & tW_TmnNum & " Trans No. " & tW_TransNo) '*KT 52-12-21 ทำการเก้บ Log การเกิด Error ที่เกี่ยวข้องการการทำ Dat File และ EJ
        If bFlagEJ And bW_EnableEJ Then Call SP_CRTxLogFile("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Events:PrintOut" & " Type:Write EJ Ticket : Terminal No. " & tW_TmnNum & " Trans No. " & tW_TransNo) '*KT 52-12-21 ทำการเก้บ Log การเกิด Error ที่เกี่ยวข้องการการทำ Dat File และ EJ
        Exit Function
ErrHandler3:
        If Not bFlagPrn3 Then
            bFlagPrn3 = oPrn.EndDoc(oW_Logo, False)
        End If
        If Not bFlagPrn2 Then
            bFlagPrn2 = oPrn.EndDoc(oW_Logo, False)
        End If
        
End Function

Private Sub Class_Initialize()
        cW_PageWidth = 5  'defaule 5 cm.
        cW_MrakOffSet = 0.3  'hardcode position to mark DT Sign
        bW_EnableEJ = True
        
End Sub

Private Function W_GETtDateTimeFmt() As String
Dim orsTmp As ADODB.Recordset
Dim tSql As String

Const tFmt1 = "yyyy/mm/dd"
Const tFmt2 = "dd/mm/yyyy"
 
' tSql = "SELECT FTStpSelValue FROM TCNMStorePara WHERE FTStpCode ='02' "
' Set orsTmp = ocnAll.Execute(tSql)
' If Not orsTmp.EOF Then
'        If SP_FEDtChkString(orsTmp, "FTStpSelValue") = "1" Then
'                W_GETtDateTimeFmt = tFmt1
'        Else
'                W_GETtDateTimeFmt = tFmt2
'        End If
'Else
'                W_GETtDateTimeFmt = tFmt1
' End If

'*KT 51-05-22 Comsheet ML-SF-122
Select Case Mid(tW_OptAlw, 1, 1)
    Case "0" 'yyyy/mm/dd
        W_GETtDateTimeFmt = tFmt2
    Case "1" 'dd/mm/yyyy
        W_GETtDateTimeFmt = tFmt1
    Case Else 'default yyyy/mm/dd
        W_GETtDateTimeFmt = tFmt1
End Select
End Function

Private Function W_DATtGetDetail(ByVal ptCode As String, ByVal ptFedCode As String, ByVal ptTable As String, ByVal ptFedDetail As String) As String
Dim tSql As String
Dim orsTmp As ADODB.Recordset
Dim tTmp As String

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler

            tSql = "SELECT " & ptFedDetail & " FROM " & ptTable & " WHERE " & ptFedCode & "='" & ptCode & "'"
            'Set orsTmp = ocnAll.Execute(tSql)
            If SP_SQLvExecute(tSql, orsTmp, oDB.ocnOnLine, , , , , True, True) = 0 Then
                If Not (orsTmp.BOF And orsTmp.EOF) Then
                    tTmp = SP_FEDtChkString(orsTmp, ptFedDetail)
                End If
            Else
                tTmp = ""
            End If
'Return
W_DATtGetDetail = tTmp

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function SP_FEDtChkString(ByVal poRs As ADODB.Recordset, ByVal ptField$, Optional pvRetOld) As String
'----------------------------------------------------------
'   Call:   poRs as ADODB.Recordset
'               ptField is field to check
'               pvRetOld(variant) if empty, or null want to return old
'   Ret:    variant, only text and numeric
'----------------------------------------------------------
    Dim bEmpty As Boolean
    
    If poRs.BOF And poRs.EOF Then
        bEmpty = True
    Else
        If IsNull(poRs.Fields(ptField)) Then
            bEmpty = True
        Else
            If Len(Trim(poRs.Fields(ptField))) = 0 Then
                bEmpty = True
            Else
                SP_FEDtChkString = poRs.Fields(ptField)
            End If
        End If
    End If
    If bEmpty Then
        If IsMissing(pvRetOld) Then
            SP_FEDtChkString = ""
        Else
            SP_FEDtChkString = pvRetOld
        End If
    End If
End Function

Private Function SP_FEDdChkDate(ByVal poRs As ADODB.Recordset, ByVal ptField$, Optional pvRetOld) As Date
'----------------------------------------------------------
'   Call:   poRs as ADODB.Recordset
'               ptField is field to check
'               pvRetOld(variant) if empty, or null want to return old
'   Ret:    numeric value as variant(integer,long,single,double ...)
'----------------------------------------------------------
    Dim bEmpty As Boolean
    
    If poRs.BOF And poRs.EOF Then
        bEmpty = True
    Else
        If IsNull(poRs.Fields(ptField)) Then
            bEmpty = True
        Else
            'If Len(Trim(poRs.Fields(ptField))) = 0 Then
            '    bEmpty = True
            'Else
                SP_FEDdChkDate = poRs.Fields(ptField)
            'End If
        End If
    End If
    If bEmpty Then
        If IsMissing(pvRetOld) Then
            SP_FEDdChkDate = Date
        Else
            SP_FEDdChkDate = pvRetOld
        End If
    End If
End Function

Private Function SP_FEDdChkDateTime(ByVal poRs As ADODB.Recordset, ByVal ptField$, Optional pvRetOld) As Date
'----------------------------------------------------------
'   Call:   poRs as ADODB.Recordset
'               ptField is field to check
'               pvRetOld(variant) if empty, or null want to return old
'   Ret:    numeric value as variant(integer,long,single,double ...)
'----------------------------------------------------------
    Dim bEmpty As Boolean
    
    If poRs.BOF And poRs.EOF Then
        bEmpty = True
    Else
        If IsNull(poRs.Fields(ptField)) Then
            bEmpty = True
        Else
            'If Len(Trim(poRs.Fields(ptField))) = 0 Then
            '    bEmpty = True
            'Else
                SP_FEDdChkDateTime = Format(poRs.Fields(ptField).Value, "YYYY/MM/DD HH:MM:SS")
            'End If
        End If
    End If
    If bEmpty Then
        If IsMissing(pvRetOld) Then
            SP_FEDdChkDateTime = Date
        Else
            SP_FEDdChkDateTime = pvRetOld
        End If
    End If
End Function

Private Function SP_FEDvChkNumeric(ByVal poRs As ADODB.Recordset, ByVal ptField$, Optional pvRetOld) As Variant
'----------------------------------------------------------
'   Call:   poRs as ADODB.Recordset
'               ptField is field to check
'               pvRetOld(variant) if empty, or null want to return old
'   Ret:    numeric value as variant(integer,long,single,double ...)
'----------------------------------------------------------
    Dim bEmpty As Boolean
    
    If poRs.BOF And poRs.EOF Then
        bEmpty = True
    Else
        If IsNull(poRs.Fields(ptField)) Then
            bEmpty = True
        Else
            If Len(Trim(poRs.Fields(ptField))) = 0 Then
                bEmpty = True
            Else
                SP_FEDvChkNumeric = poRs.Fields(ptField)
            End If
        End If
    End If
    If bEmpty Then
        If IsMissing(pvRetOld) Then
            SP_FEDvChkNumeric = 0
        Else
            SP_FEDvChkNumeric = pvRetOld
        End If
    End If
End Function


Private Function W_GETbSplitTransName(ByVal ptText As String, ByRef ptRet1st As String, ByRef ptRet2nd As String) As Boolean
Dim nTmp As Long
Dim cVatRateDT As Double
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
        'Line_Chars
        '*KT 51-04-19 แยก Line จาก ';'
        nTmp = InStr(1, ptText, ";", vbTextCompare)
        nTmp = IIf(nTmp = 0, 9, nTmp)
        
        '*Amorn 65-07-11 เพิ่มข้อความ Return 7%
        'cVB_CNGSTRate
      '  If Mid(ptText, 8, 1) = cVB_OldGstRate Then
        If bVB_ChkRetGST7 Then '*KT 66-08-03   TTK-ISS3000-829
'        If bVB_ChkRetGST7 = True Then
'        If bVB_ChkRetGST7 = True Then '*Amorn 2023-02-20 TTK-ISS3000-818 TK ISSUE -Pickup layaway Vat 9_ GST show wording wrong 7_ GST
        
            'ptRet1st = Replace(Mid(Replace(Trim(Mid(ptText, 1, nTmp)), ";", ""), 1, 10), " ", "")   '*KT 66-08-03   TTK-ISS3000-829
            'ptRet1st = cVB_OldGstRate & "%" & Mid(Replace(Trim(Mid(ptRet1st, 1, nTmp)), ";", ""), 1, 6)   '*KT 66-08-03   TTK-ISS3000-829
            
            ptRet1st = Replace(Mid(Replace(Trim(Mid(ptText, 1, nTmp)), ";", ""), 1, 10), " ", "")  '*KT 66-08-03   TTK-ISS3000-829

        Else
            ptRet1st = Mid(Replace(Trim(Mid(ptText, 1, nTmp)), ";", ""), 1, 7)    '*ของเก่า
            ptRet2nd = Replace(Trim(Mid(ptText, nTmp + 1, 9)), ";", "")              '*ของเก่า
            
            Select Case Val(tW_TransTypeSub)
                Case nEN_TrnRedeem, nEN_TrnDepositCancel, nEN_TrnAddDeposit, nEN_TrnAddDepositCancel
                    If W_CHKbOldVatRate(cVatRateDT) = True Then
                        '*KT 66-08-28 CFMPOS-External Specification-Enhance SCO Ph1-00.02.00.docx
                        If bW_EnableRePrint Then
                            ptRet2nd = ptRet2nd & IIf(cW_HDRate = 0, cVatRateDT, cW_HDRate) & "%"
                        Else
                            ptRet2nd = ptRet2nd & IIf(cVB_OldGstRate = 0, cVatRateDT, cVB_OldGstRate) & "%"
                        End If
                       '*KT 66-08-30  CommSheet TTK-ISS3000-829_SIT-005 Wording show Pickup 0_00.01.00_pae_.xlsx
                        If cW_HDRate = 0 Then cW_HDRate = cVatRateDT
                        If cVB_OldGstRate = 0 Then cVB_OldGstRate = cVatRateDT
                    End If
            End Select

        End If
'Return
'*KT 66-08-03   TTK-ISS3000-829
'    '*Amorn 65-07-11 เพิ่มข้อความ Return 7%
'        If bVB_ChkRetGST7 = True And tVB_CNMnuCurrent = 1723 Then
'            ptRet1st = ptRet1st & cVB_OldGstRate & "%"
'        End If
'   '*Amorn 65-07-11 เพิ่มข้อความ Return 7%
W_GETbSplitTransName = True

Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function W_PRNbTicketFreeGift() As Boolean
Dim nTmp As Long, tTmp As String
Dim cWidthLimit As Double

'Const tTkFreeGift = "Free Gift"
Dim tTkFreeGift
tTkFreeGift = W_GETxVocabularyTicketFromArray("1952")
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler


                        If Not orsCD Is Nothing And orsCD.State = adStateOpen Then
                             orsCD.Filter = "FNDctNo = 16 " 'free gift
                             If Not orsCD.EOF Then
                                     
                                     cWidthLimit = cW_PageWidth - cXC
                                     
                                     orsCD.Sort = "FNScdSeqNo ASC"
                                     Set oPrn.Font = oW_Font(1)
                                     Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
                                     Call oPrn.PrintText(tTkFreeGift, cXL, , vbLeftJustify)
                                     
                                     Do While Not orsCD.EOF
                                     
                                             nTmp = SP_FEDvChkNumeric(orsCD, "FCScdQty")
                                             'Call oPrn.PrintText(nTmp & "@", cXHL, , vbLeftJustify, False)
                                             Call oPrn.PrintText(nTmp & "@", IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False)
                                             tTmp = SP_FEDvChkNumeric(orsCD, "FTDctName")
                                             Call oPrn.PrintText(tTmp, cXC, , vbLeftJustify, True, , cWidthLimit, True)
        
                                            orsCD.MoveNext
                                     Loop
                                     
                             End If
                         
                            orsCD.Filter = vbNullString
                        
                        End If

        'Return
        W_PRNbTicketFreeGift = True
        
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        'Return
        W_PRNbTicketFreeGift = True
End Function

Public Function W_PRNbTicketPicture() As Boolean
Dim nTmp As Long, tTmp As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler

                
        If Not oW_Logo Is Nothing Then oPrn.PrintPicture oW_Logo, cXC, 0, vbCenter
    
'Return
W_PRNbTicketPicture = True

Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function W_PRNbTicketModeStart() As Boolean
Dim nTmp As Long, tTmp As String

'Line_Chars
'Const tTkMdTraining = "** TRAINING MODE **"
Dim tTkMdTraining$
tTkMdTraining = Mid(W_GETxVocabularyTicketFromArray("1961"), 1, 17)
'Const tTkMdVoid = "** VOID VOID VOID **"""
Dim tTkMdVoid$
'tTkMdVoid = Mid(W_GETxVocabularyTicketFromArray("1962"), 1, 17)
tTkMdVoid = W_GETxVocabularyTicketFromArray("1962")
'Const tTkMdSuspend = "SUSPENDED TRANSACT'N"
Dim tTkMdSuspend$
tTkMdSuspend = Mid(W_GETxVocabularyTicketFromArray("1973"), 1, 17)
'Const tTkMdRecall = "RECALLED TRANSACTION"
Dim tTkMdRecall$
tTkMdRecall = Mid(W_GETxVocabularyTicketFromArray("1974"), 1, 17)
'Const tTkMdDulicate = "** DUPLICATE **"
Dim tTkMdDulicate$
'tTkMdDulicate = Mid(W_GETxVocabularyTicketFromArray("1975"), 1, 17)
tTkMdDulicate = W_GETxVocabularyTicketFromArray("1975")
Dim tTkRate As String
tTkRate = Mid(W_GETxVocabularyTicketFromArray("1998"), 1, 17)                               '*Ao 2012/07/25  PRJ10001TK-External Specification Open Trx-00.01.00.doc
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
        
         're Print
         If nW_PrintCnt > 0 And Not bW_PRMGenEJ Then   '*KT 51-12-25 ตรวจสอบเงื่อนไขว่าใช้ Tools Generate EJ หรือไม่
                Select Case Val(tW_TransType)
                    Case 37, 41, 46 'Over/Short
                        If nW_PrintCnt > 1 Then
                             Set oPrn.Font = oW_Font(2)
            '                 Call oPrn.PrintText(tTkMdDulicate, cXC, , vbCenter, , nEN_PrinterOnly, , , True)      'Ao 2012/07/25 BUG
                             Call oPrn.PrintText(tTkMdDulicate, cXC, , vbCenter, , 1, , , True)
                        End If
                    Case Else
                         Set oPrn.Font = oW_Font(2)
        '                 Call oPrn.PrintText(tTkMdDulicate, cXC, , vbCenter, , nEN_PrinterOnly, , , True)      'Ao 2012/07/25 BUG
                         Call oPrn.PrintText(tTkMdDulicate, cXC, , vbCenter, , 1, , , True)
                End Select
         End If
         
         'trainingmode
        Select Case Val(tW_TransType)
            Case 25
                 If bW_TrainMode Then
                         Set oPrn.Font = oW_Font(2)
                         Call oPrn.PrintText(tTkMdTraining, cXC, , vbCenter, , , , , True)
                         bW_PrnTraining = True
                 End If
         End Select
         
        'void
        Select Case Val(tW_TransType)
'            Case 5, 14, 16, 17, 20, 21, 26, 27, 28, 32, 33
            Case 5, 14, 17, 21, 26, 27, 28, 32, 33   '*Eaw 56-03-26  เอา Type 20,16 ออก CommSheet TK-ISS3000-110 00.00.01.xlsx
                 Set oPrn.Font = oW_Font(2)
                 Call oPrn.PrintText(tTkMdVoid, cXC, , vbCenter, , , , , True)
                 bW_PrnVoid = True
        End Select
        
        'suspend
        Select Case Val(tW_TransType)
            Case 29
                 Set oPrn.Font = oW_Font(2)
                 Call oPrn.PrintText(tTkMdSuspend, cXC, , vbCenter, , , , , True)
        End Select
                        
          'recall
        If bW_RecallMode Then
                Set oPrn.Font = oW_Font(2)
                Call oPrn.PrintText(tTkMdRecall, cXC, , vbCenter, , , , , True)
        End If
        
        Select Case Val(tW_TransType)           '*Ao 2012/07/25  PRJ10001TK-External Specification Open Trx-00.01.00.doc
            Case 1
                 Set oPrn.Font = oW_Font(2)
                 Call oPrn.PrintText(tTkRate, cXC, , vbCenter, , , , , True)
        End Select
        
        Select Case Val(tW_TransType)           '*Ao 2012/07/25  VOID
'            Case 3, 4, 6, 11, 13, 22, 23, 30
'            Case 3, 4, 6, 11, 13, 22, 23, 30, 38, 39, 43, 44 '*Eaw 56-06-27  เพิ่ม Type 38,39,43,44 CommSheet TK-ISS3000-242-TR.xlsx
            Case 3, 4, 6, 11, 13, 19, 22, 23, 30, 38, 39, 43, 44                    '*Ao 2014/07/23
                If UCase(tW_StaCancel) = "Y" Then
                    Set oPrn.Font = oW_Font(2)
                    Call oPrn.PrintText(tTkMdVoid, cXC, , vbCenter, , , , , True)
                End If
        End Select
        
        'Return
        W_PRNbTicketModeStart = True
    
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
         'Return
        W_PRNbTicketModeStart = True
End Function

Private Function W_PRNbTicketTransReEntry() As Boolean
    Dim tTmp As String
    Dim cWidthLimit As Double
    '*Eaw 56-07-03
    'tTkTaxIns = "Sale Adjustment (Tax Include)"
    Dim tTkTaxInc As String
    tTkTaxInc = W_GETxVocabularyTicketFromArray("14")
'    tTkTaxMan="Sale Adjustment (Manual Tax)"
    Dim tTkTaxMan As String
    tTkTaxMan = W_GETxVocabularyTicketFromArray("15")
    Dim tTkDate As String
    tTkDate = W_GETxVocabularyTicketFromArray("16")
    Dim tTkTmn As String
'    tTkTmn = W_GETxVocabularyTicketFromArray("17")
    tTkTmn = W_GETxVocabularyTicketFromArray("21") '*BG 56-09-23 Comsheet-356 Print EJ tenant sale
    Dim tTkReEntry As String
    tTkReEntry = W_GETxVocabularyTicketFromArray("18")
    '------------------------------------------------------
    Dim tTkStartBill As String
    tTkStartBill = W_GETxVocabularyTicketFromArray("19")
    Dim tTkAcc As String
    tTkAcc = W_GETxVocabularyTicketFromArray("20")
On Error GoTo ErrHandler
    Set oPrn.Font = oW_Font(1)
    Select Case Val(tW_TransType)
            Case 38, 39, 40, 41, 42, 48 '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls เพิ่ม 48
                    cWidthLimit = cW_PageWidth - cXC
'                    Call oPrn.PrintText("Sale Adjustment (Tax Include)", cXL, , vbLeftJustify, True)
'                    Call oPrn.PrintText("Date", cXL, , vbLeftJustify, False)
                    Call oPrn.PrintText(tTkTaxInc, cXL, , vbLeftJustify, True)  '*Eaw 56-07-03
                    '*Eaw 56-08-29
                    Call oPrn.PrintText("", cXR, , vbRightJustify, True)
                    Call oPrn.PrintText(tTkAcc, cXL, , vbLeftJustify, False)
                    tTmp = tW_AccTypeReEntry
                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True)
                    Call oPrn.PrintText(tTkDate, cXL, , vbLeftJustify, False) '*Eaw 56-07-03
                    tTmp = Format(dW_TransDate, W_GETtDateTimeFmt)
                    tTmp = Replace(tTmp, "-", "/")
                    tTmp = Replace(tTmp, ".", "/")
'                   Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, True, , cWidthLimit)
'                    tTmp = tTmp & "    "
                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True) '*Eaw 56-07-25 CommSheet TK-ISS3000-282-TR.xls
'                    Call oPrn.PrintText("Physical Terminal", cXL, , vbLeftJustify, False)
                    Call oPrn.PrintText(tTkTmn, cXL, , vbLeftJustify, False) '*Eaw 56-07-03
'                    Call oPrn.PrintText(tW_TmnNum, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, True, , cWidthLimit)
'                    tW_PhyTmnNum = tW_PhyTmnNum & "         "
'                    Call oPrn.PrintText(tW_PhyTmnNum, cXMK, , vbRightJustify, True) '*Eaw 56-07-25 CommSheet TK-ISS3000-282-TR.xls
                    '*Eaw 56-08-14 ใช้ tTmp แทน tW_PhyTmnNum
'                    tTmp = tW_PhyTmnNum
                    'tTmp = tW_VirTerminal  '*BG 56-09-23  Comsheet-356 Print EJ
                    tTmp = SP_FEDtChkString(orsHD, "FTShdPOSCN") '*TON 59-01-07 TK-ISS3000-607
'                    tTmp = tTmp & "         "

                    '*Thirapong 2020-10-12 CommSheet TTK-ISS3000-769 TK ISSUE - Re-entry Closing Report printed wrong Virtual POSID.xlsx
                    If tTmp <> tVB_CNVirTerminalNum Then tTmp = tVB_CNVirTerminalNum
                    '*****************
                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True) '*Eaw 56-07-25 CommSheet TK-ISS3000-282-TR.xls
                    Select Case Val(tW_TransType) '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
                        Case 48
'                            tTmp = "Starting Transaction Number"
                            Call oPrn.PrintText(tTkStartBill, cXL, , vbLeftJustify, False)
                            tTmp = tW_TransNo
'                            tTmp = tTmp & "         "
                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True)
                    End Select
                    Call oPrn.PrintText("", cXR, , vbRightJustify, True)
                    Set oPrn.Font = oW_Font(2)
'                    Call oPrn.PrintText("** REENTRY MODE **", cXC, , vbCenter, , , , , True)
                    Call oPrn.PrintText(tTkReEntry, cXC, , vbCenter, , , , , True) '*Eaw 56-07-03
            Case 43, 44, 45, 46, 47, 49 '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls เพิ่ม 49
                    cWidthLimit = cW_PageWidth - cXC
'                    Call oPrn.PrintText("Sale Adjustment (Manual Tax)", cXL, , vbLeftJustify, True)
'                    Call oPrn.PrintText("Date", cXL, , vbLeftJustify, False)
                    Call oPrn.PrintText(tTkTaxMan, cXL, , vbLeftJustify, True) '*Eaw 56-07-03
                    '*Eaw 56-08-29
                    Call oPrn.PrintText("", cXR, , vbRightJustify, True)
                    Call oPrn.PrintText(tTkAcc, cXL, , vbLeftJustify, False)
                    tTmp = tW_AccTypeReEntry
                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True)
                    Call oPrn.PrintText(tTkDate, cXL, , vbLeftJustify, False) '*Eaw 56-07-03
                    tTmp = Format(dW_TransDate, W_GETtDateTimeFmt)
                    tTmp = Replace(tTmp, "-", "/")
                    tTmp = Replace(tTmp, ".", "/")
'                    Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, True, , cWidthLimit)
'                    tTmp = tTmp & "    "
                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True) '*Eaw 56-07-25 CommSheet TK-ISS3000-282-TR.xls
'                    Call oPrn.PrintText("Physical Terminal", cXL, , vbLeftJustify, False)
                    Call oPrn.PrintText(tTkTmn, cXL, , vbLeftJustify, False) '*Eaw 56-07-03
'                    Call oPrn.PrintText(tW_PhyTmnNum, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, True, , cWidthLimit)
'                    tW_PhyTmnNum = tW_PhyTmnNum & "         "
'                    Call oPrn.PrintText(tW_PhyTmnNum, cXMK, , vbRightJustify, True) '*Eaw 56-07-25 CommSheet TK-ISS3000-282-TR.xls
                    '*Eaw 56-08-14 ใช้ tTmp แทน tW_PhyTmnNum
'                    tTmp = tW_PhyTmnNum
                    'tTmp = tW_VirTerminal  '*BG 56-09-24  Comsheet-356 Print EJ
                    tTmp = SP_FEDtChkString(orsHD, "FTShdPOSCN") '*TON 59-01-07 TK-ISS3000-607
                    '*Thirapong 2020-10-12 CommSheet TTK-ISS3000-769 TK ISSUE - Re-entry Closing Report printed wrong Virtual POSID.xlsx
                    If tTmp <> tVB_CNVirTerminalNum Then tTmp = tVB_CNVirTerminalNum
                    If (tTmp = SP_FEDtChkString(orsHD, "FTTmnNum")) And (tVB_VirtualPosNo <> "") Then tTmp = tVB_VirtualPosNo   '*Thirapong 64-06-09 '*Thirapong 64-06-09 CommSheet TTK-ISS3000-789 TK ISSUE - Re-entry Closing Report printed wrong Virtual POSID_7823.xlsx
                    '========================
'                    tTmp = tTmp & "         "
                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True) '*Eaw 56-07-25 CommSheet TK-ISS3000-282-TR.xls
                    Select Case Val(tW_TransType) '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
                        Case 49
'                            tTmp = "Starting Transaction Number"
                            Call oPrn.PrintText(tTkStartBill, cXL, , vbLeftJustify, False)
                            tTmp = tW_TransNo
'                            tTmp = tTmp & "         "
                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True)
                    End Select
                    Call oPrn.PrintText("", cXR, , vbRightJustify, True)
                    Set oPrn.Font = oW_Font(2)
'                    Call oPrn.PrintText("** REENTRY MODE **", cXC, , vbCenter, , , , , True)
                    Call oPrn.PrintText(tTkReEntry, cXC, , vbCenter, , , , , True) '*Eaw 56-07-03
    End Select
    W_PRNbTicketTransReEntry = True
    Exit Function
ErrHandler:
    'Return
    W_PRNbTicketTransReEntry = True
End Function

Private Function W_PRNbTicketModeEnd() As Boolean
Dim nTmp As Long, tTmp As String
'Line_Chars
'Const tTkMdTraining = "** TRAINING MODE **"
Dim tTkMdTraining$
tTkMdTraining = Mid(W_GETxVocabularyTicketFromArray("1961"), 1, 17)

'Const tTkMdVoid = "** VOID VOID VOID **"
Dim tTkMdVoid$
'tTkMdVoid = Mid(W_GETxVocabularyTicketFromArray("1962"), 1, 17)
tTkMdVoid = W_GETxVocabularyTicketFromArray("1962")                     '*Ao 2014/07/22    4.2.140
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
      
               'void
        Select Case Val(tW_TransType)
'            Case 5, 14, 17, 16, 20, 21, 26, 27, 28, 32, 33
            Case 5, 14, 17, 21, 26, 27, 28, 32, 33   '*Eaw 56-03-26  เอา Type 20,16 ออก CommSheet TK-ISS3000-110 00.00.01.xlsx
'                    Set oPrn.Font = oW_Font(2)
'                    Call oPrn.PrintText(tTkMdVoid, cXC, , vbCenter, , , , , True)
                    tW_SlipFooterStr = tW_SlipFooterStr & "4" & tTkMdVoid & ";"
                    bW_PrnVoid = True
        End Select
           
            'trainingmode
        Select Case Val(tW_TransType)
            Case 25
                    If bW_TrainMode Then
'                            Set oPrn.Font = oW_Font(2)
'                            Call oPrn.PrintText(tTkMdTraining, cXC, , vbCenter, , , , , True)
                            tW_SlipFooterStr = tW_SlipFooterStr & "4" & tTkMdTraining & ";"
                            bW_PrnTraining = True
                    End If
        End Select
            
        Select Case Val(tW_TransType)           '*Ao 2012/07/25  VOID
'            Case 3, 4, 6, 11, 13, 22, 23, 30
'            Case 3, 4, 6, 11, 13, 22, 23, 30, 38, 39, 43, 44 '*Eaw 56-06-27  เพิ่ม Type 38,39,43,44 CommSheet TK-ISS3000-242-TR.xlsx
            Case 3, 4, 6, 11, 13, 19, 22, 23, 30, 38, 39, 43, 44                    '*Ao 2014/07/23
                If UCase(tW_StaCancel) = "Y" Then
                    Set oPrn.Font = oW_Font(2)
    '                        Call oPrn.PrintText(tTkMdVoid, cXC, , vbCenter, , , , , True)
                    tW_SlipFooterStr = tW_SlipFooterStr & "4" & tTkMdVoid & ";"
                End If
        End Select
        'Return
        W_PRNbTicketModeEnd = True
        
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        'Return
        W_PRNbTicketModeEnd = True
End Function

Private Function W_PRNbTicketPeople() As Boolean
Dim nTmp As Long, tTmp As String
Dim tRefDeToken As String '*Bmp 61-04-05
'Const tTkCustID = "Customer ID: "
Dim tTkCustID$
tTkCustID = W_GETxVocabularyTicketFromArray("1963")
'Const tTkSalesManID = "SlaesMan ID: "
Dim tTkSalesManID$
tTkSalesManID = W_GETxVocabularyTicketFromArray("1964")
'Const tTkCashierName = "Name: "
Dim tTkCashierName$
tTkCashierName = W_GETxVocabularyTicketFromArray("1965")

'Remain discount
Dim tTkRemainDisc$
tTkRemainDisc = W_GETxVocabularyTicketFromArray("1976")

'*KT 52-02-16 V 1.5.10
Dim tTkMallID$
tTkMallID = W_GETxVocabularyTicketFromArray("1977")
Dim tTkEmpID$
tTkEmpID = W_GETxVocabularyTicketFromArray("1978")
Dim tTkMallIDReg$
tTkMallIDReg = W_GETxVocabularyTicketFromArray("1979")
'*KT 52-04-10 กรณีมีการใช้ Member Price หรือ Member Point ให้พิมพ์ Member ID:  และหากเป็น member ที่ได้ส่วนลดด้วย ก็พิมพ์ แค่ครั้งเดียว ไม่จำเป็นต้องพิมพ์ Customer Id: อีก หากเป็น Card ที่มีส่วนลดอย่างเดียว ก็พิมพ์คำว่า Customer ID:
Dim tTkMemPrice$
tTkMemPrice = W_GETxVocabularyTicketFromArray("1988")

Dim MallCard(3) As String
Dim nI As Long

Dim tMallID As String
Dim tMallTmp  As String

Dim tEmpID As String
Dim tEmpTmp  As String

Dim tEmpName As String  '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount

Dim bPass As Boolean '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
Dim tLimit As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler

                        'People
                            tLimit = "Discountable Limit"               '*Ao 2013/03/06     'E-mail Krittapak 6 มีนาคม 2013 13:31
                            If tTkCustID = "" Then tTkCustID = "Customer ID :" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                            Set oPrn.Font = oW_Font(1)
                            Select Case Val(tW_TransTypeSub)
                            Case 10, 30 'VIP Service/Create VIP Sale
                                If tW_CustomerVIPID <> vbNullString Then
                                       tTmp = tTkCustID & tW_CustomerVIPID
'                                       Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                        tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                                End If
                            Case Else
                                If tW_CustomerID <> vbNullString Then
                                       tTmp = tTkCustID & tW_CustomerID
'                                       Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                       tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                                End If
                            End Select
                            
                            '*KT 52-04-10 กรณีมีการใช้ Member Price หรือ Member Point ให้พิมพ์ Member ID:  และหากเป็น member ที่ได้ส่วนลดด้วย ก็พิมพ์ แค่ครั้งเดียว ไม่จำเป็นต้องพิมพ์ Customer Id: อีก หากเป็น Card ที่มีส่วนลดอย่างเดียว ก็พิมพ์คำว่า Customer ID:
                            If tTkMemPrice = "" Then tTkMemPrice = "Customer ID :" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                            If tW_MemberPriceID <> "" Then
                                tTmp = tTkMemPrice & " " & tW_MemberPriceID
'                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                            End If
                            
                            '*KT 51-08-02
                            'แก้ไข Format
                            'Cash Tender: XXX
                            'Change Tender: XXXX
                            'Customer ID: [VIP Sale]
                            'Customer ID: [Mall Card]
                            'Customer ID: [EmpCard]
                            'Last Point: XXXX
                            'Today Point: XXXX
                            'Saleman ID: XXX
                            'Name: XXX
                            
                            '--------------------------------------------------------------------------------------------------------------------------
                            '*KT 51-09-17 ให้สลับ Mall Card[Discount] ขึ้นก่อน  'Member : Mall Card[Register]
                           'Member : Mall Card[Discount]
                            tMallID = ""
                            tMallTmp = ""
                            If Not orsCD Is Nothing Then
                                If orsCD.State = adStateOpen Then
                                    If orsCD.RecordCount > 0 Then
                                        orsCD.Filter = vbNullString
                                        orsCD.MoveFirst
                                        orsCD.Filter = "FNDctNo = 11"
                                        orsCD.Sort = "FNScdSeqNo ASC,FTScdCardID ASC"
                                        Do While Not orsCD.EOF
                                            tMallID = SP_FEDtChkString(orsCD, "FTScdCardID")
                                            If tMallTmp <> tMallID Then
                                                If tMallID <> vbNullString Then
'                                                   tTmp = "Customer ID: " & tMallID '*KT 52-02-16 V 1.5.10
                                                    If tTkMallID = "" Then tTkMallID = "Customer ID :" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                                                    If tW_MemberPriceID <> tMallID Then
                                                        '*Bmp 61-04-05 deoken v.5.1804.1
                                                        '7=Layaway Pickup ,16 = Layaway cancel, 20 =Cancel Add Layaway หมายเลข CARD ID: ต้อง detoken ก่อน
'                                                        If ((nW_TokenTransType = 7)) Or (nW_TokenTransType = 16) Or (nW_TokenTransType = 20) Then
'                                                            'tRefDeToken
'                                                            If SP_PRCbDeToken(tMallID, tRefDeToken) Then
'                                                                If Trim(tRefDeToken) <> "" Then tMallID = tRefDeToken
'                                                            End If
'                                                        End If
'                                                        Call SP_CRTxLogMallCardExtend("W_PRNbTicketPeople:tMallID:[" & tMallID & "]") '[5.1804.4][61/04/06]*Bmp 61-04-06
                                                        tTmp = W_CHKtChangeCardID(tMallID, "0") '*Eaw 56-08-17 CommSheet TK-ISS3000-269-TR_Req.xls  ถ้าเป็นบัตร Memder จะส่ง "0" ไปแทน
                                                        tTmp = tTkMallID & " " & tTmp
'                                                        Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                                        tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                                                    End If
                                                End If
                                            End If
                                            tMallTmp = tMallID
                                            orsCD.MoveNext
                                        Loop
                                        orsCD.Filter = vbNullString
                                    End If
                                End If
                            End If
                            
                            'Member : Mall Card[Register]
                            MallCard(0) = SP_FEDtChkString(orsHD, "FTCstRegCardCode")
                            MallCard(1) = SP_FEDtChkString(orsHD, "FTCstRegCardCode2")
                            MallCard(2) = SP_FEDtChkString(orsHD, "FTCstRegCardCode3")
                            For nI = 0 To 2
                                If MallCard(nI) <> vbNullString Then
                                       'tTmp = "Customer ID: " & MallCard(nI) '*KT 52-02-16 V 1.5.10
                                       If tTkMallIDReg = "" Then tTkMallIDReg = "Customer ID :" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                                       tTmp = W_CHKtChangeCardID(MallCard(nI), "0") '*Eaw 56-08-17 CommSheet TK-ISS3000-269-TR_Req.xls  ถ้าเป็นบัตร Member จะส่ง "0" ไปแทน
                                       tTmp = tTkMallIDReg & " " & tTmp
'                                       Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                        tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                                End If
                            Next nI
                            
                            'Member : Employee Card[Discount]
                            tEmpID = ""
                            tEmpTmp = ""
                            If Not orsCD Is Nothing Then
                                If orsCD.State = adStateOpen Then
                                    If orsCD.RecordCount > 0 Then
                                        orsCD.Filter = vbNullString
                                        orsCD.MoveFirst
                                        orsCD.Filter = "FNDctNo = 12"
                                        orsCD.Sort = "FNScdSeqNo ASC,FTScdCardID ASC"
                                        Do While Not orsCD.EOF
                                            tEmpID = SP_FEDtChkString(orsCD, "FTScdCardID")
                                            If tEmpTmp <> tEmpID Then
                                                If tEmpID <> vbNullString Then
                                                   'tTmp = "Customer ID: " & tEmpID '*KT 52-02-16 V 1.5.10
                                                   If tTkEmpID = "" Then tTkEmpID = "Customer ID :" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                                                    tTmp = W_CHKtChangeCardID(tEmpID, "0") '*Eaw 56-08-17 CommSheet TK-ISS3000-269-TR_Req.xls  ถ้าเป็นบัตร Emp จะส่ง "0" ไปแทน
                                                    tTmp = tTkEmpID & " " & tTmp
'                                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                                                    tEmpName = Mid(SP_FEDtChkString(orsCD, "FTScdCardName"), 1, nVB_MaxChar)
                                                    If tEmpName <> "" Then
'                                                        Call oPrn.PrintText(tEmpName, cXL, , vbLeftJustify, True) ' *KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
                                                        tW_SlipFooterStr = tW_SlipFooterStr & "1" & tEmpName & ";"
                                                    End If
                                                   '*KT 51-09-19 ML-SF- 558 Print  Remain Discount for Employee Card
'                                                    If SP_FEDtChkString(orsCD, "FTScdStaTrans") = "1" Then                     '*Ao 2013/10/30   TK EDC
'                                                        'show balance amount วงเงินส่วนลดคงเหลือของบัตรพนักงาน
'                                                        '-----------------------------------------------------------------------------------------
'                                                        tTmp = IIf(tTkRemainDisc <> "", tTkRemainDisc, "Purchase Limit")
'                                                        'Line_Chars
''                                                        Call oPrn.PrintText(Mid(Trim(tTmp), 1, 17), cXL, , vbLeftJustify, False)
'                                                        tW_SlipFooterStr = tW_SlipFooterStr & "8" & Mid(Trim(tTmp), 1, 17) & "{"
''                                                        tTmp = Format(oDisc.C_EMPcBalDis(tEmpID), tW_CurFmt)
'                                                        tTmp = Format(W_EMPcBalDis(tEmpID), tW_CurFmt)
''                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
'                                                        tW_SlipFooterStr = tW_SlipFooterStr & "2" & tTmp & "{"
''                                                        Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
'                                                        tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
'                                                        '-------------------------------------------------------------------------------------------
'                                                    End If
'                                                    If SP_FEDtChkString(orsCD, "FTRemark") <> "" Then               '*Ao 2013/03/06    TK Emp EDC
'                                                        tTmp = SP_FEDtChkString(orsCD, "FTRemark")
'                                                        tTmp = cUT.UT_STRtToken(tTmp, ",", 2)
'                                                        tTmp = tLimit & " " & tTmp
''                                                        Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                                        tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
'                                                    End If
                                                    If SP_FEDtChkString(orsCD, "FTRemark") <> "" Then       '63402800000001309500000001309400000000000000
                                                        tTmp = IIf(tTkRemainDisc <> "", tTkRemainDisc, "Purchase Limit")
                                                        tW_SlipFooterStr = tW_SlipFooterStr & "8" & Mid(Trim(tTmp), 1, 17) & "{"
                                                        
                                                        tTmp = SP_FEDtChkString(orsCD, "FTRemark")              'Tag : 63 และ 40 ตัว  = 2800000001309500000001309400000000000000
                                                        tTmp = Right(tTmp, Len(tTmp) - 2)                                           '28 ตัดทิ้ง    และ ตัดทิ้ง 12 ตัว และเอา 12 ตัว =  000000013094
                                                        tTmp = Mid(tTmp, 13, 12)
                                                        tTmp = Format(tTmp, "#,##0")
                                                        tW_SlipFooterStr = tW_SlipFooterStr & "$" & tTmp & ";"
                                                    End If
                                                End If
                                            End If
                                            tEmpTmp = tEmpID
                                            orsCD.MoveNext
                                        Loop
                                        orsCD.Filter = vbNullString
                                    End If
                                End If
                            End If
'                            Call W_PRNbTicketPoint '*KT 52-02-23 member point
                            '-----------------------------------------------------------------------------------------------------------------------------
'*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
'                            If tW_SalesManID <> vbNullString Then
'                                   tTmp = tTkSalesManID & tW_SalesManID
'                                   Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                            End If
'                            If bVB_CNAlwInputSpn Then
                            If bVB_CNAlwInputSpn Or bVB_CNShortCutSpn Then  '*Eaw 56-06-24 หากกด Hot Key ของ Salesman ต้องพิมพ์ SalesmanID ด้วย
'                                With orsDT
'                                    ReDim aW_SalemanID(0)
'                                    .Filter = "FTSdtSpnCode <> ''"
'                                    .Sort = "FTSdtSpnCode ASC"
'                                    Do While Not orsDT.EOF
'                                        tW_SalesManIDDT = SP_FEDtChkString(orsDT, "FTSdtSpnCode")
'                                        bPass = True
'                                        For nI = 0 To UBound(aW_SalemanID) - 1
'                                            If aW_SalemanID(nI) = tW_SalesManIDDT Then
'                                                bPass = False
'                                                Exit For
'                                            End If
'                                        Next nI
'                                        If bPass Then
'                                            If tW_SalesManIDDT <> vbNullString Then
'                                                   tTmp = tTkSalesManID & tW_SalesManIDDT
'                                                   Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                                    aW_SalemanID(UBound(aW_SalemanID)) = tW_SalesManIDDT
'                                                    ReDim Preserve aW_SalemanID(UBound(aW_SalemanID) + 1)
'                                            End If
'                                        End If
'                                        .MoveNext
'                                    Loop
'                                    .Filter = adFilterNone 'vbNullString
'                                End With
                                If tW_SalesManID <> "" Then
                                    If tTkSalesManID = "" Then tTkSalesManID = "Salesman ID :" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
'                                    tTmp = tTkSalesManID & W_PRNtSalesManID(tW_SalesManID)                     '*Ao 2013/02/28   CommSheet TK-ISS3000-044 00.00.01.xlsx
                                    tTmp = tTkSalesManID & tW_SalesManID            '*Ao 2013/04/18  CommSheet TK-ISS3000-151-IT2 00.00.01.xls
'                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                    tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                                End If
                            End If
                            
                            If tW_CashierName <> vbNullString Then
                                If tTkCashierName = "" Then tTkCashierName = "Cashier Name :" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                                tTmp = tTkCashierName & tW_CashierName
'                                nC_Line = nC_Line - 1    '*BG 56-06-06 TK-ISS3000-078-ST            '*Ao 2013/07/25    ไม่ได้ใช้
'                                bVB_LastLine = False  '*BG 56-06-06 TK-ISS3000-078-ST
'                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
'                                bVB_LastLine = True   '*BG 56-06-06 TK-ISS3000-078-ST
                
                            End If
                            
                            'Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -Follow by format Receipt/EJ
                            
        'Return
        W_PRNbTicketPeople = True
        
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        'Return
        W_PRNbTicketPeople = True
End Function

'Private Function W_PRNbTicketPoint() As Boolean
''-------------------------------------------------
''Call:
''Cmt: Print Point
''------------------------------------------------
'Dim nTmp As Long, tTmp As String
'
''Const tTkLastPoint = "Last Point: "
'Dim tTkLastPoint$
'tTkLastPoint = W_GETxVocabularyTicketFromArray("1966")
''Const tTkToDaypoint = "Today Point: "
'Dim tTkToDaypoint$
'tTkToDaypoint = W_GETxVocabularyTicketFromArray("1967")
'
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'On Error GoTo ErrHandler
'
'                        'Point
'                            Set oPrn.Font = oW_Font(1)
'
'                            '*KT 51-04-08 ตรวจสอบว่า StorePara ให้มีการ Print หรือไม่
'                            'CstCard 1
'                            If tW_CstCardCode <> "" Then
'                                If Mid(tW_OptAlw, 2, 1) = "1" Then
'                                    tTmp = IIf(Len(Mid(tTkLastPoint, 1, 12)) < 12, Mid(tTkLastPoint, 1, 12) & Space$(12 - Len(Mid(tTkLastPoint, 1, 12))), Mid(tTkLastPoint, 1, 12)) & " " & tW_CstCardCode & "  " & nW_LPoint
'                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                End If
'                                 '*KT 51-04-08 ตรวจสอบว่า StorePara ให้มีการ Print หรือไม่
'                                If Mid(tW_OptAlw, 3, 1) = "1" Then
'                                    tTmp = IIf(Len(Mid(tTkToDaypoint, 1, 12)) < 12, Mid(tTkLastPoint, 1, 12) & Space$(12 - Len(Mid(tTkToDaypoint, 1, 12))), Mid(tTkToDaypoint, 1, 12)) & " " & tW_CstCardCode & "  " & nW_ToDayPoint
'                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                End If
'                            End If
'
'                            'CstCard 2
'                            If tW_CstCardCode2 <> "" Then
'                                If Mid(tW_OptAlw, 2, 1) = "1" Then
'                                    'tTmp = tTkLastPoint & "" & tW_CstCardCode2 & "     " & nW_LPoint2
'                                    tTmp = IIf(Len(Mid(tTkLastPoint, 1, 12)) < 12, Mid(tTkLastPoint, 1, 12) & Space$(12 - Len(Mid(tTkLastPoint, 1, 12))), Mid(tTkLastPoint, 1, 12)) & " " & tW_CstCardCode2 & "  " & nW_LPoint2
'                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                End If
'                                 '*KT 51-04-08 ตรวจสอบว่า StorePara ให้มีการ Print หรือไม่
'                                If Mid(tW_OptAlw, 3, 1) = "1" Then
'                                   ' tTmp = tTkToDaypoint & "/" & tW_CstCardCode2 & "     " & nW_ToDayPoint2
'                                   tTmp = IIf(Len(Mid(tTkToDaypoint, 1, 12)) < 12, Mid(tTkLastPoint, 1, 12) & Space$(12 - Len(Mid(tTkToDaypoint, 1, 12))), Mid(tTkToDaypoint, 1, 12)) & " " & tW_CstCardCode2 & "  " & nW_ToDayPoint2
'                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                End If
'                            End If
'
'                            'CstCard 3
'                            If tW_CstCardCode3 <> "" Then
'                                If Mid(tW_OptAlw, 2, 1) = "1" Then
'                                   ' tTmp = tTkLastPoint & "" & tW_CstCardCode3 & "     " & nW_LPoint3
'                                   tTmp = IIf(Len(Mid(tTkLastPoint, 1, 12)) < 12, Mid(tTkLastPoint, 1, 12) & Space$(12 - Len(Mid(tTkLastPoint, 1, 12))), Mid(tTkLastPoint, 1, 12)) & " " & tW_CstCardCode3 & "  " & nW_LPoint3
'                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                End If
'                                 '*KT 51-04-08 ตรวจสอบว่า StorePara ให้มีการ Print หรือไม่
'                                If Mid(tW_OptAlw, 3, 1) = "1" Then
'                                    'tTmp = tTkToDaypoint & "" & tW_CstCardCode3 & "     " & nW_ToDayPoint3
'                                      tTmp = IIf(Len(Mid(tTkToDaypoint, 1, 12)) < 12, Mid(tTkLastPoint, 1, 12) & Space$(12 - Len(Mid(tTkToDaypoint, 1, 12))), Mid(tTkToDaypoint, 1, 12)) & " " & tW_CstCardCode3 & "  " & nW_ToDayPoint3
'                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                End If
'                            End If
'
'                            'If tTmp <> "" Then Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True) ' space - - - - - -Follow by format Receipt/EJ
''Return
'W_PRNbTicketPoint = True
'
'Exit Function
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'
'End Function

'Private Function W_PRNbTicketPointEx20120916() As Boolean
''-------------------------------------------------
''Call:
''Cmt: Print Point อ้างอิงเอกสาร ISS3000 Release 2.4(Mall Department Store Ph 3)External Specification)
''------------------------------------------------
'Dim nTmp As Long, tTmp As String
'
''Member ID(line1)
'Dim tTkPointL1$
'tTkPointL1 = W_GETxVocabularyTicketFromArray("1980")
'
''Member Name(line2)
'Dim tTkPointL2$
'tTkPointL2 = W_GETxVocabularyTicketFromArray("1981")
'
''Member description(line3)
'Dim tTkPointL3$
'tTkPointL3 = W_GETxVocabularyTicketFromArray("1982")
'
''Member description(line4)
'Dim tTkPointL4$
'tTkPointL4 = W_GETxVocabularyTicketFromArray("1983")
'
''Member description(line5)
'Dim tTkPointL5$
'tTkPointL5 = W_GETxVocabularyTicketFromArray("1984")
'
''Member description(line6)
'Dim tTkPointL6$
'tTkPointL6 = W_GETxVocabularyTicketFromArray("1985")
'
''Member description(line7)
'Dim tTkPointL7$
'tTkPointL7 = W_GETxVocabularyTicketFromArray("1986")
'
''Member description(line8)
'Dim tTkPointL8$
'tTkPointL8 = W_GETxVocabularyTicketFromArray("1987")
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'On Error GoTo ErrHandler
'
'                Select Case Val(Format(tW_TransTypeSub))
'                    Case tEN_TrnIPV, tEN_TrnIPVReturn, tEN_TrnIPVRedeem, tEN_TrnIPVVIPService, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVOtherReceive, tEN_TrnIPVOtherDisbursement ''IPV or Void Transaction
'                        'IPV Transaction ไม่ต้องแสดงเรื่อง Member Point ที่ ท้ายบิล ตามเอกสาร phase1.8 UAT status on 09-04-2009
'                    Case tEN_TrnTaxExceptSale, tEN_TrnAddDepositCancel, tEN_TrnDepositCancel 'ตามเอกสาร ISS3000 Release 2.4 External Specification  '*KT 54-09-13 V.4.0.1 Req3 Phase4.0(Member Point) Member Point on Other Receive เอา tEN_TrnOtherReceive ออก '*KT 54-10-26 V.4.0.1 tEN_TrnOtherDisbursement เอาออก
'                        'none
'                    Case Else
'                            If tW_PrnPoint = "1" Then
'                                'Point
'                                    Set oPrn.Font = oW_Font(1)
'                                    'Member ID: xxxxxxxxxxxxxxxxxxxx
'                                    'ชื่อ-สกุล (TH/ENG)
'                                    'Acc.Balance as of YESTERDAY
'                                    'x,xxx คะแนน/points
'                                    '
'                                    '
'                                    'Points earn today,Use tomorrow
'                                    'คะแนนวันนี้จะใช้ได้วันถัดไป
'
'                                    If tW_CstCardCode <> "" Then
'                                            tTmp = ""
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                            tTmp = tTkPointL1 & tW_CstCardCode
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                            'Member Name(line2)
'                                            tTmp = tW_NamePoint & " " & tTkPointL2
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                             If Mid(tW_OptAlw, 2, 1) = "1" Then
'                                                'Member description(line3)
'                                                tTmp = tTkPointL3
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True, , , , True, True)
'                                                'Member description(line4)
'                                                tTmp = nW_LPoint & " " & tTkPointL4
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                            End If
'                                            If Mid(tW_OptAlw, 3, 1) = "1" Then
'                                                'Member description(line5)
'                                                tTmp = tTkPointL5
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                                'Member description(line6)
'                                                tTmp = nW_ToDayPoint & " " & tTkPointL6
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                                'Member description(line7)
'                                            End If
'                                            If Mid(tW_OptAlw, 2, 1) = "1" Or Mid(tW_OptAlw, 3, 1) = "1" Then
'                                                tTmp = tTkPointL7
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                                'Member description(line8)
'                                                tTmp = tTkPointL8
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                            End If
'                                    End If
'                            End If
'                    End Select
''Return
'W_PRNbTicketPointEx20120916 = True
'
'Exit Function
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'
'End Function

'Private Function W_PRNbTicketPointEx() As Boolean
''-------------------------------------------------
''Call:
''Cmt: Print Point อ้างอิงเอกสาร ISS3000 Release 2.4(Mall Department Store Ph 3)External Specification)
''------------------------------------------------
'Dim nTmp As Long, tTmp As String
'
''Member ID(line1)
'Dim tTkPointL1$
'tTkPointL1 = W_GETxVocabularyTicketFromArray("1980")
'
''Member Name(line2)
'Dim tTkPointL2$
'tTkPointL2 = W_GETxVocabularyTicketFromArray("1981")
'
''Member description(line3)
'Dim tTkPointL3$
'tTkPointL3 = W_GETxVocabularyTicketFromArray("1982")
'
''Member description(line4)
'Dim tTkPointL4$
'tTkPointL4 = W_GETxVocabularyTicketFromArray("1983")
'
''Member description(line5)
'Dim tTkPointL5$
'tTkPointL5 = W_GETxVocabularyTicketFromArray("1984")
'
''Member description(line6)
'Dim tTkPointL6$
'tTkPointL6 = W_GETxVocabularyTicketFromArray("1985")
'
''Member description(line7)
'Dim tTkPointL7$
'tTkPointL7 = W_GETxVocabularyTicketFromArray("1986")
'
''Member description(line8)
'Dim tTkPointL8$
'tTkPointL8 = W_GETxVocabularyTicketFromArray("1987")
'
'
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'On Error GoTo ErrHandler
'
'                Select Case Val(Format(tW_TransTypeSub))
'                    Case tEN_TrnIPV, tEN_TrnIPVReturn, tEN_TrnIPVRedeem, tEN_TrnIPVVIPService, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVOtherReceive, tEN_TrnIPVOtherDisbursement ''IPV or Void Transaction
'                        'IPV Transaction ไม่ต้องแสดงเรื่อง Member Point ที่ ท้ายบิล ตามเอกสาร phase1.8 UAT status on 09-04-2009
'                    Case tEN_TrnTaxExceptSale, tEN_TrnAddDepositCancel, tEN_TrnDepositCancel 'ตามเอกสาร ISS3000 Release 2.4 External Specification  '*KT 54-09-13 V.4.0.1 Req3 Phase4.0(Member Point) Member Point on Other Receive เอา tEN_TrnOtherReceive ออก '*KT 54-10-26 V.4.0.1 tEN_TrnOtherDisbursement เอาออก
'                        'none
'                    Case tEN_TrnReturn
'                         If tW_PrnPoint = "1" Then
'                                'Point
'                                    Set oPrn.Font = oW_Font(1)
'                                    'BP Award ID: xxxxxxxxxxxxxxxx
'                                    'Name Member
'                                    'Point Before Update                                                             point
'                                    'Bonus Points Credited                                                         point
'                                    'Total Bonus Points                                                                point
'                                    If tW_CstCardCode <> "" Then
'                                            tTmp = ""
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                            tW_CstCardCode = W_GETtCreditNo(tW_CstCardCode)     '*Eaw 56-06-07
'                                            tTmp = tTkPointL1 & tW_CstCardCode
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                            'Member Name(line2)
''                                            tTmp = tW_NamePoint & " " & tTkPointL2                                                '*Eaw 55-12-26  Member Name: XXXXXX
'                                            tTmp = tTkPointL2 & " " & tW_NamePoint
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                             If Mid(tW_OptAlw, 2, 1) = "1" Then 'Print Before Points
'                                                'Member description(line3)
'                                                tTmp = tTkPointL3
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
'                                                'Member description(line4)
''                                                tTmp = nW_LPoint - nW_ToDayPoint                                                     '*Eaw 55-12-26  Last point
'                                                tTmp = nW_LPoint
''                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)                           '*Ao 2012/12/13   ต้องพิมพ์ ทางขวา
'                                                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
'                                            End If
'                                            If Mid(tW_OptAlw, 3, 1) = "1" Then 'Print Bonus Points
'                                                'Member description(line5)
'                                                tTmp = tTkPointL5
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
'                                                'Member description(line6)
'                                                tTmp = (-1) * nW_ToDayPoint
''                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)                          '*Ao 2012/12/13   ต้องพิมพ์ ทางขวา
'                                                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
'                                            End If
'                                            If Mid(tW_OptAlw, 2, 1) = "1" Or Mid(tW_OptAlw, 3, 1) = "1" Then
'                                                'Member description(line7)
'                                                tTmp = tTkPointL7
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
'                                                'Member description(line8)
''                                                tTmp = nW_LPoint                                                                                          '*Eaw 55-12-26 Last Point + Earn Point
'                                                 tTmp = nW_LPoint - nW_ToDayPoint
''                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)                           '*Ao 2012/12/13   ต้องพิมพ์ ทางขวา
'                                                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
'                                            End If
'                                    End If
'                            End If
'                    Case Else
'                            If tW_PrnPoint = "1" Then
'                                'Point
'                                    Set oPrn.Font = oW_Font(1)
'                                    'BP Award ID: xxxxxxxxxxxxxxxx
'                                    'Name Member
'                                    'Point Before Update                                                             point
'                                    'Bonus Points Credited                                                         point
'                                    'Total Bonus Points                                                                point
'                                    If tW_CstCardCode <> "" Then
'                                            tTmp = ""
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                            tW_CstCardCode = W_GETtCreditNo(tW_CstCardCode)     '*Eaw 56-06-07
'                                            tTmp = tTkPointL1 & tW_CstCardCode
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                            'Member Name(line2)
''                                            tTmp = tW_NamePoint & " " & tTkPointL2                                                '*Eaw 55-12-26  Member Name: XXXXXX
'                                            tTmp = tTkPointL2 & " " & tW_NamePoint
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
'                                             If Mid(tW_OptAlw, 2, 1) = "1" Then 'Print Before Points
'                                                'Member description(line3)
'                                                tTmp = tTkPointL3
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
'                                                'Member description(line4)
''                                                tTmp = nW_LPoint - nW_ToDayPoint                                                     '*Eaw 55-12-26  Last point
'                                                tTmp = nW_LPoint
''                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)                          '*Ao 2012/12/13   ต้องพิมพ์ ทางขวา
'                                                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
'                                            End If
'                                            If Mid(tW_OptAlw, 3, 1) = "1" Then 'Print Bonus Points
'                                                'Member description(line5)
'                                                tTmp = tTkPointL5
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
'                                                'Member description(line6)
'                                                tTmp = nW_ToDayPoint
''                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)                          '*Ao 2012/12/13   ต้องพิมพ์ ทางขวา
'                                                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
'                                            End If
'                                            If Mid(tW_OptAlw, 2, 1) = "1" Or Mid(tW_OptAlw, 3, 1) = "1" Then
'                                                'Member description(line7)
'                                                tTmp = tTkPointL7
'                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
'                                                'Member description(line8)
''                                                tTmp = nW_LPoint                                                                                          '*Eaw 55-12-26 Last Point + Earn Point
'                                                 tTmp = nW_LPoint + nW_ToDayPoint
''                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)                           '*Ao 2012/12/13   ต้องพิมพ์ ทางขวา
'                                                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
'                                            End If
'                                    End If
'                            End If
'                    End Select
''Return
'W_PRNbTicketPointEx = True
'
'Exit Function
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'
'End Function

Private Function W_PRNbTicketPointEx() As Boolean
'-------------------------------------------------
'Call:
'Cmt: Print Point อ้างอิงเอกสาร ISS3000 Release 2.4(Mall Department Store Ph 3)External Specification)
'------------------------------------------------
Dim nTmp As Long, tTmp As String

'Member ID(line1)
Dim tTkPointL1$
tTkPointL1 = W_GETxVocabularyTicketFromArray("1980")

'Member Name(line2)
Dim tTkPointL2$
tTkPointL2 = W_GETxVocabularyTicketFromArray("1981")

'Member description(line3)
Dim tTkPointL3$
tTkPointL3 = W_GETxVocabularyTicketFromArray("1982")

'Member description(line4)
Dim tTkPointL4$
tTkPointL4 = W_GETxVocabularyTicketFromArray("1983")

'Member description(line5)
Dim tTkPointL5$
tTkPointL5 = W_GETxVocabularyTicketFromArray("1984")

'Member description(line6)
Dim tTkPointL6$
tTkPointL6 = W_GETxVocabularyTicketFromArray("1985")

'Member description(line7)
Dim tTkPointL7$
tTkPointL7 = W_GETxVocabularyTicketFromArray("1986")

'Member description(line8)
Dim tTkPointL8$
tTkPointL8 = W_GETxVocabularyTicketFromArray("1987")

Dim tSql As String, tMemID As String, tName As String, tPointBefore As String, tBonusPoints As String, tTotalBonus As String
Dim tExpDate As String
Dim nPoint As Long, nBalance As Long
Dim orsTemp As ADODB.Recordset
Dim orsMem As ADODB.Recordset
Dim bChk As Boolean
Dim tSaveID As String
Dim tSaveName As String
Dim tSavePoint As String
Dim tBP1 As String, tBP2 As String, tExpiry As String
Dim tCardNo As String '*Bmp 61-02-15 Token
Dim tTokenCardNo As String
Dim tRefDeToken As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
    
    W_PRNbTicketPointEx = True
    Select Case Val(Format(tW_TransTypeSub))
        Case tEN_TrnIPV, tEN_TrnIPVReturn, tEN_TrnIPVRedeem, tEN_TrnIPVVIPService, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVOtherReceive, tEN_TrnIPVOtherDisbursement
            Exit Function
'        Case tEN_TrnTaxExceptSale, tEN_TrnAddDepositCancel, tEN_TrnDepositCancel, tEN_TrnClose, tEN_TrnPrintClose
        Case tEN_TrnTaxExceptSale, tEN_TrnClose, tEN_TrnPrintClose, tEN_TrnOverShortReport, tEN_TrnOpen, tEN_TrnStartBank, tEN_TrnPickUp
            Exit Function
    End Select
'*Eaw 56-07-08
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        If bVB_CNCheckMem Then
            bChk = True
        Else
            bChk = False
        End If
    Else
        bChk = False
    End If
'    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
'        If SP_CONbMemberServer(True, "", False) Then
'            bChk = True
'        Else
'            bChk = False
'        End If
'    Else
'        bChk = False
'    End If
'----------------------------------------------------------
'    If SP_CONbMemberServer(True, "", False) Then
'        bChk = True
'    Else
'        bChk = False
'    End If
    tSaveID = ""
    tSaveName = ""
    tSavePoint = ""

    tSql = "SELECT SUM(FCSpoPoint) AS FCSpoPoint,MIN(CONVERT(INT,FTSpoID)) AS FTSpoID,FTSpoMemID,FDDateIns"
    tSql = tSql & vbCrLf & "FROM TPSTSalePoint WITH(NOLOCK)"
    tSql = tSql & vbCrLf & "WHERE " & tW_TransWhere
'    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'    tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tVB_Running & "'"
'    tSql = tSql & vbCrLf & "AND FTShdTransType= '" & tVB_Transaction & "'"
'    tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "GROUP BY FTSpoMemID,FDDateIns"
    tSql = tSql & vbCrLf & "ORDER BY FTSpoID"
    'Call SP_SQLvExecute(tSql, orsTemp) '*Tao 60-06-07 ของเดิม
    Call SP_SQLvExecute(tSql, orsTemp, IIf(bW_UseServer, oDB.ocnOnLine, oDB.ocnOffLine), , , , , IIf(bW_UseServer, True, False)) '*Tao 60-06-07 Reprint หลังจาก Closed Teminal
    If bW_UseServer Then
        '*Tao 61-09-18 ถ้าหาที่ POSServer ไม่เจอ ให้ไปหาที่ Local DB ก่อน (เผื่อยัง Transfer ไปไม่ได้)
        If (orsTemp.BOF And orsTemp.EOF) Then Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine)
    Else
        '*Tao 61-09-18 ถ้าหาที่ Local DB ไม่เจอ ให้ไปหาที่ POSServer ก่อน
        If (orsTemp.BOF And orsTemp.EOF) Then Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine)
    End If
    If (orsTemp.BOF And orsTemp.EOF) Then Exit Function
        
    tBP1 = W_GETxVocabularyTicketFromArray("22")                '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
    tBP2 = W_GETxVocabularyTicketFromArray("23")
    tExpiry = W_GETxVocabularyTicketFromArray("24")
     
    orsTemp.MoveFirst
    Set oPrn.Font = oW_Font(1)
    Do While Not orsTemp.EOF
        
        tMemID = SP_FEDtChkString(orsTemp, "FTSpoMemID")
        nPoint = SP_FEDvChkNumeric(orsTemp, "FCSpoPoint")
        tExpDate = SP_DTEtStrSTD(SP_FEDdChkDate(orsTemp, "FDDateIns"))
               
        If (nPoint > 0) Then  '*Tao 60-04-28 ถ้า Point = 0 ไม่ต้องพิมพ์
        
            'If bChk Then '*Tao 61-05-28 ของเดิม
            If bChk And bVB_ChkWSOFFLINE Then '*Tao 61-06-26 ถ้า WebService Down ไม่ต้องหา Point มาพิมพ์
                'BP Award ID: xxxxxxxxxxxxxxxx
                'Name Member
                'Point Before Update                                                             point
                'Bonus Points Credited                                                         point
                'Total Bonus Points                                                                point
                '===================================
                '*Bmp 61-02-15 Token
                tCardNo = tMemID
                '===================================================================================
                Call SP_CRTxLogMallCardExtend("W_PRNbTicketPointEx: tMemID:[" & tMemID & "] TransType:[" & nW_TokenTransType & "]") '*Bmp 61-04-06
                If ((nW_TokenTransType = 7)) Or (nW_TokenTransType = 16) Or (nW_TokenTransType = 20) Then
                    '* อาจจะต้อง Detoken ก่อน
                End If
                '===================================================================================
                tTokenCardNo = ""
                '*Tao 61-10-05 ไม่ต้องค้นหาจาก MemServer
'                If SP_PRCbTokenMemSrv("W_PRNbTicketPointEx", tCardNo, tTokenCardNo) Then '*Bmp 61-04-04 หาบน Member ก่อน
'                    tCardNo = tTokenCardNo
'                    Call SP_CRTxLogMallCardExtend("W_PRNbTicketPointEx:SP_PRCbTokenMemSrv: tMemID:[" & tMemID & "]TokenNo:[" & tTokenCardNo & "] TransType:[" & nW_TokenTransType & "]") '*Bmp 61-04-06
'                Else
                    tTokenCardNo = ""
                    If SP_PRCbToken("W_PRNbTicketPointEx", tCardNo, tTokenCardNo) Then
                        tCardNo = tTokenCardNo
                        Call SP_CRTxLogMallCardExtend("W_PRNbTicketPointEx:SP_PRCbToken: tMemID:[" & tMemID & "]TokenNo:[" & tTokenCardNo & "] TransType:[" & nW_TokenTransType & "]") '*Bmp 61-04-06
                    Else
                        tCardNo = tMemID
                    End If
'                End If
                '===================================
                tSql = "SELECT TOP 1 FTMcdCode, FDMcdExpDate, FTMcdName, FTMcdSurName ,FCBalance  FROM TCNMMallCard WITH(NOLOCK)"
                'tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tMemID & "'"
                tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardNo & "'" '*Bmp 61-02-15 Token
                tSql = tSql & vbCrLf & "AND FDMcdExpDate = '" & tExpDate & "'"
                tSql = tSql & vbCrLf & "ORDER BY FDMcdExpDate DESC"                             'อาจจะมีมากกว่า 1 Record
                If SP_SQLvExecute(tSql, orsMem, ocnMember, , , , True) = 0 Then
                    If Not (orsMem.BOF And orsMem.EOF) Then
                        tTmp = ""
                        tName = SP_FEDtChkString(orsMem, "FTMcdName") & " " & SP_FEDtChkString(orsMem, "FTMcdSurName")
                        nBalance = SP_FEDvChkNumeric(orsMem, "FCBalance")
                        tExpDate = Format(Month(CDate(tExpDate)), "00") & "/" & Right(Format(Year(CDate(tExpDate)), "0000"), 2)
                            
                        Select Case Val(Format(tW_TransTypeSub))
                            Case tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                tPointBefore = nBalance + nPoint
                                tBonusPoints = (-1) * nPoint
                                tTotalBonus = nBalance
                            Case Else
                                tPointBefore = nBalance - nPoint
                                tBonusPoints = nPoint
                                tTotalBonus = nBalance
                        End Select
                            
    '                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                        tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
    '                    tMemID = W_GETtCreditNo(tMemID)                '*Ao 2013/09/23
    '                    tTmp = tTkPointL1 & tMemID
                        '========================================================================
                        '*Bmp 61-04-05 BP Award ID: พิมพ์ หมายเลขที่ Token ที่ถูกต้อง ต้องพิมพ์ หมายเลขที่ไม่ต้อง Token
                        Call SP_CRTxLogMallCardExtend("W_PRNbTicketPointEx: tMemID:[" & tMemID & "] TransType:[" & nW_TokenTransType & "]") '*Bmp 61-04-06
'                        If ((nW_TokenTransType = 7)) Or (nW_TokenTransType = 16) Or (nW_TokenTransType = 20) Then
'                            'tRefDeToken
'                            If SP_PRCbDeToken(tMemID, tRefDeToken) Then
'                                If Trim(tRefDeToken) <> "" Then tMemID = tRefDeToken
'                            End If
'                        End If
                        '========================================================================
                        tTmp = W_CHKtChangeCardID(tMemID, "0") '*Eaw 56-08-17 CommSheet TK-ISS3000-269-TR_Req.xls  ถ้าเป็นบัตร Memder จะส่ง "0" ไปแทน
                        Call SP_CRTxLogMallCardExtend("W_PRNbTicketPointEx : tTmp, tMemID[" & tTmp & "," & tMemID & "]")   '*[5.2209.0][65/09/10]*Bmp 65-09-10 TTK-ISS3000-809
                        
                        If tMemID = SP_PRCtGetMemPointID Then tTmp = tTmp & tVB_MallCardIDRef '*Tao 60-04-21 ถ้าได้ MallCard แบบ Manual ให้ใส่ * ต่อท้าย (TK-ISS3000-632)
                        Call SP_CRTxLogMallCardExtend("W_PRNbTicketPointEx: tTmp :[" & tTmp & "]" & ",tVB_MallCardIDRef:" & tVB_MallCardIDRef)   '*[5.2209.0][65/09
                
                        tSaveID = tSaveID & tTmp & "," & tExpDate & ";"                           '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
                        tTmp = tTkPointL1 & tTmp    'BP Award: 1111 1111 1111 1111
    '                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                        tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                        tTmp = tTkPointL2 & " " & tName
    '                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                        tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                        '-------------------------------------------------------------------------------------
                        tSaveName = tSaveName & tName & ";"                              '*Ao 2013/11/30             ABCD;XYZ
                        tSavePoint = tSavePoint & tPointBefore & "," & tBonusPoints & "," & tTotalBonus & ";"
                        '-------------------------------------------------------------------------------------
                        If Mid(tW_OptAlw, 2, 1) = "1" Then
                            tTmp = tTkPointL3
    '                        Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                            tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                            tTmp = tPointBefore
    '                        Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                            tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                        End If
                        If Mid(tW_OptAlw, 3, 1) = "1" Then
                            tTmp = tTkPointL5
    '                        Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                            tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                            tTmp = tBonusPoints
    '                        Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                            tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                        End If
                        If Mid(tW_OptAlw, 2, 1) = "1" Or Mid(tW_OptAlw, 3, 1) = "1" Then
                            tTmp = tTkPointL7
    '                        Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                            tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                            tTmp = tTotalBonus
    '                        Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                            tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                        End If
                        
    '                    tW_SlipFooterStr = tW_SlipFooterStr & "1" & tExpiry & " " & tExpDate & ";"          '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
                        If Trim(tExpiry) <> "" Then tW_SlipFooterStr = tW_SlipFooterStr & "1" & tExpiry & " " & tExpDate & ";"      '*Bump 58-02-12 CommSheet TK-ISS3000-578.xls        version 4.2.165 12/02/2015
                    Else
                        GoTo PrintNewMember '*Tao 61-04-24 ให้พิมพ์ NewMember
                    End If
                Else
                    GoTo PrintNewMember '*Tao 61-04-24 ให้พิมพ์ NewMember
                End If
            Else
PrintNewMember:
    '            tMemID = SP_FEDtChkString(orsTemp, "FTSpoMemID")
    '            nPoint = SP_FEDvChkNumeric(orsTemp, "FCSpoPoint")
                tExpDate = Format(Month(CDate(tExpDate)), "00") & "/" & Right(Format(Year(CDate(tExpDate)), "0000"), 2)
                tTmp = ""
                'tName = ""
                tName = "New Member" '*Bump 57-09-12 15.57pm. CommSheet TK-ISS3000-552.xlsx
                nBalance = 0
                Select Case Val(Format(tW_TransTypeSub))
                    Case tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                        tPointBefore = nBalance
                        'tBonusPoints = (-1) * nPoint '*Tao 61-07-19 ของเดิม
                        'tBonusPoints = nBalance '*Tao 61-07-19 TK-ISS3000-683 ถ้าติดต่อ Member Server ไม่ได้ ไม่ต้องเอาไปลบ ให้แสดงเป็นเลข 0
                        tBonusPoints = (-1) * nPoint '*Tao 61-08-09 TK-ISS3000-691 ถ้าติดต่อ Member Server ไม่ได้ ให้เอาไปลบ (กลับไปเป็นเหมือนเดิม)
                        'tTotalBonus = nBalance '*Tao 61-08-09 ของเดิม
                        tTotalBonus = (-1) * nPoint   '*Tao 61-08-09 TK-ISS3000-691 ถ้าติดต่อ Member Server ไม่ได้ ให้เอาไปลบ
                    Case Else
                        tPointBefore = nBalance
                        tBonusPoints = nPoint
                        tTotalBonus = nPoint
                End Select
                
    '            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
    '            tMemID = W_GETtCreditNo(tMemID)                                     '*Ao 2013/09/23
    '            tTmp = tTkPointL1 & tMemID
                tTmp = W_CHKtChangeCardID(tMemID, "0") '*Eaw 56-08-17 CommSheet TK-ISS3000-269-TR_Req.xls  ถ้าเป็นบัตร Memder จะส่ง "0" ไปแทน
                Call SP_CRTxLogMallCardExtend("W_PRNbTicketPointEx(PrintNewMember) : tTmp, tMemID[" & tTmp & "," & tMemID & "]")   '*[5.2209.0][65/09/10]*Bmp 65-09-10 TTK-ISS3000-809
                
                If tMemID = SP_PRCtGetMemPointID Then tTmp = tTmp & tVB_MallCardIDRef '*Tao 60-04-21 ถ้าได้ MallCard แบบ Manual ให้ใส่ * ต่อท้าย (TK-ISS3000-632)
                Call SP_CRTxLogMallCardExtend("W_PRNbTicketPointEx(PrintNewMember): tTmp :[" & tTmp & "]" & ",tVB_MallCardIDRef:" & tVB_MallCardIDRef)   '*[5.2209.0][65/09/10]*Bmp 65-09-10 TTK-ISS3000-809
                
                tSaveID = tSaveID & tTmp & "," & tExpDate & ";"                            '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
                tTmp = tTkPointL1 & tTmp        'BP Award: 1111 1111 1111 1111
    '            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                tTmp = tTkPointL2 & " " & tName
    '            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                tW_SlipFooterStr = tW_SlipFooterStr & "1" & tTmp & ";"
                '-------------------------------------------------------------------------------------
                tSaveName = tSaveName & tName & ";"                              '*Ao 2013/11/30             ABCD;XYZ
                tSavePoint = tSavePoint & tPointBefore & "," & tBonusPoints & "," & tTotalBonus & ";"
                '-------------------------------------------------------------------------------------
                If Mid(tW_OptAlw, 2, 1) = "1" Then
                    tTmp = tTkPointL3
    '                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                    tTmp = tPointBefore
    '                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                    tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                End If
                If Mid(tW_OptAlw, 3, 1) = "1" Then
                    tTmp = tTkPointL5
    '                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                    tTmp = tBonusPoints
    '                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                    tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                End If
                If Mid(tW_OptAlw, 2, 1) = "1" Or Mid(tW_OptAlw, 3, 1) = "1" Then
                    tTmp = tTkPointL7
    '                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                    tTmp = tTotalBonus
    '                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                    tW_SlipFooterStr = tW_SlipFooterStr & "7" & tTmp & ";"
                End If
              
    '            tW_SlipFooterStr = tW_SlipFooterStr & "1" & tExpiry & " " & tExpDate & ";"          '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
                If Trim(tExpiry) <> "" Then tW_SlipFooterStr = tW_SlipFooterStr & "1" & tExpiry & " " & tExpDate & ";" '*Bump 58-02-12 CommSheet TK-ISS3000-578.xls         version 4.2.165 12/02/2015
            End If
        End If '*Tao 60-04-28
        orsTemp.MoveNext
    Loop
    
'    tW_SlipFooterStr = tW_SlipFooterStr & "1" & "   " & ";"
'    tW_SlipFooterStr = tW_SlipFooterStr & "*" & tBP1 & ";"              '*Ao 2015/01/05  CommSheet TK-ISS3000-570_Confirm
'    tW_SlipFooterStr = tW_SlipFooterStr & "*" & tBP2 & ";"

    '*Bump 58-02-12 CommSheet TK-ISS3000-578.xls            version 4.2.165 12/02/2015
    If Trim(tBP1) = "" And Trim(tBP2) = "" Then
        '* ไม่ต้องพิมพ์
    Else
            tW_SlipFooterStr = tW_SlipFooterStr & "1" & "   " & ";"
            If Trim(tBP1) <> "" Then tW_SlipFooterStr = tW_SlipFooterStr & "*" & tBP1 & ";"
            If Trim(tBP2) <> "" Then tW_SlipFooterStr = tW_SlipFooterStr & "*" & tBP2 & ";"
    End If
    
'    '*Tao 60-06-07 ของเดิม
'    tSql = "UPDATE TPSTSalHD"
'    tSql = tSql & vbCrLf & "SET FTCstPointName = '" & Left(tSaveID, 45) & "'"
'    If Len(tSaveID) > 45 Then                   '*Ao 2014/07/21   V 4.2.138   TK-ISS3000PH2-536
'        tSql = tSql & vbCrLf & ",FTShdDocRefMin = '" & Right(tSaveID, Len(tSaveID) - 45) & "'"
'    Else
'        tSql = tSql & vbCrLf & ",FTShdDocRefMin = '' "
'    End If
'    tSql = tSql & vbCrLf & ",FTCstPointName2 = '" & tSaveName & "'"
'    tSql = tSql & vbCrLf & ",FTCstPointName3 = '" & tSavePoint & "'"
'    tSql = tSql & vbCrLf & "WHERE " & tW_TransWhere
'    Call SP_SQLvExecute(tSql)
'    tSql = "UPDATE " & tVB_CNTblHD
''    tSql = tSql & vbCrLf & "SET FTCstPointName = '" & tSaveID & "'"
'    tSql = tSql & vbCrLf & "SET FTCstPointName = '" & Left(tSaveID, 45) & "'"
'    If Len(tSaveID) > 45 Then                   '*Ao 2014/07/21   V 4.2.138   TK-ISS3000PH2-536
'        tSql = tSql & vbCrLf & ",FTShdDocRefMin = '" & Right(tSaveID, Len(tSaveID) - 45) & "'"
'    Else
'        tSql = tSql & vbCrLf & ",FTShdDocRefMin = '' "
'    End If
'    tSql = tSql & vbCrLf & ",FTCstPointName2 = '" & tSaveName & "'"
'    tSql = tSql & vbCrLf & ",FTCstPointName3 = '" & tSavePoint & "'"
'    tSql = tSql & vbCrLf & "WHERE " & tW_TransWhere
'    Call SP_SQLvExecute(tSql)
    
    '*Tao 60-06-07 บันทึกลง HD ได้ไม่ครบ ทำให้ RePrint ไม่ได้
    tSql = "SET FTCstPointName = '" & Left(tSaveID, 50) & "'"
    If Len(tSaveID) > 50 Then
        tSql = tSql & vbCrLf & ",FTRemark = '" & Mid(tSaveID, 51, 100) & "'"
    Else
        'tSql = tSql & vbCrLf & ",FTRemark = ''" '*Tao 61-05-16 ปิดไว้
    End If
    tSql = tSql & vbCrLf & ",FTCstPointName2 = '" & Left(tSaveName, 50) & "'"
    If Len(tSaveName) > 50 Then
        tSql = tSql & vbCrLf & ",FTShdDocRefMin = '" & Mid(tSaveName, 51, 20) & "'"
    Else
        tSql = tSql & vbCrLf & ",FTShdDocRefMin = ''"
    End If
    tSql = tSql & vbCrLf & ",FTCstPointName3 = '" & Left(tSavePoint, 50) & "'"
    If Len(tSavePoint) > 50 Then
        tSql = tSql & vbCrLf & ",FTShdDocRefMax = '" & Mid(tSavePoint, 51, 20) & "'"
    Else
        tSql = tSql & vbCrLf & ",FTShdDocRefMax = ''"
    End If
    tSql = tSql & vbCrLf & "WHERE " & tW_TransWhere
    Call SP_SQLvExecute("UPDATE TPSTSalHD " & tSql)
    Call SP_SQLvExecute("UPDATE " & tVB_CNTblHD & " " & tSql)
    '===================================
Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

Public Function W_GETtCreditNo(ByVal ptCstCardNo As String) As String
'---------------------------------------------------
'call:
'cmt: '*Eaw 56-06-07
'-------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim tCardNo As String
On Error GoTo ErrHandler

    tSql = "SELECT * FROM " & tW_TblRC
    tSql = tSql & vbCrLf & "  WHERE " & tW_TransWhere
'    tSql = tSql & vbCrLf & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
'    tSql = tSql & vbCrLf & "  AND FTShdTransNo='" & tVB_Running & "'"
'    tSql = tSql & vbCrLf & "  AND FTShdTransType='" & tVB_Transaction & "'"
'    tSql = tSql & vbCrLf & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "  AND FTTdmCode  IN ('T002','T003')"
    tSql = tSql & vbCrLf & "  AND FTSrcCardNo ='" & ptCstCardNo & "'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine)
    If Not (orsTemp.EOF And orsTemp.BOF) Then
        tCardNo = W_CHKtChangeCardID(ptCstCardNo, "T002")
    Else
        tCardNo = ptCstCardNo
    End If
    W_GETtCreditNo = tCardNo
    Exit Function
ErrHandler:
    W_GETtCreditNo = ptCstCardNo
End Function

Public Function W_PRNbTicketHeader() As Boolean
Dim nTmp As Long, tTmp As String
Dim tTransNameCut1 As String, tTransNameCut2 As String

'Const tTkTaxInvoiceABB = "TAX INVOICE (ABB.)"
Dim tTkTaxInvoiceABB$
tTkTaxInvoiceABB = W_GETxVocabularyTicketFromArray("1971")
'Const tTkTerminalNo = "TERMINAL NO"
Dim tTkTerminalNo$
tTkTerminalNo = W_GETxVocabularyTicketFromArray("1972")
Dim tTkTaxinv As String
tTkTaxinv = W_GETxVocabularyTicketFromArray("11")
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler

        Set oPrn.Font = oW_Font(1)
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

        'Head
        Call W_GETbSplitTransName(tW_TransNameABB, tTransNameCut1, tTransNameCut2)
         'Line_Chars
         'store  name
         '*Bump 57-04-05 CommSheet TK-ISS3000-461-00.00.04.xls
        'tW_PrintType = "0" คือ พิมพ์ ออก Slip Printer  ไม่ต้องพิมพ์ ส่วนหัว 3 บรรทัด
        '-------
        'TSC                                                        TAX INVOICE
        'GST Reg #                                           M2-0087370-0
        'Company Reg #                                    198902636N
         If bW_PRMGenEJ Then
            If Not (tW_PrintType = "0") Then
                Call oPrn.PrintText(Mid(tW_StoreNameABB, 1, 17), cXL, , vbLeftJustify, False)
                If tTkTaxinv = "" Then tTkTaxinv = "TAX INVOICE"
                Call oPrn.PrintText(tTkTaxinv, cXR, , vbRightJustify)
            End If
         Else
            Call oPrn.PrintText(Mid(tW_StoreNameABB, 1, 17), cXL, , vbLeftJustify, False)
            If tTkTaxinv = "" Then tTkTaxinv = "TAX INVOICE"
            Call oPrn.PrintText(tTkTaxinv, cXR, , vbRightJustify)
         End If
         '*Bump 57-04-05 CommSheet TK-ISS3000-461-00.00.04.xls  Comment 4 Line ข้างล่าง
'        Call oPrn.PrintText(Mid(tW_StoreNameABB, 1, 17), cXL, , vbLeftJustify, False)
''        Call oPrn.PrintText("Tel." & tW_StoreTel, cXR, , vbRightJustify)       '*Ao 2013/04/11  CommSheet TK-ISS3000-147-IT2 00.00.01.xls
'        If tTkTaxinv = "" Then tTkTaxinv = "TAX INVOICE" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
'        Call oPrn.PrintText(tTkTaxinv, cXR, , vbRightJustify)
        '----------------------------------------------------------------------------------------------------------------------
        
'        'store name
'        Call oPrn.PrintText(tTkTaxInvoiceABB, cXL, , vbLeftJustify, False)
'        Call oPrn.PrintText(tW_TaxInvoice, cXR, , vbRightJustify)
'
'        'therminal
'        Call oPrn.PrintText(tTkTerminalNo, cXL, , vbLeftJustify, False)
'        Call oPrn.PrintText(tW_TaxInvoiceABB, cXR, , vbRightJustify)
        
        '55-07-20 Neung TK สลับ TERMINAL NO กับ TAX INVOICE (ABB.)   -----------------------
        'therminal
        '---------------------------------------------------------
        '*Bump 57-04-05 CommSheet TK-ISS3000-461-00.00.04.xls
        'tW_PrintType = "0" คือ พิมพ์ ออก Slip Printer  ไม่ต้องพิมพ์ ส่วนหัว 3 บรรทัด
        '-------
        'TSC                                                        TAX INVOICE
        'GST Reg #                                           M2-0087370-0
        'Company Reg #                                    198902636N
        If bW_PRMGenEJ Then
            If Not (tW_PrintType = "0") Then
                'therminal
                If tTkTerminalNo = "" Then tTkTerminalNo = "GST Reg # "
                Call oPrn.PrintText(tTkTerminalNo, cXL, , vbLeftJustify, False)
                Call oPrn.PrintText(tW_TaxInvoiceABB, cXR, , vbRightJustify)
                
                'store name
                If tTkTaxInvoiceABB = "" Then tTkTaxInvoiceABB = "Company Reg # " '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                Call oPrn.PrintText(tTkTaxInvoiceABB, cXL, , vbLeftJustify, False)
                Call oPrn.PrintText(tW_TaxInvoice, cXR, , vbRightJustify)
            End If
        Else
            'therminal
            If tTkTerminalNo = "" Then tTkTerminalNo = "GST Reg # " '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
            Call oPrn.PrintText(tTkTerminalNo, cXL, , vbLeftJustify, False)
            Call oPrn.PrintText(tW_TaxInvoiceABB, cXR, , vbRightJustify)
            
            'store name
            If tTkTaxInvoiceABB = "" Then tTkTaxInvoiceABB = "Company Reg # " '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
            Call oPrn.PrintText(tTkTaxInvoiceABB, cXL, , vbLeftJustify, False)
            Call oPrn.PrintText(tW_TaxInvoice, cXR, , vbRightJustify)
        End If
        '---------------------------
        '*Bump 57-04-05 CommSheet TK-ISS3000-461-00.00.04.xls  Comment 8 Line ข้างล่าง
        '---------------------------
'        If tTkTerminalNo = "" Then tTkTerminalNo = "GST Reg # " '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
'        Call oPrn.PrintText(tTkTerminalNo, cXL, , vbLeftJustify, False)
'        Call oPrn.PrintText(tW_TaxInvoiceABB, cXR, , vbRightJustify)
        
'        'store name
'        If tTkTaxInvoiceABB = "" Then tTkTaxInvoiceABB = "Company Reg # " '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
'        Call oPrn.PrintText(tTkTaxInvoiceABB, cXL, , vbLeftJustify, False)
'        Call oPrn.PrintText(tW_TaxInvoice, cXR, , vbRightJustify)
        '----------------------------------------------------------------------------------------------------------------------
          '*Bump 57-04-05 CommSheet TK-ISS3000-461-00.00.04.xls
        'tW_PrintType = "0" คือ พิมพ์ ออก Slip Printer  ไม่ต้องพิมพ์ ส่วนหัว 3 บรรทัด และ 1 บรรทัดว่าง
        '-------
        'TSC                                                        TAX INVOICE
        'GST Reg #                                           M2-0087370-0
        'Company Reg #                                    198902636N
        If bW_PRMGenEJ Then
            If Not (tW_PrintType = "0") Then
                Call oPrn.PrintText("", cXC, , vbCenter, , , , , True)                          '*Ao 2013/08/26  CommSheet TK-ISS3000-322-UAT TK-ISS3000-FAPL-101.xlsx
            End If
        Else
            Call oPrn.PrintText("", cXC, , vbCenter, , , , , True)                          '*Ao 2013/08/26  CommSheet TK-ISS3000-322-UAT TK-ISS3000-FAPL-101.xlsx
        End If
'        Call oPrn.PrintText("", cXC, , vbCenter, , , , , True)                          '*Ao 2013/08/26  CommSheet TK-ISS3000-322-UAT TK-ISS3000-FAPL-101.xlsx
        'trans name 1
        Set oPrn.Font = oW_Font(2)
        Call oPrn.PrintText(tTransNameCut1, cXL, , vbLeftJustify, False, , , , True)
                    
        'trans no
        Set oPrn.Font = oW_Font(1)
        tTmp = tW_EmpCdoe & " " & IIf(Len(tW_StoreNum) < 4, Space$(4 - Len(tW_StoreNum)), "") & tW_StoreNum & " " & IIf(Len(tW_StoreNum) < 5, Space$(5 - Len(tW_TmnNum)), "") & tW_TmnNum & " " & IIf(Len(tW_TransNo) < 5, Space$(5 - Len(tW_TransNo)), "") & tW_TransNo
        Call oPrn.PrintText(tTmp, cXR, , vbRightJustify)
        
        'trans name 2
        If tTransNameCut2 <> vbNullString Then
            Set oPrn.Font = oW_Font(2)
            Call oPrn.PrintText(tTransNameCut2, cXL, , vbLeftJustify, False, , , , True)
        End If
        
        'trans date
        Set oPrn.Font = oW_Font(1)
        'Line_Chars
        tTmp = Format(dW_SysDate, tW_DateFmt) & " " & Mid(tW_SysTime, 1, 5)
        tTmp = Replace(tTmp, "-", "/")
        tTmp = Replace(tTmp, ".", "/")
        Call oPrn.PrintText(tTmp, cXR, , vbRightJustify)

        'Return
        W_PRNbTicketHeader = True
    
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
      'Return
        W_PRNbTicketHeader = True
End Function

Public Function W_PRNbTicketHeaderSlp() As Boolean
Dim nTmp As Long, tTmp As String
Dim tTransNameCut1 As String, tTransNameCut2 As String
Dim nLoop As Long
Dim tTkMdDulicate$
Dim tSpace As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
    
'        oPrn.ClearBufferOPOSPrn '*Eaw 56-07-06 Work shop
    
'        For nLoop = 0 To 3
'             Call oPrn.PrintText("", cXL, , vbLeftJustify, True)
'        Next nLoop
        tSpace = ""         '*Ao 2013/04/30  CommSheet TK-ISS3000-095 00.00.04.xls
        Call oPrn.PrintText("", cXC, , vbCenter, , , , , True)       '*Ao 2013/08/26  CommSheet TK-ISS3000-322-UAT TK-ISS3000-FAPL-101.xlsx
        Set oPrn.Font = oW_Font(1)
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        'Head
        Call W_GETbSplitTransName(tW_TransNameABB, tTransNameCut1, tTransNameCut2)
        
        'trans name 1
        Set oPrn.Font = oW_Font(2)
        Select Case SP_VALnText2Double(tW_TransTypeSub)
                Case nEN_TrnReturn, nEN_TrnReturnTaxInc, nEN_TrnReturnManualTax, nEN_TrnDepositCancel, nEN_TrnAddDepositCancel         '*Ao 2013/08/15   CommSheet TK-ISS3000-315-UAT-FAPL-092.xlsx
                    Call oPrn.PrintText(tTransNameCut1, cXL, , vbLeftJustify, False, , , , True)
                    Call oPrn.LineHeader(0, tTransNameCut1, cXL, vbLeftJustify, False)
                Case Else
                    'If bW_StaPrnCV2 Then
                    If bW_StaPrnCV2 And bW_StaPrnCVReIssue Then '*Thirapong 2021-03-27
                        Call oPrn.PrintText("Re-issue", cXL, , vbLeftJustify, False, , , , True)
                        Call oPrn.LineHeader(0, "Re-issue", cXL, vbLeftJustify, False)
                    Else
                        Call oPrn.PrintText(tTransNameCut1, cXL, , vbLeftJustify, False, , , , True)
                        Call oPrn.LineHeader(0, tTransNameCut1, cXL, vbLeftJustify, False)
                    End If
        End Select
                    
        'trans no
        Set oPrn.Font = oW_Font(1)
        tTmp = tW_EmpCdoe & " " & IIf(Len(tW_StoreNum) < 4, Space$(4 - Len(tW_StoreNum)), "") & tW_StoreNum & " " & IIf(Len(tW_StoreNum) < 5, Space$(5 - Len(tW_TmnNum)), "") & tW_TmnNum & " " & IIf(Len(tW_TransNo) < 5, Space$(5 - Len(tW_TransNo)), "") & tW_TransNo
        tTmp = tTmp & tSpace                                                            '*Ao 2013/04/22  CommSheet TK-ISS3000-095 00.00.02.xls
        Call oPrn.PrintText(tTmp, cXR, , vbRightJustify)
        Call oPrn.LineHeader(1, tTmp, cXR, vbRightJustify, True)
        
        'trans name 2
        'TK-ISS3000-078 Counting number of lines to be printed at Slip printer is incorrected
        Select Case SP_VALnText2Double(tW_TransTypeSub)
                Case nEN_TrnReturn, nEN_TrnReturnTaxInc, nEN_TrnReturnManualTax, nEN_TrnDepositCancel, nEN_TrnAddDepositCancel         '*Ao 2013/08/15   CommSheet TK-ISS3000-315-UAT-FAPL-092.xlsx
                    If tTransNameCut2 <> vbNullString Then
                        Set oPrn.Font = oW_Font(2)
                        Call oPrn.PrintText(tTransNameCut2, cXL, , vbLeftJustify, False, , , , True)
                        Call oPrn.LineHeader(2, tTransNameCut2, cXL, vbLeftJustify, False)
                    Else
                        Call oPrn.LineHeader(2, "", cXL, vbLeftJustify, False)
                    End If
                Case Else
                    If Not bW_StaPrnCV2 Then
                        If tTransNameCut2 <> vbNullString Then
                            Set oPrn.Font = oW_Font(2)
                            Call oPrn.PrintText(tTransNameCut2, cXL, , vbLeftJustify, False, , , , True)
                            Call oPrn.LineHeader(2, tTransNameCut2, cXL, vbLeftJustify, False)
                        Else
                            Call oPrn.LineHeader(2, "", cXL, vbLeftJustify, False)
                        End If
                    Else
                        Call oPrn.LineHeader(2, "", cXL, vbLeftJustify, False)
                    End If
        End Select
        
        'trans date
        Set oPrn.Font = oW_Font(1)
        'Line_Chars
        tTmp = Format(dW_SysDate, tW_DateFmt) & " " & Mid(tW_SysTime, 1, 5)
        tTmp = Replace(tTmp, "-", "/")
        tTmp = Replace(tTmp, ".", "/")
        tTmp = tTmp & tSpace                                                        '*Ao 2013/04/22  CommSheet TK-ISS3000-095 00.00.02.xls
        Call oPrn.PrintText(tTmp, cXR, , vbRightJustify)
        Call oPrn.LineHeader(3, tTmp, cXR, vbRightJustify, True)
        'TK-ISS3000-064 CV receipt value and Reprint value isn't the same
        If bW_EnableRePrintCV1 Then
            Set oPrn.Font = oW_Font(2)
            tTkMdDulicate = Mid(W_GETxVocabularyTicketFromArray("1975"), 1, 17)
            Call oPrn.PrintText(tTkMdDulicate, cXC, , vbCenter, , 1, , , True)
            Call oPrn.PrintText("", cXC, , vbCenter, , 1, , , True)
            Set oPrn.Font = oW_Font(1)
        End If
        'Return
        W_PRNbTicketHeaderSlp = True
    
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
      'Return
        W_PRNbTicketHeaderSlp = True
End Function

Public Function W_PRNbTicketHeaderSlp2() As Boolean
Dim nTmp As Long, tTmp As String
'Dim tTransNameCut1 As String, tTransNameCut2 As String
Dim nLoop As Long
Dim tTkMdDulicate$
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
        'TK-ISS3000-102 00.00.02-Pae
        'If bC_PrnByOPOS And tW_TypePrinter = 0 Then
        'TK-ISS3000-150-IT2 Credit Voucher Receipt must be print on Sales Endorsement area
        'If bC_PrnByOPOS Then
            For nLoop = 0 To 20 '22 line
                 Call oPrn.PrintText("", cXL, , vbLeftJustify, True)
                 oPrn.LineNo = 0
            Next nLoop
        'End If
        Set oPrn.Font = oW_Font(1)
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        'Head
       ' Call W_GETbSplitTransName(tW_TransNameABB, tTransNameCut1, tTransNameCut2)
        
        'trans name 1
        Set oPrn.Font = oW_Font(2)
        Call oPrn.PrintText("Receipt", cXL, , vbLeftJustify, False, , , , True)
                    
        'trans no
        Set oPrn.Font = oW_Font(1)
        tTmp = tW_EmpCdoe & " " & IIf(Len(tW_StoreNum) < 4, Space$(4 - Len(tW_StoreNum)), "") & tW_StoreNum & " " & IIf(Len(tW_StoreNum) < 5, Space$(5 - Len(tW_TmnNum)), "") & tW_TmnNum & " " & IIf(Len(tW_TransNo) < 5, Space$(5 - Len(tW_TransNo)), "") & tW_TransNo
        Call oPrn.PrintText(tTmp, cXR, , vbRightJustify)
        
        'trans name 2
        'If tTransNameCut2 <> vbNullString Then
        '    Set oPrn.Font = oW_Font(2)
       '     Call oPrn.PrintText(tTransNameCut2, cXL, , vbLeftJustify, False, , , , True)
        'End If
        
        'trans date
        Set oPrn.Font = oW_Font(1)
        'Line_Chars
        tTmp = Format(dW_SysDate, tW_DateFmt) & " " & Mid(tW_SysTime, 1, 5)
        tTmp = Replace(tTmp, "-", "/")
        tTmp = Replace(tTmp, ".", "/")
        Call oPrn.PrintText(tTmp, cXR, , vbRightJustify)
        'TK-ISS3000-064 CV receipt value and Reprint value isn't the same
        If bW_EnableRePrintCV2 Then
            Set oPrn.Font = oW_Font(2)
            tTkMdDulicate = Mid(W_GETxVocabularyTicketFromArray("1975"), 1, 17)
            Call oPrn.PrintText(tTkMdDulicate, cXC, , vbCenter, , 1, , , True)
            Call oPrn.PrintText("", cXC, , vbCenter, , 1, , , True)
            Set oPrn.Font = oW_Font(1)
        End If
        'Return
        W_PRNbTicketHeaderSlp2 = True
    
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
      'Return
        W_PRNbTicketHeaderSlp2 = True
End Function


Private Function W_PRNbTicketOverShort() As Boolean
    Dim tTkLnOverShort$
    tTkLnOverShort = W_GETxVocabularyTicketFromArray("1999")
    Dim tTkLnOverShortTransfer$
    tTkLnOverShortTransfer = W_GETxVocabularyTicketFromArray("1")
    Dim tTkLnIssuedBy$
    tTkLnIssuedBy = W_GETxVocabularyTicketFromArray("2")
    Dim tTkLnApprovedBy$
    tTkLnApprovedBy = W_GETxVocabularyTicketFromArray("3")
    Dim tTkLnCashierCopy$
    tTkLnCashierCopy = W_GETxVocabularyTicketFromArray("4")
    Dim tTkLnCashierCopy2$
    tTkLnCashierCopy2 = W_GETxVocabularyTicketFromArray("10") 'TK-CommSheet-007-Print OverShort
    Dim tTkOverShort$
    tTkOverShort = W_GETxVocabularyTicketFromArray("8")
    
'    Dim tPdtName As String
'    Dim tPdtName2 As String
'    Dim tTempName1 As String
'    Dim tTempName2 As String
    Dim cEnter As Double
    Dim cACTUAL As Double
    Dim cQtyEnter As Double
    Dim cQtyACTUAL As Double
    Dim cNet As Double
    
    'Dim orsPickup As ADODB.Recordset
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    Dim tTransWhere As String
    Dim nError As Long
    Dim tCode As String
    Dim tName As String
    Dim cTotEnt As Double
    Dim cTotAct As Double
    Dim nSpace As Integer, nSpa As Integer, nSpaceEx As Integer, nSpaEx As Integer
    Dim tTmp  As String
    Dim cReg As Double
    Dim cACTUALPrice As Double, cCashTrans As Double
    Dim orsFor As ADODB.Recordset
    Dim cForeignEnter As Double, cTransfer As Double, cForeignEx As Double, cStkFac As Double
    
On Error GoTo ErrHandler
    
    Set orsTemp = orsDT.Clone

    'tSql = wFunctionKB.W_GETtCloseTenderSQL(tW_TblHD, tW_TblRC, "FTShdTransType", True)
    'Call SP_SQLvExecute(tSql, orsTemp)
    tSql = "SELECT LTRIM(RTRIM(FTFcrCode)) AS FTFcrCode,FTFcrSign,FTFcrType,FCFcrExchang FROM TCNMForCurrency" '*Ao 2013/05/03 CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
    Call SP_SQLvExecute(tSql, orsFor, oDB.ocnOnLine, , , , , True, True)
    
   ' ***** OVER / SHORT REPORT ****
   Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
   Call oPrn.PrintText(tTkLnOverShort, cXC, , vbCenter, , , , , True)
   Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
   
'   Call oPrn.PrintText("ACTUAL", cXC - cX2, , vbRightJustify, False) 'AMT
'   Call oPrn.PrintText("ENTER", cXHR, , vbRightJustify, True)  'AMT  cXHR + cX4
'   Call oPrn.PrintText("TNDR", cXL, , vbLeftJustify, False) 'TNDR
'   Call oPrn.PrintText("AMT", cXC - cX2, , vbRightJustify, False) 'AMT
'   Call oPrn.PrintText("CNT", cXC + cX4, , vbRightJustify, False) 'CNT
'   Call oPrn.PrintText("AMT", cXHR, , vbRightJustify, False)  'AMT
'   Call oPrn.PrintText("CNT", cXMK, , vbRightJustify, True) 'CNT
    '*Ao 2013/01/29
    If nVB_MaxChar > 42 Then
        nSpace = 9
        nSpa = 9
    Else
        nSpace = 8
        nSpa = 5
    End If
    tTmp = Space(nSpa) & Space(nSpace) & "ACTUAL" & Space(5)
    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
    tTmp = "ENTER" & Space(5)
    Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
    
    tTmp = "TNDR" & Space(nSpa - 4) & Space(nSpace) & Space(3) & "AMT" & Space(2) & "CNT"
    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
    tTmp = "AMT" & Space(2) & "CNT"
    Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
    '----------------------------------------------------------------------------------------
   'cash
    With orsDT
        .Filter = "FTSdtStaPdt='4'"
        .Sort = "FNSdtSeqNo ASC"
        If Not (orsDT.BOF And orsDT.EOF) Then
            Do While Not .EOF
                cEnter = 0
                cQtyEnter = 0
                tCode = SP_FEDtChkString(orsDT, "FTSkuCode")
                tName = SP_FEDtChkString(orsDT, "FTSkuAbbName")
                cACTUAL = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
                cACTUALPrice = SP_FEDvChkNumeric(orsDT, "FCSdtSalePrice")                    '*Ao แก้ไขการพิมพ์  OVER/SHORT REPORT
                cQtyACTUAL = SP_FEDvChkNumeric(orsDT, "FCSdtQtyAll")
                
'                If cACTUAL = 0 Then cQtyACTUAL = 0 '*BG 56-09-12 Comsheet - 353
                Select Case tCode
                    Case "Cash"
                            orsTemp.Filter = "FTSdtStaPdt='0' OR FTSdtStaPdt='1'"
                            orsTemp.Sort = "FNSdtSeqNo ASC"
                            If Not (orsTemp.BOF And orsTemp.EOF) Then
                                Do While Not orsTemp.EOF
                                        cEnter = cEnter + SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt")
'                                        cQtyEnter = cQtyEnter + SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll")
                                        '*BG 56-09-03  Comsheet-285
                                        If SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt") <> 0 Then
                                            cQtyEnter = cQtyEnter + 1
                                        End If
                                        '****
                                        orsTemp.MoveNext
                                Loop
'                                If cEnter > 0 Then
'                                    cQtyEnter = 1 'TK-ComSheet-002
'                                Else
'                                    cQtyEnter = 0
'                                End If
                            End If
                            orsTemp.Filter = adFilterNone
                    Case "GVch"
                            orsTemp.Filter = "FTSdtStaPdt='2'"
                            orsTemp.Sort = "FNSdtSeqNo ASC"
                            If Not (orsTemp.BOF And orsTemp.EOF) Then
                                Do While Not orsTemp.EOF
                                        cEnter = cEnter + SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt")
'                                        cQtyEnter = cQtyEnter + SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll")
                                        '*BG 56-09-03   Comsheet-285
                                        If SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt") <> 0 Then
                                            cQtyEnter = cQtyEnter + 1
                                        End If
                                        '****
                                        orsTemp.MoveNext
                                Loop
                            End If
                            orsTemp.Filter = adFilterNone
                    Case "T006" '*Tao 62-08-06 Vendor Voucher
                            orsTemp.Filter = "FTSkuCode='T006' AND FTSdtStaPdt='6'" '6=ที่ป้อนข้อมูลไว้
                            orsTemp.Sort = "FNSdtSeqNo ASC"
                            If Not (orsTemp.BOF And orsTemp.EOF) Then
                                Do While Not orsTemp.EOF
                                    cEnter = cEnter + SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt")
                                    cQtyEnter = cQtyEnter + SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll")
                                    orsTemp.MoveNext
                                Loop
                            End If
                            orsTemp.Filter = adFilterNone
                    Case Else
                            '--------------------------------------------------------------------
                            orsFor.Filter = "FTFcrCode = '" & tCode & "'"
                            If Not (orsFor.BOF And orsFor.EOF) Then
                                cStkFac = SP_FEDvChkNumeric(orsFor, "FCFcrExchang")
                            Else
                                cStkFac = 1
                            End If
                            orsFor.Filter = adFilterNone
                            '--------------------------------------------------------------------
                            orsTemp.Filter = "FTSdtStaPdt='3' AND FTSkuCode ='" & tCode & "'"
                            orsTemp.Sort = "FNSdtSeqNo ASC"
                            If Not (orsTemp.BOF And orsTemp.EOF) Then
                                Do While Not orsTemp.EOF
                                        'cEnter = cEnter + SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt") 'TK-ISS3000-164 Over/Short Report - Cash & Foreign Currency problem
                                    cEnter = cEnter + SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")                       '*Ao 2013/10/11    V4.0.78
                                    cStkFac = SP_FEDvChkNumeric(orsTemp, "FCSdtStkFac")                         '*Ao 2014/06/27
                                    cForeignEnter = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice") * cStkFac, True, False)  '*Ao 2013/10/11    V4.0.78
                                    If SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice") <> 0 Then
                                        cQtyEnter = cQtyEnter + 1
                                    End If
                                    orsTemp.MoveNext
                                Loop
                            End If
                            orsTemp.Filter = adFilterNone
                            '---------------------------------------------------------------------------------------------------------------'*Ao 2013/10/11    V4.0.78
                            orsTemp.Filter = "FTSdtStaPdt='5' AND FTSkuCode ='" & tCode & "'"
                            If Not (orsTemp.BOF And orsTemp.EOF) Then
                                cTransfer = SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")
                            Else
                                cTransfer = 0
                            End If
                            cForeignEx = SP_FMTcCurPoint((cACTUALPrice - cTransfer) * cStkFac, True, False)
                            orsTemp.Filter = adFilterNone
                            '---------------------------------------------------------------------------------------------------------------
                End Select
'                Call oPrn.PrintText(tName, cXL, , vbLeftJustify, False) '
'                Call oPrn.PrintText(Format(cACTUAL, tW_CurFmt), cXC - cX2, , vbRightJustify, False) 'ACTUAL
'                Call oPrn.PrintText(Format(cQtyACTUAL, "###0"), cXC + cX4, , vbRightJustify, False) 'CNT
'                Call oPrn.PrintText(Format(cEnter, tW_CurFmt), cXHR, , vbRightJustify, False)  'ENTER
'                Call oPrn.PrintText(Format(cQtyEnter, "###0"), cXMK, , vbRightJustify, True) 'CNT
                '*Ao 2013/01/29
                If nVB_MaxChar > 42 Then
                    nSpace = 15
                    nSpaceEx = 14
                Else
                    nSpace = 14
                    nSpaceEx = 13
                End If
                tName = tName & Space(nSpa - IIf(Len(tName) > nSpa, nSpa, Len(tName)))
                If Len(tName) > nSpa Then
                    nSpaEx = Abs((nSpa - Len(tName)))
                Else
                    nSpaEx = 0
                End If
'                If (nSpace - nSpaEx) < Len(Format(cACTUAL, tW_CurFmt)) Then
'                    nSpaEx = 0
'                Else
'                    nSpaEx = (nSpace - nSpaEx) - Len(Format(cACTUAL, tW_CurFmt))
'                End If
''                tTmp = tName & Space((nSpace - nSpaEx) - Len(Format(cACTUAL, tW_CurFmt))) & Format(cACTUAL, tW_CurFmt) & Space(5 - Len(Format(cQtyACTUAL, "####0"))) & Format(cQtyACTUAL, "####0")
'                tTmp = tName & Space(nSpaEx) & Format(cACTUAL, tW_CurFmt)
                If (nSpace - nSpaEx) < Len(Format(cACTUALPrice, tW_CurFmt)) Then         '*Ao แก้ไขการพิมพ์  OVER/SHORT REPORT
                    nSpaEx = 0
                Else
                    nSpaEx = (nSpace - nSpaEx) - Len(Format(cACTUALPrice, tW_CurFmt))
                End If
'                tTmp = tName & Space((nSpace - nSpaEx) - Len(Format(cACTUAL, tW_CurFmt))) & Format(cACTUAL, tW_CurFmt) & Space(5 - Len(Format(cQtyACTUAL, "####0"))) & Format(cQtyACTUAL, "####0")
                tTmp = tName & Space(nSpaEx) & Format(cACTUALPrice, tW_CurFmt)
                '---------------------------------------------------------------------------------------------------------------
                If 5 < Len(Format(cQtyACTUAL, "####0")) Then
                    nSpaEx = 0
                Else
                    nSpaEx = 5 - Len(Format(cQtyACTUAL, "####0"))
                End If
                tTmp = tTmp & Space(nSpaEx) & Format(cQtyACTUAL, "####0")
                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                '----------------------------------------------------------------------------------------------------------------------------------
                If nSpaceEx < Len(Format(cEnter, tW_CurFmt)) Then
                    nSpaEx = 0
                Else
                    nSpaEx = nSpaceEx - Len(Format(cEnter, tW_CurFmt))
                End If
'                tTmp = Space(nSpaceEx - Len(Format(cEnter, tW_CurFmt))) & Format(cEnter, tW_CurFmt) & Space(5 - Len(Format(cQtyEnter, "####0"))) & Format(cQtyEnter, "####0")
                tTmp = Space(nSpaEx) & Format(cEnter, tW_CurFmt)
                If 5 < Len(Format(cQtyEnter, "####0")) Then
                    nSpaEx = 0
                Else
                    nSpaEx = 5 - Len(Format(cQtyEnter, "####0"))
                End If
                '----------------------------
                If Len(Format(cEnter, tW_CurFmt)) >= 14 Then
                    If Len(Format(cEnter, tW_CurFmt)) - 12 > 5 Then
                        nSpaEx = 0
                    Else
                        nSpaEx = 5 - (Len(Format(cEnter, tW_CurFmt)) - 12)
                    End If
                End If
                '----------------------------
                tTmp = tTmp & Space(nSpaEx) & Format(cQtyEnter, "####0")
                Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                '--------------------------------------------------------------------------------------------------
'                cTotEnt = cTotEnt + cEnter
'                cTotAct = cTotAct + cACTUAL
                Select Case tCode
                    Case "Cash"
                        cTotEnt = cTotEnt + cEnter
                        cTotAct = cTotAct + cACTUAL
                    Case "GVch", "T006" '*Tao 62-08-19 ให้ Vendor Voucher ไปรวมด้วย
                        cTotEnt = cTotEnt + cEnter
                        cTotAct = cTotAct + cACTUAL
                    Case Else
                        cTotEnt = cTotEnt + cForeignEnter
                        cTotAct = cTotAct + cForeignEx
                End Select
'                cTotAct = cTotAct + cACTUAL
                .MoveNext
            Loop
        End If
        .Filter = adFilterNone
   End With
   
 '***CASH TRANSFER REPORT***
 '57-05-08   20130504-ReEntry.xls Re-Entry ไม่ต้องแสดง ASH TRANSFER REPORT
cCashTrans = 0
Select Case Val(Format(tW_TransTypeSub))
    Case 37
        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
        Call oPrn.PrintText(tTkLnOverShortTransfer, cXC, , vbCenter, , , , , True)
        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
         With orsDT
                .Filter = "FTSdtStaPdt='5'"
                .Sort = "FNSdtSeqNo ASC"
                If Not (orsDT.BOF And orsDT.EOF) Then
                    Do While Not .EOF
                        'TK-ISS3000-164 Over/Short Report - Cash & Foreign Currency problem
                        cNet = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
                        cReg = SP_FEDvChkNumeric(orsDT, "FCSdtRegPrice")
                        tCode = SP_FEDtChkString(orsDT, "FTSkuCode")
                        tName = SP_FEDtChkString(orsDT, "FTSkuAbbName")
                        Call oPrn.PrintText(tName, cXL, , vbLeftJustify, False) '
'                        If tCode = "T012" Then
'                           Call oPrn.PrintText(Format(cReg, tW_CurFmt), cXMK, , vbRightJustify, True) 'amt
'                        Else
'                           Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXMK, , vbRightJustify, True) 'amt
'                        End If
                        Call oPrn.PrintText(Format(cReg, tW_CurFmt), cXMK, , vbRightJustify, True) 'amt   *BG 56-09-17 Comsheet-355
                        '----------------------------------------------------------------------------------
'                        cCashTrans = cCashTrans + cNet
                        Select Case tCode
                            Case "T001"
                                cCashTrans = cCashTrans + cNet
                            Case "T005"
                                cCashTrans = cCashTrans + cNet
                            Case Else
                                
                        End Select
                        .MoveNext
                    Loop
                End If
           End With
   End Select
   Call oPrn.PrintText(tTkOverShort, cXC, , vbRightJustify, False)
'    Call oPrn.PrintText(Format(cW_Total - cTotAct , tW_CurFmt), cXMK, , vbRightJustify, True)      '*Ao 2013/09/23  CommSheet TK-ISS3000-305-UAT_00.00.03 FSBT NG
'   Call oPrn.PrintText(Format(cW_Total - (cTotAct - cCashTrans), tW_CurFmt), cXMK, , vbRightJustify, True)
   Call oPrn.PrintText(Format(cTotEnt - (cTotAct - cCashTrans), tW_CurFmt), cXMK, , vbRightJustify, True)

   Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
   Call oPrn.PrintText(tTkLnIssuedBy, cXL, , vbLeftJustify, , , , , True)
   Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
   
   Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
   Call oPrn.PrintText(tTkLnApprovedBy, cXL, , vbLeftJustify, , , , , True)
   Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
   'TK-CommSheet-007-Print OverShort
   Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
   If nW_PrintCnt = 1 Then
        Call oPrn.PrintText(tTkLnCashierCopy2, cXL, , vbLeftJustify, , , , , True)
   Else
        Call oPrn.PrintText(tTkLnCashierCopy, cXL, , vbLeftJustify, , , , , True)
   End If
   Exit Function
ErrHandler:

End Function

'ชื่อ 20 digit
Private Function W_PRNbTicketDetail() As Boolean
Dim nSeq As Long
Dim nTmp As Long, tTmp As String, cTmp As String
Dim cWidthLimit As Double
Dim tPdtBar As String, tPdtName As String, tPdtName2 As String
Dim tPdtDpt As String, tPdtDptCut As String
Dim cQty As Double, cRegPrice As Double, cSalePrice As Double, cNet As Double
Dim tTaxType As String, cVatRate As Double, cVatAmt As Double
Dim tPdtType As String
Dim tCancelFlag  As String
Dim tTxMk As String, tTrMk As String, tNdMk As String
Dim tMallCardFlag As String
Dim tEmpCardFlag As String
Dim tBPFlag As String
Dim cWeight As String
Dim cDiscVal As Double
Dim cDutyFreeDeposit As Double
Dim bHasLineDisc As Boolean
Dim bHasPromo As Boolean
Dim cHasPromo As Double
'Dim nIsPromo As Boolean '3 : Promo,2 : Promo Time
Dim nIsPromo As Long '3 : Promo,2 : Promo Time    '*Ao 2011/02/15 CommSheet ML-PRV-040-EJ Shows Promo Price wrong-.xls
Dim bHasDeal As Boolean
Dim nDealGrp As Long
Dim tDealNum  As String
Dim tPdtSkuCode As String 'Bump 56-03-02 TK -ISS3000 - 55 :copy จาก  Foodland, CommSheet ST-FL-047-SF- Print UPC or SKU on Receipt and EJ.xls
Dim orsTmp As ADODB.Recordset
Dim orsTmp2 As ADODB.Recordset
Dim orsNew1 As ADODB.Recordset
Dim orsNew2 As ADODB.Recordset

Dim tSql As String

Dim tLogType As String
Dim cLogQty As String
Dim nLogDTSeq As Long

Dim tPriceOver As String

Dim nDealGroup  As Long

Dim cTaxExceptAmt As Double, cDiscPer As Double
Dim tMallEmpFlag As String
Dim tSkuItemType As String '*Tao 60-04-10 SkuItemType ต่อท้าย SkuCode

'Line_Chars
'Const tTkWeight = "Weight"
Dim tTkWeight$
tTkWeight = W_GETxVocabularyTicketFromArray("1942")
'Const tTkDsRegPri = "Reg.Price"
Dim tTkDsRegPri$
tTkDsRegPri = Mid(W_GETxVocabularyTicketFromArray("1900"), 1, 7)

'Const tTkDsDis = "Coop DIS."
Dim tTkDsCoDis$
 tTkDsCoDis = W_GETxVocabularyTicketFromArray("1908")

'Const tTkDsDis = "DIS."
Dim tTkDsDis$
 tTkDsDis = W_GETxVocabularyTicketFromArray("190601")
  'Const tTkDsDisAmt = "DISA."
Dim tTkDsDisAmt$
tTkDsDisAmt = W_GETxVocabularyTicketFromArray("190602")
 'Const tTkDsDisAmt = "DIS."
Dim tTkDsDisNew$
tTkDsDisNew = W_GETxVocabularyTicketFromArray("190603")

 
 'Const tTkDsMkDn = "MKDN."
Dim tTkDsMkDn$
tTkDsMkDn = W_GETxVocabularyTicketFromArray("190701")
Dim tTkDsMkDnAmt$
tTkDsMkDnAmt = W_GETxVocabularyTicketFromArray("190702")
 'Const tTkDsMkDn = "MKDN."
Dim tTkDsMkDnNew$
tTkDsMkDnNew = W_GETxVocabularyTicketFromArray("190703")
 
'Const tTkDsPromo = "Promo"
Dim tTkDsPromo$
tTkDsPromo = Mid(W_GETxVocabularyTicketFromArray("1903"), 1, 7)
'Const tTkDsPromo = "Promo"
Dim tTkDsPromoTime$
tTkDsPromoTime = Mid(W_GETxVocabularyTicketFromArray("1904"), 1, 7)

'Const tTkDsCoCpn = "CoCpn."
Dim tTkDsCoCpn$
tTkDsCoCpn = W_GETxVocabularyTicketFromArray("1917")
'Const tTkDsCoBnd = "CoBnd."
Dim tTkDsCoBnd$
tTkDsCoBnd = W_GETxVocabularyTicketFromArray("1918")

'Const tTkDsDeal = "Deal"
Dim tTkDsDeal$
tTkDsDeal = W_GETxVocabularyTicketFromArray("1905")

'Const tTkPost = "POST#"
Dim tTkPost$
tTkPost = W_GETxVocabularyTicketFromArray("1927")
'Const tTkTran = "TRAN#"
Dim tTkTran$
tTkTran = W_GETxVocabularyTicketFromArray("1928")
'Const tTkAmt = "AMOUNT"
Dim tTkAmt$
tTkAmt = W_GETxVocabularyTicketFromArray("1929")

'Const tTkOrgTotalUpd = "Original Post Total Updated"
Dim tTkOrgTotalUpd$
tTkOrgTotalUpd = W_GETxVocabularyTicketFromArray("1930")

'Const tTkDepNo = "Deposit No.: "
Dim tTkDepNo$
tTkDepNo = W_GETxVocabularyTicketFromArray("1931")

'*Neung 55-07-23 Layaway TK -------------------------------------------------
'Const tTkPhoneNum = "Phone Number: "
Dim tTkPhoneNum$
tTkPhoneNum = W_GETxVocabularyTicketFromArray("1997")
'----------------------------------------------------------------------------------------------

'Const tTkCurrency = "Currency"
Dim tTkCurrency$
tTkCurrency = W_GETxVocabularyTicketFromArray("1932")
Dim tTkCoin$
tTkCoin = W_GETxVocabularyTicketFromArray("5")
Dim tTkGV$
tTkGV = W_GETxVocabularyTicketFromArray("6")
Dim tTkForeign$
tTkForeign = W_GETxVocabularyTicketFromArray("7")
Dim tTkVVch$ '*Tao 62-09-13 เพิ่ม Vendor Voucher (VVch)
tTkVVch = W_GETxVocabularyTicketFromArray("25")


'Const tTkTotal = "TOTAL"
Dim tTkTotal$
tTkTotal = W_GETxVocabularyTicketFromArray("1933")

'Const tTkDepAmt = "Deposit Amt. "
Dim tTkDepAmt$
tTkDepAmt = W_GETxVocabularyTicketFromArray("1934")

'Const tTkVat = "Vat"
Dim tTkVat$
tTkVat = W_GETxVocabularyTicketFromArray("1935")


'Const tTkVoucherNo = "Voucher No.:"
Dim tTkVoucherNo$
tTkVoucherNo = W_GETxVocabularyTicketFromArray("1936")

'Const tTkLine = "- - LINE "
Dim tTkLine$
tTkLine = W_GETxVocabularyTicketFromArray("1937")
'Const tTkLineCancel = "CANCELLED "
Dim tTkLineCancel$
tTkLineCancel = W_GETxVocabularyTicketFromArray("1938")
'Const tTkLineQtyChange = "Qty Changed TO "
Dim tTkLineQtyChange$
tTkLineQtyChange = W_GETxVocabularyTicketFromArray("1939")
'Const tTkLineNagativeScan = "Cancelled "
Dim tTkLineNagativeScan$
tTkLineNagativeScan = W_GETxVocabularyTicketFromArray("1940")
'Const tTkLineQuantity = "Quantity "
Dim tTkLineQuantity$
tTkLineQuantity = W_GETxVocabularyTicketFromArray("1941")

'Const tTkDsRegPri = "Reg.Price"
Dim tTkDsMBPri$
tTkDsMBPri = Mid(W_GETxVocabularyTicketFromArray("1989"), 1, 7)

'Const tTkDsRegPri = "Reg.Price"
Dim tTkDsMPDis$
tTkDsMPDis = Mid(W_GETxVocabularyTicketFromArray("1990"), 1, 7)

'Const tTkLnVendorCpn = "Line Vendor Coupon"
Dim tTkLnVendorCpn$
tTkLnVendorCpn = W_GETxVocabularyTicketFromArray("1991") '*KT 53-10-16 Line Vendor Coupon
Dim tIndexLnVendorCpn As String '*KT 53-10-16 Line Vendor Coupon
Dim tQtyLnVendorCpn As String '*KT 53-10-16 Line Vendor Coupon
Dim nQtyLnVendorCpn As Double '*KT 53-10-16 Line Vendor Coupon
Dim cLnVendorCpnNet As Double '*KT 53-10-16 Line Vendor Coupon
Dim cLnVendorCpnPerQty  As Double '*KT 53-10-16 Line Vendor Coupon

Dim bPass As Boolean
Dim nI As Long

Dim tTkCooCpn As String
tTkCooCpn = "Cooperate Coupon "

Dim tTkCooBrn As String
tTkCooBrn = "Cooperate Brand "

Dim tLen As String

Dim bDealTypeFree  As Boolean

Dim tTmpBlank$

Dim nPosNo As Long
Dim tTrnNo As String
Dim cNetBDiscAll As Double
Dim cSumDiscOnLn As Double '*KT 51-10-01  แก้ไข net amount
Dim cB4Dis As Double '*KT 51-01-24 เก็บค่า Net Amount (qty*price)ก่อนหักส่วนลดใดๆๆ

'Deal
Dim bPntDeal As Boolean
Dim bPntDealST As Boolean
Dim bPntDealEnd As Boolean
Dim bCross As Boolean
Dim nOldPosition As Long
Dim bIsKeep As Boolean
Dim cHasDeal As Double
    
Dim tInStoreBar As String
Dim tSdtStaRfd As String
Dim cMPrice  As Double
Dim bHasMPrice As Boolean
Dim bHasMPDisc As Boolean
Dim nLenTha As Long
Dim tTempName1 As String
Dim tTempName2 As String

Dim tPrintWeight As String
Dim cHasMember21  As Double

Dim tSKUCode As String  '*Eaw 56-02-20
Dim orsOverP As ADODB.Recordset '*Eaw 56-03-13
Dim cPriceOver As Double '*Eaw 56-03-13
Dim tForFac As String  '*Eaw 56-04-11
'--------------------------------------
'*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
'Dim tDealPromo As String
'Dim tHaveDealPromo As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'*KT 55-08-07 Takashimaya
Dim nLoop As Long
Dim tType As String
Dim nRecordCount As Integer, nCount As Integer
Dim tSpace As String

Dim orsFor As ADODB.Recordset

On Error GoTo ErrHandler

        'Detail
'                tHaveDealPromo = "" '*KT 52-11-30 Deal&Promo ''*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
                bW_ChkMark = True           '*Ao 2013/03/06
                bW_ChkMEM = False
                bW_ChkEmp = False
'                tSpace = "  "                                        '*Ao 2013/04/22  CommSheet TK-ISS3000-095 00.00.02.xls
                tSpace = ""                                        '*Ao 2013/04/30  CommSheet TK-ISS3000-095 00.00.04.xls
                
                Set oPrn.Font = oW_Font(1)
                                              
                tTrMk = tW_TrMk
                cWidthLimit = cW_PageWidth - cXC
                
                Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
                                 
               '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
               
               '*KT 51-07-30 Communication Sheet ML-SF-372(2) Cooperate Coupon and Cooperate Brand are wrong.xls
                '==================================================================
                With orsCD
                    If Not orsCD Is Nothing And orsCD.State = adStateOpen Then
                    ReDim aW_CooCpnPer(0) 'clear array
                    .Filter = "FNDctNo = 9  OR  FNDctNo = 10" 'CooCpn,CooBrn
                    .Sort = "FNScdSeqNo ASC"
                    If Not .EOF Then
                        Do While Not .EOF
                            tTmp = SP_FEDtChkString(orsCD, "FCScdDiscPerOrg")
                            nTmp = SP_FEDvChkNumeric(orsCD, "FNDctNo")
                            bPass = True
                            For nI = 0 To UBound(aW_CooCpnPer) - 1
                                If aW_CooCpnPer(nI) = tTmp Then bPass = False: Exit For
                            Next nI
                            aW_CooCpnPer(UBound(aW_CooCpnPer)) = tTmp
                            ReDim Preserve aW_CooCpnPer(UBound(aW_CooCpnPer) + 1)
                            If bPass Then
                                If tTmp <> "" Then
                                    Select Case nTmp
                                        Case 9
                                            tTmp = tTkCooCpn & " " & IIf(tTmp = "", 0, tTmp) & "%"
                                        Case 10
                                            tTmp = tTkCooBrn & " " & IIf(tTmp = "", 0, tTmp) & "%"
                                    End Select
                                    Set oPrn.Font = oW_Font(1)
                                    Call oPrn.PrintText(Mid(tTmp, 1, nVB_MaxChar), cXC, , vbCenter, True)
                                End If
                            End If
                            .MoveNext
                        Loop
                    End If
                    .Filter = adFilterNone 'vbNullString
                End If
                End With
                '============================================================
'               orsDT.MoveFirst
                If Not (orsDT.BOF And orsDT.EOF) Then           '*Ao 2013/02/04
                    orsDT.Filter = adFilterNone
                    orsDT.MoveFirst
                End If
               Select Case Val(tW_TransTypeSub) '
               Case 2, 12, 24, 37, 41, 46 'sart bank,pickup money,deposit money,over/short report
                             tType = SP_FEDtChkString(orsHD, "FTCcyCode")
                             If Val(tW_TransTypeSub) = 2 And orsDT.RecordCount = 2 Then
                                With orsDT
                                        nLoop = 0
                                         Do While Not orsDT.EOF
                                            nLoop = nLoop + 1
                                            If nLoop = 1 Then
                                                'coin
                                                Call oPrn.PrintText(tTkCoin, cXL, , vbLeftJustify, False) 'tTkCurrency
                                                Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                                Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                            Else
                                                'currency
                                                Call oPrn.PrintText(tTkCurrency, cXL, , vbLeftJustify, False) 'tTkCurrency
                                                Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                                Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                            End If
                                            orsDT.MoveNext
                                        Loop
                                        'total
                                        Call oPrn.PrintText(tTkTotal, cXL, , vbLeftJustify, False) 'tTkTotal
                                        Call oPrn.PrintText(Format(cW_Total, tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                        Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                        'blank
                                        Call oPrn.PrintText("", cXL, , vbLeftJustify, True)
                                            
                                End With
                            ElseIf (Val(tW_TransTypeSub) = 12 Or Val(tW_TransTypeSub) = 24 Or Val(tW_TransTypeSub) = 37 Or Val(tW_TransTypeSub) = 41 Or Val(tW_TransTypeSub) = 46) And tType = "0" Then
                                Dim cTotalCashOverShort As Double
                                cTotalCashOverShort = 0
                                '*BG 57-04-29  PRJ10001TK-PH2-External Spec-Cash Transfer 00.02.03 FSBT.docx
                                If Val(tW_TransTypeSub) = 12 Then
                                    Call W_PRNxTicketCashTrDeno(orsDT)
                                Else
                                    With orsDT
                                            nLoop = 0
                                             Do While Not orsDT.EOF
                                                nLoop = nLoop + 1
                                               Select Case nLoop
                                                    Case 1
                                                        'coin
                                                        Call oPrn.PrintText(tTkCoin, cXL, , vbLeftJustify, False) 'tTkCurrency
                                                        Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                                        Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                                        cTotalCashOverShort = cTotalCashOverShort + SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
                                                    Case 2
                                                        'currency
                                                        Call oPrn.PrintText(tTkCurrency, cXL, , vbLeftJustify, False) 'tTkCurrency
                                                        Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                                        Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                                        cTotalCashOverShort = cTotalCashOverShort + SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
                                                    Case 3
                                                        'GV
                                                        Call oPrn.PrintText(tTkGV, cXL, , vbLeftJustify, False) 'tTkGV
                                                        Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                                        Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                                    Case 4
                                                        '*Tao 62-09-13 ตรวจสอบก่อนว่า tVB_TransactionSub Type ไหน
                                                        Select Case Val(tW_TransTypeSub)
                                                            Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                                                                'Vendor Voucher
                                                                Call oPrn.PrintText(tTkVVch, cXL, , vbLeftJustify, False) 'tTkForeign
                                                                Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                                                Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                                            Case Else
                                                                'ของเดิม
                                                                'Foreign(Base)
                                                                Call oPrn.PrintText(tTkForeign, cXL, , vbLeftJustify, False) 'tTkForeign
                                                                Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                                                Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                                        End Select
                                                    Case 5
                                                        '*Tao 62-09-13 ตรวจสอบก่อนว่า tVB_TransactionSub Type ไหน
                                                        Select Case Val(tW_TransTypeSub)
                                                            Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                                                                'ถ้ามี Vendor Voucher Seq จะเลื่อนลงมา
                                                                'Foreign(Base)
                                                                Call oPrn.PrintText(tTkForeign, cXL, , vbLeftJustify, False) 'tTkForeign
                                                                Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                                                Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                                        End Select
                                                End Select
                                                
                                                If nLoop = 2 And (Val(tW_TransTypeSub) = 37 Or Val(tW_TransTypeSub) = 41 Or Val(tW_TransTypeSub) = 46) Then
                                                    'Call oPrn.PrintText(tTkTotal, cXL, , vbLeftJustify, False) 'tTkTotal
                                                    Call oPrn.PrintText("TOTAL CASH", cXL, , vbLeftJustify, False)
                                                    Call oPrn.PrintText(Format(cTotalCashOverShort, tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                                    Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                                End If
                                                orsDT.MoveNext
                                            Loop
                                            'total
                                            If Val(tW_TransTypeSub) <> 37 And Val(tW_TransTypeSub) <> 41 And Val(tW_TransTypeSub) <> 46 Then
                                                Call oPrn.PrintText(tTkTotal, cXL, , vbLeftJustify, False) 'tTkTotal
                                                'Call oPrn.PrintText("TOTAL CASH", cXL, , vbLeftJustify, False)
                                                Call oPrn.PrintText(Format(cW_Total, tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                                Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
                                            End If
                                            'blank
                                            orsDT.Filter = adFilterNone
                                            'If Val(tW_TransTypeSub) = 37 Then
                                            Select Case Val(Format(tW_TransTypeSub))
                                                Case 37, 41, 46
                                                    Call W_PRNbTicketOverShort
                                            End Select
                                            'End If
                                            Call oPrn.PrintText("", cXL, , vbLeftJustify, True)
                                    End With
                                End If
                                '--------------------------------------------------------------------------------------------------------------------------
'                                With orsDT
'                                        nLoop = 0
'                                         Do While Not orsDT.EOF
'                                            nLoop = nLoop + 1
'                                           Select Case nLoop
'                                                Case 1
'                                                    'coin
'                                                    Call oPrn.PrintText(tTkCoin, cXL, , vbLeftJustify, False) 'tTkCurrency
'                                                    Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
'                                                    Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
'                                                    cTotalCashOverShort = cTotalCashOverShort + SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                                Case 2
'                                                    'currency
'                                                    Call oPrn.PrintText(tTkCurrency, cXL, , vbLeftJustify, False) 'tTkCurrency
'                                                    Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
'                                                    Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
'                                                    cTotalCashOverShort = cTotalCashOverShort + SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                                Case 3
'                                                    'GV
'                                                    Call oPrn.PrintText(tTkGV, cXL, , vbLeftJustify, False) 'tTkGV
'                                                    Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
'                                                    Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
'                                                Case 4
'                                                    'Foreign(Base)
'                                                    Call oPrn.PrintText(tTkForeign, cXL, , vbLeftJustify, False) 'tTkForeign
'                                                    Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt"), tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
'                                                    Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
'                                            End Select
'
'                                            If nLoop = 2 And (Val(tW_TransTypeSub) = 37 Or Val(tW_TransTypeSub) = 41 Or Val(tW_TransTypeSub) = 46) Then
'                                                'Call oPrn.PrintText(tTkTotal, cXL, , vbLeftJustify, False) 'tTkTotal
'                                                Call oPrn.PrintText("TOTAL CASH", cXL, , vbLeftJustify, False)
'                                                Call oPrn.PrintText(Format(cTotalCashOverShort, tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
'                                                Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
'                                            End If
'                                            orsDT.MoveNext
'                                        Loop
'                                        'total
'                                        If Val(tW_TransTypeSub) <> 37 And Val(tW_TransTypeSub) <> 41 And Val(tW_TransTypeSub) <> 46 Then
'                                            Call oPrn.PrintText(tTkTotal, cXL, , vbLeftJustify, False) 'tTkTotal
'                                            'Call oPrn.PrintText("TOTAL CASH", cXL, , vbLeftJustify, False)
'                                            Call oPrn.PrintText(Format(cW_Total, tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
'                                            Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
'                                        End If
'                                        'blank
'                                        orsDT.Filter = adFilterNone
'                                        'If Val(tW_TransTypeSub) = 37 Then
'                                        Select Case Val(Format(tW_TransTypeSub))
'                                            Case 37, 41, 46
'                                                Call W_PRNbTicketOverShort
'                                        End Select
'                                        'End If
'                                        Call oPrn.PrintText("", cXL, , vbLeftJustify, True)
'                                End With
'--------------------------------------------------------------------------------------------------------------------------------------------------
                            Else
                                Dim cTotalAmt  As Double
                                Dim cRateCurrency As Double
                                Dim nSpace As Integer, nSpaceEx As Integer
                                If Val(tW_TransTypeSub) = 12 Or Val(tW_TransTypeSub) = 24 Or Val(tW_TransTypeSub) = 37 Or Val(tW_TransTypeSub) = 41 Or Val(tW_TransTypeSub) = 46 Then           '*Ao 2013/09/04   เพิ่ม Val(tW_TransTypeSub) = 24
'                                    Call W_PRNxTicketCashTr(orsDT) '*BG 57-04-28  PRJ10001TK-PH2-External Spec-Cash Transfer 00.02.03 FSBT.docx
                                    If Val(tW_TransTypeSub) = 12 Then
                                        Call W_PRNxTicketCashTrNO(orsDT)
                                    Else
                                        Call W_PRNxTicketCashTr(orsDT)
                                    End If
                                    
                                Else
                                    '*Ao 2013/05/03  CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
                                    tSql = "SELECT FTFcrCode,FTFcrSign,FTFcrType,FCFcrExchang FROM TCNMForCurrency" '*Ao 2013/05/03 CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
                                    Call SP_SQLvExecute(tSql, orsFor, oDB.ocnOnLine, , , , , True, True) 'read
                                    nSpace = 25
                                    For nLoop = 0 To 3
                                        With orsDT
                                            .Filter = "FTSdtStaPdt='" & CStr(nLoop) & "'"
                                            .Sort = "FNSdtSeqNo ASC"
                                            
                                            If Not (orsDT.BOF And orsDT.EOF) Then
                                                Do While Not orsDT.EOF
                                                    tTmp = ""
                                                    If nLoop = 3 Then
'                                                                cRateCurrency = SP_FEDvChkNumeric(orsDT, "FCSdtStkFac")
'                                                                Call oPrn.PrintText(Format(cRateCurrency, "#0.00000"), cXL, , vbLeftJustify, False)
                                                        cRateCurrency = SP_FEDvChkNumeric(orsDT, "FCSdtStkFac")
                                                        tTmp = Format(cRateCurrency, "#0.00000")
                                                    Else
                                                        If nLoop <> 2 Then
                                                            cQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
                                                            If Val(tW_TransTypeSub) = 37 Then
                                                                tTmp = Format(cQty, "###0") & "X"
                                                            Else
                                                                tTmp = Format(cQty, "###0")
                                                            End If
                                                        End If
                                                    End If
                                                    tPdtName = SP_FEDtChkString(orsDT, "FTSkuAbbName") 'SKU Name  Abb Eng
                                                    tPdtName = Mid(tPdtName, 1, 20) 'SKU Name  Abb Eng
                                                    If nSpace > (Len(tTmp) + Len(tPdtName)) Then
                                                        nSpaceEx = nSpace - (Len(tTmp) + Len(tPdtName))
                                                    Else
                                                        nSpaceEx = 0
                                                    End If
                                                    If nLoop = 3 Then
                                                        Call oPrn.PrintText(tPdtName & Space(nSpaceEx) & tTmp, cXL, , vbLeftJustify, False)     '*Ao 2013/05/03  CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
                                                    Else
                                                        Call oPrn.PrintText(tTmp & Space(nSpaceEx) & tPdtName, cXL, , vbLeftJustify, False)
                                                    End If
                                                    
                                                    If nLoop = 3 Then
                                                        cNet = SP_FEDvChkNumeric(orsDT, "FCSdtRegPrice")
                                                        '----------------------------------------------------------------------------------------------
                                                        tTmp = SP_FEDtChkString(orsDT, "FTSkuCode")        '*Ao 2013/05/03  CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
                                                        tTmp = Trim(tTmp)
                                                        orsFor.Filter = "FTFcrCode='" & tTmp & "'"
                                                        If Not (orsFor.BOF And orsFor.EOF) Then
                                                            tTmp = SP_FEDtChkString(orsFor, "FTFcrSign")
                                                            tTmp = Trim(tTmp)
                                                        Else
                                                            tTmp = ""
                                                        End If
                                                        orsFor.Filter = adFilterNone
                                                        '----------------------------------------------------------------------------------------------
                                                    Else
                                                        cNet = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
                                                        tTmp = ""
                                                    End If
                                                    
                                                    Call oPrn.PrintText(tTmp & Format(cNet, tW_CurFmt), cXMK, , vbRightJustify, True) 'amt          '*Ao 2013/05/03  CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
'                                                            Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXMK, , vbRightJustify, True) 'amt
                                                    
                                                    If nLoop = 0 Or nLoop = 1 Then 'Coin + Curreny = Cash Total
                                                        cTotalAmt = cTotalAmt + cNet
                                                    End If
                                                    .MoveNext
                                                Loop
                                            End If
                                                'Print Cash Total
                                                'If nLoop = 1 And Val(tW_TransTypeSub) <> 37 Then
                                            If nLoop = 1 Then
'                                                    Call oPrn.PrintText("TOTAL CASH", cXC, , vbRightJustify, False)
'                                                    Call oPrn.PrintText(Format(cTotalAmt, tW_CurFmt), cXMK, , vbRightJustify, True) 'amt       '*Ao 2013/05/03  CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
                                                tTmp = "TOTAL CASH"
                                                nSpaceEx = nSpace - Len(tTmp)
                                                tTmp = Space(nSpaceEx) & tTmp
                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                                                Call oPrn.PrintText(Format(cTotalAmt, tW_CurFmt), cXMK, , vbRightJustify, True)     '*Ao 2013/05/03  CommSheet TK-ISS3000-108 00.00.03 FSBT NG.xlsx
                                            End If
                                        End With
                                    Next nLoop
                                End If
                                orsDT.Filter = adFilterNone
                                'If Val(tW_TransTypeSub) = 37 Then
                                Select Case Val(Format(tW_TransTypeSub))
                                    Case 37, 41, 46
                                        Call W_PRNbTicketOverShort
                                End Select
                                'End If
                                'blank
                                Call oPrn.PrintText("", cXL, , vbLeftJustify, True)
                            End If
                           
'               Case 12, 24 'pickup,depsit money
'
'                                    'currency
'                                    Call oPrn.PrintText(tTkCurrency, cXL, , vbLeftJustify, False) 'tTkCurrency
'                                    Call oPrn.PrintText(Format(cW_Total, tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
'                                    Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
'
'                                    'total
'                                    Call oPrn.PrintText(tTkTotal, cXL, , vbLeftJustify, False) 'tTkTotal
'                                    Call oPrn.PrintText(Format(cW_Total, tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
'                                    Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign
               
               Case 15, 21

                                    'voucher no
                                    Call oPrn.PrintText(tTkVoucherNo, cXHL - cXHC, , vbLeftJustify, False) 'voucher
                                    Call oPrn.PrintText(tW_VoicherNo, cXC, , vbLeftJustify, True)  'amt
                                    
                                    'amount
                                    Call oPrn.PrintText(tTkAmt, cXL, , vbLeftJustify, False) 'amount
                                    Call oPrn.PrintText(Format(cW_Total, tW_CurFmt), cXMK, , vbRightJustify, False) 'amt
                                    Call oPrn.PrintText(tTrMk, cXMK, , vbLeftJustify, True) 'sign

                                    
               Case 3, 4, 6, 7, 10, 11, 13, 16, 19, 20, 22, 23, 29, 30, 38, 39, 43, 44 '*KT 51-08-14 Communication Sheet ML-SF-459 Other receive and Other disbursement format receipt wrong.xls
               
                        If orsDT.BOF And orsDT.EOF Then
                            Set orsTmp = Nothing
                            Set orsTmp2 = Nothing
                            Set orsNew1 = Nothing
                            Set orsNew2 = Nothing
                            Exit Function
                        End If
                        'deposit =  =  =  =  =  =  =  =  =  =  =  =  =  = =  =  =  =  =  =  =  =  =  =  = =  =  =  =  =  =  =  =  =  =  = =  =  =  =  =  =  =  =  =  =  =
                        If bW_DepositMode Then
                            If tW_TypePrinter <> "0" Then
                                tTmp = eW_Deposit.tEN_DepositNo             '*Ao 2013/04/22  CommSheet TK-ISS3000-095 00.00.02.xls
                            Else
                                If eW_Deposit.tEN_DepositNo = "" Then
                                    tTmp = "---" & tSpace
                                Else
                                    tTmp = eW_Deposit.tEN_DepositNo & tSpace
                                End If
                            End If
                                    

                                    
                            Call oPrn.PrintText(tTkDepNo, cXL, , vbLeftJustify, False)   'dep no
'                            Call oPrn.PrintText(IIf(eW_Deposit.tEN_DepositNo = "", "---", eW_Deposit.tEN_DepositNo), cXR, , vbRightJustify, True) ' OrgTotalTpd ''*KT 52-10-02 DCR_10004 Phase 3.0
'                            Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)          '*Ao 2013/04/22  CommSheet TK-ISS3000-095 00.00.02.xls
                            Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True, , , True, True)        '*Ao 2013/07/16
                            
                            '*Neung 55-07-23 Layaway TK ---------------------------------------------------------------------
                            Call oPrn.PrintText(tTkPhoneNum, cXL, , vbLeftJustify, False)   'Phone Number
'                            Call oPrn.PrintText(IIf(eW_Deposit.tEN_PhoneNum = "", "---", eW_Deposit.tEN_PhoneNum), cXR, , vbRightJustify, True)
                            If tW_TypePrinter <> "0" Then
                                tTmp = eW_Deposit.tEN_PhoneNum             '*Ao 2013/04/22  CommSheet TK-ISS3000-095 00.00.02.xls
                            Else
                                If eW_Deposit.tEN_PhoneNum = "" Then
                                    tTmp = "---" & tSpace
                                Else
                                    tTmp = eW_Deposit.tEN_PhoneNum & tSpace
                                End If
                            End If
                            Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)
                            '------------------------------------------------------------------------------------------------------------------
                        End If
                        Select Case Val(tW_TransTypeSub) '
                            Case nEN_TrnDepositCancel, nEN_TrnAddDepositCancel 'Add deposit cancel      '*Ao 2013/04/05   CommSheet TK-ISS3000-110 00.00.03 FSBT.xlsx
'                                Call oPrn.PrintText("RETURN", cXL, , vbLeftJustify, True)  'Sysbom of Deposit cancel,Add deposit cancel
                        End Select
                        
                        'LogDT print  to ej only =  =  =  =  =  =  =  =  =  =  =  =  =  = =  =  =  =  =  =  =  =  =  =  = =  =  =  =  =  =  =  =  =  =  = =  =  =  =  =  =  =  =  =  =  =
                        '*Ao 2014/02/07
'                        tSql = "SELECT FNSdtSeqNo,FTLogType,FCLogQty FROM TPSTLogDT  WHERE " & tW_TransWhere & " ORDER BY FNLogSeqNo "
                        tSql = "SELECT FNSdtSeqNo,FTLogType,FCLogQty FROM TPSTLogDT  WHERE " & tW_WhereCD & " ORDER BY FNLogSeqNo "
                        'Set orsTmp = ocnAll.Execute(tSql)
                        If SP_SQLvExecute(tSql, orsTmp) = 0 Then
                            With orsTmp
                                Do While Not .EOF
                                            nLogDTSeq = SP_FEDvChkNumeric(orsTmp, "FNSdtSeqNo")
                                            tLogType = SP_FEDtChkString(orsTmp, "FTLogType")
                                            cLogQty = SP_FEDvChkNumeric(orsTmp, "FCLogQty")
                                            
                                            Call oPrn.PrintText(tTkLine & nLogDTSeq, cXL, , vbLeftJustify, False, eEN_EJOnly)    'dep no
                                            Select Case tLogType
                                            Case 0
                                                    Call oPrn.PrintText(tTkLineCancel, cXHL, , vbLeftJustify, True, eEN_EJOnly)    'dep no
                                            Case 1
                                                    Call oPrn.PrintText(tTkLineQtyChange & cLogQty, cXHL, , vbLeftJustify, True, eEN_EJOnly)    'dep no
                                            Case 2
                                                    Call oPrn.PrintText(tTkLineNagativeScan & " " & cLogQty & " " & tTkLineQuantity, cXHL, , vbLeftJustify, True, eEN_EJOnly)      'dep no
                                            End Select
                                orsTmp.MoveNext
                                Loop
                            End With
                        End If
                        'DT =  =  =  =  =  =  =  =  =  =  =  =  =  = =  =  =  =  =  =  =  =  =  =  = =  =  =  =  =  =  =  =  =  =  = =  =  =  =  =  =  =  =  =  =  =
                        ReDim aW_CrossDeal(0)
                        bIsKeep = False
                        
                        Set orsTmp = orsCD.Clone 'for Any Discount
                        orsTmp.Filter = adFilterNone 'vbNullString
                        Set orsTmp2 = orsDT.Clone
                        orsTmp2.Filter = adFilterNone 'vbNullString
                        
                        'ตรวจสอบว่ารายการแรกที่จะพิมพ์เป็น Deal หรือไม่
                        'ถ้าเป็น Deal ต้องนำรายการแรกในกลุ่มที่ได้ Deal นั้นมาพิพพ์ก่อน และต้อง Mark รายการเดิมไว้เพื่อเมื่อพิมพ์รายการ Deal เสร็จจะกลับมาได้
                        bIsKeep = False
                        If Not orsTmp2.EOF Then
                            orsTmp.Filter = "FNSdtSeqNo =  1"
                            orsTmp.Sort = "FNScdSeqNo ASC"  'เรียงตามการเกิด Discount
                            
                            tDealNum = SP_FEDtChkString(orsTmp, "FTDctRefCode")
                            nDealGroup = SP_FEDvChkNumeric(orsTmp, "FNScdDealGroup")
                            'check DT in deal group
                            '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                            'orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FTScdDiscType<>'4' AND FNSdtSeqNo =  1 AND  FNDctNo = " & nEN_Disc01Deal
                            orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & "  AND FNSdtSeqNo =  1 AND  FNDctNo = " & nEN_Disc01Deal
                            orsTmp.Sort = "FNScdSeqNo ASC"
                            If orsTmp.RecordCount > 0 Then
                                     '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                     'orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FTScdDiscType<>'4' AND  FNDctNo = " & nEN_Disc01Deal
                                     orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & "  AND  FNDctNo = " & nEN_Disc01Deal
                                     orsTmp.Sort = "FNScdSeqNo ASC"
                                    '*KT 54-05-12 V.3.0.58 แก้ไขเรื่อง Free Type 002  เก็บ Record แรกของกลุ่ม FTDctRefCode,FNScdDealGroup
                                    'CommSheet ML-PRV-055 Not free item on Free Type2.xls
                                    '*KT 54-09-02 V.4.0.1 แก้ไขเรื่อง Deal Type 003 ให้สามารถกำหนดกลุ่ม Get และ Buy มีสินค้าเดียวกัน
                                    '*KT 54-09-23 V.4.0.1 แก้ไขเรื่อง Deal Type 003 ให้สามารถกำหนดกลุ่ม Get และ Buy มีสินค้าเดียวกัน ทั้งแบบ %off,amount off,new price
'                                    If ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "002") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "4")) Or ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "003") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "1")) Then
                                    If ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "002") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "4")) Or ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "003") And ((SP_FEDtChkString(orsTmp, "FTScdDiscType") = "1") Or (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "2") Or (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "3"))) Then
                                            orsTmp.Sort = "FCScdAmt ASC,FNScdSeqNo ASC"
                                    End If
                                     nSeq = SP_FEDvChkNumeric(orsTmp, "FNSdtSeqNo")
                                     nOldPosition = orsTmp2.AbsolutePosition
                                     Do While Not orsTmp2.EOF
                                        If nSeq = SP_FEDvChkNumeric(orsTmp2, "FNSdtSeqNo") Then
                                            Exit Do
                                        End If
                                        orsTmp2.MoveNext
                                     Loop
                                     bIsKeep = True
                            End If
                         End If

                        Do While Not orsTmp2.EOF
                                    'move DT Record
                                    orsDT.Find "FNSdtSeqNo = " & SP_FEDvChkNumeric(orsTmp2, "FNSdtSeqNo"), , , 1
                                    If orsDT.EOF Then
                                        Set orsTmp = Nothing
                                        Set orsTmp2 = Nothing
                                        Set orsNew1 = Nothing
                                        Set orsNew2 = Nothing
                                        Exit Function ' Error
                                    End If
                                    'get
                                    nSeq = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
                                    tInStoreBar = SP_FEDtChkString(orsDT, "FTSdtInstoreBar")
                                    tPdtBar = SP_FEDtChkString(orsDT, "FTSdtBarCode")
                                    tCancelFlag = SP_FEDtChkString(orsDT, "FTSdtCancelFlag")
                                    tPrintWeight = Mid(SP_FEDtChkString(orsDT, "FTSdtStaPdt"), 13, 1)
                                    tPdtSkuCode = SP_FEDtChkString(orsDT, "FTSkuCode") '*Bump 56-02-03 TK-ISS3000-055 : Copy จาก  Fooland, CommSheet ST-FL-047-SF- Print UPC or SKU on Receipt and EJ.xls
                                    tSkuItemType = SP_PRCxSkuItemType(tPdtSkuCode, tW_TblDT, bW_UseServer, tW_TransWhere) '*Tao 60-04-10 SkuItemType ต่อท้าย SkuCode
                                    tIndexLnVendorCpn = SP_FEDtChkString(orsDT, "FTSkuIndex") '*KT 53-10-16 Line Vendor Coupon
                                     '*KT 52-05-02 PH 1.8 ML-SF-046 ในกรณีที่ Barcode  20 ตัว และ ชิ่อสินค้ายาว 20 อักษร  ระบบพิมพ์ Receipt ออกมาโดยชื่อสินค้าตกบรรทัด
                                    If tCancelFlag = "1" Then
                                        tPdtName = SP_FEDtChkString(orsDT, "FTSkuAbbName")
                                        tPdtName = Replace(tPdtName, "Canceled - ", "")
                                        tTempName1 = tPdtName
                                        tPdtName = Mid(tPdtName, 1, 20) 'SKU Name  Abb Eng
                                        tPdtName2 = SP_FEDtChkString(orsDT, "FTSkuAbbNameSnd")
                                        tPdtName2 = Replace(tPdtName2, "Canceled - ", "")
                                        tTempName2 = tPdtName2
                                        tPdtName2 = Mid(tPdtName2, 1, 20) 'SKU Name  Abb Thai
                                    Else
                                        'Line_Chars
                                        tPdtName = SP_FEDtChkString(orsDT, "FTSkuAbbName") 'SKU Name  Abb Eng
                                        tTempName1 = tPdtName
                                        tPdtName = Mid(tPdtName, 1, 20) 'SKU Name  Abb Eng
                                        tPdtName2 = SP_FEDtChkString(orsDT, "FTSkuAbbNameSnd") 'SKU Name  Abb Thai
                                        tTempName2 = tPdtName2
                                        tPdtName2 = Mid(tPdtName2, 1, 20) 'SKU Name  Abb Thai
                                    End If
                                    nLenTha = cUT_Print.UT_nThaAbsoluteCount(tPdtName)
                                    Do While Not nLenTha = 0
                                        tPdtName = Mid(tTempName1, 1, 20 + nLenTha) 'SKU Name  Abb Eng
                                        If Len(tPdtName) - nLenTha <= 20 Or Len(tTempName1) <= Len(tPdtName) + nLenTha Then
                                            nLenTha = 0
                                        Else
                                            nLenTha = cUT_Print.UT_nThaAbsoluteCount(tPdtName)
                                        End If
                                    Loop
                                    nLenTha = cUT_Print.UT_nThaAbsoluteCount(tPdtName2)
                                    Do While Not nLenTha = 0
                                        tPdtName2 = Mid(tTempName2, 1, 20 + nLenTha) 'SKU Name  Abb Thai
                                         If Len(tPdtName2) - nLenTha <= 20 Or Len(tTempName2) <= Len(tPdtName2) + nLenTha Then
                                            nLenTha = 0
                                        Else
                                            nLenTha = cUT_Print.UT_nThaAbsoluteCount(tPdtName2)
                                        End If
                                    Loop
                                                                        
                                    tPdtDpt = SP_FEDtChkString(orsDT, "FTDpcCode")
                                    tPriceOver = SP_FEDtChkString(orsDT, "FTSdtPriceOverideFlag") 'Y : Price Overide N : Nomal
                                    
                                    cQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
                                    cWeight = SP_FEDvChkNumeric(orsDT, "FCSdtWeight")
                                    cRegPrice = IIf(UCase(tPriceOver) = "Y", SP_FEDvChkNumeric(orsDT, "FCSdtB4DisChg"), SP_FEDvChkNumeric(orsDT, "FCSdtRegPrice")) '*KT 51-07-25 กรณีสินค้ามี Price Overide จะใช้ RegPrice ไม่ได้ ต้องใช้ FCSdtB4DisChg
                                    cSalePrice = SP_FEDvChkNumeric(orsDT, "FCSdtSalePrice")
'                                    tMallCardFlag = SP_FEDtChkString(orsDT, "FTSkuStaMallEmpCard")   '*Eaw 55-12-07 test
'                                    tEmpCardFlag = SP_FEDtChkString(orsDT, "FTSkuStaEmpCard")   '*Eaw 55-12-06 เก็บ Flag ของ Alw Emp Card
                                    '*KT 55-08-09 Takashimaya PRJ10001TK-External Specification SKU-00.01.00.doc
'                                    If Mid(SP_FEDtChkString(orsDT, "FTSdtStaPdt"), 9, 1) = "1" Then  '*eaw 55-12-07 test                   '*Ao 2013/03/06 Comment
'                                         tMallCardFlag = "Y"
'                                    Else
'                                        tMallCardFlag = "N"
'                                    End If
'                                    If Mid(SP_FEDtChkString(orsDT, "FTSdtStaPdt"), 20, 1) = "1" Then
'                                         tEmpCardFlag = "Y"
'                                    Else
'                                        tEmpCardFlag = "N"
'                                    End If
'                                    If Mid(SP_FEDtChkString(orsDT, "FTSdtStaPdt"), 21, 1) = "1" Then
'                                         tBPFlag = "Y"
'                                    Else
'                                         tBPFlag = "N"
'                                    End If

                                    '*Eaw 56-03-13   CommSheet TK-ISS3000-060 00.00.02 FSBT NG.xlsx
                                    tSql = "SELECT FCLineOver FROM  " & tW_TblDT & " WHERE " & tW_TransWhere & " AND FTSdtBarCode = '" & tPdtBar & "'" & "AND FNSdtSeqNo = " & nSeq
                                    If SP_SQLvExecute(tSql, orsOverP) = 0 Then
                                        cPriceOver = SP_FEDvChkNumeric(orsOverP, "FCLineOver")
                                        If cPriceOver > 0 Then
                                            cRegPrice = SP_FEDvChkNumeric(orsDT, "FCSdtRegPrice")
                                        End If
                                    End If
                                    '*Eaw 56-08-14
                                    Select Case Val(Format(tW_TransTypeSub))
                                        Case 43, 44
                                            cRegPrice = SP_FEDvChkNumeric(orsDT, "FCSdtTotalBFVat") / cQty
                                            cSalePrice = cRegPrice
                                    End Select
                                    '----------------------------------------
                                    cDutyFreeDeposit = SP_FEDvChkNumeric(orsDT, "FCSdtDutyDeposit")
                                    cB4Dis = SP_FEDvChkNumeric(orsDT, "FCSdtB4DisChg")
                                    cMPrice = SP_FEDvChkNumeric(orsDT, "FCSdtMemberPrice")
                                    tSdtStaRfd = SP_FEDtChkString(orsDT, "FTSdtStaRfd")
                                    
                                    '*KT 52-09-08 Member Price + Consignment
                                    orsTmp.Filter = adFilterNone
                                    orsTmp.Filter = "FNDctNo = 21"  ' type member price + non consignment"
                                    orsTmp.Find "FNSdtSeqNo =  " & nSeq, , , 1
                                    bHasMPrice = IIf(orsTmp.EOF, False, True)
                                    
                                    
                                    orsTmp.Filter = adFilterNone
                                    orsTmp.Filter = "FNDctNo = 22" ' type member price + consignment
                                    orsTmp.Find "FNSdtSeqNo =  " & nSeq, , , 1
                                    bHasMPDisc = IIf(orsTmp.EOF, False, True)
                                    
                                    cNet = cQty * cSalePrice
                                    cNetBDiscAll = cQty * cRegPrice
                                    
                                    tPdtDptCut = Left(tPdtDpt, 5)
                                    '*KT 55-08-09 Takashimaya PRJ10001TK-External Specification SKU-00.01.00.doc
                                    'การแสดงเครื่องหมาย "#","@" หน้ารหัสสินค้า
                                    ' tNdMk = IIf(tMallCardFlag = "N", "*", vbNullString) 'none disc mall card
'                                    tNdMk = ""
                                    '*Eaw 55-12-06  เปลี่ยนเป็นใช้เครื่องหมาย *สำหรับ Alw Member Card ,# สำหรับ Alw Emp Card
'                                    If UCase(tMallCardFlag = "N") Then tNdMk = "*" 'ไม่ให้ส่วนลดบัตรสมาชิก( Member Discount)                           '*Ao 2013/03/06  Comment
'                                    If UCase(tEmpCardFlag = "N") Then tNdMk = "#"  'ไม่ให้ส่วนลดพนักงาน (Employee Discount)
'                                    If UCase(tMallCardFlag = "N") And UCase(tEmpCardFlag = "N") Then tNdMk = "*"   'ไม่ให้ส่วนลดบัตรสมาชิก( Member Discount)  + ไม่ให้ส่วนลดพนักงาน (Employee Discount)
                                    
'                                    tSKUCode = SP_FEDtChkString(orsDT, "FTSKUCode")                   '*Eaw 56-02-20  New Req Discountable
'                                    tNdMk = W_GETxSkuSign(tNdMk, tSKUCode, tMallCardFlag)         '*Eaw 56-02-20  New Req Discountable
'                                    If UCase(tMallCardFlag) = "Y" And UCase(tBPFlag) = "Y" Then
'                                        tNdMk = "@"
'                                    ElseIf UCase(tMallCardFlag) = "Y" Then
'                                        tNdMk = "#"
'                                    End If
                                    '
                                    tNdMk = ""
                                    tMallCardFlag = SP_FEDtChkString(orsDT, "FTSkuStaMallEmpCard")
                                    If Mid(SP_FEDtChkString(orsDT, "FTSdtStaPdt"), 20, 1) = "1" Then
                                         tEmpCardFlag = "Y"
                                    Else
                                        tEmpCardFlag = "N"
                                    End If
                                    tMallEmpFlag = SP_FEDtChkString(orsDT, "FTPdtNoDis")
                                    tNdMk = W_GETtSkuSign(tMallCardFlag, tEmpCardFlag, tMallEmpFlag)
                                    
                                    tTxMk = IIf(SP_FEDtChkString(orsDT, "FTTaxStaSign") = "Y", SP_FEDtChkString(orsDT, "FTTaxSign"), vbNullString) '  sign
                                    '*KT 51-09-24 แก้ไขให้จำนวนชิ้นที่ Show ใน Ticket และ EJ ไม่รวมสินค้าที่ Cancel
                                    'nW_Items = nW_Items + Abs(cQty)
                                    nW_Items = nW_Items + IIf(cQty < 0, 0, cQty)
                                    
                                    orsTmp.Filter = adFilterNone 'vbNullString
                                    orsTmp.Filter = "FNDctNo = 4 OR FNDctNo = 5 OR FNDctNo = 6 OR FNDctNo = 9 OR FNDctNo = 10  OR FNDctNo = 22"
                                    orsTmp.Find "FNSdtSeqNo =  " & nSeq, , , 1
                                    bHasLineDisc = IIf(orsTmp.EOF, False, True)
                                    
                                    orsTmp.Filter = adFilterNone 'vbNullString
                                    orsTmp.Filter = "(FNDctNo = 2 OR FNDctNo = 3) "
                                    orsTmp.Find "FNSdtSeqNo =  " & nSeq, , , 1
                                    bHasPromo = IIf(orsTmp.EOF, False, True)
                                    If orsTmp.EOF Then
                                        cHasPromo = 0
                                    Else
                                        '*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
                                        If SP_FEDvChkNumeric(orsTmp, "FTSdtStkCode") = "1" Then
                                            cHasPromo = 0
                                        Else
                                            cHasPromo = SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
                                        End If
                                        nIsPromo = SP_FEDvChkNumeric(orsTmp, "FNDctNo")
                                    End If
                                    
'*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
'                                    orsTmp.Filter = adFilterNone 'vbNullString
'                                    orsTmp.Filter = "(FNDctNo = 2 OR FNDctNo = 3 OR FNDctNo = 21)"
'                                    orsTmp.Find "FNSdtSeqNo =  " & nSeq, , , 1
'                                    bHasPromo = IIf(orsTmp.EOF, False, True)
'                                    If orsTmp.EOF Then
'                                        tDealPromo = ""
'                                    Else
'                                        tDealPromo = SP_FEDvChkNumeric(orsTmp, "FTSdtStkCode")
'                                    End If
'                                    If tDealPromo = "1" Then '*KT 52-11-30 Deal&Promo
'                                        tHaveDealPromo = "1"
'                                    End If

                                    orsTmp.Filter = adFilterNone 'vbNullString
                                    orsTmp.Filter = "(FNDctNo = 21) "
                                    orsTmp.Find "FNSdtSeqNo =  " & nSeq, , , 1
                                    If bHasPromo = False Then                                           '*Ao 2011/02/15 CommSheet ML-PRV-040-EJ Shows Promo Price wrong-.xls
                                        bHasPromo = IIf(orsTmp.EOF, False, True)
                                    End If
                                    If orsTmp.EOF Then
                                        cHasMember21 = 0
                                    Else
                                        '*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
                                        If SP_FEDvChkNumeric(orsTmp, "FTSdtStkCode") = "1" Then
                                            cHasMember21 = 0
                                        Else
                                            cHasMember21 = SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
                                        End If
                                    End If
                                    
                                    orsTmp.Filter = adFilterNone ' vbNullString
                                    orsTmp.Filter = "FNDctNo = 1 " ' type promo  and deal
                                    orsTmp.Find "FNSdtSeqNo =  " & nSeq, , , 1
                                    bHasDeal = IIf(orsTmp.EOF, False, True)
                                    If orsTmp.EOF Then
                                        cHasDeal = 0
                                    Else
                                        cHasDeal = SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
                                    End If
                                    
                                    If bHasDeal Then
                                        orsTmp.Filter = adFilterNone 'vbNullString
                                        orsTmp.Filter = "FNDctNo = 1 AND FTScdDiscType = '4' AND FCScdAmt > 0 AND FNSdtSeqNo =" & nSeq  ' type promo  and deal
'                                        orsTmp.Find "FNSdtSeqNo =  " & nSeq, , , 1
                                        bDealTypeFree = IIf(orsTmp.EOF, False, True)
                                    Else
                                        bDealTypeFree = False
                                    End If
                                    
                                    orsTmp.Filter = adFilterNone
                                    
                                    tTaxType = SP_FEDtChkString(orsDT, "FTTaxCode")
                                    cVatRate = SP_FEDvChkNumeric(orsDT, "FCSdtTax")

                                    
                                    
                            If tCancelFlag = "1" Then  'print  to ej only
                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                                    'section1 bar - name
                                    tPdtBar = tNdMk & IIf(tInStoreBar <> "", tInStoreBar, tPdtBar)
                                    '*************************
                                    'Call oPrn.PrintText(tPdtBar, cXL, , vbLeftJustify, False, eEN_EJOnly) 'bar'Bump 56-03-02 Comment บรรทัดนี้
                                     '*Bump 53-03-02 TK-ISS3000-055 Copy จาก Foodland-CommSheet ST-FL-047-SF- Print UPC or SKU on Receipt and EJ.xls
                                     '*เพิ่ม Code 8 บรรทัด
                                    If nVB_CNPrintUPC = 1 Then
                                        'พิมพ์ barcode
                                        Call oPrn.PrintText(tPdtBar, cXL, , vbLeftJustify, False, eEN_EJOnly) '*Bump 56-03-02
                                    Else
                                        'พิมพ์ Skucode
                                        tPdtSkuCode = tNdMk & IIf(tInStoreBar <> "", tInStoreBar, tPdtSkuCode)
                                        'Call oPrn.PrintText(tPdtSkuCode, cXL, , vbLeftJustify, False, eEN_EJOnly) '*Tao 60-04-10 ของเดิม
                                        Call oPrn.PrintText(tPdtSkuCode & tSkuItemType, cXL, , vbLeftJustify, False, eEN_EJOnly) '*Tao 60-04-10 SkuItemType ต่อท้าย SkuCode
                                    End If
                                    '*************************
                                    '*KT 51-08-11 Communication Sheet ML-SF-462 Wrong display ABB Thai on SKU description.xls
                                    Call oPrn.PrintText(tPdtName, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, True, eEN_EJOnly, cWidthLimit)      'pdt name
                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                                    'section2 depart - cancelld
                                    'Call oPrn.PrintText(tPdtDptCut, cXL, , vbLeftJustify, False, eEN_EJOnly) 'deprat
                                    Call oPrn.PrintText(tTkLineCancel, cXHL, , vbLeftJustify, True, eEN_EJOnly)  'cancelled
                            Else
                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                                    'section1 bar - name
                                    tPdtBar = tNdMk & IIf(tInStoreBar <> "", tInStoreBar, tPdtBar)
                                    '******************************
                                    'Call oPrn.PrintText(tPdtBar, cXL, , vbLeftJustify, False) 'bar '*Bump 56-03-02 Comment บรรทัดนี้
                                    '*Bump 56-03-02 TK-ISS3000-055 :Copy จาก Foodland  CommSheet ST-FL-047-SF- Print UPC or SKU on Receipt and EJ.xls
                                    If nVB_CNPrintUPC = 1 Then
                                        'พิมพ์ barcode
                                        Call oPrn.PrintText(tPdtBar, cXL, , vbLeftJustify, False) 'bar
                                    Else
                                        'พิมพ์ Skucode
                                        tPdtSkuCode = tNdMk & IIf(tInStoreBar <> "", tInStoreBar, tPdtSkuCode)
                                        'Call oPrn.PrintText(tPdtSkuCode, cXL, , vbLeftJustify, False) '*Tao 60-04-10 ของเดิม
                                        Call oPrn.PrintText(tPdtSkuCode & tSkuItemType, cXL, , vbLeftJustify, False) '*Tao 60-04-10 SkuItemType ต่อท้าย SkuCode
                                    End If
                                    '******************************
                                    '*KT 51-08-11 Communication Sheet ML-SF-462 Wrong display ABB Thai on SKU description.xls
                                    'Call oPrn.PrintText(tPdtName, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, True, , cWidthLimit)     'pdt name
                                    Call oPrn.PrintText(tPdtName2, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, True, , cWidthLimit)     'pdt name shot '*Tao 60-09-25
                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                                    'section2 depart - qty - price
                                    If Not bHasLineDisc And Not bHasPromo And Not bHasDeal Then  'แสดงแบบปกติ ไม่มีส่วนลด
                                                    'Call oPrn.PrintText(tPdtDptCut, cXL, , vbLeftJustify, False) 'deprat
                                                    If Val(tW_TransTypeSub) = 11 Then '  แบบ except sale- - - - - - - - - - - - - - - - - - - - - -
                                                            If bHasMPrice Then
                                                                tTmp = cQty & "@" & Format(cMPrice, tW_CurFmt)
                                                                tLen = tTmp
'                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price  '*Eaw 56-02-26  TK-ISS3000-068
                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                tTmp = Format(cMPrice * cQty, tW_CurFmt)
                                                                cTaxExceptAmt = cMPrice
                                                            Else
                                                                tTmp = cQty & "@" & Format(cRegPrice, tW_CurFmt)
                                                                tLen = tTmp
'                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price   '*Eaw 56-02-26  TK-ISS3000-068
                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                tTmp = Format(cRegPrice * cQty, tW_CurFmt)
                                                                cTaxExceptAmt = cRegPrice
                                                            End If
                                                    Else ' แบบปกติ- - - - - - - - - - -  - - - - - - - - - - -
                                                            '*KT 52-09-08 Member Price + Consignment
                                                            If bHasMPrice Then
                                                                tTmp = cQty & "@" & Format(cMPrice, tW_CurFmt)
                                                                tLen = tTmp
'                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price   '*Eaw 56-02-26  TK-ISS3000-068
                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                tTmp = Format(cMPrice * cQty, tW_CurFmt)
                                                            Else
                                                                tTmp = cQty & "@" & Format(cSalePrice, tW_CurFmt)
                                                                tLen = tTmp
'                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price   '*Eaw 56-02-26  TK-ISS3000-068
                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                tTmp = Format(cNet, tW_CurFmt)
                                                            End If
                                                    End If
                                                     tLen = tLen & tTmp
                                                    If Len(tLen) > 25 Then 'จำนวน SubTotal เกิน 12 Digit ขึ้นไปให้ขึ้นบรรทัดใหม่
                                                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' ขึ้นบรรทัดใหม่
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                        tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                                    Else
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                        tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                                    End If
                                    Else 'แสดงแบบมีส่วนลด
                                            'Call oPrn.PrintText(tPdtDptCut, cXL, , vbLeftJustify, False) 'deprat ' '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ให้ Free แสดงเหมือนกับ Deal
                                            If bHasPromo And bHasLineDisc And Not bHasDeal Then   'แสดงราคา promo - - - - - - - - - - -  - - - - - - - - - - -'*KT 51-06-20  Communication Sheet ML-SF-345
                                                    'Call oPrn.PrintText(tPdtDptCut, cXL, , vbLeftJustify, False) 'deprat '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ให้ Free แสดงเหมือนกับ Deal
'                                                   If SP_FEDvChkNumeric(orsTmp, "FNDctNo") = nEN_Disc03Pmt Then '*KT 53-02-11
                                                    If nIsPromo = 3 Then
                                                        Call oPrn.PrintText(tTkDsPromo, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) ' promo
                                                    Else
                                                        Call oPrn.PrintText(tTkDsPromoTime, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) ' promo time sale
                                                    End If
                                                    cTmp = W_GETcPromoAmt(nSeq)
                                                    Select Case Val(tW_TransTypeSub)
                                                        Case "11" 'Tax Except Sale
                                                            cTaxExceptAmt = SP_FMTcCurPoint((cRegPrice - (100 * cVatRate / (100 + cVatRate))) - (cTmp / cQty), True, False)
                                                            tTmp = cQty & "@" & Format(cTaxExceptAmt, tW_CurFmt) '*KT 51-12-11 Phase 2 ข้อ 12
                                                            Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, True) 'qty@price
                                                            tTmp = Format(cNet, tW_CurFmt)
                                                        Case Else
                                                            tTmp = cQty & "@" & Format(cRegPrice - (cTmp / cQty), tW_CurFmt) '*KT 51-12-11 Phase 2 ข้อ 12
                                                            Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, True) 'qty@price
                                                            tTmp = Format(cNet, tW_CurFmt)
                                                    End Select
                                            ElseIf bHasPromo And Not bHasLineDisc And Not bHasDeal Then    'แบบปกติ สินค้า promo - - - - - - - - - - -  - - - - - - - - - - -'*KT 51-06-20  Communication Sheet ML-SF-345
                                                    'Call oPrn.PrintText(tPdtDptCut, cXL, , vbLeftJustify, False) 'deprat '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ให้ Free แสดงเหมือนกับ Deal
                                                    Select Case Val(tW_TransTypeSub)
                                                        Case "11" 'Tax Except Sale
                                                                cTaxExceptAmt = cSalePrice + SP_FMTcCurPoint((cVatRate * cSalePrice) / (100), True, False)     '*KT 52-04-07 แก้ไขการคำนวนค่า vat ของรายการ Tax Except Sale ใหม่ ไม่ต้องเอา100+ Rate
                                                                tTmp = cQty & "@" & Format(cTaxExceptAmt, tW_CurFmt)
                                                                tLen = tTmp
'                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price  '*Eaw 56-02-26  TK-ISS3000-068
                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                tTmp = Format(cQty * cTaxExceptAmt, tW_CurFmt)
                                                                tLen = tLen & tTmp
                                                        Case Else
                                                                tTmp = cQty & "@" & Format(cSalePrice, tW_CurFmt)
                                                                 tLen = tTmp
'                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price   '*Eaw 56-02-26  TK-ISS3000-068
                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                tTmp = Format(cNet, tW_CurFmt)
                                                                tLen = tLen & tTmp
                                                    End Select
                                                    If Len(tLen) > 25 Then   'จำนวน SubTotal เกิน 12 Digit ขึ้นไปให้ขึ้นบรรทัดใหม่
                                                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' ขึ้นบรรทัดใหม่
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                        tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                                    Else
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                        tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                                    End If
                                            ElseIf bHasPromo And bHasDeal And cHasDeal <= 0 Then     'แบบปกติ สินค้า deal / promo
                                                    'Call oPrn.PrintText(tPdtDptCut, cXL, , vbLeftJustify, False) 'deprat '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ให้ Free แสดงเหมือนกับ Deal
                                                    Select Case Val(tW_TransTypeSub)
                                                        Case "11" 'Tax Except Sale
                                                                cTaxExceptAmt = cSalePrice + SP_FMTcCurPoint((cVatRate * cSalePrice) / (100), True, False)       '*KT 52-04-07 แก้ไขการคำนวนค่า vat ของรายการ Tax Except Sale ใหม่ ไม่ต้องเอา100+ Rate
                                                                tTmp = cQty & "@" & Format(cTaxExceptAmt, tW_CurFmt)
                                                                tLen = tTmp
'                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price   '*Eaw 56-02-26  TK-ISS3000-068
                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                tTmp = Format(cQty * cTaxExceptAmt, tW_CurFmt)
                                                                tLen = tLen & tTmp
                                                        Case Else
                                                        
'*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
'                                                                 If bHasDeal And bHasPromo And tDealPromo = "1" Then
'                                                                     tTmp = cQty & "@" & Format((cSalePrice - (cHasPromo / cQty)), tW_CurFmt)
'                                                                     tLen = tTmp
'                                                                    Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price
'                                                                    tTmp = Format((cSalePrice - (cHasPromo / cQty)) * cQty, tW_CurFmt)
'                                                                    tLen = tLen & tTmp
'                                                                 Else
                                                                    tTmp = cQty & "@" & Format(cSalePrice, tW_CurFmt)
                                                                     tLen = tTmp
'                                                                    Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price  '*Eaw 56-02-26  TK-ISS3000-068
                                                                    Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                    tTmp = Format(cNet, tW_CurFmt)
                                                                    tLen = tLen & tTmp
'                                                                End If

                                                    End Select
                                                    If Len(tLen) > 25 Then   'จำนวน SubTotal เกิน 12 Digit ขึ้นไปให้ขึ้นบรรทัดใหม่
                                                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' ขึ้นบรรทัดใหม่
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                        tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                                    Else
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                        tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                        Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                                    End If
                                            ElseIf bHasDeal Then
                                                    cTaxExceptAmt = cRegPrice 'cSalePrice + SP_FMTcCurPoint((cVatRate * cSalePrice) / (100), True, False)
                                                    If (Not (bDealTypeFree)) Or (bDealTypeFree And cHasDeal <= 0) Then   '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ให้ Free แสดงเหมือนกับ Deal
                                                        'Call oPrn.PrintText(tPdtDptCut, cXL, , vbLeftJustify, False) 'deprat '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ให้ Free แสดงเหมือนกับ Deal

'*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
'                                                        If bHasDeal And bHasPromo And tDealPromo = "1" Then
'                                                             tTmp = cQty & "@" & Format(cRegPrice - (cHasPromo / cQty), tW_CurFmt)
'                                                             tLen = tTmp
'                                                            Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price
'                                                             'tTmp = IIf(bDealTypeFree, "Free", Format((cRegPrice - (cHasPromo / cQty)) * cQty, tW_CurFmt))
'                                                             tTmp = Format((cRegPrice - (cHasPromo / cQty)) * cQty, tW_CurFmt) '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ให้ Free แสดงเหมือนกับ Deal
'                                                        Else
'*KT 53-01-12 แก้ไขกรณีที่มีการตั้ง Double Policy ว่าได้ Dealแล้วลด Promotime,Promo,Member Price(non consignment)
                                                            If cHasPromo > 0 Then
                                                                tTmp = cQty & "@" & Format(cRegPrice - (cHasPromo / cQty), tW_CurFmt)
                                                                tLen = tTmp
'                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price  '*Eaw 56-02-26  TK-ISS3000-068
                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                tTmp = Format((cRegPrice - (cHasPromo / cQty)) * cQty, tW_CurFmt) '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                                            ElseIf cHasMember21 > 0 Then
                                                                tTmp = cQty & "@" & Format(cRegPrice - (cHasMember21 / cQty), tW_CurFmt)
                                                                tLen = tTmp
'                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price    '*Eaw 56-02-26  TK-ISS3000-068
                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                tTmp = Format((cRegPrice - (cHasMember21 / cQty)) * cQty, tW_CurFmt) '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                                            Else
                                                                tTmp = cQty & "@" & Format(cRegPrice, tW_CurFmt)
                                                                tLen = tTmp
'                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price  '*Eaw 56-02-26  TK-ISS3000-068
                                                                Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                tTmp = Format(cRegPrice * cQty, tW_CurFmt) '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                                            End If
'                                                        End If

                                                        tLen = tLen & tTmp
                                                        If Len(tLen) > 25 Then  'จำนวน SubTotal เกิน 12 Digit ขึ้นไปให้ขึ้นบรรทัดใหม่
                                                            Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' ขึ้นบรรทัดใหม่
                                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                            tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                            Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                                        Else
                                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                            tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                            'Call oPrn.PrintText(IIf(bDealTypeFree, vbNullString, tTmp), cXMK, , vbLeftJustify, True) '  sign
                                                            Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                                        End If
                                                    End If
                                            Else  'แสดงราคา reg pri - - - - - - - - - - -  - - - - - - - - - - -
                                                    'Call oPrn.PrintText(tPdtDptCut, cXL, , vbLeftJustify, False) 'deprat '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ให้ Free แสดงเหมือนกับ Deal
                                                    cTaxExceptAmt = cRegPrice  'cSalePrice + SP_FMTcCurPoint((cVatRate * cSalePrice) / (100), True, False)
                                                    '*KT 52-08-14 PH2-ML-SF- 118 สินค้าที่มีส่วนลด Member Price และ Cooperate Dis ซ้อนกัน ใน Receipt/EJ จะพิมพ์คำว่า Reg Price ให้เปลี่ยนเป็น MB Price
                                                    If bHasMPrice And bHasLineDisc Then
                                                        Call oPrn.PrintText(tTkDsMBPri, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'reg price
                                                    Else
                                                        Call oPrn.PrintText(tTkDsRegPri, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'reg price
                                                    End If
                                                    If bHasMPrice Then
                                                        tTmp = cQty & "@" & Format(cMPrice, tW_CurFmt)
                                                    Else
                                                        tTmp = cQty & "@" & Format(cRegPrice, tW_CurFmt)
                                                    End If
                                                    Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                    tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                    tTmp = ""                       '*Ao 2013/02/05   บังคับว่าง
                                                    If tTmp = "" Then
                                                        tTmp = IIf(Len(tTmp) >= 2, tTmp, tTmp & Space$(2 - Len(tTmp)))
                                                        Call oPrn.PrintText(vbNullString, cXR, , vbRightJustify, True)  ' space - - - - - -
                                                    Else
                                                       tTmp = IIf(Len(tTmp) >= 2, tTmp, tTmp & Space$(2 - Len(tTmp)))
                                                        Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True) '  sign
                                                    End If
                                            End If
                                            
                                    End If
                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                                    'section3 wegith
                                    If cWeight > 0 And (tPrintWeight = "1" Or tPrintWeight = "") Then   'Weight
                                                '*KT 52-04-30 PH1.8-ML-SF-045  พิมพ์ น้ำหนักสินค้าไม่ตรงกับ นน บนBarcode ที่ได้จากเครื่องชั่ง
                                                If cWeight >= 1 Then
                                                        tTmp = cWeight & " " & "Kg."
                                                Else
                                                        cWeight = cWeight * 1000
                                                        tTmp = cWeight & " " & "g."
                                                End If
                                                '*KT 52-05-06 PH 1.8 ML-SF-055 สินค้า Weight Item ในกรณีทีทีการ เปลี่ยนจำนวน  1.  EJ  ต้องมีการ X จำนวนด้วย  ถ้าซื้อของ 0.035g 10 ชิ้น Receipt ต้องออกมา 10@0.035g
                                                If cQty > 1 Then
                                                         tTmp = " " & cQty & "@" & tTmp
                                                Else
                                                         tTmp = " " & tTmp
                                                End If
                                                Call oPrn.PrintText(tTkWeight, cXL, , vbLeftJustify, False)
                                                Call oPrn.PrintText(tTmp, cXHL, , vbLeftJustify, True) '
                                    End If
                                    
                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                                    'section3 tax ecept
'                                    If Val(tW_TransTypeSub) = 11 And tTaxType <> "00" Then
                                    '*Eaw 56-08-02
                                    If Val(tW_TransTypeSub) = 11 Then
'                                            If tTaxType = "01" Then
'                                                    cVatAmt = SP_FMTcCurPoint((cVatRate * (cTaxExceptAmt)) / (100 + cVatRate), True, False) '*KT 52-05-07 หาค่า Vat ของ Item/หน่วยก่อนแล้วค่อยคูณ ๐ty ระบบเก่าใช้แบบนี้
'                                                    cVatAmt = cVatAmt * cQty '*KT 52-05-07 หาค่า Vat ของ Item/หน่วยก่อนแล้วค่อยคูณ ๐ty ระบบเก่าใช้แบบนี้
''                                                    cVatAmt = SP_FMTcCurPoint((cVatRate * (cQty * cTaxExceptAmt)) / (100 + cVatRate), True, False) 'หา vat จากราคาหลัง pro *KT 52-04-07 ทำการคำนวน Vat ใหม่ของรายการ Tax Except Sale
''                                                    cVatAmt = SP_FMTcCurPoint((cVatRate * (cB4Dis)) / (100 + cVatRate), True, False) 'หา vat จากราคาหลัง pro '*KT 52-01-24 คำนวน Vat หลังหักส่วนลดจากค่า Net Amount ก่อนหักส่วนลดใดๆๆ
'                                            Else  ' none vat
'                                                    cVatAmt = 0
'                                            End If
                                            If cVatRate <> 0 Then
                                                cVatAmt = SP_FMTcCurPoint((cVatRate * (cTaxExceptAmt)) / (100 + cVatRate), True, False)
                                                cVatAmt = cVatAmt * cQty
                                            Else
                                                cVatAmt = 0
                                            End If
                                            '--------------------------------------------
                                                    tTmp = Format(cVatAmt, tW_CurFmt)
                                                    Call oPrn.PrintText(tTkVat & " -" & tTmp, cXL, , vbLeftJustify, False)  '
                                                    tTmp = Format((cQty * cTaxExceptAmt) - cVatAmt, tW_CurFmt) '*KT 52-01-24 คำนวน Net Amount หลังหัก Vat
'                                                     tTmp = Format(cNet, tW_CurFmt) '*KT 52-01-24 คำนวน Net Amount หลังหัก Vat  *KT 52-04-07 คำนวนยอดสุทธิที่ต้องจ่ายหลังหัก vat
                                                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) '  Amt,
                                                    tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                    Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                    End If
                                                                        
                                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                                    'section4 duty free deposit
                                    If Val(tW_TransTypeSub) = 13 Then
                                                    tTmp = Format(cDutyFreeDeposit, tW_CurFmt)
                                                    Call oPrn.PrintText(tTkDepAmt & tTmp, cXL, , vbLeftJustify, False) '
                                                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' DutyFreeDepAmt,
                                                    
                                                    tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                    Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                    End If
                                    
                                    'section deal discount
                                    If orsCD.State = 1 Then
                                            '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                            'orsTmp.Filter = "FNSdtSeqNo =  " & nSeq & " AND FNDctNo = '1' AND FTScdDiscType <> '4'"
                                            orsTmp.Filter = "FNSdtSeqNo =  " & nSeq & " AND FNDctNo = '1'"
                                            orsTmp.Sort = "FNScdSeqNo ASC"  'เรียงตามการเกิด Discount
                                            cSumDiscOnLn = 0 '*KT 51-10-01  แก้ไข net amount
                                            Do While Not orsTmp.EOF
                                                    nTmp = SP_FEDvChkNumeric(orsTmp, "FNDctNo")
                                                    Select Case nTmp
                                                    Case nEN_Disc01Deal '- - - - - - - - - - -
                                                            '*KT 51-08-13 Communication Sheet ML-SF-443 Deal print in receipt wrong.xls
                                                            '-------------------------------------------------------------------------------------------------------------------------------
                                                            tDealNum = SP_FEDtChkString(orsTmp, "FTDctRefCode")
                                                            nDealGroup = SP_FEDvChkNumeric(orsTmp, "FNScdDealGroup")
                                                            'check last DT in deal group
                                                            '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                                            'orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FTScdDiscType<>'4' AND FNSdtSeqNo =  " & nSeq
                                                            orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FNSdtSeqNo =  " & nSeq
                                                            orsTmp.Sort = "FNScdSeqNo ASC"
                                                            If orsTmp.RecordCount > 0 Then
                                                                 '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                                                'orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FTScdDiscType<>'4'"
                                                                orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup
                                                                orsTmp.Sort = "FNScdSeqNo ASC"
                                                                '*KT 54-05-12 V.3.0.58 แก้ไขเรื่อง Free Type 002  เก็บ Record แรกของกลุ่ม FTDctRefCode,FNScdDealGroup
                                                                ''CommSheet ML-PRV-055 Not free item on Free Type2.xls
                                                                '*KT 54-09-02 V.4.0.1 แก้ไขเรื่อง Deal Type 003 ให้สามารถกำหนดกลุ่ม Get และ Buy มีสินค้าเดียวกัน
                                                                '*KT 54-09-23 V.4.0.1 แก้ไขเรื่อง Deal Type 003 ให้สามารถกำหนดกลุ่ม Get และ Buy มีสินค้าเดียวกัน ทั้งแบบ %off,amount off,new price
'                                                                If ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "002") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "4")) Or ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "003") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "1")) Then
                                                                If ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "002") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "4")) Or ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "003") And ((SP_FEDtChkString(orsTmp, "FTScdDiscType") = "1") Or (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "2") Or (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "3"))) Then
                                                                        orsTmp.Sort = "FCScdAmt ASC,FNScdSeqNo ASC"
                                                                End If
                                                                orsTmp.MoveLast
                                                                If SP_FEDvChkNumeric(orsTmp, "FNSdtSeqNo") = nSeq Then
                                                                            If SP_FEDvChkNumeric(orsTmp, "FCScdAmt") > 0 Then
                                                                                 '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                                                                If SP_FEDtChkString(orsTmp, "FTScdDiscType") = "4" Then
                                                                                    'Call oPrn.PrintText(tPdtDptCut, cXL, , vbLeftJustify, False) 'deprat '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ให้ Free แสดงเหมือนกับ Deal

'*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
'                                                                                    If tHaveDealPromo = "1" Then '*KT 52-12-02 Deal&Promo
'                                                                                         tTmp = cQty & "@" & Format(cRegPrice - (cHasPromo / cQty), tW_CurFmt)
'                                                                                         tLen = tTmp
'                                                                                        Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price
'                                                                                         tTmp = "Free"
'                                                                                    Else
                                                                                        tTmp = cQty & "@" & Format(cRegPrice - (cHasPromo / cQty), tW_CurFmt)
                                                                                         tLen = tTmp
'                                                                                        Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price   '*Eaw 56-02-26  TK-ISS3000-068
                                                                                        Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                                        tTmp = "Free"
'                                                                                    End If

                                                                                    tLen = tLen & tTmp
                                                                                    If Len(tLen) > 25 Then  'จำนวน SubTotal เกิน 12 Digit ขึ้นไปให้ขึ้นบรรทัดใหม่
                                                                                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' ขึ้นบรรทัดใหม่
                                                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                                                        tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                                                        Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                                                                    Else
                                                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                                                        tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                                                        Call oPrn.PrintText(vbNullString, cXMK, , vbLeftJustify, True) '  sign
                                                                                    End If
                                                                                    bPntDealEnd = True
                                                                                Else
                                                                                    cTmp = W_GETcDealAmt(tDealNum, nDealGroup)
                                                                                    
                                                                                    '*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
                                                                                    '*KT 52-11-28 Deal&Promo
'                                                                                    If tHaveDealPromo = "1" Then
'                                                                                        cTmp = cTmp - W_PRCcTotalDealPromo(orsTmp, tDealNum, nDealGroup)
'                                                                                        '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
'                                                                                        'orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FTScdDiscType<>'4'"
'                                                                                        orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup
'                                                                                        orsTmp.Sort = "FNScdSeqNo ASC"
'                                                                                        orsTmp.MoveLast
'                                                                                    End If

                                                                                    Call oPrn.PrintText(tTkDsDeal, cXL, , vbLeftJustify, False) 'Deal
                                                                                    tTmp = Format((-1) * cTmp, tW_CurFmt)
                                                                                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                                                    tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                                                    Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                                                                    bPntDealEnd = True
                                                                                End If
                                                                            Else
                                                                                bPntDealEnd = True
                                                                            End If
                                                                Else
                                                                            If SP_FEDvChkNumeric(orsTmp, "FCScdAmt") > 0 And SP_FEDtChkString(orsTmp, "FTScdDiscType") = "4" Then   'กรณีที่ได้ Free ที่เป็นกลุ่ม Get มากกว่า1 record ในกลุ่มเดียวกัน
                                                                                orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FNSdtSeqNo =  " & nSeq & " AND FCScdAmt > 0"
                                                                                orsTmp.Sort = "FNScdSeqNo ASC"
                                                                                If orsTmp.RecordCount > 0 Then
                                                                                    'Call oPrn.PrintText(tPdtDptCut, cXL, , vbLeftJustify, False) 'deprat '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ให้ Free แสดงเหมือนกับ Deal

'*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls .ให้การพิมพ์ Ticket/EJ
'                                                                                    If tHaveDealPromo = "1" Then '*KT 52-12-02 Deal&Promo
'                                                                                         tTmp = cQty & "@" & Format(cRegPrice - (cHasPromo / cQty), tW_CurFmt)
'                                                                                         tLen = tTmp
'                                                                                        Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price
'                                                                                         tTmp = "Free"
'                                                                                    Else
                                                                                        tTmp = cQty & "@" & Format(cRegPrice - (cHasPromo / cQty), tW_CurFmt)
                                                                                         tLen = tTmp
'                                                                                        Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price   '*Eaw 56-02-26  TK-ISS3000-068
                                                                                        Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                                                        tTmp = "Free"
'                                                                                    End If

                                                                                    tLen = tLen & tTmp
                                                                                    If Len(tLen) > 25 Then  'จำนวน SubTotal เกิน 12 Digit ขึ้นไปให้ขึ้นบรรทัดใหม่
                                                                                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' ขึ้นบรรทัดใหม่
                                                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                                                        tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                                                        Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) '  sign
                                                                                    Else
                                                                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) ' net
                                                                                        tTmp = IIf(tTrMk = vbNullString, tTxMk, tTrMk)
                                                                                        Call oPrn.PrintText(vbNullString, cXMK, , vbLeftJustify, True) '  sign
                                                                                    End If
                                                                                    bPntDealEnd = False
                                                                                    orsTmp.Filter = adFilterNone
                                                                                    orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup
                                                                                    orsTmp.Sort = "FNScdSeqNo ASC"
                                                                                    orsTmp.MoveLast
                                                                                Else
                                                                                    orsTmp.Filter = adFilterNone
                                                                                    orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup
                                                                                    orsTmp.Sort = "FNScdSeqNo ASC"
                                                                                    orsTmp.MoveLast
                                                                                    bPntDealEnd = False
                                                                                End If
                                                                            Else
                                                                                bPntDealEnd = False
                                                                            End If
                                                                End If
                                                                aW_CrossDeal(UBound(aW_CrossDeal)) = nSeq
                                                                ReDim Preserve aW_CrossDeal(UBound(aW_CrossDeal) + 1)
                                                                bPntDeal = True
                                                                If Not bPntDealST Then
                                                                    'Set orsNew1 = orsTmp.Clone
                                                                    Call SP_RSTbClone2OfflineEx(orsTmp, orsNew1, True)
                                                                    orsTmp.MoveLast
                                                                    If Not bIsKeep Then
                                                                        nOldPosition = orsTmp2.AbsolutePosition
                                                                    End If
                                                                    bPntDealST = True
                                                                End If
                                                            Else
                                                                'none
                                                            End If
                                                            tTmp = vbNullString
                                                            '-------------------------------------------------------------------------------------------------------------------------------
                                                    Case Else '- - - - - - - - - - -
                                                            tTmp = vbNullString
                                                    End Select
                                                    If tTmp <> vbNullString Then
                                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False) 'disc name
                                                            cDiscVal = SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
                                                            cSumDiscOnLn = cSumDiscOnLn + cDiscVal '*KT 51-10-01  แก้ไข net amount
                                                            Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXHL, , vbLeftJustify, False) 'dis amt
                                                            'cDiscVal = cNet - cDiscVal 'SP_FEDvChkNumeric(orsTmp, "FCScdAmtAF") '*KT 51-09-30 แก้ไข net amount
                                                            If Val(tW_TransTypeSub) = 11 Then '*KT 52-01-24 กรณีที่เป็นรายการ TaxExcept นัน้ต้องลบ Vat ด้วย
                                                                cDiscVal = cNetBDiscAll - cSumDiscOnLn - cVatAmt '*KT 51-10-01  แก้ไข net amount
                                                            Else
                                                                cDiscVal = cNetBDiscAll - cSumDiscOnLn  '*KT 51-10-01  แก้ไข net amount
                                                            End If
                                                            Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXMK, , vbRightJustify, False) 'disc amt after
                                                            tTmp = IIf(tTrMk = vbNullString, "", tTrMk)
                                                            Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) 'sign
                                                    End If
                                                    orsTmp.MoveNext
                                            Loop
                                    End If
                                        
                                    'section 5 line discount
                                    If orsCD.State = 1 Then
                                                orsTmp.Filter = "FNSdtSeqNo = " & nSeq     '*Ao 2013/03/27    TK"
'                                                orsTmp.Filter = "FNSdtSeqNo =  " & nSeq & " AND FNDctNo = 1 OR FNDctNo = 2 OR FNDctNo = 3 OR FNDctNo = 4 OR FNDctNo = 5 OR FNDctNo = 6 OR FNDctNo = 9 OR FNDctNo = 10 OR FNDctNo = 22"
                                                orsTmp.Sort = "FNScdSeqNo ASC"  'เรียงตามการเกิด Discount
                                                cSumDiscOnLn = 0 '*KT 51-10-01  แก้ไข net amount
                                                tTmp = ""
                                                nRecordCount = W_CHKbDisCount(nSeq)                  '*Ao 2013/03/21  CommSheet TK-ISS3000-116 00.00.01-Ao.xlsx
                                                nCount = 0
                                                Do While Not orsTmp.EOF
                                                        nTmp = SP_FEDvChkNumeric(orsTmp, "FNDctNo")
                                                        nCount = nCount + 1
                                                        Select Case nTmp
                                                        Case nEN_Disc01Deal
                                                                cSumDiscOnLn = cSumDiscOnLn + SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
                                                                 tTmp = vbNullString
                                                        '*KT 51-12-11 Phase 2 ข้อ 12
                                                        '-------------------------------------------------------------------------------------------------------------------------------------
                                                        Case nEN_Disc02PmtTime
                                                                cSumDiscOnLn = cSumDiscOnLn + SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
                                                                 tTmp = vbNullString
                                                        Case nEN_Disc03Pmt
                                                                cSumDiscOnLn = cSumDiscOnLn + SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
                                                                 tTmp = vbNullString
                                                        '--------------------------------------------------------------------------------------------------------------------------------------
                                                        Case nEN_Disc04CoopDisc '- - - - - - - - - - -
'                                                                tTmp = tTkDsCoDis   '*Eaw 55-12-18-19
                                                                cDiscPer = SP_FEDvChkNumeric(orsTmp, "FCScdDiscPer")    '*Eaw 55-12-18-19
                                                                tTmp = cDiscPer & tTkDsCoDis 'Per      '*Eaw 55-12-18-19
                                                        Case nEN_Disc05LnDisc '- - - - - - - - - - -
                                                                If SP_FEDvChkNumeric(orsTmp, "FTScdDiscType") = "2" Then
                                                                        tTmp = tTkDsDisAmt
                                                                ElseIf SP_FEDtChkString(orsTmp, "FTScdDiscType") = "1" Then
                                                                        cDiscPer = SP_FEDvChkNumeric(orsTmp, "FCScdDiscPer")    '*Ao 2012/07/27 PRJ10001TK-External Specification Line Discount-00.01.00.doc
                                                                        tTmp = cDiscPer & tTkDsDis 'Per
                                                                Else
                                                                         tTmp = tTkDsDisNew 'New Price
                                                                End If
                                                        Case nEN_Disc06LnMDown '- - - - - - - - - - -
                                                                If SP_FEDvChkNumeric(orsTmp, "FTScdDiscType") = "2" Then
                                                                        tTmp = tTkDsMkDnAmt
                                                                ElseIf SP_FEDtChkString(orsTmp, "FTScdDiscType") = "1" Then
                                                                        cDiscPer = SP_FEDvChkNumeric(orsTmp, "FCScdDiscPer")    '*Ao 2012/07/27 PRJ10001TK-External Specification Line Discount-00.01.00.doc
                                                                        tTmp = cDiscPer & tTkDsMkDn 'Per
                                                                Else
                                                                         tTmp = tTkDsMkDnNew 'New Price
                                                                End If

                                                        Case nEN_Disc09CoopCpn '- - - - - - - - - - -
                                                                tTmp = tTkDsCoCpn

                                                        Case nEN_Disc10CoopBrand '- - - - - - - - - - -
                                                                tTmp = tTkDsCoBnd

                                                        Case nEN_Disc22MDisc '*KT 52-09-08 Member Price + Consignment
                                                                tTmp = tTkDsMPDis

                                                        Case Else '- - - - - - - - - - -
                                                                tTmp = vbNullString

                                                        End Select

                                                        If tTmp <> vbNullString Then
                                                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False) 'disc name
                                                                cDiscVal = SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
'                                                                cSumDiscOnLn = cSumDiscOnLn + cDiscVal '*KT 51-10-01  แก้ไข net amount
'                                                                Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXHL, , vbLeftJustify, False) 'dis amt
'                                                                'cDiscVal = cNet - cDiscVal 'SP_FEDvChkNumeric(orsTmp, "FCScdAmtAF") '*KT 51-09-30 แก้ไข net amount
'                                                                If Val(tW_TransTypeSub) = 11 Then '*KT 52-01-24 กรณีที่เป็นรายการ TaxExcept นัน้ต้องลบ Vat ด้วย
'                                                                    If bHasMPrice Then
'                                                                        cDiscVal = (cMPrice * cQty) - cSumDiscOnLn - cVatAmt
'                                                                    Else
'                                                                        'cDiscVal = cNetBDiscAll - cSumDiscOnLn - cVatAmt  '*KT 51-10-01  แก้ไข net amount
'                                                                        cDiscVal = (cNetBDiscAll - (cNetBDiscAll * cVatRate / (100 + cVatRate))) - cSumDiscOnLn '*KT 51-10-01  แก้ไข net amount
'                                                                    End If
'                                                                Else
'                                                                     If bHasMPrice Then
'                                                                        cDiscVal = (cMPrice * cQty) - cSumDiscOnLn
'                                                                    Else
'                                                                        cDiscVal = cNetBDiscAll - cSumDiscOnLn  '*KT 51-10-01  แก้ไข net amount
'                                                                    End If
'
'                                                                End If
'                                                                Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXMK, , vbRightJustify, False) 'disc amt after
'                                                                '-----------------------------------------------------------------
''                                                                tTmp = IIf(tTrMk = vbNullString, "", tTrMk)
'                                                                tTmp = IIf(tTrMk = vbNullString, tTxMk, "")
'                                                                Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) 'sign               '*Ao 2013/02/27  TK CommSheet TK-ISS3000-045 00.00.02 FSBT NG.xlsx
                                                                Select Case nTmp
                                                                    Case nEN_Disc05LnDisc, nEN_Disc06LnMDown
                                                                        If SP_FEDtChkString(orsTmp, "FTScdDiscType") = "3" Then
                                                                            cDiscVal = SP_FEDvChkNumeric(orsTmp, "FCScdDiscNew")    '*Ao 2013/03/21  CommSheet TK-ISS3000-116 00.00.01-Ao.xlsx
                                                                            Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXHL, , vbLeftJustify, False) 'dis amt
                                                                            cDiscVal = SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
                                                                            cSumDiscOnLn = cSumDiscOnLn + cDiscVal '*KT 51-10-01  แก้ไข net amount
                                                                        Else
                                                                            cSumDiscOnLn = cSumDiscOnLn + cDiscVal '*KT 51-10-01  แก้ไข net amount
                                                                            Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXHL, , vbLeftJustify, False) 'dis amt
                                                                        End If
                                                                    Case Else
                                                                        cSumDiscOnLn = cSumDiscOnLn + cDiscVal '*KT 51-10-01  แก้ไข net amount
                                                                        Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXHL, , vbLeftJustify, False) 'dis amt
                                                                End Select
'                                                                cSumDiscOnLn = cSumDiscOnLn + cDiscVal '*KT 51-10-01  แก้ไข net amount
'                                                                Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXHL, , vbLeftJustify, False) 'dis amt
                                                                'cDiscVal = cNet - cDiscVal 'SP_FEDvChkNumeric(orsTmp, "FCScdAmtAF") '*KT 51-09-30 แก้ไข net amount
                                                                If Val(tW_TransTypeSub) = 11 Then '*KT 52-01-24 กรณีที่เป็นรายการ TaxExcept นัน้ต้องลบ Vat ด้วย
                                                                    If bHasMPrice Then
                                                                        cDiscVal = (cMPrice * cQty) - cSumDiscOnLn - cVatAmt
                                                                    Else
                                                                        'cDiscVal = cNetBDiscAll - cSumDiscOnLn - cVatAmt  '*KT 51-10-01  แก้ไข net amount
                                                                        cDiscVal = (cNetBDiscAll - (cNetBDiscAll * cVatRate / (100 + cVatRate))) - cSumDiscOnLn '*KT 51-10-01  แก้ไข net amount
                                                                    End If
                                                                Else
                                                                    If bHasMPrice Then
                                                                        cDiscVal = (cMPrice * cQty) - cSumDiscOnLn
                                                                    Else
'                                                                        cDiscVal = cNetBDiscAll - cSumDiscOnLn  '*KT 51-10-01  แก้ไข net amount
                                                                        '--------------------------------------------------------------------
                                                                        Select Case nTmp
                                                                            Case nEN_Disc05LnDisc, nEN_Disc06LnMDown
                                                                                cDiscVal = SP_FEDvChkNumeric(orsTmp, "FCScdAmtAF")          '*Ao 2013/03/21  CommSheet TK-ISS3000-116 00.00.01-Ao.xlsx
                                                                            Case Else
                                                                                cDiscVal = cNetBDiscAll - cSumDiscOnLn  '*KT 51-10-01  แก้ไข net amount
                                                                        End Select
                                                                    '------------------------------------------------------------------
                                                                    End If
                                                                End If
                                                                
                                                                If nRecordCount = nCount Then                       '*Ao 2013/03/21  CommSheet TK-ISS3000-116 00.00.01-Ao.xlsx
                                                                    Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXMK, , vbRightJustify, False) 'disc amt after
                                                                    tTmp = IIf(tTrMk = vbNullString, tTxMk, "")
                                                                    Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) 'sign
                                                                Else
                                                                    tTmp = ""
                                                                    Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) 'sign
                                                                End If
'                                                                Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXMK, , vbRightJustify, False) 'disc amt after
'                                                                    '-----------------------------------------------------------------
'    '                                                                tTmp = IIf(tTrMk = vbNullString, "", tTrMk)
'                                                                tTmp = IIf(tTrMk = vbNullString, tTxMk, "")
'                                                                Call oPrn.PrintText(tTmp, cXMK, , vbLeftJustify, True) 'sign               '*Ao 2013/02/27  TK CommSheet TK-ISS3000-045 00.00.02 FSBT NG.xlsx
   '-------------------------------------------------------------------------------------------------------------------------
                                                        End If
                                                        orsTmp.MoveNext
                                                Loop
                                        End If
                                            
                                        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                                End If
                                
                                '*KT 53-10-16 Vendor Coupon ที่ผูกกับสินค้าตามเอกสาร External Spec PRJ09010TheMall-R3-SE-External Specification Vendor Coupon-00.06.00.doc
                                '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                'tTkLnVendorCpn
                                'Step 1 : กรองหา RC ที่เป็น Line Vendor Cpn มีหรือไม่ ด้วยเงื่อนไข Index,Tender Type,Flag Line Vendor Cpn
                                'Step 2 : ถ้ามีให้แสดงข้อความตาม External Spec
                                orsRC.Filter = adFilterNone
                                orsRC.Filter = "FTSkuIndex = '" & tIndexLnVendorCpn & "' AND FTTdmCode = 'T006' AND FTVendorCnpFlag = '1'"
                                orsRC.Sort = "FNSrcSeqNo ASC"
                                If orsRC.RecordCount > 0 Then
                                        Do While Not orsRC.EOF
                                            tQtyLnVendorCpn = SP_FEDtChkString(orsRC, "FTVendorCnpQty")
                                            If tQtyLnVendorCpn <> "" Then
                                                nQtyLnVendorCpn = CDbl(tQtyLnVendorCpn)
                                                If nQtyLnVendorCpn > 0 Then
                                                    'Print tTkLnVendorCpn
                                                    Call oPrn.PrintText(Mid(tTkLnVendorCpn, 1, 7), cXL, , vbLeftJustify, False) 'V-CPN
                                                    'Print @Qty
                                                    cLnVendorCpnNet = SP_FEDvChkNumeric(orsRC, "FCSrcAmt") 'ที่ไม่ใช้ FCSrcNet เพราะกรณี vender Cpn มีการทอน
                                                    cLnVendorCpnPerQty = cLnVendorCpnNet / nQtyLnVendorCpn
                                                    tTmp = tQtyLnVendorCpn & "@" & Format(cLnVendorCpnPerQty, tW_CurFmt)
'                                                    Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX8, cX8x), , vbLeftJustify, False) 'qty *price  '*Eaw 56-02-26  TK-ISS3000-068
                                                    Call oPrn.PrintText(tTmp, IIf(bC_PrnByOPOS, cX15, cXC), , vbLeftJustify, False) 'qty@ price
                                                    'Print TOTAL
                                                     Call oPrn.PrintText(Format(cLnVendorCpnNet, tW_CurFmt), IIf(bC_PrnByOPOS, cXC, cXC), , vbLeftJustify, True, , cWidthLimit)
                                                End If
                                            End If
                                            orsRC.MoveNext
                                        Loop
                                End If
                                orsRC.Filter = adFilterNone
                                '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                                
                                If bPntDeal Then
                                     '------------------------------------
                                    'Set orsNew2 = orsNew1.Clone
                                    Call SP_RSTbClone2OfflineEx(orsNew1, orsNew2, True)
                                    orsNew2.Filter = "FNSdtSeqNo <> " & nSeq
                                    orsNew2.Sort = "FNScdSeqNo ASC"
                                    '*KT 54-05-12 V.3.0.58 แก้ไขเรื่อง Free Type 002  เก็บ Record แรกของกลุ่ม FTDctRefCode,FNScdDealGroup
                                    ''CommSheet ML-PRV-055 Not free item on Free Type2.xls
                                    '*KT 54-09-02 V.4.0.1 แก้ไขเรื่อง Deal Type 003 ให้สามารถกำหนดกลุ่ม Get และ Buy มีสินค้าเดียวกัน
                                    '*KT 54-09-23 V.4.0.1 แก้ไขเรื่อง Deal Type 003 ให้สามารถกำหนดกลุ่ม Get และ Buy มีสินค้าเดียวกัน ทั้งแบบ %off,amount off,new price
'                                    If ((SP_FEDtChkString(orsNew2, "FTScdCardType") = "002") And (SP_FEDtChkString(orsNew2, "FTScdDiscType") = "4")) Or ((SP_FEDtChkString(orsNew2, "FTScdCardType") = "003") And (SP_FEDtChkString(orsNew2, "FTScdDiscType") = "1")) Then
                                    If ((SP_FEDtChkString(orsNew2, "FTScdCardType") = "002") And (SP_FEDtChkString(orsNew2, "FTScdDiscType") = "4")) Or ((SP_FEDtChkString(orsNew2, "FTScdCardType") = "003") And ((SP_FEDtChkString(orsNew2, "FTScdDiscType") = "1") Or (SP_FEDtChkString(orsNew2, "FTScdDiscType") = "2") Or (SP_FEDtChkString(orsNew2, "FTScdDiscType") = "3"))) Then
                                            orsNew2.Sort = "FCScdAmt ASC,FNScdSeqNo ASC"
                                    End If
                                    If Not (orsNew2.EOF And orsNew2.BOF) Then
                                         Call SP_RSTbClone2OfflineEx(orsNew2, orsNew1, True)
                                         orsNew2.MoveFirst
                                         orsTmp2.Filter = "FNSdtSeqNo =  " & SP_FEDvChkNumeric(orsNew2, "FNSdtSeqNo")
                                    End If
                                    If bPntDealEnd Then
                                         orsTmp2.Filter = adFilterNone
                                         orsTmp2.AbsolutePosition = nOldPosition
                                         bPntDeal = False
                                         bPntDealST = False
                                         bPntDealEnd = False
                                         bCross = True
                                         Do While bCross
                                             orsTmp2.MoveNext
                                             If Not orsTmp2.EOF Then
                                                 bCross = False
                                                 nSeq = SP_FEDvChkNumeric(orsTmp2, "FNSdtSeqNo")
                                                 For nI = 0 To UBound(aW_CrossDeal) - 1
                                                     If Val(aW_CrossDeal(nI)) = nSeq Then
                                                         bCross = True
                                                         Exit For
                                                     Else
                                                         bCross = False
                                                     End If
                                                 Next nI
                                             Else
                                                 bCross = False
                                             End If
                                         Loop
                                        'ตรวจสอบว่ารายการต่อไปที่จะพมพ์เป็น Deal หรือไม่
                                        'ถ้าเป็น Deal ต้องนำรายการแรกในกลุ่มที่ได้ Deal นั้นมาพิพพ์ก่อน และต้อง Mark รายการเดิมไว้เพื่อเมื่อพิมพ์รายการ Deal เสร็จจะกลับมาได้
                                         bIsKeep = False
                                         If Not orsTmp2.EOF Then
                                             orsTmp.Filter = "FNSdtSeqNo =  " & nSeq & " "
                                             orsTmp.Sort = "FNScdSeqNo ASC"  'เรียงตามการเกิด Discount
                                             tDealNum = SP_FEDtChkString(orsTmp, "FTDctRefCode")
                                             nDealGroup = SP_FEDvChkNumeric(orsTmp, "FNScdDealGroup")
                                             'check DT in deal group
                                             '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                             'orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FTScdDiscType<>'4' AND FNSdtSeqNo=" & nSeq & " AND  FNDctNo = " & nEN_Disc01Deal
                                             orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FNSdtSeqNo=" & nSeq & " AND  FNDctNo = " & nEN_Disc01Deal
                                             orsTmp.Sort = "FNScdSeqNo ASC"
                                             If orsTmp.RecordCount > 0 Then
                                                     '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                                   ' orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FTScdDiscType<>'4' AND  FNDctNo = " & nEN_Disc01Deal
                                                    orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & "  AND  FNDctNo = " & nEN_Disc01Deal
                                                    orsTmp.Sort = "FNScdSeqNo ASC"
                                                    '*KT 54-05-12 V.3.0.58 แก้ไขเรื่อง Free Type 002  เก็บ Record แรกของกลุ่ม FTDctRefCode,FNScdDealGroup
                                                    ''CommSheet ML-PRV-055 Not free item on Free Type2.xls
                                                    '*KT 54-09-02 V.4.0.1 แก้ไขเรื่อง Deal Type 003 ให้สามารถกำหนดกลุ่ม Get และ Buy มีสินค้าเดียวกัน
                                                    '*KT 54-09-23 V.4.0.1 แก้ไขเรื่อง Deal Type 003 ให้สามารถกำหนดกลุ่ม Get และ Buy มีสินค้าเดียวกัน ทั้งแบบ %off,amount off,new price
'                                                    If ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "002") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "4")) Or ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "003") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "1")) Then
                                                     If ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "002") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "4")) Or ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "003") And ((SP_FEDtChkString(orsTmp, "FTScdDiscType") = "1") Or (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "2") Or (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "3"))) Then
                                                            orsTmp.Sort = "FCScdAmt ASC,FNScdSeqNo ASC"
                                                    End If
                                                      nSeq = SP_FEDvChkNumeric(orsTmp, "FNSdtSeqNo")
                                                      nOldPosition = orsTmp2.AbsolutePosition
                                                      Do While Not orsTmp2.EOF
                                                         If nSeq = SP_FEDvChkNumeric(orsTmp2, "FNSdtSeqNo") Then
                                                             Exit Do
                                                         End If
                                                         orsTmp2.MoveNext
                                                      Loop
                                                      bIsKeep = True
                                             End If
                                         End If
                                    End If
                                    '------------------------------------
                              Else
    '                               orsPrintSeq.MoveNext
                                    bCross = True
                                    Do While bCross
                                        orsTmp2.MoveNext
                                        If Not orsTmp2.EOF Then
                                            bCross = False
                                            nSeq = SP_FEDvChkNumeric(orsTmp2, "FNSdtSeqNo")
                                            For nI = 0 To UBound(aW_CrossDeal) - 1
                                                If Val(aW_CrossDeal(nI)) = nSeq Then
                                                    bCross = True
                                                    Exit For
                                                Else
                                                    bCross = False
                                                End If
                                            Next nI
                                        Else
                                            bCross = False
                                        End If
                                    Loop
                                     'ตรวจสอบว่ารายการต่อไปที่จะพมพ์เป็น Deal หรือไม่
                                     'ถ้าเป็น Deal ต้องนำรายการแรกในกลุ่มที่ได้ Deal นั้นมาพิพพ์ก่อน และต้อง Mark รายการเดิมไว้เพื่อเมื่อพิมพ์รายการ Deal เสร็จจะกลับมาได้
                                    bIsKeep = False
                                    If Not orsTmp2.EOF Then
                                        orsTmp.Filter = "FNSdtSeqNo =  " & nSeq
                                        orsTmp.Sort = "FNScdSeqNo ASC"  'เรียงตามการเกิด Discount
                                        tDealNum = SP_FEDtChkString(orsTmp, "FTDctRefCode")
                                        nDealGroup = SP_FEDvChkNumeric(orsTmp, "FNScdDealGroup")
                                        'check DT in deal group
                                        '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                        'orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FTScdDiscType<>'4' AND FNSdtSeqNo=" & nSeq & " AND FNDctNo =" & nEN_Disc01Deal
                                        orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & "  AND FNSdtSeqNo=" & nSeq & " AND FNDctNo =" & nEN_Disc01Deal
                                        orsTmp.Sort = "FNScdSeqNo ASC"
                                        If orsTmp.RecordCount > 0 Then
                                                '*KT 52-12-02 PH 3.0-ML-SF-028,PH 3.0-ML-SF-029,PH 3.0-ML-SF-030,PH 3.0-ML-SF-031 ห้ Free แสดงเหมือนกับ Deal
                                                'orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FTScdDiscType<>'4'  AND  FNDctNo = " & nEN_Disc01Deal
                                                orsTmp.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & "   AND  FNDctNo = " & nEN_Disc01Deal
                                                orsTmp.Sort = "FNScdSeqNo ASC"
                                                '*KT 54-05-12 V.3.0.58 แก้ไขเรื่อง Free Type 002  เก็บ Record แรกของกลุ่ม FTDctRefCode,FNScdDealGroup
                                                ''CommSheet ML-PRV-055 Not free item on Free Type2.xls
                                                '*KT 54-09-02 V.4.0.1แก้ไขเรื่อง Deal Type 003 ให้สามารถกำหนดกลุ่ม Get และ Buy มีสินค้าเดียวกัน
                                               '*KT 54-09-23 V.4.0.1 แก้ไขเรื่อง Deal Type 003 ให้สามารถกำหนดกลุ่ม Get และ Buy มีสินค้าเดียวกัน ทั้งแบบ %off,amount off,new price
'                                                If ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "002") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "4")) Or ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "003") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "1")) Then
                                                If ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "002") And (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "4")) Or ((SP_FEDtChkString(orsTmp, "FTScdCardType") = "003") And ((SP_FEDtChkString(orsTmp, "FTScdDiscType") = "1") Or (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "2") Or (SP_FEDtChkString(orsTmp, "FTScdDiscType") = "3"))) Then
                                                        orsTmp.Sort = "FCScdAmt ASC,FNScdSeqNo ASC"
                                                End If
                                                 nSeq = SP_FEDvChkNumeric(orsTmp, "FNSdtSeqNo")
                                                 nOldPosition = orsTmp2.AbsolutePosition
                                                 Do While Not orsTmp2.EOF
                                                    If nSeq = SP_FEDvChkNumeric(orsTmp2, "FNSdtSeqNo") Then
                                                        Exit Do
                                                    End If
                                                    orsTmp2.MoveNext
                                                 Loop
                                                 bIsKeep = True
                                        End If
                                End If
                             End If
                        Loop
      '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
               Case 5, 14, 17, 26, 27, 28, 32, 33
                            
                                    Call oPrn.PrintText(tTkPost, cXL, , vbLeftJustify, False)  'post#
                                    Call oPrn.PrintText(tTkTran, cXC - cXHC, , vbLeftJustify, False) 'tran#
                                    Call oPrn.PrintText(tTkAmt, cXHR, , vbLeftJustify, True) ' amt

                                    tTmp = SP_FEDtChkString(orsHD, "FTShdDepRefTmnNum")
                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)  'post#
                                    
                                    tTmp = SP_FEDtChkString(orsHD, "FTShdDepReTransNo")
                                    Call oPrn.PrintText(tTmp, cXC - cXHC, , vbLeftJustify, False) 'tran#
                                     Select Case Val(tW_TransTypeSub)
                                        Case 17
                                            tTmp = Format(SP_FEDvChkNumeric(orsHD, "FCShdDepositAmt") + SP_FEDvChkNumeric(orsHD, "FCShdRnd"), tW_CurFmt) '*KT 51-11-10 ML-SB-634
                                        Case Else
                                            tTmp = Format(SP_FEDvChkNumeric(orsHD, "FCShdGrand") + SP_FEDvChkNumeric(orsHD, "FCShdRnd"), tW_CurFmt) '*KT 51-11-10 ML-SB-634
                                    End Select
                                    'Call oPrn.PrintText(tTmp, cXHR, , vbLeftJustify, True) ' amt
                                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True) ' amt
                                    
                                    Call oPrn.PrintText(tTkOrgTotalUpd, cXL, , vbLeftJustify, True)  ' OrgTotalTpd
                                    
                                    Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
                                                                 
                                    Call W_GETxDataToReprint
     
      '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
               Case 9, 31, 40, 42, 45, 47 'close preport
                    If Not orsDT Is Nothing Then
                        If orsDT.RecordCount > 0 Then
                            
                            Dim tForeign As String
'                            tSql = "SELECT FTFcrCode,FTFcrSign,FTFcrType,FCFcrExchang FROM TCNMForCurrency"   '*Eaw 56-03-19 CommSheet TK-ISS3000-083 00.00.01.xlsx
                            tSql = "SELECT LTRIM(RTRIM(FTFcrCode)) AS FTFcrCode,FTFcrSign,FTFcrType,FCFcrExchang FROM TCNMForCurrency"   '*BG 56-09-09  Comsheet-285
'                            Call SP_SQLvExecute(tSql, orsFor)
                            Call SP_SQLvExecute(tSql, orsFor, oDB.ocnOnLine, , , , , True, True) 'read TCNMForCurrency          '*Ao 2013/03/01  TK
                            orsDT.MoveFirst
                            With orsDT
                                Do While Not .EOF
                                            'Line_Chars
                                            '*KT 51-08-11 Communication Sheet ML-SF-462 Wrong display ABB Thai on SKU description.xls
                                            tPdtName = Mid(SP_FEDtChkString(orsDT, "FTSkuAbbName"), 1, 17) 'Default SKU Name Abb Eng
                                            tPdtName2 = Mid(SP_FEDtChkString(orsDT, "FTSkuAbbNameSnd"), 1, 17) 'Default SKU Name  Abb Thai
                                            cQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
                                            cNet = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
                                            Select Case Val(tW_TransTypeSub) '
                                                Case 9, 42, 47 'Closing Report
                                                    If tPdtName = "" Then
                                                            If tTmpBlank <> "" Then
                                                                Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
                                                                tTmpBlank = tPdtName
                                                            End If
                                                    Else
                                                            If cNet <> 0 Then
                                                                tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                If tTmp <> vbNullString Then
                                                                     Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                End If
                                                                If UCase(tPdtName2) = UCase("Start Bank") Then
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                    Call oPrn.PrintText("Cash", cX5, , vbLeftJustify, False)
                                                                    Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                ElseIf UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T012") Then
                                                                    orsFor.Filter = "FTFcrCode='" & Trim(tPdtName) & "'"
                                                                    If Not (orsFor.BOF And orsFor.EOF) Then
                                                                        tForeign = SP_FEDtChkString(orsFor, "FTFcrSign")
                                                                        tForeign = Trim(tForeign)   '*Eaw 56-05-27 CommSheet TK-ISS3000-108 00.00.05 FSBT.xls
                                                                        tPdtName2 = Trim(SP_FEDtChkString(orsFor, "FTFcrType"))               '*Ao 2013/02/28  CommSheet TK-ISS3000-083 00.00.01.xlsx
                                                                        tForFac = SP_FEDtChkString(orsDT, "FTSdtStkCode") '*Eaw 56-04-09 CommSheet TK-ISS3000-106 00.00.03.xlsx
                                                                        If tForFac = "1" Then
                                                                            cNet = SP_FEDvChkNumeric(orsDT, "FCSdtStkFac")
                                                                            tPdtName2 = Space$(1) & tPdtName2
                                                                            Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)             '*Ao 2014/06/27 เพิ่ม ELSE
                                                                            Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtDutyFreeAmt"), "#,##0.00000"), cXC, , vbLeftJustify, False)
                                                                            Call oPrn.PrintText(tForeign & Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                        Else
                                                                            Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)             '*Ao 2014/06/27 เพิ่ม ELSE
                                                                            Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtStkFac"), "#,##0.00000"), cXC, , vbLeftJustify, False)  '*Ao 2014/06/27
                                                                            Call oPrn.PrintText(tForeign & Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                        End If
'                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
'                                                                        Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsFor, "FCFcrExchang"), "#,##0.00000"), cXC, , vbLeftJustify, False)  '*Eaw 56-03-19 CommSheet TK-ISS3000-083 00.00.01.xlsx
'                                                                        Call oPrn.PrintText(tForeign & Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                    Else
                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                        Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)       '*Ao 2013/05/31 Foreign(Base)
                                                                    End If
                                                                    orsFor.Filter = adFilterNone
                                                                ElseIf UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T002") Or UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T003") Or UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T017") Or UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T015") Then
                                                                    If cQty = 0 Then
                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                    Else
                                                                        Call oPrn.PrintText(Space$(1) & tPdtName2, cX5, , vbLeftJustify, False)
                                                                        Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                    End If      '*Ao 2013/06/03  CommSheet TK-ISS3000-222-ST 00 00 01-Aoo.xls
'                                                                ElseIf UCase(tPdtName2) = UCase("OFF-LINE") Or UCase(tPdtName2) = UCase("VOID") Or UCase(tPdtName2) = UCase("Drawer Open") Or UCase(tPdtName2) = UCase("CV Issue") Or UCase(tPdtName2) = UCase("CV Cancel") Or UCase(tPdtName2) = UCase("CV Pickup") Or UCase(tPdtName2) = UCase("TranSuspendIn") Or UCase(tPdtName2) = UCase("Total IPV In") Or UCase(tPdtName2) = UCase("Total IPV Out") Then
                                                                ElseIf UCase(tPdtName2) = UCase("OFF-LINE") Or UCase(tPdtName2) = UCase("VOID") Or UCase(tPdtName2) = UCase("Drawer Open") Or UCase(tPdtName2) = UCase("CV Cancel") Or UCase(tPdtName2) = UCase("TranSuspendIn") Or UCase(tPdtName2) = UCase("Total IPV In") Or UCase(tPdtName2) = UCase("Total IPV Out") Then
                                                                    If cQty > 0 Then
                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                    End If
                                                                ElseIf UCase(tPdtName) = UCase("Cash Transfer") Then  '*Eaw 56-04-09 CommSheet TK-ISS3000-106 00.00.03.xlsx
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                ElseIf UCase(tPdtName) = UCase("Total Return") Then  '*Ao 2013/08/07       -Total Return
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                    Call oPrn.PrintText(Format("-" & cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                Else
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                    Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                End If
                                                                tTmpBlank = tPdtName
                                                            Else
                                                                If UCase(tPdtName2) = UCase("Total Sale") Then
                                                                         tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                        If tTmp <> vbNullString Then
                                                                             Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                        End If
                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                        Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                        tTmpBlank = tPdtName
                                                                End If
                                                                If UCase(tPdtName2) = UCase("Start Bank") Then
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                    Call oPrn.PrintText(Space$(1) & "Cash", cX5, , vbLeftJustify, False)
                                                                    Call oPrn.PrintText(Format(0, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                End If
                                                                '*Eaw 56-08-02 CommSheet TK-ISS3000-299-UAT.xls
                                                                Select Case UCase(tPdtName2) '*Eaw 56-08-29 เปลี่ยน
                                                                    Case UCase("Layaway Deposit"), UCase("Rounding")
                                                                        If cQty > 0 Then
                                                                            tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                            Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                            Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                            Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                            tTmpBlank = tPdtName
                                                                        End If
                                                                End Select
                                                                Select Case UCase(tPdtName2)
'                                                                    Case UCase("Drawer Open"), UCase("CV Issue"), UCase("CV Cancel"), UCase("CV Pickup"), UCase("TranSuspendIn"), UCase("Total IPV In"), UCase("Total IPV Out")
                                                                    Case UCase("Drawer Open"), UCase("CV Cancel"), UCase("CV Pickup"), UCase("TranSuspendIn"), UCase("Total IPV In"), UCase("Total IPV Out"), UCase("VOID") '*Eaw 56-07-02 CommSheet TK-ISS3000-243-TR.xls เพิ่ม VOID
                                                                        If cQty > 0 Then
                                                                            tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                            If tTmp <> vbNullString Then
                                                                                 Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                            End If
                                                                            Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                        End If
                                                                End Select
                                                                '*Eaw 56-09-12 CommSheet TK-ISS3000-351-UAT-FAPL-119.xls
'                                                                Select Case UCase(SP_FEDtChkString(orsDT, "FTSkuCode"))
'                                                                    Case UCase("T001"), UCase("T002"), UCase("T003"), UCase("T004"), UCase("T005"), UCase("T006"), UCase("T007"), UCase("T008"), UCase("T009"), UCase("T010"), UCase("T011"), UCase("T012"), UCase("T013"), UCase("T014"), UCase("T015"), UCase("T016"), UCase("T017"), UCase("T018")
'                                                                    tTmp = IIf(cQty = 0, vbNullString, cQty)
'                                                                    If tTmp <> vbNullString Then
'                                                                        Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
'                                                                    End If
'                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
'                                                                    Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
'                                                                End Select
                                                                '*Ao 2014/07/31   CommSheet TK-ISS3000-460-00 00 03   V4.1.129
                                                                Select Case UCase(SP_FEDtChkString(orsDT, "FTSkuCode"))
                                                                    Case UCase("T001"), UCase("T004"), UCase("T005"), UCase("T006"), UCase("T007"), UCase("T008"), UCase("T009"), UCase("T010"), UCase("T011"), UCase("T012"), UCase("T013"), UCase("T014"), UCase("T016"), UCase("T018")
                                                                        tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                        If tTmp <> vbNullString Then
                                                                            Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                        End If
                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                        Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                    Case UCase("T002"), UCase("T003"), UCase("T015"), UCase("T017")         '*Ao 2014/07/31   CommSheet TK-ISS3000-460-00 00 03   V4.1.129
                                                                        If cQty = 0 Then
                                                                            Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                        Else
                                                                            tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                            If tTmp <> vbNullString Then
                                                                                Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                            End If
                                                                            Call oPrn.PrintText(Space$(1) & tPdtName2, cX5, , vbLeftJustify, False)
                                                                            Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                        End If
                                                                End Select
                                                            End If
                                                    End If
                                            Case 31, 40, 45 'Total Report
                                                    If tPdtName = "" Then
                                                            If tTmpBlank <> "" Then
                                                                Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
                                                                 tTmpBlank = tPdtName
                                                           End If
                                                    Else
                                                            If cNet <> 0 Then
                                                                tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                If tTmp <> vbNullString Then
                                                                    Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                End If
                                                                If UCase(tPdtName2) = UCase("Start Bank") Then
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                    Call oPrn.PrintText(Space$(1) & "Cash", cX5, , vbLeftJustify, False)
                                                                    Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                ElseIf UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T012") Then
                                                                    orsFor.Filter = "FTFcrCode='" & Trim(tPdtName) & "'"
                                                                    If Not (orsFor.BOF And orsFor.EOF) Then
                                                                        tForeign = SP_FEDtChkString(orsFor, "FTFcrSign")
                                                                        tForeign = Trim(tForeign)               '*Eaw 56-05-27 CommSheet TK-ISS3000-108 00.00.05 FSBT.xls
                                                                        tPdtName2 = Trim(SP_FEDtChkString(orsFor, "FTFcrType"))               '*Ao 2013/02/28  CommSheet TK-ISS3000-083 00.00.01.xlsx
                                                                        tForFac = SP_FEDtChkString(orsDT, "FTSdtStkCode") '*Eaw 56-04-09 CommSheet TK-ISS3000-106 00.00.03.xlsx
                                                                        If tForFac = "1" Then
                                                                            cNet = SP_FEDvChkNumeric(orsDT, "FCSdtStkFac")
                                                                            tPdtName2 = Space$(1) & tPdtName2
                                                                            Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                            Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtDutyFreeAmt"), "#,##0.00000"), cXC, , vbLeftJustify, False)
                                                                            Call oPrn.PrintText(tForeign & Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                        Else
                                                                            Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)             '*Ao 2014/06/27   Add ELSE
                                                                            Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtStkFac"), "#,##0.00000"), cXC, , vbLeftJustify, False)  '*Eaw 56-03-19 CommSheet TK-ISS3000-083 00.00.01.xlsx
                                                                            Call oPrn.PrintText(tForeign & Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                        End If
'                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
'                                                                        Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsFor, "FCFcrExchang"), "#,##0.00000"), cXC, , vbLeftJustify, False)  '*Eaw 56-03-19 CommSheet TK-ISS3000-083 00.00.01.xlsx
'                                                                        Call oPrn.PrintText(tForeign & Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                    Else
                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                        Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)       '*Ao 2013/05/31 Foreign(Base)
                                                                    End If
                                                                    orsFor.Filter = adFilterNone
                                                                ElseIf UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T002") Or UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T003") Or UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T017") Or UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T015") Then
                                                                    If cQty = 0 Then
                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                    Else
                                                                        Call oPrn.PrintText(Space$(1) & tPdtName2, cX5, , vbLeftJustify, False)
                                                                        Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                    End If      '*Ao 2013/06/03  CommSheet TK-ISS3000-222-ST 00 00 01-Aoo.xls
'                                                                ElseIf UCase(tPdtName2) = UCase("OFF-LINE") Or UCase(tPdtName2) = UCase("VOID") Or UCase(tPdtName2) = UCase("Drawer Open") Or UCase(tPdtName2) = UCase("CV Issue") Or UCase(tPdtName2) = UCase("CV Cancel") Or UCase(tPdtName2) = UCase("CV Pickup") Or UCase(tPdtName2) = UCase("TranSuspendIn") Or UCase(tPdtName2) = UCase("Total IPV In") Or UCase(tPdtName2) = UCase("Total IPV Out") Then
                                                                ElseIf UCase(tPdtName2) = UCase("OFF-LINE") Or UCase(tPdtName2) = UCase("VOID") Or UCase(tPdtName2) = UCase("Drawer Open") Or UCase(tPdtName2) = UCase("CV Cancel") Or UCase(tPdtName2) = UCase("TranSuspendIn") Or UCase(tPdtName2) = UCase("Total IPV In") Or UCase(tPdtName2) = UCase("Total IPV Out") Then
                                                                    If cQty > 0 Then
                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                    End If
                                                                ElseIf UCase(tPdtName) = UCase("Cash Transfer") Then  '*Eaw 56-04-09 CommSheet TK-ISS3000-106 00.00.03.xlsx
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                ElseIf UCase(tPdtName) = UCase("Total Return") Then  '*Ao 2013/08/07       -Total Return
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                    Call oPrn.PrintText(Format("-" & cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                Else
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                    Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                End If
                                                                tTmpBlank = tPdtName
                                                            Else
                                                                If UCase(tPdtName2) = UCase("Total Sale") Then
                                                                    tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                    If tTmp <> vbNullString Then
                                                                        Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                    End If
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                    Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                    tTmpBlank = tPdtName
                                                                End If
                                                                If UCase(tPdtName2) = UCase("Start Bank") Then
                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                    Call oPrn.PrintText("Cash", cX5, , vbLeftJustify, False)
                                                                    Call oPrn.PrintText(Format(0, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                    tTmpBlank = tPdtName
                                                                End If
                                                                '*Eaw 56-08-02 CommSheet TK-ISS3000-299-UAT.xls
                                                                Select Case UCase(tPdtName2) '*Eaw 56-08-29 เปลี่ยน
                                                                    Case UCase("Layaway Deposit"), UCase("Rounding")
                                                                        If cQty > 0 Then
                                                                            tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                            Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                            Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                            Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                            tTmpBlank = tPdtName
                                                                        End If
                                                                End Select
                                                                Select Case UCase(tPdtName2)
'                                                                    Case UCase("Drawer Open"), UCase("CV Cancel"), UCase("CV Pickup"), UCase("TranSuspendIn"), UCase("Total IPV In"), UCase("Total IPV Out")
                                                                    Case UCase("Drawer Open"), UCase("CV Cancel"), UCase("CV Pickup"), UCase("TranSuspendIn"), UCase("Total IPV In"), UCase("Total IPV Out"), UCase("VOID") '*Eaw 56-07-02 CommSheet TK-ISS3000-243-TR.xls เพิ่ม VOID
                                                                        If cQty > 0 Then
                                                                            tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                            If tTmp <> vbNullString Then
                                                                                Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                            End If
                                                                            Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                        End If
                                                                End Select
'                                                                If UCase(SP_FEDtChkString(orsDT, "FTSkuCode")) = UCase("T012") Then
'                                                                    Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
'                                                                    Call oPrn.PrintText(Format(SP_FEDvChkNumeric(orsDT, "FCSdtStkFac"), "#,###.00000"), cXC, , vbLeftJustify, False)
'                                                                    Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
'                                                                End
                                                                '*Eaw 56-09-12 CommSheet TK-ISS3000-351-UAT-FAPL-119.xls
'                                                                Select Case UCase(SP_FEDtChkString(orsDT, "FTSkuCode"))
'                                                                    Case UCase("T001"), UCase("T002"), UCase("T003"), UCase("T004"), UCase("T005"), UCase("T006"), UCase("T007"), UCase("T008"), UCase("T009"), UCase("T010"), UCase("T011"), UCase("T012"), UCase("T013"), UCase("T014"), UCase("T015"), UCase("T016"), UCase("T017"), UCase("T018")
'                                                                        tTmp = IIf(cQty = 0, vbNullString, cQty)
'                                                                        If tTmp <> vbNullString Then
'                                                                            Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
'                                                                        End If
'                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
'                                                                        Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
'                                                                End Select
                                                                '*Ao 2014/07/31   CommSheet TK-ISS3000-460-00 00 03   V4.1.129
                                                                Select Case UCase(SP_FEDtChkString(orsDT, "FTSkuCode"))
                                                                    Case UCase("T001"), UCase("T004"), UCase("T005"), UCase("T006"), UCase("T007"), UCase("T008"), UCase("T009"), UCase("T010"), UCase("T011"), UCase("T012"), UCase("T013"), UCase("T014"), UCase("T016"), UCase("T018")
                                                                        tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                        If tTmp <> vbNullString Then
                                                                            Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                        End If
                                                                        Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, False)
                                                                        Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                 
                                                                    Case UCase("T002"), UCase("T003"), UCase("T015"), UCase("T017")         '*Ao 2014/07/31   CommSheet TK-ISS3000-460-00 00 03   V4.1.129
                                                                        If cQty = 0 Then
                                                                            Call oPrn.PrintText(tPdtName2, cX5, , vbLeftJustify, True)
                                                                        Else
                                                                            tTmp = IIf(cQty = 0, vbNullString, cQty)
                                                                            If tTmp <> vbNullString Then
                                                                                Call oPrn.PrintText(tTmp, cX4, , vbRightJustify, False)
                                                                            End If
                                                                            Call oPrn.PrintText(Space$(1) & tPdtName2, cX5, , vbLeftJustify, False)
                                                                            Call oPrn.PrintText(Format(cNet, tW_CurFmt), cXR, , vbRightJustify, True)
                                                                        End If
                                                                End Select
                                                            End If
                                                    End If
                                            End Select
                                .MoveNext
                                Loop
                                End With
                            End If
                    End If
      '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
               Case 34, 35 'Copy Trans,Copy From Dev
               
                    nPosNo = SP_FEDvChkNumeric(orsHD, "FNShdStaRef")
                    tTrnNo = SP_FEDtChkString(orsHD, "FTShdDocRefMin") & "-" & SP_FEDtChkString(orsHD, "FTShdDocRefMax")
                    
                    tTmp = "TERMINAL NO:"
                    Call oPrn.PrintText(tTmp, cX5, , vbLeftJustify, False)
                    tTmp = nPosNo
                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True)
                    
                    tTmp = "TRAN  NO:"
                    Call oPrn.PrintText(tTmp, cX5, , vbLeftJustify, False)
                    tTmp = tTrnNo
                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True)

      '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
               Case Else
               
               End Select
      '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
      
        Set orsTmp = Nothing
        Set orsTmp2 = Nothing
        Set orsNew1 = Nothing
        Set orsNew2 = Nothing
        Set orsFor = Nothing
        
        'Return
        W_PRNbTicketDetail = True
       
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        'Return
        W_PRNbTicketDetail = True
        Set orsTmp = Nothing
        Set orsTmp2 = Nothing
        Set orsNew1 = Nothing
        Set orsNew2 = Nothing
         Set orsFor = Nothing
End Function

Private Function W_GETcDisc(pnDiscType As Long, Optional ByVal nDTSeq As Long) As Double
Dim nTmp As Long, tTmp As String
Dim cTmp
On Error GoTo ErrHandler
            tTmp = IIf(nDTSeq = 0, vbNullString, " AND FNSdtSeqNo =" & nDTSeq)
            
            If Not orsCD Is Nothing And orsCD.State = adStateOpen Then
                orsCD.Filter = "FNDctNo = " & pnDiscType & tTmp
                
                Do While Not orsCD.EOF
                        cTmp = cTmp + SP_FEDvChkNumeric(orsCD, "FCScdAmt")
                 orsCD.MoveNext
                Loop
            End If
'Return
W_GETcDisc = cTmp

Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function W_GETcDiscByCrdID(ByVal ptCrdID As String, Optional ByVal nDctNo As Long, Optional ByVal nDTSeq As Long) As Double
Dim nTmp As Long, tTmp As String
Dim cTmp
On Error GoTo ErrHandler
        If ptCrdID = "" Then W_GETcDiscByCrdID = 0: Exit Function
            
            tTmp = IIf(nDTSeq = 0, vbNullString, " AND FNSdtSeqNo =" & nDTSeq)
            tTmp = tTmp & IIf(nDctNo = 0, vbNullString, " AND FNDctNo =" & nDctNo)
            
            If Not orsCD Is Nothing And orsCD.State = adStateOpen Then
                orsCD.Filter = "FTScdCardID = '" & ptCrdID & "'" & tTmp '*KT 52-01-12 ให้ใส่เครื่องหมาย Single Code ด้วยในการ Filter
                
                Do While Not orsCD.EOF
                        cTmp = cTmp + SP_FEDvChkNumeric(orsCD, "FCScdAmt")
                 orsCD.MoveNext
                Loop
            End If
'Return
W_GETcDiscByCrdID = cTmp

Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function W_GETcDiscByCrdIDEx(ByVal ptCrdID As String, Optional ByVal nDctNo As Long, Optional ByVal nDTSeq As Long) As Double
'*KT 51-10-31 แก้ไขเฉพาะหน้า ให้ AO เพิ่มหน้าจอการคืนของส่วนลด Emp card ที่ต้องระบุหมายเลขบัตรด้วย
Dim nTmp As Long, tTmp As String
Dim cTmp
On Error GoTo ErrHandler
             tTmp = "FNDctNo =" & nDctNo & " AND FTScdCardID = '" & ptCrdID & "'"
           ' tTmp = "FNDctNo =" & nDctNo
            If Not orsCD Is Nothing And orsCD.State = adStateOpen Then
                orsCD.Filter = tTmp
                Do While Not orsCD.EOF
                        cTmp = cTmp + SP_FEDvChkNumeric(orsCD, "FCScdAmt")
                 orsCD.MoveNext
                Loop
            End If
'Return
W_GETcDiscByCrdIDEx = cTmp

Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function W_GETtCrdTypeByCrdID(ByVal ptCrdID As String, Optional ByVal nDctNo As Long) As String
Dim tTmp As String
On Error GoTo ErrHandler
        If ptCrdID = "" Then W_GETtCrdTypeByCrdID = "": Exit Function
            tTmp = tTmp & IIf(nDctNo = 0, vbNullString, " AND FNDctNo =" & nDctNo)
            If Not orsCD Is Nothing And orsCD.State = adStateOpen Then
                orsCD.Filter = "FTScdCardID = '" & ptCrdID & "'" & tTmp
                tTmp = ""
                Do While Not orsCD.EOF
                        tTmp = SP_FEDtChkString(orsCD, "FTScdCardType")
                        orsCD.MoveNext
                Loop
            End If
'Return
W_GETtCrdTypeByCrdID = tTmp
Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
End Function

Private Function W_GETtRCByCrdID(ByVal ptCrdID As String) As String
Dim tTmp As String
On Error GoTo ErrHandler
        If ptCrdID = "" Then W_GETtRCByCrdID = "": Exit Function
            If Not orsRC Is Nothing And orsRC.State = adStateOpen Then
                orsRC.Filter = "(FTTdmCode = 'T002' AND FTSrcCardNo ='" & ptCrdID & "')  OR  (FTTdmCode = 'T003' AND FTSrcCardNo ='" & ptCrdID & "')"
                Do While Not orsRC.EOF
                    If SP_VALcText2Double(SP_FEDvChkNumeric(orsRC, "FCSrcOntopAmt")) > 0 Then
                        tTmp = SP_FEDtChkString(orsRC, "FTSrcCardNo")
                    End If
                    orsRC.MoveNext
                Loop
            End If
            orsRC.Filter = vbNullString
'Return
W_GETtRCByCrdID = tTmp

Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function W_GETcCardTypeDesc(ptDiscType As String) As String
Dim tSql$
Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandler
            W_GETcCardTypeDesc = ""
            tSql = "SELECT  FTMcdDesc1,FTMcdDesc2"
            tSql = tSql & " From TCNMMallCardMtn"
            tSql = tSql & " WHERE  FTMcdCardType = '" & ptDiscType & "'"
'            Set orsTemp = ocnAll.Execute(tSql)
            If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then
                If Not (orsTemp.BOF Or orsTemp.EOF) Then
                    W_GETcCardTypeDesc = SP_FEDtChkString(orsTemp, "FTMcdDesc2")
                Else
                    W_GETcCardTypeDesc = ""
                End If
            Else
                W_GETcCardTypeDesc = ""
            End If
            orsTemp.Close
            Set orsTemp = Nothing
Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
    Set orsTemp = Nothing
End Function

Private Function W_GETcEmpCardTypeDesc(ptDiscType As String) As String
Dim tSql$
Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandler
            W_GETcEmpCardTypeDesc = ""
            tSql = "SELECT  FTEcdDesc1,FTEcdDesc2"
            tSql = tSql & " From TCNMEmpCardMtn"
            tSql = tSql & " WHERE  FTEcdCardType = '" & ptDiscType & "'"
'            Set orsTemp = ocnAll.Execute(tSql)
             If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then
                If Not (orsTemp.BOF Or orsTemp.EOF) Then
                        W_GETcEmpCardTypeDesc = SP_FEDtChkString(orsTemp, "FTEcdDesc2")
                End If
            Else
                W_GETcEmpCardTypeDesc = ""
            End If
            orsTemp.Close
            Set orsTemp = Nothing
Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function W_GETcPromoAmt(Optional ByVal nDTSeq As Long) As Double
Dim nTmp As Long, tTmp As String
Dim cTmp
On Error GoTo ErrHandler
            If nDTSeq = 0 Then
                    orsCD.Filter = "(FNDctNo = 2) OR (FNDctNo = 3)  "
            Else
                    orsCD.Filter = "(FNSdtSeqNo =" & nDTSeq & " AND FNDctNo = 2) OR (FNSdtSeqNo =" & nDTSeq & " AND FNDctNo = 3)"
            End If
            If Not orsCD Is Nothing And orsCD.State = adStateOpen Then
                Do While Not orsCD.EOF
                        cTmp = cTmp + SP_FEDvChkNumeric(orsCD, "FCScdAmt")
                 orsCD.MoveNext
                Loop
            End If
'Return
W_GETcPromoAmt = cTmp

Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function W_GETcPromoDealAmt(Optional ByVal nDTSeq As Long) As Double
Dim nTmp As Long, tTmp As String
Dim cTmp
On Error GoTo ErrHandler
            W_GETcPromoDealAmt = 0
            orsCD.Filter = "(FNSdtSeqNo =" & nDTSeq & " AND FNDctNo = 1) OR (FNSdtSeqNo =" & nDTSeq & " AND FNDctNo = 2 ) OR (FNSdtSeqNo =" & nDTSeq & " AND FNDctNo = 3)"
            If Not orsCD Is Nothing And orsCD.State = adStateOpen Then
                Do While Not orsCD.EOF
                        cTmp = cTmp + SP_FEDvChkNumeric(orsCD, "FCScdAmt")
                        orsCD.MoveNext
                Loop
            End If
'Return
W_GETcPromoDealAmt = cTmp

Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function W_GETcDealAmt(ByVal tDealNum As String, ByVal pnDealGroup As Long) As Double
Dim nTmp As Long, tTmp As String
Dim cTmp
On Error GoTo ErrHandler
            If Not orsCD Is Nothing And orsCD.State = adStateOpen Then
                orsCD.Filter = "(FNDctNo = 1) AND (FTDctRefCode ='" & tDealNum & "')  AND  (FNScdDealGroup =" & pnDealGroup & ")"
    
                Do While Not orsCD.EOF
                        cTmp = cTmp + SP_FEDvChkNumeric(orsCD, "FCScdAmt")
                 orsCD.MoveNext
                Loop
            End If
'Return
W_GETcDealAmt = cTmp

Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
  
End Function

Private Function W_PRNbTicketTransDisc() As Boolean
'KT 51-10-31
Dim nTmp As Long, tTmp As String, cTmp As Double
Dim tTmpID As String
Dim tTmpType As String

Dim nDiscType As Long, cDiscVal As Double
Dim tPriceType As String

Dim orsTmp As ADODB.Recordset

Dim bPass As Boolean
Dim nI As Long

'Const tTkDsMallCpn = "The Mall Coupon"
Dim tTkDsMallCpn$
tTkDsMallCpn = W_GETxVocabularyTicketFromArray("1911")
'Const tTkDsVipCrd = "VIP Card"
Dim tTkDsVipCrd$
tTkDsVipCrd = W_GETxVocabularyTicketFromArray("1920")
'Const tTkDsEmpCrd = "EMPL"
Dim tTkDsEmpCrd$
tTkDsEmpCrd = W_GETxVocabularyTicketFromArray("1922")

'Const tTkDsTransDisc = "TRAN.DISC."
Dim tTkDsTransDisc$
tTkDsTransDisc = W_GETxVocabularyTicketFromArray("190901")
'Const tTkDsTransDisc = "TRAN.DISC."
Dim tTkDsTransDiscAmt$
tTkDsTransDiscAmt = W_GETxVocabularyTicketFromArray("190902")
'Const tTkDsTransDisc = "TRAN.DISC."
Dim tTkDsTransDiscNew$
tTkDsTransDiscNew = W_GETxVocabularyTicketFromArray("190903")

'Const tTkDsTransMkDn = "TRAN.MKDN."
Dim tTkDsTransMkDn$
tTkDsTransMkDn = W_GETxVocabularyTicketFromArray("191001")
'Const tTkDsTransMkDn = "TRAN.MKDN."
Dim tTkDsTransMkDnAmt$
tTkDsTransMkDnAmt = W_GETxVocabularyTicketFromArray("191002")
'Const tTkDsTransMkDn = "TRAN.MKDN."
Dim tTkDsTransMkDnNew$
tTkDsTransMkDnNew = W_GETxVocabularyTicketFromArray("191003")

'Const tTkDsAutoTransDisc = "AUTO TRAN.DISC."
Dim tTkDsAutoTransDisc$
tTkDsAutoTransDisc = W_GETxVocabularyTicketFromArray("1913")

'Const tTkDsOnTop = "On Top"
Dim tTkDsOnTop$
tTkDsOnTop = W_GETxVocabularyTicketFromArray("1912")

'Const tTkSubAll = "SUBTOTAL-ALL"
Dim tTkSubAll$
tTkSubAll = W_GETxVocabularyTicketFromArray("1944")
'Const tTkSubReg = "SUBTOTAL-REG"
Dim tTkSubReg$
tTkSubReg = W_GETxVocabularyTicketFromArray("1945")

'Const tTkDsTransMkDn1 = "TRAN.MKDN.1"
Dim tTkDsTransMkDn1$
tTkDsTransMkDn1 = W_GETxVocabularyTicketFromArray("199201")
'Const tTkDsTransMkDn1 = "TRAN.MKDN.1"
Dim tTkDsTransMkDnAmt1$
tTkDsTransMkDnAmt1 = W_GETxVocabularyTicketFromArray("199202")
'Const tTkDsTransMkDn1 = "TRAN.MKDN.1"
Dim tTkDsTransMkDnNew1$
tTkDsTransMkDnNew1 = W_GETxVocabularyTicketFromArray("199203")

'Const tTkDsTransMkDn2 = "TRAN.MKDN.2"
Dim tTkDsTransMkDn2$
tTkDsTransMkDn2 = W_GETxVocabularyTicketFromArray("199301")
'Const tTkDsTransMkDn2 = "TRAN.MKDN.2"
Dim tTkDsTransMkDnAmt2$
tTkDsTransMkDnAmt2 = W_GETxVocabularyTicketFromArray("199302")
'Const tTkDsTransMkDn2 = "TRAN.MKDN.2"
Dim tTkDsTransMkDnNew2$
tTkDsTransMkDnNew2 = W_GETxVocabularyTicketFromArray("199303")

'Const tTkDsTransMkDn3 = "TRAN.MKDN.3"
Dim tTkDsTransMkDn3$
tTkDsTransMkDn3 = W_GETxVocabularyTicketFromArray("199401")
'Const tTkDsTransMkDn3 = "TRAN.MKDN.3"
Dim tTkDsTransMkDnAmt3$
tTkDsTransMkDnAmt3 = W_GETxVocabularyTicketFromArray("199402")
'Const tTkDsTransMkDn3 = "TRAN.MKDN.3"
Dim tTkDsTransMkDnNew3$
tTkDsTransMkDnNew3 = W_GETxVocabularyTicketFromArray("199403")

'Const tTkDsTransMkDn4 = "TRAN.MKDN.4"
Dim tTkDsTransMkDn4$
tTkDsTransMkDn4 = W_GETxVocabularyTicketFromArray("199501")
'Const tTkDsTransMkDn4 = "TRAN.MKDN.4"
Dim tTkDsTransMkDnAmt4$
tTkDsTransMkDnAmt4 = W_GETxVocabularyTicketFromArray("199502")
'Const tTkDsTransMkDn4 = "TRAN.MKDN.4"
Dim tTkDsTransMkDnNew4$
tTkDsTransMkDnNew4 = W_GETxVocabularyTicketFromArray("199503")

'Const tTkDsTransMkDn5 = "TRAN.MKDN.5"
Dim tTkDsTransMkDn5$
tTkDsTransMkDn5 = W_GETxVocabularyTicketFromArray("199601")
'Const tTkDsTransMkDn5 = "TRAN.MKDN.5"
Dim tTkDsTransMkDnAmt5$
tTkDsTransMkDnAmt5 = W_GETxVocabularyTicketFromArray("199602")
'Const tTkDsTransMkDn5 = "TRAN.MKDN.5"
Dim tTkDsTransMkDnNew5$
tTkDsTransMkDnNew5 = W_GETxVocabularyTicketFromArray("199603")


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
                        'cW_DiscTotalTran = 0
                        'cW_DiscTotalTran2 = 0

                        '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
                        'Trans Discount B4 SubTotal
                        ReDim aW_DctNo(0) 'clear array
                        
                        aW_DctNo(0) = ""
                        If orsCD.State = 1 Then
                                Set orsTmp = orsCD.Clone
                                orsTmp.Filter = vbNullString '"FNSdtSeqNo=1    "
                                orsTmp.Sort = "FNScdSeqNo ASC"
                                Do While Not orsTmp.EOF
                                        nTmp = SP_FEDvChkNumeric(orsTmp, "FNDctNo")
                                        
'                                        Select Case nTmp
'                                            Case nEN_Disc11MallCrd, nEN_Disc12EmpCrd
'                                                cW_DiscTotalTran = cW_DiscTotalTran + SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
'                                        End Select
'
'                                        Select Case nTmp
'                                            Case nEN_Disc11MallCrd, nEN_Disc12EmpCrd, nEN_Disc15Ontop
'                                                cW_DiscTotalTran2 = cW_DiscTotalTran2 + SP_FEDvChkNumeric(orsTmp, "FCScdAmt")
'                                        End Select

                                        bPass = True
                                        For nI = 0 To UBound(aW_DctNo) - 1
                                            If aW_DctNo(nI) = nTmp Then bPass = False: Exit For
                                        Next nI
                                        aW_DctNo(UBound(aW_DctNo)) = nTmp
                                        ReDim Preserve aW_DctNo(UBound(aW_DctNo) + 1)
                                        If bPass Then
                                            cDiscVal = W_GETcDisc(nTmp)
                                            
                                            If cDiscVal >= 0 Then
                                            
                                                    Select Case nTmp
                                                    Case nEN_Disc08MallCpn
                                                            tTmp = tTkDsMallCpn
                                                            
                                                    Case nEN_Disc13TranDisc
                                                            Select Case SP_FEDtChkString(orsTmp, "FTScdDiscType")
                                                                Case "1"
                                                                    tTmp = tTkDsTransDisc
                                                                Case "2"
                                                                    tTmp = tTkDsTransDiscAmt
                                                                Case "3"
                                                                    tTmp = tTkDsTransDiscNew
                                                            End Select
                                                            
                                                    Case nEN_Disc14TranMDown
                                                            Select Case SP_FEDtChkString(orsTmp, "FTScdDiscType")
                                                                Case "1"
                                                                    tTmp = tTkDsTransMkDn
                                                                Case "2"
                                                                    tTmp = tTkDsTransMkDnAmt
                                                                Case "3"
                                                                    tTmp = tTkDsTransMkDnNew
                                                            End Select
                                                            
                                                    Case nEN_Disc23TranMDown1  '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                                                            Select Case SP_FEDtChkString(orsTmp, "FTScdDiscType")
                                                                Case "1"
                                                                    tTmp = tTkDsTransMkDn1
                                                                Case "2"
                                                                    tTmp = tTkDsTransMkDnAmt1
                                                                Case "3"
                                                                    tTmp = tTkDsTransMkDnNew1
                                                            End Select
                                                            
                                                    Case nEN_Disc24TranMDown2 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                                                            Select Case SP_FEDtChkString(orsTmp, "FTScdDiscType")
                                                                Case "1"
                                                                    tTmp = tTkDsTransMkDn2
                                                                Case "2"
                                                                    tTmp = tTkDsTransMkDnAmt2
                                                                Case "3"
                                                                    tTmp = tTkDsTransMkDnNew2
                                                            End Select
                                                            
                                                    Case nEN_Disc25TranMDown3 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                                                            Select Case SP_FEDtChkString(orsTmp, "FTScdDiscType")
                                                                Case "1"
                                                                    tTmp = tTkDsTransMkDn3
                                                                Case "2"
                                                                    tTmp = tTkDsTransMkDnAmt3
                                                                Case "3"
                                                                    tTmp = tTkDsTransMkDnNew3
                                                            End Select
                                                            
                                                    Case nEN_Disc26TranMDown4 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                                                            Select Case SP_FEDtChkString(orsTmp, "FTScdDiscType")
                                                                Case "1"
                                                                    tTmp = tTkDsTransMkDn4
                                                                Case "2"
                                                                    tTmp = tTkDsTransMkDnAmt4
                                                                Case "3"
                                                                    tTmp = tTkDsTransMkDnNew4
                                                            End Select
                                                            
                                                    Case nEN_Disc27TranMDown5 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                                                            Select Case SP_FEDtChkString(orsTmp, "FTScdDiscType")
                                                                Case "1"
                                                                    tTmp = tTkDsTransMkDn5
                                                                Case "2"
                                                                    tTmp = tTkDsTransMkDnAmt5
                                                                Case "3"
                                                                    tTmp = tTkDsTransMkDnNew5
                                                            End Select
                                                            
                                                    Case nEN_Disc07AutoTran
                                                            tTmp = tTkDsAutoTransDisc
                                                            
                                                    Case Else
                                                            tTmp = vbNullString
                                                            
                                                    End Select
                                                    
                                                    If tTmp <> vbNullString Then
                                                        Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                                                        Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXC, , vbLeftJustify, False)
                                                        Call oPrn.PrintText(" OFF", cXHR, , vbLeftJustify, True)
                                                    End If
                                                                 
                                            End If
                                        End If
                                        orsTmp.MoveNext
                                Loop
                                
                        End If
                        
                        '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
                        'SubTotal
                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
                        'Sub All
                        If tTkSubAll = "" Then tTkSubAll = "SUBTOTAL-ALL" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                        Call oPrn.PrintText(tTkSubAll, cX5, , vbLeftJustify, False)
                        
                        Select Case Val(tW_TransTypeSub)
                            Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnDepositCancel, nEN_TrnAddDepositCancel, nEN_TrnRedeem
                                ' tTmp = Format(cW_SubAll + cW_DiscTotalTran2, tW_CurFmt)
                                 tTmp = Format(cW_SubAll, tW_CurFmt)
                            Case Else
                                ' tTmp = Format(cW_SubAll + cW_DiscTotalTran2, tW_CurFmt)
                                 tTmp = Format(cW_SubAll, tW_CurFmt)
                        End Select
                        
                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                        Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                        'Sub Reg
                        If bW_PrintMem Then                                     '*Ao 2013/03/21  CommSheet TK-ISS3000-115 00.00.01.xlsx
                            Call oPrn.PrintText(tTkSubReg, cX5, , vbLeftJustify, False)
                            tTmp = Format(cW_SubReg, tW_CurFmt)                                             '*Ao 2013/03/08  CommSheet TK-ISS3000-026-00 00 04.xls
                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                        End If
'                        Call oPrn.PrintText(tTkSubReg, cX5, , vbLeftJustify, False)
'                        'tTmp = Format(cW_SubReg + cW_DiscTotalTran, tW_CurFmt)
'                        tTmp = Format(cW_SubReg, tW_CurFmt)                                             '*Ao 2013/03/08  CommSheet TK-ISS3000-026-00 00 04.xls
'                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
'                        Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                
                        '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
                        'Trans Discount  After SubTotal
                        cW_OnTopDps = 0
                        bW_ShwSubTotal = False
                        ReDim aW_DctNo(0) 'clear array
                        ReDim aW_DctNo2(0) 'clear array
                        aW_DctNo(0) = ""
                        aW_DctNo2(0) = ""
                        If orsCD.State = 1 Then
                                Set orsTmp = orsCD.Clone
                                
                                orsTmp.Filter = vbNullString '"FNSdtSeqNo=1  "
                                orsTmp.Filter = "FNDctNo = 11 OR FNDctNo = 12  OR FNDctNo = 15"
                                orsTmp.Sort = "FNScdSeqNo ASC"
                                Do While Not orsTmp.EOF
                                
'Comsheet ML-SF-270 แสดงส่วนลดแต่ละหมายเลขบัตร
'----------------------------------------------------------------------------------------------------------------------------------
                                        nTmp = SP_FEDvChkNumeric(orsTmp, "FNDctNo")
                                        tTmpID = SP_FEDtChkString(orsTmp, "FTScdCardID")
                                        tTmpType = SP_FEDtChkString(orsTmp, "FTScdCardType")
                                        
                                        bPass = True
                                        For nI = 0 To UBound(aW_DctNo) - 1
                                            If aW_DctNo(nI) = tTmpID And aW_DctNo2(nI) = nTmp Then bPass = False: Exit For
                                        Next nI
                                        aW_DctNo(UBound(aW_DctNo)) = tTmpID
                                        aW_DctNo2(UBound(aW_DctNo)) = nTmp
'----------------------------------------------------------------------------------------------------------------------------------
                                        ReDim Preserve aW_DctNo(UBound(aW_DctNo) + 1)
                                        ReDim Preserve aW_DctNo2(UBound(aW_DctNo) + 1)
                                        
                                        If bPass Then
                                            Select Case Val(tW_TransTypeSub)
                                                Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                                    cDiscVal = W_GETcDiscByCrdIDEx(tTmpID, nTmp)
                                                Case Else
                                                    cDiscVal = W_GETcDiscByCrdID(tTmpID, nTmp)
                                            End Select
'                                            If cDiscVal >= 0 Then
                                            If cDiscVal > 0 Then                            '*Ao 2013/03/21   TK
                                                    Select Case nTmp
                                                        Case nEN_Disc11MallCrd
                                                                tTmp = W_GETcCardTypeDesc(SP_FEDtChkString(orsTmp, "FTScdCardType"))  'tTkDsVipCrd
                                                                If tTmp = vbNullString Then
                                                                    tTmp = tTkDsVipCrd
                                                                End If
                                                                If tTmp <> "" Then                  '*Ao 2013/01/03   พิมพ์ %
                                                                    tTmp = tTmp & " " & SP_FEDtChkString(orsTmp, "FCScdDiscPerOrg") & "%"
                                                                End If
                                                                bW_ShwSubTotal = True
                                                        Case nEN_Disc12EmpCrd
                                                                tTmp = W_GETcEmpCardTypeDesc(SP_FEDtChkString(orsTmp, "FTScdCardType"))   'tTkDsEmpCrd
                                                                If tTmp = vbNullString Then
                                                                    tTmp = tTkDsEmpCrd
                                                                End If
                                                                If tTmp <> "" Then                  '*Ao 2013/01/03   พิมพ์ %
                                                                    tTmp = tTmp & " " & SP_FEDtChkString(orsTmp, "FCScdDiscPerOrg") & "%"
                                                                End If
                                                                bW_ShwSubTotal = True
                                                        Case nEN_Disc15Ontop
                                                                tTmp = tTkDsOnTop & " " & tTmpType
                                                                bW_ShwSubTotal = True
                                                        Case Else
                                                                tTmp = vbNullString
                                                    End Select
                                                    
                                                    'Communication Sheet ML-SF-478 Error on Deposit Transaction.xls
                                                    '*KT 51-11-20
                                                    '-------------------------------------------------------------------------------------------------------
                                                    Select Case Val(tW_TransTypeSub)
                                                        Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnRedeem, nEN_TrnDepositCancel, nEN_TrnAddDepositCancel
                                                            Select Case nTmp
                                                                Case nEN_Disc15Ontop
                                                                    If W_GETtRCByCrdID(tTmpID) <> "" Then
                                                                        cW_OnTopDps = cW_OnTopDps + cDiscVal
                                                                        tTmp = vbNullString
                                                                    End If
                                                            End Select
                                                    End Select
                                                    '----------------------------------------------------------------------------------------------------
                                                    
                                                    If tTmp <> vbNullString Then
                                                            'Line_Chars
'                                                            Call oPrn.PrintText(Mid(tTmp, 1, 17), cX5, , vbLeftJustify, False)             '*Ao 2013/02/02         ไม่ต้องตัดข้อความ
                                                            Call oPrn.PrintText(tTmp, cX5, , vbLeftJustify, False)
                                                            Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXMK, , vbRightJustify, False)
                                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                                    End If
                                            End If
                                        End If
                                        orsTmp.MoveNext
                                Loop
                        End If
                        
                        '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

            Set orsTmp = Nothing
            'Return
            W_PRNbTicketTransDisc = True
            
            Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
              'Return
            W_PRNbTicketTransDisc = True
            Set orsTmp = Nothing
End Function

Private Function W_PRNbTicketDepInfo() As Boolean
Dim nTmp As Long, tTmp As String

'Const tTkPost = "POST#"
Dim tTkPost$
tTkPost = W_GETxVocabularyTicketFromArray("1968")
'Const tTkTran = "TRAN#"
Dim tTkTran$
tTkTran = W_GETxVocabularyTicketFromArray("1969")
'Const tTkTranDate = "TRANDATE"
Dim tTkTranDate$
tTkTranDate = W_GETxVocabularyTicketFromArray("1970")

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler

'                                Call oPrn.PrintText(tTkPost, cXL, , vbLeftJustify, False)  'post#
                                tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTkPost & "{"
'                                Call oPrn.PrintText(tTkTran, cXC - cXHC, , vbLeftJustify, False) 'tran#
                                tW_SlipFooterStr = tW_SlipFooterStr & "5" & tTkTran & "{"
'                                Call oPrn.PrintText(tTkTranDate, cXHR - cX2, , vbLeftJustify, True) ' trandate
                                tW_SlipFooterStr = tW_SlipFooterStr & "6" & tTkTranDate & ";"

                                tTmp = SP_FEDtChkString(orsHD, "FTShdDepRefTmnNum")
'                                Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)  'post#
                                tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                                
                                tTmp = SP_FEDtChkString(orsHD, "FTShdDepReTransNo")
'                                Call oPrn.PrintText(tTmp, cXC - cXHC, , vbLeftJustify, False) 'tran#
                                tW_SlipFooterStr = tW_SlipFooterStr & "5" & tTmp & "{"
                                 
'                                 tTmp = Format(SP_FEDdChkDate(orsHD, "FDShdTransDate"), tW_DateFmt)
                                tTmp = Format(SP_FEDdChkDate(orsHD, "FDShdDepReTransDate"), tW_DateFmt) '*Eaw 56-07-11 CommSheet TK-ISS3000-256-TR.xls
'                                Call oPrn.PrintText(tTmp, cXHR - cX2, , vbLeftJustify, True) ' amt
                                tW_SlipFooterStr = tW_SlipFooterStr & "6" & tTmp & ";"
                                    
         '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    
        'Return
        W_PRNbTicketDepInfo = True
        
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        'Return
        W_PRNbTicketDepInfo = True
End Function

Private Function W_PRNbTicketBarcode() As Boolean
Dim nTmp As Long, tTmp As String

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
                        tTmp = W_GETxVocabularyTicketFromArray("18")

                        Select Case Val(tW_TransType)
                            Case 38, 39, 40, 41, 42
                                    Set oPrn.Font = oW_Font(2)
'                                    Call oPrn.PrintText("** REENTRY MODE **", cXC, , vbCenter, , , , , True)       '*Ao 2014/07/19
                                    Call oPrn.PrintText(tTmp, cXC, , vbCenter, , , , , True)
                            Case 43, 44, 45, 46, 47, 48, 49 '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls เพิ่ม 48,49
                                    Set oPrn.Font = oW_Font(2)
'                                    Call oPrn.PrintText("** REENTRY MODE **", cXC, , vbCenter, , , , , True)           '*Ao 2014/07/19
                                    Call oPrn.PrintText(tTmp, cXC, , vbCenter, , , , , True)
                        End Select
                        
                        '*KT 51-04-08 ตรวจสอบว่า StorePara ให้มีการ Print หรือไม่
                        If Mid(tW_OptAlw, 4, 1) = "0" Or Mid(tW_OptAlw, 4, 1) = "" Then W_PRNbTicketBarcode = True: Exit Function

                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True, nEN_PrinterOnly)  ' space - - - - - -
                        'Barcode
                        tTmp = Format(Left(tW_TmnNum, 5), "00000") & Format(Right(tW_TransNo, 5), "00000")
                        If bC_PrnByOPOS Then 'Print by OPOS
                           If oOpos.RecLineWidth <= nVB_MaxChar * 10 Then
                                oOpos.PrintBarcode PTR_S_RECEIPT, tTmp, PTR_BCS_Code39, 50, oOpos.RecLineWidth, DISP_BM_CENTER, PTR_BC_TEXT_NONE
                           Else
                                oOpos.PrintBarcode PTR_S_RECEIPT, tTmp, PTR_BCS_Code39, 50, nVB_MaxChar * 10, DISP_BM_CENTER, PTR_BC_TEXT_NONE
                           End If
                           Call oPrn.PrintText(tTmp, cXC, , vbCenter, True, nEN_PrinterOnly)
                        Else  'Print by Driver
                            Set oPrn.Font = oW_Font(3)
                            Call oPrn.PrintText("@" & tTmp & "@", cXC, , vbCenter, True, nEN_PrinterOnly)   'for barcode 39
                            Set oPrn.Font = oW_Font(1)
                            Call oPrn.PrintText(tTmp, cXC, , vbCenter, True, nEN_PrinterOnly)
                        End If
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        'Return
        W_PRNbTicketBarcode = True
        
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        'Return
        W_PRNbTicketBarcode = True
End Function

Private Function W_PRNbTicketFooterMsg() As Boolean
Dim nTmp As Long, tTmp As String

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler


        'Slip Trailer
    If aW_Trailer(1) <> vbNullString Then
'        Call oPrn.PrintText("", cXC, , vbCenter, True, nEN_PrinterOnly)                 '*Ao 2014/07/19  TK-ISS3000PH2-525   V 4.2.138
        tTmp = " "
        Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True, nEN_PrinterOnly)
    End If

    Set oPrn.Font = oW_Font(1)
    For nTmp = 1 To 7
        If aW_Trailer(nTmp) <> vbNullString Then
            Call oPrn.PrintText(aW_Trailer(nTmp), cXC, , vbCenter, True, nEN_PrinterOnly)  '*Ao 2014/06/20
        End If
    Next

    W_PRNbTicketFooterMsg = True
        
    Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        'Return
        W_PRNbTicketFooterMsg = True
End Function

Private Sub W_PRNbTicketShowOffline()
Dim nTmp As Long, tTmp As String

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler

    'Show Status Offline
    If tW_StatusOffline = "0" Then '*KT 52-06-25 เปลี่ยนจากการขายแบบ Local
'        bVB_LastLine = False '*Eaw 56-10-03
        Set oPrn.Font = oW_Font(2)
'        Call oPrn.PrintText("OFFLINE", cXC, , vbCenter, , , , , True)
        tW_SlipFooterStr = tW_SlipFooterStr & "4" & "OFFLINE" & ";"
'        bVB_LastLine = True '*Eaw 56-10-03 CommSheet TK-ISS3000-364-00.00.03.xls
        Set oPrn.Font = oW_Font(1)
    Else
        '*Tao 61-04-27 ตรวจสอบบาง TransTypeSub
        Select Case tW_TransTypeSub
        Case "01", "02", "05", "08", "09", "12", "14"
        Case "17", "18", "21", "24", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37"
        Case "40", "41", "42", "45", "46", "47", "48", "49", "50"
        Case Else
            '*Bmp 61-04-12 v.เพิ่มการพิม์ กรณี Web service Token เกิด Offline POSFront[5.1804.9].exe
            'bVB_ChkWSOFFLINE เรียกใช้เฉพาะ ใน wFunctionKB.W_UPDxPointToMem
            'If bVB_ChkWSOFFLINE Then   '* [5.1805.0][61/05/03] *Bmp 61-05-03 '*Tao 61-05-25 ปิดไว้
            If Not bW_PRMGenEJ Then '*Thirapong 63-08-03 Gen EJ ไม่ต้องเช็ค bVB_ChkWSOFFLINE
                If Not bVB_ChkWSOFFLINE Then '*Tao 61-05-25 ตรวจสอบมาแล้ว จากการ Token ใน W_UPDxPointToMem
                    'If Not SP_CHKbWebservice Then '*Tao 61-05-25 ไม่ต้องตรวจสอบอีกแว้ว
                        tW_SlipFooterStr = tW_SlipFooterStr & "1" & Space(1) & ";"  '*Tao 61-05-25 เว้น 1 บรรทัด
                        Set oPrn.Font = oW_Font(2)
                        'tW_SlipFooterStr = tW_SlipFooterStr & "4" & "OFFLINE" & ";" '*Tao 61-04-20 ของเดิม
                        tW_SlipFooterStr = tW_SlipFooterStr & "4" & "OFFLINE*" & ";" '*Tao 61-04-20 เพิ่ม *
                        Set oPrn.Font = oW_Font(1)
                    'End If
                End If
                'bVB_ChkWSOFFLINE = False '*Reset '*Tao 61-05-25 ปิดไว้
            End If
        End Select
    End If
Exit Sub
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
End Sub

Public Function W_PRNbTicketHeaderMsg() As Boolean
Dim nTmp As Long, tTmp As String

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler


                        'Slip Head
                        Set oPrn.Font = oW_Font(1)
                        For nTmp = 1 To 7
                                If aW_Header(nTmp) <> vbNullString Then
                                        Call oPrn.PrintText(aW_Header(nTmp), cXC, , vbCenter, True, nEN_PrinterOnly)
                                End If
                        Next
        'Return
        W_PRNbTicketHeaderMsg = True

        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        'Return
        W_PRNbTicketHeaderMsg = True
End Function

Private Function W_PRNbTicketTotal() As Boolean
Dim nTmp As Long, tTmp As String

'Const tTkSubTotal = "SUBTOTAL"
Dim tTkSubTotal$
tTkSubTotal = W_GETxVocabularyTicketFromArray("1946")
'Const tTkMiscRcv = "Misc Recv."
Dim tTkMiscRcv$
tTkMiscRcv = W_GETxVocabularyTicketFromArray("1947")
'Const tTkTotal = "TOTAL"
Dim tTkTotal$
tTkTotal = W_GETxVocabularyTicketFromArray("1948")
'Const tTkItem = "Items"
Dim tTkItem$
tTkItem = W_GETxVocabularyTicketFromArray("1949")
'Const tTkVatInc = "VAT(Inclouded)"
Dim tTkVatInc$
tTkVatInc = W_GETxVocabularyTicketFromArray("1950")
'Const tTkVatExc = "VAT(Exclouded)"
Dim tTkVatExc$
tTkVatExc = W_GETxVocabularyTicketFromArray("1951")
 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 
 '*Tao 60-04-10 ยอดรวม Normal Amount
Dim tTkRegAmt As String
tTkRegAmt = W_GETxVocabularyTicketFromArray("2052")
 '*Tao 60-04-10 ยอดรวม LowBP Amount
Dim tTkLowBPAmt As String
tTkLowBPAmt = W_GETxVocabularyTicketFromArray("2053")

On Error GoTo ErrHandler
                    'SubTotal
                    If tTkSubTotal = "" Then tTkSubTotal = "SUBTOTAL" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                    
                    Select Case Val(tW_TransTypeSub)
'*Eaw 56-07-24 Layaway ไม่แสดง Subtotal CommSheet TK-ISS3000-235-ST-00.00.02.xls
'                    Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnRedeem
'                            Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
'                            Call oPrn.PrintText(tTkSubTotal, cX5, , vbLeftJustify, False)  '*KT 51-06-11 oOPOS
'                            tTmp = Format(cW_Total + cW_OnTopDps, tW_CurFmt)
'                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
'                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                    Case nEN_TrnSuspended
                            Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
                             Call oPrn.PrintText(tTkSubTotal, cX5, , vbLeftJustify, False) '*KT 51-06-11 oOPOS
'                             tTmp = Format(cW_Total + cW_Rnd, tW_CurFmt)
                             tTmp = Format(cW_Total, tW_CurFmt)                         '*Eaw 56-04-29   CommSheet TK-ISS3000-105-FSBT Answer.xlsx
                             Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                             Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                    Case Else
                            If bW_ShwSubTotal Then
                                Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
                                Call oPrn.PrintText(tTkSubTotal, cX5, , vbLeftJustify, False) '*KT 51-06-11 oOPOS
                                tTmp = Format(cW_Total, tW_CurFmt)
                                Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                            End If
                    End Select
                    
'*Eaw 56-04-29   CommSheet TK-ISS3000-105-FSBT Answer.xlsx
'                    Select Case Val(tW_TransTypeSub)
'                        Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnRedeem, nEN_TrnDepositCancel, nEN_TrnAddDepositCancel
'                            'none
'                        Case Else
'                            'Misc Recv
'                            If cW_Rnd <> 0 Then
'                                    Call oPrn.PrintText(tTkMiscRcv, cX5, , vbLeftJustify, False)
'                                    tTmp = Format(cW_Rnd, tW_CurFmt)
'                                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
'                                    Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
'                            End If
'                    End Select
                    
                    '*Tao 60-04-10 เพิ่ม Low BP
                    If bVB_PrintRegSale Then '*Tao 60-09-04 Allow printing of "REGULAR","SALES"(2lines)
                        'REGULAR
                        If tTkRegAmt = "" Then tTkRegAmt = "REGULAR"
                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  'space
                        Call oPrn.PrintText(tTkRegAmt, cX5, , vbLeftJustify, False)
                        'tTmp = Format(cW_RegAmt, tW_CurFmt) '*Tao 60-05-02 ของเดิม
                        tTmp = Format(cW_RegAmt, tW_CurFmt) & Space(14) '*Tao 60-05-02 เว้น 14 ตัว
                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) '*Tao 60-04-27 ของเดิม แนวเดียวกัน
                        'SALES
                        If tTkLowBPAmt = "" Then tTkLowBPAmt = "SALES"
                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  'space
                        Call oPrn.PrintText(tTkLowBPAmt, cX5, , vbLeftJustify, False)
                        'tTmp = Format(cW_LowBPAmt, tW_CurFmt) '*Tao 60-05-02 ของเดิม
                        tTmp = Format(cW_LowBPAmt, tW_CurFmt) & Space(14) '*Tao 60-05-02 เว้น 14 ตัว
                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False) '*Tao 60-04-27 ของเดิม แนวเดียวกัน
                    End If
                    '======================
                    
                    Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -
                     '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                     'TOTAL
                     'Total
                     If tTkTotal = "" Then tTkTotal = "TOTAL" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                    Select Case Val(tW_TransTypeSub)
                        Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnRedeem, nEN_TrnDepositCancel, nEN_TrnAddDepositCancel 'Redeem
                            Set oPrn.Font = oW_Font(1)
                            Call oPrn.PrintText(tTkTotal, cX5, , vbLeftJustify, False)
                            Set oPrn.Font = oW_Font(2)
                            tTmp = Format(cW_Total + cW_OnTopDps, tW_CurFmt)
''                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False, , , , True, True)              '*BG 57-04-29
                            If tW_TypePrinter <> "0" Then
                                Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False, , , , True, True)          '*BG 57-04-29
                            Else
                                Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                            End If
                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                        Case Else
                            Set oPrn.Font = oW_Font(1)
                            Call oPrn.PrintText(tTkTotal, cX5, , vbLeftJustify, False)
                            Set oPrn.Font = oW_Font(2)
'                            tTmp = Format(cW_Total + cW_Rnd, tW_CurFmt)
                            tTmp = Format(cW_Total, tW_CurFmt)                         '*Eaw 56-04-29   CommSheet TK-ISS3000-105-FSBT Answer.xlsx
'                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False, , , , True, True)
                            If tW_TypePrinter <> "0" Then
                                Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False, , , , True, True)          '*BG 57-04-29
                            Else
                                Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                            End If
                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                     End Select
                     
                    'item count
                    If nW_Items > 0 Then
                            Set oPrn.Font = oW_Font(1)
'                            tTmp = nW_Items & " " & tTkItem & "  " & tTkVatInc
                            '*Eaw 56-07-31CommSheet TK-ISS3000-283 TR FAPL-063.xls แยกคิด ManualTax ออกจากค่าอื่น
                            Select Case Val(tW_TransTypeSub)
                                Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax
                                    tTmp = nW_Items & " " & tTkItem & "  " & tTkVatExc
                                    If tTkItem = "" Then tTmp = nW_Items & " " & "Item" & "  " & "GST (Excl)" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                                Case tEN_TrnReturn '*KT 66-08-03 TTK-ISS3000-829
                                    If (cW_HDRate > 0) And (Not (bVB_Reference)) Then
                                         If InStr(1, tTkVatInc, CStr(cW_TaxRate)) > 0 Then
                                             tTkVatInc = Replace(tTkVatInc, CStr(cW_TaxRate), CStr(cW_HDRate))
                                        Else
                                            If W_GETxVocabularyRate(tTkVatInc) <> CStr(cW_HDRate) Then
                                                tTkVatInc = "GST-" & CStr(cW_HDRate) & "% (Incl)"
                                            End If
                                        End If
                                    End If
                                     tTmp = nW_Items & " " & tTkItem & "  " & tTkVatInc
                                    If tTkItem = "" Then tTmp = nW_Items & " " & "Item" & "  " & "GST (Incl)"
                                Case tEN_TrnAddDepositCancel, tEN_TrnRedeem, tEN_TrnAddDeposit
                                    If (cW_HDRate > 0) Then
                                         If InStr(1, tTkVatInc, CStr(cW_TaxRate)) > 0 Then
                                             tTkVatInc = Replace(tTkVatInc, CStr(cW_TaxRate), CStr(cW_HDRate))
                                        Else
                                            If W_GETxVocabularyRate(tTkVatInc) <> CStr(cW_HDRate) Then
                                                tTkVatInc = "GST-" & CStr(cW_HDRate) & "% (Incl)"
                                            End If
                                        End If
                                    End If
                                    tTmp = nW_Items & " " & tTkItem & "  " & tTkVatInc
                                    If tTkItem = "" Then tTmp = nW_Items & " " & "Item" & "  " & "GST (Incl)"
                                Case Else
                                    tTmp = nW_Items & " " & tTkItem & "  " & tTkVatInc
                                    If tTkItem = "" Then tTmp = nW_Items & " " & "Item" & "  " & "GST (Incl)" '*Eaw 56-10-31 CommSheet TK-ISS3000-409.xlsx
                            End Select
                            'Call oPrn.PrintText(tTmp, cXC, , vbCenter, False)
                            Call oPrn.PrintText(tTmp, cX5, , vbLeftJustify, False)
                             tTmp = Format(cW_Vat, tW_CurFmt)
                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                    End If
                     '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

        'Return
        W_PRNbTicketTotal = True
        
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        'Return
        W_PRNbTicketTotal = True
End Function

Private Function W_PRNbTicketTender() As Boolean
Dim nTmp As Long, tTmp As String
Dim tRcvName As String, cRcvAmt As Double
Dim tCardNo As String, tCardExp As String
Dim tSrcRetDocRef As String '*Tao 60-08-28 ฝากเก็บ TransID
Dim tRemark As String '*Tao 60-09-01 Remark
Dim tCustCode As String '*Tao 60-09-01 QR Code
Dim tReceiptText As String '*Tao 60-09-01 ReceiptText
Dim aReceiptText() As String '*Tao 60-09-01 ReceiptText
Dim nLine As Integer '*Tao 60-09-01
Dim bStaSend As Boolean '*Tao 60-09-06
Dim tWalletName As String '*Tao 60-09-08
Dim tEDCDetail As String '*Tao 62-07-22 EDC Detail
Dim aEDCdetail() As String '*Tao 62-07-22 EDC Detail Array
Dim tCardHolder As String '*Tao 62-07-22 Card Holder Name
Dim tRedeemCardNo As String '*Tao 62-07-22 Redeem Card Number
Dim tRedeemCardNoPrn As String '*Tao 62-07-22 Redeem Card Number Print
Dim tRedeemCardHolder As String '*Tao 62-07-22 Redeem Card Holder

Const tTkChn = "Change Tender"

'Const tTkPay = "PAYMENTS"
Dim tTkPay$
tTkPay = W_GETxVocabularyTicketFromArray("1953")
'Const tTkPrvPay = "PREVIOUS PAYMENTS"
Dim tTkPrvPay$
tTkPrvPay = W_GETxVocabularyTicketFromArray("1954")
'Const tTkPayAmt = "PAYMENTS AMOUNT"
Dim tTkPayAmt$
tTkPayAmt = W_GETxVocabularyTicketFromArray("1955")
'Const tTkPrvBalDue = "PREVIOUS BALANCE DUE"
Dim tTkPrvBalDue$
tTkPrvBalDue = W_GETxVocabularyTicketFromArray("1956")
'Const tTkDep = "DEPOSIT"
Dim tTkDep$
tTkDep = W_GETxVocabularyTicketFromArray("1957")
'Const tTkBalDue = "BALANCE DUE"
Dim tTkBalDue$
tTkBalDue = W_GETxVocabularyTicketFromArray("1958")
'Const tTkDueCust = "DUE CUSTOMER"
Dim tTkDueCust$
tTkDueCust = W_GETxVocabularyTicketFromArray("1959")
'Const tTkAmtDue = "AMOUNT DUE"
Dim tTkAmtDue$
tTkAmtDue = W_GETxVocabularyTicketFromArray("1960")
'Const tTkTotal = "TOTAL"
Dim tTkTotal$
tTkTotal = W_GETxVocabularyTicketFromArray("1948")
'Const tTkMiscRcv = "Misc Recv."
Dim tTkMiscRcv$
tTkMiscRcv = W_GETxVocabularyTicketFromArray("1947")
'Const tTkDsOnTop = "On Top"
Dim tTkDsOnTop$
tTkDsOnTop = W_GETxVocabularyTicketFromArray("1912")

Dim tTkAmtDueDate$
Dim tRcvCode$

tTkAmtDueDate = W_GETxVocabularyTicketFromArray("1943")

Dim tTkCooCpnID As String
tTkCooCpnID = "Cooperate Coupon No."

Dim tTkCooBrnID As String
tTkCooBrnID = "Cooperate Brand No."

Dim bPass As Boolean
Dim nI As Long

Dim tApproveCode As String

Dim cDiscVal As Double
Dim tTmpType As String

Dim bPrnRC  As Boolean '*KT 51-12-09 Phase 2 ข้อ 23  กรณีที่มีการ Redeem เป็น 0 จะไม่แสดงลงใน Receipt/EJ
Dim tStaNoPay As String

Dim tTkMiscTnd As String
tTkMiscTnd = W_GETxVocabularyTicketFromArray("9")
'Const tTkMiscTnd = "Misc Tender"
Dim tSpace  As String

'*Eaw 56-04-29   CommSheet TK-ISS3000-105-FSBT Answer.xlsx
Dim cRteAmt As Double
Dim tExchgRte As String
Const tTkExchg = "Ex."
Dim tTkAddDep As String
tTkAddDep = W_GETxVocabularyTicketFromArray("12")
'Const tTkAddDep = "ADD DEPOSIT"
Dim tTkAddDepCan As String
tTkAddDepCan = W_GETxVocabularyTicketFromArray("13")
'Const tTkAddDepCan = "ADD DEPOSIT CANCEL"
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
                            
                            tSpace = ""               '*Ao 2013/04/30  CommSheet TK-ISS3000-095 00.00.04.xls
                            Set oPrn.Font = oW_Font(1)
                            
                            Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)  ' space - - - - - -                    'deposit
                    
                            '- - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                            If bW_DepositMode Then
                            
                                    Select Case SP_VALnText2Double(tW_TransTypeSub)
                                    Case nEN_TrnDeposit
                                            'Line_Chars
                                            'deposit
                                            Call W_GETxBalRefDeposit
                                            Call oPrn.PrintText(Mid(tTkDep, 1, 18), cXC, , vbRightJustify, False)
                                            'Call oPrn.PrintText(tTkDep, cX5, , vbLeftJustify, False)
'                                            tTmp = Format(eW_Deposit.cEN_Deposit + cW_Rnd, tW_CurFmt)
'                                            tTmp = Format(cW_Total, tW_CurFmt)                         '*Eaw 56-04-29   CommSheet TK-ISS3000-105-FSBT Answer.xlsx
                                            tTmp = Format(eW_Deposit.cEN_Deposit, tW_CurFmt) '*Eaw 56-06-03 CommSheet TK-ISS3000-198-ST 00 00 01.xls
                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                    Case nEN_TrnRedeem
                                            ''Line_Chars
                                            'prev pay
                                            'Call oPrn.PrintText(tTkPrvPay, cX5, , vbLeftJustify, False)
                                            Call oPrn.PrintText(Mid(tTkPrvPay, 1, 18), cXC, , vbRightJustify, False)
                                            Call W_GETxBalRefDeposit '*Eaw 56-07-20 CommSheet TK-ISS3000-287 TR-FAPL-067.xls
                                            tTmp = Format(eW_Deposit.cEN_PevPayment, tW_CurFmt)
                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                            'Line_Chars
                                            'amt due
                                            'Call oPrn.PrintText(tTkAmtDue, cX5, , vbLeftJustify, False)
                                            Call oPrn.PrintText(Mid(tTkAmtDue, 1, 18), cXC, , vbRightJustify, False)
                                            tTmp = Format(eW_Deposit.cEN_PrevBalDue, tW_CurFmt)
                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                            'Line_Chars
                                            'prev bal due
                                            'Call oPrn.PrintText(tTkPrvBalDue, cX5, , vbLeftJustify, False)
'                                            Call oPrn.PrintText(Mid(tTkPrvBalDue, 1, 18), cXC, , vbRightJustify, False)  '*Eaw 56-06-20 ไม่พิมพ์ PreviousBalance Due
'                                            tTmp = Format(eW_Deposit.cEN_PrevBalDue, tW_CurFmt)
'                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
'                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                            '*KT 51-10-31 Communication Sheet ML-SF-568 Deposit Is Wrong.xls
                                            '*Eaw 56-10-256 ส่วนของ Due date ย้ายไปพิมพ์หลัง balance Due CommSheet TK-ISS3000-399-FAPL-144.xls
                                            'Due Date
'                                            Call oPrn.PrintText(Mid(tTkAmtDueDate, 1, 18), cXC, , vbRightJustify, False)
'
'                                            tTmp = Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt)
'
''                                            Call oPrn.PrintText(tTmp, cXR, , vbRightJustify, True)             '*Ao 2013/04/22  CommSheet TK-ISS3000-095 00.00.02.xls
'                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True)
                                            
'                                    Case nEN_TrnDepositCancel, nEN_TrnAddDepositCancel     '*Ao  2013/04/05  CommSheet TK-ISS3000-110 00.00.03 FSBT.xlsx
                                    Case nEN_TrnDepositCancel
                                            'Line_Chars
                                            'total
                                            '*Eaw 56-10-24 CommSheet TK-ISS3000-385-FAPL-144.xlsx ให้เอา Total ออกไปเลย ไม่ต้องพิมพ์ซ้ำ
'                                            'Call oPrn.PrintText(tTkTotal, cX5, , vbLeftJustify, False)
'                                            Call oPrn.PrintText(Mid(tTkTotal, 1, 18), cXC, , vbRightJustify, False)
''                                            tTmp = Format(cW_Total + cW_Rnd, tW_CurFmt)
'                                            tTmp = Format(cW_Total, tW_CurFmt)                         '*Eaw 56-04-29   CommSheet TK-ISS3000-105-FSBT Answer.xlsx
'                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
'                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
'                                            'Line_Chars
                                            'pay amt
                                            'Call oPrn.PrintText(tTkPayAmt, cX5, , vbLeftJustify, False)
'                                            Call oPrn.PrintText(Mid(tTkPayAmt, 1, 18), cXC, , vbRightJustify, False)
                                            Call oPrn.PrintText(Mid(tTkPay, 1, 18), cXC, , vbRightJustify, False)  '*Eaw 56-06-20 CommSheet TK-ISS3000-198-ST 00 00 03 FSBT NG.xlsx
'                                            tTmp = Format(eW_Deposit.cEN_Deposit + cW_Rnd, tW_CurFmt)
                                            tTmp = Format(eW_Deposit.cEN_Deposit, tW_CurFmt)                       '*Eaw 56-04-29   CommSheet TK-ISS3000-105-FSBT Answer.xlsx
                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                            '*Eaw 56-06-20 CommSheet TK-ISS3000-198-ST 00 00 03 FSBT NG.xlsx
                                            Call oPrn.PrintText(Mid(tTkPrvBalDue, 1, 18), cXC, , vbRightJustify, False)
                                            tTmp = Format(eW_Deposit.cEN_BalDue, tW_CurFmt)
                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
'                                    Case nEN_TrnAddDeposit
                                    Case nEN_TrnAddDeposit, nEN_TrnAddDepositCancel                 '*Ao  2013/04/05  CommSheet TK-ISS3000-110 00.00.03 FSBT.xlsx
                                            'Line_Chars
                                            'pay amt
                                            'Call oPrn.PrintText(tTkPay, cX5, , vbLeftJustify, False)
                                            Call W_GETxBalRefDeposit   '*Eaw 56-06-19 แก้ปัญหาเรื่อง Rounding deposit amount '*Eaw 56-06-03 CommSheet TK-ISS3000-198-ST 00 00 01.xls
                                            Call oPrn.PrintText(Mid(tTkPay, 1, 18), cXC, , vbRightJustify, False)
                                            tTmp = Format(eW_Deposit.cEN_PevPayment, tW_CurFmt)
                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                            'Line_Chars
                                            'prev bal due
                                            'Call oPrn.PrintText(tTkPrvBalDue, cX5, , vbLeftJustify, False)
                                            Call oPrn.PrintText(Mid(tTkPrvBalDue, 1, 18), cXC, , vbRightJustify, False)
                                            tTmp = Format(eW_Deposit.cEN_PrevBalDue, tW_CurFmt)
                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                            'Line_Chars
                                            'deposit
                                            'Call oPrn.PrintText(tTkDep, cX5, , vbLeftJustify, False)
'                                            Call oPrn.PrintText(Mid(tTkDep, 1, 18), cXC, , vbRightJustify, False)
                                            If SP_VALnText2Double(tW_TransTypeSub) = nEN_TrnAddDepositCancel Then   '*Eaw 56-06-03 CommSheet TK-ISS3000-198-ST 00 00 01.xls
                                                Call oPrn.PrintText(Mid(tTkAddDepCan, 1, 18), cXC, , vbRightJustify, False)
                                            Else
                                                Call oPrn.PrintText(Mid(tTkAddDep, 1, 18), cXC, , vbRightJustify, False)
                                            End If
'                                            tTmp = Format(eW_Deposit.cEN_Deposit + cW_Rnd, tW_CurFmt)
'                                            tTmp = Format(cW_Total, tW_CurFmt)                         '*Eaw 56-04-29   CommSheet TK-ISS3000-105-FSBT Answer.xlsx
                                            tTmp = Format(eW_Deposit.cEN_Deposit, tW_CurFmt)   '*Eaw 56-06-03 CommSheet TK-ISS3000-198-ST 00 00 01.xls
                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                            
'                                            If SP_VALnText2Double(tW_TransTypeSub) = nEN_TrnAddDepositCancel Then   '*Ao 2013/04/23 CommSheet TK-ISS3000-110.xls
'                                                tTmp = "CANCEL ADD DEPOSIT"
'                                                'CANCEL ADD DEPOSIT
'                                                Call oPrn.PrintText(tTmp, cXC, , vbRightJustify, True)
'                                            End If
                                        End Select
                                    
                            End If
                            
                    'due customer - - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

                                    
                            Select Case Val(tW_TransTypeSub)
                            Case nEN_TrnDepositCancel, nEN_TrnAddDepositCancel, nEN_TrnCancelVoucher        '*Ao 2013/04/05   CommSheet TK-ISS3000-110 00.00.03 FSBT.xlsx
'                                    Call oPrn.PrintText(tTkDueCust, cXC, , vbCenter, True) 'due customer
                            End Select
                            
                     'RCV- - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                            With orsRC
                            If Not orsRC.EOF Then .MoveFirst
                            Do While Not .EOF
                                    'Line_Chars
                                    'get var
                                    tRcvName = Mid(SP_FEDtChkString(orsRC, "FTSrcAbb"), 1, 30)
                                    cRcvAmt = SP_FEDvChkNumeric(orsRC, "FCSrcAmt")
                                    tCardNo = SP_FEDtChkString(orsRC, "FTSrcCardNo")
                                    tRcvCode = SP_FEDtChkString(orsRC, "FTTdmCode")
                                    tApproveCode = SP_FEDtChkString(orsRC, "FTApprCode")
                                    tStaNoPay = SP_FEDtChkString(orsRC, "FTSrcStaNoPay")
                                    cRteAmt = SP_FEDvChkNumeric(orsRC, "FCSrcRteAmt")
                                    tExchgRte = Format(SP_FEDvChkNumeric(orsRC, "FCSrcRteFac"), "#,##0.00000")
                                    tSrcRetDocRef = SP_FEDtChkString(orsRC, "FTSrcRetDocRef") '*Tao 60-08-28 ฝากเก็บ WPSTransID
                                     '*Tao 60-09-08 เช็คว่า ส่งไป web service หรือไม่
                                    bStaSend = IIf(SP_FEDtChkString(orsRC, "FTSrcStaPrc") = "Y", True, False)
                                    If bStaSend Then
                                        'Pay OnLine
                                        tRemark = SP_FEDtChkString(orsRC, "FTSrcRmk")
                                        If tRemark <> "" Then
                                            tWalletName = UT_tToken(tRemark, ";", 8, True)
                                            tCustCode = UT_tToken(tRemark, ";", 10, True)
                                            tReceiptText = UT_tToken(tRemark, ";", 11, True)
                                        End If
                                    End If
                                    
                                    '*Tao 62-07-22 Card Holder Name
                                    tEDCDetail = SP_FEDtChkString(orsRC, "FTEDCDetail") '*Tao 62-07-22 EDC Detail จะมี Tag98 ฝากมาด้วย
                                    If (Trim(tEDCDetail) <> "") Then
                                        aEDCdetail = Split(tEDCDetail, ",")
                                        For nI = 0 To UBound(aEDCdetail)
                                            If aEDCdetail(nI) <> "" Then
                                                If Left(aEDCdetail(nI), 2) = "98" Then
                                                    tCardHolder = Mid(aEDCdetail(nI), 3)
                                                    Exit For
                                                End If
                                            End If
                                        Next nI
                                    Else
                                        tCardHolder = ""
                                    End If
                                    
                                    If Trim(tStaNoPay) <> "1" Then           '*Ao 2011/01/25  Vendor Coupon   CommSheet ML-PRV-020-Change Tender is wrong concept.xls
                                        '*KT 51-12-09 Phase 2 ข้อ 23 กรณีที่มีการ Redeem เป็น 0 จะไม่แสดงลงใน Receipt/EJ
                                        '-------------------------------------------------------------------------------------------------------------------------
                                        Select Case SP_VALnText2Double(tW_TransTypeSub)
                                            Case nEN_TrnRedeem
                                                If cRcvAmt > 0 Then
                                                    bPrnRC = True
                                                Else
                                                    bPrnRC = False
                                                End If
                                            Case Else
                                                bPrnRC = True
                                        End Select
                                        '---------------------------------------------------------------------------------------------------------------------------
                                        
                                        '*Tao 62-07-22 ถ้าเป็น Credit Redeem เต็มจำนวน จะไม่มีรายการ Credit Card ให้พิมพ์ข้อมูลที่เก็บไว้แทน
                                        If (tRcvCode = "T002") And (Trim(tCardNo) <> tRedeemCardNo) And (tRedeemCardNo <> "") Then 'ถ้าคนละหมายเลขบัตร แสดงว่า Redeem เต็มจำนวน
                                            tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tRedeemCardNoPrn, 1, 38) & ";" 'Redeem CardNo + Expire + AppCode
                                            If (tRedeemCardHolder <> "") Then tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tRedeemCardHolder, 1, 38) & ";" 'Card Holder Name
                                            tRedeemCardNo = ""
                                            tTmp = vbNullString
                                            Select Case Val(tW_TransTypeSub)
                                                Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnRedeem
                                                    cDiscVal = SP_FEDvChkNumeric(orsRC, "FCSrcOntopAmt")
                                                    tTmpType = W_GETtCrdTypeByCrdID(tCardNo, 15)
                                                    If cDiscVal > 0 Then
                                                            tTmp = tTkDsOnTop & " " & tTmpType
                                                            If tTmp <> vbNullString Then
                                                                    'Line_Chars
                                                                   tW_SlipFooterStr = tW_SlipFooterStr & "#" & Mid(tTmp, 1, 40) & "{"
                                                                   tW_SlipFooterStr = tW_SlipFooterStr & "2" & Format(cDiscVal, tW_CurFmt) & "{"
                                                                   tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
                                                            End If
                                                    End If
                                            End Select
                                        ElseIf (tRcvCode = "T002") And (Trim(tCardNo) = tRedeemCardNo) And (cRcvAmt = 0) Then 'ถ้าหมายเลขบัตรเดียวกัน และยอดเงินเป้น 0 แสดงว่า Redeem เต็มจำนวน
                                            tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tRedeemCardNoPrn, 1, 38) & ";" 'Redeem CardNo + Expire + AppCode
                                            If (tRedeemCardHolder <> "") Then tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tRedeemCardHolder, 1, 38) & ";" 'Card Holder Name
                                            tRedeemCardNo = ""
                                            tTmp = vbNullString
                                            Select Case Val(tW_TransTypeSub)
                                                Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnRedeem
                                                    cDiscVal = SP_FEDvChkNumeric(orsRC, "FCSrcOntopAmt")
                                                    tTmpType = W_GETtCrdTypeByCrdID(tCardNo, 15)
                                                    If cDiscVal > 0 Then
                                                            tTmp = tTkDsOnTop & " " & tTmpType
                                                            If tTmp <> vbNullString Then
                                                                    'Line_Chars
                                                                   tW_SlipFooterStr = tW_SlipFooterStr & "#" & Mid(tTmp, 1, 40) & "{"
                                                                   tW_SlipFooterStr = tW_SlipFooterStr & "2" & Format(cDiscVal, tW_CurFmt) & "{"
                                                                   tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
                                                            End If
                                                    End If
                                            End Select
                                            GoTo NextRC 'ให้พิมพ์ Tender ถัดไปเลย
                                        End If
                                        
                                        If bPrnRC Then '*KT 51-12-09 Phase 2 ข้อ 23 กรณีที่มีการ Redeem เป็น 0 จะไม่แสดงลงใน Receipt/EJ
                                            'general tender
                                            Select Case tRcvCode
                                                Case "T002", "T003", "T004", "T005", "T006", "T007", "T008", "T009", "T010", "T011", "T014", "T017" 'mall/emp card '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher" '*Eaw 56-07-11 เพิ่ม "T017" TC card
'                                                    Call oPrn.PrintText(Trim(tRcvName), cXL, , vbLeftJustify, False)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & Trim(tRcvName) & "{"
                                                Case "T018" '*KT 55-08-28 Takashimaya CV
'                                                    Call oPrn.PrintText(Trim(tRcvName), cXL, , vbLeftJustify, False)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & Trim(tRcvName) & "{"
                                                Case "T012"  '*Eaw 56-04-29   CommSheet TK-ISS3000-105-FSBT Answer.xlsx
                                                    tTmp = Trim(tRcvName) & " " & Format(cRteAmt, tW_CurFmt) & " " & tTkExchg & tExchgRte         '*Ao ลดช่องว่างลง  '*Eaw 56-06-04 CommSheet TK-ISS3000-173-ST 00 00 03 FSBT NG.xls เอาช่องว่างระหว่าง tTkExchg กับ tExchgRte  ออก
'                                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                                                Case "T015", "T013", "T016" '*Eaw 56-06-04 CommSheet TK-ISS3000-223-ST 00 00 01.xls '*Eaw 56-09-02 add T013,T016
'                                                    Call oPrn.PrintText(Mid(Trim(tRcvName & " " & tApproveCode), 1, 30), cXL, , vbLeftJustify, False)
                                                    tTmp = Mid(Trim(tRcvName & " " & tApproveCode), 1, 30)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                                                Case "T019", "T020", "T021", "T022" '*Tao 60-08-28 WalletPay, WeChatPay, BluePay, Others
                                                    '*Tao 60-09-08 ถ้า Pay OnLine ให้เอาชื่อที่ WPA ส่งมาไปพิมพ์
                                                    If bStaSend Then
                                                        tW_SlipFooterStr = tW_SlipFooterStr & "8" & tWalletName & "{"  '*Tao 60-09-11 ถ้ามีชื่อ Wallet ส่งมา ให้เอาไปพิมพ์
                                                    Else
                                                        tW_SlipFooterStr = tW_SlipFooterStr & "8" & tRcvName & "{"   '*Tao 60-09-11 ถ้าไม่มีชื่อ Wallet ส่งมา ให้เอา Tender Name ไปพิมพ์
                                                    End If
                                                Case Else
                                                    'Line_Chars
'                                                    Call oPrn.PrintText(Mid(Trim(tRcvName & " " & tCardNo), 1, 20), cXL, , vbLeftJustify, False)
                                                    tTmp = Mid(Trim(tRcvName & " " & tCardNo), 1, 30)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                                            End Select
                                            Select Case tRcvCode
                                                Case "T018"
                                                    'If bW_EnableRePrint Then
                                                        Call W_GETbCVIssueEx(tCardNo)
                                                        tTmp = Format(oW_CV.cEN_CVAmt, tW_CurFmt)
                                                    'Else
                                                    '    tTmp = Format(cVB_CNCVAmt, tW_CurFmt)
                                                    'End If
                                                Case Else
                                                    tTmp = Format(cRcvAmt, tW_CurFmt)
                                            End Select
'                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "2" & tTmp & "{"
'                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
                                            
                                            'card tender
                                            Select Case tRcvCode
                                                Case "T002", "T003", "T004", "T005", "T006", "T007", "T008", "T009", "T010", "T011", "T014", "T017" '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher" '*Eaw 56-07-11 เพิ่ม "T017" TC card
                                                    If tCardNo <> vbNullString Then
                                                            Select Case tRcvCode
                                                                Case "T002", "T003", "T017"  '*Eaw 56-07-11 เพิ่ม "T017" TC card
                                                                    'Line_Chars
                                                                    '*Bump 58-02-12 CommSheet TK-ISS3000-578.xls
                                                                    '* bVB_Mark_X_CRexp = true คือ วันที่บัตรเครดิตหมดอายุ จะถูก mark เป็น XX/XX ตอนพิมพ์ และ EJ
                                                                    '*KT 67-01-08 CR Mobile Voucher
                                                                    If UCase(SP_FEDtChkString(orsRC, "FTREFNO")) = UCase("BPVoucher") And tRcvCode = "T002" Then
                                                                        tCardExp = SP_FEDtChkString(orsRC, "FTSrcCardExp")
                                                                         tTmp = SP_FEDtChkString(orsRC, "FTSrcCardNo")
                                                                    Else
                                                                        If bVB_Mark_X_CRexp = True Then
                                                                            tCardExp = "XX/XX"
                                                                        Else
                                                                            tCardExp = SP_FEDtChkString(orsRC, "FTSrcCardExp")
                                                                        End If
                                                                         tTmp = W_CHKtChangeCardID(SP_FEDtChkString(orsRC, "FTSrcCardNo"), tRcvCode)     '*Ao 09-08-01: PH2-ML-SF- 088
                                                                    End If
                                                                   
'                                                                     tTmp = tCardNo & "  " & tCardExp & "  " & tApproveCode
                                                                     tTmp = tTmp & " " & tCardExp & " " & tApproveCode '*Eaw 56-07-06
'                                                                    Call oPrn.PrintText(Mid(tTmp, 1, 40), cX2, , vbLeftJustify, True)

                                                                    '*Tao 62-07-24 ของเดิม (ถ้าเป็น Credit Redeem ยังไม่ต้องพิมพ์)
                                                                    'tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                                    
                                                                    '*Tao 62-07-22 ถ้าเป็น Credit Redeem ให้เก็บข้อมูลบัตรไว้ก่อน
                                                                    If (tRcvCode = "T002") And (Trim(tRcvName) = tVB_CreditRedeemName) And ((orsRC.RecordCount - orsRC.AbsolutePosition) > 0) Then
                                                                         'tRedeemCardNoPrn = W_CHKtChangeCardID(tCardNo, tRcvCode) '*Tao 62-07-22 Redeem Card Number Mark X
                                                                        tRedeemCardHolder = tCardHolder '*Tao 62-07-22 Redeem Card Holder
                                                                        tRedeemCardNo = tCardNo '*Tao 62-07-22 Redeem Card No
                                                                        tRedeemCardNoPrn = Mid(tTmp, 1, 38)  '*Tao 62-07-22 Redeem Card Number
'                                                                        '*KT 66-12-08  CR Bank Instant Point Redemption
                                                                        'ของเดิมมีการเก็บ T002  ปกติ Credit Card   แบบ Redeem ทั้งหมด หรือ Redeem อย่างเดี่ยวเป็น 0 เสมอ
'                                                                        'หาว่ามีรายการ T002 ที่จ่ายร่วมหรือไม่
'                                                                        Dim orsRCTmp As ADODB.Recordset
'                                                                        Set orsRCTmp = orsRC.Clone
'                                                                        orsRCTmp.Filter = "FTSrcCardNo='" & "tCardNo" & "' AND FTTdmCode='T002' AND FTVendorCnpFlag<>'1'"
'                                                                        If (orsRCTmp.EOF And orsRCTmp.EOF) Then
'                                                                                tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tRedeemCardNoPrn, 1, 38) & ";" 'Redeem CardNo + Expire + AppCode
'                                                                                If (tRedeemCardHolder <> "") Then tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tRedeemCardHolder, 1, 38) & ";" 'Card Holder Name
'                                                                        End If
                                                                    Else
                                                                        tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                                        If tCardHolder <> "" Then tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tCardHolder, 1, 38) & ";" 'Card Holder Name
                                                                    End If
                                                                    '=============================================
                                                                
                                                                Case "T014"
                                                                    tCardExp = SP_FEDtChkString(orsRC, "FTSrcCardExp")                      '*Ao 2014/10/22    eCoupon   V 4.2.152
'                                                                    tTmp = "ECOUPON          " & tCardNo
                                                                    'tTmp = "ECOUPON           " & tCardNo
                                                                    tTmp = "ECOUPON" & "         " & tCardNo '*Bump 57-11-13 ปรับ เป็น 9 Space  ใช้ font 'Courier MonoThai'  เทียบดูใน notepad V.4.2.156
                                                                    tW_SlipFooterStr = tW_SlipFooterStr & "0" & tTmp & ";"
                                                                    tTmp = "EXPIRY DATE     " & tCardExp
                                                                    tW_SlipFooterStr = tW_SlipFooterStr & "0" & tTmp & ";"
                                                                Case Else
                                                                    'Line_Chars
                                                                    tTmp = tCardNo
'                                                                    Call oPrn.PrintText(Mid(tTmp, 1, 40), cX2, , vbLeftJustify, True)
                                                                    tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                            End Select
                                                    End If
                                                Case "T019", "T020", "T021", "T022" '*Tao 60-08-28 WalletPay, WeChatPay, BluePay, Others
                                                    Select Case SP_VALnText2Double(tW_TransTypeSub)
                                                    Case tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnIPVTaxExceptSale, tEN_TrnCancelVoucher, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVVIPService, tEN_TrnIPVReturn, tEN_TrnIPVRedeem, tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc
'                                                        If Trim(tSrcRetDocRef) <> vbNullString Then '*Tao 60-08-28 พิมพ์ WPSTransID
'                                                            tTmp = "Partner Trans ID: " & Trim(tSrcRetDocRef)
'                                                            tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
'                                                        End If
                                                        If Trim(tCardNo) <> vbNullString Then '*Tao 60-08-28 พิมพ์ PartnerTransID
                                                            'tTmp = "PartnerTransID: " & Trim(tCardNo) '*Tao 61-08-02 ของเดิม
                                                            
                                                            '*Tao 61-08-09 ของเดิม
                                                            'tTmp = "Reference: " & Trim(tCardNo) '*Tao 61-08-02 เปลี่ยนเป็น Reference
                                                            'tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                            
                                                            '*Tao 61-08-09 ย้ายไปบันทัดถัดไป เพื่อให้พิมพ์ได้ครบ 30 ตัวเลข
                                                            'tTmp = "Reference: " '*Tao 61-08-02 เปลี่ยนเป็น Reference
                                                            tTmp = "WPS TRANS ID: " '*Tao 61-08-20 เปลี่ยนเป็น WPS TRANS ID (CommSheet TTK-ISS3000-693)
                                                            If Len(Trim(tCardNo)) > 24 Then
                                                                tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                                tTmp = Trim(tCardNo)
                                                                tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                            Else
                                                                tTmp = tTmp & Trim(tCardNo)
                                                                tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                            End If
                                                            
                                                        End If
                                                    Case tEN_TrnSale, tEN_TrnDeposit, tEN_TrnAddDeposit, tEN_TrnRedeem, tEN_TrnTaxExceptSale, tEN_TrnVIPSale, tEN_TrnVoucher, tEN_TrnDutyFreeSale, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc
                                                        If bStaSend Then
                                                            'Pay OnLine
                                                            If Trim(tCustCode) <> vbNullString Then '*Tao 60-08-28 พิมพ์ Customer QR Code (Token)
                                                                tTmp = "Buyer Identity Code: " & Trim(tCustCode)
                                                                tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                            End If
'                                                            '*Tao 60-09-07 ไม่ต้องแสดง เพราะใน ReceiptText มีครบแล้ว
'                                                            If Trim(tCardNo) <> vbNullString Then '*Tao 60-08-28 พิมพ์ PartnerTransID
'                                                                tTmp = "Partner TransID: " & Trim(tCardNo)
'                                                                tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
'                                                            End If
'                                                            If Trim(tSrcRetDocRef) <> vbNullString Then '*Tao 60-08-28 พิมพ์ WPSTransID
'                                                                tTmp = "TransID: " & Replace(Trim(tSrcRetDocRef), "-", "")
'                                                                tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
'                                                            End If
                                                            If Trim(tReceiptText) <> vbNullString Then '*Tao 60-09-01 พิมพ์ ReceiptText
                                                                aReceiptText() = Split(tReceiptText, vbLf)
                                                                For nLine = 0 To UBound(aReceiptText)
                                                                    tTmp = Trim(aReceiptText(nLine))
                                                                    tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                                Next nLine
                                                            End If
                                                        Else
                                                            'Pay OffLine
                                                            If Trim(tCardNo) <> vbNullString Then '*Tao 60-08-28 พิมพ์ PartnerTransID
                                                                'tTmp = "PartnerTransID: " & Trim(tCardNo) '*Tao 61-08-02 ของเดิม
                                                                '*Tao 61-08-09 ของเดิม
                                                                'tTmp = "Reference: " & Trim(tCardNo) '*Tao 61-08-02 เปลี่ยนเป็น Reference
                                                                'tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                                
                                                                '*Tao 61-08-09 ย้ายไปบันทัดถัดไป เพื่อให้พิมพ์ได้ครบ 30 ตัวเลข
                                                                'tTmp = "Reference: " '*Tao 61-08-02 เปลี่ยนเป็น Reference
                                                                tTmp = "WPS TRANS ID: " '*Tao 61-08-20 เปลี่ยนเป็น WPS TRANS ID (CommSheet TTK-ISS3000-693)
                                                                If Len(Trim(tCardNo)) > 24 Then
                                                                    tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                                    tTmp = Trim(tCardNo)
                                                                    tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                                Else
                                                                    tTmp = tTmp & Trim(tCardNo) '*Tao 61-08-02 เปลี่ยนเป็น Reference
                                                                    tW_SlipFooterStr = tW_SlipFooterStr & "0" & Mid(tTmp, 1, 38) & ";"
                                                                End If
                                                            End If
                                                        End If
                                                    End Select
                                            End Select
NextRC:
                                            '*Tao 62-07-22 ถ้าเป็น Credit Redeem ยังไม่ต้องพิมพ์
                                            If (Trim(tRcvName) <> tVB_CreditRedeemName) Or ((orsRC.RecordCount - orsRC.AbsolutePosition) = 0) Then
                                            
                                                'Deposit ,Add Deposit
                                                'Communication Sheet ML-SF-478 Error on Deposit Transaction.xls
                                                '-------------------------------------------------------------------------------------------------------
                                                tTmp = vbNullString
                                                Select Case tRcvCode
                                                  Case "T002", "T003", "T017" '*Eaw 56-07-11 เพิ่ม "T017" TC card
                                                        Select Case Val(tW_TransTypeSub)
                                                            Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnRedeem
                                                                cDiscVal = SP_FEDvChkNumeric(orsRC, "FCSrcOntopAmt") 'W_GETcDiscByCrdID(tCardNo, 15) 'อ่าน Ontop จาก Field ใน RC ไม่ต้องไปที่ CD แล้ว
                                                                tTmpType = W_GETtCrdTypeByCrdID(tCardNo, 15)
                                                                If cDiscVal > 0 Then
                                                                        tTmp = tTkDsOnTop & " " & tTmpType
                                                                        If tTmp <> vbNullString Then
                                                                                'Line_Chars
    '                                                                           Call oPrn.PrintText(Mid(tTmp, 1, 40), cX5, , vbLeftJustify, False) '*Eaw 56-08-22 เปลี่ยนจาก 17 เป็น 40
                                                                               tW_SlipFooterStr = tW_SlipFooterStr & "#" & Mid(tTmp, 1, 40) & "{"
    '                                                                           Call oPrn.PrintText(Format(cDiscVal, tW_CurFmt), cXMK, , vbRightJustify, False)
                                                                               tW_SlipFooterStr = tW_SlipFooterStr & "2" & Format(cDiscVal, tW_CurFmt) & "{"
    '                                                                           Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                                                               tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
                                                                        End If
                                                                End If
                                                        End Select
                                                 End Select
                                                '-------------------------------------------------------------------------------------------------------
                                            End If
                                            
                                            '*KT 55-08-28 Takashimaya PRJ10001TK-External Specification Credit Voucher-00.04.00
                                            '-------------------------------------------------------------------------------------------------------
                                            Select Case SP_VALnText2Double(tW_TransTypeSub)
                                                Case nEN_TrnReturn, nEN_TrnReturnTaxInc, nEN_TrnReturnManualTax, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel 'TK-ISS3000-037
                                                     Select Case tRcvCode
                                                            Case "T018"
                                                                    tTmp = "CREDIT VCR" & "      " & tCardNo '*Eaw 56-08-07 ขยับช่องว่างเพิ่มอีก 3 ช่อง
'                                                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                                                     tW_SlipFooterStr = tW_SlipFooterStr & "0" & tTmp & ";"
                                                                     tCardExp = SP_FEDtChkString(orsRC, "FTSrcCardExp")
                                                                     tTmp = "EXPIRY DATE" & "     " & tCardExp '*Eaw 56-08-07 ขยับช่องว่างเพิ่มอีก 3 ช่อง
'                                                                     Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                                                     tW_SlipFooterStr = tW_SlipFooterStr & "0" & tTmp & ";"
                                                                    '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
                                                                    'tTmp = "CONTACT NUMBER" & "  " & SP_FEDtChkString(orsRC, "FTSrcRetDocRef")
                                                                    'Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                                                     bW_StaPrnCV2 = True
                                                                     bW_StaPrnCV3 = False
                                                     End Select
                                                Case nEN_TrnIPV, nEN_TrnIPVDutyFreeSale, nEN_TrnIPVRedeem, nEN_TrnIPVReturn, nEN_TrnIPVTaxExceptSale, nEN_TrnIPVVIPService
                                                    Select Case tRcvCode
                                                            Case "T018"
                                                                     bW_StaPrnCV2 = W_GETbCVIssue(tW_TransType, tW_TmnNum, tW_TransNo)
                                                                     bW_StaPrnCV3 = False
                                                     End Select
                                                Case Else
                                                     Select Case tRcvCode
                                                            Case "T018"
                                                                    tTmp = "CREDIT VCR" & "      " & tCardNo '*Eaw 56-08-07 ขยับช่องว่างเพิ่มอีก 3 ช่อง
'                                                                    Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                                                    tW_SlipFooterStr = tW_SlipFooterStr & "0" & tTmp & ";"
                                                                     tCardExp = SP_FEDtChkString(orsRC, "FTSrcCardExp")
                                                                     tTmp = "EXPIRY DATE" & "     " & tCardExp '*Eaw 56-08-07 ขยับช่องว่างเพิ่มอีก 3 ช่อง
'                                                                     Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                                                    tW_SlipFooterStr = tW_SlipFooterStr & "0" & tTmp & ";"
                                                                    '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
                                                                    'tTmp = "CONTACT NUMBER" & "  " & SP_FEDtChkString(orsRC, "FTSrcRetDocRef")
                                                                    'Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                                                     bW_StaPrnCV2 = W_GETbCVIssue(tW_TransType, tW_TmnNum, tW_TransNo)
                                                                     bW_StaPrnCV3 = True

                                                     End Select
                                            End Select
                                            '-------------------------------------------------------------------------------------------------------
                                            
                                        End If
                                    End If
                                    .MoveNext
                            Loop
                            End With
                            
                            'TK-ISS3000-037
                            Select Case SP_VALnText2Double(tW_TransTypeSub)
                                Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                                    'none
                                Case Else
                                    'bW_StaPrnCVReIssue Then '*Thirapong 2021-03-27
                                    If oW_ReIssue.cEN_ReIssueAmt > 0 And oW_ReIssue.tEN_ReIssueNo <> "" Then
                                        bW_StaPrnCVReIssue = True
                                        Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ":W_PRNbTicketTender :tEN_ReIssueNo :" & oW_ReIssue.tEN_ReIssueNo & ",cEN_ReIssueAmt :" & oW_ReIssue.cEN_ReIssueAmt & ",bW_StaPrnCV2:" & bW_StaPrnCV2)
                                    Else
                                        bW_StaPrnCVReIssue = False
                                        Call SP_CRTxLogPrg("SLIP PRINTER-cTransaction.PrintOut-TransType:" & tW_TransTypeSub & ":W_PRNbTicketTender :tEN_ReIssueNo :" & oW_ReIssue.tEN_ReIssueNo & ",cEN_ReIssueAmt :" & oW_ReIssue.cEN_ReIssueAmt & ",bW_StaPrnCV2:" & bW_StaPrnCV2)
                                    End If
                                    'If bW_StaPrnCV2 Then
                                    If bW_StaPrnCV2 And bW_StaPrnCVReIssue Then '*Thirapong 2021-03-27
                                        If SP_VALnText2Double(tW_TransTypeSub) <> nEN_TrnReturn And SP_VALnText2Double(tW_TransTypeSub) <> nEN_TrnReturnTaxInc And SP_VALnText2Double(tW_TransTypeSub) <> nEN_TrnReturnManualTax Then
'                                            Call oPrn.PrintText("", cXL, , vbLeftJustify, True)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "0" & "" & ";"
                                            tTmp = "Re-issue"
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "0" & tTmp & ";"
                                            tTmp = "Cr vchr"
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, False)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                                            tTmp = Format(oW_ReIssue.cEN_ReIssueAmt, tW_CurFmt)
'                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, True)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "$" & tTmp & ";"
                                            tTmp = "CREDIT VCR" & "      " & oW_ReIssue.tEN_ReIssueNo '*Eaw 56-08-07 ขยับช่องว่างเพิ่มอีก 3 ช่อง
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True) '*Eaw 56-03-02  TK-ISS3000-069
                                            tW_SlipFooterStr = tW_SlipFooterStr & "0" & tTmp & ";"
                                            tTmp = "EXPIRY DATE" & "     " & oW_ReIssue.tEN_ReIssueExp '*Eaw 56-08-07 ขยับช่องว่างเพิ่มอีก 3 ช่อง
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True) '*Eaw 56-03-02  TK-ISS3000-069
                                            tW_SlipFooterStr = tW_SlipFooterStr & "0" & tTmp & ";"
                                            '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
'                                            tTmp = "CONTACT NUMBER" & "  " & oW_ReIssue.tEN_ReIssueTel
'                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True)
                                        End If
                                End If
                            End Select
                            
                            'Misc Recv
                            Select Case Val(tW_TransTypeSub)
                                Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnRedeem, nEN_TrnDepositCancel, nEN_TrnAddDepositCancel '*KT 51-11-13
                                    'Misc Recv
                                    If cW_Rnd <> 0 Then
'                                            Call oPrn.PrintText(tTkMiscRcv, cX5, , vbLeftJustify, False)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "#" & tTkMiscRcv & "{"
                                            tTmp = Format(cW_Rnd, tW_CurFmt)
'                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "2" & tTmp & "{"
'                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
                                    End If
                            End Select
                             
                            'change
                            If cW_Chg > 0 Then
                            '*Eaw 56-02-26 ไม่พิมพ์ Change tender เมื่อมีการพิมพ์ CV Re-Issue TK-ISS3000-067
'                                        Call oPrn.PrintText(tTkChn, cXL, , vbLeftJustify, False)
'                                        tTmp = Format(cW_Chg, tW_CurFmt)
'                                        Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
'                                        Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                        Select Case tRcvCode
                                               Case "T018"
                                                    'none
                                                Case Else
'                                                    Call oPrn.PrintText(tTkChn, cXL, , vbLeftJustify, False)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTkChn & "{"
                                                    tTmp = Format(cW_Chg, tW_CurFmt)
'                                                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "2" & tTmp & "{"
'                                                    Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
                                        End Select
                                        '----------------------------------------------------------------------------------------------
                            End If
                            
                            'Misc Tender  '*Eaw 56-04-29   CommSheet TK-ISS3000-105-FSBT Answer.xlsx
                            Select Case Val(tW_TransTypeSub)
                                Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnRedeem, nEN_TrnDepositCancel, nEN_TrnAddDepositCancel '*KT 51-11-13
                                    'none
                                Case Else
        '                            If cW_MiscTnd <> 0 Then
                                    If cW_Rnd <> 0 Then
'                                            Call oPrn.PrintText(tTkMiscTnd, cXL, , vbLeftJustify, False)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTkMiscTnd & "{"
                                            tTmp = Format(cW_Rnd, tW_CurFmt)
'                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "2" & tTmp & "{"
'                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                            tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
                                    End If
                            End Select
                            '- - - - - - - -- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                            If bW_DepositMode Then
                                    Select Case Val(tW_TransTypeSub)
'                                    Case nEN_TrnDeposit, nEN_TrnRedeem, nEN_TrnAddDeposit                      '*Ao 2013/04/05   CommSheet TK-ISS3000-110 00.00.03 FSBT.xlsx
                                    Case nEN_TrnDeposit, nEN_TrnRedeem, nEN_TrnAddDeposit, nEN_TrnAddDepositCancel, nEN_TrnDepositCancel '*Eaw 56-10-28 เพิ่ม nEN_TrnDepositCancel
                                            If eW_Deposit.cEN_BalDue >= 0 Then
                                                    ' bal due
                                                    'Call oPrn.PrintText(tTkBalDue, cXC + cXHC, , vbRightJustify, False)
'                                                    Call oPrn.PrintText(tTkBalDue, cX5, , vbLeftJustify, False)
'                                                    Call oPrn.PrintText(Mid(tTkBalDue, 1, 18), cXC, , vbRightJustify, False)  '*Eaw 56-06-03 CommSheet TK-ISS3000-198-ST 00 00 01.xls
'                                                    tW_SlipFooterStr = tW_SlipFooterStr & "9" & Mid(tTkBalDue, 1, 18) & ","
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "8" & "          " & Mid(tTkBalDue, 1, 18) & "{"
'                                                    tTmp = Format(eW_Deposit.cEN_BalDue, tW_CurFmt)
                                                    Select Case Val(tW_TransTypeSub)                                                                            '*Ao 2014/07/22   V 4.2.139
                                                        Case nEN_TrnDeposit
                                                            If nW_PrintCnt > 0 Or bW_PRMGenEJ Then
                                                                tTmp = Format(eW_Deposit.cEN_Reprint, tW_CurFmt)
                                                            Else
                                                                tTmp = Format(eW_Deposit.cEN_BalDue, tW_CurFmt)
                                                            End If
                                                        Case Else
                                                            tTmp = Format(eW_Deposit.cEN_BalDue, tW_CurFmt)
                                                    End Select
                                                    
'                                                    Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "2" & tTmp & "{"
'                                                    Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
                                                    tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
                                                    'due date
                                                    Select Case Val(tW_TransTypeSub)
'                                                       Case nEN_TrnDeposit, nEN_TrnAddDeposit             '*Ao 2013/04/05   CommSheet TK-ISS3000-110 00.00.03 FSBT.xlsx
                                                        Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnAddDepositCancel, nEN_TrnRedeem, nEN_TrnDepositCancel '*Eaw 56-10-26 เพิ่ม nEN_TrnRedeem,nEN_TrnDepositCancel CommSheet TK-ISS3000-399-FAPL-144.xls
                                                            'Call oPrn.PrintText(tTkAmtDueDate, cXC + cXHC, , vbRightJustify, False)
'                                                           Call oPrn.PrintText(tTkAmtDueDate, cX5, , vbLeftJustify, False)
'                                                            Call oPrn.PrintText(Mid(tTkAmtDueDate, 1, 18), cXC, , vbRightJustify, False)  '*Eaw 56-06-03 CommSheet TK-ISS3000-198-ST 00 00 01.xls
'                                                            tW_SlipFooterStr = tW_SlipFooterStr & "9" & Mid(tTkAmtDueDate, 1, 18) & "," '*Eaw 56-10-11 เอาไปรวมกับวันที่
'                                                           Call oPrn.PrintText(Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt), cXR, , vbRightJustify, True)    '*Ao 2013/04/22  CommSheet TK-ISS3000-095 00.00.02.xls
'                                                            Call oPrn.PrintText(Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt), cXMK, , vbRightJustify, True)
'                                                            tW_SlipFooterStr = tW_SlipFooterStr & "10" & Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt) & ";" '*Eaw 56-10-11 เอาไปรวมกับวันที่
                                                            tTmp = "          " & Mid(tTkAmtDueDate, 1, 18)
                                                            tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                                                            tTmp = Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt)
                                                            tW_SlipFooterStr = tW_SlipFooterStr & "$" & tTmp & ";"
                                                    End Select
                                            Else
                                                     '*KT 52-04-17 Communication Sheet PH1.5-ML-SF-125 Deposit full payment not print due date.xls
                                                     'ในรายการ Deposit/Add Deposit  เมื่อมีการจ่ายเงินเต็มจำนวนโปรแกรมไม่พิมพ์ Duedate ต้องการให้พิมพ์ Duedate ด้วย
                                                     'due date
                                                    Select Case Val(tW_TransTypeSub)
'                                                        Case nEN_TrnDeposit, nEN_TrnAddDeposit                  '*Ao 2013/04/05   CommSheet TK-ISS3000-110 00.00.03 FSBT.xlsx
                                                        Case nEN_TrnDeposit, nEN_TrnAddDeposit, nEN_TrnAddDepositCancel, nEN_TrnDepositCancel '*Eaw 56-10-28 เพิ่ม nEN_TrnDepositCancel
                                                            'Call oPrn.PrintText(tTkAmtDueDate, cXC + cXHC, , vbRightJustify, False)
'                                                            Call oPrn.PrintText(tTkAmtDueDate, cX5, , vbLeftJustify, False)
'                                                            tW_SlipFooterStr = tW_SlipFooterStr & "10" & tTkAmtDueDate & "," '*Eaw 56-10-11 เอาไปรวมกับวันที่
'                                                            Call oPrn.PrintText(Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt), cXR, , vbRightJustify, True)    '*Ao 2013/04/22  CommSheet TK-ISS3000-095 00.00.02.xls
'                                                            Call oPrn.PrintText(Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt), cXMK, , vbRightJustify, True)
'                                                            tW_SlipFooterStr = tW_SlipFooterStr & "11" & Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt) & ";" '*Eaw 56-10-11 เอาไปรวมกับวันที่
'                                                            tTmp = Mid(tTkAmtDueDate, 1, 18) & "   " & Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt)
'                                                            tW_SlipFooterStr = tW_SlipFooterStr & "$" & tTmp & ";"


                                                            tTmp = "          " & Mid(tTkAmtDueDate, 1, 18)
                                                            tW_SlipFooterStr = tW_SlipFooterStr & "8" & tTmp & "{"
                                                            tTmp = Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt)
                                                            tW_SlipFooterStr = tW_SlipFooterStr & "$" & tTmp & ";"
                                                            
                                                        End Select
                                            End If
                                    '*Eaw 56-10-28 เอาไปรวมกับเคสอื่น
'                                    Case nEN_TrnDepositCancel '*Eaw 56-06-20 CommSheet TK-ISS3000-198-ST 00 00 03 FSBT NG.xlsx เพิ่ม nEN_TrnDepositCancel
'                                            'bal due
''                                            Call oPrn.PrintText(tTkBalDue, cX5, , vbLeftJustify, False)
'                                            tW_SlipFooterStr = tW_SlipFooterStr & "#" & tTkBalDue & "{"
'                                            tTmp = Format(0, tW_CurFmt)
''                                            Call oPrn.PrintText(tTmp, cXMK, , vbRightJustify, False)
'                                            tW_SlipFooterStr = tW_SlipFooterStr & "2" & tTmp & "{"
''                                            Call oPrn.PrintText(tW_TrMk, cXMK, , vbLeftJustify, True)
'                                            tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
'                                            '*Eaw 56-10-26 เพิ่มส่วนของ Due date CommSheet TK-ISS3000-399-FAPL-144.xls
'                                            'Due date
'
'                                            tTmp = Mid(tTkAmtDueDate, 1, 18)
'                                            tW_SlipFooterStr = tW_SlipFooterStr & "#" & tTmp & "{"
'                                            tTmp = Format(eW_Deposit.cEN_BalDueDate, tW_DateFmt)
'                                            tW_SlipFooterStr = tW_SlipFooterStr & "2" & tTmp & "{"
'                                            tW_SlipFooterStr = tW_SlipFooterStr & "3" & tW_TrMk & ";"
                                    End Select
                            End If
                            
                            'Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True)   ' space - - - - - -Follow by format Receipt/EJ
                            
                            '*KT 51-07-15 Comsheet ML-SF-372
                            '==================================================================

                            With orsCD
                                If Not orsCD Is Nothing And orsCD.State = adStateOpen Then
                                ReDim aW_CooCpnID(0) 'clear array
                                .Filter = "FNDctNo = 9  OR  FNDctNo = 10" 'CooCpn,CooBrn
                                .Sort = "FNScdSeqNo ASC"
                                If Not .EOF Then
                                    Do While Not .EOF
                                        tTmp = SP_FEDtChkString(orsCD, "FTScdCardID")
                                        nTmp = SP_FEDvChkNumeric(orsCD, "FNDctNo")
                                        bPass = True
                                        For nI = 0 To UBound(aW_CooCpnID) - 1
                                            If aW_CooCpnID(nI) = tTmp Then bPass = False: Exit For
                                        Next nI
                                        aW_CooCpnID(UBound(aW_CooCpnID)) = tTmp
                                        ReDim Preserve aW_CooCpnID(UBound(aW_CooCpnID) + 1)
                                        If bPass Then
                                            If tTmp <> "" Then
                                                Select Case nTmp
                                                    Case 9
                                                        tTmp = tTkCooCpnID & " " & tTmp
                                                    Case 10
                                                        tTmp = tTkCooBrnID & " " & tTmp
                                                End Select
'                                                Call oPrn.PrintText(Mid(tTmp, 1, nVB_MaxChar), cXL, , vbLeftJustify, True)
                                                tW_SlipFooterStr = tW_SlipFooterStr & "1" & Mid(tTmp, 1, nVB_MaxChar) & ";"
                                            End If
                                        End If
                                        .MoveNext
                                    Loop
                                End If
                                .Filter = vbNullString
                            End If
                            End With
                            
                            '============================================================
        'Return
        W_PRNbTicketTender = True
        
        Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        'Return
        W_PRNbTicketTender = True
End Function

Private Function W_GETeDeposit(ByVal pdTransDate As Date, ByVal ptTerminalNo As String, ByVal ptTransNo As String) As eEN_Deposit
Dim tSql As String
Dim eTmp As eEN_Deposit
'Dim eParent As eEN_TransPrimary

Dim orsTmp As ADODB.Recordset

On Error GoTo ErrHandler
        
        'self
        eTmp.tEN_DepositNo = SP_FEDtChkString(orsHD, "FTCdpDepositNo")
        eTmp.cEN_PevPayment = SP_FEDvChkNumeric(orsHD, "FCShdDepositBPay") 'เงินวางมัดจำก่อนหน้า
        eTmp.cEN_PrevBalDue = SP_FEDvChkNumeric(orsHD, "FCShdDepositPrvBal") 'เงินที่ยังขาดอยู่ก่อนหน้า
        eTmp.cEN_Deposit = SP_FEDvChkNumeric(orsHD, "FCShdDepositAmt") 'เงินมัดจำ ในใบนี้
        eTmp.cEN_BalDue = SP_FEDvChkNumeric(orsHD, "FCShdDepositBal") 'เงินมัดจำที่ยังขาดอยู่
        eTmp.cEN_BalDueDate = SP_FEDdChkDate(orsHD, "FDShdDepositDueDate") '
        eTmp.eEN_Parrent.dEN_TransDate = SP_FEDdChkDate(orsHD, "FDShdDepReTransDate")
        eTmp.eEN_Parrent.tEN_TmnNum = SP_FEDtChkString(orsHD, "FTShdDepRefTmnNum")
        eTmp.eEN_Parrent.tEN_TransNo = SP_FEDtChkString(orsHD, "FTShdDepReTransNo")
        eTmp.tEN_PhoneNum = SP_FEDtChkString(orsHD, "FTPmhDocNoBill") '*Neung 55-07-23 Layaway TK
'        eTmp.cEN_RndAmt = SP_FEDvChkNumeric(orsHD, "FCShdDepositTax") '*Eaw 56-07-20 ใช้เก็บ Rounding '*Eaw 56-07-26 เอามาผิดฟิลด์
        eTmp.cEN_Reprint = SP_FEDvChkNumeric(orsHD, "FCCstCardAmt3")                    '*Ao 2014/07/22
        W_GETeDeposit = eTmp
        Set orsTmp = Nothing
        
Exit Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

Private Function SP_DTEtStrSTD(pdDate As Date, Optional ByVal pbChrRet As Boolean = False, Optional ByVal ptTime$, Optional ByVal pbRetWithTime As Boolean = False) As String
'----------------------------------------------------------
'   Call:   pdDate as selected date
'   Ret:   string of date format "YYYY/MM/DD"
'----------------------------------------------------------
    Dim tRet$
    
    SP_DTEtStrSTD = Format(Year(pdDate), "0000") & "/" & Format(Month(pdDate), "00") & "/" & Format(Day(pdDate), "00")
    tRet = SP_DTEtStrSTD
    If pbChrRet Then

                  If pbRetWithTime Then
                       SP_DTEtStrSTD = "'" & tRet & " " & ptTime & "'"
                  Else
                       SP_DTEtStrSTD = "'" & tRet & "'"
                  End If

    Else
        If pbRetWithTime Then
            SP_DTEtStrSTD = tRet & " " & ptTime
        End If
    End If
End Function


'Public Property Get SaveField() As Boolean
'    SaveField = bW_SaveField
'End Property
'
'Public Property Let SaveField(ByVal pbSave As Boolean)
'    bW_SaveField = pbSave
'End Property


Public Property Get PrintEJ() As Boolean
    PrintEJ = bW_EnableEJ
End Property

Public Property Let PrintEJ(ByVal pbEnable As Boolean)
    bW_EnableEJ = pbEnable
End Property

Public Property Let ExportFolder(ByVal ptFolder As String)
    If Right(ptFolder, 1) <> "\" Then
            ptFolder = ptFolder & "\"
    End If
    
    tW_ExportFolder = ptFolder
End Property

Public Property Get ExportFolder() As String
    ExportFolder = tW_ExportFolder
End Property

Public Property Let EJFolder(ByVal ptFolder As String)
    If Right(ptFolder, 1) <> "\" Then
            ptFolder = ptFolder & "\"
    End If
    
    tW_EJFolder = ptFolder
End Property

Public Property Get EJFolder() As String
    EJFolder = tW_EJFolder
End Property

'Public Function ExportClosing(ByVal pdTransDate As Date, ByVal ptTmnNum As String, Optional ByVal ptFormat As eEN_BackupFormat = 1, Optional ByVal pbSaveField As Boolean)
''*KT 51-11-03
'Dim tSql As String
'Dim tDate As String
'Dim tTmp As String, tTblName As String, nTmp As Long
'Dim tFile As String, tFileGroup As String
'Dim tPrimary As String
'Dim aTbl() As Variant
'Dim tField As String
'Dim tTblNo As String
'Dim nDataSize  As Long
'Dim nPointer As Long
'Dim nLocal As Long, tBuffer As String * 100
'Dim orsTmp As ADODB.Recordset
'
'On Error GoTo ErrHandler
'
'        If ocnAll.State <> 1 Then Exit Function
'        'set  format date to yyyy-MM-dd  (format for bcp)
'        nLocal = GetUserDefaultLCID()
'        Call GetLocaleInfo(nLocal, LOCALE_SSHORTDATE, tBuffer, 99)
'        Call SetLocaleInfo(nLocal, LOCALE_SSHORTDATE, "yyyy-MM-dd")
'
'         tDate = Format(pdTransDate, "yyyy/mm/dd")
'         tTmp = "C" & Format(pdTransDate, "yyyy-mm-dd") & "-" & ptTmnNum   'TW 51-08-26
'         tFileGroup = tVB_CNPthBnkTr & "\" & tTmp & ".dat"
'         If SP_FLEbManFile(tFileGroup, nEN_Fle1Exist) Then
'                Call SP_FLEbManFile(tFileGroup, nEN_Fle2Delete)
'         End If
'
'         tPrimary = "1=1"
'         If SP_FLDbManFolder(tVB_CNPthBnkTr, nEN_Fld1Exist) = False Then Exit Function 'check exist folder
'
'        Open tFileGroup For Output As #1 'open new
'        Call WritePrivateProfileString("Main", "Primary", tPrimary, tFileGroup)
'        Call WritePrivateProfileString("Main", "ColumnDelimiter", CStr(Asc("|")), tFileGroup)
'        Call WritePrivateProfileString("Main", "Tables", CStr(4), tFileGroup)
'        Seek #1, LOF(1) 'EOF
'
'        tSql = "SELECT * FROM TPSTClose"
'        Call SP_SQLvExecute(tSql, orsTmp, oDB.ocnOffLine)
'        tTblNo = "Table " & CStr(1)
'        tTblName = "TPSTClose"
'
'        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'        tField = IIf(pbSaveField, W_SQLtGetField(orsTmp), vbNullString)
'         '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'         nPointer = LOF(1) ' EOF
'         Seek #1, nPointer
'         Print #1, vbCrLf
'         Print #1, "[" & tTblNo & " Data]"
'
'         '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'         nPointer = Seek(1)
'         If orsTmp Is Nothing Then
'                 Print #1, vbNullString
'                 nDataSize = 0
'         Else
'                 If orsTmp.State = 1 Then
'                         If Not (orsTmp.EOF And orsTmp.BOF) Then
'                             orsTmp.MoveFirst
'                             Print #1, orsTmp.GetString(, , "|", vbCrLf)
'                             nDataSize = (Seek(1) - nPointer) - 3
'                         End If
'                 End If
'         End If
'
'         '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'         Seek #1, 1 ' BOF
'         Call WritePrivateProfileString(tTblNo, "Name", tTblName, tFileGroup)
'         Call WritePrivateProfileString(tTblNo, "Pointer", CStr(nPointer), tFileGroup)
'         Call WritePrivateProfileString(tTblNo, "DataSize", CStr(nDataSize), tFileGroup)
'         Call WritePrivateProfileString(tTblNo, "Fields", tField, tFileGroup)
'         '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
''before exit
'        Set orsTmp = Nothing
'        Call SetLocaleInfo(nLocal, LOCALE_SSHORTDATE, Replace(tBuffer, Chr(0), vbNullString)) 'reset date format to old setting
'        Close #1
'
''return
'        ExportClosing = True
'        Exit Function
'ErrHandler: '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'    Close #1
'
'End Function

Public Function Export(ByVal pdTransDate As Date, ByVal ptTmnNum As String, ByVal ptTransNo As String, Optional ByVal ptFormat As eEN_BackupFormat = 1, Optional ByVal pbSaveField As Boolean) As Boolean
'*KT 51-11-03
Dim tSql As String
Dim tDate As String
Dim tTmp As String, tTblName As String, nTmp As Long
Dim tFile As String, tFileGroup As String
Dim tPrimary As String
Dim aTbl() As Variant
Dim tField As String
Dim tTblNo As String
Dim nDataSize  As Long
Dim nPointer As Long
Dim nLocal As Long, tBuffer As String * 100
Dim orsTmp As ADODB.Recordset
Dim oFSO As New FileSystemObject
Dim nTable As Integer
Dim tCVPrimary As String '*Eaw 56-10-18

On Error GoTo ErrHandler

        If ocnAll.State <> 1 Then Exit Function
        'set  format date to yyyy-MM-dd  (format for bcp)
'        nTable = 5                     '*Ao 2013/07/05
        nTable = 6                      '*Eaw 56-10-18 เพิ่ม TPSTCv เข้ามาจึงเป็น 6
        nLocal = GetUserDefaultLCID()
        Call GetLocaleInfo(nLocal, LOCALE_SSHORTDATE, tBuffer, 99)
        Call SetLocaleInfo(nLocal, LOCALE_SSHORTDATE, "yyyy-MM-dd")
         
        tDate = Format(pdTransDate, "yyyy/mm/dd")
        tTmp = Format(pdTransDate, "yyyy-mm-dd") & "_" & tUT_LoginCode & "-" & ptTmnNum & "-" & ptTransNo 'TW 51-08-26
        tFileGroup = tW_ExportFolder & tTmp & ".dat"
        tPrimary = "FDShdTransDate = '" & tDate & "' AND FTTmnNum='" & ptTmnNum & "' AND  FTShdTransNo='" & ptTransNo & "' "
         
         '*KT 51-11-06 ทำการลบ Folder ที่เก็บของ File dat ตอนที่เปิด Open Terminal ครั้งแรก
         '-----------------------------------------------------------------------------------------------------------------------
        If SP_VALnText2Double(tVB_Transaction) = tEN_TrnOpen Then
            If SP_FLDbManFolder(tW_ExportFolder & Format(pdTransDate, "yyyy-mm-dd") & "_" & tUT_LoginCode, nEN_Fld1Exist) Then
                    nVB_CNCountFolderBck = nVB_CNCountFolderBck + 1
                    oFSO.CopyFolder tW_ExportFolder & Format(pdTransDate, "yyyy-mm-dd") & "_" & tUT_LoginCode, tW_ExportFolder & Format(pdTransDate, "yyyy-mm-dd") & "_" & tUT_LoginCode & "[" & nVB_CNCountFolderBck & "]" '*KT 52-01-14 ทำการ Backup Folder ที่จะลบทิ้งก่อนโดยชื่อ Folder ต่อท้ายด้วย Bck
                    Call SP_FLDbManFolder(tW_ExportFolder & Format(pdTransDate, "yyyy-mm-dd") & "_" & tUT_LoginCode, nEN_Fld2Delete)
            End If
        End If
        '---------------------------------------------------------------------------------------------------------------------
        If SP_FLDbManFolder(tW_ExportFolder, nEN_Fld1Exist) = False Then Exit Function 'check exist folder
        If SP_FLEbManFile(tFileGroup, nEN_Fle1Exist) Then
            Call SP_FLEbManFile(tFileGroup, nEN_Fld2Delete)
        End If
        Open tFileGroup For Output As #1 'open new
        Call WritePrivateProfileString("Main", "Primary", tPrimary, tFileGroup)
        Call WritePrivateProfileString("Main", "ColumnDelimiter", CStr(Asc("|")), tFileGroup)
'        Call WritePrivateProfileString("Main", "Tables", CStr(4), tFileGroup)                                              '*Ao 2013/07/05
        Call WritePrivateProfileString("Main", "Tables", CStr(nTable), tFileGroup)
        Seek #1, LOF(1) 'EOF
        
        For nTmp = 1 To nTable  'UBound(aTbl)            '*Ao 2013/07/05
            
                tTblNo = "Table " & CStr(nTmp)
               '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
               'ใช้ ors จากการ open transacrtion
               
               Select Case nTmp
               Case 1
                    orsHD.Filter = adFilterNone '*Eaw 56-11-19
'                    Set orsTmp = orsHD
                    '*Ao 2014/09/16    ต้อง Select ใหม่เพราะมีการ Update HD             V 4.2.145
                    tSql = "SELECT *  FROM " & tW_TblHD & " WITH (NOLOCK) WHERE " & tW_TransWhere
                    Call SP_SQLvExecute(tSql, orsTmp)
                    If (orsTmp.BOF And orsTmp.EOF) Then
                        Set orsTmp = orsHD
                    End If
                    '------------------------------------------------------------------------
                    tTblName = tVB_CNTblHD
               Case 2
                    orsDT.Filter = adFilterNone '*Eaw 56-11-19
                    Set orsTmp = orsDT
                    tTblName = tVB_CNTblDT
               Case 3
                    orsRC.Filter = adFilterNone '*Eaw 56-11-19
                    Set orsTmp = orsRC
                    tTblName = tVB_CNTblRC
               Case 4
                    orsCD.Filter = adFilterNone '*Eaw 56-11-19
                    Set orsTmp = orsCD
                    tTblName = tVB_CNTblCD
                Case 5
                    Call SP_DATxRsNothing(orsTmp)
                    Call W_GETbGetTBSalePoint(orsTmp, "TPSTSalePoint", tPrimary)
                    tTblName = "TPSTSalePoint"
                Case 6 '*Eaw 56-10-18
                    tCVPrimary = "FTPOSNo ='" & ptTmnNum & "' AND  FTTransNo='" & ptTransNo & "' "
                    Call SP_DATxRsNothing(orsTmp)
                    Call W_GETbGetTBCV(orsTmp, "TPSTCV", tCVPrimary)
                    tTblName = "TPSTCV"
               End Select

                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                tField = IIf(pbSaveField, W_SQLtGetField(orsTmp), vbNullString)
                 
                 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                 nPointer = LOF(1) ' EOF
                 Seek #1, nPointer
                 Print #1, vbCrLf
                 Print #1, "[" & tTblNo & " Data]"
                 
                 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                nPointer = Seek(1)
                If orsTmp Is Nothing Then
'                        Print #1, vbNullString             '*Ao 2013/10/03
                        nDataSize = 0
                Else
                    If orsTmp.State = 1 Then
                        If Not (orsTmp.EOF And orsTmp.BOF) Then
                            orsTmp.MoveFirst
                            'Print #1, Replace(orsTmp.GetString(, , "|", vbCrLf), "'", "''") '*KT 52-06-05 PH 2.0 ML-SF-018 ในกรณี Offline ถ้าขายสินค้าที่มีชื่อของสินค้ามีเครื่องหมาย '   เมื่อ Online จะเกิด Missing Transaction และไม่สามารถ Recovery Dat file โดยใช้ Import Text ได้
                            Print #1, orsTmp.GetString(, , "|", vbCrLf)
                            nDataSize = (Seek(1) - nPointer) - 3
                        Else
                            nDataSize = 0       '*Ao 2013/10/03
                        End If
                    Else
                        nDataSize = 0       '*Ao 2013/10/03
                    End If
                End If
                 
                 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                 Seek #1, 1 ' BOF
                 Call WritePrivateProfileString(tTblNo, "Name", tTblName, tFileGroup)
                 Call WritePrivateProfileString(tTblNo, "Pointer", CStr(nPointer), tFileGroup)
                 Call WritePrivateProfileString(tTblNo, "DataSize", CStr(nDataSize), tFileGroup)
                 Call WritePrivateProfileString(tTblNo, "Fields", tField, tFileGroup)
                 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        Next
         
'before exit
        Set orsTmp = Nothing
        Call SetLocaleInfo(nLocal, LOCALE_SSHORTDATE, Replace(tBuffer, Chr(0), vbNullString)) 'reset date format to old setting
        Close #1
        Set oFSO = Nothing
'return
        Export = True
        Exit Function
ErrHandler: '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Close #1
    Call SP_CRTxLogFile("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Events:Export" & " Type:Write DAT Ticket : Terminal No. " & ptTmnNum & " Trans No. " & ptTransNo) '*KT 52-12-21 ทำการเก้บ Log การเกิด Error ที่เกี่ยวข้องการการทำ Dat File และ EJ
End Function

'Public Function Import(ByVal ptBakFile As String, Optional ByVal ptServer As String, Optional ByVal ptUser As String, Optional ByVal ptPass As String, Optional ByVal ptDatabase As String, Optional ByVal ptSal As Long = 1) As Boolean 'ptSal = 1 Sale จริง 2
''*KT 51-10-31
''*KT 51-11-03
'Dim tSql As String
'Dim tFile As String
'Dim nTmp As Long, tTmp As String
'Dim tRet  As String * 255
'Dim nLen As Long
'Dim aByte() As Byte
'Dim tPrimary As String, tColumnDelimiter As String, nTblCnt As Long
'Dim tTbl As String, nPointer As Long, nDataSize As Long
'Dim orsTemp As ADODB.Recordset
'Dim bPass As Boolean
'
'
'On Error GoTo ErrHandler
'        Import = False
'        bPass = False
'
'        'use DMO  ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'        If ptServer <> vbNullString Then 're conenct
'                With oW_Server
'                        .DisConnect
'                        .EnableBcp = True
'                        '.LoginTimeout = "10"
'                        .Connect ptServer, ptUser, ptPass
'                        Set oW_Database = .Databases(ptDatabase)
'                        'oW_Database.DBOption.SelectIntoBulkCopy = False
'               End With
'        End If  ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'                        With oW_BulkCp
'                                .UseExistingConnection = True
'                                .UseBulkCopyOption = True
'                                .DataFileType = SQLDMODataFile_SpecialDelimitedChar
'                                .UseServerSideBCP = False
''                                .ColumnDelimiter = "|"
'                                .RowDelimiter = vbCrLf
'                                .MaximumErrorsBeforeAbort = 1
'                                .SetCodePage SQLDMOBCP_RAW
'                                .MaximumErrorsBeforeAbort = 1
'                                .ImportRowsPerBatch = 1000
'                                .ServerBCPDataFileType = SQLDMOBCPDataFile_Char
'                                .UseExistingConnection = True
'                        End With
'
'                        nLen = GetPrivateProfileString("Main", "Primary", "", tRet, 255, ptBakFile)
'                        tPrimary = Left(tRet, nLen)
'                        If tPrimary = vbNullString Then Exit Function
'
'                        nLen = GetPrivateProfileString("Main", "ColumnDelimiter", "", tRet, 255, ptBakFile)
'                        tColumnDelimiter = Chr(Val(tRet))
'                        oW_BulkCp.ColumnDelimiter = tColumnDelimiter
'
'                        nLen = GetPrivateProfileString("Main", "Tables", "1", tRet, 255, ptBakFile)
'                        nTblCnt = Val(tRet)
'
'                        tFile = W_FLEtGetTempFile
'
''oW_Server.BeginTransaction
'                        bPass = True
'                        For nTmp = 1 To nTblCnt
'                                        nLen = GetPrivateProfileString("Table " & nTmp, "Name", vbNullString, tRet, 255, ptBakFile)
'                                        tTbl = Left(tRet, nLen)
'                                        '*KT 51-12-15 V 1.0.3 Phase 2 เรื่องของการนำเข้าข้อมูลจาก File Dat
'                                        '--------------------------------------------------------------------------------------------------------------------------------------
'                                        If ptSal = 1 Then
'                                            Select Case UCase(Mid(tTbl, 1, 5))
'                                                    Case "TPSHD"
'                                                        tTbl = "TPSTSalHD"
'                                                    Case "TPSDT"
'                                                        tTbl = "TPSTSalDT"
'                                                    Case "TPSRC"
'                                                        tTbl = "TPSTSalRC"
'                                                    Case "TPSCD"
'                                                        tTbl = "TPSTSalCD"
'                                                    Case Else
'                                                        tTbl = tTbl
'                                                End Select
'                                        End If
'                                        '---------------------------------------------------------------------------------------------------------------------------------------
'                                        nLen = GetPrivateProfileString("Table " & nTmp, "Pointer", vbNullString, tRet, 255, ptBakFile)
'                                        nPointer = Val(tRet)
'
'                                        nLen = GetPrivateProfileString("Table " & nTmp, "DataSize", vbNullString, tRet, 255, ptBakFile)
'                                        nDataSize = Val(tRet)
'
'                                        If nDataSize <> 0 Then
'                                                'delete exist data
'                                                'tSql = "DELETE  FROM [" & tTbl & "] WHERE " & tPrimary
'                                                'Call oW_Database.ExecuteImmediate(tSql)
'                                                'select
'                                                tSql = "SELECT FTShdTransNo  FROM [" & tTbl & "] WHERE " & tPrimary
'                                                Set orsTemp = oDB.ocnOnLine.Execute(tSql)
'                                                If orsTemp.BOF And orsTemp.EOF Then
'                                                    'Extract
'                                                    If Not Extract(ptBakFile, tTbl, tFile, ptSal) Then Exit Function
'                                                    'import
'                                                    oW_BulkCp.DataFilePath = tFile
'                                                    oW_Database.Tables(tTbl).ImportData oW_BulkCp
'                                                     'delete temp file
'                                                     Kill tFile
'                                                End If
'                                        End If
'                        Next
'
'
' On Error GoTo ErrWithRollback
'
' 'before exit
''        oW_Server.CommitTransaction
'        Close #1
'        Import = True
''return
'        Import = True
'        If Not orsTemp Is Nothing Then
'                If orsTemp.State = adStateOpen Then orsTemp.Close          'ไม่ใช้  State เพราะกรณี recordset ยังไม่ได้เปิดจะ error
'                Set orsTemp = Nothing
'        End If
'
'Exit Function
'ErrWithRollback: '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
'        Close #1
''       If bPass Then oW_Server.RollbackTransaction
'        If Not orsTemp Is Nothing Then
'                If orsTemp.State = adStateOpen Then orsTemp.Close          'ไม่ใช้  State เพราะกรณี recordset ยังไม่ได้เปิดจะ error
'                Set orsTemp = Nothing
'        End If
'
'Exit Function
'ErrHandler: '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'        'MsgBox Err.Description
'        Close #1
'        Close #2
''         If bPass Then oW_Server.RollbackTransaction
'        If Not orsTemp Is Nothing Then
'                If orsTemp.State = adStateOpen Then orsTemp.Close          'ไม่ใช้  State เพราะกรณี recordset ยังไม่ได้เปิดจะ error
'                Set orsTemp = Nothing
'        End If
'
'End Function
Public Function Import(ByVal ptBakFile As String, Optional ByVal ptServer As String, Optional ByVal ptUser As String, Optional ByVal ptPass As String, Optional ByVal ptDatabase As String)
'*Bump 57-03-06 copy มาจาก ImportText.exe V.4.0.9
Dim tSql As String
Dim tFile As String
Dim nTmp As Long, tTmp As String
Dim tRet  As String * 255
Dim nLen As Long
Dim aByte() As Byte
Dim tPrimary As String, tColumnDelimiter As String, nTblCnt As Long
Dim tTbl As String, nPointer As Long, nDataSize As Long
Dim orsTemp As ADODB.Recordset
Dim bPass As Boolean
'*Eaw 56-10-18
Dim tPosNo As String
Dim tTransNo As String
Dim nPos As Integer
Dim tCVPrimary As String
Dim tTransType As String  '*BG 56-10-21
On Error GoTo ErrHandler
        bPass = False

        'use DMO  ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        If ptServer <> vbNullString Then 're conenct
                
                With oW_Server
                        .DisConnect
                        .EnableBcp = True
                        '.LoginTimeout = "10"
                        .Connect ptServer, ptUser, ptPass
                    
                        Set oW_Database = .Databases(ptDatabase)
                         
                        'oW_Database.DBOption.SelectIntoBulkCopy = False
               
               End With
                         
'                With oW_BulkCp
'                        .DataFileType = SQLDMODataFile_SpecialDelimitedChar
'                        .ColumnDelimiter = "|" 'vbTab
'                        .RowDelimiter = vbCrLf
'                        .UseServerSideBCP = False
'                        .MaximumErrorsBeforeAbort = 1
'                        .SetCodePage (SQLDMOBCP_RAW)
'                        .MaximumErrorsBeforeAbort = 1
'                        .ImportRowsPerBatch = 1000
'                        .ServerBCPDataFileType = SQLDMOBCPDataFile_Char
'                        .UseExistingConnection = True
'                        .UseBulkCopyOption = True
'                End With

        End If  ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
               
                        With oW_BulkCp
                                .UseExistingConnection = True
                               'Bump 56-02-07    SQL2012 เปลี่ยนวิธีการ จึงไม่ต้องใช้ Option นี้  CommSheet TK-ISS3000-036.xls ***
                                If nVB_SQLVersion = 9 Or nVB_SQLVersion = 10 Then
                                    .UseBulkCopyOption = True
                                End If
                                '**************************************************************************************
                                .DataFileType = SQLDMODataFile_SpecialDelimitedChar
                                .UseServerSideBCP = False
'                                .ColumnDelimiter = "|"
                                .RowDelimiter = vbCrLf
                                .MaximumErrorsBeforeAbort = 1
                                .SetCodePage SQLDMOBCP_RAW
                                .MaximumErrorsBeforeAbort = 1
                                .ImportRowsPerBatch = 1000
                                .ServerBCPDataFileType = SQLDMOBCPDataFile_Char
                                .UseExistingConnection = True
                        End With
                        
                        nLen = GetPrivateProfileString("Main", "Primary", "", tRet, 255, ptBakFile)
                        tPrimary = Left(tRet, nLen)
                        If tPrimary = vbNullString Then Exit Function
        
                        nLen = GetPrivateProfileString("Main", "ColumnDelimiter", "", tRet, 255, ptBakFile)
                        tColumnDelimiter = Chr(Val(tRet))
                        oW_BulkCp.ColumnDelimiter = "|" 'tColumnDelimiter
                        
                        nLen = GetPrivateProfileString("Main", "Tables", "1", tRet, 255, ptBakFile)
                        nTblCnt = Val(tRet)
                        
                        tFile = W_FLEtGetTempFile
                        
'oW_Server.BeginTransaction
                        bPass = True
                        For nTmp = 1 To nTblCnt
                        
                                        nLen = GetPrivateProfileString("Table " & nTmp, "Name", vbNullString, tRet, 255, ptBakFile)
                                        tTbl = Left(tRet, nLen)
                                        Select Case UCase(Mid(tTbl, 1, 5))
                                            Case "TPSHD"
                                                tTbl = "TPSTSalHD"
                                            Case "TPSDT"
                                                tTbl = "TPSTSalDT"
                                            Case "TPSRC"
                                                tTbl = "TPSTSalRC"
                                            Case "TPSCD"
                                                tTbl = "TPSTSalCD"
                                            Case "TPSTC" '*Eaw 56-10-18
                                                tTbl = "TPSTCV"
                                                nPos = InStr(1, tPrimary, "FTTmnNum", vbBinaryCompare)
                                                tPosNo = Mid(tPrimary, nPos + 8, 8)
                                                nPos = InStr(1, tPrimary, "FTShdTransNo", vbBinaryCompare)
                                                tTransNo = Mid(tPrimary, nPos + 12, 8)
                                                tCVPrimary = "FTPosNo " & tPosNo & " AND FTTransNo " & tTransNo
                                            Case Else
                                                tTbl = tTbl
                                        End Select
                                        
                                        nLen = GetPrivateProfileString("Table " & nTmp, "Pointer", vbNullString, tRet, 255, ptBakFile)
                                        nPointer = Val(tRet)
                                        
                                        nLen = GetPrivateProfileString("Table " & nTmp, "DataSize", vbNullString, tRet, 255, ptBakFile)
                                        nDataSize = Val(tRet)
                                        
                                        If nDataSize <> 0 Then
                                                
                                                If Not Extract(ptBakFile, tTbl, tFile) Then Exit Function
                                                'wMain.olaTbl.Caption = tTbl
                                                DoEvents
                                                
                                                Select Case UCase(Mid(tTbl, 1, 5))
                                                    Case "TPSHD"
                                                        tSql = "SELECT FTShdTransNo  FROM [" & tTbl & "] WHERE " & tPrimary
                                                    Case "TPSDT"
                                                         tSql = "SELECT FTShdTransNo  FROM [" & tTbl & "] WHERE " & tPrimary
                                                    Case "TPSRC"
                                                        tSql = "SELECT FTShdTransNo  FROM [" & tTbl & "] WHERE " & tPrimary
                                                    Case "TPSCD"
                                                         tSql = "SELECT FTShdTransNo  FROM [" & tTbl & "] WHERE " & tPrimary
                                                    Case "TPSTC" '*Eaw 56-10-18
                                                        tSql = "SELECT *  FROM [" & tTbl & "] WHERE " & tCVPrimary
                                                    Case Else
                                                         tSql = "SELECT *  FROM [" & tTbl & "] WHERE " & tPrimary
                                                End Select

                                                If UCase(tTbl) = UCase("TPSTClose") Then
                                                     tSql = "DELETE TPSTClose"
                                                     
                                                     Call ocnAll.Execute(tSql)
                                                     tSql = "SELECT *  FROM [" & tTbl & "] WHERE " & tPrimary
                                                End If
                                                
                                                'Set orsTemp = ocnAll.Execute(tSql)
                                                Set orsTemp = oDB.ocnOffLine.Execute(tSql) '* Bump 57-03-06 import เข้า OFFLINE
                                                'oDB.ocnOnLine.
                                                'MsgBox "oDB.ocnOffLine.Execute(tSql)=" & tSql
                                                If orsTemp.BOF And orsTemp.EOF Then
                                                    'bVB_IsDuplicate = False '*Bump 56-10-18
                                                    'import
                                                    'MsgBox "tFile=" & tFile
                                                    oW_BulkCp.DataFilePath = tFile
                                                    oW_Database.Tables(tTbl).ImportData oW_BulkCp 'Bump 56-02-07 ดูอย่างเดียว
                                                Else
                                                    'bVB_IsDuplicate = True '*Bump 56-10-18
                                                End If
                                                'bVB_Result = True
                                                'delete temp file
                                                Kill tFile
                                                '*BG 56-10-21 Import Deposit table
                                                Call SP_INSxDepositTbl(tTbl, tPrimary)
                                                '****
                                        End If
                        Next
                        Call C_UPDxUpdatePoint(tPrimary)

 On Error GoTo ErrWithRollback
         
 'before exit
'        oW_Server.CommitTransaction
        Close #1
        
        
'return
        Import = True
        If Not orsTemp Is Nothing Then
                If orsTemp.State = adStateOpen Then orsTemp.Close          'ไม่ใช้  State เพราะกรณี recordset ยังไม่ได้เปิดจะ error
                Set orsTemp = Nothing
        End If
        
Exit Function
ErrWithRollback: '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'        MsgBox Err.Number & vbCrLf & Err.Description
        Close #1
'        If bPass Then oW_Server.RollbackTransaction
        If Not orsTemp Is Nothing Then
                If orsTemp.State = adStateOpen Then orsTemp.Close          'ไม่ใช้  State เพราะกรณี recordset ยังไม่ได้เปิดจะ error
                Set orsTemp = Nothing
        End If
        'bVB_Result = False
Exit Function
ErrHandler: '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        MsgBox Err.Number & vbCrLf & Err.Description
'        MsgBox tFile
'        MsgBox tTbl
        Close #1
        Close #2
'        If bPass Then oW_Server.RollbackTransaction
        If Not orsTemp Is Nothing Then
                If orsTemp.State = adStateOpen Then orsTemp.Close          'ไม่ใช้  State เพราะกรณี recordset ยังไม่ได้เปิดจะ error
                Set orsTemp = Nothing
        End If
        'bVB_Result = False
End Function

Public Function ImportEx(ByVal ptBakFile As String)
Dim tSql As String
Dim tFile As String
Dim nTmp As Long, tTmp As String
Dim tRet  As String * 255
Dim nLen As Long
Dim aByte() As Byte
Dim tPrimary As String, tColumnDelimiter As String, nTblCnt As Long
Dim tTbl As String, nPointer As Long, nDataSize As Long

On Error GoTo ErrHandler

   
                        nLen = GetPrivateProfileString("Main", "Primary", "", tRet, 255, ptBakFile)
                        tPrimary = Left(tRet, nLen)
                        If tPrimary = vbNullString Then Exit Function
        
                        nLen = GetPrivateProfileString("Main", "ColumnDelimiter", "", tRet, 255, ptBakFile)
                        tColumnDelimiter = Chr(Val(tRet))
                        
                        nLen = GetPrivateProfileString("Main", "Tables", "1", tRet, 255, ptBakFile)
                        nTblCnt = Val(tRet)
                        

                        For nTmp = 1 To nTblCnt
                                        nLen = GetPrivateProfileString("Table " & nTmp, "Name", vbNullString, tRet, 255, ptBakFile)
                                        tTbl = Left(tRet, nLen)
                                        
                                        nLen = GetPrivateProfileString("Table " & nTmp, "Pointer", vbNullString, tRet, 255, ptBakFile)
                                        nPointer = Val(tRet)
                                        
                                        nLen = GetPrivateProfileString("Table " & nTmp, "DataSize", vbNullString, tRet, 255, ptBakFile)
                                        nDataSize = Val(tRet)
                                        tFile = W_FLEtGetTempFileEx(tTbl)
                                        
                                        If nDataSize <> 0 Then
                                                'Extract
                                                If Not ExtractEx(ptBakFile, tTbl, tFile) Then Exit Function
                                                Select Case UCase(tTbl)
                                                    Case UCase("TPSTSalHD"), UCase(tVB_CNTblHD)
                                                        aVB_PathFileBck(0) = tFile
                                                    Case UCase("TPSTSalDT"), UCase(tVB_CNTblDT)
                                                        aVB_PathFileBck(1) = tFile
                                                    Case UCase("TPSTSalCD"), UCase(tVB_CNTblCD)
                                                        aVB_PathFileBck(2) = tFile
                                                    Case UCase("TPSTSalRC"), UCase(tVB_CNTblRC)
                                                        aVB_PathFileBck(3) = tFile
                                                End Select
                                        End If
                        Next
        
Exit Function
ErrHandler: '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

End Function


Public Function Extract(ByVal ptBakFile As String, ByVal ptTable As String, ByVal ptOutFile As String, Optional ByVal ptSal As Long = 1) As Boolean
Dim nTmp As Long
Dim nLen As Long, tRet As String * 255
Dim tTbl As String, nPointer As Long, nDataSize As Long
Dim aByte() As Byte
Dim bFoundTbl As Boolean
Dim nTblCnt As Long

On Error GoTo ErrHandler
        
        If ptOutFile = vbNullString Then Exit Function
        
        nLen = GetPrivateProfileString("Main", "Tables", "1", tRet, 255, ptBakFile)
        nTblCnt = Val(tRet)
                        
        'find  table
        For nTmp = 1 To nTblCnt
        
                    nLen = GetPrivateProfileString("Table " & nTmp, "Name", vbNullString, tRet, 255, ptBakFile)
                    tTbl = Left(tRet, nLen)
                    '*KT 52-01-21
                    If ptSal = 1 Then
                     '*KT 51-12-15 V 1.0.3 Phase 2 เรื่องของการนำเข้าข้อมูลจาก File Dat
                         '-----------------------------------------------------------------------------------------------------------------
                        Select Case UCase(Mid(tTbl, 1, 5))
                            Case "TPSHD"
                                tTbl = "TPSTSalHD"
                            Case "TPSDT"
                                tTbl = "TPSTSalDT"
                            Case "TPSRC"
                                tTbl = "TPSTSalRC"
                            Case "TPSCD"
                                tTbl = "TPSTSalCD"
                            Case Else
                                tTbl = tTbl
                        End Select
                        '-------------------------------------------------------------------------------------------------------------------
                    End If
                    If tTbl = ptTable Then
                    
                            nLen = GetPrivateProfileString("Table " & nTmp, "Pointer", vbNullString, tRet, 255, ptBakFile)
                            nPointer = Val(tRet)
                            nLen = GetPrivateProfileString("Table " & nTmp, "DataSize", vbNullString, tRet, 255, ptBakFile)
                            nDataSize = Val(tRet)
                            bFoundTbl = True
                            Exit For
                            
                    End If
        Next
        
        If Not bFoundTbl Then Exit Function
        
        'get and put to out file
        Open ptBakFile For Binary As 1
        ReDim aByte(nDataSize + 1) As Byte
        Get #1, nPointer, aByte
        Close #1
        If Dir(ptOutFile) <> vbNullString Then Kill ptOutFile
        Open ptOutFile For Binary Access Write As #2
        Put #2, , aByte
        Close #2
                                        
        
 'before exit
Close #1
Close #2

'return
Extract = True

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

Close #1
Close #2

End Function

Public Function ExtractEx(ByVal ptBakFile As String, ByVal ptTable As String, ByVal ptOutFile As String) As Boolean
Dim nTmp As Long
Dim nLen As Long, tRet As String * 255
Dim tTbl As String, nPointer As Long, nDataSize As Long
Dim aByte() As Byte
Dim bFoundTbl As Boolean
Dim nTblCnt As Long

On Error GoTo ErrHandler
        
        If ptOutFile = vbNullString Then Exit Function
        
        nLen = GetPrivateProfileString("Main", "Tables", "1", tRet, 255, ptBakFile)
        nTblCnt = Val(tRet)
                        
        'find  table
        For nTmp = 1 To nTblCnt
        
                    nLen = GetPrivateProfileString("Table " & nTmp, "Name", vbNullString, tRet, 255, ptBakFile)
                    tTbl = Left(tRet, nLen)
                    
                    If tTbl = ptTable Then
                    
                            nLen = GetPrivateProfileString("Table " & nTmp, "Pointer", vbNullString, tRet, 255, ptBakFile)
                            nPointer = Val(tRet)
                            nLen = GetPrivateProfileString("Table " & nTmp, "DataSize", vbNullString, tRet, 255, ptBakFile)
                            nDataSize = Val(tRet)
                            bFoundTbl = True
                            Exit For
                            
                    End If
        Next
        
        If Not bFoundTbl Then Exit Function
        
        'get and put to out file
        Open ptBakFile For Binary As 1
        ReDim aByte(nDataSize + 1) As Byte
        Get #1, nPointer, aByte
        Close #1
        Open ptOutFile For Binary Access Write As #2
        Put #2, , aByte
        Close #2
                                        
        
 'before exit
Close #1
Close #2

'return
ExtractEx = True

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

Close #1
Close #2

End Function


Private Function W_SQLtGetField(ByRef poRecordset As ADODB.Recordset) As String
Dim nTmp As Long, tTmp As String
On Error GoTo ErrHandler
        If poRecordset Is Nothing Then Exit Function
        
        For nTmp = 0 To poRecordset.Fields.Count - 1
                tTmp = tTmp & "," & poRecordset.Fields(nTmp).Name
        Next
        
        W_SQLtGetField = Mid(tTmp, 2)
        
        
Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

Private Function W_FLEtGetTempFile() As String

Dim nTmp As Long, tTmp As String
Dim tPath As String, tFile As String, tFileFormat As String

Dim tFolder As String ' Name of the folder
Dim nLen As Long ' Return Value

On Error GoTo ErrHandler
        
        tFolder = String(MAX_PATH, 0)
        nLen = GetTempPath(MAX_PATH, tFolder)
        
        If nLen <> 0 Then
            tPath = Left(tFolder, nLen)
        Else
            Exit Function
        End If
    
        tFileFormat = "Temp" & Format(Now, "YYMMDDHHMMSS")
        For nTmp = 1 To 99
            tFile = tPath & tFileFormat & Format(nTmp, "00") & ".bcp"
            If Dir(tFile, vbNormal) = vbNullString Then Exit For
        Next
        
'return
W_FLEtGetTempFile = tFile

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function


Private Function W_FLEtGetTempFileEx(ByVal ptTbl$) As String

Dim nTmp As Long, tTmp As String
Dim tPath As String, tFile As String, tFileFormat As String

Dim tFolder As String ' Name of the folder
Dim nLen As Long ' Return Value

On Error GoTo ErrHandler
        
        tFolder = String(MAX_PATH, 0)
        nLen = GetTempPath(MAX_PATH, tFolder)
        
        If nLen <> 0 Then
            tPath = Left(tFolder, nLen)
        Else
            Exit Function
        End If
    
        tFileFormat = "Temp" & Format(Now, "YYMMDD") & "-" & ptTbl
        tFile = tPath & tFileFormat & ".tmp"

'return
W_FLEtGetTempFileEx = tFile

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function


Private Function GetPrimary(ByVal ptBakFile As String) As String
Dim tRet As String
Dim nLen As Long
On Error GoTo ErrHandler
    GetPrimary = ""
    nLen = GetPrivateProfileString("Main", "Primary", "", tRet, 255, ptBakFile)
    GetPrimary = Left(tRet, nLen)
Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
    GetPrimary = ""
End Function

Private Function SP_FEDoChkImage(ByVal poRs As ADODB.Recordset, ByVal ptField$) As IPictureDisp
'----------------------------------------------------------
'   Cmt :   Load Field Image
'----------------------------------------------------------
    Dim tSql$, tFile$, nFF%, nSize&
    Dim aBuff() As Byte

On Error GoTo ErrHandle
    
    If poRs.BOF And poRs.EOF Then
        Set SP_FEDoChkImage = Nothing
        Exit Function
    End If
    
    With poRs
        nSize = .Fields(ptField).ActualSize
        If nSize > 0 Then aBuff = .Fields(ptField).GetChunk(nSize)
        
    End With
    
    If nSize = 0 Then
        Set SP_FEDoChkImage = Nothing
        Exit Function
    End If

'Save buffer to file
    tFile = SP_GETtNewTempFile
    If tFile = "" Then Exit Function
    nFF = FreeFile
    Open tFile For Binary As #nFF
    Put #nFF, , aBuff()
    Close #nFF

'Load file to picture
    Set SP_FEDoChkImage = LoadPicture(tFile)
    Call Kill(tFile)
    Exit Function

ErrHandle:
    Call SP_MSGShowError

End Function

Private Function SP_GETtNewTempFile() As String
'----------------------------------------------------------
'   Cmt :   กำหนดชื่อ File Temporary ใหม่
'----------------------------------------------------------
    Dim tTmp As String, tFile As String, tPath As String
    Dim nLen As Long, nC As Long

'Temporary Path
    tTmp = String(255, 0)
    nLen = GetTempPath(255, tTmp)
    If nLen <> 0 Then
        tPath = Left$(tTmp, nLen)
    Else
        SP_GETtNewTempFile = ""
        Exit Function
    End If
    
    nC = 0
    Do
        nC = nC + 1
        tFile = tPath & Format$(Now(), "_yyyymmddhhnnss") & Format$(nC, "00000000") & ".tmp"
    Loop Until Dir(tFile) = ""
    SP_GETtNewTempFile = tFile
    
End Function

Private Sub W_GETxVocabularyTicketToArray()
'----------------------------------------------------------
'   Cmt :   เก็บค่าคำศัพทืที่ใช้สำหรับพิมพ์ลง Ticket
'----------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    ReDim aW_VocabularyPrn(0)
    
    tSql = "SELECT FNSpmSeq,FTSpmID,FTSpmNameUser"
    tSql = tSql & " FROM  TSysPosMenu"
    tSql = tSql & " WHERE FTSpmGrp='Print Slip'"
    tSql = tSql & " ORDER BY FNSpmSeq"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        Do While Not orsTemp.EOF
            aW_VocabularyPrn(UBound(aW_VocabularyPrn)).nEN_Seq = SP_FEDvChkNumeric(orsTemp, "FNSpmSeq")
            aW_VocabularyPrn(UBound(aW_VocabularyPrn)).tEN_SpmID = SP_FEDtChkString(orsTemp, "FTSpmID")
            aW_VocabularyPrn(UBound(aW_VocabularyPrn)).tEN_Name = SP_FEDtChkString(orsTemp, "FTSpmNameUser")
            ReDim Preserve aW_VocabularyPrn(UBound(aW_VocabularyPrn) + 1)
            orsTemp.MoveNext
        Loop
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
     Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGShowError
End Sub

Private Function W_GETxVocabularyRate(ByVal ptText As String) As String
'------------------------------------------------------------------------
'*KT 66-08-09
'ตัดจำนวน Rate จากข้อความที่ส่งมา โดยยึดว่าเจอตัวเลขชุดแรก
'------------------------------------------------------------------------
    Dim tVal As String
    Dim tChr As String
    Dim nLoop As Integer
On Error GoTo ErrHandle
    If ptText <> "" Then
        For nLoop = 1 To Len(ptText)
            tChr = Mid(ptText, nLoop, 1)
            If IsNumeric(tChr) Then
                tVal = tVal & tChr
            Else
                If tVal <> "" Then Exit For
            End If
        Next nLoop
    End If
    W_GETxVocabularyRate = tVal
    Exit Function
ErrHandle:
    W_GETxVocabularyRate = tVal
End Function


Private Function W_GETxVocabularyTicketFromArray(ByVal ptSpmID As String) As String
'----------------------------------------------------------
'   Cmt :  หาคำศัพทืที่ใช้สำหรับพิมพ์ลง Ticket จาก SpmID ที่ส่งมา
'----------------------------------------------------------
    Dim nI As Long
On Error GoTo ErrHandle
    W_GETxVocabularyTicketFromArray = ""
    For nI = LBound(aW_VocabularyPrn) To UBound(aW_VocabularyPrn) - 1
        If aW_VocabularyPrn(nI).tEN_SpmID = ptSpmID Then
            '*Amorn 65-07-11 เปิดการทำ Return 7%
'            If aW_VocabularyPrn(nI).tEN_Name = "GST-8%(Incl)" Then
'                ptSpmID = "1723"
'            End If
            '*Bump 66-02-07 [6.2302.1] [66/02/09] GST Rate
            If (ptSpmID = "1950" Or tVB_CNMnuCurrent = "1723" Or ptSpmID) Or bVB_ChkRetGST7 Then
'            If (ptSpmID = "1950" Or tVB_CNMnuCurrent = "1723") Or bVB_ChkRetGST7 Then
                If (bVB_ChkRetGST7 = True And ptSpmID = "1950") Or (W_CHKbOldVatRate = True And ptSpmID = "1950") Or (bVB_ChkRetGST7 = True And tVB_CNMnuCurrent = "1723") Then
                    'W_GETxVocabularyTicketFromArray = Replace(aW_VocabularyPrn(nI).tEN_Name, Mid(aW_VocabularyPrn(nI).tEN_Name, 5, 1), cVB_OldGstRate)
                    W_GETxVocabularyTicketFromArray = Replace(aW_VocabularyPrn(nI).tEN_Name, W_GETxVocabularyRate(aW_VocabularyPrn(nI).tEN_Name), cVB_OldGstRate)
                Else
                    W_GETxVocabularyTicketFromArray = aW_VocabularyPrn(nI).tEN_Name
                End If
            Else
                W_GETxVocabularyTicketFromArray = aW_VocabularyPrn(nI).tEN_Name
            End If
'            If (bVB_ChkRetGST7 = True And ptSpmID = "1950") Or (W_CHKbOldVatRate = True And ptSpmID = "1950") Or (bVB_ChkRetGST7 = True And tVB_CNMnuCurrent = "1723") Then
'                W_GETxVocabularyTicketFromArray = Replace(aW_VocabularyPrn(nI).tEN_Name, Mid(aW_VocabularyPrn(nI).tEN_Name, 5, 1), cVB_OldGstRate)
'            Else
'                W_GETxVocabularyTicketFromArray = aW_VocabularyPrn(nI).tEN_Name
'            End If
'            W_GETxVocabularyTicketFromArray = aW_VocabularyPrn(nI).tEN_Name
            Exit For
        End If
    Next nI
    Exit Function
ErrHandle:
    Call SP_MSGShowError
End Function

Private Function W_EMPcBalDis(ByVal ptEmpD As String) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # หายอด Balance Limit ของ Employee
'ARGUMENT LIST # ptEmpD = หมายเลข Employee ID
'
'RETURN VALUE  # ยอด Balance Limit ของ Employee   "ึ
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim bOnline As Boolean
On Error GoTo Err_Handler
    bOnline = Not (ocnMember Is Nothing)
    If bOnline = False Then
        'If ocnAll.ConnectionString <> oDB.ocnOffLine.ConnectionString Then 'PH 2.0.0 OLD
        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
            If SP_CONbMemberServer Then
                bOnline = True
            End If
        End If
    End If
    If bOnline Then
        W_EMPcBalDis = 0
        tSql = "SELECT  Top 1 FCEcdBalDis"
        tSql = tSql & vbCrLf & "FROM TCNMEmpCard"
        tSql = tSql & vbCrLf & "WHERE FTEcdCode = '" & ptEmpD & "'"
        If SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True) = 0 Then
            'Error จะไม่เช็ค
            W_EMPcBalDis = SP_FEDvChkNumeric(orsTemp, "FCEcdBalDis")
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Public Function W_CHKtChangeCardID(ByVal ptCardID As String, ptRcvCode As String) As String
'__________________________________________________________________
'   Call:
'   Cmt:    '*Ao 09-08-01 ChangeCardID "12345678XXXX1234"
'                กรณีเกิด Error ให้ Function Return ค่า ptCardID กลับไปให้ *KT 52-09-26
'                PH2-ML-SF-127 เมื่อชำระด้วย Installment Creditจะไม่พิมพ์ หมายเลขบัตรเครดิตรบน Receipt / EJ (ให้พิมพ์แบบเดียวกับบัตรเครดิต)
'__________________________________________________________________
Dim tFirst As String, tLast As String
Const nFirst = 8
Const nLast = 13
'Const tTmp = "XXXX"
'Const tTmp = "XXXX XXXX XXXX "   '*Eaw 56-02-19 เปลี่ยนรูปแบบการพิมพ์หมายเลขบัตรเป็น XXXX XXXX XXXX 1234
Dim tTmp As String
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim tStaValue As String
Dim tStaAlw As String
Dim nI As Integer
Dim tCard As String
Dim nCount As Integer
Dim tRefDeToken As String
On Error GoTo Err_Handler
    Select Case UCase(ptRcvCode)
        Case "T002", "T003", "T017", "0" '*Eaw เคส 0 คือกรณีพิมพ์ Emp Card+Member Card
            If (Trim(ptCardID) = "") Then
                W_CHKtChangeCardID = ptCardID
                Exit Function
            End If
'*Eaw 56-02-19
'            tFirst = Left(ptCardID, nFirst)
'            tLast = Mid(ptCardID, nLast, Len(ptCardID))
'            W_CHKtChangeCardID = tFirst & tTmp & tLast

'            tLast = Mid(ptCardID, Len(ptCardID) - 3)
'            W_CHKtChangeCardID = tTmp & tLast
'-------------------------------------------------------------------------------------------
'*Eaw 56-08-17 CommSheet TK-ISS3000-269-TR_Req.xls ถ้า tStaValue=1 พิมพ์ความยาวเท่าหมายเลขบัตรจริง,2=พิมพ์ความยาวบัตร 16 หลัก
'        tSql = "SELECT FTStpCode,FTStpName,FTStpSelValue,FTStpStaAlw FROM TCNMStorePara WHERE FTStpCode='22'"
'        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
'        If Not (orsTemp.BOF And orsTemp.EOF) Then
'            tStaValue = SP_FEDtChkString(orsTemp, "FTStpSelValue")
'            tStaAlw = SP_FEDtChkString(orsTemp, "FTStpStaAlw")

        'SP_PRCbDeToken
        '*Bmp 61-04-03 Transaction Type ต้อง Detoken ก่อน
        '05  --IPV Sale
        '14  --IPV return
        '16 --Layaway cancel
        '20  --Cancel Add Layaway
        '17  --IPV layaway pickup
        '26  --IPV VIP Service Sale
        '27  --IPV Tax Except Sale
        '28  --IPV Duty Free Sale
        
        '03  Sale Transaction
        '04 Return Transaction
        '05  IPV Transaction
        '10 VIP Service Sale
        '11 Tax Except Sale
        '13 Duty Free Sale
        '14 IPV Return Transaction
        '15 Voucher
        Call SP_CRTxLogMallCardExtend("W_CHKtChangeCardID:ptCardID:[" & ptCardID & "] TransType:[" & nW_TokenTransType & "]")   '*[5.1804.2][61/04/05]*Bmp 61-04-05
        '7=Layaway Pickup ,16 = Layaway cancel, 20 =Cancel Add Layaway หมายเลข CARD ID: ต้อง detoken ก่อน
        If bVB_AlwToken Then '*ถ้า Store Para บอกว่าให้ token
            If bW_PRMGenEJ Then
            'If nW_PrintCnt > 0 And Not bW_PRMGenEJ Then
            '*TU 2020-07-29 TTK-ISS3000-762
                Select Case nW_TokenTransType
                    Case 7, 16, 19, 20 '*, 3, 4, 5, 10, 11, 13, 14, 15
                        '*ต้อง Detoken หมายเลข ก่อนพิมพ์
                        If SP_PRCbDeToken(ptCardID, tRefDeToken) Then
                            If Trim(tRefDeToken) <> "" Then ptCardID = tRefDeToken
                            Call SP_CRTxLogMallCardExtend("W_CHKtChangeCardID:tRefDeToken(bW_PRMGenEJ=True/Deposit):" & ptCardID)   '*[5.1804.2][61/04/05]*Bmp 61-04-05
                        End If
'                    Case 3, 4, 5, 10, 11, 13, 14, 15
'                        If bVB_RePrint_NotDeToken Then
'                            If SP_PRCbDeToken(ptCardID, tRefDeToken) Then
'                                If Trim(tRefDeToken) <> "" Then ptCardID = tRefDeToken
'                                Call SP_CRTxLogMallCardExtend("W_CHKtChangeCardID:tRefDeToken(bW_PRMGenEJ=True/Deposit):" & ptCardID)   '*[5.1804.2][61/04/05]*Bmp 61-04-05
'                            End If
'                        End If
                    Case Else
                        Call SP_CRTxLogMallCardExtend("W_CHKtChangeCardID:ptCardID:[" & ptCardID & "] (bW_PRMGenEJ=True/Sale,Return) TransType:[" & nW_TokenTransType & "]")   '*[5.2209.0][65/09/10]*Bmp 65-09-10 TTK-ISS3000-809
                End Select
            Else
                Select Case nW_TokenTransType
                    Case 7, 16, 19, 20
                        '*ต้อง Detoken หมายเลข ก่อนพิมพ์
                        If SP_PRCbDeToken(ptCardID, tRefDeToken) Then
                            If Trim(tRefDeToken) <> "" Then ptCardID = tRefDeToken
                            Call SP_CRTxLogMallCardExtend("W_CHKtChangeCardID:tRefDeToken(bW_PRMGenEJ=False/Deposit):" & ptCardID)   '*[5.1804.2][61/04/05]*Bmp 61-04-05
                        End If
                        Case 3, 4, 5, 10, 11, 13, 14, 15 '*Thirapong 2020-10-14 CommSheet TTK-ISS3000-769
                        If bVB_RePrint_DeToken Then
                            If SP_PRCbDeToken(ptCardID, tRefDeToken) Then
                                If Trim(tRefDeToken) <> "" Then ptCardID = tRefDeToken
                                Call SP_CRTxLogMallCardExtend("W_CHKtChangeCardID:tRefDeToken(bW_PRMGenEJ=False/Sale):" & ptCardID)   '*[5.1804.2][61/04/05]*Bmp 61-04-05
                            End If
                        End If
                    Case Else
                End Select
            End If
            
        
        End If
        '====================================================================================================
'        If ((nW_TokenTransType = 7)) Or (nW_TokenTransType = 16) Or (nW_TokenTransType = 20) Then
'            'tRefDeToken
'            If SP_PRCbDeToken(ptCardID, tRefDeToken) Then
'                If Trim(tRefDeToken) <> "" Then ptCardID = tRefDeToken
'                Call SP_CRTxLogMallCardExtend("W_CHKtChangeCardID:tRefDeToken:" & tRefDeToken)  '*[5.1804.2][61/04/05]*Bmp 61-04-05
'            End If
'        End If
        '====================================================================================================
            If bVB_ChkMarkX Then
                Select Case Trim(tVB_OptionMarkX)
                    Case "1"
    '                    If Len(ptCardID) < 4 Then
    '                        W_CHKtChangeCardID = ptCardID
    '                        Exit Function
    '                    Else
    '                        For nI = 1 To Len(ptCardID) - 4
    '                            tTmp = tTmp & "X"
    '                        Next nI
    '                        tLast = Mid(ptCardID, Len(ptCardID) - 3)
    '                        tCard = tTmp & tLast
    '                        tTmp = ""
    '                        nCount = 1
    '                        For nI = 1 To Len(tCard)
    '                            If nCount <> 5 Then
    '                                nCount = nCount + 1
    '                                tTmp = tTmp & Mid(tCard, nI, 1)
    '                            Else
    '                                tTmp = tTmp & " "
    '                                nI = nI - 1
    '                                nCount = 1
    '                            End If
    '                        Next nI
    '                        W_CHKtChangeCardID = tTmp
    '                         Exit Function
    '                    End If
                        tTmp = ""                                                                   '*Ao 2013/08/26
                        If Len(ptCardID) > 8 Then
                            For nI = 1 To Len(ptCardID) - 8
                                tTmp = tTmp & "X"
                            Next nI
                            tFirst = Left(ptCardID, 4)
                            tLast = Right(ptCardID, 4)
                            tCard = tFirst & tTmp & tLast
                            tTmp = ""
                            For nI = 1 To Len(tCard)
                                Select Case nI
                                    Case 5, 9, 13, 17, 21, 25
                                        tTmp = tTmp & " "
                                End Select
                                tTmp = tTmp & Mid(tCard, nI, 1)
                            Next nI
                        Else
                            If Len(ptCardID) < 5 Then
                                tTmp = ptCardID
                            Else
                                For nI = 1 To Len(ptCardID) - 4
                                    tTmp = tTmp & "X"
                                Next nI
                                tLast = Right(ptCardID, 4)
                                tCard = tTmp & tLast
                                tTmp = ""
                                For nI = 1 To Len(tCard)
                                    Select Case nI
                                        Case 5, 9, 13
                                            tTmp = tTmp & " "
                                    End Select
                                    tTmp = tTmp & Mid(tCard, nI, 1)
                                Next nI
                            End If
                        End If
                        W_CHKtChangeCardID = tTmp
                        Exit Function
                    Case "2"
                        tTmp = "XXXX XXXX XXXX "
                        '*Eaw 56-05-15 CommSheet TK-ISS3000-159-IT2 00.00.03.xls
                        Select Case Len(ptCardID)
                            Case 1
                                tLast = "000" & ptCardID
                                W_CHKtChangeCardID = tTmp & tLast
                            Case 2
                                tLast = "00" & ptCardID
                                W_CHKtChangeCardID = tTmp & tLast
                            Case 3
                                tLast = "0" & ptCardID
                                W_CHKtChangeCardID = tTmp & tLast
                            Case 4
                                tLast = ptCardID
                                W_CHKtChangeCardID = tTmp & tLast
                            Case Else
                                tLast = Mid(ptCardID, Len(ptCardID) - 3)
                                W_CHKtChangeCardID = tTmp & tLast
                         End Select
                         '--------------------------------------------------------------
                         Exit Function
                End Select
            Else
                W_CHKtChangeCardID = ptCardID
                Exit Function
            End If
    End Select
Err_Handler:
    W_CHKtChangeCardID = ptCardID
End Function

'Private Function W_PRCcTotalDealPromo(ByVal porsCD As ADODB.Recordset, ByVal tDealNum As String, ByVal nDealGroup As String) As Double
''-----------------------------------------------------------
''Cmt : ทำการรวบรวมการ Total ของส่วนลด Promotin Time,Promotion,Member Price(Non Consignment)
''*KT 52-11-28 Deal&Promo
''----------------------------------------------------------
'    Dim orsTemp As ADODB.Recordset
'    Dim nSdtSeqNo As Long
'On Error GoTo Err_Handler
'    W_PRCcTotalDealPromo = 0
'    Set orsTemp = porsCD.Clone
'    porsCD.Filter = "FTDctRefCode = '" & tDealNum & "'  AND FNScdDealGroup =" & nDealGroup & " AND FTScdDiscType<>'4'"
'    porsCD.Sort = "FNScdSeqNo ASC"
'    If Not (porsCD.BOF And porsCD.EOF) Then
'        Do While Not porsCD.EOF
'            nSdtSeqNo = SP_FEDvChkNumeric(porsCD, "FNSdtSeqNo")
'            orsTemp.Filter = "FNSdtSeqNo = " & nSdtSeqNo & " AND FTSdtStkCode  =  '1'"
'            If Not (orsTemp.BOF And orsTemp.EOF) Then
'                Do While Not orsTemp.EOF
'                    W_PRCcTotalDealPromo = W_PRCcTotalDealPromo + SP_FEDvChkNumeric(orsTemp, "FCScdAmt")
'                    orsTemp.MoveNext
'                Loop
'            End If
'            orsTemp.Filter = adFilterNone
'            porsCD.MoveNext
'        Loop
'    Else
'        W_PRCcTotalDealPromo = 0
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'    Exit Function
'Err_Handler:
'    W_PRCcTotalDealPromo = 0
'    Call SP_DATxRsNothing(orsTemp)
'End Function

Public Sub W_PRNbTicketPmtCpn()
'-----------------------------------------------------------
'Cmt : ทำการพิมพ์ Slip Promotion Coupon ตาม format
'----------------------------------------------------------
    Dim tSql As String
    Dim orsSlip As ADODB.Recordset
    Dim orsMessage  As ADODB.Recordset
    Dim tTmp As String
    Dim nI As Long
    Dim tFed As String
    Dim tCpnID As String
    Dim nCut As Long
    Dim nLoop As Long
    Dim bMessage As Boolean
    Dim nH As Long
    Dim oPicture As IPictureDisp
    Dim tKey As String
    Dim tReg As String
On Error GoTo Err_Handler

    bW_StaPrnPmtCpn = False
    If Not (orsPmtCpn.BOF And orsPmtCpn.EOF) Then orsPmtCpn.MoveFirst
    'oOpos.TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'Initial start  print by OPOS
    tSql = "SELECT FTPcnDesc,FTPcnUnit,FTPcnLogo,FTPcnSlpHead,FTPcnSlpTrail,FTPcnSlpDisp "
    tSql = tSql & vbCrLf & "FROM TCNMPmtCpnJN "
    tSql = tSql & vbCrLf & "WHERE  FTPcnID=  " & SP_FEDvChkNumeric(orsPmtCpn, "FTPcnID")
    If SP_SQLvExecute(tSql, orsSlip, oDB.ocnOnLine, , , , , True, True) = 0 Then
        If Not (orsSlip.BOF And orsSlip.EOF) Then
            tKey = SP_FEDvChkNumeric(orsPmtCpn, "FTPcnID")
            'Get Message
            bMessage = False
            tSql = "SELECT *  "
            tSql = tSql & vbCrLf & "FROM TCNMSlipMsg "
            tSql = tSql & vbCrLf & "WHERE  FTSlpCode IN ('" & SP_FEDtChkString(orsSlip, "FTPcnSlpHead") & "','" & SP_FEDtChkString(orsSlip, "FTPcnSlpTrail") & "','" & SP_FEDtChkString(orsSlip, "FTPcnSlpDisp") & "')"
            If SP_SQLvExecute(tSql, orsMessage, oDB.ocnOnLine, , , , , True, True) = 0 Then
                If Not (orsMessage.BOF And orsMessage.EOF) Then
                    bMessage = True
                End If
            End If
            
            'Set Logo
            tReg = GetSetting("AdaPosMall", "POSFront", "nLogoCpn")
            If bVB_SetImageLogoCpn Then
                For nI = 0 To UBound(aVB_PicLogoCpn)
                    If aVB_PicLogoCpn(nI).tID = tKey Then
                        If tReg <> CStr(aVB_PicLogoCpn(nI).tPicName) Then
                                Call oPrn.SetBitmapCpn(aVB_PicLogoCpn(nI).oPic, tKey)
                                If bVB_ImageLogoCpn Then
                                    Call SaveSetting("AdaPosMall", "POSFront", "nLogoCpn", CStr(aVB_PicLogoCpn(nI).tPicName))
                                End If
                                Set oVB_PicLogoCpn = aVB_PicLogoCpn(nI).oPic
                        End If
                    End If
                Next nI
            End If
        
            If Not (orsPmtCpn.BOF And orsPmtCpn.EOF) Then orsPmtCpn.MoveFirst
            Do While Not orsPmtCpn.EOF
                tCpnID = SP_FEDtChkString(orsPmtCpn, "FTShdPmtCpnID")
                nCut = SP_FEDvChkNumeric(orsPmtCpn, "FNShdPmtCpnQty")
                If tCpnID <> "" Then
                    
                    
                    For nLoop = 1 To nCut
                        nH = 0
                         '*KT 54-11-14 V.4.0.4 DCR_5411001  ต้องการให้ Slip กับ Promotion Coupon เป็นแผ่นเดียว เนื่องจากกลัวว่าลูกค้าจะทำ Promotion Coupon หาย
'                        If bC_PrnByOPOS Then '*KT 54-11-10 V.4.0.3 การพิมพ์ Logo Cpn กรณี Driver
'                            oOpos.TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'Initial start  print by OPOS
'                        End If
                        'Section 1: Header
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        If bMessage Then
                                 orsMessage.Filter = "FTSlpCode = '" & SP_FEDtChkString(orsSlip, "FTPcnSlpHead") & "'"
                                 If Not (orsMessage.BOF And orsMessage.EOF) Then
                                    For nI = 1 To 5
                                        tFed = "FTSlpDescLine" & nI
                                        tTmp = SP_FEDtChkString(orsMessage, tFed)
                                        If tTmp <> "" Then
                                            nH = nH + 1
                                        End If
                                    Next nI
                                    '*KT 54-11-14 V.4.0.4 DCR_5411001  ต้องการให้ Slip กับ Promotion Coupon เป็นแผ่นเดียว เนื่องจากกลัวว่าลูกค้าจะทำ Promotion Coupon หาย
                                    Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True, nEN_PrinterOnly) ' space - - - - - -
'                                    If bC_PrnByOPOS Then '*KT 54-11-10 V.4.0.3  การพิมพ์ Logo Cpn กรณี Driver
'                                        If nH < 8 Then
'                                            For nI = 1 To 8 - nH
'                                                Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True, nEN_PrinterOnly) ' space - - - - - -
'                                            Next nI
'                                        End If
'                                    End If
                                     'line1-5
                                    For nI = 1 To 5
                                        tFed = "FTSlpDescLine" & nI
                                        tTmp = SP_FEDtChkString(orsMessage, tFed)
                                        If tTmp <> "" Then
                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True, nEN_PrinterOnly)
                                            nH = nH + 1
                                        End If
                                    Next nI
                                Else
                                    '*KT 54-11-14 V.4.0.4 DCR_5411001  ต้องการให้ Slip กับ Promotion Coupon เป็นแผ่นเดียว เนื่องจากกลัวว่าลูกค้าจะทำ Promotion Coupon หาย
                                    Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True, nEN_PrinterOnly) ' space - - - - - -
'                                    If bC_PrnByOPOS Then '*KT 54-11-10 V.4.0.3  การพิมพ์ Logo Cpn กรณี Driver
'                                        For nI = 1 To 8
'                                            Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True, nEN_PrinterOnly) ' space - - - - - -
'                                        Next nI
'                                    End If
                                End If
                                orsMessage.Filter = adFilterNone
                        Else
                            '*KT 54-11-14 V.4.0.4 DCR_5411001  ต้องการให้ Slip กับ Promotion Coupon เป็นแผ่นเดียว เนื่องจากกลัวว่าลูกค้าจะทำ Promotion Coupon หาย
                            Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True, nEN_PrinterOnly) ' space - - - - - -
'                            For nI = 1 To 8
'                                Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True, nEN_PrinterOnly) ' space - - - - - -
'                            Next nI
                        End If
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        '*KT 54-11-14 V.4.0.4 DCR_5411001  ต้องการให้ Slip กับ Promotion Coupon เป็นแผ่นเดียว เนื่องจากกลัวว่าลูกค้าจะทำ Promotion Coupon หาย
'                        If bC_PrnByOPOS Then '*KT 54-11-09 การพิมพ์ Logo Cpn กรณี Driver
'                            oOpos.CutPaper 90
'                        End If

                        'Section 2: Logo
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        If bC_PrnByOPOS Then
                            If bVB_ImageLogoCpn Then
                                oPrn.PrintLogoCpn
                            End If
                        Else
                            If bVB_SetImageLogoCpn Then
                                'oPrn.PrintPicture oPicture, 2.5, 0, vbCenter '*KT 54-11-09 V.4.0.3  การพิมพ์ Logo Cpn กรณี Driver
                                oPrn.PrintPicture oVB_PicLogoCpn, 2.5, oPrn.CurrentY, vbCenter
                            End If
                        End If
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        'Section 3: Description
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        tTmp = SP_FEDtChkString(orsSlip, "FTPcnDesc")
                        If tTmp <> "" Then
                            Set oPrn.Font = oW_Font(2)
                            Call oPrn.PrintText(tTmp, cXC, , vbCenter, True, nEN_PrinterOnly, , , True)
                            Set oPrn.Font = oW_Font(1)
                        End If
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        'Section 4: มูลค่าคูปอง
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        tTmp = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsPmtCpn, "FCShdPmtCpnVal"), False, True) & " " & SP_FEDtChkString(orsSlip, "FTPcnUnit")
                        tTmp = Trim(tTmp)
                        If tTmp <> "" Then
                            Set oPrn.Font = oW_Font(2)
                            Call oPrn.PrintText(tTmp, cXC, , vbCenter, True, nEN_PrinterOnly, , , True)
                            Set oPrn.Font = oW_Font(1)
                        End If
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        'Section 5: Detail
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                         If bMessage Then
                                orsMessage.Filter = "FTSlpCode = '" & SP_FEDtChkString(orsSlip, "FTPcnSlpDisp") & "'"
                                If Not (orsMessage.BOF And orsMessage.EOF) Then
                                    'line1-5
                                    For nI = 1 To 5
                                        tFed = "FTSlpDescLine" & nI
                                        tTmp = SP_FEDtChkString(orsMessage, tFed)
                                        If tTmp <> "" Then
                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True, nEN_PrinterOnly)
                                        End If
                                    Next nI
                                End If
                                orsMessage.Filter = adFilterNone
                        End If
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        'Section 6: แท่ง Barcode
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True, nEN_PrinterOnly) ' space - - - - - -
                        tTmp = tCpnID
                        If tTmp <> "" Then
                            If bC_PrnByOPOS Then 'Print by OPOS
                               If oOpos.RecLineWidth <= nVB_MaxChar * 10 Then
                                    oOpos.PrintBarcode PTR_S_RECEIPT, tTmp, PTR_BCS_Code39, 50, oOpos.RecLineWidth, DISP_BM_CENTER, PTR_BC_TEXT_NONE
                               Else
                                    oOpos.PrintBarcode PTR_S_RECEIPT, tTmp, PTR_BCS_Code39, 50, nVB_MaxChar * 10, DISP_BM_CENTER, PTR_BC_TEXT_NONE
                               End If
                               Call oPrn.PrintText(tTmp, cXC, , vbCenter, True, nEN_PrinterOnly)
                            Else  'Print by Driver
                                Set oPrn.Font = oW_Font(3)
                                Call oPrn.PrintText("@" & tTmp & "@", cXC, , vbCenter, True, nEN_PrinterOnly)   'for barcode 39
                                Set oPrn.Font = oW_Font(1)
                                Call oPrn.PrintText(tTmp, cXC, , vbCenter, True, nEN_PrinterOnly)
                            End If
                        End If
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        'Section 7 : Reprint
                         '----------------------------------------------------------------------------------------------------------------------------------------------------
                        If bW_EnableRePrint Then
                            Call oPrn.PrintText("Reprint", cXC, , vbCenter, True, nEN_PrinterOnly, , , True) 'ตัวหนา
                        End If
                        Call oPrn.PrintText(vbNullString, cXL, , vbLeftJustify, True, nEN_PrinterOnly) ' space - - - - - -
                         '----------------------------------------------------------------------------------------------------------------------------------------------------
                        'Section 8: Trail
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
                        If bMessage Then
                                orsMessage.Filter = "FTSlpCode = '" & SP_FEDtChkString(orsSlip, "FTPcnSlpTrail") & "'"
                                If Not (orsMessage.BOF And orsMessage.EOF) Then
                                    'line1-5
                                    For nI = 1 To 5
                                        tFed = "FTSlpDescLine" & nI
                                        tTmp = SP_FEDtChkString(orsMessage, tFed)
                                        If tTmp <> "" Then
                                            Call oPrn.PrintText(tTmp, cXL, , vbLeftJustify, True, nEN_PrinterOnly)
                                        End If
                                    Next nI
                                End If
                                orsMessage.Filter = adFilterNone
                        End If
                        '----------------------------------------------------------------------------------------------------------------------------------------------------
'                        If nLoop = nCut Then
'                                If Not oPrn.EndDoc(oW_Logo, oOpos.DrawerOpen, True) Then
'                                    bW_StaPrnPmtCpn = False
'                                End If
'                        Else
'                                If Not oPrn.EndDoc(oW_Logo, oOpos.DrawerOpen, False) Then
'                                    bW_StaPrnPmtCpn = False
'                                End If
'                        End If
                         '*KT 54-11-14 V.4.0.4 DCR_5411001  ต้องการให้ Slip กับ Promotion Coupon เป็นแผ่นเดียว เนื่องจากกลัวว่าลูกค้าจะทำ Promotion Coupon หาย
                         bW_StaPrnPmtCpn = False
                    Next nLoop
                End If
                orsPmtCpn.MoveNext
            Loop
        End If
    End If

    Call SP_DATxRsNothing(orsSlip)
    Call SP_DATxRsNothing(orsMessage)
    Exit Sub
Err_Handler:
    bW_StaPrnPmtCpn = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsSlip)
    Call SP_DATxRsNothing(orsMessage)
End Sub

Public Sub W_GETxDataToReprint()
Dim tRcvCode As String
On Error GoTo ErrHandle
    If Not (orsRC.BOF And orsRC.EOF) Then
        orsRC.MoveFirst
        Do While Not orsRC.EOF
            tRcvCode = SP_FEDtChkString(orsRC, "FTTdmCode")
            '*KT 55-08-28 Takashimaya PRJ10001TK-External Specification Credit Voucher-00.04.00
             '-------------------------------------------------------------------------------------------------------
             Select Case SP_VALnText2Double(tW_TransTypeSub)
                 Case nEN_TrnReturn
                      Select Case tRcvCode
                             Case "T018"
                                      bW_StaPrnCV2 = True
                                      bW_StaPrnCV3 = False
                                      Call W_GETxReprintCV(tW_TransType, tW_TmnNum, tW_TransNo)
                      End Select
                 Case nEN_TrnIPV, nEN_TrnIPVDutyFreeSale, nEN_TrnIPVRedeem, nEN_TrnIPVReturn, nEN_TrnIPVTaxExceptSale, nEN_TrnIPVVIPService
                     Select Case tRcvCode
                             Case "T018"
                                      bW_StaPrnCV2 = W_GETbCVIssue(tW_TransType, tW_TmnNum, tW_TransNo)
                                      bW_StaPrnCV3 = False
                                     Call W_GETxReprintCV(tW_TransType, tW_TmnNum, tW_TransNo)
                      End Select
                 Case Else
                      Select Case tRcvCode
                             Case "T018"
                                      bW_StaPrnCV2 = W_GETbCVIssue(tW_TransType, tW_TmnNum, tW_TransNo)
                                      bW_StaPrnCV3 = True
                                      Call W_GETxReprintCV(tW_TransType, tW_TmnNum, tW_TransNo)
                      End Select
             End Select
             '-------------------------------------------------------------------------------------------------------
            orsRC.MoveNext
         Loop
         orsRC.MoveFirst
    Else
        bW_StaPrnCV2 = False
        bW_StaPrnCV3 = False
    End If
    Exit Sub
ErrHandle:
    
End Sub

Public Property Get RePrint() As Boolean
    PrintEJ = bW_EnableRePrint
End Property

Public Property Let RePrint(ByVal pbEnable As Boolean)
    bW_EnableRePrint = pbEnable
End Property

Public Property Get RePrintNormal() As Boolean
    RePrintNormal = bW_EnableRePrintNormal
End Property

Public Property Let RePrintNormal(ByVal pbEnable As Boolean)
    bW_EnableRePrintNormal = pbEnable
End Property

Public Property Get RePrintCV1() As Boolean
    RePrintCV1 = bW_EnableRePrintCV1
End Property

Public Property Let RePrintCV1(ByVal pbEnable As Boolean)
    bW_EnableRePrintCV1 = pbEnable
End Property

Public Property Get RePrintCV2() As Boolean
    RePrintCV2 = bW_EnableRePrintCV2
End Property

Public Property Let RePrintCV2(ByVal pbEnable As Boolean)
    bW_EnableRePrintCV2 = pbEnable
End Property

Public Property Let TypePriner(ByVal ptType As String)
    tW_TypePrinter = ptType
End Property

Public Property Let ReIssueNo(ByVal ptReIssueNo As String)
    oW_ReIssue.tEN_ReIssueNo = ptReIssueNo
End Property

Public Property Let ReIssueExp(ByVal ptReIssueExp As String)
    oW_ReIssue.tEN_ReIssueExp = ptReIssueExp
End Property

Public Property Let ReIssueAmt(ByVal pcReIssueAmt As String)
    oW_ReIssue.cEN_ReIssueAmt = pcReIssueAmt
End Property

Public Property Let ReIssueTel(ByVal ptReIssueTel As String)
'----------------------------------------------------------
'Call:
'Cmt:'*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
'----------------------------------------------------------
    oW_ReIssue.tEN_ReIssueTel = ptReIssueTel
End Property

Public Property Let PhysicalTmnNum(ByVal ptPhysicalNum As String)
    tW_PhyTmnNum = ptPhysicalNum
End Property

'Public Function W_GETxSkuSign(ByVal ptNdMk As String, ByVal ptSKUCode As String, ByVal ptAlwSku As String) As String
''-------------------------------------------------------------------------------
''call:
''cmt : *Eaw 56-02-20 เปลี่ยนเครื่องหมายของ SKU
''-------------------------------------------------------------------------------
'    Dim orsTmp As ADODB.Recordset
'    Dim orsMall As ADODB.Recordset
'    Dim tSql As String
'    Dim tMemCard As String
'    Dim nDctNo As Long
'On Error GoTo ErrHandler
'    Set orsTmp = orsCD.Clone
'    orsTmp.Filter = "FNDctNo = '11' OR FNDctNo='12'"
'    If Not (orsTmp.EOF And orsTmp.BOF) Then
'        tMemCard = SP_FEDtChkString(orsTmp, "FTScdCardID")
'        tSql = "SELECT  TOP 1 TCNMMallCardMtn.FTMcdCode,TCNMMallCardSKU.FTSKUCode,TCNMMallCardSKU.FTMcsAlw"
'        tSql = tSql & vbCrLf & "FROM  (TCNMMallCardMtn INNER JOIN TCNMMallCardRange"
'        tSql = tSql & vbCrLf & " ON TCNMMallCardMtn.FTMcdCode=TCNMMallCardRange.FTMcdCode)"
'        tSql = tSql & vbCrLf & " INNER JOIN TCNMMallCardSKU ON TCNMMallCardMtn.FTMcdCode=TCNMMallCardSKU.FTMcdCode"
'        tSql = tSql & vbCrLf & " WHERE LEFT('" & tMemCard & "' ,TCNMMallCardMtn.FNMcdDigitChk) BETWEEN LEFT(TCNMMallCardRange.FTMcrRangeFrom,TCNMMallCardMtn.FNMcdDigitChk)"
'        tSql = tSql & vbCrLf & " AND LEFT(TCNMMallCardRange.FTMcrRangeTo,TCNMMallCardMtn.FNMcdDigitChk)"
'        tSql = tSql & vbCrLf & " AND LEN('" & tMemCard & "') = TCNMMallCardMtn.FNMcdDigits"
'        tSql = tSql & vbCrLf & " AND TCNMMallCardSKU.FTSKUCode='" & ptSKUCode & "'"
'        If ptAlwSku = "Y" Then
'            tSql = tSql & vbCrLf & " AND TCNMMallCardSKU.FTMcsAlw='1'"
'            Call SP_SQLvExecute(tSql, orsMall, oDB.ocnOnLine, , , , , True, True) 'read master
'            If Not (orsMall.EOF And orsMall.BOF) Then
'                W_GETxSkuSign = "*"
'            Else
'                W_GETxSkuSign = ptNdMk
'            End If
'        Else
'            tSql = tSql & vbCrLf & " AND TCNMMallCardSKU.FTMcsAlw='0'"
'            Call SP_SQLvExecute(tSql, orsMall, oDB.ocnOnLine, , , , , True, True) 'read master
'            If Not (orsMall.EOF And orsMall.BOF) Then
'                W_GETxSkuSign = ""
'            Else
'                W_GETxSkuSign = ptNdMk
'            End If
'        End If
'    Else
'        W_GETxSkuSign = ptNdMk
'    End If
'    Exit Function
'
'ErrHandler:
'    W_GETxSkuSign = ptNdMk
'End Function

'Public Function W_GETcRegPrice(ByVal pcSubReg As Double, ByVal pnTransType As Long, ByVal ptDate As String, ByVal ptTransTmn As String, ByVal ptTransNo As String) As Double
''-------------------------------------------------------------------------------
''call:
''cmt : *Eaw 56-02-20 คิดราคา Regular price ตามบัตร Member/Employee
''-------------------------------------------------------------------------------
'    Dim orsTmp As ADODB.Recordset
'    Dim tSql As String
'On Error GoTo ErrHandler
'    tSql = "SELECT  ISNULL(SUM(DT.FCSdtQty * DT.FCSdtSalePrice ) ,-99999)AS RegPrice "
'    tSql = tSql & vbCrLf & " FROM " & tW_TblDT & " AS DT  INNER JOIN " & tW_TblCD & " AS CD "
'    tSql = tSql & vbCrLf & " ON DT.FNSdtSeqNo = CD.FNSdtSeqNo"
'    tSql = tSql & vbCrLf & " AND DT.FTSkuCode = CD.FTSkuCode"
'    tSql = tSql & vbCrLf & " AND DT.FTTmnNum = CD.FTTmnNum"
'    tSql = tSql & vbCrLf & " AND DT.FTShdTransNo =  CD.FTShdTransNo"
'    tSql = tSql & vbCrLf & " AND DT.FTShdTransType = CD.FTShdTransType"
'    tSql = tSql & vbCrLf & " AND DT.FDShdTransDate = CD.FDShdTransDate"
'    tSql = tSql & vbCrLf & " WHERE CD.FTTmnNum = '" & ptTransTmn & "'"
'    tSql = tSql & vbCrLf & " AND CD.FTShdTransNo ='" & ptTransNo & "'"
'    tSql = tSql & vbCrLf & " AND CD.FTShdTransType = '" & Format(pnTransType, "00") & "'"
'    tSql = tSql & vbCrLf & " AND CD.FDShdTransDate ='" & ptDate & "'"
'    tSql = tSql & vbCrLf & " AND CD.FNDctNo='12'"
'    Call SP_SQLvExecute(tSql, orsTmp)
'    If Not (orsTmp.EOF And orsTmp.BOF) Then
'        W_GETcRegPrice = SP_FEDvChkNumeric(orsTmp, "RegPrice")
'        If W_GETcRegPrice < 0 Then W_GETcRegPrice = pcSubReg
'    Else
'        W_GETcRegPrice = pcSubReg
'    End If
'    Exit Function
'
'ErrHandler:
'    W_GETcRegPrice = pcSubReg
'End Function
Private Function C_CHKbCheckUpdatePoint(ByVal ptWhere As String) As Boolean
'---------------------------------------------------------------------------------------------
'
'   '*Ao 2014/02/11
'---------------------------------------------------------------------------------------------
Dim tSql As String
Dim orsMem As New ADODB.Recordset
Dim bChk As Boolean

On Error GoTo ErrHandler
    bChk = True
    
    tSql = "SELECT * FROM TPSTBPHis"
    tSql = tSql & vbCrLf & "WHERE " & ptWhere
    Set orsMem = ocnMember.Execute(tSql)
    If Not (orsMem.EOF And orsMem.BOF) Then
        bChk = True
    Else
        bChk = False
    End If
    C_CHKbCheckUpdatePoint = bChk
    Exit Function
ErrHandler:
    C_CHKbCheckUpdatePoint = True
End Function
'Private Sub C_UPDxUpdatePoint(ByVal ptWhere As String)
''---------------------------------------------------------------------------------------------
''
''   '*Ao 2014/02/11
''   '*Bump 57-03-06 copy มาจาก ImportText.exe V.4.0.9
''---------------------------------------------------------------------------------------------
'Dim dDate As Date
'Dim tVB_CNTerminalNum As String, tVB_Running As String, tVB_Transaction As String, tUT_SaleDate As String
'Dim tSql As String, tCardCode As String, tTransTime As String, tExpDate As String, tMcdExpDate As String, tChkExp As String
'Dim nCardPoint As Long, nEff As Long
'Dim cTotalPoint As Double, cBalance As Double
'Dim bChk As Boolean
'Dim tYearExp As String, tMonthExp As String
'
'Dim orsTmp As New ADODB.Recordset
'Dim orsMem As New ADODB.Recordset
'On Error GoTo ErrHandler
'
''    tSql = "SELECT * FROM TPSTSalePoint"
''    tSql = tSql & vbCrLf & "WHERE " & ptWhere
'
'    tSql = "SELECT FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate,FCSpoPoint,FTSpoID"
'    tSql = tSql & vbCrLf & ",FTSpoMemID,CONVERT(VARCHAR,FDDateUpd,012) AS FDDateUpd,FTWhoUpd"
'    tSql = tSql & vbCrLf & "FROM TPSTSalePoint"
'    tSql = tSql & vbCrLf & "WHERE " & ptWhere
'    tSql = tSql & vbCrLf & "ORDER BY FTSpoID"
'    Set orsTmp = ocnAll.Execute(tSql)
'    If Not (orsTmp.EOF And orsTmp.BOF) Then
'
'        If C_CHKbCheckUpdatePoint(ptWhere) Then     '************
'            Exit Sub
'        End If
'
'        orsTmp.MoveFirst
'        Do While Not orsTmp.EOF
'            tVB_CNTerminalNum = SP_FEDtChkString(orsTmp, "FTTmnNum")
'            tVB_Running = SP_FEDtChkString(orsTmp, "FTShdTransNo")
'            tVB_Transaction = SP_FEDtChkString(orsTmp, "FTShdTransType")
'            tUT_SaleDate = SP_DTEtStrSTD(SP_FEDdChkDate(orsTmp, "FDShdTransDate"))
'            tTransTime = Format(Time, "HH:MM:SS")
'            tCardCode = SP_FEDtChkString(orsTmp, "FTSpoMemID")
'            nCardPoint = SP_FEDvChkNumeric(orsTmp, "FCSpoPoint")
'
'            tExpDate = SP_FEDtChkString(orsTmp, "FDDateUpd")           '*  แปลงเป็น String แล้ว
'            If tExpDate = "" Then
'                tExpDate = Right(Format(Year(Date), "0000"), 2) & "12"              'Year & Month
'            Else
'                tExpDate = Left(tExpDate, 4)
'            End If
'            tChkExp = SP_FEDtChkString(orsTmp, "FTWhoUpd")
'            If tChkExp = "" Then
'                tChkExp = "N"
'            End If
'            Select Case SP_VALnText2Double(tVB_Transaction)             'IPV,Return จะเป็นการลบ
'                Case tEN_TrnReturn, tEN_TrnIPV, tEN_TrnCancelVoucher, tEN_TrnDepositCancel, tEN_TrnIPVRedeem, tEN_TrnAddDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax      '*Ao 2013/10/23   เพิ่ม tEN_TrnReturnTaxInc,tEN_TrnReturnManualTax
'                    nCardPoint = (-1) * nCardPoint
'                Case tEN_TrnIPVVIPService, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVOtherReceive, tEN_TrnOtherDisbursement
'                    nCardPoint = (-1) * nCardPoint
'                    'tEN_TrnOtherDisbursement       - POINT
'                    'tEN_TrnIPVOtherDisbursement +POINT
'            End Select
'            Select Case SP_VALnText2Double(tVB_Transaction)
'                Case tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnIPVRedeem, tEN_TrnAddDepositCancel, tEN_TrnRedeem
'                    tChkExp = "N"
'            End Select
'            If UCase(tChkExp) = UCase("Y") Then
'                tSql = "SELECT TOP 1 * FROM TCNMMallCard"
'                tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardCode & "'"
'                tSql = tSql & vbCrLf & "AND SUBSTRING(CONVERT(VARCHAR,FDMcdExpDate,012),1,4) = '" & tExpDate & "'"
'                tSql = tSql & vbCrLf & "ORDER BY FDMcdExpDate DESC"                             'อาจจะมีมากกว่า 1 Record
'            Else
'                tSql = "SELECT TOP 1 * FROM TCNMMallCard"
'                tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardCode & "'"
'                tSql = tSql & vbCrLf & "ORDER BY FDMcdExpDate DESC"                             'อาจจะมีมากกว่า 1 Record
'            End If
'            Set orsMem = ocnMember.Execute(tSql)
'            If Not (orsMem.EOF And orsMem.BOF) Then
'                bChk = True
'            Else
'                bChk = False
'            End If
'            If bChk Then
'                cTotalPoint = SP_FEDvChkNumeric(orsMem, "FCMcdTotalPoint")
'                tMcdExpDate = SP_DTEtStrSTD(SP_FEDdChkDate(orsMem, "FDMcdExpDate"))           '***
'                cBalance = SP_FEDvChkNumeric(orsMem, "FCBalance")
'
'                tSql = "UPDATE TCNMMallCard"
'                tSql = tSql & vbCrLf & "SET FCEarned = FCEarned + " & nCardPoint
'                tSql = tSql & vbCrLf & ",FCBalance = FCMcdTotalPoint + FCEarned +" & nCardPoint
'                tSql = tSql & vbCrLf & ",FTTimeUpd = '" & Format(Time, "HH:MM:SS") & "'"
'                tSql = tSql & vbCrLf & ",FTWhoUpd = 'ImportText'"
'                tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardCode & "'"
'                tSql = tSql & vbCrLf & "AND FDMcdExpDate = '" & tMcdExpDate & "'"
'                Call ocnMember.Execute(tSql, nEff)
'                If nEff > 0 Then
'                    bChk = True
'                Else
'                    bChk = False
'                End If
'            Else
'                dDate = DateSerial(Left(Format(Year(Date), "0000"), 2) & Left(tExpDate, 2), Right(tExpDate, 2) + 1, 0)
'                tMcdExpDate = SP_DTEtStrSTD(dDate)
'                tSql = "INSERT INTO TCNMMallCard(FTMcdStaAct,FTMcdCode,FTMcdName,FDMcdBirthday,FCMcdAmountPoint,FCMcdTotalPoint,FTMcdSta"
'                tSql = tSql & vbCrLf & ",FDMcdStaDate,FDMcdExpDate,FCEarned,FCBalance"
'                tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
'                tSql = tSql & vbCrLf & "VALUES('A','" & tCardCode & "','New Member','" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "',0,0,'A'"
'                tSql = tSql & vbCrLf & ",'" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "','" & tMcdExpDate & "'"
'                tSql = tSql & vbCrLf & "," & nCardPoint & "," & nCardPoint
'                tSql = tSql & vbCrLf & "," & cUT.UT_SQLtLastIns("ImportText")
'                Call ocnMember.Execute(tSql, nEff)
'                If nEff > 0 Then
'                    bChk = True
'                Else
'                    bChk = False
'                End If
'            End If
'            If bChk Then
'                tSql = "UPDATE TPSTSalePoint"
'                tSql = tSql & vbCrLf & "SET FTRemark = '1'"
'                tSql = tSql & vbCrLf & ",FDDateIns = '" & tMcdExpDate & "'"
'                tSql = tSql & vbCrLf & "WHERE " & ptWhere
''                tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
''                tSql = tSql & vbCrLf & "AND FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
''                tSql = tSql & vbCrLf & "AND FTSpoMemID = '" & tCardCode & "'"
'                Call ocnAll.Execute(tSql, nEff)
'                tSql = "INSERT INTO TPSTBPHis(FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate,FTShdTransTime,FTShdMemberID,FCCstCardPoint ,FCTotalBP"
'                tSql = tSql & vbCrLf & ",FCBalanceBP,FDDateUpd,FTTimeUpd)"
'                tSql = tSql & vbCrLf & "VALUES('" & tVB_CNTerminalNum & "','" & tVB_Running & "','" & tVB_Transaction & "','" & tUT_SaleDate & "','" & tTransTime & "'"
'                tSql = tSql & vbCrLf & ",'" & tCardCode & "'," & nCardPoint & "," & cTotalPoint & "," & cBalance & ",'" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "','" & Format(Time, "HH:MM:SS") & "')"
'                Call ocnMember.Execute(tSql, nEff)
'                'ยังไม่มี Process การลบข้อมูล ออกจาก Table TPSTBPHis        '*Ao 2013/01/11
'            End If
'            orsTmp.MoveNext
'        Loop
'    End If
'    Exit Sub
'ErrHandler:
'    MsgBox "C_UPDxUpdatePoint:" & Err.Number & vbCrLf & Err.Description
'    'tSql
'    Call SP_CRTxLogMallCardEvent("AdminMenu", "ImportText", "Command", "ErrNumber" & Err.Number & "Description" & Err.Description & "tSql" & tSql) '*Bump 57-03-06
'End Sub
Private Sub C_UPDxUpdatePoint(ByVal ptWhere As String)
'---------------------------------------------------------------------------------------------
'
'   '*Ao 2014/02/11
'   '*Bump 57-03-06 copy มาจาก ImportText.exe V.4.0.9
'---------------------------------------------------------------------------------------------
Dim dDate As Date
Dim tVB_CNTerminalNum As String, tVB_Running As String, tVB_Transaction As String, tUT_SaleDate As String
Dim tSql As String, tCardCode As String, tTransTime As String, tExpDate As String, tMcdExpDate As String, tChkExp As String
Dim nCardPoint As Long, nEff As Long
Dim cTotalPoint As Double, cBalance As Double
Dim bChk As Boolean
Dim tYearExp As String, tMonthExp As String

Dim orsTmp As New ADODB.Recordset
Dim orsMem As New ADODB.Recordset
Dim tCardNo As String '*Bmp 61-02-15 Token
Dim tTokenCardNo As String
Dim bChkToken As Boolean

On Error GoTo ErrHandler

'    tSql = "SELECT * FROM TPSTSalePoint"
'    tSql = tSql & vbCrLf & "WHERE " & ptWhere

    '* Bump 57-08-13 Comment    ให้ใช้ด้านล่าง การ Update Point ให้ Sum up ก่อน
'    tSql = "SELECT FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate,FCSpoPoint,FTSpoID"
'    tSql = tSql & vbCrLf & ",FTSpoMemID,CONVERT(VARCHAR,FDDateUpd,012) AS FDDateUpd,FTWhoUpd"
'    tSql = tSql & vbCrLf & "FROM TPSTSalePoint"
'    tSql = tSql & vbCrLf & "WHERE " & ptWhere
'    tSql = tSql & vbCrLf & "ORDER BY FTSpoID"

    '* Bump 57-08-13 การ Update Point ให้ Sum up ก่อน  เพื่อไม่ให้ duplicate ตอน INSERT INTO TPSTBPHis ด้านล่าง
    tSql = "SELECT FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate,SUM(FCSpoPoint) AS FCSpoPoint,MIN(FTSpoID) AS FTSpoID"
    tSql = tSql & vbCrLf & ",FTSpoMemID,CONVERT(VARCHAR,FDDateUpd,012) AS FDDateUpd,FTWhoUpd "
    tSql = tSql & vbCrLf & "FROM TPSTSalePoint"
    tSql = tSql & vbCrLf & "WHERE " & ptWhere
    tSql = tSql & vbCrLf & "GROUP BY FTSpoMemID,FDDateUpd,FTWhoUpd,FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate"
    tSql = tSql & vbCrLf & "ORDER BY FTSpoID"
    Set orsTmp = ocnAll.Execute(tSql)
    If Not (orsTmp.EOF And orsTmp.BOF) Then
        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
            '**********
            If C_CHKbCheckUpdatePoint(ptWhere) Then     '************ ตรวจสอบว่ามี  การ Update Member Point แล้ว
                Exit Sub
            End If
            orsTmp.MoveFirst
            Do While Not orsTmp.EOF
                tVB_CNTerminalNum = SP_FEDtChkString(orsTmp, "FTTmnNum")
                tVB_Running = SP_FEDtChkString(orsTmp, "FTShdTransNo")
                tVB_Transaction = SP_FEDtChkString(orsTmp, "FTShdTransType")
                tUT_SaleDate = SP_DTEtStrSTD(SP_FEDdChkDate(orsTmp, "FDShdTransDate"))
                tTransTime = Format(Time, "HH:MM:SS")
                tCardCode = SP_FEDtChkString(orsTmp, "FTSpoMemID")
                nCardPoint = SP_FEDvChkNumeric(orsTmp, "FCSpoPoint")
                '*Bmp 61-02-15 Token
                tCardNo = tCardCode
                tTokenCardNo = ""
                
                bChkToken = False
                If SP_PRCbToken("C_UPDxUpdatePoint", tCardNo, tTokenCardNo) Then
                    tCardNo = tTokenCardNo
                    bChkToken = True
                Else
                    bChkToken = False ''*ถ้า token ไม่ได้ ไม่ต้อง update
                End If
                '===================================
                If Not (bChkToken) Then ' = False ''*ถ้า token ไม่ได้ ไม่ต้อง update
                    tExpDate = SP_FEDtChkString(orsTmp, "FDDateUpd")           '*  แปลงเป็น String แล้ว
                    If tExpDate = "" Then
                        tExpDate = Right(Format(Year(Date), "0000"), 2) & "12"              'Year & Month
                    Else
                        tExpDate = Left(tExpDate, 4)
                    End If
                    tChkExp = SP_FEDtChkString(orsTmp, "FTWhoUpd")
                    If tChkExp = "" Then
                        tChkExp = "N"
                    End If
                    Select Case SP_VALnText2Double(tVB_Transaction)             'IPV,Return จะเป็นการลบ
                        Case tEN_TrnReturn, tEN_TrnIPV, tEN_TrnCancelVoucher, tEN_TrnDepositCancel, tEN_TrnIPVRedeem, tEN_TrnAddDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax      '*Ao 2013/10/23   เพิ่ม tEN_TrnReturnTaxInc,tEN_TrnReturnManualTax
                            nCardPoint = (-1) * nCardPoint
                        Case tEN_TrnIPVVIPService, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVOtherReceive, tEN_TrnOtherDisbursement
                            nCardPoint = (-1) * nCardPoint
                            'tEN_TrnOtherDisbursement       - POINT
                            'tEN_TrnIPVOtherDisbursement +POINT
                    End Select
                    Select Case SP_VALnText2Double(tVB_Transaction)
                        Case tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnIPVRedeem, tEN_TrnAddDepositCancel, tEN_TrnRedeem
                            tChkExp = "N"
                    End Select
                    If UCase(tChkExp) = UCase("Y") Then
                        tSql = "SELECT TOP 1 * FROM TCNMMallCard"
                        'tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardCode & "'"
                        tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardNo & "'" '*Bmp 61-02-15 Token
                        tSql = tSql & vbCrLf & "AND SUBSTRING(CONVERT(VARCHAR,FDMcdExpDate,012),1,4) = '" & tExpDate & "'"
                        tSql = tSql & vbCrLf & "ORDER BY FDMcdExpDate DESC"                             'อาจจะมีมากกว่า 1 Record
                    Else
                        tSql = "SELECT TOP 1 * FROM TCNMMallCard"
                        'tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardCode & "'"
                        tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardNo & "'" '*Bmp 61-02-15 Token
                        tSql = tSql & vbCrLf & "ORDER BY FDMcdExpDate DESC"                             'อาจจะมีมากกว่า 1 Record
                    End If
                    Set orsMem = ocnMember.Execute(tSql)
                    If Not (orsMem.EOF And orsMem.BOF) Then
                        bChk = True
                    Else
                        bChk = False
                    End If
                    If bChk Then
                        cTotalPoint = SP_FEDvChkNumeric(orsMem, "FCMcdTotalPoint")
                        tMcdExpDate = SP_DTEtStrSTD(SP_FEDdChkDate(orsMem, "FDMcdExpDate"))           '***
                        cBalance = SP_FEDvChkNumeric(orsMem, "FCBalance")
                        
                        tSql = "UPDATE TCNMMallCard" '+-*/
                        tSql = tSql & vbCrLf & "SET FCEarned = FCEarned + " & nCardPoint
                        tSql = tSql & vbCrLf & ",FCBalance = FCMcdTotalPoint + FCEarned +" & nCardPoint
                        tSql = tSql & vbCrLf & ",FTTimeUpd = '" & Format(Time, "HH:MM:SS") & "'"
                        tSql = tSql & vbCrLf & ",FTWhoUpd = 'ImportText'"
                        'tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardCode & "'"
                        tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardNo & "'" '*Bmp 61-02-15 Token
                        tSql = tSql & vbCrLf & "AND FDMcdExpDate = '" & tMcdExpDate & "'"
                        Call ocnMember.Execute(tSql, nEff)
                        If nEff > 0 Then
                            bChk = True
                        Else
                            bChk = False
                        End If
                    Else
                        dDate = DateSerial(Left(Format(Year(Date), "0000"), 2) & Left(tExpDate, 2), Right(tExpDate, 2) + 1, 0)
                        tMcdExpDate = SP_DTEtStrSTD(dDate)
                        tSql = "INSERT INTO TCNMMallCard(FTMcdStaAct,FTMcdCode,FTMcdName,FDMcdBirthday,FCMcdAmountPoint,FCMcdTotalPoint,FTMcdSta"
                        tSql = tSql & vbCrLf & ",FDMcdStaDate,FDMcdExpDate,FCEarned,FCBalance"
                        tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                        'tSql = tSql & vbCrLf & "VALUES('A','" & tCardCode & "','New Member','" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "',0,0,'A'"
                        tSql = tSql & vbCrLf & "VALUES('A','" & tCardNo & "','New Member','" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "',0,0,'A'" '*Bmp 61-02-15 Token tCardNo
                        tSql = tSql & vbCrLf & ",'" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "','" & tMcdExpDate & "'"
                        tSql = tSql & vbCrLf & "," & nCardPoint & "," & nCardPoint
                        tSql = tSql & vbCrLf & "," & cUT.UT_SQLtLastIns("ImportText")
                        Call ocnMember.Execute(tSql, nEff)
                        If nEff > 0 Then
                            bChk = True
                        Else
                            bChk = False
                        End If
                    End If
                    If bChk Then
                        tSql = "UPDATE TPSTSalePoint"
                        tSql = tSql & vbCrLf & "SET FTRemark = '1'"
                        tSql = tSql & vbCrLf & ",FDDateIns = '" & tMcdExpDate & "'"
                        tSql = tSql & vbCrLf & "WHERE " & ptWhere
        '                tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
        '                tSql = tSql & vbCrLf & "AND FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
        '                tSql = tSql & vbCrLf & "AND FTSpoMemID = '" & tCardCode & "'"
                        Call ocnAll.Execute(tSql, nEff)
                        Call SP_CRTxLogMallCardExtend("Date:" & Date & " Time:" & Time$ & " C_UPDxUpdatePoint: " & tSql) '*Tao 59-04-18
                        tSql = "INSERT INTO TPSTBPHis(FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate,FTShdTransTime,FTShdMemberID,FCCstCardPoint ,FCTotalBP"
                        tSql = tSql & vbCrLf & ",FCBalanceBP,FDDateUpd,FTTimeUpd)"
                        tSql = tSql & vbCrLf & "VALUES('" & tVB_CNTerminalNum & "','" & tVB_Running & "','" & tVB_Transaction & "','" & tUT_SaleDate & "','" & tTransTime & "'"
                        'tSql = tSql & vbCrLf & ",'" & tCardCode & "'," & nCardPoint & "," & cTotalPoint & "," & cBalance & ",'" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "','" & Format(Time, "HH:MM:SS") & "')"
                        tSql = tSql & vbCrLf & ",'" & tCardNo & "'," & nCardPoint & "," & cTotalPoint & "," & cBalance & ",'" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "','" & Format(Time, "HH:MM:SS") & "')" ''*Bmp 61-02-15 Token tCardNo
                        Call ocnMember.Execute(tSql, nEff)
                        'ยังไม่มี Process การลบข้อมูล ออกจาก Table TPSTBPHis        '*Ao 2013/01/11
                     Else
                        tSql = "UPDATE TPSTSalePoint"
                        tSql = tSql & vbCrLf & "SET FTRemark = '0'"
                        tSql = tSql & vbCrLf & ",FDDateIns = '" & tMcdExpDate & "'"
                        tSql = tSql & vbCrLf & "WHERE " & ptWhere
                        tSql = tSql & vbCrLf & "AND ISNULL(FTRemark,'') <> '1'"
                        Call SP_SQLvExecute(tSql)
                        Call SP_CRTxLogMallCardExtend("Date:" & Date & " Time:" & Time$ & " C_UPDxUpdatePoint: " & tSql) '*Tao 59-04-18
                    End If
                End If ' = False ''*ถ้า token ไม่ได้ ไม่ต้อง update
                orsTmp.MoveNext
            Loop
        End If
    End If
    Exit Sub
ErrHandler:
    MsgBox "C_UPDxUpdatePoint:" & Err.Number & vbCrLf & Err.Description
    'tSql
    Call SP_CRTxLogMallCardEvent("AdminMenu", "ImportText", "Command", "ErrNumber" & Err.Number & "Description" & Err.Description & "tSql" & tSql) '*Bump 57-03-06
End Sub

'*KT 66-08-30 CommSheet TTK-ISS3000-829_SIT-005 Wording show Pickup 0_00.01.00_pae_.xlsx
Private Function W_CHKbOldVatRate(Optional ByRef pcVatRateDT As Double = 0) As Boolean
'*Amorn
Dim orsTmpT As New ADODB.Recordset
Dim OTax, nTax As Integer
Dim TaxC As String
Dim tSql As String
On Error GoTo ErrHandler
    W_CHKbOldVatRate = False
    OTax = 0
    nTax = 0
    If tVB_DepositNoGST7 = "" Then Exit Function
    
    '*Bump 66-01-18 POSFront[6.2301.2].exe
    
'    tSql = "SELECT DISTINCT TOP 1 DT.FTTaxCode,DT.FCSdtTax,TA.FCTaxRate  FROM TPSTDepositDT DT WITH(NOLOCK) "
'    tSql = tSql & vbCrLf & "INNER JOIN TPSTDepositHD HD WITH(NOLOCK)  ON DT.FTShdTransNo=HD.FTShdTransNo "
'    tSql = tSql & vbCrLf & "INNER JOIN TCNMTaxMtn TA ON DT.FTTaxCode=TA.FTTaxCode "
'    tSql = tSql & vbCrLf & "WHERE HD.FTCdpDepositNo<> '' AND HD.FTCdpDepositNo= '" & tVB_DepositNoGST7 & "'"

'    Call SP_SQLvExecute(tSql, orsTmpT, oDB.ocnOnLine, , , , , True, True) 'read master
    
    tSql = "SELECT DISTINCT TOP 1 DT.FTTaxCode,DT.FCSdtTax  FROM TPSTDepositDT DT WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "INNER JOIN TPSTDepositHD HD WITH(NOLOCK)  ON DT.FTShdTransNo=HD.FTShdTransNo "
    tSql = tSql & vbCrLf & " AND (DT.FTTmnNum = HD.FTTmnNum ) AND ( DT.FDShdTransDate = HD.FDShdTransDate)" '*Bump 66-02-07 [6.2302.1] [66/02/09] GST Rate
    tSql = tSql & vbCrLf & " WHERE (DT.FCSdtTax <> 0) AND (HD.FTCdpDepositNo<> '') AND (HD.FTCdpDepositNo= '" & tVB_DepositNoGST7 & "')"
    
    Call SP_SQLvExecute(tSql, orsTmpT, oDB.ocnOnLine) 'Pconnect database server '*Bump 66-01-18
    
    If Not (orsTmpT.BOF And orsTmpT.EOF) Then
        OTax = SP_FEDvChkNumeric(orsTmpT, "FCSdtTax")
        TaxC = SP_FEDtChkString(orsTmpT, "FTTaxCode")
'        nTax = SP_FEDvChkNumeric(orsTmpT, "FCTaxRate") '*Bump 66-01-18
    Else
        Exit Function
   End If
   'cVB_CNGSTRate
'    If OTax <> nTax Then
     If OTax <> cVB_CNGSTRate Then '*Bump 66-01-18
        W_CHKbOldVatRate = True
        pcVatRateDT = OTax
        Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:cTransection Events:W_CHKbOldVatRate: FTCdpDepositNo:" & tVB_DepositNoGST7 & ",GST Rate:" & OTax & ", Master GST Rate:" & cVB_CNGSTRate) '*Bump 66-01-18
    End If
    Exit Function
ErrHandler:
    
End Function
