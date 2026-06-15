VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPSDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private bC_ChkEMP As Boolean                    '*Ao 2013/10/29   'ใช้ INSERT EMP
Private cC_DisEMP As Double
Private tC_TmnNum As String 'หมายเลขเครื่อง
Private tC_TransNo As String 'เลขที่เอกสาร
Private dC_TranDate As Date 'วันที่เอกสาร
Private nC_DTSeqNo As Integer 'ลำดับรายการสินค้า
Private cC_Qty As Double
Private cC_QtyAF As Double

Private cC_Amt As Double
Private cC_AmtAF As Double
Private cC_AmtAFAvi As Double

Private cC_DiscPer As Double
Private cC_DiscAmt As Double
Private cC_DiscNew As Double
Private cC_DiscPerOrg As Double
Private cC_DiscAmtOrg As Double
Private cC_DiscNewOrg As Double

Private nC_ENLnDisType As eEN_LineDiscType 'ประเภทการลด
Private nC_ENDiscType As eEN_DiscType 'ประเภทรายการลด
Private nC_SptSeqNo As Integer 'ลำดับรายการสินค้าที่ถูก Split

Private cC_LastDiscount As Double
Private cC_LastDiscountPerUnit As Double
Private tC_TheBestType As String 'เก็บประเภทส่วนลด ก่อนหน้าที่เป็นแบบ Thebest เช่น 2,3,8

'Private cC_B4DiscChg As Double
'Private cC_B4DiscChgAvi As Double
'Private cC_B4DiscChgQty As Double
'Private cC_B4DiscChgQtyAvi As Double
Private nC_ActSeqNo As Integer

Private tC_CardID As String 'หมายเลขบัตร
Private tC_CusID As String 'หมายเลขสมาชิก
Private tC_CardType As String 'ประเภทบัตร
Private tC_CardKey As String 'Key ของ Format บัตร
Private cC_DiscValue As Double 'ค่าที่เป็นจำนวนส่วนลด
Private cC_DiscAmtAlw As Double  'มูลค่าที่นำไปหาส่วนลด
Private tC_RefCode As String 'รหัสอ้างอิง ของ Deal
Private tC_EventID As String 'รหัสอ้างอิง ของ Deal

Private dC_CardExpDate As Date 'วันที่หมดอายุ

Private tC_StaChkExpire 'สถานะการตรวจสอบวันหมดอายุ
Private tC_StaChkDepart 'สถานะการตรวจสอบ Depart ของ SKU
Private cC_AlwDis As Double 'มูลค่าที่อนุญาตให้คำนวณส่วนลด
Private cC_DiscOntopAlw As Double 'มูลค่าที่อนุญาตให้คำนวณส่วนลด ontop
Private cC_AlwDisQty As Double
Private cC_TotalDisAmt As Double
Private cC_TotalDisQty As Double
Private cC_OntopFactor As Double
Private cC_Balance As Double
Private nC_DealGroup As Integer

Private bC_PromoAftDeal  As Boolean '*KT 52-11-25 Deal&Promo


Const tC_Msg001 = "กรุณากำหนดพร็อบเพอร์ตี้  tTerminalNo, tTransNo  และ dTranDate ให้ครบถ้วน ;Property tTerminalNo, tTransNo or dTranDate not found."
Const tC_Msg002 = "เกิดข้อผิดพลาดในการรวมข้อมูลส่วนลด;There are some error occer on discount summary."
Const tC_Msg003 = "เกิดข้อผิดพลาดในการเพิ่มข้อมูลส่วนลด;There are some error occer on add discount data."
Const tC_Msg004 = "ไม่สามารถให้ส่วนลดเกิน 100% ได้;Unable to discount, the exceed discount is 100%."
Const tC_Msg005 = "ไม่สามารถให้ส่วนลดเกินจำนวนที่สามารถลดได้;Unable to discount more than available."
Const tC_Msg006 = "ไม่สามารถให้ส่วนลด Line Discount ได้;Line Discount Not Allowed."
Const tC_Msg007 = "ไม่สามารถให้ส่วนลด Line Markdown ได้;Line Markdown Not Allowed."
Const tC_Msg008 = "คูปองที่ท่านใช้ หมดอายุแล้ว;This coupon has been Expired."
Const tC_Msg009 = "ไม่สามารถให้ส่วนลดเกินจำนวนที่สามารถลดได้" & vbCrLf & "ระบบจะปรับส่วนลดให้เท่ากับค่าสูงสุดที่สามารถลดได้ คุณต้องการดำเนินการต่อหรือไม่?;Unable to discount more than available." & vbCrLf & "Discount will be regenerate to maximum posible value. Continue?"
Const tC_Msg010 = "ไม่สามารถเปลี่ยนแปลงราคาใหม่ได้;New Price Exceeds Current Price."
Const tC_Msg011 = "เครื่องจุดขายนี้ไม่อนุญาตส่วนลด On Top;This terminal not allow On Top discount"

Private ass As Integer
'Recalcurate ----------------------------------------
Dim atCollectDT() As String
Dim atCollectCD() As String
Dim bC_StaRecal As Boolean

Dim bC_OnTopDepAlw As Boolean 'เงื่อนไขการคิด Ontop
Dim bC_OnTopUseReg  As Boolean 'เงื่อนไขการคิด Ontop เฉพาะ Reg Price
Dim bC_OnTopUsePro  As Boolean ''เงื่อนไขการคิด Ontop เฉพาะ Pro Price


'*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
Dim bC_POSNoOnTopFlag As Boolean ''เงื่อนไขเครื่องจุดขายที่ได้ OnTop
Dim tC_POSNoOnTopAlwType As String ''เงื่อนไขเครื่องจุดขายที่ได้ OnTop แบบ Terminal Num หรือ Floor
Dim tC_FloorAlwOnTop As String ''เงื่อนไขเครื่องจุดขายที่ได้ OnTop แบบ Terminal Num หรือ Floor
Dim cC_SubTotalAll As Double 'Sub Total All
Dim cC_SubTotal As Double 'SubTotalAll - Discount
Dim tC_StaOnTopAlwTierAmtLast As String 'Y : Allow Calculate from Last Tier N: Not Allow Calculate from Last Tie   *** Tier Amount only

Dim aW_MallCard() As MALLPRORATE
'
Public bCIsDisc As Boolean

Private Type U_CollectCD
    FTStmCode As String
    FTTmnNum As String
    FTShdTransNo As String
    FNSdtSeqNo As Integer
    FNScdSeqNo As Integer
    FTSkuCode As String
    FTShdTransType As String
    FDShdTransDate As Date
    FTSdtBarCode As String
    FTSdtStkCode As String
    FCSdtStkFac As Double
    FTSkuDepCode As String
    FTSkuBrdCode As String
    FTPunCode As String
    FCScdB4DisChgQty As Double
    FCScdB4DisChgQtyAvi As Double
    FCScdQty As Double
    FCScdQtyAll As Double
    FCScdAftDisChgQty   As Double
    FCScdAftDisChgQtyAvi    As Double
    FCScdB4DisChg As Double
    FCScdB4DisChgAvi    As Double
    FCScdDisRate    As Double
    FCScdDis As Double
    FCScdDisAll As Double
    FCScdAftDisChg  As Double
    FCScdAftDisChgAvi   As Double
    FNDctNo As Integer
    FTDctName As String
    FTDctRefCode As String
    FTSdtStaPdt  As String
    FTScdEventID As String
    FNScdActSeqNo   As Integer
    FTScdCardID  As String
    FCScdDiscValue  As Double
    FDScdExpDate    As Date
    FTScdCardType    As String
    FNScdDiscType   As Integer
End Type
Dim auCD() As U_CollectCD
Private Type U_CollectEmpDis
    cC_DiscAmt As Double
    nFNSdtSeqNo As Integer
    tEventID As String
    cFCSdtQtyAllDisAvi As Double
    cSourceAmt As Double
    cAlwDisQty As Double
End Type
Dim aEmpDis() As U_CollectEmpDis
'*Ao 51-06-19
Dim cC_QtyAll As Double
Dim cC_EmpMaxDis As Double
Dim cC_SpecialPrice As Double
Private bW_ChkMallCard As Boolean '*Ao 51-05-08
Private cC_AmtAllMall As Double          '*Ao 51-07-30 หาส่วนทีเอาไปลดจริง

Private Type UDT_CollectCPN
    udt_cDiscAmt  As Double
    udt_nSeqNo As Integer
    udt_tEventID As String
    udt_cQtyAllAvi As Double
    udt_cAlwSource As Double
End Type

Private Type UDT_EmpCard
    udt_nSeqNo As Integer
    udt_tSkuCode As String
    udt_tDptCode As String
    udt_cCalRate As Double
    udt_cQtyAllAvi As Double
    udt_cB4CalAmt As Double
    udt_cB4UnCalAmt As Double
    udt_cB4Disc As Double
    udt_cB4CalBalance As Double
    udt_cB4AllCalBalance As Double
    udt_cAFCalAmt As Double
    udt_cAFUnCalAmt As Double
    udt_cAFDisc As Double
    udt_cAFCalBalance As Double
    udt_cAFAllCalBalance As Double
    udt_cAllLimit As Double
    udt_cMallCrdSkuRate As Double
    udt_tStaDisMallCrd As String
    udt_tStaDisSKUOrDpt  As String '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
End Type

Private bC_FlagInsCD As Boolean '*KT 52-01-52 Flag ในการบอกว่ามีการ Insert Discount ลง Table CD สำเร็จ Defalse = False

'mall card discount
Private cC_AmtMallMax As Double
Private cC_AmtMallMin  As Double
Private tC_MallRegAlw As String
Private tC_MallProAlw As String
Private dC_MallEndDate As Date
Private dC_MallStartDate As Date
Private tC_MallFrmAmt As String
Private tC_DiscMallAlw As String
Private tC_MallAlwPoint As String

'Recal by item
Private bC_RecalByItem As Boolean
Private nC_ScdSeqNo  As Long
Private nC_RecalActSeqNo As Long

Private bC_ReConnectMember As Boolean
Private tC_CardName As String '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
Private bC_IsLastDeal As Boolean '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls

'Preview Day
Private bC_IsPreviewDay As Boolean

Private oC_RsTempOnTop  As ADODB.Recordset

Private Type AutoTran
    tSKUCode As String
    tDeptCode As String
    tBndCode As String
    nSeqNo As Long
    cAmt As Double
    tDiscountID As String
    tAutoTranID As String
    cQtyAllDisAvi As Double
    tEventID As String
End Type
Private aW_AutoTran() As AutoTran
Private aW_AutoTranTemp() As String

'Mall Coupon
Private bC_MallCpnInStoreBar As Boolean
Private tC_Emplimit As String
Private tC_EmpRemark As String

Private Sub C_EMPxDisEmp(ByVal pcDisc As Double, ByVal pnSeqNo As Integer, ByVal ptEventID As String, ByVal pcQtyAvi As Double, ByVal pcAlwCal As Double, ByVal pcDiscPer As Double, ByRef pcDisSumAmt As Double, ByRef pcDisSumAmtEx As Double, ByRef pcDisEMP As Double, Optional ByVal pbChk As Boolean = False)
'----------------------------------------------------------------------------------
'Call:
'Cmt:       '*Ao 2013/10/29
'----------------------------------------------------------------------------------
Dim tEventNo As String
    
    cC_DiscAmt = pcDisc
    tEventNo = ptEventID
    If cC_DiscAmt > 0 Then
        pcDisSumAmt = pcDisSumAmt + cC_DiscAmt
        '*Ao 2011/06/22  CommSheet ML-PRV-061 Employee Disc Over limit.xls
        pcDisSumAmtEx = pcDisSumAmtEx + SP_FMTcCurPoint(cC_DiscAmt, True, False)
        If pbChk = True Then
            If pcDisSumAmtEx > pcDisSumAmt Then
                cC_DiscAmt = cC_DiscAmt - (pcDisSumAmtEx - pcDisSumAmt)
            Else
                If pcDisSumAmtEx < pcDisSumAmt Then
                    cC_DiscAmt = cC_DiscAmt + (pcDisSumAmt - pcDisSumAmtEx)
                End If
            End If
        End If
        '______________________________________________________________________________________
        cC_DiscPer = pcDiscPer
'        Call C_DATxInsertDiscCD(cC_DiscAmt, pnSeqNo, tEventNo, pcQtyAvi, pcAlwCal, 0)
        pcDisEMP = pcDisEMP + SP_FMTcCurPoint(cC_DiscAmt, True, False)
    Else
        pcDisSumAmt = pcDisSumAmt + cC_DiscAmt
        cC_DiscPer = pcDiscPer
'        Call C_DATxInsertDiscCD(cC_DiscAmt, pnSeqNo, tEventNo, pcQtyAvi, pcAlwCal, 0)
    End If
End Sub

Private Sub W_DATxUpdateDT(ByVal pnDTSeq As Integer, ByVal ptSKUBar As String)
'--------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    Update Coop flag ให้ SKU ในตาราง DT '*Eaw 56-08-26 CommSheet TK-ISS3000-317-UAT-FAPL-93.xls
'--------------------------------------------------------------------------------------------------
On Error GoTo Err_Handler
    Dim tSql As String
    tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) "
    tSql = tSql & vbCrLf & "SET FTSdtTaxInv= 'Y' "
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & "AND FNSdtSeqNo ='" & pnDTSeq & "'"
    tSql = tSql & vbCrLf & "AND FTSkuCode = '" & ptSKUBar & "'"
    Call SP_SQLbExecute(tSql)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function W_GETxDiscEmpCard() As Double
'--------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Eaw 56-05-30 sum รวม Discount ใหม่จาก database  CommSheet TK-ISS3000-211-ST 00 00 01.xls
'--------------------------------------------------------------------------------------------------
On Error GoTo Err_Handler
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    
    tSql = "SELECT SUM(FCScdAmt) AS FCScdAmt "
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD & " WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & "AND  ISNULL(FTSdtStkCode,'') <> '1'"
    tSql = tSql & vbCrLf & " AND FNDctNo ='12'"
    Call SP_SQLbExecute(tSql, orsTemp)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        W_GETxDiscEmpCard = SP_FEDvChkNumeric(orsTemp, "FCScdAmt")
    End If
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_GETcDiscMallCardEx(ByVal porsTemp As ADODB.Recordset) As Double
'--------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 2013/03/08    TK
'--------------------------------------------------------------------------------------------------
Dim aEmpCard() As UDT_EmpCard

Dim orsMallSKUAlw As ADODB.Recordset
Dim orsMallDptAlw As ADODB.Recordset
'Dim orsSum As ADODB.Recordset
Dim cAllPrice As Double  '*Eaw 56-06-28

Dim nI As Integer, nFNSdtSeqNo As Integer, nSeqNoLast As Integer, nMall As Integer
Dim tSql  As String, tFTSkuCode As String, tFTSkuDepCode As String, tFTSkuBrdCode As String, tSKUBlue As String
Dim bMallSKUAlw As Boolean, bMallDptAlw As Boolean, bMallAlw As Boolean, bEmpCard As Boolean, bChk As Boolean

Dim cFCSdtQtyAllDisAvi As Double, cSaleAmt As Double, cSkuRate As Double, cTempSum As Double, cTempQty As Double
Dim cDptRate As Double, cBlueYellowAmt As Double, cTempRegSum As Double, cTempProSum As Double, cSumArray As Double
Dim cBluePrice As Double, cYellowPrice As Double, cDisSumAmt As Double, cDiscHD As Double, cCalAmt As Double

Dim nEff As Long

On Error GoTo Err_Handler
    Select Case tVB_CNSmnID '*Eaw 56-07-03 CommSheet TK-ISS3000-263.xls
        Case "1809" 'Member Card
            nC_ENLnDisType = 1
        Case "1887"  'TC Card
            nC_ENLnDisType = 2
    End Select
'    nC_ENLnDisType = nEN_LnDis1Per
    cC_DiscPer = cC_DiscValue
    cC_DiscAmt = 0
    cC_DiscNew = 0
    cC_DiscPerOrg = cC_DiscPer
    cC_DiscAmtOrg = cC_DiscAmt
    cC_DiscNewOrg = cC_DiscNew
    
    If cC_DiscValue = 0 Then
'        Call C_DATxInsertDiscCD(0, 1, tC_CardKey, 0, 0, 0, , , 0)          '*Ao 2013/12/17   แก้ไข ไม่เก็บค่า  tC_CardKey
        Call C_DATxInsertDiscCD(0, 1, "", 0, 0, 0, , , 0)
        Call DISxUpdateDiscTable(0)
        Exit Function
    End If          '*Ao 2013/12/06    Check  TCNMMallCardMtn.FDMcdStartDate,FDMcdEndDate
    If Not (Format(SP_GETdChkDateDisc, "yyyy/MM/dd") >= Format(dC_MallStartDate, "yyyy/MM/dd") And Format(SP_GETdChkDateDisc, "yyyy/MM/dd") <= Format(dC_MallEndDate, "yyyy/MM/dd")) Then
'        Call C_DATxInsertDiscCD(0, 1, tC_CardKey, 0, 0, 0, , , 0)
        Call C_DATxInsertDiscCD(0, 1, "", 0, 0, 0, , , 0)                               '*Ao 2013/12/17   แก้ไข ไม่เก็บค่า  tC_CardKey
        Call DISxUpdateDiscTable(0)
        Exit Function
    End If
    '*KT 52-02-18 mall card discount dpt/sku allowed
    If UCase(tC_StaChkDepart) = "Y" Then
        tSql = "SELECT MM.FTMcdCode, MK.FTSkuCode, MK.FTMcsAlw,MK.FDMcsStartDate,MK.FDMcsEndDate,MK.FCMcsRate,MK.FCMcsAmtMin,MK.FCMcsAmtMax"
        tSql = tSql & vbCrLf & "FROM TCNMMallCardMtn MM WITH(NOLOCK)  INNER JOIN"
        tSql = tSql & vbCrLf & "TCNMMallCardSKU MK WITH(NOLOCK)  ON MM.FTMcdCode = MK.FTMcdCode"
        tSql = tSql & vbCrLf & "WHERE MM.FTMcdCode = '" & tVB_DiscMallCardKey & "'"
        tSql = tSql & vbCrLf & "AND MK.FTMcsAlwPoint = '0'"
        Call SP_SQLvExecute(tSql, orsMallSKUAlw, oDB.ocnOnLine, , , , , True, True) 'read master
        If Not (orsMallSKUAlw.BOF And orsMallSKUAlw.EOF) Then
            bMallSKUAlw = True
        Else
            bMallSKUAlw = False
        End If
        tSql = "SELECT MM.FTMcdCode, MD.FTMpcCode,MD.FTMpcBrdCode,MD.FTMpcAlw,MD.FDMpcStartDate,MD.FDMpcEndDate,MD.FCMpcRate,MD.FTMpcReg,MD.FTMpcPro,MD.FCMpcAmtMin,MD.FCMpcAmtMax"
        tSql = tSql & vbCrLf & "FROM TCNMMallCardMtn MM WITH(NOLOCK)  INNER JOIN"
        tSql = tSql & vbCrLf & "TCNMMallCardDep MD WITH(NOLOCK)  ON MM.FTMcdCode = MD.FTMcdCode"
        tSql = tSql & vbCrLf & "WHERE MM.FTMcdCode = '" & tVB_DiscMallCardKey & "'"
        tSql = tSql & vbCrLf & "AND MD.FTMpcAlwPoint = '0'"
        Call SP_SQLvExecute(tSql, orsMallDptAlw, oDB.ocnOnLine, , , , , True, True) 'read master
        If Not (orsMallDptAlw.BOF And orsMallDptAlw.EOF) Then
            bMallDptAlw = True
        Else
            bMallDptAlw = False
        End If
    End If
    '---------------------------------------------------------------------------------------------------------------------------------
    '1 not allow sku
    '2 allow sku
    '3 not allow dpt
    '4 allow dpt
    '---------------------------------------------------------------------------------------------------------------------------------
'    porsTemp.MoveFirst
''    Set orsSum = porsTemp.Clone  '*KT 52-02-18 mall card discount dpt/sku allowed
'    If Not (porsTemp.BOF And porsTemp.EOF) Then
'        Set orsSum = New ADODB.Recordset
'        For nI = 1 To porsTemp.Fields.Count
'            Select Case UCase(Mid(porsTemp.Fields(nI - 1).Name, 2, 1))
'                Case "T"
'                    Call orsSum.Fields.Append(porsTemp.Fields(nI - 1).Name, porsTemp.Fields(nI - 1).Type, 256)
'                Case Else
'                    Call orsSum.Fields.Append(porsTemp.Fields(nI - 1).Name, porsTemp.Fields(nI - 1).Type, porsTemp.Fields(nI - 1).DefinedSize)
'            End Select
'        Next nI
'        orsSum.CursorType = adOpenStatic
'        orsSum.Open
'    End If
    ReDim aW_MallCard(0)    'clear array
    aW_MallCard(0).MALLAMT = 0
    aW_MallCard(0).MALLRATE = 0
    If Not (porsTemp.BOF And porsTemp.EOF) Then
        porsTemp.Sort = "FCSdtSaleAmt ASC"           '*Ao 2011/03/21  Communication Sheet PH3-ML-SB-134 Mall Card  discount amount minus in last item.xls
        porsTemp.MoveFirst
    End If
    
    ReDim aEmpCard(0) As UDT_EmpCard
    nI = 0
    
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        nI = nI + 1
        ReDim Preserve aEmpCard(nI) As UDT_EmpCard
        nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
        tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
        tFTSkuDepCode = SP_FEDtChkString(porsTemp, "FTSkuDepCode")
        tFTSkuBrdCode = SP_FEDtChkString(porsTemp, "FTSkuBrdCode")
        cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
        cSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
        cSkuRate = CDbl(SP_FEDvChkNumeric(porsTemp, "FCSkuMallCardRate"))
        
        aEmpCard(nI).udt_nSeqNo = nFNSdtSeqNo
        aEmpCard(nI).udt_tSkuCode = tFTSkuCode
        aEmpCard(nI).udt_tDptCode = tFTSkuDepCode
        aEmpCard(nI).udt_cQtyAllAvi = cFCSdtQtyAllDisAvi
        
        Select Case C_CHKnDDP(nFNSdtSeqNo) 'Check Double Policy
            Case 0 'ไม่อนุญาต
                'คิดเฉพาะส่วนที่เหลือ
                bEmpCard = False
            Case 1 'อนุญาต
                bEmpCard = True
            Case 2 'หาผลที่ดีที่สุด
               bEmpCard = True
        End Select
        
        bMallAlw = True
        If bEmpCard Then
            If UCase(tC_StaChkDepart) = "Y" Then 'ตรวจสอบระดับ Allow และ Not Allow ของ Dept Code/Sku Code
                If bMallSKUAlw Then 'ระดับ  Sku Code
                    orsMallSKUAlw.Filter = adFilterNone
                    orsMallSKUAlw.Filter = "FTSkuCode = '" & tFTSkuCode & "'"
                    If Not (orsMallSKUAlw.BOF Or orsMallSKUAlw.EOF) Then
                        'not allow sku/allow sku
                        If C_EMPbGetSKURate(orsMallSKUAlw, cSaleAmt, cDptRate, cBlueYellowAmt) Then
                            aEmpCard(nI).udt_cCalRate = cDptRate
                            aEmpCard(nI).udt_cB4CalAmt = cBlueYellowAmt
                            aEmpCard(nI).udt_cB4UnCalAmt = 0
                            aEmpCard(nI).udt_tStaDisMallCrd = "0"
                            aEmpCard(nI).udt_tStaDisSKUOrDpt = "1"  'SKU
                            bMallAlw = False
                        Else
                            bMallAlw = True
                        End If
                    Else
                        bMallAlw = True
                    End If
                End If
                If bMallAlw Then
                    If bMallDptAlw Then 'ระดับ Dpt
                        orsMallDptAlw.Filter = adFilterNone
'                        orsMallDptAlw.Filter = "(FTMpcCode = '" & tFTSkuDepCode & "' AND FTMpcBrdCode ='" & tFTSkuBrdCode & "') OR (FTMpcCode = '" & tFTSkuDepCode & "' AND FTMpcBrdCode ='') OR (FTMpcCode = '' AND FTMpcBrdCode = '" & tFTSkuBrdCode & "') OR (FTMpcCode = '" & tFTSkuDepCode & "')"
                        orsMallDptAlw.Filter = "(FTMpcCode = '" & tFTSkuDepCode & "' AND FTMpcBrdCode ='" & tFTSkuBrdCode & "') OR (FTMpcCode = '" & tFTSkuDepCode & "' AND FTMpcBrdCode ='') OR (FTMpcCode = '' AND FTMpcBrdCode = '" & tFTSkuBrdCode & "')"
                        orsMallDptAlw.Sort = "FTMpcCode DESC,FTMpcBrdCode DESC"
                        If Not (orsMallDptAlw.BOF Or orsMallDptAlw.EOF) Then
                            'not allow dpt/allow dpt
                            If C_EMPbGetRateDpt(orsMallDptAlw, cSaleAmt, cDptRate, cBlueYellowAmt) Then
                                aEmpCard(nI).udt_cCalRate = cDptRate
                                aEmpCard(nI).udt_cB4CalAmt = cBlueYellowAmt
                                aEmpCard(nI).udt_cB4UnCalAmt = 0
                                aEmpCard(nI).udt_tStaDisMallCrd = "0"
                                aEmpCard(nI).udt_tStaDisSKUOrDpt = "2"  'Dpt
                                bMallAlw = False
                            Else
                                bMallAlw = True
                            End If
                        Else
                            bMallAlw = True
                        End If
                    Else
                        bMallAlw = True
                    End If
                End If
            End If
        Else
            aEmpCard(nI).udt_cCalRate = 0
            aEmpCard(nI).udt_cB4CalAmt = 0
            aEmpCard(nI).udt_cB4UnCalAmt = 0
            aEmpCard(nI).udt_tStaDisMallCrd = "0"
            aEmpCard(nI).udt_tStaDisSKUOrDpt = "0"
            bMallAlw = False
        End If
        
        If bMallAlw Then
            If SP_FEDtChkString(porsTemp, "FTSkuStaMallEmpCardEx") = "Y" Then
'                If cSkuRate = 0 Then
'                    aEmpCard(nI).udt_cCalRate = cC_DiscValue                       '*Ao 2013/04/09  CommSheet TK-ISS3000-139-IT2 00.00.01.xlsx
'                Else
'                    aEmpCard(nI).udt_cCalRate = IIf(cC_DiscValue > cSkuRate, cSkuRate, cC_DiscValue)
'                End If
                aEmpCard(nI).udt_cCalRate = IIf(cC_DiscValue > cSkuRate, cSkuRate, cC_DiscValue)
                aEmpCard(nI).udt_cB4CalAmt = cSaleAmt
                aEmpCard(nI).udt_cB4UnCalAmt = 0
                aEmpCard(nI).udt_tStaDisMallCrd = "0"
                aEmpCard(nI).udt_tStaDisSKUOrDpt = "0"
            Else
                aEmpCard(nI).udt_cCalRate = 0
                aEmpCard(nI).udt_cB4CalAmt = 0
                aEmpCard(nI).udt_cB4UnCalAmt = 0
                aEmpCard(nI).udt_tStaDisMallCrd = "0"
                aEmpCard(nI).udt_tStaDisSKUOrDpt = "0"
            End If
        End If

        aEmpCard(nI).udt_cB4Disc = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalAmt * aEmpCard(nI).udt_cCalRate / 100, True, False)
        aEmpCard(nI).udt_cB4CalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalAmt - aEmpCard(nI).udt_cB4Disc, True, False)
        aEmpCard(nI).udt_cB4AllCalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalBalance + aEmpCard(nI).udt_cB4UnCalAmt, True, False)
        aEmpCard(nI).udt_cAFCalAmt = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalAmt, True, False)
        aEmpCard(nI).udt_cAFUnCalAmt = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4UnCalAmt, True, False)
        aEmpCard(nI).udt_cAFDisc = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4Disc, True, False)
        aEmpCard(nI).udt_cAFCalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalBalance, True, False)
        aEmpCard(nI).udt_cAFAllCalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cAFUnCalAmt, True, False)
        aEmpCard(nI).udt_cAllLimit = 0
        aEmpCard(nI).udt_cMallCrdSkuRate = cSkuRate
        '**********************************************************
        porsTemp.MoveNext
    Loop
    '*Eaw 56-06-28
     For nI = 1 To UBound(aEmpCard)
            cAllPrice = cAllPrice + SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalAmt)
    Next nI
    If cAllPrice < cC_AmtMallMin Then
'        Call C_DATxInsertDiscCD(0, 1, tC_CardKey, 0, 0, 0, , , 0)
        Call C_DATxInsertDiscCD(0, 1, "", 0, 0, 0, , , 0)                                   '*Ao 2013/12/17   แก้ไข ไม่เก็บค่า  tC_CardKey
        Call DISxUpdateDiscTable(0)
        ReDim aW_MallCard(0)    'clear array
        Exit Function
    End If
    If cAllPrice > cC_AmtMallMax Then  '*Eaw 56-07-01
        For nI = 1 To UBound(aEmpCard)
            aEmpCard(nI).udt_cB4CalAmt = (aEmpCard(nI).udt_cB4CalAmt * cC_AmtMallMax) / cAllPrice
            aEmpCard(nI).udt_cB4Disc = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalAmt * aEmpCard(nI).udt_cCalRate / 100, True, False)
            aEmpCard(nI).udt_cB4CalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalAmt - aEmpCard(nI).udt_cB4Disc, True, False)
            aEmpCard(nI).udt_cB4AllCalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalBalance + aEmpCard(nI).udt_cB4UnCalAmt, True, False)
            aEmpCard(nI).udt_cAFCalAmt = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalAmt, True, False)
            aEmpCard(nI).udt_cAFUnCalAmt = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4UnCalAmt, True, False)
            aEmpCard(nI).udt_cAFDisc = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4Disc, True, False)
            aEmpCard(nI).udt_cAFCalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalBalance, True, False)
            aEmpCard(nI).udt_cAFAllCalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cAFUnCalAmt, True, False)
            aEmpCard(nI).udt_cAllLimit = 0
            aEmpCard(nI).udt_cMallCrdSkuRate = cSkuRate
        Next nI
    End If
    '-------------------------------------------------------------------------------------------------------------------------------
    For nI = 1 To UBound(aEmpCard)
        bChk = False
        If UBound(aW_MallCard) > 0 Then
            For nMall = 0 To UBound(aW_MallCard) - 1
                If aW_MallCard(nMall).MALLRATE = aEmpCard(nI).udt_cCalRate Then
                    bChk = True
                    aW_MallCard(nMall).MALLAMT = CDbl(Val(Format(aW_MallCard(nMall).MALLAMT))) + aEmpCard(nI).udt_cB4CalAmt
                    Exit For
                End If
            Next nMall
        End If
        If Not bChk Then
            If aEmpCard(nI).udt_cCalRate > 0 Then
                aW_MallCard(UBound(aW_MallCard)).MALLAMT = aEmpCard(nI).udt_cB4CalAmt
                aW_MallCard(UBound(aW_MallCard)).MALLRATE = aEmpCard(nI).udt_cCalRate
                ReDim Preserve aW_MallCard(UBound(aW_MallCard) + 1)
            End If
        End If

        If aEmpCard(nI).udt_cB4Disc > 0 Then
            nSeqNoLast = aEmpCard(nI).udt_nSeqNo
        End If
    Next nI
    
    cDiscHD = 0
    For nMall = 0 To UBound(aW_MallCard) - 1
            cDiscHD = cDiscHD + SP_FMTcCurPoint(CDbl(Val(Format(aW_MallCard(nMall).MALLAMT))) * CDbl(Val(Format(aW_MallCard(nMall).MALLRATE))) / 100, True, False)
    Next nMall

    cDisSumAmt = 0
    cCalAmt = 0
    For nI = 1 To UBound(aEmpCard)
        If aEmpCard(nI).udt_nSeqNo = nSeqNoLast Then
            bChk = True
        Else
            bChk = False
        End If
        '____________________________________________
        cCalAmt = cCalAmt + aEmpCard(nI).udt_cB4CalAmt                  '*Ao 2013/07/31   เพิ่ม  tC_CardKey
        Call C_MALxInsMall(aEmpCard(nI).udt_cAFDisc, aEmpCard(nI).udt_nSeqNo, tC_CardKey, aEmpCard(nI).udt_cQtyAllAvi, aEmpCard(nI).udt_cAFCalAmt, aEmpCard(nI).udt_cCalRate, cDisSumAmt, aEmpCard(nI).udt_cMallCrdSkuRate, cDiscHD, bChk)
        
        If aEmpCard(nI).udt_cCalRate > 0 Then
            Call C_EMPxUpdateSingEmp(aEmpCard(nI).udt_nSeqNo)
        End If
        '_____________________________________
    Next nI
    Call DISxUpdateDiscTable(cCalAmt)
    
'    porsTemp.Filter = adFilterNone
'    porsTemp.MoveFirst
'    If Not (porsTemp.BOF And porsTemp.EOF) Then
'         porsTemp.Filter = "FTSdtMemberPointAlw = 'Y' AND FTSdtStaPdt14 <> '1' AND FTSkuStaMallEmpCardEx = 'Y'"
'         If Not (porsTemp.BOF And porsTemp.EOF) Then
'                  Do While Not porsTemp.EOF
'                         tSql = "UPDATE " & tVB_CNTblDT
'                         tSql = tSql & vbCrLf & "SET FCSdtBSubPro = " & SP_FEDvChkNumeric(porsTemp, "FCSdtBSubReg")
'                         tSql = tSql & vbCrLf & ",FCSdtDiscSubPro =" & SP_FEDvChkNumeric(porsTemp, "FCSdtDiscSubReg")
'                         tSql = tSql & vbCrLf & ",FCSdtBSubReg =" & SP_FEDvChkNumeric(porsTemp, "FCSdtBSubPro")
'                         tSql = tSql & vbCrLf & ",FCSdtDiscSubReg =" & SP_FEDvChkNumeric(porsTemp, "FCSdtDiscSubPro")
'                         tSql = tSql & vbCrLf & ",FCSdtBluePrice =" & SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
'                         tSql = tSql & vbCrLf & ",FCSdtYellowPrice =" & SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
'                         tSql = tSql & vbCrLf & ",FTSkuStaMallEmpCard ='N'"
'                         tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                         tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
'                         tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                         tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
'                         tSql = tSql & vbCrLf & " AND FTSkuStaMallEmpCard = 'Y'"
'                         nEff = 0
'                         Call SP_SQLvExecute(tSql, , oDB.ocnOffLine, , , nEff)
'                         If nEff = 0 Then
'                              Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[11]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
'                         End If
'                         porsTemp.MoveNext
'                  Loop
'         End If
'    End If
    If cDisSumAmt = 0 Then
'        Call C_DATxInsertDiscCD(0, 1, tC_CardKey, 0, 0, 0, , , 0)               '*Ao 2013/12/17   แก้ไข ไม่เก็บค่า  tC_CardKey
        Call C_DATxInsertDiscCD(0, 1, "", 0, 0, 0, , , 0)
    End If
    C_GETcDiscMallCardEx = cDisSumAmt
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub C_EMPxUpdateSingEmp(ByVal pnSeqNo As Integer)
'--------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 2013/03/08    TK
'--------------------------------------------------------------------------------------------------
Dim tSql As String
Dim bChk As Boolean
Dim orsTemp As ADODB.Recordset

On Error GoTo Err_Handler
    If nC_ENDiscType = nEN_Disc12EmpCrd Then
        tSql = "SELECT TOP 1.* FROM " & tVB_CNTblCD & " WITH(NOLOCK)  "
        tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & vbCrLf & "AND FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & "AND FNDctNo = 11"               'Member Discount
        If SP_SQLvExecute(tSql, orsTemp) = 0 Then
            If (orsTemp.BOF And orsTemp.EOF) Then
                bChk = True
            Else
                bChk = False
            End If
        End If
    Else
        bChk = True
    End If
    
    If bChk Then
        tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) "
        tSql = tSql & vbCrLf & "SET FTPdtNoDis = 'Y'"                   'Y    Allow Emp
        tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & vbCrLf & "AND FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & "AND FNSdtSeqNo = " & pnSeqNo
        Call SP_SQLbExecute(tSql)
    End If
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DISxUpdateDiscTable(ByVal pcDiscAmt As Double)
'---------------------------------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 2013/03/06
'---------------------------------------------------------------------------------------------
Dim tSql As String
Dim bChk  As Boolean
Dim orsTemp As ADODB.Recordset
Dim cSaleAmt As Double

On Error GoTo Err_Handler

'    If nC_ENDiscType = nEN_Disc12EmpCrd Then
'        tSql = "SELECT TOP 1.* FROM " & tVB_CNTblCD
'        tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
'        tSql = tSql & vbCrLf & "AND FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
'        tSql = tSql & vbCrLf & "AND FNDctNo = 11"               'Member Discount
'        If SP_SQLvExecute(tSql, orsTemp) = 0 Then
'            If (orsTemp.BOF And orsTemp.EOF) Then
'                bChk = True
'            Else
'                bChk = False
'            End If
'        End If
'    Else
'        If nC_ENDiscType = nEN_Disc11MallCrd Then '*Eaw 56-07-03
'            tSql = "SELECT TOP 1.* FROM " & tVB_CNTblCD
'            tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
'            tSql = tSql & vbCrLf & "AND FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
'            tSql = tSql & vbCrLf & "AND FNDctNo = 11"               'Member Discount
'            If SP_SQLvExecute(tSql, orsTemp) = 0 Then
'                If (orsTemp.BOF And orsTemp.EOF) Then
'                    bChk = True
'                Else
'                    bChk = False
'                End If
'            End If
'        End If
'    End If
    tSql = "SELECT FCShdWpTax FROM " & tVB_CNTblHD & " WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & "WHERE  FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "' AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND FCShdWpTax > 0"
    If SP_SQLvExecute(tSql, orsTemp) = 0 Then
        If (orsTemp.BOF And orsTemp.EOF) Then
            bChk = True
        Else
            bChk = False
        End If
    Else
        bChk = True
    End If

    If bChk Then
        tSql = "UPDATE " & tVB_CNTblHD & " WITH(ROWLOCK) "
        tSql = tSql & vbCrLf & "SET FCShdWpTax = " & SP_FMTcCurPoint(pcDiscAmt, True, False)
        tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
        Call SP_SQLbExecute(tSql)
    End If
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Public Function DISbChkAlwTranDisc(ByVal pnDiscType As Integer) As Boolean
'_____________________________
'   Call:
'   Cmt:'*Ao 2012/01/05   CommSheet ML-PH4-R3-016 Select Tran MarkDown and Tran Disc 1 time in Transaction.xls
'_____________________________
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim bChk As Boolean

On Error GoTo Err_Handler

    tSql = "IF EXISTS(SELECT FNScdSeqNo FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & "WHERE FNDctNo IN (" & pnDiscType & ")"
    tSql = tSql & vbCrLf & "AND " & tVB_CNWheBill & ")"
    tSql = tSql & vbCrLf & "SELECT 1 AS Result"
    tSql = tSql & vbCrLf & "ELSE"
    tSql = tSql & vbCrLf & "SELECT 0 AS Result"
    If SP_SQLbExecute(tSql, orsTemp) Then 'ตรวจสอบว่ามีส่วนลดตัวมันเองหรือไม่
        If SP_FEDvChkNumeric(orsTemp, "Result") = 1 Then 'มี
            bChk = False
        Else 'ไม่มี
            bChk = True
        End If
    End If
    DISbChkAlwTranDisc = bChk
Exit Function
Err_Handler:
    DISbChkAlwTranDisc = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Public Function DISbChkAlw(ByVal pnDiscType As Integer, ByVal pbChkCoop As Boolean) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # ตรวจสอบว่าส่วนลดปัจจุบันสามารถกดได้หรือไม่
'ARGUMENT LIST # pnDiscType = Discount Type
'                                         pbChkCoop = Check Cooperate Coupon (Brand)?
'RETURN VALUE  # True if Allow
'                                         False if Not Allow
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    On Error GoTo Err_Handler
    DISbChkAlw = Not C_CHKbUnseq(pnDiscType) 'ส่วนลดปัจจุบัน เรียงลำดับหรือไม่
    If pbChkCoop Then
        tSql = "IF EXISTS(SELECT FNScdSeqNo FROM " & tVB_CNTblCD
        tSql = tSql & vbCrLf & "WHERE FNDctNo IN (" & nEN_Disc09CoopCpn & "," & nEN_Disc10CoopBrand & ")"
        tSql = tSql & vbCrLf & "AND " & tVB_CNWheBill & ")"
        tSql = tSql & vbCrLf & "SELECT 1 AS Result"
        tSql = tSql & vbCrLf & "ELSE"
        tSql = tSql & vbCrLf & "SELECT 0 AS Result"
        If SP_SQLbExecute(tSql, orsTemp) Then 'ตรวจสอบว่ามี Cooperate Coupon หรือ Cooperate Brand หรือไม่
            If SP_FEDvChkNumeric(orsTemp, "Result") = 1 Then 'มี
                DISbChkAlw = False And DISbChkAlw
            Else 'ไม่มี
                DISbChkAlw = True And DISbChkAlw
            End If
        End If
    Else
        '*Ao 2012/01/05   CommSheet ML-PH4-R3-016 Select Tran MarkDown and Tran Disc 1 time in Transaction.xls
        Select Case pnDiscType
            Case nEN_Disc13TranDisc, nEN_Disc14TranMDown, nEN_Disc23TranMDown1, nEN_Disc24TranMDown2, nEN_Disc25TranMDown3, nEN_Disc26TranMDown4, nEN_Disc27TranMDown5
                DISbChkAlw = (DISbChkAlwTranDisc(pnDiscType)) And (DISbChkAlw)
            Case Else
            
        End Select
        '____________________________________
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    DISbChkAlw = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Public Sub C_UPDxLabel(Optional ByVal pbSpecial As Boolean = False, Optional ByVal pcAddReg As Double = 0, Optional ByVal pnDiscType As eEN_DiscType = 0, Optional ByVal pbRetReg As Boolean = False)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # ทำการคำนวณหาค่า Yellow Price และ Blue Price พร้อมทั้ง Update Flag MallCard ด้วย
'ARGUMENT LIST # pbSpecial = Call From Special Gift or Special CPN
'                                     # pcAddReg = Insert FCSdtCostEX When Call From Special Gift or Special CPN
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cProPrice As Double
    Dim cRegPrice As Double
    Dim nI As Integer
    Dim tSql As String
    Dim tWhe As String
    Dim cAddYellow As Double
    Dim orsDT As ADODB.Recordset
    Dim orsCD As ADODB.Recordset
    Dim orsTemp As ADODB.Recordset
    Dim cKeppAddReg As Double
    Dim cCostEx As Double
    Dim cBluePrice As Double
    Dim cTmp As Double
    Dim cSumCosEx As Double
On Error GoTo Err_Handler
    
    If pnDiscType <> 0 Then
        nC_ENDiscType = pnDiscType
    End If
    cKeppAddReg = pcAddReg
    
    tSql = "SELECT FNSdtSeqNo,FCScdAmt ,FNDctNo"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD & " WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & "AND  ISNULL(FTSdtStkCode,'') <> '1'" '*KT 52-11-25
    Call SP_SQLbExecute(tSql, orsTemp, , , , , , , True)
    Call SP_RSTbClone2OfflineEx2(orsTemp, orsCD) '*KT 54-03-25 ป้องกัน recordset เน่าพร้อม connection ที่ reconnect ใหม่
    
    tSql = "SELECT FNSdtSeqNo,FCSdtCostEX, FCSdtBluePrice"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
    tSql = tSql & vbCrLf & "ORDER BY  FNSdtSeqNo ASC" '*KT 54-12-02 V.4.0.8 ML-PH4-R3-015 Wrong Sub-Total Reg when paid with Tender not allow Mcard  and CoopCpn , CoopBnd
    Call SP_DATxRsClose(orsTemp)
    Call SP_SQLbExecute(tSql, orsTemp, , , , , , , True)
    Call SP_RSTbClone2OfflineEx2(orsTemp, orsDT) '*KT 54-03-25 ป้องกัน recordset เน่าพร้อม connection ที่ reconnect ใหม่
    If Not orsDT Is Nothing Then
        If Not (orsDT.BOF And orsDT.EOF) Then
                
    '*KT 52-03-06 ปรับปรุงให้ดีขึ้น
    '------------------------------------------------------------
            cCostEx = 0
            cBluePrice = 0
            orsDT.MoveFirst                         '*Aoo 2012/09/14
            Do While Not orsDT.EOF
                cTmp = CDbl(Val(Format(SP_FEDvChkNumeric(orsDT, "FCSdtCostEX"))))
'                If cCostEx < cTmp Then
'                    cCostEx = cTmp
'                End If
                cBluePrice = cBluePrice + SP_FEDvChkNumeric(orsDT, "FCSdtBluePrice")
                orsDT.MoveNext
            Loop
    '------------------------------------------------------------
            
    '*KT 52-03-06 ปรับปรุงให้ดีขึ้น
    '------------------------------------------------------------
            orsDT.Sort = "FNSdtSeqNo ASC" '*KT 54-12-02 V.4.0.8 ML-PH4-R3-015 Wrong Sub-Total Reg when paid with Tender not allow Mcard  and CoopCpn , CoopBnd
            orsDT.MoveFirst
            Do While Not orsDT.EOF
                pcAddReg = cKeppAddReg
                nI = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
                cRegPrice = 0: cProPrice = 0
                Call C_GETxAllowMallCard(nI, cRegPrice, cProPrice, pcAddReg, pbSpecial)
                If cRegPrice = 0 Then pcAddReg = 0
                tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) "   'tpsdt08107tmpontop
                tSql = tSql & vbCrLf & "SET "
                tSql = tSql & vbCrLf & "FCSdtBluePrice = " & SP_FMTcCurPoint(cRegPrice, True, False)
                tSql = tSql & vbCrLf & ",FCSdtYellowPrice = " & SP_FMTcCurPoint(cProPrice, True, False)
                ''*KT 51-09-15 แก้ไขเรื่องจ่ายด้วย Special Gift Voucher/Special Coupon หลังหักส่วนลด Mall Cardด้วยสินค้าป้ายน้ำเงินอย่างเดียว
                If pbSpecial And pcAddReg > 0 Then
                    tSql = tSql & vbCrLf & ",FCSdtCostEX = " & SP_FMTcCurPoint(IIf(pcAddReg > cBluePrice, cBluePrice, pcAddReg), True, False)
                Else
                    If pcAddReg > 0 Then            '*Ao 2011/06/28  CommSheet ML-PRV-062-SubTotal Reg Incorrect.xls
                        cSumCosEx = cSumCosEx + pcAddReg
                        tSql = tSql & vbCrLf & ",FCSdtCostEX = " & SP_FMTcCurPoint(cSumCosEx, True, False)
                    End If
                End If
                tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill & " And FNSdtSeqNo = " & nI
                If SP_SQLbExecute(tSql) Then
                    If nC_ENDiscType > nEN_Disc07AutoTran Then
                        cAddYellow = 0
'                        If Not orsCD Is Nothing Then               '*Ao 2011/03/26 ถ้า orsCD ไม่มีข้อมูล จะทำให้ Error
                        If orsCD.State = 1 Then
                            If Not (orsCD.BOF And orsCD.EOF) Then
                                orsCD.Filter = "FNSdtSeqNo = " & nI & " AND FNDctNo >= 8 AND FNDctNo <> 15 AND FNDctNo <> 21 AND FNDctNo <> 22 "
                                If Not (orsCD.BOF And orsCD.EOF) Then
                                    Do While Not orsCD.EOF
                                        cAddYellow = cAddYellow + CDbl(Val(Format(SP_FEDvChkNumeric(orsCD, "FCScdAmt"))))
                                        orsCD.MoveNext
                                    Loop
                                End If
                                orsCD.Filter = adFilterNone
                            End If
                        End If
                        If cAddYellow <> 0 Then
                            Call C_DATxUpdProDT(nI, "FCSdtYellowPrice + " & SP_FMTcCurPoint(cAddYellow, True, False), "FCSdtBluePrice", 0, 0)
                        Else
                            '*KT 52-12-09 กรณีที่สินค้าเป็น Member Price + Consignment และเป็น Double Policy กับ Mall Card ต้องเป็นสินค้าป้ายน้ำเงิน
                             If nC_ENDiscType = nEN_Disc21MPrice Or nC_ENDiscType = nEN_Disc22MDisc Then
                                Call C_DATxUpdProDT(nI, "FCSdtYellowPrice ", "FCSdtBluePrice", 0, 0)
                             End If
                        End If
                    Else
                        Call C_DATxUpdProDT(nI, "FCSdtYellowPrice", "FCSdtBluePrice", 0, 0)
                    End If
                End If
                orsDT.MoveNext
            Loop
    '-----------------------------------------------------------
    
            tSql = "SELECT MAX(FCSdtCostEX) AS FCSdtCostEX ,SUM(FCSdtBluePrice) AS FCSdtSumBluePrice"
            tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "  'tpsdt08107tmpontop
            tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
            Call SP_SQLbExecute(tSql, orsDT)
            If Not (orsDT.BOF Or orsDT.EOF) Then
                     cCostEx = SP_FEDvChkNumeric(orsDT, "FCSdtCostEX")
                     cBluePrice = SP_FEDvChkNumeric(orsDT, "FCSdtSumBluePrice")
            Else
                    cCostEx = 0
                    cBluePrice = 0
            End If
            '*KT 51-09-12
            'แก้ไขเรื่องการชำระด้วย Special Gift Voucher / Special Coupon กรณีที่ไม่มีส่วนลดใดเลย
            'โดยชำระ Special Gift Voucher / Special Coupon ส่วนหนึ่ง แล้วนำก้อนที่จ่าย  Special Gift Voucher / Special Coupon
            'ไปลดสินค้าป้ายเหลืองก่อนแล้วส่วนที่เหลือไปลดแบบ Pro Rate ป้ายน้ำเงิน
            tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) " 'tpsdt08107tmpontop
            tSql = tSql & vbCrLf & "SET FTSkuStaMallEmpCard = 'N'"
            tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
            If pbSpecial And Not (pbRetReg) Then '*TW 51-09-12
                If cBluePrice <> 0 Then
                    tSql = tSql & vbCrLf & "AND (FCSdtBluePrice - (FCSdtBluePrice * " & cCostEx / cBluePrice & ")) <= 0"
                Else
                    tSql = tSql & vbCrLf & "AND (FCSdtBluePrice) <= 0"
               End If
            Else
                tSql = tSql & vbCrLf & "AND (FCSdtBluePrice) <= 0"
            End If
            Call SP_SQLbExecute(tSql)
        End If
    End If
    Call SP_DATxRsNothing(orsDT)
    Call SP_DATxRsNothing(orsCD)
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsDT)
    Call SP_DATxRsNothing(orsCD)
    Call SP_DATxRsNothing(orsTemp)
End Sub
        

Private Sub C_GETxAllowMallCard(ByVal pnSeqNo As Integer, ByRef pcRegPrice As Double, ByRef pcProPrice As Double, ByRef pcAddReg As Double, ByRef pbSpecial As Boolean)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #  คำนวณหายอด Yellow Price และ Blue Price แล้วส่งกลับในรูป Argument
'ARGUMENT LIST # pnSeqNo = Sequence No of DT
'                                         pcRegPrice = Blue Price (ราคาปกติ)
'                                         pcProPrice = Yellow Price (ราคาโปรโมชั่น)
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim orsDDP As ADODB.Recordset
    Dim orsTemp2 As ADODB.Recordset
    Dim cProPrice As Double
    Dim cB4 As Double
    Dim cDisc As Double
    Dim cAF As Double
    Dim nDctNo As Long
    Dim tDiscType As String
    
    On Error GoTo Err_Handler
    
    tSql = "SELECT FTSkuStaMallEmpCard,FCSdtSaleAmt,FCSdtAmtDisAvi"
    tSql = tSql & vbCrLf & "FROM  " & tVB_CNTblDT & " DT WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & "AND DT.FNSdtSeqNo = " & pnSeqNo
    
    If SP_SQLbExecute(tSql, orsTemp, , , , , , , True) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            pcRegPrice = 0
            
            pcProPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt")
            cProPrice = pcProPrice
            
            If UCase(SP_FEDtChkString(orsTemp, "FTSkuStaMallEmpCard")) = "Y" Then
                pcRegPrice = pcProPrice
                pcProPrice = 0
                cProPrice = pcProPrice
                tSql = "SELECT FCScdAmtB4Avi,FCScdAmtAfAvi,FCScdAmt,FNDctNo,FNScdSeqNo,FTScdDiscType"
                tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD & " WITH(NOLOCK)  "
                tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill & " AND FNSdtSeqNo = " & pnSeqNo
                tSql = tSql & vbCrLf & "AND  ISNULL(FTSdtStkCode,'') <> '1'" '*KT 52-11-25 Deal&Promo
                tSql = tSql & vbCrLf & "ORDER BY FNScdSeqNo ASC" '*KT 54-12-02 V.4.0.8 ML-PH4-R3-015 Wrong Sub-Total Reg when paid with Tender not allow Mcard  and CoopCpn , CoopBnd
                If SP_SQLbExecute(tSql, orsTemp2, , , , , , , True) Then
                    If Not (orsTemp2.BOF And orsTemp2.EOF) Then
                        Call SP_RSTbClone2OfflineEx2(orsTemp2, orsTemp) '*KT 54-03-25 ป้องกัน recordset เน่าพร้อม connection ที่ reconnect ใหม่
                        orsTemp.Sort = "FNScdSeqNo ASC" '*KT 54-12-02 V.4.0.8 ML-PH4-R3-015 Wrong Sub-Total Reg when paid with Tender not allow Mcard  and CoopCpn , CoopBnd
                        orsTemp.MoveFirst '*KT 54-12-02 V.4.0.8 ML-PH4-R3-015 Wrong Sub-Total Reg when paid with Tender not allow Mcard  and CoopCpn , CoopBnd
                        Do While Not orsTemp.EOF
                            cB4 = SP_FEDvChkNumeric(orsTemp, "FCScdAmtB4Avi")
                            cAF = SP_FEDvChkNumeric(orsTemp, "FCScdAmtAfAvi")
                            cDisc = SP_FEDvChkNumeric(orsTemp, "FCScdAmt")
                            nDctNo = SP_FEDvChkNumeric(orsTemp, "FNDctNo")
                            tDiscType = SP_FEDvChkNumeric(orsTemp, "FTScdDiscType")
                            If cDisc > 0 Then
                                    tSql = "SELECT FTDdpStaPolicy,FNDdpNo1,FNDdpNo2"
                                    tSql = tSql & vbCrLf & "FROM TCNMDoubleDiscPolicy WITH(NOLOCK) "
                                    tSql = tSql & vbCrLf & "WHERE FNDdpNo1 = " & SP_FEDvChkNumeric(orsTemp, "FNDctNo")
                                    tSql = tSql & vbCrLf & "    AND FTDdpStaPolicy = 'N'"
                                    '*KT 52-02-02  PH2-ML-SF-014 แก้ไขใน POSFront 1.5.6
                                    '*KT 52-02-11 PH2-ML-SF-090โ ทำให้ Sub all ผิด เมื่อมีส่วนลดแบบ Employee discount
                                    '-------------------------------------------------------------------------------------------------------------
'                                    If SP_FEDvChkNumeric(orsTemp, "FNDctNo") = nEN_Disc12EmpCrd Then
'                                        tSql = tSql & vbCrLf & "    AND FNDdpNo2 = " & nEN_Disc12EmpCrd
'                                    Else
                                    tSql = tSql & vbCrLf & "    AND FNDdpNo2 = " & nEN_Disc11MallCrd
'                                    End If
                                    '--------------------------------------------------------------------------------------------------------------
                                    If SP_SQLbExecute(tSql, orsDDP, , , , oDB.ocnOnLine, True, True) Then 'read master
                                        If Not (orsDDP.BOF And orsDDP.EOF) Then
                                                '*KT 51-11-14 Communication Sheet ML-SF-645 TranDis show sub reg wrong
                                                '-------------------------------------------------------------------------------------------------------------------
''                                                Select Case SP_FEDvChkNumeric(orsTemp, "FNDctNo")
''                                                    Case nEN_Disc09CoopCpn, nEN_Disc10CoopBrand
''                                                        'none
''                                                        tSql = "SELECT FCScdAmtB4Avi,FCScdAmtAfAvi,FCScdAmt,FNDctNo"
''                                                        tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD
''                                                        tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill & " AND FNSdtSeqNo = " & pnSeqNo
''                                                        tSql = tSql & vbCrLf & " AND FNDctNo > " & nEN_Disc11MallCrd
''                                                        Call SP_SQLvExecute(tSql, orsTemp2)
''                                                        If (orsTemp2.BOF And orsTemp2.EOF) Then
''                                                            cProPrice = SP_FMTcCurPoint(SP_VALcText2Double(cProPrice + (cB4 - cAF - cDisc)), True, False)
''                                                        Else
''                                                             cProPrice = SP_FMTcCurPoint(SP_VALcText2Double(cProPrice + (cB4 - cAF)), True, False)
''                                                        End If
''                                                    Case Else
'                                                        cProPrice = SP_FMTcCurPoint(SP_VALcText2Double(cProPrice + (cB4 - cAF - cDisc)), True, False)
''                                                End Select
                                                '-------------------------------------------------------------------------------------------------------------------
                                                '*Ao 2011/06/28  CommSheet ML-PRV-062-SubTotal Reg Incorrect.xls
                                                If nDctNo = nEN_Disc12EmpCrd Then
                                                    cProPrice = SP_FMTcCurPoint(SP_VALcText2Double(cProPrice + (cB4 - cDisc)), True, False)
                                                Else
                                                    cProPrice = SP_FMTcCurPoint(SP_VALcText2Double(cProPrice + (cB4 - cAF - cDisc)), True, False)
                                                End If
                                                '__________________________________________________________________________________________________
                                        Else
'                                            'If SP_FEDvChkNumeric(orsTemp, "FNDctNo") > nEN_Disc12EmpCrd Then '*TW 51-08-28 เปลี่ยนจาก nEN_Disc11MallCrd เป็น nEN_Disc12EmpCrd
'                                            If nDctNo > nEN_Disc12EmpCrd Then '*TW 51-08-28 เปลี่ยนจาก nEN_Disc11MallCrd เป็น nEN_Disc12EmpCrd
'                                                If pbSpecial Then
'                                                    'Do Nothing
'                                                Else
''                                                    cProPrice = SP_FMTcCurPoint(SP_VALcText2Double(cProPrice + (cB4 - cAF - cDisc)), True, False)  '*KT 51-11-14 Communication Sheet ML-SF-645 TranDis show sub reg wrong
'
'                                                End If
'                                            End If
                                            '*Ao 2011/06/28  CommSheet ML-PRV-062-SubTotal Reg Incorrect.xls
                                            If nDctNo = nEN_Disc12EmpCrd Then
                                                cProPrice = SP_FMTcCurPoint(SP_VALcText2Double(cProPrice + (cB4 - cAF - cDisc)), True, False)
                                                pcAddReg = cDisc
                                                pcAddReg = IIf(pcAddReg < 0, 0, pcAddReg)
                                            End If
                                            '____________________________________________________________
                                        End If
                                        cProPrice = IIf(cProPrice < 0, 0, cProPrice)
                                    End If
                            End If
                            orsTemp.MoveNext
                        Loop
                        If pbSpecial Then
                            If pcAddReg >= 0 Then
                                pcProPrice = cProPrice
                                pcRegPrice = pcRegPrice - pcProPrice
                            Else
                                pcProPrice = IIf(cProPrice + pcAddReg <= 0, 0, cProPrice + pcAddReg)
                                pcRegPrice = pcRegPrice - pcProPrice
                            End If
                        Else
                            pcProPrice = cProPrice
                            pcRegPrice = pcRegPrice - pcProPrice
                        End If
                    Else
                        '*TW 51-09-12 เพิ่ม ว่าถ้าเป็น Special ไม่ต้องหัก AddReg ออก
                        If pbSpecial Then
                            'Do Nothing
                        Else
                            If pcAddReg > 0 Then
                                pcProPrice = SP_FMTcCurPoint(SP_VALcText2Double(pcProPrice + IIf(pcRegPrice - pcAddReg >= 0, pcAddReg, pcRegPrice)), True, False)
                                pcRegPrice = SP_FMTcCurPoint(SP_VALcText2Double(IIf(pcRegPrice - pcAddReg > 0, pcRegPrice - pcAddReg, 0)), True, False)
                            End If
                        End If
                    End If
                End If
            Else
                If pbSpecial Then
                    If pcAddReg >= 0 Then
                        pcProPrice = cProPrice
                        pcRegPrice = 0
                    Else
                        pcProPrice = SP_FMTcCurPoint(SP_VALcText2Double(IIf(cProPrice + pcAddReg <= 0, 0, cProPrice + pcAddReg)), True, False)
                        pcRegPrice = 0
                    End If
                End If
            End If
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    'Call SP_DATxRsNothing(orsTemp2)
    Call SP_DATxRsNothing(orsDDP)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
    'Call SP_DATxRsNothing(orsTemp2)
    Call SP_DATxRsNothing(orsDDP)
End Sub

Public Sub DISxGetSubRegPro(ByRef pcSubReg As Double, ByRef pcSubPro As Double)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # Get Sub Reg and Sub Pro (Screen) From Database
'ARGUMENT LIST # pcSubReg = Sub Reg (Return)
'                                         pcSubPro = Sub Pro (Return)
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsTemp  As ADODB.Recordset
    Dim tSql As String
    On Error GoTo Err_Handler
    'หา Sub Pro ทั้งบิล ขณะนั้น
    'tSql = "SELECT SUM(FCSdtBSubPro) AS FCSdtBSubPro, SUM(FCSdtBSubReg) AS FCSdtBSubReg" '*Tao 62-01-23 ของเดิม
    tSql = "SELECT ROUND(SUM(FCSdtBSubPro),2) AS FCSdtBSubPro, ROUND(SUM(FCSdtBSubReg),2) AS FCSdtBSubReg" '*Tao 62-01-23 ปัดเศษ
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
    If SP_SQLbExecute(tSql, orsTemp) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            pcSubReg = SP_FEDvChkNumeric(orsTemp, "FCSdtBSubReg")
            pcSubPro = SP_FEDvChkNumeric(orsTemp, "FCSdtBSubPro")
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Sub

Public Sub DISxGetSubRegProForBill(ByRef pcSubReg As Double, ByRef pcSubPro As Double)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # Get Sub Reg and Sub Pro (Bill) From Database
'ARGUMENT LIST # pcSubReg = Blue Price (Return)
'                                         pcSubPro = Yellow Price (Return)
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsTemp  As ADODB.Recordset
    Dim tSql As String
    Dim cDiscPlus
    Dim cCostEx As Double
On Error GoTo Err_Handler
    'หา Sub Pro ทั้งบิล ขณะนั้น
    pcSubReg = 0
    pcSubPro = 0
    cVB_CNAddSubTotal = 0
    tSql = "SELECT MAX(ISNULL(FCSdtCostEX,0)) AS FCSdtCostEx"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
    If SP_SQLbExecute(tSql, orsTemp) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            cCostEx = SP_FEDvChkNumeric(orsTemp, "FCSdtCostEx")
        End If
    End If
    
    tSql = "SELECT SUM(FCSdtYellowPrice) AS FCSdtYellowPrice,SUM(FCSdtBluePrice) AS FCSdtBluePrice,SUM(FCSdtSaleAmt) AS FCSdtSaleAmt"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
    If SP_SQLbExecute(tSql, orsTemp) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            pcSubReg = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSdtBluePrice"), True, False) '*KT 51-10-28 ใส่ Format เพราะบางครั้งเป็นค่า x.xxxxxxxxx e
            pcSubPro = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt"), True, False) '*KT 51-10-28 ใส่ Format เพราะบางครั้งเป็นค่า x.xxxxxxxxx e
            pcSubReg = IIf(pcSubReg - cCostEx > 0, pcSubReg - cCostEx, 0)
            If SP_FMTcCurPoint(pcSubPro - pcSubReg, True, False) > 0 Then '*KT 51-10-28 ใส่ Format เพราะบางครั้งเป็นค่า x.xxxxxxxxx e
                cVB_CNAddSubTotal = 0
                pcSubPro = pcSubPro - pcSubReg
            Else
                cVB_CNAddSubTotal = pcSubReg - pcSubPro
                pcSubPro = 0
            End If
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Sub

Private Sub C_DATxUpdProDT(ByVal pnSeqNo As Integer, ByVal ptValSubPro As String, ByVal ptValSubReg As String, ByVal ptValDiscPro As String, ByVal ptValDiscReg As String)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # Update Sub Reg Sub Pro
'ARGUMENT LIST # pnSeqNo = Sequence of DT
'                                         ptValSubPro = Value of FCSdtBSubPro
'                                         ptValSubReg = Value of FCSdtBSubReg
'                                         ptValDiscPro = Value of FCSdtDiscSubPro
'                                         ptValDiscReg = Value of FCSdtDiscSubReg
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    On Error GoTo Err_Handler
    tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) "
    tSql = tSql & vbCrLf & "SET FCSdtBSubPro = " & ptValSubPro
    tSql = tSql & vbCrLf & ",FCSdtDiscSubPro = " & ptValDiscPro
    tSql = tSql & vbCrLf & ",FCSdtBSubReg = " & ptValSubReg
    tSql = tSql & vbCrLf & ",FCSdtDiscSubReg = " & ptValDiscReg
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & "AND FNSdtSeqNo = " & pnSeqNo
    Call SP_SQLbExecute(tSql)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function C_GETcDiscCoopCpn(ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # Find Cooperate Coupon And Cooperate Brand Discount
'ARGUMENT LIST # porsTemp = Recordset of DT
'
'RETURN VALUE  # Discount value
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim bLimited As Boolean
    Dim cLimitAmt As Double
    Dim cYellowPrice As Double
    Dim cBluePrice As Double
    Dim cDisSumAmt As Double
    Dim tFilter As String
    Dim orsTemp2 As ADODB.Recordset
    Dim cFCSdtSaleAmt As Double
    Dim cSourceAmt As Double
    Dim orsTarget   As ADODB.Recordset
    Dim nI As Long
    Dim nSeqNo As Long
On Error GoTo Err_Handler
    nC_ENLnDisType = nEN_LnDis1Per
    cC_DiscPerOrg = cC_DiscValue
    cC_DiscAmtOrg = 0
    cC_DiscNewOrg = 0

    cC_DiscPer = cC_DiscValue
    cC_DiscAmt = 0
    cC_DiscNew = 0
    cDisSumAmt = 0
    If C_CHKbCoopFromMaster(cC_DiscPerOrg, bLimited, cLimitAmt) Then
        If nC_ENDiscType = nEN_Disc09CoopCpn Then
                porsTemp.Filter = "FTSkuStaCoCpnAlw = 'Y' AND FDSkuCoCpnEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
        Else
                porsTemp.Filter = "FTSkuStaCoBndAlw = 'Y' AND FDSkuCoBndEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
        End If
        
        Set orsTarget = New ADODB.Recordset
        For nI = 1 To porsTemp.Fields.Count
            Call orsTarget.Fields.Append(porsTemp.Fields(nI - 1).Name, porsTemp.Fields(nI - 1).Type, porsTemp.Fields(nI - 1).DefinedSize)
        Next nI
        orsTarget.CursorType = adOpenStatic
        orsTarget.Open
        
        Do While Not porsTemp.EOF
            nSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
            cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
            cSourceAmt = cFCSdtSaleAmt
            If cSourceAmt > 0 Then
                If C_CHKnDDP(nSeqNo) = 1 Then
                    cYellowPrice = cYellowPrice + SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice") 'SP_FMTcCurPoint(C_CPNcSumByRec(porsTemp, "FCSdtYellowPrice"), True, False) '*KT 51-10-28
                    cBluePrice = cBluePrice + SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice") 'SP_FMTcCurPoint(C_CPNcSumByRec(porsTemp, "FCSdtBluePrice"), True, False) '*KT 51-10-28
                    orsTarget.AddNew
                    For nI = 1 To porsTemp.Fields.Count
                        Select Case UCase(Mid(porsTemp.Fields(nI - 1).Name, 2, 1))
                            Case "T"
                                orsTarget.Fields(nI - 1).Value = porsTemp.Fields(nI - 1).Value & vbNullString
                            Case "C"
                                orsTarget.Fields(nI - 1).Value = SP_VALcText2Double(porsTemp.Fields(nI - 1).Value & vbNullString)
                            Case "N"
                                orsTarget.Fields(nI - 1).Value = SP_VALnText2Double(porsTemp.Fields(nI - 1).Value & vbNullString)
                            Case "D"
                                If IsNull(porsTemp.Fields(nI - 1).Value) Then
                                    orsTarget.Fields(nI - 1).Value = "Null"
                                Else
                                   orsTarget.Fields(nI - 1).Value = porsTemp.Fields(nI - 1).Value
                                End If
                        End Select
                    Next nI
                End If
            End If
            porsTemp.MoveNext
        Loop
        
        If bLimited Then 'เช็ค Limited
            If (cYellowPrice + cBluePrice) > cLimitAmt Then
                If cYellowPrice >= cLimitAmt Then
                    If cYellowPrice > 0 Then
                        If Not (orsTarget.BOF And orsTarget.EOF) Then orsTarget.MoveFirst '*KT 52-04-08 กรณีมีการส่ง recordset เข้าไปทำงานที่ Procedure อื่นให้ทำการ move record ไปที่ต้นทางเสมอเพราะว่าไม่รุ้ว่าก่อนหน้ามีการใช้งานอะไรมาบ้าง
                        Set orsTemp2 = orsTarget.Clone
                        cDisSumAmt = cDisSumAmt + C_CPNcDiscYellowLim(orsTemp2, cYellowPrice, cC_DiscPerOrg, cLimitAmt) 'ทำ Yellow Limit
                    End If
                Else
                    If cYellowPrice > 0 Then
                        If Not (orsTarget.BOF And orsTarget.EOF) Then orsTarget.MoveFirst '*KT 52-04-08 กรณีมีการส่ง recordset เข้าไปทำงานที่ Procedure อื่นให้ทำการ move record ไปที่ต้นทางเสมอเพราะว่าไม่รุ้ว่าก่อนหน้ามีการใช้งานอะไรมาบ้าง
                        Set orsTemp2 = orsTarget.Clone
                        cDisSumAmt = cDisSumAmt + C_CPNcDiscYellowUnl(orsTemp2, cYellowPrice, cC_DiscPerOrg, cLimitAmt) 'ทำ Yellow Unlimit
                        cLimitAmt = cLimitAmt - cYellowPrice
                    End If
                    If cBluePrice >= cLimitAmt Then
                        If Not (orsTarget.BOF And orsTarget.EOF) Then orsTarget.MoveFirst '*KT 52-04-08 กรณีมีการส่ง recordset เข้าไปทำงานที่ Procedure อื่นให้ทำการ move record ไปที่ต้นทางเสมอเพราะว่าไม่รุ้ว่าก่อนหน้ามีการใช้งานอะไรมาบ้าง
                        Set orsTemp2 = orsTarget.Clone
                        cDisSumAmt = cDisSumAmt + C_CPNcDiscBlueLim(orsTemp2, cBluePrice, cC_DiscPerOrg, cLimitAmt) 'ทำ Blue Limit
                    Else
                        If cBluePrice > 0 Then
                            If Not (orsTarget.BOF And orsTarget.EOF) Then orsTarget.MoveFirst '*KT 52-04-08 กรณีมีการส่ง recordset เข้าไปทำงานที่ Procedure อื่นให้ทำการ move record ไปที่ต้นทางเสมอเพราะว่าไม่รุ้ว่าก่อนหน้ามีการใช้งานอะไรมาบ้าง
                            Set orsTemp2 = orsTarget.Clone
                            cDisSumAmt = cDisSumAmt + C_CPNcDiscBlueUnl(orsTemp2, cBluePrice, cC_DiscPerOrg, cLimitAmt) 'ทำ Blue Unlimit
                        End If
                    End If
                End If
            Else
                If CDbl(SP_FMTcCurPoint(cBluePrice + cYellowPrice, True, False)) > 0 Then '*KT 51-10-28
                    If Not (orsTarget.BOF And orsTarget.EOF) Then orsTarget.MoveFirst '*KT 52-04-08 กรณีมีการส่ง recordset เข้าไปทำงานที่ Procedure อื่นให้ทำการ move record ไปที่ต้นทางเสมอเพราะว่าไม่รุ้ว่าก่อนหน้ามีการใช้งานอะไรมาบ้าง
                    Set orsTemp2 = orsTarget.Clone
                    cDisSumAmt = cDisSumAmt + C_CPNcDiscTotalLim(orsTemp2, cBluePrice + cYellowPrice, cC_DiscPerOrg, cLimitAmt) 'ทำ Total Limit
                End If
            End If
        Else
            If SP_FMTcCurPoint(cBluePrice + cYellowPrice, True, False) > 0 Then
                If Not (orsTarget.BOF And orsTarget.EOF) Then orsTarget.MoveFirst '*KT 52-04-08 กรณีมีการส่ง recordset เข้าไปทำงานที่ Procedure อื่นให้ทำการ move record ไปที่ต้นทางเสมอเพราะว่าไม่รุ้ว่าก่อนหน้ามีการใช้งานอะไรมาบ้าง
                Set orsTemp2 = orsTarget.Clone
                cDisSumAmt = cDisSumAmt + C_CPNcDiscTotalUnl(orsTemp2, cBluePrice + cYellowPrice, cC_DiscPerOrg, cLimitAmt) 'ทำ Total Unlimit
            End If
        End If
    End If
    C_GETcDiscCoopCpn = cDisSumAmt
    Call SP_DATxRsNothing(orsTemp2)
    Call SP_DATxRsNothing(orsTarget)
    Exit Function
Err_Handler:
    C_GETcDiscCoopCpn = 0
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_GETcDiscCoopCpn)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp2)
    Call SP_DATxRsNothing(orsTarget)
End Function

Private Function C_CPNcSumByRec(ByVal porsTemp As ADODB.Recordset, ByVal ptRefField As String) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # Return sum Price Refer ptRefField that Can allow by DDP
'ARGUMENT LIST # porsTemp = Recordset of DT
'                                         ptRefField = Amount Filed Name
'
'RETURN VALUE  # Sum Price
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSum As Double
    Dim nSeqNo As Integer
On Error GoTo Err_Handler
    With porsTemp
        If (.BOF And .EOF) Then Exit Function
        .MoveFirst
        Do While Not .EOF
            nSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
            If C_CHKnDDP(nSeqNo) = 1 Then
                cSum = cSum + SP_FEDvChkNumeric(porsTemp, ptRefField)
            End If
            .MoveNext
        Loop
        .MoveFirst '*KT 52-04-08 ให้ทำการ move recordset กลับมาที่ต้นทาง
    End With
    C_CPNcSumByRec = cSum
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_CPNcProrate(ByVal porsTemp As ADODB.Recordset, ByVal ptRefField As String, ByVal pcSumAll As Double, ByVal pcLimitAmt As Double, ByVal ptEventTemp As String, ByVal pcPer As Double, ByVal ptTotal As String) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # Calculate and Prepare data to Save on CD
'ARGUMENT LIST # porsTemp = Recordset of DT
'                                         ptRefField = Amount Filed Name
'                                         pcSumAll = ยอดคำนวณทั้งบิล
'                                         pcLimitAmt = ยอด Limit ที่สามารถคำนวณได้ ทั้งบิล
'                                         ptEventTemp = Event ID ของ Discount นี้
'                                         pcPer = Discount Percentage
'                                         ptTotal = ชื่อ Field ที่รวมกับ ptRefField แล้วเป็น Total เพราะการคำนวณแบบ Total ต้องใช้ ทั้ง Yellow Price และ Blue Price ในการคำนวณ
'
'RETURN VALUE  # Sum Discount
'*KT 53-02-09 Communication Sheet PH3-ML-SF-083 Record Corperate Coupon rate wrong when SKU allow lower % than Corporate Coupon % input.xls
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตั้งต้น 3.0.19
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cCal As Double
    Dim nFNSdtSeqNo As Integer, nRecordCount As Integer
    Dim tEventID As String
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cAlwSource As Double
    Dim tSKUCode As String
    Dim cSumDiscount As Double
    Dim cDisTotal As Double                     '*Ao 2011/02/15  CommSheet ML-PRV-038-Wrong rounding for Cooperate Coupon & Cooperate Brand.xls
    On Error GoTo Err_Handler
    With porsTemp
        If (.BOF And .EOF) Then Exit Function
        cDisTotal = pcLimitAmt * pcPer / 100            '*Ao 2011/02/15  CommSheet ML-PRV-038-Wrong rounding for Cooperate Coupon & Cooperate Brand.xls
        .MoveFirst
        Do While Not .EOF 'ดำเนินการ
            nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
            nRecordCount = nRecordCount + 1
            If C_CHKnDDP(nFNSdtSeqNo) = 1 Then
                If ptTotal = "" Then
                    cCal = SP_FEDvChkNumeric(porsTemp, ptRefField)
                Else
                    cCal = SP_FEDvChkNumeric(porsTemp, ptRefField)
                    cCal = cCal + SP_FEDvChkNumeric(porsTemp, ptTotal) 'Total Price
                End If
                tSKUCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
                tEventID = SP_SQLtFunction("", ptEventTemp, tVB_CNTblSku, "WHERE FTSkuCode = '" & tSKUCode & "'", oDB.ocnOnLine, True, True) 'Read Master
                cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
                cAlwSource = cCal / pcSumAll * pcLimitAmt
                cC_DiscAmt = cAlwSource * pcPer / 100
                cC_DiscPer = pcPer '*KT 53-02-09 Communication Sheet PH3-ML-SF-083 Record Corperate Coupon rate wrong when SKU allow lower % than Corporate Coupon % input.xls
                '*Ao 2011/02/15  CommSheet ML-PRV-038-Wrong rounding for Cooperate Coupon & Cooperate Brand.xls
                If .RecordCount = nRecordCount Then
                    If cSumDiscount + SP_FMTcCurPoint(cC_DiscAmt, True, False) > cDisTotal Then
                        cAlwSource = cAlwSource + ((cSumDiscount + SP_FMTcCurPoint(cC_DiscAmt, True, False)) - cDisTotal)
                        cC_DiscAmt = cC_DiscAmt - ((cSumDiscount + SP_FMTcCurPoint(cC_DiscAmt, True, False)) - cDisTotal)
                    Else
                        If cSumDiscount + SP_FMTcCurPoint(cC_DiscAmt, True, False) < cDisTotal Then
                            cAlwSource = cAlwSource - (cDisTotal - (cSumDiscount + SP_FMTcCurPoint(cC_DiscAmt, True, False)))
                            cC_DiscAmt = cC_DiscAmt + (cDisTotal - (cSumDiscount + SP_FMTcCurPoint(cC_DiscAmt, True, False)))
                        End If
                    End If
                End If
                '____________________________________________________________________________________________
                Call C_DATxInsertDiscCD(cC_DiscAmt, nFNSdtSeqNo, tEventID, cFCSdtQtyAllDisAvi, cAlwSource, 0)
'                cSumDiscount = cSumDiscount +cC_DiscAmt
                cSumDiscount = cSumDiscount + SP_FMTcCurPoint(cC_DiscAmt, True, False)              '*Ao 2011/02/15  CommSheet ML-PRV-038-Wrong rounding for Cooperate Coupon & Cooperate Brand.xls
            End If
            .MoveNext
        Loop
       .MoveFirst  '*KT 52-04-08 ให้ทำการ move recordset กลับมาที่ต้นทาง
        C_CPNcProrate = cSumDiscount
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_CPNcFull(ByVal porsTemp As ADODB.Recordset, ByVal ptRefField As String, ByVal ptEventTemp As String, ByVal pcPer As Double, ByVal ptTemp As String, ByVal ptTotal As String) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # Calculate and Prepare data to Save on CD (Full)
'ARGUMENT LIST # porsTemp = Recordset of DT
'                                         ptRefField = Amount Filed Name
'                                         pcSumAll = ยอดคำนวณทั้งบิล
'                                         pcLimitAmt = ยอด Limit ที่สามารถคำนวณได้ ทั้งบิล
'                                         ptEventTemp = Event ID ของ Discount นี้
'                                         pcPer = Discount Percentage
'                                         ptTemp = ชื่อ Field Rate จาก SKU
'                                         ptTotal = ชื่อ Field ที่รวมกับ ptRefField แล้วเป็น Total เพราะการคำนวณแบบ Total ต้องใช้ ทั้ง Yellow Price และ Blue Price ในการคำนวณ
'
'RETURN VALUE  #
'*KT 53-02-09 Communication Sheet PH3-ML-SF-083 Record Corperate Coupon rate wrong when SKU allow lower % than Corporate Coupon % input.xls
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตั้งต้น 3.0.19
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cCal As Double
    Dim nFNSdtSeqNo As Integer
    Dim tEventID As String
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cAlwSource As Double
    Dim tSKUCode As String
    Dim cSumDiscount As Double
    Dim cRateSKU As Double
    Dim cRateCal As Double
    On Error GoTo Err_Handler
    With porsTemp
        If (.BOF And .EOF) Then Exit Function
        .MoveFirst
        Do While Not .EOF 'ดำเนินการ
            nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
            If C_CHKnDDP(nFNSdtSeqNo) = 1 Then
                If ptTotal = "" Then
                    cCal = SP_FEDvChkNumeric(porsTemp, ptRefField)
                Else
                    cCal = SP_FEDvChkNumeric(porsTemp, ptRefField)
                    cCal = cCal + SP_FEDvChkNumeric(porsTemp, ptTotal)
                End If
                
                tSKUCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
                tEventID = SP_SQLtFunction("", ptEventTemp, tVB_CNTblSku, "WHERE FTSkuCode = '" & tSKUCode & "'", oDB.ocnOnLine, True, True) 'read master
                cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
                cRateSKU = SP_FEDvChkNumeric(porsTemp, ptTemp)
                If cRateSKU > pcPer Then
                    cRateCal = pcPer
                Else
                    cRateCal = cRateSKU
                End If
                cAlwSource = cCal
                cC_DiscAmt = cAlwSource * cRateCal / 100
                cC_DiscPer = cRateCal '*KT 53-02-09 Communication Sheet PH3-ML-SF-083 Record Corperate Coupon rate wrong when SKU allow lower % than Corporate Coupon % input.xls
                Call C_DATxInsertDiscCD(cC_DiscAmt, nFNSdtSeqNo, tEventID, cFCSdtQtyAllDisAvi, cAlwSource, 0)
                cSumDiscount = cSumDiscount + cC_DiscAmt
            End If
            .MoveNext
        Loop
        .MoveFirst '*KT 52-04-08 set record  กลับมาที่ต้นทางเสมอ
        C_CPNcFull = cSumDiscount
    End With
    Exit Function
Err_Handler:
     Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_CPNcFull)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_CPNcDiscYellowLim(ByVal porsTemp As ADODB.Recordset, ByVal pcSumYellow As Double, ByVal pcCoopPer As Double, ByRef pcLimitAmt As Double) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # หา Discount Cooperate Coupon or Brand เฉพาะส่วนป้ายเหลือง แบบมี Limit
'ARGUMENT LIST # porsTemp = Recordset of DT
'                                         pcSumYellow = ยอดรวมราคาป้ายเหลืองที่สามารถคำนวณได้
'                                         pcCoopPer = Percent of Cooperate
'                                 Ref pcLimitAmt = Current Limit (After calculate)
'RETURN VALUE  # ส่วนลดที่ได้
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tTemp As String
    Dim tEvnTemp As String
    Dim acKeepRate() As Double
    Dim cOldRate As Double
    Dim nI As Integer
    Dim cSumCalGTEQ As Double
    Dim cSumCalLEQByRate As Double
    Dim cSumDiscount As Double
    Dim tFilter As String
    Dim cB4DisChg As Double
On Error GoTo Err_Handler
    If nC_ENDiscType = nEN_Disc09CoopCpn Then
        tTemp = "FCSkuCoCpnRate"
        tEvnTemp = "FTSkuEventCoDisCpn"
        tFilter = "FTSkuStaCoCpnAlw = 'Y' AND FDSkuCoCpnEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    Else
        tTemp = "FCSkuCoBndRate"
        tEvnTemp = "FTSkuEventCoBnd"
        tFilter = "FTSkuStaCoBndAlw = 'Y' AND FDSkuCoBndEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    End If

    With porsTemp
        If Not (.BOF And .EOF) Then
            .Filter = "FCSdtYellowPrice > 0 AND " & tTemp & " >= " & pcCoopPer & " AND " & tFilter
            cSumCalGTEQ = C_CPNcSumByRec(porsTemp, "FCSdtYellowPrice")
            If cSumCalGTEQ > pcLimitAmt Then            'Prorate
                cSumDiscount = cSumDiscount + C_CPNcProrate(porsTemp, "FCSdtYellowPrice", cSumCalGTEQ, pcLimitAmt, tEvnTemp, pcCoopPer, "")
                pcLimitAmt = 0
            Else            'DiscFull (cSumCal)
                cSumDiscount = cSumDiscount + C_CPNcFull(porsTemp, "FCSdtYellowPrice", tEvnTemp, pcCoopPer, tTemp, "")
                pcLimitAmt = pcLimitAmt - cSumCalGTEQ 'ราคาที่ CPN อนุญาตให้ลดได้
                cSumCalGTEQ = pcSumYellow - cSumCalGTEQ 'ส่วนที่เหลือที่จะลดต่อ ของป้ายเหลือ
                If pcLimitAmt > 0 And cSumCalGTEQ > 0 Then
                    .Filter = "FCSdtYellowPrice > 0 AND " & tTemp & " < " & pcCoopPer & " AND " & tFilter
                    .Sort = tTemp & " DESC"
                    If Not (.BOF And .EOF) Then '*KT 52-03-23 ตรวจสอบก่อนมีการ Move First
                        ReDim acKeepRate(0) As Double
                        .MoveFirst
                        cOldRate = 99999
                        Do While Not .EOF ' เก็บ Rate ในแต่ละระดับไว้
                            If cOldRate > SP_FEDvChkNumeric(porsTemp, tTemp) Then
                                ReDim Preserve acKeepRate(UBound(acKeepRate) + 1) As Double
                                acKeepRate(UBound(acKeepRate)) = SP_FEDvChkNumeric(porsTemp, tTemp)
                            End If
                            .MoveNext
                        Loop
                        
                        For nI = 1 To UBound(acKeepRate) 'หาทีละระดับ  เอาอันมากหาก่อน
                            .Filter = "FCSdtYellowPrice > 0 AND " & tTemp & " = " & acKeepRate(nI) & " AND " & tFilter
                            cSumCalLEQByRate = C_CPNcSumByRec(porsTemp, "FCSdtYellowPrice")
                            If cSumCalLEQByRate > pcLimitAmt Then 'แบบ Prorate
                                cSumDiscount = cSumDiscount + C_CPNcProrate(porsTemp, "FCSdtYellowPrice", cSumCalLEQByRate, pcLimitAmt, tEvnTemp, acKeepRate(nI), "")
                                pcLimitAmt = 0
                            Else 'แบบหาทั้งหมด
                                cSumDiscount = cSumDiscount + C_CPNcFull(porsTemp, "FCSdtYellowPrice", tEvnTemp, acKeepRate(nI), tTemp, "")
                                pcLimitAmt = pcLimitAmt - cSumCalLEQByRate
                            End If
                            If pcLimitAmt <= 0 Then Exit For
                        Next nI
                    End If
                End If
            End If
        End If
    End With
    C_CPNcDiscYellowLim = cSumDiscount
    Exit Function
Err_Handler:
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_CPNcDiscYellowLim)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_CPNcDiscYellowUnl(ByVal porsTemp As ADODB.Recordset, ByVal pcSumYellow As Double, ByVal pcCoopPer As Double, ByRef pcLimitAmt As Double) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # หา Discount Cooperate Coupon or Brand เฉพาะส่วนป้ายเหลือง แบบไม่มี Limit
'ARGUMENT LIST # porsTemp = Recordset of DT
'                                         pcSumYellow = ยอดรวมราคาป้ายเหลืองที่สามารถคำนวณได้
'                                         pcCoopPer = Percent of Cooperate
'                                 Ref pcLimitAmt = Current Limit (After calculate)
'RETURN VALUE  # ส่วนลดที่ได้
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tTemp As String
    Dim tEvnTemp As String
    Dim acKeepRate() As Double
    Dim cOldRate As Double
    Dim nI As Integer
    Dim cSumCalGTEQ As Double
    Dim cSumCalLEQByRate As Double
    Dim cSumDiscount As Double
    Dim tFilter As String
    Dim cB4DisChg As Double
On Error GoTo ErrHandle
    If nC_ENDiscType = nEN_Disc09CoopCpn Then
        tTemp = "FCSkuCoCpnRate"
        tEvnTemp = "FTSkuEventCoDisCpn"
        tFilter = "FTSkuStaCoCpnAlw = 'Y' AND FDSkuCoCpnEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    Else
        tTemp = "FCSkuCoBndRate"
        tEvnTemp = "FTSkuEventCoBnd"
        tFilter = "FTSkuStaCoBndAlw = 'Y' AND FDSkuCoBndEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    End If
    
    With porsTemp
        If Not (.BOF Or .EOF) Then '*KT 52-04-08 ห้ามใช้ AND เพราะมีการ Filter จากข้างนอกก่อนหน้ามีโอกาศที่จะเป็น  BOF = true  EOF = false
            .Filter = "FCSdtYellowPrice > 0" & " AND " & tFilter
            If Not (.BOF Or .EOF) Then '*KT 52-03-23 ตรวจสอบ BOF กับ EOF  '*KT 52-04-08 ห้ามใช้ AND เพราะมีการ Filter จากข้างนอกก่อนหน้ามีโอกาศที่จะเป็น  BOF = true  EOF = false
                cSumDiscount = cSumDiscount + C_CPNcFull(porsTemp, "FCSdtYellowPrice", tEvnTemp, pcCoopPer, tTemp, "")
            End If
        End If
    End With
    C_CPNcDiscYellowUnl = cSumDiscount
    Exit Function
ErrHandle:
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_CPNcDiscYellowUnl)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_CPNcDiscBlueLim(ByVal porsTemp As ADODB.Recordset, ByVal pcSumBlue As Double, ByVal pcCoopPer As Double, ByRef pcLimitAmt As Double) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # หา Discount Cooperate Coupon or Brand เฉพาะส่วนป้ายน้ำเงิน แบบมี Limit
'ARGUMENT LIST # porsTemp = Recordset of DT
'                                         pcSumBlue = ยอดรวมราคาป้ายน้ำเงินที่สามารถคำนวณได้
'                                         pcCoopPer = Percent of Cooperate
'                                 Ref pcLimitAmt = Current Limit (After calculate)
'RETURN VALUE  # ส่วนลดที่ได้
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tTemp As String
    Dim tEvnTemp As String
    Dim acKeepRate() As Double
    Dim cOldRate As Double
    Dim nI As Integer
    Dim cSumCalGTEQ As Double
    Dim cSumCalLEQByRate As Double
    Dim cSumDiscount As Double
    Dim tFilter As String
    Dim cB4DisChg As Double
On Error GoTo ErrHandle
    If nC_ENDiscType = nEN_Disc09CoopCpn Then
        tTemp = "FCSkuCoCpnRate"
        tEvnTemp = "FTSkuEventCoDisCpn"
        tFilter = "FTSkuStaCoCpnAlw = 'Y' AND FDSkuCoCpnEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    Else
        tTemp = "FCSkuCoBndRate"
        tEvnTemp = "FTSkuEventCoBnd"
        tFilter = "FTSkuStaCoBndAlw = 'Y' AND FDSkuCoBndEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    End If
    
    With porsTemp
        If Not (.BOF And .EOF) Then
            .Filter = "FCSdtBluePrice > 0 AND " & tTemp & " >= " & pcCoopPer & " AND " & tFilter
            cSumCalGTEQ = C_CPNcSumByRec(porsTemp, "FCSdtBluePrice")
            If cSumCalGTEQ > pcLimitAmt Then            'Prorate
                cSumDiscount = cSumDiscount + SP_FMTcCurPoint(C_CPNcProrate(porsTemp, "FCSdtBluePrice", cSumCalGTEQ, pcLimitAmt, tEvnTemp, pcCoopPer, ""), True, False) '*KT 51-10-28
                pcLimitAmt = 0
            Else            'DiscFull (cSumCal)
                cSumDiscount = cSumDiscount + SP_FMTcCurPoint(C_CPNcFull(porsTemp, "FCSdtBluePrice", tEvnTemp, pcCoopPer, tTemp, ""), True, False)  '*KT 51-10-28
                pcLimitAmt = pcLimitAmt - cSumCalGTEQ 'ราคาที่ CPN อนุญาตให้ลดได้
                cSumCalGTEQ = pcSumBlue - cSumCalGTEQ 'ส่วนที่เหลือที่จะลดต่อ ของป้ายเหลือ
                If pcLimitAmt > 0 And cSumCalGTEQ > 0 Then
                    .Filter = "FCSdtBluePrice > 0 AND " & tTemp & " < " & pcCoopPer & " AND " & tFilter
                    .Sort = tTemp & " DESC"
                    If Not (.BOF And .EOF) Then '*KT 52-03-23 ตรวจสอบก่อนมีการ Move First
                        ReDim acKeepRate(0) As Double
                        .MoveFirst
                        cOldRate = 99999
                        Do While Not .EOF ' เก็บ Rate ในแต่ละระดับไว้
                            If cOldRate > CDbl(SP_FMTcCurPoint(SP_FEDvChkNumeric(porsTemp, tTemp), True, False)) Then   '*KT 51-10-28
                                ReDim Preserve acKeepRate(UBound(acKeepRate) + 1) As Double
                                acKeepRate(UBound(acKeepRate)) = SP_FMTcCurPoint(SP_FEDvChkNumeric(porsTemp, tTemp), True, False)  '*KT 51-10-28
                            End If
                            .MoveNext
                        Loop
                        
                        For nI = 1 To UBound(acKeepRate) 'หาทีละระดับ  เอาอันมากหาก่อน
                            .Filter = "FCSdtBluePrice > 0 AND " & tTemp & " = " & acKeepRate(nI) & " AND " & tFilter
                            cSumCalLEQByRate = SP_FMTcCurPoint(C_CPNcSumByRec(porsTemp, "FCSdtBluePrice"), True, False)  '*KT 51-10-28
                            If cSumCalLEQByRate > pcLimitAmt Then 'แบบ Prorate
                                cSumDiscount = cSumDiscount + SP_FMTcCurPoint(C_CPNcProrate(porsTemp, "FCSdtBluePrice", cSumCalLEQByRate, pcLimitAmt, tEvnTemp, acKeepRate(nI), ""), True, False)  '*KT 51-10-28
                                pcLimitAmt = 0
                            Else 'แบบหาทั้งหมด
                                cSumDiscount = cSumDiscount + SP_FMTcCurPoint(C_CPNcFull(porsTemp, "FCSdtBluePrice", tEvnTemp, acKeepRate(nI), tTemp, ""), True, False)  '*KT 51-10-28
                                pcLimitAmt = pcLimitAmt - cSumCalLEQByRate
                            End If
                            If pcLimitAmt <= 0 Then Exit For
                        Next nI
                    End If
                End If
            End If
        End If
    End With
    C_CPNcDiscBlueLim = cSumDiscount
    Exit Function
ErrHandle:
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_CPNcDiscBlueLim)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_CPNcDiscBlueUnl(ByVal porsTemp As ADODB.Recordset, ByVal pcSumBlue As Double, ByVal pcCoopPer As Double, ByRef pcLimitAmt As Double) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # หา Discount Cooperate Coupon or Brand เฉพาะส่วนป้ายน้ำเงิน แบบไม่มี Limit
'ARGUMENT LIST # porsTemp = Recordset of DT
'                                         pcSumBlue = ยอดรวมราคาป้ายน้ำเงินที่สามารถคำนวณได้
'                                         pcCoopPer = Percent of Cooperate
'                                 Ref pcLimitAmt = Current Limit (After calculate)
'RETURN VALUE  # ส่วนลดที่ได้
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tTemp As String
    Dim tEvnTemp As String
    Dim acKeepRate() As Double
    Dim cOldRate As Double
    Dim nI As Integer
    Dim cSumCalGTEQ As Double
    Dim cSumCalLEQByRate As Double
    Dim cSumDiscount As Double
    Dim tFilter As String
    Dim cB4DisChg As Double
On Error GoTo ErrHandle
    If nC_ENDiscType = nEN_Disc09CoopCpn Then
        tTemp = "FCSkuCoCpnRate"
        tEvnTemp = "FTSkuEventCoDisCpn"
        tFilter = "FTSkuStaCoCpnAlw = 'Y' AND FDSkuCoCpnEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    Else
        tTemp = "FCSkuCoBndRate"
        tEvnTemp = "FTSkuEventCoBnd"
        tFilter = "FTSkuStaCoBndAlw = 'Y' AND FDSkuCoBndEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    End If
    With porsTemp
        If Not (.BOF And .EOF) Then
            .Filter = "FCSdtBluePrice > 0" & " AND " & tFilter
             If Not (.BOF And .EOF) Then '*KT 52-03-23 ตรวจสอบ BOF กับ EOF
                cSumDiscount = cSumDiscount + C_CPNcFull(porsTemp, "FCSdtBluePrice", tEvnTemp, pcCoopPer, tTemp, "")
            End If
        End If
    End With
    C_CPNcDiscBlueUnl = cSumDiscount
    Exit Function
ErrHandle:
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_CPNcDiscBlueUnl)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_CPNcDiscTotalLim(ByVal porsTemp As ADODB.Recordset, ByVal pcSumTotal As Double, ByVal pcCoopPer As Double, ByRef pcLimitAmt As Double) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # หา Discount Cooperate Coupon or Brand แบบมี Limit
'ARGUMENT LIST # porsTemp = Recordset of DT
'                                         pcSumTotal = ยอดรวมที่สามารถคำนวณได้
'                                         pcCoopPer = Percent of Cooperate
'                                 Ref pcLimitAmt = Current Limit (After calculate)
'RETURN VALUE  # ส่วนลดที่ได้
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim tTemp As String
    Dim tEvnTemp As String
    Dim acKeepRate() As Double
    Dim cOldRate As Double
    Dim nI As Integer
    Dim cSumCalGTEQ As Double
    Dim cSumCalLEQByRate As Double
    Dim cSumDiscount As Double
    Dim tFilter As String
    Dim cB4DisChg As Double
On Error GoTo ErrHandle
    If nC_ENDiscType = nEN_Disc09CoopCpn Then
        tTemp = "FCSkuCoCpnRate"
        tEvnTemp = "FTSkuEventCoDisCpn"
        tFilter = "FTSkuStaCoCpnAlw = 'Y' AND FDSkuCoCpnEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    Else
        tTemp = "FCSkuCoBndRate"
        tEvnTemp = "FTSkuEventCoBnd"
        tFilter = "FTSkuStaCoBndAlw = 'Y' AND FDSkuCoBndEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    End If
    
    With porsTemp
        If Not (.BOF And .EOF) Then
            .Filter = tTemp & " >= " & pcCoopPer & " AND " & tFilter
            cSumCalGTEQ = SP_FMTcCurPoint(C_CPNcSumByRec(porsTemp, "FCSdtBluePrice"), True, False) '*KT 51-10-28
            cSumCalGTEQ = cSumCalGTEQ + SP_FMTcCurPoint(C_CPNcSumByRec(porsTemp, "FCSdtYellowPrice"), True, False) '*KT 51-10-28
            If cSumCalGTEQ > pcLimitAmt Then            'Prorate
                cSumDiscount = cSumDiscount + SP_FMTcCurPoint(C_CPNcProrate(porsTemp, "FCSdtBluePrice", cSumCalGTEQ, pcLimitAmt, tEvnTemp, pcCoopPer, "FCSdtYellowPrice"), True, False) '*KT 51-10-28
                pcLimitAmt = 0
            Else            'DiscFull (cSumCal)
                cSumDiscount = cSumDiscount + SP_FMTcCurPoint(C_CPNcFull(porsTemp, "FCSdtBluePrice", tEvnTemp, pcCoopPer, tTemp, "FCSdtYellowPrice"), True, False) '*KT 51-10-28
                pcLimitAmt = pcLimitAmt - cSumCalGTEQ 'ราคาที่ CPN อนุญาตให้ลดได้
                cSumCalGTEQ = pcSumTotal - cSumCalGTEQ 'ส่วนที่เหลือที่จะลดต่อ ของป้ายเหลือ
                If pcLimitAmt > 0 And cSumCalGTEQ > 0 Then
                    .Filter = tTemp & " < " & pcCoopPer & " AND " & tFilter
                    .Sort = tTemp & " DESC"
                     If Not (.BOF And .EOF) Then '*KT 52-03-23 ตรวจสอบ BOF กับ EOF
                        ReDim acKeepRate(0) As Double
                        .MoveFirst
                        cOldRate = 99999
                        Do While Not .EOF ' เก็บ Rate ในแต่ละระดับไว้
                            If cOldRate > SP_FEDvChkNumeric(porsTemp, tTemp) Then
                                ReDim Preserve acKeepRate(UBound(acKeepRate) + 1) As Double
                                acKeepRate(UBound(acKeepRate)) = SP_FMTcCurPoint(SP_FEDvChkNumeric(porsTemp, tTemp), True, False) '*KT 51-10-28
                            End If
                            .MoveNext
                        Loop
                        
                        For nI = 1 To UBound(acKeepRate) 'หาทีละระดับ  เอาอันมากหาก่อน
                            .Filter = tTemp & " = " & acKeepRate(nI) & " AND " & tFilter
                            cSumCalLEQByRate = SP_FMTcCurPoint(C_CPNcSumByRec(porsTemp, "FCSdtBluePrice"), True, False) '*KT 51-10-28
                            If cSumCalLEQByRate > pcLimitAmt Then 'แบบ Prorate
                                cSumDiscount = cSumDiscount + SP_FMTcCurPoint(C_CPNcProrate(porsTemp, "FCSdtBluePrice", cSumCalLEQByRate, pcLimitAmt, tEvnTemp, acKeepRate(nI), "FCSdtYellowPrice"), True, False) '*KT 51-10-28
                                pcLimitAmt = 0
                            Else 'แบบหาทั้งหมด
                                cSumDiscount = cSumDiscount + SP_FMTcCurPoint(C_CPNcFull(porsTemp, "FCSdtBluePrice", tEvnTemp, acKeepRate(nI), tTemp, "FCSdtYellowPrice"), True, False) '*KT 51-10-28
                                pcLimitAmt = pcLimitAmt - cSumCalLEQByRate
                            End If
                            If pcLimitAmt <= 0 Then Exit For
                        Next nI
                    End If
                End If
            End If
        End If
    End With
    C_CPNcDiscTotalLim = cSumDiscount
    Exit Function
ErrHandle:
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_CPNcDiscTotalLim)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_CPNcDiscTotalUnl(ByVal porsTemp As ADODB.Recordset, ByVal pcSumTotal As Double, ByVal pcCoopPer As Double, ByRef pcLimitAmt As Double) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # หา Discount Cooperate Coupon or Brand แบบไม่มี Limit
'ARGUMENT LIST # porsTemp = Recordset of DT
'                                         pcSumTotal = ยอดรวมที่สามารถคำนวณได้
'                                         pcCoopPer = Percent of Cooperate
'                                 Ref pcLimitAmt = Current Limit (After calculate)
'RETURN VALUE  # ส่วนลดที่ได้
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tTemp As String
    Dim tEvnTemp As String
    Dim cSumDiscount As Double
    Dim tFilter As String
    Dim cB4DisChg As Double
On Error GoTo ErrHandle
    If nC_ENDiscType = nEN_Disc09CoopCpn Then
        tTemp = "FCSkuCoCpnRate"
        tEvnTemp = "FTSkuEventCoDisCpn"
        tFilter = "FTSkuStaCoCpnAlw = 'Y' AND FDSkuCoCpnEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    Else
        tTemp = "FCSkuCoBndRate"
        tEvnTemp = "FTSkuEventCoBnd"
        tFilter = "FTSkuStaCoBndAlw = 'Y' AND FDSkuCoBndEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "'" '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    End If
    With porsTemp
        If Not (.BOF And .EOF) Then
            .Filter = "(FCSdtBluePrice > 0 AND " & tFilter & ") OR (FCSdtYellowPrice > 0 AND " & tFilter & ")"
            If Not (.BOF And .EOF) Then '*KT 52-03-23 ตรวจสอบ BOF กับ EOF
                cSumDiscount = cSumDiscount + SP_FMTcCurPoint(C_CPNcFull(porsTemp, "FCSdtBluePrice", tEvnTemp, pcCoopPer, tTemp, "FCSdtYellowPrice"), True, False) '*KT 51-10-28
            End If
        End If
    End With
    C_CPNcDiscTotalUnl = cSumDiscount
    Exit Function
ErrHandle:
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_CPNcDiscTotalUnl)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_EMPbGetDptRate(ByVal ptTblName As String, ByVal ptFieldName As String, ByVal ptDptCode As String, ByVal ptBrdCode As String, ByVal ptCardKey As String, ByRef pcRate As Double) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #  คืนค่า True ถ้า มีการกำหนด Depart   False ถ้า ไม่มีการกำหนด Depart
'ARGUMENT LIST # ptTblName = ชื่อตาราง
'                                         ptFieldName = ชื่อ Field Code
'                                         ptDptCode = Sku Depart Code
'                                         ptCardKey = Code ของ Mall Card หรือ Employee Card
'                                 Ref pcRate = Percent of ..
'RETURN VALUE  # คืนค่า True ถ้า มีการกำหนด Depart   False ถ้า ไม่มีการกำหนด Depart
'*KT 52-02-18
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tWhe As String
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    On Error GoTo Err_Handler
    tWhe = "WHERE FTMpcCode = '" & ptDptCode & "'"
    If UCase(ptTblName) = UCase("TCNMEmpCardDep") Then
        tWhe = tWhe & " AND FTEcdCode = '" & ptCardKey & "'"
    Else
        tWhe = tWhe & " AND FTMcdCode = '" & ptCardKey & "'"
        tWhe = tWhe & " AND '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "' BETWEEN FDMpcStartDate AND FDMpcEndDate" '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    End If
    
    tSql = "SELECT  " & ptFieldName
    tSql = tSql & vbCrLf & "FROM " & ptTblName
    tSql = tSql & vbCrLf & tWhe
    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then 'read master
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            pcRate = SP_FEDvChkNumeric(orsTemp, ptFieldName)
            C_EMPbGetDptRate = True
        Else
            pcRate = 0
            C_EMPbGetDptRate = False
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function C_EMPbGetRateDpt(ByVal porsMallDptAlw As ADODB.Recordset, ByVal pcSaleAmt As Double, ByRef pcRate As Double, ByRef pcMallAmt As Double) As Boolean
'-----------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 2013/03/12  TK
'-----------------------------------------------------------------------
Dim bChk As Boolean
Dim tDptAllow As String
Dim cMin As Double, cMax As Double
Dim dSTDate As Date, dEDDate As Date
    
On Error GoTo Err_Handler
    bChk = False
    
    If Not (porsMallDptAlw.BOF Or porsMallDptAlw.EOF) Then
        tDptAllow = SP_FEDtChkString(porsMallDptAlw, "FTMpcAlw")
        cMin = SP_FEDtChkString(porsMallDptAlw, "FCMpcAmtMin")
        cMax = SP_FEDtChkString(porsMallDptAlw, "FCMpcAmtMax")
        dSTDate = SP_FEDdChkDate(porsMallDptAlw, "FDMpcStartDate")
        dEDDate = SP_FEDdChkDate(porsMallDptAlw, "FDMpcEndDate")
        If (SP_DTEtStrSTD(dSTDate) <= SP_DTEtStrSTD(SP_GETdChkDateDisc) And SP_DTEtStrSTD(dEDDate) >= SP_DTEtStrSTD(SP_GETdChkDateDisc)) Then
            Select Case tDptAllow
                Case "0"
                    If pcSaleAmt >= cMin Then
                        If pcSaleAmt > cMax Then
                            pcSaleAmt = cMax
                        End If
                        pcRate = SP_VALcText2Double(SP_FEDvChkNumeric(porsMallDptAlw, "FCMpcRate"))
                        pcMallAmt = pcSaleAmt
                    Else
                        pcRate = 0
                        pcMallAmt = 0
                    End If
                Case "1"
                    pcRate = 0
                    pcMallAmt = 0
                Case Else
                    pcRate = 0
                    pcMallAmt = 0
            End Select
        Else
            pcRate = 0
            pcMallAmt = 0
        End If
        bChk = True
    Else
        bChk = False
    End If
    C_EMPbGetRateDpt = bChk
Exit Function
Err_Handler:
    C_EMPbGetRateDpt = False
End Function

Private Function C_EMPbGetDptRateEx(ByVal porsMallDptAlw As ADODB.Recordset, ByVal ptSKUBlue, ByVal pcBluePrice As Double, ByVal pcYellowPrice As Double, ByVal pcRegSumAmt As Double, ByVal pcProSumAmt As Double, ByVal pcSumQty As Double, ByVal ocSumArray As Double, ByRef pcRate As Double, ByRef pcMallAmt As Double) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #  คืนค่า True ถ้า มีการกำหนด Depart   False ถ้า ไม่มีการกำหนด Depart
'ARGUMENT LIST # ptTblName = ชื่อตาราง
'                                         ptFieldName = ชื่อ Field Code
'                                         ptDptCode = Sku Depart Code
'                                         ptCardKey = Code ของ Mall Card หรือ Employee Card
'                                 Ref pcRate = Percent of ..
'RETURN VALUE  # คืนค่า True ถ้า มีการกำหนด Depart   False ถ้า ไม่มีการกำหนด Depart
'*KT 52-02-18
'Ret : ptBlueYellow = 0: Blue 1: Yellow 2: All
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    Dim tReg As String
    Dim tPro As String
    Dim tDptAllow As String
    Dim cMin As Double
    Dim cMax  As Double
    Dim cMaxAvg As Double
    Dim dSTDate As Date
    Dim dEDDate As Date

On Error GoTo Err_Handler
    C_EMPbGetDptRateEx = False
    
    If Not (porsMallDptAlw.BOF Or porsMallDptAlw.EOF) Then
            tDptAllow = SP_FEDtChkString(porsMallDptAlw, "FTMpcAlw")
            pcRate = SP_VALcText2Double(SP_FEDvChkNumeric(porsMallDptAlw, "FCMpcRate"))
            tReg = SP_FEDtChkString(porsMallDptAlw, "FTMpcReg")
            tPro = SP_FEDtChkString(porsMallDptAlw, "FTMpcPro")
            cMin = SP_FEDtChkString(porsMallDptAlw, "FCMpcAmtMin")
            cMax = SP_FEDtChkString(porsMallDptAlw, "FCMpcAmtMax")
            dSTDate = SP_FEDdChkDate(porsMallDptAlw, "FDMpcStartDate")
            dEDDate = SP_FEDdChkDate(porsMallDptAlw, "FDMpcEndDate")
            If (SP_DTEtStrSTD(dSTDate) <= SP_DTEtStrSTD(SP_GETdChkDateDisc) And SP_DTEtStrSTD(dEDDate) >= SP_DTEtStrSTD(SP_GETdChkDateDisc)) Then  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุของ point ตามที่กำหนดใน parameter
                 Select Case tDptAllow
                     Case "0" 'allow
                             If pcSumQty = 0 Then pcSumQty = 1
'                             cMaxAvg = CDbl(SP_FMTcCurPoint(cMax / pcSumQty, True, False))         '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                             If tReg = "Y" And tPro = "Y" Then
                                 If ((pcRegSumAmt + pcProSumAmt) >= cMin And (pcRegSumAmt + pcProSumAmt) <= cMax) Or (pcRegSumAmt + pcProSumAmt) > cMax Then
                                     If (pcRegSumAmt + pcProSumAmt) > cMax Then
'                                         pcMallAmt = cMaxAvg               '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                                        If ocSumArray >= cMax Then
                                            pcMallAmt = 0
                                            pcRate = 0
                                        ElseIf ocSumArray > 0 Then
                                            pcMallAmt = cMax - ocSumArray
                                            If pcMallAmt > pcBluePrice + pcYellowPrice Then
                                                pcMallAmt = pcBluePrice + pcYellowPrice
                                            End If
                                        Else
                                            If pcBluePrice + pcYellowPrice < cMax Then
                                                pcMallAmt = pcBluePrice + pcYellowPrice
                                            Else
                                                pcMallAmt = cMax
                                            End If
                                        End If
                                     Else
                                         pcMallAmt = pcBluePrice + pcYellowPrice
                                     End If
                                     C_EMPbGetDptRateEx = True
                                  Else
                                     pcRate = 0
                                     pcMallAmt = 0
                                     C_EMPbGetDptRateEx = False
                                  End If
                             ElseIf tReg = "Y" And tPro = "N" Then
                                 If ((pcRegSumAmt) >= cMin And (pcRegSumAmt) <= cMax) Or (pcRegSumAmt) > cMax Then
                                     If (pcRegSumAmt) > cMax Then
'                                         pcMallAmt = cMaxAvg               '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                                        If ocSumArray >= cMax Then
                                            pcMallAmt = 0
                                            pcRate = 0
                                        ElseIf ocSumArray > 0 Then
                                            pcMallAmt = cMax - ocSumArray
                                            If pcMallAmt > pcBluePrice Then
                                                pcMallAmt = pcBluePrice
                                            End If
                                        Else
                                            If pcBluePrice < cMax Then
                                                pcMallAmt = pcBluePrice
                                            Else
                                                pcMallAmt = cMax
                                            End If
                                        End If
                                     Else
                                         pcMallAmt = pcBluePrice
                                     End If
                                     C_EMPbGetDptRateEx = True
                                 Else
                                     pcRate = 0
                                     pcMallAmt = 0
                                     C_EMPbGetDptRateEx = False
                                 End If
                             ElseIf tReg = "N" And tPro = "Y" Then
                                  If ((pcProSumAmt) >= cMin And (pcProSumAmt) <= cMax) Or (pcProSumAmt) > cMax Then
                                      If (pcProSumAmt) > cMax Then
'                                         pcMallAmt = cMaxAvg                   '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                                        If ocSumArray >= cMax Then
                                            pcMallAmt = 0
                                            pcRate = 0
                                        ElseIf ocSumArray > 0 Then
                                            pcMallAmt = cMax - ocSumArray
                                            If pcMallAmt > pcYellowPrice Then
                                                pcMallAmt = pcYellowPrice
                                            End If
                                        Else
                                            If pcYellowPrice < cMax Then
                                                pcMallAmt = pcYellowPrice
                                            Else
                                                pcMallAmt = cMax
                                            End If
                                        End If
                                     Else
                                         pcMallAmt = pcYellowPrice
                                     End If
                                     C_EMPbGetDptRateEx = True
                                 Else
                                     pcRate = 0
                                     pcMallAmt = 0
                                     C_EMPbGetDptRateEx = False
                                 End If
                             Else
                                 pcRate = 0
                                 pcMallAmt = 0
                                 C_EMPbGetDptRateEx = False
                             End If
                     Case "1" 'not allow
                             tReg = SP_FEDtChkString(porsMallDptAlw, "FTMpcReg")
                             tPro = SP_FEDtChkString(porsMallDptAlw, "FTMpcPro")
                             If tReg = "Y" And tPro = "Y" Then
                                 If ((pcRegSumAmt + pcProSumAmt) >= cMin And (pcRegSumAmt + pcProSumAmt) <= cMax) Or (pcRegSumAmt + pcProSumAmt) > cMax Then
                                     pcRate = 0
                                     pcMallAmt = 0
                                     C_EMPbGetDptRateEx = True
                                 Else
                                     pcRate = 0
                                     pcMallAmt = 0
                                     C_EMPbGetDptRateEx = False
                                 End If
                             ElseIf tReg = "Y" And tPro = "N" Then
                                 If ((pcRegSumAmt) >= cMin And (pcRegSumAmt) <= cMax) Or (pcRegSumAmt) > cMax Then
                                     pcRate = pcRate
                                     pcMallAmt = pcYellowPrice
                                     C_EMPbGetDptRateEx = True
                                 Else
                                     pcRate = 0
                                     pcMallAmt = 0
                                     C_EMPbGetDptRateEx = False
                                 End If
                             ElseIf tReg = "N" And tPro = "Y" Then
                                 If ((pcProSumAmt) >= cMin And (pcProSumAmt) <= cMax) Or (pcProSumAmt) > cMax Then
                                     pcRate = pcRate
                                     pcMallAmt = pcBluePrice
                                     C_EMPbGetDptRateEx = True
                                 Else
                                     pcRate = 0
                                     pcMallAmt = 0
                                     C_EMPbGetDptRateEx = False
                                 End If
                             Else
                                 pcRate = 0
                                 pcMallAmt = 0
                                 C_EMPbGetDptRateEx = False
                             End If
                     Case Else
                             pcRate = 0
                             pcMallAmt = 0
                             C_EMPbGetDptRateEx = False
                End Select
        Else
                pcRate = 0
                pcMallAmt = 0
                C_EMPbGetDptRateEx = False
        End If
    Else
            pcRate = 0
            pcMallAmt = 0
            C_EMPbGetDptRateEx = False
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    C_EMPbGetDptRateEx = False
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Function C_EMPbGetSKURateEx(ByVal porsMallSKUAlw As ADODB.Recordset, ByVal pcBluePrice As Double, ByVal pcYellowPrice As Double, ByVal pcSumAmt As Double, ByVal pcSumQty As Double, ByRef pcRate As Double, ByRef pcMallAmt As Double) As Boolean
''CREATE BY            # kitti                              LAST MODIFIED BY           # kitti
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
''DESCRIPTION       #  คืนค่า True ถ้า มีการกำหนด Depart   False ถ้า ไม่มีการกำหนด Depart
''ARGUMENT LIST # ptTblName = ชื่อตาราง
''                                         ptFieldName = ชื่อ Field Code
''                                         ptSKUCode = Sku Depart Code
''                                         ptCardKey = Code ของ Mall Card หรือ Employee Card
''                                 Ref pcRate = Percent of ..
''RETURN VALUE  #
''*KT 52-02-18
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim orsTemp As ADODB.Recordset
'    Dim tSql As String
'    Dim tSKUAllow As String
'   Dim cPriceAvg As Double
'    Dim cMin As Double
'    Dim cMax As Double
'    Dim dSTDate As Date
'    Dim dEDDate As Date
'
'On Error GoTo Err_Handler
'    C_EMPbGetSKURateEx = False
'
'    If Not (porsMallSKUAlw.BOF Or porsMallSKUAlw.EOF) Then
'        cMin = SP_FEDvChkNumeric(porsMallSKUAlw, "FCMcsAmtMin")
'        cMax = SP_FEDvChkNumeric(porsMallSKUAlw, "FCMcsAmtMax")
'        dSTDate = SP_FEDdChkDate(porsMallSKUAlw, "FDMcsStartDate")
'        dEDDate = SP_FEDdChkDate(porsMallSKUAlw, "FDMcsEndDate")
'        If (SP_DTEtStrSTD(dSTDate) <= SP_DTEtStrSTD(SP_GETdChkDateDisc) And SP_DTEtStrSTD(dEDDate) >= SP_DTEtStrSTD(SP_GETdChkDateDisc)) And ((cMin <= pcSumAmt And cMax >= pcSumAmt) Or cMax < pcSumAmt) Then  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุของ point ตามที่กำหนดใน parameter
'            tSKUAllow = SP_FEDtChkString(porsMallSKUAlw, "FTMcsAlw")
'            Select Case tSKUAllow
'                Case "0"
'                    pcRate = SP_VALcText2Double(SP_FEDvChkNumeric(porsMallSKUAlw, "FCMcsRate"))
''*KT 52-06-06 PH 2.0 ML-SF-016 คำนวนส่วนลด Mall Card ผิดกรณีที่ Allow SKU   1. กรณีไม่มี Limit และขายสินค้ามากกว่า 1 รายการ  2. กรณีมี Limit จะไม่ตรวจสอบ Limit
''                    If pcSumQty = 0 Then pcSumQty = 1
''                    cPriceAvg = CDbl(SP_FMTcCurPoint((pcSumAmt / pcSumQty), True, False))
''                    If CDbl(SP_FMTcCurPoint((pcBluePrice + pcYellowPrice), True, False)) > cPriceAvg Then
''                        pcMallAmt = cPriceAvg
''                    Else
''                        pcMallAmt = pcBluePrice + pcYellowPrice
''                    End If
'                    If pcSumAmt > cMax Then
'                        pcMallAmt = SP_FMTcCurPoint((pcBluePrice + pcYellowPrice) * cMax / pcSumAmt, True, False)
'                    Else
'                        pcMallAmt = pcBluePrice + pcYellowPrice
'                    End If
'                    C_EMPbGetSKURateEx = True
'                Case "1"
'                    pcRate = 0
'                    pcMallAmt = 0
'                    C_EMPbGetSKURateEx = True
'                Case Else
'                    pcRate = 0
'                    pcMallAmt = 0
'                    C_EMPbGetSKURateEx = False
'            End Select
'        Else
'            pcRate = 0
'            pcMallAmt = 0
'            C_EMPbGetSKURateEx = False
'        End If
'    Else
'        pcRate = 0
'        pcMallAmt = 0
'        C_EMPbGetSKURateEx = False
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'    Exit Function
'Err_Handler:
'    pcRate = 0
'    pcMallAmt = 0
'    C_EMPbGetSKURateEx = False
'    Call SP_DATxRsNothing(orsTemp)
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function
Private Function C_EMPbGetSKURate(ByVal porsMallSKUAlw As ADODB.Recordset, ByVal pcSaleAmt As Double, ByRef pcRate As Double, ByRef pcMallAmt As Double) As Boolean
'---------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 2013/03/12 TK
'---------------------------------------------------------------
Dim cMin As Double, cMax As Double
Dim dSTDate As Date, dEDDate As Date
Dim bChk As Boolean
Dim tSKUAllow  As String

On Error GoTo Err_Handler
    bChk = False
     If Not (porsMallSKUAlw.BOF Or porsMallSKUAlw.EOF) Then
        tSKUAllow = SP_FEDtChkString(porsMallSKUAlw, "FTMcsAlw")
        cMin = SP_FEDvChkNumeric(porsMallSKUAlw, "FCMcsAmtMin")
        cMax = SP_FEDvChkNumeric(porsMallSKUAlw, "FCMcsAmtMax")
        dSTDate = SP_FEDdChkDate(porsMallSKUAlw, "FDMcsStartDate")
        dEDDate = SP_FEDdChkDate(porsMallSKUAlw, "FDMcsEndDate")
        If (SP_DTEtStrSTD(dSTDate) <= SP_DTEtStrSTD(SP_GETdChkDateDisc) And SP_DTEtStrSTD(dEDDate) >= SP_DTEtStrSTD(SP_GETdChkDateDisc)) Then
            Select Case tSKUAllow
                Case "0"
                    If pcSaleAmt >= cMin Then
                        If pcSaleAmt > cMax Then
                            pcSaleAmt = cMax
                        End If
                        pcRate = SP_VALcText2Double(SP_FEDvChkNumeric(porsMallSKUAlw, "FCMcsRate"))
                        pcMallAmt = pcSaleAmt
                    Else
                        pcRate = 0
                        pcMallAmt = 0
                    End If
                Case "1"
                    pcRate = 0
                    pcMallAmt = 0
                Case Else
                    pcRate = 0
                    pcMallAmt = 0
            End Select
        Else
            pcRate = 0
            pcMallAmt = 0
        End If
        bChk = True
     Else
        bChk = False
     End If
     C_EMPbGetSKURate = bChk
Exit Function
Err_Handler:
    C_EMPbGetSKURate = False
End Function

Private Function C_EMPbGetSKURateEx(ByVal porsMallSKUAlw As ADODB.Recordset, ByVal pcBluePrice As Double, ByVal pcYellowPrice As Double, ByVal pcSumAmt As Double, ByVal pcSumQty As Double, ByRef pcRate As Double, ByRef pcMallAmt As Double) As Boolean
'CREATE BY            # kitti                              LAST MODIFIED BY           # kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #  คืนค่า True ถ้า มีการกำหนด Depart   False ถ้า ไม่มีการกำหนด Depart
'ARGUMENT LIST # ptTblName = ชื่อตาราง
'                                         ptFieldName = ชื่อ Field Code
'                                         ptSKUCode = Sku Depart Code
'                                         ptCardKey = Code ของ Mall Card หรือ Employee Card
'                                 Ref pcRate = Percent of ..
'RETURN VALUE  #
'*KT 52-02-18
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    Dim tSKUAllow As String
   Dim cPriceAvg As Double
    Dim cMin As Double
    Dim cMax As Double
    Dim dSTDate As Date
    Dim dEDDate As Date
    
On Error GoTo Err_Handler
    C_EMPbGetSKURateEx = False
    
    If Not (porsMallSKUAlw.BOF Or porsMallSKUAlw.EOF) Then
        cMin = SP_FEDvChkNumeric(porsMallSKUAlw, "FCMcsAmtMin")
        cMax = SP_FEDvChkNumeric(porsMallSKUAlw, "FCMcsAmtMax")
        dSTDate = SP_FEDdChkDate(porsMallSKUAlw, "FDMcsStartDate")
        dEDDate = SP_FEDdChkDate(porsMallSKUAlw, "FDMcsEndDate")
        If (SP_DTEtStrSTD(dSTDate) <= SP_DTEtStrSTD(SP_GETdChkDateDisc) And SP_DTEtStrSTD(dEDDate) >= SP_DTEtStrSTD(SP_GETdChkDateDisc)) And ((cMin <= pcSumAmt And cMax >= pcSumAmt) Or cMax < pcSumAmt) Then  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุของ point ตามที่กำหนดใน parameter
            tSKUAllow = SP_FEDtChkString(porsMallSKUAlw, "FTMcsAlw")
            Select Case tSKUAllow
                Case "0"
                    pcRate = SP_VALcText2Double(SP_FEDvChkNumeric(porsMallSKUAlw, "FCMcsRate"))
'*KT 52-06-06 PH 2.0 ML-SF-016 คำนวนส่วนลด Mall Card ผิดกรณีที่ Allow SKU   1. กรณีไม่มี Limit และขายสินค้ามากกว่า 1 รายการ  2. กรณีมี Limit จะไม่ตรวจสอบ Limit
'                    If pcSumQty = 0 Then pcSumQty = 1
'                    cPriceAvg = CDbl(SP_FMTcCurPoint((pcSumAmt / pcSumQty), True, False))
'                    If CDbl(SP_FMTcCurPoint((pcBluePrice + pcYellowPrice), True, False)) > cPriceAvg Then
'                        pcMallAmt = cPriceAvg
'                    Else
'                        pcMallAmt = pcBluePrice + pcYellowPrice
'                    End If
                    If pcSumAmt > cMax Then
                        pcMallAmt = SP_FMTcCurPoint((pcBluePrice + pcYellowPrice) * cMax / pcSumAmt, True, False)
                    Else
                        pcMallAmt = pcBluePrice + pcYellowPrice
                    End If
                    C_EMPbGetSKURateEx = True
                Case "1"
                    pcRate = 0
                    pcMallAmt = 0
                    C_EMPbGetSKURateEx = True
                Case Else
                    pcRate = 0
                    pcMallAmt = 0
                    C_EMPbGetSKURateEx = False
            End Select
        Else
            pcRate = 0
            pcMallAmt = 0
            C_EMPbGetSKURateEx = False
        End If
    Else
        pcRate = 0
        pcMallAmt = 0
        C_EMPbGetSKURateEx = False
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    pcRate = 0
    pcMallAmt = 0
    C_EMPbGetSKURateEx = False
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function C_EMPcBalDis(ByVal ptEmpD As String) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # หายอด Balance Limit ของ Employee
'ARGUMENT LIST # ptEmpD = หมายเลข Employee ID
'
'RETURN VALUE  # ยอด Balance Limit ของ Employee
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
        C_EMPcBalDis = 0
        tSql = "SELECT  Top 1 FCEcdBalDis"
        tSql = tSql & vbCrLf & "FROM TCNMEmpCard WITH(NOLOCK)  "
        tSql = tSql & vbCrLf & "WHERE FTEcdCode = '" & ptEmpD & "'"
        If SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True) = 0 Then
            'Error จะไม่เช็ค
            C_EMPcBalDis = SP_FEDvChkNumeric(orsTemp, "FCEcdBalDis")
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function C_EMPbChkAlwBuy(ByVal ptEmpD As String, ByRef pcLimit As Double, ByVal ptChkAllowBuy As String) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # ตรวจสอบเวลาซื้อ
'ARGUMENT LIST # ptEmpD = หมายเลข Employee ID
'                                 Ref pcLimit = ยอด Limit ที่ต้องเช็ค
'RETURN VALUE  # True ถ้าอนุญาต และ False ถ้าไม่อนุญาต
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tPosition As String
    Dim tNow As String
    Dim bOnline As Boolean
On Error GoTo Err_Handler
    bOnline = Not (ocnMember Is Nothing)
    If bOnline = False Then
        '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
        'If ocnAll.ConnectionString <> oDB.ocnOffLine.ConnectionString And bC_ReConnectMember Then 'PH 2.0.0 OLD
         If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) And bC_ReConnectMember Then
            If SP_CONbMemberServer(, , False) Then
                bOnline = True
            End If
        End If
    End If
    If bOnline Then
        tSql = "SELECT Top 1 FTEcdPosition ,FCEcdBalDis"
        tSql = tSql & vbCrLf & "FROM TCNMEmpCard WITH(NOLOCK) "
        tSql = tSql & vbCrLf & "WHERE FTEcdCode = '" & ptEmpD & "'"
        If SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True) <> 0 Then
            'Error จะไม่เช็ค
            C_EMPbChkAlwBuy = True
            Call SP_DATxRsNothing(orsTemp)
            Exit Function
        End If
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            tPosition = SP_FEDtChkString(orsTemp, "FTEcdPosition")
            pcLimit = SP_FEDvChkNumeric(orsTemp, "FCEcdBalDis")
            If UCase(ptChkAllowBuy) = "Y" Then              '*Ao 09-12-18   PH3-ML-SF-050
                tSql = "SELECT Top 1 FTEstSTime1 AS T1,FTEstETime1 AS E1"
                tSql = tSql & vbCrLf & ",FTEstSTime2 AS T2,FTEstETime2 AS E2"
                tSql = tSql & vbCrLf & ",FTEstSTime3 AS T3,FTEstETime3 AS E3"
                tSql = tSql & vbCrLf & ",FTEstAlwFlag"
                tSql = tSql & vbCrLf & "FROM TCNMEmpShopping WITH(NOLOCK)  "
                tSql = tSql & vbCrLf & "WHERE FDEstDate = '" & Format(SP_GETdChkDateDisc, "yyyy/MM/dd") & "' AND FTEcdPosition = '" & tPosition & "'" '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
                If SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True) <> 0 Then
                    'Error จะไม่เช็ค
                    C_EMPbChkAlwBuy = True
                    Call SP_DATxRsNothing(orsTemp)
                    Exit Function
                End If
                If Not (orsTemp.BOF And orsTemp.EOF) Then
                    If SP_FEDtChkString(orsTemp, "FTEstAlwFlag") = "Y" Then
'                        tNow = Format(Now, "HH:MM")
                        tNow = Format(Time$, "hh:mm")       '*Ao 09-12-21    PH3-ML-SF-050
                        tNow = Replace(tNow, ":", "")
                        If (SP_FEDtChkString(orsTemp, "T1") <= tNow And tNow <= SP_FEDtChkString(orsTemp, "E1")) _
                        Or (SP_FEDtChkString(orsTemp, "T2") <= tNow And tNow <= SP_FEDtChkString(orsTemp, "E2")) _
                        Or (SP_FEDtChkString(orsTemp, "T3") <= tNow And tNow <= SP_FEDtChkString(orsTemp, "E3")) Then
                            C_EMPbChkAlwBuy = False
                        Else
                            C_EMPbChkAlwBuy = True
                        End If
                    Else
                        C_EMPbChkAlwBuy = False
                    End If
                Else
                    C_EMPbChkAlwBuy = False
                End If
            End If
        Else
            'ฟ้อง
            C_EMPbChkAlwBuy = False
        End If
        
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Sub C_EMPxInsEmp(ByVal pcDisc As Double, ByVal pnSeqNo As Integer, ByVal ptEventID As String, ByVal pcQtyAvi As Double, ByVal pcAlwCal As Double, ByVal pcDiscPer As Double, ByRef pcDisSumAmt As Double, ByRef pcDisSumAmtEx As Double, Optional ByVal pbChk As Boolean = False)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # สั่ง Insert ส่วนลด Employee Card ลง CD
'ARGUMENT LIST # pcDisc = ส่วนลด
'                                         pnSeqNo = DT SeqNo
'                                         ptEventID = Event ID
'                                         pcQtyAvi = จำนวนที่อนุญาต
'                                         pcAlwCal = จำนวนที่คำนวณ
'                                         pcDiscPer = Percent Discount
'                                 Ref pcDisSumAmt = ยอดรวมส่วนลด
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tEventNo As String
    
    cC_DiscAmt = pcDisc
    tEventNo = ptEventID
    If cC_DiscAmt > 0 Then
        pcDisSumAmt = pcDisSumAmt + cC_DiscAmt
        '*Ao 2011/06/22  CommSheet ML-PRV-061 Employee Disc Over limit.xls
        pcDisSumAmtEx = pcDisSumAmtEx + SP_FMTcCurPoint(cC_DiscAmt, True, False)
        If pbChk = True Then
            If pcDisSumAmtEx > pcDisSumAmt Then
                cC_DiscAmt = cC_DiscAmt - (pcDisSumAmtEx - pcDisSumAmt)
            Else
                If pcDisSumAmtEx < pcDisSumAmt Then
                    cC_DiscAmt = cC_DiscAmt + (pcDisSumAmt - pcDisSumAmtEx)
                End If
            End If
        End If
        '______________________________________________________________________________________
        cC_DiscPer = pcDiscPer
        Call C_DATxInsertDiscCD(cC_DiscAmt, pnSeqNo, tEventNo, pcQtyAvi, pcAlwCal, 0)
    Else
        bW_ChkMallCard = True
        pcDisSumAmt = pcDisSumAmt + cC_DiscAmt
        cC_DiscPer = pcDiscPer
        Call C_DATxInsertDiscCD(cC_DiscAmt, pnSeqNo, tEventNo, pcQtyAvi, pcAlwCal, 0)
    End If
End Sub

Private Sub C_EMPxUpdMember(ByVal ptCardID As String, ByVal pcSumAFDisc As Double) 'ส่งส่วนที่ลดแล้วเข้าไป
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # เก็บ Query การ Update Employee ที่ Member ไว้ เพื่อจะ Run ตอนท้ายบิล
'ARGUMENT LIST # ptCardID = หมายเลข Employee ID
'                                         pcSumAFDisc = ยอดส่วนลดที่ได้
'RETURN VALUE  #
'*KT 53-01-28 Communication Sheet PH3-ML-SB-074 Remain Discount for Employee.xls
'*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133 ใส่ ROUND(x,2)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    tSql = "UPDATE  TCNMEmpCard WITH(ROWLOCK) "
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnIPV, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVRedeem, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVVIPService, tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax '
              tSql = tSql & vbCrLf & "SET FCEcdBalDis = ROUND(FCEcdBalDis + " & SP_FMTcCurPoint(pcSumAFDisc, True, False) & "," & nVB_CNDecAmtForSav & ")"
        Case Else
              tSql = tSql & vbCrLf & "SET FCEcdBalDis = ROUND(FCEcdBalDis - " & SP_FMTcCurPoint(pcSumAFDisc, True, False) & "," & nVB_CNDecAmtForSav & ")"
     End Select
     tSql = tSql & vbCrLf & "WHERE FTEcdCode = '" & ptCardID & "'"
     tVB_CNUpdateEmpLimit = tSql '*TW 51-08-18 เก็บไว้ Run ตอนท้ายบิล
     cVB_CNRemainDiscAmt = SP_FMTcCurPoint(pcSumAFDisc, True, False)
     tVB_CNRemainDiscID = ptCardID
End Sub

Private Sub C_EMPxUpdMemberEx() 'ส่งส่วนที่ลดแล้วเข้าไป
'CREATE BY            # Kitt                                                LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 53-01-28
'DESCRIPTION       # เก็บ Query การ Update Employee ที่ Member ไว้ เพื่อจะ Run ตอนท้ายบิล
'Variable                     # tCardID = หมายเลข Employee ID
'                                         cSumAFDisc = ยอดส่วนลดที่ได้
'RETURN VALUE  #
'*KT 53-01-28 Communication Sheet PH3-ML-SB-074 Remain Discount for Employee.xls
'*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133 ใส่ ROUND(x,2)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tCardID$
    Dim cSumAFDisc As Double
On Error GoTo Err_Handler
    tSql = "SELECT FTEcdCode,FCShdCreditStore"
    tSql = tSql & vbCrLf & " FROM " & tVB_CNTblHD & " WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
    If SP_SQLvExecute(tSql, orsTemp) = 0 Then
        tCardID = SP_FEDtChkString(orsTemp, "FTEcdCode")
        cSumAFDisc = SP_FEDvChkNumeric(orsTemp, "FCShdCreditStore")
        tSql = "UPDATE  TCNMEmpCard WITH(ROWLOCK) "
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnIPV, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVRedeem, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVVIPService, tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax '
                  tSql = tSql & vbCrLf & "SET FCEcdBalDis = ROUND(FCEcdBalDis + " & SP_FMTcCurPoint(cSumAFDisc, True, False) & "," & nVB_CNDecAmtForSav & ")"
            Case Else
                  tSql = tSql & vbCrLf & "SET FCEcdBalDis = ROUND(FCEcdBalDis - " & SP_FMTcCurPoint(cSumAFDisc, True, False) & "," & nVB_CNDecAmtForSav & ")"
         End Select
         tSql = tSql & vbCrLf & "WHERE FTEcdCode = '" & tCardID & "'"
         tVB_CNUpdateEmpLimit = tSql
    Else
        tVB_CNUpdateEmpLimit = ""
        cVB_CNRemainDiscAmt = 0
        tVB_CNRemainDiscID = ""
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
Err_Handler:
    Call SP_DATxRsNothing(orsTemp)
    tVB_CNUpdateEmpLimit = ""
    cVB_CNRemainDiscAmt = 0
    tVB_CNRemainDiscID = ""
End Sub

Private Sub C_EMPxUpdMemberEx2() 'ส่งส่วนที่ลดแล้วเข้าไป
'CREATE BY            # kitti                                               LAST MODIFIED BY           # kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 53-01-28
'DESCRIPTION       # เก็บ Query การ Update Employee ที่ Member ไว้ เพื่อจะ Run ตอนท้ายบิล
'Variable                    # tCardID = หมายเลข Employee ID
'                                         cSumAFDisc = ยอดส่วนลดที่ได้
'RETURN VALUE  #
'*KT 53-01-28 Communication Sheet PH3-ML-SB-074 Remain Discount for Employee.xls
'*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133 ใส่ ROUND(x,2)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tCardID$
    Dim cSumAFDisc As Double
On Error GoTo Err_Handler
    tSql = "SELECT FTScdCardID AS FTScdCardID,SUM(FCScdAmtAF) AS FCScdAmtAF"
    tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD & " WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND  FNDctNo = 12"
    tSql = tSql & vbCrLf & " AND  FTSdtStkCode <> '1'"
    tSql = tSql & vbCrLf & " GROUP BY FTScdCardID"
    If SP_SQLvExecute(tSql, orsTemp) = 0 Then
        Do While Not orsTemp.EOF
            tCardID = SP_FEDtChkString(orsTemp, "FTScdCardID")
            cSumAFDisc = SP_FEDvChkNumeric(orsTemp, "FCScdAmtAF")
            tSql = "UPDATE  TCNMEmpCard WITH(ROWLOCK) "
            Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnIPV, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVRedeem, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVVIPService, tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax '
                      tSql = tSql & vbCrLf & "SET FCEcdBalDis = ROUND(FCEcdBalDis + " & CStr(SP_FMTcCurPoint(cSumAFDisc, True, False)) & "," & nVB_CNDecAmtForSav & ")"
                Case Else
                      tSql = tSql & vbCrLf & "SET FCEcdBalDis = ROUND(FCEcdBalDis - " & CStr(SP_FMTcCurPoint(cSumAFDisc, True, False)) & "," & nVB_CNDecAmtForSav & ")"
             End Select
             tSql = tSql & vbCrLf & "WHERE FTEcdCode = '" & tCardID & "'"
             Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
             orsTemp.MoveNext
        Loop
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
Err_Handler:
    Call SP_DATxRsNothing(orsTemp)
End Sub

Public Sub C_EMPxUpdLimit()
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 53-01-28
'DESCRIPTION       # สั่ง Run Query เพื่อ Update Employee ที่ Member
'ARGUMENT LIST #
'RETURN VALUE  #
'*KT 53-01-28 Communication Sheet PH3-ML-SB-074 Remain Discount for Employee.xls
'*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim bOnline As Boolean
    Dim tSql As String
    
On Error GoTo Err_Handler
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & "cPSDisc" & " Events:C_EMPxUpdLimit[1]")
    If tVB_CNUpdateEmpLimit <> "" Then
        bOnline = Not (ocnMember Is Nothing)
        If bOnline = False Then
           If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
                If SP_CONbMemberServer Then
                    bOnline = True
                End If
            End If
        End If
        If bOnline Then
            Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnIPV, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVRedeem, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVVIPService, tEN_TrnDepositCancel, tEN_TrnIPVReturn '
                    Call C_EMPxUpdMemberEx
                    Call SP_SQLvExecute(tVB_CNUpdateEmpLimit, , ocnMember, , , , True)
                    '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                    tSql = "UPDATE TCNMEmpCard WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tVB_CNRemainDiscID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                    Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
                Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                    Call C_EMPxUpdMemberEx2
                    '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                    tSql = "UPDATE TCNMEmpCard WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tVB_CNRemainDiscID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                    Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
                Case tEN_TrnSuspended, tEN_TrnOpen, tEN_TrnClose, tEN_TrnStartBank, tEN_TrnPrintClose, tEN_TrnPickUp, tEN_TrnNoSale, tEN_TrnOtherDisbursement, tEN_TrnOtherReceive, tEN_TrnCreateVIP
                    'none '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                Case tEN_TrnTotalReport, tEN_TrnIPVOtherReceive, tEN_TrnIPVOtherDisbursement, tEN_TrnCopyTrans, tEN_TrnCopyFromDev, tEN_TrnCopyMaster
                    'none '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                Case Else
                    Call SP_SQLvExecute(tVB_CNUpdateEmpLimit, , ocnMember, , , , True)
                    '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                    tSql = "UPDATE TCNMEmpCard WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tVB_CNRemainDiscID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                    Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
            End Select
        Else
            '*Neung 55-02-10 Urgent Check for Employee Discount ------------------------------------------------------------------------------
           If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
                If SP_CONbMemberServer Then
                    bOnline = True
                End If
            End If
            If bOnline Then
                Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnIPV, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVRedeem, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVVIPService, tEN_TrnDepositCancel, tEN_TrnIPVReturn '
                        Call C_EMPxUpdMemberEx
                        Call SP_SQLvExecute(tVB_CNUpdateEmpLimit, , ocnMember, , , , True)
                        '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                        tSql = "UPDATE TCNMEmpCard WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tVB_CNRemainDiscID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                        Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
                    Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                        Call C_EMPxUpdMemberEx2
                        '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                        tSql = "UPDATE TCNMEmpCard WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tVB_CNRemainDiscID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                        Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
                    Case tEN_TrnSuspended, tEN_TrnOpen, tEN_TrnClose, tEN_TrnStartBank, tEN_TrnPrintClose, tEN_TrnPickUp, tEN_TrnNoSale, tEN_TrnOtherDisbursement, tEN_TrnOtherReceive, tEN_TrnCreateVIP
                        'none '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                    Case tEN_TrnTotalReport, tEN_TrnIPVOtherReceive, tEN_TrnIPVOtherDisbursement, tEN_TrnCopyTrans, tEN_TrnCopyFromDev, tEN_TrnCopyMaster
                        'none '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                    Case Else
                        Call SP_SQLvExecute(tVB_CNUpdateEmpLimit, , ocnMember, , , , True)
                        '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                        tSql = "UPDATE TCNMEmpCard  WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tVB_CNRemainDiscID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                        Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
                End Select
            Else
                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & "cPSDisc" & " Events:C_EMPxUpdLimit[4]")
            End If
            '*Neung 55-02-10 Urgent Check for Employee Discount ------------------------------------------------------------------------------
        End If
    Else
        Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnIPV, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVRedeem, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVVIPService, tEN_TrnDepositCancel '
                        bOnline = Not (ocnMember Is Nothing)
                        If bOnline = False Then
                           If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
                                If SP_CONbMemberServer Then
                                    bOnline = True
                                End If
                            End If
                        End If
                        If bOnline Then
                                Call C_EMPxUpdMemberEx
                                Call SP_SQLvExecute(tVB_CNUpdateEmpLimit, , ocnMember, , , , True)
                                '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                                tSql = "UPDATE TCNMEmpCard WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tVB_CNRemainDiscID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                                Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
                        Else
                            '*Neung 55-02-10 Urgent Check for Employee Discount ------------------------------------------------------------------------------
                            If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
                                If SP_CONbMemberServer Then
                                    bOnline = True
                                End If
                            End If
                            If bOnline Then
                                Call C_EMPxUpdMemberEx
                                Call SP_SQLvExecute(tVB_CNUpdateEmpLimit, , ocnMember, , , , True)
                                '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                                tSql = "UPDATE TCNMEmpCard WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tVB_CNRemainDiscID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                                Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
                            Else
                                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & "cPSDisc" & " Events:C_EMPxUpdLimit[4]")
                            End If
                            '*Neung 55-02-10 Urgent Check for Employee Discount ------------------------------------------------------------------------------
                        End If
                Case tEN_TrnReturn, tEN_TrnIPVReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                        bOnline = Not (ocnMember Is Nothing)
                        If bOnline = False Then
                           If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
                                If SP_CONbMemberServer Then
                                    bOnline = True
                                End If
                            End If
                        End If
                        If bOnline Then
                                Call C_EMPxUpdMemberEx2
                                '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                                tSql = "UPDATE TCNMEmpCard WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tVB_CNRemainDiscID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                                Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
                        Else
                            '*Neung 55-02-10 Urgent Check for Employee Discount ------------------------------------------------------------------------------
                            If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
                                If SP_CONbMemberServer Then
                                    bOnline = True
                                End If
                            End If
                            If bOnline Then
                                Call C_EMPxUpdMemberEx2
                                '*KT 53-10-26 หมายเลข Commsheet PH3-ML-SF-133
                                tSql = "UPDATE TCNMEmpCard WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tVB_CNRemainDiscID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                                Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
                            Else
                                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & "cPSDisc" & " Events:C_EMPxUpdLimit[4]")
                            End If
                            '*Neung 55-02-10 Urgent Check for Employee Discount ------------------------------------------------------------------------------
                        End If
        End Select
    End If
    Exit Sub
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & "cPSDisc" & " Events:C_EMPxUpdLimit[2]")
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & "cPSDisc" & " Events:C_EMPxUpdLimit[3]")
End Sub

Private Function C_GETcDiscEmpCardEx(ByVal porsTemp As ADODB.Recordset) As Double
'---------------------------------------------------------------------------------------------------------------------------------------------
'
'   Cmt: Ao 2013/03/04   TK    ให้คิด Emp ไม่เหมือน The Mall
'---------------------------------------------------------------------------------------------------------------------------------------------
Dim aEmpCard() As UDT_EmpCard

Dim nSeqNo As Integer, nI As Integer
Dim bEmpCard As Boolean
Dim tFTSkuCode As String, tFTSkuDepCode As String, tFTSkuBrdCode As String, tStaChkAllowBuy As String, tEmpNoFrm As String, tEmpNoTo As String
Dim tWhere As String, tStaChkLimit As String
Dim cSaleAmt As Double, cFCSdtQtyAllDisAvi As Double, cDptRate As Double, cCalAmt As Double
Dim cSkuRate As Double, cDisSumAmt As Double, cDisSumAmtEx As Double
Dim cDisEMP As Double

On Error GoTo Err_Handler
    nC_ENLnDisType = nEN_LnDis1Per
    cC_DiscPer = cC_DiscValue
    cC_DiscAmt = 0
    cC_DiscNew = 0
    cC_DiscPerOrg = cC_DiscPer
    cC_DiscAmtOrg = cC_DiscAmt
    cC_DiscNewOrg = cC_DiscNew
    tStaChkAllowBuy = cUT.UT_STRtToken(tC_RefCode, ",", 1, True)
    tStaChkLimit = cUT.UT_STRtToken(tC_RefCode, ",", 2, True)
    tEmpNoFrm = cUT.UT_STRtToken(tC_CusID, ",", 1, True)
    tEmpNoTo = cUT.UT_STRtToken(tC_CusID, ",", 2, True)
    
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        nSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
        '-------------------------------------------------------------------------------
        If SP_FEDtChkString(porsTemp, "FTSkuStaMallEmpCard") = "Y" Then
            Select Case C_CHKnDDP(nSeqNo) 'Check Double Policy
                Case 0 'ไม่อนุญาต
                    'คิดเฉพาะส่วนที่เหลือ
                    bEmpCard = False
                Case 1 'อนุญาต
                    bEmpCard = True
                Case 2 'หาผลที่ดีที่สุด
                   bEmpCard = True
            End Select
        Else
            bEmpCard = False
        End If
'        Select Case C_CHKnDDP(nSeqNo) 'Check Double Policy
'            Case 0 'ไม่อนุญาต
'                'คิดเฉพาะส่วนที่เหลือ
'                bEmpCard = False
'            Case 1 'อนุญาต
'                bEmpCard = True
'            Case 2 'หาผลที่ดีที่สุด
'               bEmpCard = True
'        End Select
        '------------------------------------------------------------------------------
        If bEmpCard Then
            nI = nI + 1
            ReDim Preserve aEmpCard(nI) As UDT_EmpCard
            tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
            tFTSkuDepCode = SP_FEDtChkString(porsTemp, "FTSkuDepCode")
            tFTSkuBrdCode = SP_FEDtChkString(porsTemp, "FTSkuBrdCode")
            cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
'            cBluePrice = SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
'            cYellowPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
            cSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
            
            aEmpCard(nI).udt_nSeqNo = nSeqNo
            aEmpCard(nI).udt_tSkuCode = tFTSkuCode
            aEmpCard(nI).udt_tDptCode = tFTSkuDepCode
            aEmpCard(nI).udt_cQtyAllAvi = cFCSdtQtyAllDisAvi
            If UCase(tC_StaChkDepart) = "Y" Then
                If C_EMPbGetDptRate("TCNMEmpCardDep", "FCMpcRate", tFTSkuDepCode, tFTSkuBrdCode, tC_CardKey, cDptRate) Then
                    aEmpCard(nI).udt_cCalRate = cDptRate
                Else
                    tWhere = "WHERE FTSkuCode = '" & tFTSkuCode & "'"
                    tWhere = tWhere & " AND FTSkuStaEmpCard = 'Y'"
                    cSkuRate = SP_VALcText2Double(SP_SQLtFunction("", "FCSkuEmpCardRate", tVB_CNTblSku, tWhere, oDB.ocnOnLine, True, True)) 'read master
'                    If cSkuRate = 0 Then
'                        aEmpCard(nI).udt_cCalRate = cC_DiscValue
'                    Else
'                        aEmpCard(nI).udt_cCalRate = IIf(cC_DiscValue > cSkuRate, cSkuRate, cC_DiscValue)
'                    End If
                    aEmpCard(nI).udt_cCalRate = IIf(cC_DiscValue > cSkuRate, cSkuRate, cC_DiscValue)
                End If
            Else
                tWhere = "WHERE FTSkuCode = '" & tFTSkuCode & "'"
                tWhere = tWhere & " AND FTSkuStaEmpCard = 'Y'"
                cSkuRate = SP_VALcText2Double(SP_SQLtFunction("", "FCSkuEmpCardRate", tVB_CNTblSku, tWhere, oDB.ocnOnLine, True, True)) 'read master
'                If cSkuRate = 0 Then
'                    aEmpCard(nI).udt_cCalRate = cC_DiscValue
'                Else
'                    aEmpCard(nI).udt_cCalRate = IIf(cC_DiscValue > cSkuRate, cSkuRate, cC_DiscValue)
'                End If
                aEmpCard(nI).udt_cCalRate = IIf(cC_DiscValue > cSkuRate, cSkuRate, cC_DiscValue)
            End If
            If aEmpCard(nI).udt_cCalRate = 0 Then
                aEmpCard(nI).udt_cB4CalAmt = 0
                aEmpCard(nI).udt_cB4UnCalAmt = cSaleAmt
            Else
                aEmpCard(nI).udt_cB4CalAmt = cSaleAmt
                aEmpCard(nI).udt_cB4UnCalAmt = cSaleAmt
            End If
            aEmpCard(nI).udt_cB4Disc = aEmpCard(nI).udt_cB4CalAmt * aEmpCard(nI).udt_cCalRate / 100
            aEmpCard(nI).udt_cB4CalBalance = aEmpCard(nI).udt_cB4CalAmt - aEmpCard(nI).udt_cB4Disc
            aEmpCard(nI).udt_cB4AllCalBalance = aEmpCard(nI).udt_cB4CalBalance + aEmpCard(nI).udt_cB4UnCalAmt
            
            aEmpCard(nI).udt_cAFCalAmt = aEmpCard(nI).udt_cB4CalAmt
            aEmpCard(nI).udt_cAFUnCalAmt = aEmpCard(nI).udt_cB4UnCalAmt
            aEmpCard(nI).udt_cAFDisc = aEmpCard(nI).udt_cB4Disc
            aEmpCard(nI).udt_cAFCalBalance = aEmpCard(nI).udt_cB4CalBalance
            aEmpCard(nI).udt_cAFAllCalBalance = aEmpCard(nI).udt_cAFUnCalAmt
            aEmpCard(nI).udt_cAllLimit = 0
        End If
        porsTemp.MoveNext
    Loop
    
    cDisSumAmt = 0
    cDisSumAmtEx = 0
    cCalAmt = 0
    cDisEMP = 0
    If nI > 0 Then
        For nI = 1 To UBound(aEmpCard)
            If nI = UBound(aEmpCard) Then
                bEmpCard = True
            Else
                bEmpCard = False
            End If
            '_____________________________________
'            cCalAmt = cCalAmt + aEmpCard(nI).udt_cB4CalAmt                              'Ao 2013/07/31 เพิ่ม  tC_CardKey
'            Call C_EMPxInsEmp(aEmpCard(nI).udt_cAFDisc, aEmpCard(nI).udt_nSeqNo, tC_CardKey, aEmpCard(nI).udt_cQtyAllAvi, aEmpCard(nI).udt_cAFCalAmt, aEmpCard(nI).udt_cCalRate, cDisSumAmt, cDisSumAmtEx, bEmpCard)
'            If aEmpCard(nI).udt_cCalRate > 0 Then
'                Call C_EMPxUpdateSingEmp(aEmpCard(nI).udt_nSeqNo)
'            End If
            cCalAmt = cCalAmt + aEmpCard(nI).udt_cB4CalAmt
            If Not bC_ChkEMP Then
                Call C_EMPxInsEmp(aEmpCard(nI).udt_cAFDisc, aEmpCard(nI).udt_nSeqNo, tC_CardKey, aEmpCard(nI).udt_cQtyAllAvi, aEmpCard(nI).udt_cAFCalAmt, aEmpCard(nI).udt_cCalRate, cDisSumAmt, cDisSumAmtEx, bEmpCard)
                If aEmpCard(nI).udt_cCalRate > 0 Then
                    Call C_EMPxUpdateSingEmp(aEmpCard(nI).udt_nSeqNo)
                End If
            Else                '*Ao 2013/10/31   ไม่ได้ใช้แล้ว
'                Call C_EMPxDisEmp(aEmpCard(nI).udt_cAFDisc, aEmpCard(nI).udt_nSeqNo, tC_CardKey, aEmpCard(nI).udt_cQtyAllAvi, aEmpCard(nI).udt_cAFCalAmt, aEmpCard(nI).udt_cCalRate, cDisSumAmt, cDisSumAmtEx, cDisEMP, bEmpCard)
            End If
            '_____________________________________
        Next nI
        If Not bC_ChkEMP Then
            Call DISxUpdateDiscTable(cCalAmt)
        End If
    End If
    
'    C_GETcDiscEmpCardEx = cDisSumAmt                        '*Eaw 56-05-30 CommSheet TK-ISS3000-211-ST 00 00 01.xls
    C_GETcDiscEmpCardEx = W_GETxDiscEmpCard
    If Not bC_ChkEMP Then
        cC_DisEMP = cCalAmt                     '*Ao 2013/10/29
    Else
        cC_DisEMP = cCalAmt                     '*Ao 2013/10/29
    End If
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_GETcDiscEmpCard(ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # หาส่วนลด Employee Card
'ARGUMENT LIST # porsTemp = Recordset of DT
'
'RETURN VALUE  # ส่วนลด Employee Card
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDisSumAmt As Double
    
    Dim nFNSdtSeqNo As Integer
    Dim tFTSkuCode As String
    Dim tFTSkuDepCode As String
    Dim tFTSkuBrdCode As String
    Dim cFCSdtQtyAllDisAvi As Double
    
    Dim cDptRate As Double
    Dim dCurrentExpDate As Date
    
    Dim aEmpCard() As UDT_EmpCard
    Dim nI As Integer
    Dim cBluePrice As Double
    Dim cYellowPrice As Double

    Dim cSumCalBalance As Double
    Dim cSumLimit As Double
    Dim tStaChkLimit As String
    Dim tStaChkAllowBuy As String
    Dim tEmpNoFrm As String
    Dim tEmpNoTo As String
    Dim bOnline As Boolean
    Dim tWhere As String
    Dim cSkuRate As Double
    '*TW 51-08-15
    Dim cSumSpecial As Double
    Dim cSumSpecialBlue As Double
    Dim cSumSpecialReg As Double
    Dim cSumSpecialRegOrg As Double
    '----------------------------
    
    Dim bEmpCard As Boolean '***KT 51-11-31 V 1.0.1
    Dim cB4Disc As Double
    Dim tFilter As String
    Dim tSql As String
    Dim bChk As Boolean
    Dim cDisSumAmtEx  As Double
    
On Error GoTo Err_Handler
    bOnline = Not (ocnMember Is Nothing)
    If bOnline = False Then
        '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
       ' If ocnAll.ConnectionString <> oDB.ocnOffLine.ConnectionString And bC_ReConnectMember Then 'PH 2.0.0 OLD
        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) And bC_ReConnectMember Then
            If SP_CONbMemberServer(, , False) Then
                bOnline = True
            End If
        End If
    End If
    If UCase(tC_StaChkExpire) = "Y" Then
        dCurrentExpDate = DateAdd("m", 1, dC_CardExpDate) - 1
'        If dCurrentExpDate < Format(dVB_CurrentDate, "mm/dd/yyyy") Then Exit Function '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  ใส่ Format mm/dd/yyyyy เพราะ dVB_CurrentDate มีเวลาเข้ามาด้วย Communication Sheet PH2-ML-SF-027 Credit Expire.xls
        '*Ao 52-02-05  เนื่องจาก dCurrentExpDate Format เป็น dd/mm/yyyy  ทำให้การเปรียบเทียบผิด
        'เช่น วันนี้ 05/02/2009 -> Format(05/02/2009, "mm/dd/yyyy")=02/05/2009
        '01/04/2009 <  02/05/2009
        'If CDate(Format(dCurrentExpDate, "YYYY/MM/DD")) < CDate(Format(SP_GETdChkDateDisc, "YYYY/MM/DD")) Then Exit Function  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุของ point ตามที่กำหนดใน parameter
         If CDate(Format(dCurrentExpDate, "YYYY/MM/DD")) < CDate(Format(dVB_CurrentDate, "YYYY/MM/DD")) Then Exit Function  'ทาง FSBT ต้องการให้ตรวจสอบวันที่ Expire Date ตามวันที่ System Date
    End If
    tStaChkAllowBuy = cUT.UT_STRtToken(tC_RefCode, ",", 1, True)
    tStaChkLimit = cUT.UT_STRtToken(tC_RefCode, ",", 2, True)
    tEmpNoFrm = cUT.UT_STRtToken(tC_CusID, ",", 1, True)
    tEmpNoTo = cUT.UT_STRtToken(tC_CusID, ",", 2, True)
    
    '*KT 52-02-02  PH2-ML-SF-014 แก้ไขใน POSFront 1.5.6 '*KT 52-02-11 PH2-ML-SF-090โ ทำให้ Sub all ผิด เมื่อมีส่วนลดแบบ Employee discount
    '-----------------------------------------------------------------------------
    Select Case C_CHKnDDPEmp() 'Check Double Policy ของ Emp card # Emp card
             Case 0 'ไม่อนุญาต
                 'ให้ลด
                  Exit Function
    End Select
    '---------------------------------------------------------------------------
    
    nC_ENLnDisType = nEN_LnDis1Per
    cC_DiscPer = cC_DiscValue
    cC_DiscAmt = 0
    cC_DiscNew = 0
    cC_DiscPerOrg = cC_DiscPer
    cC_DiscAmtOrg = cC_DiscAmt
    cC_DiscNewOrg = cC_DiscNew
    
    If UCase(tStaChkAllowBuy) = "Y" And bOnline Then 'ตรวจสอบเวลาซื้อ
        If Not C_EMPbChkAlwBuy(tC_CardID, cSumLimit, tStaChkAllowBuy) Then
            bW_ChkMallCard = True
            Call C_DATxInsertDiscCD(0, 1, "", 0, 0, 0)
            Exit Function
        End If
    End If
    
    porsTemp.MoveFirst
    ReDim aEmpCard(0) As UDT_EmpCard
    
    
    nI = 0
    '*TW 51-08-15
    cSumSpecialRegOrg = 1
    cSumSpecialReg = 1
    cSumSpecial = C_MALcSpec
    If cSumSpecial > 0 Then
        cSumSpecialBlue = 0
        cSumSpecialRegOrg = 0
        porsTemp.MoveFirst
        Do While Not porsTemp.EOF
            cSumSpecialRegOrg = cSumSpecialRegOrg + SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
            cSumSpecialBlue = cSumSpecialBlue + SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt") - SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
            porsTemp.MoveNext
        Loop
        cSumSpecialReg = cSumSpecialRegOrg - cSumSpecial
    End If
    
    
    '*KT 52-03-02 member price
'    '--------------------------------------------------------------------------------------
'    If bVB_CNMemberPrice Then
'        If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" Then
'             cB4Disc = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChg")
'              tFilter = "(FTSdtStaRfd ='N') OR  (FCSdtMemberPrice <> " & cB4Disc & ")"
'        Else
'             tFilter = ""
'        End If
'    Else
'          tFilter = ""
'    End If
'   If tFilter <> "" Then porsTemp.Filter = tFilter
    '-----------------------------------------------------------------------------------------
    
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
        '***KT 51-11-29 V 1.0.1
        '-------------------------------------------------------------------------------
        Select Case C_CHKnDDP(nFNSdtSeqNo) 'Check Double Policy
                Case 0 'ไม่อนุญาต
                    'คิดเฉพาะส่วนที่เหลือ
                    bEmpCard = False
                Case 1 'อนุญาต
                    bEmpCard = True
                Case 2 'หาผลที่ดีที่สุด
                   bEmpCard = True
            End Select
        '------------------------------------------------------------------------------
        If bEmpCard Then
            nI = nI + 1
            ReDim Preserve aEmpCard(nI) As UDT_EmpCard
            tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
            tFTSkuDepCode = SP_FEDtChkString(porsTemp, "FTSkuDepCode")
            tFTSkuBrdCode = SP_FEDtChkString(porsTemp, "FTSkuBrdCode")
            cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
            cBluePrice = SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
            cYellowPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
            
            aEmpCard(nI).udt_nSeqNo = nFNSdtSeqNo
            aEmpCard(nI).udt_tSkuCode = tFTSkuCode
            aEmpCard(nI).udt_tDptCode = tFTSkuDepCode
            aEmpCard(nI).udt_cQtyAllAvi = cFCSdtQtyAllDisAvi
            If UCase(tC_StaChkDepart) = "Y" Then
                If C_EMPbGetDptRate("TCNMEmpCardDep", "FCMpcRate", tFTSkuDepCode, tFTSkuBrdCode, tC_CardKey, cDptRate) Then
                    aEmpCard(nI).udt_cCalRate = cDptRate
                Else
                    tWhere = "WHERE FTSkuCode = '" & tFTSkuCode & "'"
                    '*KT 55-08-09 Takashimaya PRJ10001TK-External Specification SKU-00.01.00.doc
                    'tWhere = tWhere & " AND FTSkuStaMallEmpCard = 'Y'"
                    tWhere = tWhere & " AND FTSkuStaEmpCard = 'Y'"
                    cSkuRate = SP_VALcText2Double(SP_SQLtFunction("", "FCSkuEmpCardRate", tVB_CNTblSku, tWhere, oDB.ocnOnLine, True, True)) 'read master
                    aEmpCard(nI).udt_cCalRate = IIf(cC_DiscValue > cSkuRate, cSkuRate, cC_DiscValue)
                End If
            Else
                tWhere = "WHERE FTSkuCode = '" & tFTSkuCode & "'"
                '*KT 55-08-09 Takashimaya PRJ10001TK-External Specification SKU-00.01.00.doc
               ' tWhere = tWhere & " AND FTSkuStaMallEmpCard = 'Y'"
                 tWhere = tWhere & " AND FTSkuStaEmpCard = 'Y'"
                cSkuRate = SP_VALcText2Double(SP_SQLtFunction("", "FCSkuEmpCardRate", tVB_CNTblSku, tWhere, oDB.ocnOnLine, True, True)) 'read master
                aEmpCard(nI).udt_cCalRate = IIf(cC_DiscValue > cSkuRate, cSkuRate, cC_DiscValue)
            End If
            
            If aEmpCard(nI).udt_cCalRate = 0 Then
                aEmpCard(nI).udt_cB4CalAmt = 0
                aEmpCard(nI).udt_cB4UnCalAmt = cYellowPrice + cBluePrice
            Else
                aEmpCard(nI).udt_cB4CalAmt = cBluePrice
                aEmpCard(nI).udt_cB4UnCalAmt = cYellowPrice
            End If
            aEmpCard(nI).udt_cB4Disc = aEmpCard(nI).udt_cB4CalAmt * aEmpCard(nI).udt_cCalRate / 100 'Old
            aEmpCard(nI).udt_cB4CalBalance = aEmpCard(nI).udt_cB4CalAmt - aEmpCard(nI).udt_cB4Disc 'CalBal = Cal-Disc
            aEmpCard(nI).udt_cB4AllCalBalance = aEmpCard(nI).udt_cB4CalBalance + aEmpCard(nI).udt_cB4UnCalAmt 'All Balance
            
            aEmpCard(nI).udt_cAFCalAmt = aEmpCard(nI).udt_cB4CalAmt
            aEmpCard(nI).udt_cAFUnCalAmt = aEmpCard(nI).udt_cB4UnCalAmt
            aEmpCard(nI).udt_cAFDisc = aEmpCard(nI).udt_cB4Disc
            aEmpCard(nI).udt_cAFCalBalance = aEmpCard(nI).udt_cB4CalBalance
            aEmpCard(nI).udt_cAFAllCalBalance = aEmpCard(nI).udt_cAFUnCalAmt
            aEmpCard(nI).udt_cAllLimit = 0
        
        End If
        porsTemp.MoveNext
    Loop
    cDisSumAmt = 0
    cDisSumAmtEx = 0
    
    If UCase(tStaChkLimit) = "Y" And bOnline Then 'ตรวจสอบเวลาซื้อ
        Call C_EMPbChkAlwBuy(tC_CardID, cSumLimit, "N") '*Ao 09-12-18
        cSumCalBalance = 0
        For nI = 1 To UBound(aEmpCard)
            cSumCalBalance = cSumCalBalance + aEmpCard(nI).udt_cB4CalBalance
        Next nI
        '*KT 53-10-20 หมายเลข Commsheet PH3-ML-SF-133
        '--------------------------------------------------------------------------------------------------------------------
        If cSumLimit >= 1 And cSumCalBalance > cSumLimit Then
            For nI = 1 To UBound(aEmpCard)
                aEmpCard(nI).udt_cAllLimit = aEmpCard(nI).udt_cB4CalBalance / cSumCalBalance * cSumLimit
                aEmpCard(nI).udt_cAFCalAmt = (aEmpCard(nI).udt_cAllLimit * 100) / (100 - aEmpCard(nI).udt_cCalRate)
                aEmpCard(nI).udt_cAFUnCalAmt = aEmpCard(nI).udt_cAFUnCalAmt + (aEmpCard(nI).udt_cB4CalAmt - aEmpCard(nI).udt_cAFCalAmt)
                aEmpCard(nI).udt_cAFDisc = aEmpCard(nI).udt_cAFCalAmt * aEmpCard(nI).udt_cCalRate / 100
                aEmpCard(nI).udt_cAFCalBalance = aEmpCard(nI).udt_cAFCalAmt - aEmpCard(nI).udt_cAFDisc
                aEmpCard(nI).udt_cAFAllCalBalance = aEmpCard(nI).udt_cAFCalBalance + aEmpCard(nI).udt_cAFUnCalAmt
                '*Ao 2011/06/22  CommSheet ML-PRV-061 Employee Disc Over limit.xls
                If nI = UBound(aEmpCard) Then
                    bChk = True
                Else
                    bChk = False
                End If
                '_____________________________________
                '*KT 52-12-14 PH3-ML-SF-044 Wrong Sub Reg when apply Employee Discount ส่ง parameter (aEmpCard(nI).udt_cAFCalAmt + aEmpCard(nI).udt_cAFDisc) กรณีที่ Remain discount น้อยกว่ายอดที่จะคิดหาส่วนลดตั้งแต่ต้น
'                Call C_EMPxInsEmp(aEmpCard(nI).udt_cAFDisc, aEmpCard(nI).udt_nSeqNo, "", aEmpCard(nI).udt_cQtyAllAvi, (aEmpCard(nI).udt_cAFCalAmt + aEmpCard(nI).udt_cAFDisc), aEmpCard(nI).udt_cCalRate, cDisSumAmt, cDisSumAmtEx, bChk)
                '*Ao 2011/06/28  CommSheet ML-PRV-062-SubTotal Reg Incorrect.xls
                Call C_EMPxInsEmp(aEmpCard(nI).udt_cAFDisc, aEmpCard(nI).udt_nSeqNo, "", aEmpCard(nI).udt_cQtyAllAvi, aEmpCard(nI).udt_cAFCalAmt, aEmpCard(nI).udt_cCalRate, cDisSumAmt, cDisSumAmtEx, bChk)
            Next nI
            cSumCalBalance = cSumLimit
        Else
            If cSumLimit >= 1 Then
                For nI = 1 To UBound(aEmpCard)
                    '*Ao 2011/06/22  CommSheet ML-PRV-061 Employee Disc Over limit.xls
                    If nI = UBound(aEmpCard) Then
                        bChk = True
                    Else
                        bChk = False
                    End If
                    '_____________________________________
                    Call C_EMPxInsEmp(aEmpCard(nI).udt_cAFDisc, aEmpCard(nI).udt_nSeqNo, "", aEmpCard(nI).udt_cQtyAllAvi, aEmpCard(nI).udt_cAFCalAmt, aEmpCard(nI).udt_cCalRate, cDisSumAmt, cDisSumAmtEx, bChk)
                Next nI
            Else 'Remain Discount ใน Member Server =< 0
                '*KT 53-11-03 หมายเลข Commsheet PH3-ML-SF-133 ได้รับแจ้งจาก Chanwit Atiponakarapon ให้เก็บค่าเป็น 0
                'Call SP_MSGnShowing("tms_CN005179", nCS_Warning)
                aEmpCard(1).udt_cAllLimit = 0
                aEmpCard(1).udt_cAFCalAmt = 0
                aEmpCard(1).udt_cAFUnCalAmt = aEmpCard(1).udt_cAFUnCalAmt + (aEmpCard(1).udt_cB4CalAmt - aEmpCard(1).udt_cAFCalAmt)
                aEmpCard(1).udt_cAFDisc = 0
                aEmpCard(1).udt_cAFCalBalance = aEmpCard(1).udt_cAFCalAmt - aEmpCard(1).udt_cAFDisc
                aEmpCard(1).udt_cAFAllCalBalance = aEmpCard(1).udt_cAFCalBalance + aEmpCard(1).udt_cAFUnCalAmt
'                Call C_EMPxInsEmp(aEmpCard(1).udt_cAFDisc, aEmpCard(1).udt_nSeqNo, "", aEmpCard(1).udt_cQtyAllAvi, (aEmpCard(1).udt_cAFCalAmt + aEmpCard(1).udt_cAFDisc), aEmpCard(1).udt_cCalRate, cDisSumAmt, cDisSumAmtEx, False)
                '*Ao 2011/06/28  CommSheet ML-PRV-062-SubTotal Reg Incorrect.xls
                Call C_EMPxInsEmp(aEmpCard(1).udt_cAFDisc, aEmpCard(1).udt_nSeqNo, "", aEmpCard(1).udt_cQtyAllAvi, aEmpCard(1).udt_cAFCalAmt, aEmpCard(1).udt_cCalRate, cDisSumAmt, cDisSumAmtEx, False)
                'กรณี Remain Discount  น้อยกว่า 0 ทำการ Update Remain Discount   = 0 ใน Member Server
               'If (cSumLimit > 0 And cSumLimit < 1) Or (cSumLimit < 0 And cSumLimit > -1) Or (cSumLimit < 0) Then
               If cSumLimit <> 0 Then
                    'tSql = "UPDATE TCNMEmpCard SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tC_CardID & "' AND ((FCEcdBalDis > 0 AND FCEcdBalDis < 1) OR (FCEcdBalDis < 0 AND FCEcdBalDis > - 1) )"
                    tSql = "UPDATE TCNMEmpCard WITH(ROWLOCK) SET FCEcdBalDis = 0 WHERE FTEcdCode = '" & tC_CardID & "'"
                    Call SP_SQLvExecute(tSql, , ocnMember, , , , True)
                End If
                cSumCalBalance = 0
            End If
        End If
        cSumCalBalance = CDbl(SP_FMTcCurPoint(cSumCalBalance, True, False))
        Call C_EMPxUpdMember(tC_CardID, cSumCalBalance) 'ส่งส่วนที่ลดแล้วเข้าไป
        '--------------------------------------------------------------------------------------------------------------------
    Else
        For nI = 1 To UBound(aEmpCard)
            '*Ao 2011/06/22  CommSheet ML-PRV-061 Employee Disc Over limit.xls
            If nI = UBound(aEmpCard) Then
                bChk = True
            Else
                bChk = False
            End If
            '_____________________________________
            Call C_EMPxInsEmp(aEmpCard(nI).udt_cAFDisc, aEmpCard(nI).udt_nSeqNo, "", aEmpCard(nI).udt_cQtyAllAvi, aEmpCard(nI).udt_cAFCalAmt, aEmpCard(nI).udt_cCalRate, cDisSumAmt, cDisSumAmtEx, bChk)
        Next nI
        
    End If

    C_GETcDiscEmpCard = cDisSumAmt
    Exit Function
Err_Handler:
    C_GETcDiscEmpCard = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Sub C_CALxExceptSale(ByVal ptWhe As String)
'    Dim nRow As Integer
'    With wMain.ogdForm
'        For nRow = .FixedRows To .rows - 1
'            If UCase(.TextMatrix(nRow, .ColIndex("FTSdtCancelFlag"))) <> "Y" And Trim(.TextMatrix(nRow, .ColIndex("FNSdtSeqNo"))) <> "" Then
'                Call SP_CALxExceptSale(wMain.ogdForm, nRow, SP_VALcText2Double(SP_SQLtFunction("", "FCSdtSalePrice", tVB_CNTblDT, ptWhe & " AND FNSdtSeqNo = " & nRow)))
'            End If
'        Next nRow
'    End With
'
'End Sub

Private Sub C_DISxDeal()
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # ใส่ข้อมูล Deal ลง CD
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDiscount As Double
    Call C_DATxInsertCDDeal(cC_DiscValue, nC_DTSeqNo, tC_EventID, cC_AlwDisQty, cC_AlwDis, cC_AlwDisQty)
    Select Case nC_ENLnDisType
        Case nEN_LnDis1Per
            cC_DiscPer = cC_Balance
        Case nEN_LnDis2Amt
            cC_DiscAmt = cC_DiscValue
        Case nEN_LnDis3NewPrice
            cC_DiscNew = cC_DiscValue
        Case nEN_LnDis4Free
            cC_DiscAmt = cC_DiscValue
    End Select
End Sub

Private Function C_CHKbUnseq(ByVal pnDiscType As Integer) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    On Error GoTo Err_Handler
    '*KT 52-09-23 Analyse code
    '*KT 54-12-02 V.4.0.9 ML-PH4-R3-014 สามารถเลือก Transaction Markdown 1 -5 สลับกันได้  และการให้ส่วนลดเป็นไปตาม Double Discount
    Select Case pnDiscType
        Case nEN_Disc14TranMDown, nEN_Disc23TranMDown1, nEN_Disc24TranMDown2, nEN_Disc25TranMDown3, nEN_Disc26TranMDown4, nEN_Disc27TranMDown5
             tSql = "IF EXISTS(SELECT FNScdSeqNo FROM " & tVB_CNTblCD & " WHERE (FNDctNo > " & pnDiscType & ") AND (FNDctNo  NOT IN (21,22,14,23,24,25,26,27)) AND " & tVB_CNWheBill & ")"
        Case Else
             tSql = "IF EXISTS(SELECT FNScdSeqNo FROM " & tVB_CNTblCD & " WHERE (FNDctNo > " & pnDiscType & ") AND (FNDctNo  NOT IN (21,22)) AND " & tVB_CNWheBill & ")"
    End Select
    tSql = tSql & vbCrLf & "SELECT 1 AS Result"
    tSql = tSql & vbCrLf & "ELSE"
    tSql = tSql & vbCrLf & "SELECT 0 AS Result"
    If SP_SQLbExecute(tSql, orsTemp) Then
        If SP_FEDvChkNumeric(orsTemp, "Result") = 1 Then
            C_CHKbUnseq = True
        Else
            C_CHKbUnseq = False
        End If
    Else
        C_CHKbUnseq = False
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function C_TMPbRet2Org(ByVal ptOrgName$, ByVal ptTmpName$) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    On Error GoTo Err_Handler
    tSql = "DELETE FROM " & ptOrgName
    tSql = tSql & vbCrLf & " WHERE " & tVB_CNWheBill
    If Not SP_SQLbExecute(tSql) Then GoTo Err_Handler
            
    tSql = "INSERT INTO " & ptOrgName
    tSql = tSql & vbCrLf & "SELECT * FROM " & ptTmpName
    tSql = tSql & vbCrLf & " WHERE " & tVB_CNWheBill
    If Not SP_SQLbExecute(tSql) Then GoTo Err_Handler
    
    If SP_TBLbChkExist(ptTmpName) Then
        tSql = "DROP TABLE " & ptTmpName
        If Not SP_SQLbExecute(tSql) Then GoTo Err_Handler
    End If
    C_TMPbRet2Org = True
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    C_TMPbRet2Org = False
End Function

'Public Sub DISxDropTmpSpec()
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  #
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim tSql As String
'    On Error GoTo Err_Handler
'    If SP_TBLbChkExist(tVB_CNTblCD & "OntopTmp") Then
'        tSql = "DROP TABLE " & tVB_CNTblCD & "OntopTmp"
'        If Not SP_SQLbExecute(tSql) Then GoTo Err_Handler
'    End If
'
'    If SP_TBLbChkExist(tVB_CNTblDT & "OntopTmp") Then
'        tSql = "DROP TABLE " & tVB_CNTblDT & "OntopTmp"
'        If Not SP_SQLbExecute(tSql) Then GoTo Err_Handler
'    End If
'
'    Exit Sub
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub

Private Function C_TMPbOrs2Array(ByVal porsTemp As ADODB.Recordset, ByRef parrTemp()) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  51-09-03                                   LAST MODIFIED DATE     # 51-09-03
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nCol As Integer
    Dim nI As Integer
    Dim nRow As Integer
    On Error GoTo Err_Handler
    nCol = porsTemp.Fields.Count
    ReDim parrTemp(nCol - 1, 0)
    If Not (porsTemp.BOF And porsTemp.EOF) Then
        For nI = 0 To nCol - 1
            parrTemp(nI, 0) = porsTemp.Fields(nI).Name
        Next nI
        Do While Not porsTemp.EOF
            nRow = nRow + 1
            ReDim Preserve parrTemp(nCol - 1, nRow)
            For nI = 0 To nCol - 1
                Select Case UCase(Mid(porsTemp.Fields(nI).Name, 2, 1))
                    Case "T"
                        parrTemp(nI, nRow) = SP_FEDtChkString(porsTemp, porsTemp.Fields(nI).Name)
                    Case "D"
                        parrTemp(nI, nRow) = Format(SP_FEDdChkDate(porsTemp, porsTemp.Fields(nI).Name), "YYYY/MM/DD")
                    Case "N", "C"
                        parrTemp(nI, nRow) = SP_FEDvChkNumeric(porsTemp, porsTemp.Fields(nI).Name)
                End Select
            Next nI
            porsTemp.MoveNext
        Loop
    End If
    C_TMPbOrs2Array = True
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    C_TMPbOrs2Array = False
End Function

Private Function C_TMPbArray2DB(ByRef parrTemp(), ByVal ptTblName As String) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  51-09-03                                   LAST MODIFIED DATE     # 51-09-03
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nCol As Integer
    Dim nI As Integer
    Dim nJ As Integer
    Dim nRow As Integer
    Dim tSql As String
    Dim tValue As String
    Dim orsTemp As ADODB.Recordset
    On Error GoTo Err_Handler
    If UBound(parrTemp, 2) > 0 Then
        tSql = "SELECT TOP 1 * FROM " & ptTblName & " WITH(NOLOCK)"
        If SP_SQLbExecute(tSql, orsTemp) Then
            nCol = orsTemp.Fields.Count
            tSql = "TRUNCATE TABLE " & ptTblName
            If SP_SQLbExecute(tSql) Then
                For nJ = 1 To UBound(parrTemp, 2)
                    tValue = ""
                    For nI = 0 To nCol - 1
                        Select Case UCase(Mid(orsTemp.Fields(nI).Name, 2, 1))
                            Case "T"
                                tValue = tValue & ",'" & SP_STRtSiggleCode(parrTemp(nI, nJ)) & "'"
                            Case "D"
                                tValue = tValue & ",'" & parrTemp(nI, nJ) & "'"
                            Case "N", "C"
                                tValue = tValue & "," & parrTemp(nI, nJ)
                        End Select
                    Next nI
                    tValue = Right(tValue, Len(tValue) - 1)
                    tSql = "INSERT INTO " & ptTblName
                    tSql = tSql & vbCrLf & "VALUES(" & tValue & ")"
                    Call SP_SQLbExecute(tSql)
                Next nJ
            End If
        End If
    Else
        tSql = "TRUNCATE TABLE " & ptTblName
        If SP_SQLbExecute(tSql) Then
        End If
    End If
    C_TMPbArray2DB = True
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
    C_TMPbArray2DB = False
End Function

Private Function C_TMPbCreate(ByVal ptOrgName$, ByVal ptTmpName$) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-09-03
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    On Error GoTo Err_Handler
    If SP_TBLbChkExist(ptTmpName) Then
        tSql = "DROP TABLE " & ptTmpName
        If Not SP_SQLbExecute(tSql) Then GoTo Err_Handler
    End If
    tSql = "SELECT * INTO " & ptTmpName
    tSql = tSql & vbCrLf & "FROM " & ptOrgName
    tSql = tSql & vbCrLf & " WHERE " & tVB_CNWheBill
    If Not SP_SQLbExecute(tSql) Then GoTo Err_Handler
    C_TMPbCreate = True
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    C_TMPbCreate = False
End Function

Private Sub C_DISxSpecial()
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsYellow As ADODB.Recordset
    Dim orsBlue As ADODB.Recordset
    Dim cSumYellowPrice As Double
    Dim cSumBluePrice As Double
    Dim cSumDiscExcept As Double
    Dim cSumDiscLeft As Double
    Dim cAllDiscValue As Double
    Dim tSql As String
    Dim tWhe As String
    On Error GoTo Err_Handler
    cSumYellowPrice = C_GETcB4DisDDPSpecial(False, orsYellow)
    If cC_DiscValue - cSumYellowPrice > 0 Then
            cSumBluePrice = C_GETcB4DisDDPSpecial(True, orsBlue)
            cAllDiscValue = cC_DiscValue - cSumYellowPrice
            If cAllDiscValue > cSumBluePrice Then cAllDiscValue = cSumBluePrice
            cSumDiscExcept = C_GETcSpecialWithPmt(orsBlue, cAllDiscValue, cSumBluePrice, cAllDiscValue)
            cAllDiscValue = 0 'cSumYellowPrice '*KT 51-09-15 แก้ไขเรื่องจ่ายด้วย Special Gift Voucher/Special Coupon หลังหักส่วนลด Cooperate Coupon และ Mall Card
    Else
        cAllDiscValue = cC_DiscValue
    End If

    If cSumYellowPrice > 0 Then
        '*KT 51-09-15 แก้ไขเรื่องจ่ายด้วย Special Gift Voucher/Special Coupon หลังหักส่วนลด Cooperate Coupon และ Mall Card
        'cSumDiscLeft = C_GETcSpecialWithPmt(orsYellow, cAllDiscValue, cSumYellowPrice, (-1) * cAllDiscValue) 'cAllDiscValue = ยอดลดรวมทั้งหมด,cSumYellowPrice = ยอดที่สามารถลดได้ทั้งหมด
        cSumDiscLeft = C_GETcSpecialWithPmt(orsYellow, cAllDiscValue, cSumYellowPrice, cAllDiscValue) 'cAllDiscValue = ยอดลดรวมทั้งหมด,cSumYellowPrice = ยอดที่สามารถลดได้ทั้งหมด
    End If

    Call SP_DATxRsNothing(orsBlue)
    Call SP_DATxRsNothing(orsYellow)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsBlue)
    Call SP_DATxRsNothing(orsYellow)
End Sub

Private Sub C_CLRxSpecial()
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    On Error GoTo Err_Handler
    If tVB_Running = "" Then Exit Sub
    tSql = "DELETE FROM " & tVB_CNTblCD & " WITH(ROWLOCK) " 'tpscd08107tmpontop
    tSql = tSql & vbCrLf & " WHERE FNDctNo = " & nEN_Disc20Special
    tSql = tSql & vbCrLf & " AND " & tVB_CNWheBill
    Call SP_SQLbExecute(tSql)
    Call C_DATxSumDT
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub C_CLRxMCrdSpecial(ByVal pbClrSpec As Boolean)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    On Error GoTo Err_Handler
    If tVB_Running = "" Then Exit Sub
    
    tSql = "DELETE FROM " & tVB_CNTblCD & " WITH(ROWLOCK) " 'tpscd08107tmpontop
    tSql = tSql & vbCrLf & " WHERE FNDctNo >= " & nEN_Disc11MallCrd & " AND FNDctNo NOT IN (" & nEN_Disc21MPrice & "," & nEN_Disc22MDisc & ")"   'ทำไมต้องใช้ >= ตอบ เพราะต้องเอา Mall Card ออกด้วย '*KT 52-09-24 Member Price + Consignment
    If Not pbClrSpec Then
        tSql = tSql & vbCrLf & " AND FNDctNo <> " & nEN_Disc20Special
    End If
    tSql = tSql & vbCrLf & " AND " & tVB_CNWheBill
    Call SP_SQLbExecute(tSql)
    Call C_DATxSumDT
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Public Function DISbClearDiscDT(Optional ByVal pnByType As eEN_VerifyDiscType = nEN_Type3None) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตั้งต้น 3.0.19
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim nEffect As Long
On Error GoTo Err_Handler
    tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) SET " 'tpsdt08107tmpontop
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnDutyFreeSale
            '*AO 51-06-18  แก้ไข เพราะ  FCSdtB4DisChg หักค่ามัดจำออกไปแล้ว จึงใม่ต้อง -FCSdtDutyDeposit
            tSql = tSql & vbCrLf & " FCSdtSalePrice = ROUND(FCSdtB4DisChg,2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND(FCSdtB4DisChg,2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleAmt = ROUND((FCSdtB4DisChg) * FCSdtQty,2) "
            tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = FCSdtQty"
            tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi = ROUND((FCSdtB4DisChg) * FCSdtQty,2) "
        Case tEN_TrnTaxExceptSale
            '*KT 52-05-07 ใช้คำสั่ง ROUND ก่อนแล้วค่อยคูณ qty
            tSql = tSql & vbCrLf & " FCSdtSalePrice = ROUND(FCSdtB4DisChg * 100 / (100 + FCSdtTax),2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND(FCSdtB4DisChg * 100 / (100 + FCSdtTax),2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleAmt = ROUND((FCSdtB4DisChg * 100 / (100 + FCSdtTax)),2) * FCSdtQty"
            tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = ROUND(FCSdtQty,2)"
            tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi = ROUND((FCSdtB4DisChg * 100 / (100 + FCSdtTax)),2) * FCSdtQty"
        Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
            tSql = tSql & vbCrLf & " FCSdtSalePrice = ROUND(FCSdtB4DisChg,2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND(FCSdtB4DisChg,2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleAmt = ROUND((FCSdtB4DisChg) * FCSdtQty,2) "
            tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = ROUND(FCSdtQty,2)"
            tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi = ROUND((FCSdtB4DisChg) * FCSdtQty,2) "
            '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
            'ทำการกลับค่าของ FCSdtBluePrice,FCSdtYellowPrice กรณีที่ทำ return แบบ Refer เพราะการ Cancel มีการ Recal ถ้าเป็นการ Recal ส่วนลด Mall card จะต้องใช้ FCSdtBluePrice,FCSdtYellowPrice
            If bVB_Reference Then  'แบบ  Refer
                tSql = tSql & vbCrLf & " ,FCSdtBSubReg= CASE WHEN FTSkuStaMallEmpCard = 'N' THEN 0 ELSE ROUND(FCSdtB4DisChg,2) END"
                tSql = tSql & vbCrLf & " ,FCSdtBSubPro=CASE WHEN FTSkuStaMallEmpCard = 'N' THEN ROUND(FCSdtB4DisChg,2) ELSE 0 END"
                tSql = tSql & vbCrLf & " ,FCSdtBluePrice=CASE WHEN FTSkuStaMallEmpCard = 'N' THEN 0 ELSE ROUND(FCSdtB4DisChg,2) END"
                tSql = tSql & vbCrLf & " ,FCSdtYellowPrice=CASE WHEN FTSkuStaMallEmpCard = 'N' THEN ROUND(FCSdtB4DisChg,2) ELSE 0 END"
            End If
        Case Else
            tSql = tSql & vbCrLf & " FCSdtSalePrice = ROUND(FCSdtB4DisChg,2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND(FCSdtB4DisChg,2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleAmt = ROUND((FCSdtB4DisChg) * FCSdtQty,2) "
            tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = ROUND(FCSdtQty,2)"
            tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi = ROUND((FCSdtB4DisChg) * FCSdtQty,2) "
    End Select
    tSql = tSql & vbCrLf & " ,FCSdtDis = 0"
    tSql = tSql & vbCrLf & " WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
    '*KT 52-09-23 Analyse code
    Select Case pnByType
        Case nEN_Type1ByItem
            tSql = tSql & vbCrLf & "AND FNSdtSeqNo = " & nC_DTSeqNo
        Case nEN_Type2ByDcn
            tSql = tSql & vbCrLf & "AND FNSdtSeqNo IN (SELECT DISTINCT FNSdtSeqNo FROM " & tVB_CNTblCD
            tSql = tSql & vbCrLf & "                                              WHERE " & tVB_CNWheBill
            tSql = tSql & vbCrLf & "                                              AND FNDctNo = " & nC_ENDiscType & ")"
        Case Else
            'none
    End Select
    If SP_SQLbExecute(tSql, , False, nEffect) Then
        If nEffect > 0 Then
            DISbClearDiscDT = True
        Else
            DISbClearDiscDT = False
        End If
    Else
        DISbClearDiscDT = False
    End If
    Exit Function
Err_Handler:
    DISbClearDiscDT = False
End Function

Public Function DISbClearDiscDTByItem() As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim nEffect As Long
    On Error GoTo Err_Handler
    tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) SET "
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnDutyFreeSale
            '*AO 51-06-18  แก้ไข เพราะ  FCSdtB4DisChg หักค่ามัดจำออกไปแล้ว จึงใม่ต้อง -FCSdtDutyDeposit
            tSql = tSql & vbCrLf & " FCSdtSalePrice = ROUND(FCSdtB4DisChg,2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND(FCSdtB4DisChg,2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleAmt = ROUND((FCSdtB4DisChg) * FCSdtQty,2) "
            tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = FCSdtQty"
            tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi = ROUND((FCSdtB4DisChg) * FCSdtQty,2) "
        Case tEN_TrnTaxExceptSale
            '*KT 52-05-07 ใช้คำสั่ง ROUND ก่อนแล้วค่อยคูณ qty
            tSql = tSql & vbCrLf & " FCSdtSalePrice = ROUND(FCSdtB4DisChg * 100 / (100 + FCSdtTax),2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND(FCSdtB4DisChg * 100 / (100 + FCSdtTax),2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleAmt = ROUND((FCSdtB4DisChg * 100 / (100 + FCSdtTax)),2) * FCSdtQty"
            tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = ROUND(FCSdtQty,2)"
            tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi = ROUND((FCSdtB4DisChg * 100 / (100 + FCSdtTax)),2) * FCSdtQty"
        Case Else
            tSql = tSql & vbCrLf & " FCSdtSalePrice = ROUND(FCSdtB4DisChg,2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND(FCSdtB4DisChg,2)"
            tSql = tSql & vbCrLf & " ,FCSdtSaleAmt = ROUND((FCSdtB4DisChg) * FCSdtQty,2) "
            tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = ROUND(FCSdtQty,2)"
            tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi = ROUND((FCSdtB4DisChg) * FCSdtQty,2) "
    End Select
    tSql = tSql & vbCrLf & " ,FCSdtDis = 0"
    tSql = tSql & vbCrLf & " WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & " AND  FNSdtSeqNo = " & nC_DTSeqNo
    tSql = tSql & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
    
    If SP_SQLbExecute(tSql, , False, nEffect) Then
        If nEffect > 0 Then
            DISbClearDiscDTByItem = True
        Else
            DISbClearDiscDTByItem = False
        End If
    Else
        DISbClearDiscDTByItem = False
    End If
    Exit Function
Err_Handler:
    DISbClearDiscDTByItem = False
End Function


Public Sub C_DATxSumDT(Optional ByVal pnByType As eEN_VerifyDiscType = nEN_Type3None)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    Dim nEffect As Long
    Dim nOldDctNo As Integer
    Dim nNewDctNo As Integer
    Dim nDTSeqNo As Integer
    Dim bExceptOccur As Boolean
    Dim nI As Integer
    Dim cDisc As Double
On Error GoTo Err_Handler
    
    '1. Select New Disc From CD
    tSql = "SELECT FNSdtSeqNo ,FNDctNo,FCScdAmt,FCScdQtyB4-FCScdQtyAF AS FCScdQty,FCScdAmtB4Avi,FCScdAmtAFAvi,FCScdAmtAF,FTScdDiscType"
    tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD & " WITH(NOLOCK) " 'tpscd08107tmpontop
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
    '*KT 52-09-23 Analyse code
    Select Case pnByType
        Case nEN_Type1ByItem
            tSql = tSql & vbCrLf & "AND FNSdtSeqNo = " & nC_DTSeqNo
        Case nEN_Type2ByDcn
            tSql = tSql & vbCrLf & "AND FNSdtSeqNo IN (SELECT DISTINCT FNSdtSeqNo FROM " & tVB_CNTblCD
            tSql = tSql & vbCrLf & "                                              WHERE " & tVB_CNWheBill
            tSql = tSql & vbCrLf & "                                              AND FNDctNo = " & nC_ENDiscType & ")"
        Case Else
            'none
    End Select
    tSql = tSql & vbCrLf & "AND  ISNULL(FTSdtStkCode,'') <> '1'" '*KT 52-11-25
    tSql = tSql & vbCrLf & "ORDER BY FNScdSeqNo"
    If Not SP_SQLbExecute(tSql, orsTemp) Then Exit Sub
    
    '2. Clear Old All DT Discount and Promo
    If Not DISbClearDiscDT(pnByType) Then Exit Sub '*KT 52-09-23 Analyse code

    '3.Update DT
    nOldDctNo = 0
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        Do While Not orsTemp.EOF
            nDTSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
            nNewDctNo = SP_FEDvChkNumeric(orsTemp, "FNDctNo")
            If tVB_TransactionSub = tEN_TrnTaxExceptSale Then
                If Not bVB_CNStaMain Then
                    If (nOldDctNo <= nEN_Disc04CoopDisc And nNewDctNo > nEN_Disc04CoopDisc) Then
                        bExceptOccur = True
                    End If
                End If
            End If
            nOldDctNo = nNewDctNo
            cDisc = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdAmt"))
            Dim cAmtAF As Double
            Dim tDiscType As String
            
            cAmtAF = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdAmtAF"))
            tDiscType = SP_FEDtChkString(orsTemp, "FTScdDiscType")
            
            tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) SET "
            '----------------------------------------------------------------------------------------------------------------------------------------'*Ao 2013/03/27   TK   Line Over
            Select Case nNewDctNo
                Case nEN_Disc05LnDisc, nEN_Disc06LnMDown
                    If (tDiscType = "3") And (cDisc = 0) Then
                        tSql = tSql & vbCrLf & " FCSdtSaleAmt = ROUND(" & cAmtAF & ",2)"
                    Else
                        tSql = tSql & vbCrLf & " FCSdtSaleAmt = ROUND(FCSdtSaleAmt - (" & cDisc & "),2)"
                    End If
                Case Else
                    tSql = tSql & vbCrLf & " FCSdtSaleAmt = ROUND(FCSdtSaleAmt - (" & cDisc & "),2)"
            End Select
            '----------------------------------------------------------------------------------------------------------------------------------------
'            tSql = tSql & vbCrLf & " FCSdtSaleAmt = ROUND(FCSdtSaleAmt - (" & cDisc & "),2)"
            Select Case nNewDctNo
'                Case nEN_Disc01Deal, nEN_Disc02PmtTime, nEN_Disc03Pmt, nEN_Disc05LnDisc, nEN_Disc06LnMDown, nEN_Disc04CoopDisc, nEN_Disc21MPrice, nEN_Disc22MDisc
                Case nEN_Disc01Deal, nEN_Disc02PmtTime, nEN_Disc03Pmt, nEN_Disc04CoopDisc, nEN_Disc21MPrice, nEN_Disc22MDisc
                    tSql = tSql & vbCrLf & " ,FCSdtSalePrice = CASE WHEN FCSdtQty = 0 THEN FCSdtSalePrice ELSE ROUND(FCSdtSalePrice - (" & cDisc & "/FCSdtQty),2) END"
                    tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = CASE WHEN FCSdtQty = 0 THEN FCSdtSalePrice ELSE ROUND(FCSdtSalePrice - (" & cDisc & "/FCSdtQty),2) END"    '*Ao 51-05-08   FCSdtSalePrice=FCSdtSaleUnit
                Case nEN_Disc05LnDisc, nEN_Disc06LnMDown
                    If (tDiscType = "3") And (cDisc = 0) Then
                        tSql = tSql & vbCrLf & " ,FCSdtSalePrice = ROUND((" & cAmtAF & "/FCSdtQty),2)"           '*Ao 2013/03/27   TK   Line Over"
                        tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND((" & cAmtAF & "/FCSdtQty),2)"
                    Else
                        tSql = tSql & vbCrLf & " ,FCSdtSalePrice = CASE WHEN FCSdtQty = 0 THEN FCSdtSalePrice ELSE ROUND(FCSdtSalePrice - (" & cDisc & "/FCSdtQty),2) END"
                        tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = CASE WHEN FCSdtQty = 0 THEN FCSdtSalePrice ELSE ROUND(FCSdtSalePrice - (" & cDisc & "/FCSdtQty),2) END"    '*Ao 51-05-08   FCSdtSalePrice=FCSdtSaleUnit
                    End If
                Case Else
                    '*KT 51-12-11 Phase 2 แก้ไขเรื่อง Price Overide สินค้าที่ไม่ใช่ Prompt Sale ไม่รู้กระทบอะไรหรือไม่ ต้องตรวจสอบทั้ง Service Doc /Report/Upload
                    'tSql = tSql & vbCrLf & " ,FCSdtSalePrice = CASE WHEN FCSdtQty = 0 THEN FCSdtSalePrice ELSE ROUND(FCSdtSalePrice - (" & SP_FEDvChkNumeric(orsTemp, "FCScdAmt") & "/FCSdtQty),2) END"
            End Select
            '*KT 52-12-28 แก้ไข DCR PRJ09010TheMall-R3-DCR_12002(Deal)-00.02.00.xls
            '*KT 53-01-08 Deal & Member Price
            Select Case nNewDctNo
                Case nEN_Disc02PmtTime, nEN_Disc03Pmt
                    tSql = tSql & vbCrLf & " ,FCSdtSalePricePmt = ROUND(FCSdtB4DisChg - (" & cDisc & "/FCSdtQty),2)"
                    tSql = tSql & vbCrLf & " ,FTSdtStaRfd = '" & IIf(tVB_StaMemPriceAlw = "", "N", tVB_StaMemPriceAlw) & "'"
                Case nEN_Disc21MPrice, nEN_Disc22MDisc
                    tSql = tSql & vbCrLf & " ,FCSdtSalePricePmt = ROUND(FCSdtB4DisChg - (" & cDisc & "/FCSdtQty),2)"
            End Select
            tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = ROUND(FCSdtQtyAllDisAvi -  " & SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdQty"), True, False) & ",2)"
            tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi = CASE WHEN " & SP_FEDvChkNumeric(orsTemp, "FCScdAmtAFAvi") & " < 0 "     '*Ao 51-07-25
            tSql = tSql & vbCrLf & " THEN 0 ELSE " & SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCScdAmtAFAvi"), True, False) & " END"
            tSql = tSql & vbCrLf & " ,FCSdtDis = FCSdtDis + " & cDisc
            tSql = tSql & vbCrLf & " ,FCSdtVatable = CASE WHEN FTTaxCode = '00' THEN 0 ELSE ROUND(FCSdtSaleAmt * 100/(FCSdtTax+100),2) END "
            tSql = tSql & vbCrLf & " ,FCSdtVat = CASE WHEN FTTaxCode = '00' THEN 0 ELSE ROUND(FCSdtSaleAmt * FCSdtTax/(FCSdtTax+100),2) END "
            tSql = tSql & vbCrLf & SP_GETtWhoUpdate
            tSql = tSql & vbCrLf & " WHERE " & tVB_CNWheBill
            tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nDTSeqNo
            If Not SP_SQLbExecute(tSql, , False, nEffect) Then
                Call SP_MSGnShowing(tC_Msg002, tCS_CNMsgWarn)
            Else
                If nEffect <= 0 And nDTSeqNo > 0 Then
                    Call SP_MSGnShowing(tC_Msg002, tCS_CNMsgWarn)
                End If
            End If
            orsTemp.MoveNext
        Loop
        
        If tVB_TransactionSub <> tEN_TrnTaxExceptSale Then
            tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) SET "
            tSql = tSql & vbCrLf & " FCSdtSaleUnit = (FCSdtSaleAmt/FCSdtQty)"
            tSql = tSql & vbCrLf & " ,FCSdtVatable = CASE WHEN FTTaxCode = '00' THEN 0 ELSE ROUND(FCSdtSaleAmt*100/(100+FCSdtTax),2) END "
            tSql = tSql & vbCrLf & " ,FCSdtVat = CASE WHEN FTTaxCode = '00' THEN 0 ELSE ROUND(FCSdtSaleAmt*FCSdtTax/(100+FCSdtTax),2) END "
            tSql = tSql & vbCrLf & SP_GETtWhoUpdate
            tSql = tSql & vbCrLf & " WHERE " & tVB_CNWheBill
            Call SP_SQLbExecute(tSql)
        Else
            If (Not bExceptOccur) And (Not bVB_CNStaMain) Then
'                Call C_CALxExceptSale(tWhe)
            End If
        End If
        
        
        '*TW 51-07-31
'        For nI = wMainRet.ogdForm.FixedRows To wMainRet.ogdForm.rows - 1
'            If wMainRet.ogdForm.TextMatrix(nI, wMainRet.ogdForm.ColIndex("FTSkuCode")) <> "" Then
'                If wMainRet.ogdForm.TextMatrix(nI, wMainRet.ogdForm.ColIndex("FTSdtCancelFlag")) <> "1" Then
'                    Call C_SUMxRecalSubPro(nI)
'                End If
'            Else
'                Exit For
'            End If
'        Next nI
        '----------------------------------------------------------------------------------------
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Sub

'Public Function C_DATxSumDTMemberPrice(ByVal pnSeqNo As Long, ByVal ptSKUCode As String, ByVal pnQty As Long) As Boolean
Public Function C_DATxSumDTMemberPrice(ByVal pnSeqNo As Long, ByVal ptSKUCode As String) As Boolean
'CREATE BY            # Kitti                           LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 52-02-16
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    C_DATxSumDTMemberPrice = True

'    Dim tSql As String
'    Dim nEffect As Long
'    Dim orsTemp As ADODB.Recordset
'    Dim cTotalQty As Double
'    Dim cMaxQty As Double
'    Dim cMemAmt As Double
'    Dim tDiscType As String
'    Dim tWhr As String
'On Error GoTo Err_Handler
'
'    C_DATxSumDTMemberPrice = False
'    cTotalQty = 0
'    cMaxQty = 0
'
'    tWhr = " WHERE FTSkuCode = '" & ptSKUCode & "'"
'    tSql = "SELECT FNPmpMemMaxQty,FCPmpMemAmt,FTPmpStaDiscType"
'    tSql = tSql & vbCrLf & "FROM TCNMMemberPrice"
'    tSql = tSql & vbCrLf & tWhr
'    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then
'        'cMaxQty = CDbl(Val(Format(SP_SQLtFunction("", "FNPmpMemMaxQty", "TCNMMemberPrice", tWhr, oDB.ocnOnLine, True, True)))) ''read master
'        cMaxQty = SP_VALcText2Double(SP_FEDvChkNumeric(orsTemp, "FNPmpMemMaxQty"))
'        cMemAmt = SP_VALcText2Double(SP_FEDvChkNumeric(orsTemp, "FCPmpMemAmt"))
'        tDiscType = SP_FEDtChkString(orsTemp, "FTPmpStaDiscType")
'        Select Case tDiscType
'            Case "1"
'                cUT_MemberPrice = SP_FMTcCurPoint(cUT_SalePrice - (cUT_SalePrice * cMemAmt / 100), True, False) '%off
'                cVB_CNMemPriceAmtDis = (cUT_SalePrice * cMemAmt / 100)
'                cVB_CNMemPriceAmt = cMemAmt
'            Case "2"
'                cUT_MemberPrice = SP_FMTcCurPoint(cUT_SalePrice - cMemAmt, True, False) 'amount off
'                cVB_CNMemPriceAmtDis = cMemAmt
'                cVB_CNMemPriceAmt = cMemAmt
'            Case "3"
'                cUT_MemberPrice = SP_FMTcCurPoint(cMemAmt, True, False) 'new price
'                cVB_CNMemPriceAmtDis = cUT_SalePrice - cMemAmt
'                cVB_CNMemPriceAmt = cMemAmt
'            Case Else 'No promotion PH2-ML-SF-147 Merber Price ถ้าเปลี่ยน Promotion Price เป็น No Promotion แต่ วัน End Date ยัง Active อยู่ เมื่อทำการขาย Sale Amount จะเป็น 0.00 ที่ถูกต้องจะต้องขายราคาปรกติ
'                Call SP_DATxRsNothing(orsTemp)
'                Exit Function
'        End Select
'    Else
'        Call SP_DATxRsNothing(orsTemp)
'        Exit Function
'    End If
'
'    tWhr = " WHERE " & tVB_CNWheBill
'    tWhr = tWhr & vbCrLf & " AND FTSkuCode = '" & ptSKUCode & "'"
'    tWhr = tWhr & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
'    tWhr = tWhr & vbCrLf & " AND FTSdtStaRfd = 'Y'"
''    tWhr = tWhr & vbCrLf & " AND FCSdtMemberPrice = FCSdtB4DisChg"
'    tWhr = tWhr & vbCrLf & " AND ISNULL(FTSdtEventMemPrice,'') <> ''" '*KT 52-09-08 Member Price + Consignment
'    cTotalQty = CDbl(Val(Format(SP_SQLtFunction("SUM", "FCSdtQty", tVB_CNTblDT, tWhr))))
'
'    If (cTotalQty < cMaxQty) Then
'        bVB_SKUisPromotion = True
'        tSql = "UPDATE " & tVB_CNTblDT & " SET "
'        Select Case Val(Format(tVB_TransactionSub))
'            Case tEN_TrnDutyFreeSale
''                tSql = tSql & vbCrLf & " FCSdtB4DisChg = ROUND(FCSdtMemberPrice,2)" '*KT 52-09-08 Member Price + Consignment
'                tSql = tSql & vbCrLf & " FCSdtSalePrice = ROUND(FCSdtMemberPrice,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND(FCSdtMemberPrice,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtSaleAmt = ROUND((FCSdtMemberPrice) * FCSdtQty,2) "
'                tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = ROUND(FCSdtQty,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi =0"
'            Case tEN_TrnTaxExceptSale
''                tSql = tSql & vbCrLf & " FCSdtB4DisChg = ROUND(FCSdtMemberPrice,2)" '*KT 52-09-08 Member Price + Consignment
'                tSql = tSql & vbCrLf & " FCSdtSalePrice = ROUND(FCSdtMemberPrice * 100 / (100 + FCSdtTax),2)"
'                tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND(FCSdtMemberPrice * 100 / (100 + FCSdtTax),2)"
'                tSql = tSql & vbCrLf & " ,FCSdtSaleAmt = ROUND((FCSdtMemberPrice * 100 / (100 + FCSdtTax)) * FCSdtQty,2) "
'                tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = ROUND(FCSdtQty,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi =0"
'            Case Else
''                tSql = tSql & vbCrLf & " FCSdtB4DisChg = ROUND(FCSdtMemberPrice,2)" '*KT 52-09-08 Member Price + Consignment
'                tSql = tSql & vbCrLf & " FCSdtSalePrice = ROUND(FCSdtMemberPrice,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtSaleUnit = ROUND(FCSdtMemberPrice,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtSaleAmt = ROUND((FCSdtMemberPrice) * FCSdtQty,2) "
'                tSql = tSql & vbCrLf & " ,FCSdtQtyAllDisAvi = ROUND(FCSdtQty,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtAmtDisAvi =0"
'        End Select
'        tSql = tSql & vbCrLf & " ,FCSdtDis = 0"
'        tSql = tSql & vbCrLf & " ,FTSkuStaMallEmpCard = 'N'"
'        tSql = tSql & vbCrLf & " ,FTSdtStaRfd = 'Y'"
'        If bVB_SKUisPromotion Then
'                tSql = tSql & vbCrLf & " ,FCSdtBSubPro = ROUND((FCSdtMemberPrice) * FCSdtQty,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtYellowPrice= ROUND((FCSdtMemberPrice) * FCSdtQty,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtBSubReg = 0"
'                tSql = tSql & vbCrLf & " ,FCSdtBluePrice = 0"
'                tSql = tSql & vbCrLf & " ,FCSdtDiscSubPro = 0"
'                tSql = tSql & vbCrLf & " ,FCSdtDiscSubReg = 0"
'        Else
'                tSql = tSql & vbCrLf & " ,FCSdtBSubPro = 0"
'                tSql = tSql & vbCrLf & " ,FCSdtYellowPrice= 0"
'                tSql = tSql & vbCrLf & " ,FCSdtBSubReg = ROUND((FCSdtMemberPrice) * FCSdtQty,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtBluePrice = ROUND((FCSdtMemberPrice) * FCSdtQty,2)"
'                tSql = tSql & vbCrLf & " ,FCSdtDiscSubPro = 0"
'                tSql = tSql & vbCrLf & " ,FCSdtDiscSubReg = 0"
'        End If
'        tSql = tSql & vbCrLf & " WHERE " & tVB_CNWheBill
'        tSql = tSql & vbCrLf & " AND FNSdtSeqNo =" & pnSeqNo
'        tSql = tSql & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
'        nEffect = 0
'        If SP_SQLvExecute(tSql, , , , , nEffect) = 0 Then
'                If nEffect > 0 Then
'                    '*KT 52-09-07 MemberPrice + Consignment
'                    Call DISxClearAll
'                    '------------------------------------------------------------------------------------
'                    bRecalByItem = False
''                    nScdSeqNo = pnScdSeqNo
'                    dTranDate = tUT_SaleDate
'                    tTerminalNo = tVB_CNTerminalNum
'                    tTransNo = tVB_Running
'                    nDTSeqNo = pnSeqNo
'                    If UCase(tVB_SKUConsignFlag) = "Y" Then
'                        oDisc.nENDiscType = nEN_Disc22MDisc
'                    Else
'                        oDisc.nENDiscType = nEN_Disc21MPrice
'                    End If
'                    tEventID = tVB_CNEventMemPrice
'                    cDiscValue = cVB_CNMemPriceAmt
'                    Select Case tVB_CNMemPriceType
'                        Case "1" '% Off
'                            nENLnDisType = nEN_LnDis1Per
'                            cDiscPer = cVB_CNMemPriceAmt
'                        Case "2" 'Amount Off
'                            nENLnDisType = nEN_LnDis2Amt
'                            cDiscAmt = cVB_CNMemPriceAmt
'                        Case "3" 'New Price
'                            nENLnDisType = nEN_LnDis3NewPrice
'                            cDiscNew = cVB_CNMemPriceAmt
'                        Case "4" 'No Promotion
'                            Call SP_DISxShwDisc2Grid(pnSeqNo, wMain.ogdForm)
''                            If bPass Then ocnAll.RollbackTrans
'                            Exit Function
'                        Case Else
'                            Call SP_DISxShwDisc2Grid(pnSeqNo, wMain.ogdForm)
''                            If bPass Then ocnAll.RollbackTrans
'                            Exit Function
'                    End Select
'                    cDiscPerOrg = cDiscPer
'                    cDiscAmtOrg = oDisc.cDiscAmt
'                    cDiscNewOrg = oDisc.cDiscNew
'                    nActSeqNo = pnSeqNo
'                    tCardID = xVB_CNMemberP.ID
'                    tCardType = xVB_CNMemberP.MallCardType
'                    dCardExpDate = xVB_CNMemberP.MallCardExpire
''                    Call oDisc.C_DATxInsertDiscCD(oDisc.cC_DiscAmt, oDisc.nC_DTSeqNo, oDisc.tEventID, oDisc.cFCSdtQtyAllDisAvi, oDisc.cFCSdtAmtDisAvi, oDisc.cFCSdtQtyAllDisAvi)
'                    Call oDisc.C_DATxInsertDiscCD(cVB_CNMemPriceAmtDis * pnQty, pnSeqNo, tVB_CNEventMemPrice, pnQty, cUT_RegPrice * pnQty, pnQty)
'                    '----------------------------------------------------------------------------------
'                End If
'                C_DATxSumDTMemberPrice = True
'        End If
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'    Exit Function
'Err_Handler:
'    C_DATxSumDTMemberPrice = False
'    Call SP_DATxRsNothing(orsTemp)
'    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub C_DATxGetDataDT(ByRef porsTemp As ADODB.Recordset)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'RETURN VALUE  #
'กรณีที่เป็น Online ได้ต้อง read master จาก server ส่วนกรณีที่เป็น offline สามารถ read master จาก local ได้ ดังนั้นสามารถเขียน SQl Commmand เแบบ join กันได้
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsSKU As ADODB.Recordset
     Dim orsSKUPreDay  As ADODB.Recordset
    Dim orsDT As ADODB.Recordset
    Dim tSKUCode As String
    Dim nI As Long
    Dim bPass As Boolean
    Dim tFed As String
    Dim tWhr As String
On Error GoTo Err_Handler
    tSKUCode = ""
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
            tSql = "SELECT DT.FTTmnNum ,DT.FTShdTransNo ,DT.FNSdtSeqNo ,DT.FTSkuCode ,DT.FTShdTransType"
            tSql = tSql & vbCrLf & ",DT.FDShdTransDate ,DT.FTSkuDepCode ,DT.FTSkuBrdCode ,DT.FCSdtB4DisChg"
            tSql = tSql & vbCrLf & ",DT.FCLineOver,DT.FCSdtRegPrice"   '*Eaw 56-03-15
            '*KT 52-11-25 Deal&Promo
            If bC_PromoAftDeal Then
                Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnDutyFreeSale
                        tSql = tSql & vbCrLf & " ,ROUND(FCSdtB4DisChg,2) AS FCSdtSalePrice"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg) * FCSdtQty,2) AS FCSdtSaleAmt"
                        tSql = tSql & vbCrLf & " ,FCSdtQty AS FCSdtQtyAllDisAvi"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg) * FCSdtQty,2) AS FCSdtAmtDisAvi"
                    Case tEN_TrnTaxExceptSale
                        tSql = tSql & vbCrLf & " ,ROUND(FCSdtB4DisChg * 100 / (100 + FCSdtTax),2) AS FCSdtSalePrice"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg * 100 / (100 + FCSdtTax)),2) * FCSdtQty AS FCSdtSaleAmt"
                        tSql = tSql & vbCrLf & " ,ROUND(FCSdtQty,2) AS FCSdtQtyAllDisAvi"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg * 100 / (100 + FCSdtTax)),2) * FCSdtQty AS FCSdtAmtDisAvi"
                    Case Else
                        tSql = tSql & vbCrLf & " ,ROUND(FCSdtB4DisChg,2) AS FCSdtSalePrice"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg) * FCSdtQty,2) AS FCSdtSaleAmt"
                        tSql = tSql & vbCrLf & " ,ROUND(FCSdtQty,2) AS FCSdtQtyAllDisAvi"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg) * FCSdtQty,2) AS FCSdtAmtDisAvi"
                End Select
            Else
                tSql = tSql & vbCrLf & ",DT.FCSdtSalePrice "
                tSql = tSql & vbCrLf & ",DT.FCSdtSaleAmt "
                tSql = tSql & vbCrLf & ",DT.FCSdtAmtDisAvi "
                tSql = tSql & vbCrLf & ",DT.FCSdtQtyAllDisAvi"
            End If
            tSql = tSql & vbCrLf & ",DT.FCSdtTax,DT.FCSdtQtyAll"
             '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
            Select Case nC_ENDiscType
                Case nEN_Disc11MallCrd
                    tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 'N' ELSE 'Y'  END AS FTSkuStaMallEmpCard"
                Case nEN_Disc12EmpCrd           '*Ao 2013/03/08  TK
                    tSql = tSql & vbCrLf & ",CASE WHEN SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),20,1) = '1' THEN 'Y' ELSE 'N' END AS FTSkuStaMallEmpCard"
                Case Else
                    tSql = tSql & vbCrLf & ",DT.FTSkuStaMallEmpCard"
            End Select
            tSql = tSql & vbCrLf & ",'' AS FTSkuStaOnTop"
            tSql = tSql & vbCrLf & ",DT.FCSdtQty"                '*Ao 51-06-19
            tSql = tSql & vbCrLf & ",0 AS FCSkuCoCpnRate,0 AS FCSkuCoBndRate"                '*Ao 51-06-19
             '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
            Select Case nC_ENDiscType
                Case nEN_Disc11MallCrd
                    If tC_MallAlwPoint = "Y" Then
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtBSubReg ELSE DT.FCSdtBSubPro  END AS FCSdtBSubPro"
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtDiscSubReg ELSE DT.FCSdtDiscSubPro  END AS FCSdtDiscSubPro"
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtBSubPro ELSE DT.FCSdtBSubReg  END  AS FCSdtBSubReg"
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtDiscSubPro ELSE DT.FCSdtDiscSubReg  END AS FCSdtDiscSubReg" '*Ao 51-07-31
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtYellowPrice ELSE DT.FCSdtBluePrice  END AS FCSdtBluePrice"
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtBluePrice ELSE DT.FCSdtYellowPrice  END AS FCSdtYellowPrice" '*TW 51-08-07
                        
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtBSubPro  END AS FCSdtBSubPro"
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtDiscSubPro  END AS FCSdtDiscSubPro"
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtBSubReg  END  AS FCSdtBSubReg"
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtDiscSubReg  END AS FCSdtDiscSubReg" '*Ao 51-07-31
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtBluePrice  END AS FCSdtBluePrice"
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtYellowPrice  END AS FCSdtYellowPrice" '*TW 51-08-07
                    Else
                        tSql = tSql & vbCrLf & ",DT.FCSdtBSubPro"
                        tSql = tSql & vbCrLf & ",DT.FCSdtDiscSubPro"
                        tSql = tSql & vbCrLf & ",DT.FCSdtBSubReg"
                        tSql = tSql & vbCrLf & ",DT.FCSdtDiscSubReg"
                        tSql = tSql & vbCrLf & ",DT.FCSdtBluePrice"
                        tSql = tSql & vbCrLf & ",DT.FCSdtYellowPrice"
                    End If
                Case Else
                    tSql = tSql & vbCrLf & ",DT.FCSdtBSubPro"
                    tSql = tSql & vbCrLf & ",DT.FCSdtDiscSubPro"
                    tSql = tSql & vbCrLf & ",DT.FCSdtBSubReg"
                    tSql = tSql & vbCrLf & ",DT.FCSdtDiscSubReg" '*Ao 51-07-31
                    tSql = tSql & vbCrLf & ",DT.FCSdtBluePrice"
                    tSql = tSql & vbCrLf & ",DT.FCSdtYellowPrice" '*TW 51-08-07
            End Select
            tSql = tSql & vbCrLf & ",'' AS FTSkuStaCoCpnAlw,'' AS FTSkuStaCoBndAlw,'' AS FTSkuStaMallCpn" '*TW 51-08-11
            tSql = tSql & vbCrLf & ",'1900/01/01' AS FDSkuCoCpnEndDate , '1900/01/01' AS FDSkuCoBndEndDate" '*KT 51-10-17 ให้ตรวจวันที่หมดอายุการใช้ Co.CpnและCo.Bnd ตาม SKU
            tSql = tSql & vbCrLf & ",'' AS FTSkuStaDisAlwTran,'1900/01/01' AS FDSkuDisAlwEndDate,'' AS FTSkuStaMrkAlwTran,'1900/01/01' AS FDSkuMrkAlwEndDate"
            tSql = tSql & vbCrLf & ",DT.FTSdtStaRfd,DT.FCSdtMemberPrice" '*KT 52-02-13 Member price
            tSql = tSql & vbCrLf & ",'' AS FTSkuStaMallEmpCardEx"
            tSql = tSql & vbCrLf & ",0 AS FCSkuMallCardRate"
            tSql = tSql & vbCrLf & ",0 AS FCSkuCoDisRate"
            tSql = tSql & vbCrLf & ",'' AS FTSkuEventCoDisCpn"
            tSql = tSql & vbCrLf & ",'' As FTSkuStaCoDisAlw"
            tSql = tSql & vbCrLf & ",'1900/01/01' AS FDSkuCoDisEndDate"
            tSql = tSql & vbCrLf & ",''  AS FTSkuStaAutoTran"
            tSql = tSql & vbCrLf & ",DT.FTSdtEventMemPrice" '*KT 52-09-08 Member Price + Consignment
            tSql = tSql & vbCrLf & ",DT.FTSdtStaPreview" '*KT 53-10-30 Preview Day
            tSql = tSql & vbCrLf & ",DT.FTSdtStaNotAlwPrv" '*KT 53-10-30 Preview Day
            tSql = tSql & vbCrLf & ",DT.FTSdtMemberPointAlw" '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
            tSql = tSql & vbCrLf & ",SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1) AS FTSdtStaPdt14" '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
            tSql = tSql & vbCrLf & ",DT.FTSdtConsignFlag" '*KT 54-06-23 V.4.0.0 Ph.4.0 Req.9 Auoto Transaction Discount
            tSql = tSql & vbCrLf & ",'' AS FTSkuStaMrkAlwTran1,'1900/01/01' AS FDSkuMrkAlwEndDate1"  '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
            tSql = tSql & vbCrLf & ",'' AS FTSkuStaMrkAlwTran2,'1900/01/01' AS FDSkuMrkAlwEndDate2"
            tSql = tSql & vbCrLf & ",'' AS FTSkuStaMrkAlwTran3,'1900/01/01' AS FDSkuMrkAlwEndDate3"
            tSql = tSql & vbCrLf & ",'' AS FTSkuStaMrkAlwTran4,'1900/01/01' AS FDSkuMrkAlwEndDate4"
            tSql = tSql & vbCrLf & ",'' AS FTSkuStaMrkAlwTran5,'1900/01/01' AS FDSkuMrkAlwEndDate5"
            tSql = tSql & vbCrLf & " FROM    " & tVB_CNTblDT & " DT WITH(NOLOCK)  "
            tSql = tSql & vbCrLf & " WHERE DT.FTTmnNum = '" & tVB_CNTerminalNum & "'"
            tSql = tSql & vbCrLf & " AND DT.FTShdTransNo = '" & tVB_Running & "'"
            tSql = tSql & vbCrLf & " AND DT.FDShdTransDate = '" & tUT_SaleDate & "'"
            tSql = tSql & vbCrLf & " AND ISNULL(DT.FTSdtCancelFlag,'0') <> '1'"
            '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
'            Select Case nC_ENDiscType
'                Case nEN_Disc11MallCrd
'                    If tC_MallAlwPoint = "Y" Then
'                        tSql = tSql & vbCrLf & " AND (ISNULL(DT.FTSdtMemberPointAlw,'N') ='N'  OR (ISNULL(DT.FTSdtMemberPointAlw,'N') ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  = '1')) "
'                    End If
'            End Select
            tSql = tSql & vbCrLf & " ORDER BY DT.FNSdtSeqNo ASC"
            If SP_SQLvExecute(tSql, orsDT) = 0 Then
                If Not (orsDT.BOF And orsDT.EOF) Then
                    Set porsTemp = New ADODB.Recordset
                    For nI = 1 To orsDT.Fields.Count
                        Select Case UCase(Mid(orsDT.Fields(nI - 1).Name, 2, 1))
                            Case "T"
                                Call porsTemp.Fields.Append(orsDT.Fields(nI - 1).Name, orsDT.Fields(nI - 1).Type, 256)
                            Case Else
                                Call porsTemp.Fields.Append(orsDT.Fields(nI - 1).Name, orsDT.Fields(nI - 1).Type, orsDT.Fields(nI - 1).DefinedSize)
                        End Select
                    Next nI
                    porsTemp.CursorType = adOpenStatic
                    porsTemp.Open
                    Do While Not orsDT.EOF
                        tSKUCode = tSKUCode & "''" & SP_FEDtChkString(orsDT, "FTSkuCode") & "''"
                        orsDT.MoveNext
                        If Not orsDT.EOF Then tSKUCode = tSKUCode & ","
                    Loop
                    If tSKUCode <> "" Then
'                        tSql = "SELECT SKU.FTSkuCode"
'                        tSql = tSql & vbCrLf & ",SKU.FTSkuStaOnTop"
'                        tSql = tSql & vbCrLf & ",SKU.FCSkuCoCpnRate,SKU.FCSkuCoBndRate"                '*Ao 51-06-19
'                        tSql = tSql & vbCrLf & ",SKU.FTSkuStaCoCpnAlw,SKU.FTSkuStaCoBndAlw,SKU.FTSkuStaMallCpn" '*TW 51-08-11
'                        tSql = tSql & vbCrLf & ",SKU.FDSkuCoCpnEndDate , SKU.FDSkuCoBndEndDate" '*KT 51-10-17 ให้ตรวจวันที่หมดอายุการใช้ Co.CpnและCo.Bnd ตาม SKU
'                        tSql = tSql & vbCrLf & ",SKU.FTSkuStaDisAlwTran,SKU.FDSkuDisAlwEndDate,SKU.FTSkuStaMrkAlwTran,SKU.FDSkuMrkAlwEndDate"
'                        tSql = tSql & vbCrLf & ",SKU.FTSkuStaMallEmpCard AS FTSkuStaMallEmpCardEx"
'                        tSql = tSql & vbCrLf & ",SKU.FCSkuMallCardRate"
'                        tSql = tSql & vbCrLf & ",SKU.FCSkuCoDisRate"
'                        tSql = tSql & vbCrLf & ",SKU.FTSkuEventCoDisCpn"
'                        tSql = tSql & vbCrLf & ",SKU.FTSkuStaCoDisAlw"
'                        tSql = tSql & vbCrLf & ",SKU.FDSkuCoDisEndDate"
'                        tSql = tSql & vbCrLf & ",SKU.FTSkuStaAutoTran"
'                        tSql = tSql & vbCrLf & " FROM   TCNMSku SKU"
'                        tSql = tSql & vbCrLf & " WHERE  SKU.FTSkuCode  IN (" & tSKUCode & ")"

                        '*KT 53-11-02 ปรับปรุง Performance
                        tFed = "SKU.FTSkuCode"
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaOnTop"
                        tFed = tFed & vbCrLf & ",SKU.FCSkuCoCpnRate,SKU.FCSkuCoBndRate"                '*Ao 51-06-19
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaCoCpnAlw,SKU.FTSkuStaCoBndAlw,SKU.FTSkuStaMallCpn" '*TW 51-08-11
                        tFed = tFed & vbCrLf & ",SKU.FDSkuCoCpnEndDate , SKU.FDSkuCoBndEndDate" '*KT 51-10-17 ให้ตรวจวันที่หมดอายุการใช้ Co.CpnและCo.Bnd ตาม SKU
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaDisAlwTran,SKU.FDSkuDisAlwEndDate,SKU.FTSkuStaMrkAlwTran,SKU.FDSkuMrkAlwEndDate"
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaMallEmpCard AS FTSkuStaMallEmpCardEx"
                        tFed = tFed & vbCrLf & ",SKU.FCSkuMallCardRate"
                        tFed = tFed & vbCrLf & ",SKU.FCSkuCoDisRate"
                        tFed = tFed & vbCrLf & ",SKU.FTSkuEventCoDisCpn"
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaCoDisAlw"
                        tFed = tFed & vbCrLf & ",SKU.FDSkuCoDisEndDate"
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaAutoTran"
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaMrkAlwTran1,SKU.FDSkuMrkAlwEndDate1"  '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaMrkAlwTran2,SKU.FDSkuMrkAlwEndDate2"
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaMrkAlwTran3,SKU.FDSkuMrkAlwEndDate3"
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaMrkAlwTran4,SKU.FDSkuMrkAlwEndDate4"
                        tFed = tFed & vbCrLf & ",SKU.FTSkuStaMrkAlwTran5,SKU.FDSkuMrkAlwEndDate5"
                        tWhr = " WHERE  SKU.FTSkuCode  IN (" & tSKUCode & ")"
                        
                        tSql = "STP_CNxSQLFunction @ptFunc = '" & tFed & "',@ptTable = '" & "TCNMSku SKU" & "',@ptWhere = '" & tWhr & "'"
                        If SP_SQLvExecute(tSql, orsSKU, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                            If (orsSKU.BOF And orsSKU.EOF) Then
                                If Not bVB_CNTrStaRegPreviewDay Then            '*Ao 2011/02/14    CommSheet ML-PRV-039-SF All Discount Promo  member price skupoint not working.xls
                                    Call SP_DATxRsNothing(porsTemp)                     'รหัสสินค้าใหม่ ที่ไม่มีใน Master มีแต่ใน Preview ต้องไปเช็คใน TCNMSkuPreDay
                                    Call SP_DATxRsNothing(orsDT)
                                    Call SP_DATxRsNothing(orsSKU)
                                    Call SP_DATxRsNothing(orsSKUPreDay)
                                    Exit Sub
                                End If
                            End If
                        Else
                            Call SP_DATxRsNothing(porsTemp)
                            Call SP_DATxRsNothing(orsDT)
                            Call SP_DATxRsNothing(orsSKU)
                            Call SP_DATxRsNothing(orsSKUPreDay)
                            Exit Sub
                        End If
                        '*KT 53-10-30 Preview Day
                        If bVB_CNTrStaRegPreviewDay Then
                            orsDT.Filter = "FTSdtStaNotAlwPrv = 'N'"
                            If Not (orsDT.BOF And orsDT.EOF) Then
                                bPass = True
                            Else
                                 bPass = False
                            End If
                            orsDT.Filter = adFilterNone
                            If bPass Then 'bPass
'                                tSql = "SELECT FTSkuCode"
'                                tSql = tSql & vbCrLf & ",SKU.FTSkuStaOnTop"
'                                tSql = tSql & vbCrLf & ",SKU.FCSkuCoCpnRate,SKU.FCSkuCoBndRate"                '*Ao 51-06-19
'                                tSql = tSql & vbCrLf & ",SKU.FTSkuStaCoCpnAlw,SKU.FTSkuStaCoBndAlw,SKU.FTSkuStaMallCpn" '*TW 51-08-11
'                                tSql = tSql & vbCrLf & ",SKU.FDSkuCoCpnEndDate , SKU.FDSkuCoBndEndDate" '*KT 51-10-17 ให้ตรวจวันที่หมดอายุการใช้ Co.CpnและCo.Bnd ตาม SKU
'                                tSql = tSql & vbCrLf & ",SKU.FTSkuStaDisAlwTran,SKU.FDSkuDisAlwEndDate,SKU.FTSkuStaMrkAlwTran,SKU.FDSkuMrkAlwEndDate"
'                                tSql = tSql & vbCrLf & ",SKU.FTSkuStaMallEmpCard AS FTSkuStaMallEmpCardEx"
'                                tSql = tSql & vbCrLf & ",SKU.FCSkuMallCardRate"
'                                tSql = tSql & vbCrLf & ",SKU.FCSkuCoDisRate"
'                                tSql = tSql & vbCrLf & ",SKU.FTSkuEventCoDisCpn"
'                                tSql = tSql & vbCrLf & ",SKU.FTSkuStaCoDisAlw"
'                                tSql = tSql & vbCrLf & ",SKU.FDSkuCoDisEndDate"
'                                tSql = tSql & vbCrLf & ",SKU.FTSkuStaAutoTran"
'                                tSql = tSql & vbCrLf & " FROM   TCNMSkuPreDay SKU"
'                                tSql = tSql & vbCrLf & " WHERE  SKU.FTSkuCode  IN (" & tSKUCode & ")"
                                '*KT 53-11-02 ปรับปรุง Performance
                                tSql = "STP_CNxSQLFunction @ptFunc = '" & tFed & "',@ptTable = '" & "TCNMSkuPreDay SKU" & "',@ptWhere = '" & tWhr & "'"
                                If SP_SQLvExecute(tSql, orsSKUPreDay, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                                    If (orsSKUPreDay.BOF And orsSKUPreDay.EOF) Then
                                        Call SP_DATxRsNothing(porsTemp)
                                        Call SP_DATxRsNothing(orsDT)
                                        Call SP_DATxRsNothing(orsSKU)
                                        Call SP_DATxRsNothing(orsSKUPreDay)
                                        Exit Sub
                                    End If
                                Else
                                    Call SP_DATxRsNothing(porsTemp)
                                    Call SP_DATxRsNothing(orsDT)
                                    Call SP_DATxRsNothing(orsSKU)
                                    Call SP_DATxRsNothing(orsSKUPreDay)
                                    Exit Sub
                                End If
                            End If 'bPass
                        End If
                    Else
                        Call SP_DATxRsNothing(porsTemp)
                        Call SP_DATxRsNothing(orsDT)
                        Call SP_DATxRsNothing(orsSKU)
                        Call SP_DATxRsNothing(orsSKUPreDay)
                        Exit Sub
                    End If
                    orsDT.MoveFirst
                    Do While Not orsDT.EOF
                        If bVB_CNTrStaRegPreviewDay And SP_FEDtChkString(orsDT, "FTSdtStaPreview") = "Y" And bPass Then
                            orsSKUPreDay.Filter = "FTSkuCode = '" & SP_FEDtChkString(orsDT, "FTSkuCode") & "'"
                            If Not (orsSKUPreDay.BOF And orsSKUPreDay.EOF) Then
                                porsTemp.AddNew
                                'FROM DT
                                porsTemp.Fields("FTTmnNum").Value = SP_FEDtChkString(orsDT, "FTTmnNum")
                                porsTemp.Fields("FTShdTransNo").Value = SP_FEDtChkString(orsDT, "FTShdTransNo")
                                porsTemp.Fields("FNSdtSeqNo").Value = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
                                porsTemp.Fields("FTSkuCode").Value = SP_FEDtChkString(orsDT, "FTSkuCode")
                                porsTemp.Fields("FTShdTransType").Value = SP_FEDtChkString(orsDT, "FTShdTransType")
                                porsTemp.Fields("FDShdTransDate").Value = Format(SP_FEDdChkDate(orsDT, "FDShdTransDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuDepCode").Value = SP_FEDtChkString(orsDT, "FTSkuDepCode")
                                porsTemp.Fields("FTSkuBrdCode").Value = SP_FEDtChkString(orsDT, "FTSkuBrdCode")
                                porsTemp.Fields("FCSdtSalePrice").Value = SP_FEDvChkNumeric(orsDT, "FCSdtSalePrice")
                                porsTemp.Fields("FCSdtB4DisChg").Value = SP_FEDvChkNumeric(orsDT, "FCSdtB4DisChg")
                                porsTemp.Fields("FCSdtSaleAmt").Value = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
                                porsTemp.Fields("FCSdtAmtDisAvi").Value = SP_FEDvChkNumeric(orsDT, "FCSdtAmtDisAvi")
                                porsTemp.Fields("FCSdtQtyAllDisAvi").Value = SP_FEDvChkNumeric(orsDT, "FCSdtQtyAllDisAvi")
                                porsTemp.Fields("FCSdtTax").Value = SP_FEDvChkNumeric(orsDT, "FCSdtTax")
                                porsTemp.Fields("FCSdtQtyAll").Value = SP_FEDvChkNumeric(orsDT, "FCSdtQtyAll")
                                porsTemp.Fields("FTSkuStaMallEmpCard").Value = SP_FEDtChkString(orsDT, "FTSkuStaMallEmpCard")
                                porsTemp.Fields("FCSdtQty").Value = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
                                porsTemp.Fields("FCSdtBSubPro").Value = SP_FEDvChkNumeric(orsDT, "FCSdtBSubPro")
                                porsTemp.Fields("FCSdtDiscSubPro").Value = SP_FEDvChkNumeric(orsDT, "FCSdtDiscSubPro")
                                porsTemp.Fields("FCSdtBSubReg").Value = SP_FEDvChkNumeric(orsDT, "FCSdtBSubReg")
                                porsTemp.Fields("FCSdtDiscSubReg").Value = SP_FEDvChkNumeric(orsDT, "FCSdtDiscSubReg")
                                porsTemp.Fields("FCSdtBluePrice").Value = SP_FEDvChkNumeric(orsDT, "FCSdtBluePrice")
                                porsTemp.Fields("FCSdtYellowPrice").Value = SP_FEDvChkNumeric(orsDT, "FCSdtYellowPrice")
                                porsTemp.Fields("FTSdtStaRfd").Value = SP_FEDtChkString(orsDT, "FTSdtStaRfd")
                                porsTemp.Fields("FCSdtMemberPrice").Value = SP_FEDvChkNumeric(orsDT, "FCSdtMemberPrice")
                                porsTemp.Fields("FTSdtStaPreview").Value = SP_FEDtChkString(orsDT, "FTSdtStaPreview") '*KT 53-10-30 Preview Day
                                porsTemp.Fields("FTSdtStaNotAlwPrv").Value = SP_FEDtChkString(orsDT, "FTSdtStaNotAlwPrv") '*KT 53-10-30 Preview Day
                                porsTemp.Fields("FCLineOver").Value = SP_FEDvChkNumeric(orsDT, "FCLineOver")   '*Eaw 56-03-15 Copy sale แบบ Line Over
                                porsTemp.Fields("FCSdtRegPrice").Value = SP_FEDvChkNumeric(orsDT, "FCSdtRegPrice")   '*Eaw 56-03-15    Copy sale แบบ Line Over
                                'FROM SKU
                                porsTemp.Fields("FTSkuStaOnTop").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaOnTop")
                                porsTemp.Fields("FCSkuCoCpnRate").Value = SP_FEDvChkNumeric(orsSKUPreDay, "FCSkuCoCpnRate")
                                porsTemp.Fields("FCSkuCoBndRate").Value = SP_FEDvChkNumeric(orsSKUPreDay, "FCSkuCoBndRate")
                                porsTemp.Fields("FTSkuStaCoCpnAlw").Value = SP_FEDvChkNumeric(orsSKUPreDay, "FTSkuStaCoCpnAlw")
                                porsTemp.Fields("FTSkuStaCoBndAlw").Value = SP_FEDvChkNumeric(orsSKUPreDay, "FTSkuStaCoBndAlw")
                                porsTemp.Fields("FTSkuStaMallCpn").Value = SP_FEDvChkNumeric(orsSKUPreDay, "FTSkuStaMallCpn")
                                porsTemp.Fields("FDSkuCoCpnEndDate").Value = Format(SP_FEDdChkDate(orsSKUPreDay, "FDSkuCoCpnEndDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FDSkuCoBndEndDate").Value = Format(SP_FEDdChkDate(orsSKUPreDay, "FDSkuCoBndEndDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaDisAlwTran").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaDisAlwTran")
                                porsTemp.Fields("FDSkuDisAlwEndDate").Value = Format(SP_FEDdChkDate(orsSKUPreDay, "FDSkuDisAlwEndDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMrkAlwTran").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaMrkAlwTran")
                                porsTemp.Fields("FDSkuMrkAlwEndDate").Value = Format(SP_FEDdChkDate(orsSKUPreDay, "FDSkuMrkAlwEndDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMallEmpCardEx").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaMallEmpCardEx")
                                porsTemp.Fields("FCSkuMallCardRate").Value = SP_FEDvChkNumeric(orsSKUPreDay, "FCSkuMallCardRate")
                                porsTemp.Fields("FCSkuCoDisRate").Value = SP_FEDvChkNumeric(orsSKUPreDay, "FCSkuCoDisRate")
                                porsTemp.Fields("FTSkuEventCoDisCpn").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuEventCoDisCpn")
                                porsTemp.Fields("FTSkuStaCoDisAlw").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaCoDisAlw")
                                porsTemp.Fields("FDSkuCoDisEndDate").Value = Format(SP_FEDdChkDate(orsSKUPreDay, "FDSkuCoDisEndDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaAutoTran").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaAutoTran")
                                porsTemp.Fields("FTSdtConsignFlag").Value = SP_FEDtChkString(orsDT, "FTSdtConsignFlag") '*KT 54-06-23 V.4.0.0 Ph.4.0 Req.9 Auoto Transaction Discount
                                porsTemp.Fields("FTSdtMemberPointAlw").Value = SP_FEDtChkString(orsDT, "FTSdtMemberPointAlw") '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
                                porsTemp.Fields("FTSdtStaPdt14").Value = SP_FEDtChkString(orsDT, "FTSdtStaPdt14") '*KT 54-11-14 V.4.0.4
                                porsTemp.Fields("FTSkuStaMrkAlwTran1").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaMrkAlwTran1")  '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                                porsTemp.Fields("FDSkuMrkAlwEndDate1").Value = Format(SP_FEDdChkDate(orsSKUPreDay, "FDSkuMrkAlwEndDate1"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMrkAlwTran2").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaMrkAlwTran2")
                                porsTemp.Fields("FDSkuMrkAlwEndDate2").Value = Format(SP_FEDdChkDate(orsSKUPreDay, "FDSkuMrkAlwEndDate2"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMrkAlwTran3").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaMrkAlwTran3")
                                porsTemp.Fields("FDSkuMrkAlwEndDate3").Value = Format(SP_FEDdChkDate(orsSKUPreDay, "FDSkuMrkAlwEndDate3"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMrkAlwTran4").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaMrkAlwTran4")
                                porsTemp.Fields("FDSkuMrkAlwEndDate4").Value = Format(SP_FEDdChkDate(orsSKUPreDay, "FDSkuMrkAlwEndDate4"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMrkAlwTran5").Value = SP_FEDtChkString(orsSKUPreDay, "FTSkuStaMrkAlwTran5")
                                porsTemp.Fields("FDSkuMrkAlwEndDate5").Value = Format(SP_FEDdChkDate(orsSKUPreDay, "FDSkuMrkAlwEndDate5"), "YYYY/MM/DD")
                                
                            Else
                                Call SP_DATxRsNothing(porsTemp)
                                Call SP_DATxRsNothing(orsDT)
                                Call SP_DATxRsNothing(orsSKU)
                                Call SP_DATxRsNothing(orsSKUPreDay)
                                Exit Sub
                            End If
                            orsSKUPreDay.Filter = adFilterNone
                        Else
                            orsSKU.Filter = "FTSkuCode = '" & SP_FEDtChkString(orsDT, "FTSkuCode") & "'"
                            If Not (orsSKU.BOF And orsSKU.EOF) Then
                                porsTemp.AddNew
                                'FROM DT
                                porsTemp.Fields("FTTmnNum").Value = SP_FEDtChkString(orsDT, "FTTmnNum")
                                porsTemp.Fields("FTShdTransNo").Value = SP_FEDtChkString(orsDT, "FTShdTransNo")
                                porsTemp.Fields("FNSdtSeqNo").Value = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
                                porsTemp.Fields("FTSkuCode").Value = SP_FEDtChkString(orsDT, "FTSkuCode")
                                porsTemp.Fields("FTShdTransType").Value = SP_FEDtChkString(orsDT, "FTShdTransType")
                                porsTemp.Fields("FDShdTransDate").Value = Format(SP_FEDdChkDate(orsDT, "FDShdTransDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuDepCode").Value = SP_FEDtChkString(orsDT, "FTSkuDepCode")
                                porsTemp.Fields("FTSkuBrdCode").Value = SP_FEDtChkString(orsDT, "FTSkuBrdCode")
                                porsTemp.Fields("FCSdtSalePrice").Value = SP_FEDvChkNumeric(orsDT, "FCSdtSalePrice")
                                porsTemp.Fields("FCSdtB4DisChg").Value = SP_FEDvChkNumeric(orsDT, "FCSdtB4DisChg")
                                porsTemp.Fields("FCSdtSaleAmt").Value = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
                                porsTemp.Fields("FCSdtAmtDisAvi").Value = SP_FEDvChkNumeric(orsDT, "FCSdtAmtDisAvi")
                                porsTemp.Fields("FCSdtQtyAllDisAvi").Value = SP_FEDvChkNumeric(orsDT, "FCSdtQtyAllDisAvi")
                                porsTemp.Fields("FCSdtTax").Value = SP_FEDvChkNumeric(orsDT, "FCSdtTax")
                                porsTemp.Fields("FCSdtQtyAll").Value = SP_FEDvChkNumeric(orsDT, "FCSdtQtyAll")
                                porsTemp.Fields("FTSkuStaMallEmpCard").Value = SP_FEDtChkString(orsDT, "FTSkuStaMallEmpCard")
                                porsTemp.Fields("FCSdtQty").Value = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
                                porsTemp.Fields("FCSdtBSubPro").Value = SP_FEDvChkNumeric(orsDT, "FCSdtBSubPro")
                                porsTemp.Fields("FCSdtDiscSubPro").Value = SP_FEDvChkNumeric(orsDT, "FCSdtDiscSubPro")
                                porsTemp.Fields("FCSdtBSubReg").Value = SP_FEDvChkNumeric(orsDT, "FCSdtBSubReg")
                                porsTemp.Fields("FCSdtDiscSubReg").Value = SP_FEDvChkNumeric(orsDT, "FCSdtDiscSubReg")
                                porsTemp.Fields("FCSdtBluePrice").Value = SP_FEDvChkNumeric(orsDT, "FCSdtBluePrice")
                                porsTemp.Fields("FCSdtYellowPrice").Value = SP_FEDvChkNumeric(orsDT, "FCSdtYellowPrice")
                                porsTemp.Fields("FTSdtStaRfd").Value = SP_FEDtChkString(orsDT, "FTSdtStaRfd")
                                porsTemp.Fields("FCSdtMemberPrice").Value = SP_FEDvChkNumeric(orsDT, "FCSdtMemberPrice")
                                porsTemp.Fields("FTSdtStaPreview").Value = SP_FEDtChkString(orsDT, "FTSdtStaPreview") '*KT 53-10-30 Preview Day
                                porsTemp.Fields("FTSdtStaNotAlwPrv").Value = SP_FEDtChkString(orsDT, "FTSdtStaNotAlwPrv") '*KT 53-10-30 Preview Day
                                porsTemp.Fields("FCLineOver").Value = SP_FEDvChkNumeric(orsDT, "FCLineOver")   '*Eaw 56-03-15    Copy sale แบบ Line Over
                                porsTemp.Fields("FCSdtRegPrice").Value = SP_FEDvChkNumeric(orsDT, "FCSdtRegPrice")   '*Eaw 56-03-15   Copy sale แบบ Line Over
                                'FROM SKU
                                porsTemp.Fields("FTSkuStaOnTop").Value = SP_FEDtChkString(orsSKU, "FTSkuStaOnTop")
                                porsTemp.Fields("FCSkuCoCpnRate").Value = SP_FEDvChkNumeric(orsSKU, "FCSkuCoCpnRate")
                                porsTemp.Fields("FCSkuCoBndRate").Value = SP_FEDvChkNumeric(orsSKU, "FCSkuCoBndRate")
                                porsTemp.Fields("FTSkuStaCoCpnAlw").Value = SP_FEDvChkNumeric(orsSKU, "FTSkuStaCoCpnAlw")
                                porsTemp.Fields("FTSkuStaCoBndAlw").Value = SP_FEDvChkNumeric(orsSKU, "FTSkuStaCoBndAlw")
                                porsTemp.Fields("FTSkuStaMallCpn").Value = SP_FEDvChkNumeric(orsSKU, "FTSkuStaMallCpn")
                                porsTemp.Fields("FDSkuCoCpnEndDate").Value = Format(SP_FEDdChkDate(orsSKU, "FDSkuCoCpnEndDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FDSkuCoBndEndDate").Value = Format(SP_FEDdChkDate(orsSKU, "FDSkuCoBndEndDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaDisAlwTran").Value = SP_FEDtChkString(orsSKU, "FTSkuStaDisAlwTran")
                                porsTemp.Fields("FDSkuDisAlwEndDate").Value = Format(SP_FEDdChkDate(orsSKU, "FDSkuDisAlwEndDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMrkAlwTran").Value = SP_FEDtChkString(orsSKU, "FTSkuStaMrkAlwTran")
                                porsTemp.Fields("FDSkuMrkAlwEndDate").Value = Format(SP_FEDdChkDate(orsSKU, "FDSkuMrkAlwEndDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMallEmpCardEx").Value = SP_FEDtChkString(orsSKU, "FTSkuStaMallEmpCardEx")
                                porsTemp.Fields("FCSkuMallCardRate").Value = SP_FEDvChkNumeric(orsSKU, "FCSkuMallCardRate")
                                porsTemp.Fields("FCSkuCoDisRate").Value = SP_FEDvChkNumeric(orsSKU, "FCSkuCoDisRate")
                                porsTemp.Fields("FTSkuEventCoDisCpn").Value = SP_FEDtChkString(orsSKU, "FTSkuEventCoDisCpn")
                                porsTemp.Fields("FTSkuStaCoDisAlw").Value = SP_FEDtChkString(orsSKU, "FTSkuStaCoDisAlw")
                                porsTemp.Fields("FDSkuCoDisEndDate").Value = Format(SP_FEDdChkDate(orsSKU, "FDSkuCoDisEndDate"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaAutoTran").Value = SP_FEDtChkString(orsSKU, "FTSkuStaAutoTran")
                                porsTemp.Fields("FTSdtConsignFlag").Value = SP_FEDtChkString(orsDT, "FTSdtConsignFlag") '*KT 54-06-23 V.4.0.0 Ph.4.0 Req.9 Auoto Transaction Discount
                                porsTemp.Fields("FTSdtMemberPointAlw").Value = SP_FEDtChkString(orsDT, "FTSdtMemberPointAlw") '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
                                porsTemp.Fields("FTSdtStaPdt14").Value = SP_FEDtChkString(orsDT, "FTSdtStaPdt14") '*KT 54-11-14 V.4.0.4
                                porsTemp.Fields("FTSkuStaMrkAlwTran1").Value = SP_FEDtChkString(orsSKU, "FTSkuStaMrkAlwTran1")  '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                                porsTemp.Fields("FDSkuMrkAlwEndDate1").Value = Format(SP_FEDdChkDate(orsSKU, "FDSkuMrkAlwEndDate1"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMrkAlwTran2").Value = SP_FEDtChkString(orsSKU, "FTSkuStaMrkAlwTran2")
                                porsTemp.Fields("FDSkuMrkAlwEndDate2").Value = Format(SP_FEDdChkDate(orsSKU, "FDSkuMrkAlwEndDate2"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMrkAlwTran3").Value = SP_FEDtChkString(orsSKU, "FTSkuStaMrkAlwTran3")
                                porsTemp.Fields("FDSkuMrkAlwEndDate3").Value = Format(SP_FEDdChkDate(orsSKU, "FDSkuMrkAlwEndDate3"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMrkAlwTran4").Value = SP_FEDtChkString(orsSKU, "FTSkuStaMrkAlwTran4")
                                porsTemp.Fields("FDSkuMrkAlwEndDate4").Value = Format(SP_FEDdChkDate(orsSKU, "FDSkuMrkAlwEndDate4"), "YYYY/MM/DD")
                                porsTemp.Fields("FTSkuStaMrkAlwTran5").Value = SP_FEDtChkString(orsSKU, "FTSkuStaMrkAlwTran5")
                                porsTemp.Fields("FDSkuMrkAlwEndDate5").Value = Format(SP_FEDdChkDate(orsSKU, "FDSkuMrkAlwEndDate5"), "YYYY/MM/DD")
                                
                            Else
                                Call SP_DATxRsNothing(porsTemp)
                                Call SP_DATxRsNothing(orsDT)
                                Call SP_DATxRsNothing(orsSKU)
                                Call SP_DATxRsNothing(orsSKUPreDay)
                                Exit Sub
                            End If
                            orsSKU.Filter = adFilterNone
                        End If
                        
                        orsDT.MoveNext
                    Loop
                End If
            Else
                Call SP_DATxRsNothing(porsTemp)
                Call SP_DATxRsNothing(orsDT)
                Call SP_DATxRsNothing(orsSKU)
                Call SP_DATxRsNothing(orsSKUPreDay)
                Exit Sub
            End If
    Else
            tSql = "SELECT DT.FTTmnNum ,DT.FTShdTransNo ,DT.FNSdtSeqNo ,DT.FTSkuCode ,DT.FTShdTransType"
            tSql = tSql & vbCrLf & ",DT.FDShdTransDate ,DT.FTSkuDepCode ,DT.FTSkuBrdCode ,DT.FCSdtB4DisChg"
            tSql = tSql & vbCrLf & ",DT.FCLineOver,DT.FCSdtRegPrice"   '*Eaw 56-06-17  CommSheet TK-ISS3000-234-ST.xlsx
            '*KT 52-11-25 Deal&Promo
            If bC_PromoAftDeal Then
                Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnDutyFreeSale
                        tSql = tSql & vbCrLf & " ,ROUND(FCSdtB4DisChg,2) AS FCSdtSalePrice"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg) * FCSdtQty,2) AS FCSdtSaleAmt"
                        tSql = tSql & vbCrLf & " ,FCSdtQty AS FCSdtQtyAllDisAvi"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg) * FCSdtQty,2) AS FCSdtAmtDisAvi"
                    Case tEN_TrnTaxExceptSale
                        tSql = tSql & vbCrLf & " ,ROUND(FCSdtB4DisChg * 100 / (100 + FCSdtTax),2) AS FCSdtSalePrice"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg * 100 / (100 + FCSdtTax)),2) * FCSdtQty AS FCSdtSaleAmt"
                        tSql = tSql & vbCrLf & " ,ROUND(FCSdtQty,2) AS FCSdtQtyAllDisAvi"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg * 100 / (100 + FCSdtTax)),2) * FCSdtQty AS FCSdtAmtDisAvi"
                    Case Else
                        tSql = tSql & vbCrLf & " ,ROUND(FCSdtB4DisChg,2) AS FCSdtSalePrice"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg) * FCSdtQty,2) AS FCSdtSaleAmt"
                        tSql = tSql & vbCrLf & " ,ROUND(FCSdtQty,2) AS FCSdtQtyAllDisAvi"
                        tSql = tSql & vbCrLf & " ,ROUND((FCSdtB4DisChg) * FCSdtQty,2) AS FCSdtAmtDisAvi"
                End Select
            Else
                tSql = tSql & vbCrLf & ",DT.FCSdtSalePrice "
                tSql = tSql & vbCrLf & ",DT.FCSdtSaleAmt "
                tSql = tSql & vbCrLf & ",DT.FCSdtAmtDisAvi "
                tSql = tSql & vbCrLf & ",DT.FCSdtQtyAllDisAvi"
            End If
            tSql = tSql & vbCrLf & ",DT.FCSdtTax,DT.FCSdtQtyAll"
             '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
            Select Case nC_ENDiscType
                Case nEN_Disc11MallCrd
                    tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 'N' ELSE 'Y'  END AS FTSkuStaMallEmpCard"
                Case nEN_Disc12EmpCrd           '*Ao 2013/03/08  TK
                    tSql = tSql & vbCrLf & ",CASE WHEN SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),20,1) = '1' THEN 'Y' ELSE 'N' END AS FTSkuStaMallEmpCard"
                Case Else
                    tSql = tSql & vbCrLf & ",DT.FTSkuStaMallEmpCard"
            End Select
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaOnTop"
            tSql = tSql & vbCrLf & ",DT.FCSdtQty"                '*Ao 51-06-19
            tSql = tSql & vbCrLf & ",SKU.FCSkuCoCpnRate,SKU.FCSkuCoBndRate"                '*Ao 51-06-19
             '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
'            tSql = tSql & vbCrLf & ",DT.FCSdtBSubPro,DT.FCSdtDiscSubPro,DT.FCSdtBSubReg,DT.FCSdtDiscSubReg" '*Ao 51-07-31
'            tSql = tSql & vbCrLf & ",DT.FCSdtBluePrice,DT.FCSdtYellowPrice" '*TW 51-08-07
            Select Case nC_ENDiscType
                Case nEN_Disc11MallCrd
                    If tC_MallAlwPoint = "Y" Then
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtBSubReg ELSE DT.FCSdtBSubPro  END AS FCSdtBSubPro"
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtDiscSubReg ELSE DT.FCSdtDiscSubPro  END AS FCSdtDiscSubPro"
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtBSubPro ELSE DT.FCSdtBSubReg  END  AS FCSdtBSubReg"
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtDiscSubPro ELSE DT.FCSdtDiscSubReg  END AS FCSdtDiscSubReg" '*Ao 51-07-31
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtYellowPrice ELSE DT.FCSdtBluePrice  END AS FCSdtBluePrice"
'                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN DT.FCSdtBluePrice ELSE DT.FCSdtYellowPrice  END AS FCSdtYellowPrice" '*TW 51-08-07
                    
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtBSubPro  END AS FCSdtBSubPro"
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtDiscSubPro  END AS FCSdtDiscSubPro"
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtBSubReg  END  AS FCSdtBSubReg"
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtDiscSubReg  END AS FCSdtDiscSubReg" '*Ao 51-07-31
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtBluePrice  END AS FCSdtBluePrice"
                        tSql = tSql & vbCrLf & ",CASE WHEN FTSdtMemberPointAlw ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  <> '1' THEN 0 ELSE DT.FCSdtYellowPrice  END AS FCSdtYellowPrice" '*TW 51-08-07
                    Else
                        tSql = tSql & vbCrLf & ",DT.FCSdtBSubPro"
                        tSql = tSql & vbCrLf & ",DT.FCSdtDiscSubPro"
                        tSql = tSql & vbCrLf & ",DT.FCSdtBSubReg"
                        tSql = tSql & vbCrLf & ",DT.FCSdtDiscSubReg"
                        tSql = tSql & vbCrLf & ",DT.FCSdtBluePrice"
                        tSql = tSql & vbCrLf & ",DT.FCSdtYellowPrice"
                    End If
                Case Else
                    tSql = tSql & vbCrLf & ",DT.FCSdtBSubPro"
                    tSql = tSql & vbCrLf & ",DT.FCSdtDiscSubPro"
                    tSql = tSql & vbCrLf & ",DT.FCSdtBSubReg"
                    tSql = tSql & vbCrLf & ",DT.FCSdtDiscSubReg" '*Ao 51-07-31
                    tSql = tSql & vbCrLf & ",DT.FCSdtBluePrice"
                    tSql = tSql & vbCrLf & ",DT.FCSdtYellowPrice" '*TW 51-08-07
            End Select
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaCoCpnAlw,SKU.FTSkuStaCoBndAlw,SKU.FTSkuStaMallCpn" '*TW 51-08-11
            tSql = tSql & vbCrLf & ",SKU.FDSkuCoCpnEndDate , SKU.FDSkuCoBndEndDate" '*KT 51-10-17 ให้ตรวจวันที่หมดอายุการใช้ Co.CpnและCo.Bnd ตาม SKU
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaDisAlwTran,SKU.FDSkuDisAlwEndDate,SKU.FTSkuStaMrkAlwTran,SKU.FDSkuMrkAlwEndDate"
            tSql = tSql & vbCrLf & ",DT.FTSdtStaRfd,DT.FCSdtMemberPrice" '*KT 52-02-13 Member price
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaMallEmpCard AS FTSkuStaMallEmpCardEx"
            tSql = tSql & vbCrLf & ",SKU.FCSkuMallCardRate"
            tSql = tSql & vbCrLf & ",SKU.FCSkuCoDisRate"
            tSql = tSql & vbCrLf & ",SKU.FTSkuEventCoDisCpn"
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaCoDisAlw"
            tSql = tSql & vbCrLf & ",SKU.FDSkuCoDisEndDate"
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaAutoTran"
            tSql = tSql & vbCrLf & ",DT.FTSdtEventMemPrice" '*KT 52-09-08 Member Price + Consignment
            tSql = tSql & vbCrLf & ",DT.FTSdtStaPreview" '*KT 53-10-30 Preview Day
            tSql = tSql & vbCrLf & ",DT.FTSdtStaNotAlwPrv" '*KT 53-10-30 Preview Day
            tSql = tSql & vbCrLf & ",DT.FTSdtMemberPointAlw" '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
            tSql = tSql & vbCrLf & ",SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1) AS FTSdtStaPdt14" '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
            tSql = tSql & vbCrLf & ",DT.FTSdtConsignFlag" '*KT 54-06-23 V.4.0.0 Ph.4.0 Req.9 Auoto Transaction Discount
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaMrkAlwTran1,SKU.FDSkuMrkAlwEndDate1"  '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaMrkAlwTran2,SKU.FDSkuMrkAlwEndDate2"
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaMrkAlwTran3,SKU.FDSkuMrkAlwEndDate3"
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaMrkAlwTran4,SKU.FDSkuMrkAlwEndDate4"
            tSql = tSql & vbCrLf & ",SKU.FTSkuStaMrkAlwTran5,SKU.FDSkuMrkAlwEndDate5"
            tSql = tSql & vbCrLf & " FROM    " & tVB_CNTblDT & " DT WITH(NOLOCK)  INNER JOIN "
            tSql = tSql & vbCrLf & tVB_CNTblSku & "  SKU WITH(NOLOCK)  ON  DT.FTSkuCode = SKU.FTSkuCode"
            tSql = tSql & vbCrLf & " WHERE DT.FTTmnNum = '" & tVB_CNTerminalNum & "'"
            tSql = tSql & vbCrLf & " AND DT.FTShdTransNo = '" & tVB_Running & "'"
            tSql = tSql & vbCrLf & " AND DT.FDShdTransDate = '" & tUT_SaleDate & "'"
            tSql = tSql & vbCrLf & " AND ISNULL(DT.FTSdtCancelFlag,'0') <> '1'"
            '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
'            Select Case nC_ENDiscType
'                Case nEN_Disc11MallCrd
'                    If tC_MallAlwPoint = "Y" Then
'                        tSql = tSql & vbCrLf & " AND (ISNULL(DT.FTSdtMemberPointAlw,'N') ='N'  OR (ISNULL(DT.FTSdtMemberPointAlw,'N') ='Y' AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),14,1)  = '1')) "
'                    End If
'            End Select
            tSql = tSql & vbCrLf & " ORDER BY DT.FNSdtSeqNo ASC"
            If Not SP_SQLbExecute(tSql, porsTemp) Then
                Call SP_DATxRsNothing(porsTemp)
                Call SP_DATxRsNothing(orsDT)
                Call SP_DATxRsNothing(orsSKU)
                Call SP_DATxRsNothing(orsSKUPreDay)
                Exit Sub
            End If
    End If
    Exit Sub
Err_Handler:
    Call SP_DATxRsNothing(porsTemp)
    Call SP_DATxRsNothing(orsDT)
    Call SP_DATxRsNothing(orsSKU)
    Call SP_DATxRsNothing(orsSKUPreDay)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub C_DATxInsertCDDeal(ByVal pcDisAmt#, ByVal pnFNSdtSeqNo%, ByVal ptEventID$, ByVal pcDisQty#, ByVal pcAlwDis#, ByVal pcAlwDisQty#)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # ใส่ข้อมูล Deal ลง CD
'ARGUMENT LIST # pcDisAmt =
'                                         pnFNSdtSeqNo =
'                                         ptEventID =
'                                         pcDisQty =
'                                         pcAlwDis =
'                                         pcAlwDisQty =
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim tField As String
    Dim tValue As String
    Dim nMaxSeqNo As Integer
    Dim tWhe As String
    Dim nEffect As Long
    Dim TWheMax As String
    Dim cDisPmt As Double
    On Error GoTo Err_Handler
    nC_DealGroup = nDE_Group
    tField = "FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FNScdSeqNo"
    tField = tField & ",FTSkuCode,FTShdTransType,FDShdTransDate,FTSdtBarCode,FTSdtStkCode"
    tField = tField & ",FCSdtStkFac,FTSkuDepCode,FTSkuBrdCode,FTPunCode,FCScdQtyB4"
    tField = tField & ",FCScdQty,FCScdQtyAF,FCScdAmtB4,FCScdAmtB4Avi,FCScdAmt"
    tField = tField & ",FCScdAmtAF,FCScdAmtAFAvi,FCScdDiscPer,FCScdDiscAmt,FCScdDiscNew"
    tField = tField & ",FTScdDiscType,FNDctNo,FTDctName,FTDctRefCode,FTScdEventID"
    tField = tField & ",FTScdCardID,FDScdExpDate,FTScdCardType,FTScdStaTrans"
    tField = tField & ",FTScdCstID,FTScdCardKey,FTScdStaChkExpire,FTScdStaChkDepart,FCScdMaxAlwDisc"
    tField = tField & ",FNScdActSeqNo"
    tField = tField & ",FNScdOrder,FCScdAmtTime,FCScdAmtPmt,FNScdDealGroup"
    tField = tField & ",FCScdDiscPerOrg,FCScdDiscAmtOrg,FCScdDiscNewOrg"
    tField = tField & ",FTScdStaPreview" '*KT 53-11-02 Preview Day
    tWhe = " WHERE FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " And  FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " And FDShdTransDate = '" & tUT_SaleDate & "'"
    TWheMax = tWhe
    tWhe = tWhe & " And FNSdtSeqNo = " & pnFNSdtSeqNo
    
    nMaxSeqNo = SP_VALcText2Double(SP_SQLtFunction("MAX", "FNScdSeqNo", "" & tVB_CNTblCD & "", TWheMax))
    nMaxSeqNo = nMaxSeqNo + 1
    
    tSql = "INSERT INTO " & tVB_CNTblCD & " WITH(ROWLOCK) (" & tField & SP_GETtWhoInsertField
    tSql = tSql & vbCrLf & " SELECT "
    tSql = tSql & vbCrLf & "FTStmCode , FTTmnNum, FTShdTransNo, FNSdtSeqNo"
    tSql = tSql & vbCrLf & "," & nMaxSeqNo & " AS FNScdSeqNo "
    tSql = tSql & vbCrLf & ",FTSkuCode,FTShdTransType,FDShdTransDate,FTSdtBarCode,FTSdtStkCode"
    tSql = tSql & vbCrLf & ",FCSdtStkFac,FTSkuDepCode,FTSkuBrdCode,FTPunCode"
    If pcDisAmt > 0 Then
        tSql = tSql & vbCrLf & "," & "FCSdtQtyAll" & " AS FCScdQtyB4 "
    Else
        tSql = tSql & vbCrLf & "," & "0 AS FCScdQtyB4 "
    End If
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(pcDisQty, True, False) & " AS FCScdQty"
    If pcDisAmt > 0 Then
        tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtQtyAll - " & SP_FMTcCurPoint(pcAlwDisQty, True, False) & " < 0 THEN 0 ELSE FCSdtQtyAll - " & SP_FMTcCurPoint(pcAlwDisQty, True, False) & " END AS FCScdQtyAF"
        tSql = tSql & vbCrLf & "," & "FCSdtSaleAmt" & " AS FCScdAmtB4"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(pcAlwDis, True, False) & " AS FCScdAmtB4Avi"
    Else
        tSql = tSql & vbCrLf & "," & "0 AS FCScdQtyAF"
        tSql = tSql & vbCrLf & "," & "0" & " AS FCScdAmtB4"
        tSql = tSql & vbCrLf & "," & "0" & " AS FCScdAmtB4Avi"
    End If
    
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(pcDisAmt, True, False) & " AS FCScdAmt"
    tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtSaleAmt - " & SP_FMTcCurPoint(pcDisAmt, True, False) & " < 0 THEN 0 ELSE FCSdtSaleAmt - " & SP_FMTcCurPoint(pcDisAmt, True, False) & " END AS FCScdAmtAF"
    tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtAmtDisAvi - " & SP_FMTcCurPoint(pcAlwDis, True, False) & " < 0 THEN 0 ELSE FCSdtAmtDisAvi -" & SP_FMTcCurPoint(pcAlwDis, True, False) & " END AS FCScdAmtAFAvi"
    tSql = tSql & vbCrLf & "," & cC_DiscPer & " AS FCScdDiscPer"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscAmt, True, False) & " AS FCScdDiscAmt"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscNew, True, False) & " AS FCScdDiscNew"
    tSql = tSql & vbCrLf & "," & nC_ENLnDisType & " AS FTScdDiscType"
    tSql = tSql & vbCrLf & "," & nC_ENDiscType & " AS FNDctNo"
    tSql = tSql & vbCrLf & ",'" & SP_SQLtFunction("", "FTDctDiscType", "TCNMDiscountType", "WHERE FNDctNo = " & nC_ENDiscType, oDB.ocnOnLine, True, True) & "' AS FTDctName" 'read master
    tSql = tSql & vbCrLf & ",'" & tC_RefCode & "' AS FTDctRefCode"
    tSql = tSql & vbCrLf & ",'" & ptEventID & "' AS FTScdEventID"
    tSql = tSql & vbCrLf & ",'" & tC_CardID & "' AS FTScdCardID"
    tSql = tSql & vbCrLf & ",'" & SP_DTEtStrSTD(dC_CardExpDate) & "' AS FDScdExpDate"
    tSql = tSql & vbCrLf & ",'" & Left(tC_CardType, 3) & "' AS FTScdCardType"
    tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
    tSql = tSql & vbCrLf & ",'" & tC_CusID & "' AS FTScdCstID"
    tSql = tSql & vbCrLf & ",'" & Left(tC_CardType, 3) & "' AS FTScdCardKey"
    tSql = tSql & vbCrLf & ",'" & tC_StaChkExpire & "' AS FTScdStaChkExpire"
    tSql = tSql & vbCrLf & ",'" & tC_StaChkDepart & "' AS FTScdStaChkDepart"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_AlwDis, True, False) & " AS FCScdMaxAlwDisc"
    tSql = tSql & vbCrLf & "," & nC_ActSeqNo & " AS FNScdActSeqNo"
    tSql = tSql & vbCrLf & "," & C_GETnGroup & " AS FNScdOrder"
    cDisPmt = 0
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cDisPmt, True, False) & " AS FCScdAmtTime"
    cDisPmt = 0
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cDisPmt, True, False) & " AS FCScdAmtPmt"
    tSql = tSql & vbCrLf & "," & nC_DealGroup & " AS FNScdDealGroup"
    tSql = tSql & vbCrLf & "," & cC_DiscPerOrg & " AS FCScdDiscPerOrg"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscAmtOrg, True, False) & " AS FCScdDiscAmtOrg"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscNewOrg, True, False) & " AS FCScdDiscNewOrg"
    tSql = tSql & vbCrLf & ",'" & IIf(bVB_CNModePreviewDay, "Y", "N") & "' AS FTScdStaPreview" '*KT 53-11-01 Preview Day
    tSql = tSql & vbCrLf & Left(SP_GETtWhoInsertValue, Len(SP_GETtWhoInsertValue) - 1)
    tSql = tSql & vbCrLf & " FROM " & tVB_CNTblDT & ""
    tSql = tSql & vbCrLf & tWhe
    If SP_SQLbExecute(tSql, , False, nEffect) Then
        If nEffect <= 0 Then
            Call SP_MSGnShowing(tC_Msg003, tCS_CNMsgWarn)
        End If
    Else
        Call SP_MSGnShowing(tC_Msg003, tCS_CNMsgWarn)
    End If
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Public Sub C_DATxInsertDiscCD(ByVal pcDisAmt#, ByVal pnFNSdtSeqNo%, ByVal ptEventID$, ByVal pcDisQty#, ByVal pcAlwDis#, ByVal pcAlwDisQty#, Optional ByVal pcDiscReg As Double = 0, Optional ByVal pcDiscPro As Double = 0, Optional ByVal pcMallCrdRate As Double = 0, Optional ByRef pbFlagInsCD As Boolean = False)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตั้งต้น 3.0.19
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim tField As String
    Dim tValue As String
    Dim nMaxSeqNo As Integer
    Dim tWhe As String
    Dim nEffect As Long
    Dim TWheMax As String
    Dim cDisPmt As Double
    Dim nErr As Long
    On Error GoTo Err_Handler
    
    tField = "FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FNScdSeqNo"
    tField = tField & ",FTSkuCode,FTShdTransType,FDShdTransDate,FTSdtBarCode,FTSdtStkCode"
    tField = tField & ",FCSdtStkFac,FTSkuDepCode,FTSkuBrdCode,FTPunCode,FCScdQtyB4"
    tField = tField & ",FCScdQty,FCScdQtyAF,FCScdAmtB4,FCScdAmtB4Avi,FCScdAmt"
    tField = tField & ",FCScdAmtAF,FCScdAmtAFAvi,FCScdDiscPer,FCScdDiscAmt,FCScdDiscNew"
    tField = tField & ",FTScdDiscType,FNDctNo,FTDctName,FTDctRefCode,FTScdEventID"
    tField = tField & ",FTScdCardID,FDScdExpDate,FTScdCardType,FTScdStaTrans"
    tField = tField & ",FTScdCstID,FTScdCardKey,FTScdStaChkExpire,FTScdStaChkDepart,FCScdMaxAlwDisc"
    tField = tField & ",FNScdActSeqNo"
    tField = tField & ",FNScdOrder,FCScdAmtTime,FCScdAmtPmt,FNScdDealGroup"
    tField = tField & ",FCScdDiscPerOrg,FCScdDiscAmtOrg,FCScdDiscNewOrg"
    tField = tField & ",FCScdDiscSubPro, FCScdDiscSubReg"
    tField = tField & ",FCSkuMallCardRate" '*KT 51-09-25 เก็บ Rate ของ Mall Card ที่ตั้งค่าที่ SKU
    tField = tField & ",FTScdCardName"  '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
    tField = tField & ",FTScdStaPreview" '*KT 53-11-02 PreviewDay
    tField = tField & ",FTRemark"               '*Ao 2013/03/06     TK Emp EDC
    
    tWhe = " WHERE FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " And  FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " And FDShdTransDate = '" & tUT_SaleDate & "'"
    TWheMax = tWhe
    tWhe = tWhe & " And FNSdtSeqNo = " & pnFNSdtSeqNo
    
    If bC_RecalByItem Then
        nMaxSeqNo = nC_ScdSeqNo
    Else
        nMaxSeqNo = SP_VALcText2Double(SP_SQLtFunction("MAX", "FNScdSeqNo", "" & tVB_CNTblCD & "", TWheMax))
        nMaxSeqNo = nMaxSeqNo + 1
    End If
    
    If Not (bW_ChkMallCard) = True Then
        tSql = "INSERT INTO " & tVB_CNTblCD & " WITH(ROWLOCK) (" & tField & SP_GETtWhoInsertField
        tSql = tSql & vbCrLf & " SELECT "
        tSql = tSql & vbCrLf & "FTStmCode , FTTmnNum, FTShdTransNo, FNSdtSeqNo"
        tSql = tSql & vbCrLf & "," & nMaxSeqNo & " AS FNScdSeqNo "
        tSql = tSql & vbCrLf & ",FTSkuCode,FTShdTransType,FDShdTransDate,FTSdtBarCode"
        '*KT 52-11-25 Deal&Promo
        If bC_PromoAftDeal Then
            tSql = tSql & vbCrLf & ",'1' AS FTSdtStkCode "
            tSql = tSql & vbCrLf & ",FCSdtStkFac,FTSkuDepCode,FTSkuBrdCode,FTPunCode"
            tSql = tSql & vbCrLf & "," & pcDisQty & " AS FCScdQtyB4 "
            tSql = tSql & vbCrLf & "," & pcDisQty & " AS FCScdQty"
            tSql = tSql & vbCrLf & "," & "CASE WHEN " & pcDisQty - pcAlwDisQty & " < 0 THEN 0 ELSE " & pcDisQty - pcAlwDisQty & " END AS FCScdQtyAF"
            tSql = tSql & vbCrLf & "," & pcAlwDis & " AS FCScdAmtB4"
            tSql = tSql & vbCrLf & "," & pcAlwDis & " AS FCScdAmtB4Avi"
            tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(pcDisAmt, True, False) & " AS FCScdAmt"
            tSql = tSql & vbCrLf & "," & "CASE WHEN " & pcAlwDis - SP_FMTcCurPoint(pcDisAmt, True, False) & " < 0 THEN 0 ELSE " & pcAlwDis - SP_FMTcCurPoint(pcDisAmt, True, False) & " END AS FCScdAmtAF"
            tSql = tSql & vbCrLf & "," & "CASE WHEN " & pcAlwDis - SP_FMTcCurPoint(pcAlwDis, True, False) & " <0 THEN 0 ELSE " & pcAlwDis - SP_FMTcCurPoint(pcAlwDis, True, False) & " END AS FCScdAmtAFAvi"
        Else
            tSql = tSql & vbCrLf & ",'' AS FTSdtStkCode "
            tSql = tSql & vbCrLf & ",FCSdtStkFac,FTSkuDepCode,FTSkuBrdCode,FTPunCode"
            tSql = tSql & vbCrLf & "," & "FCSdtQtyAllDisAvi" & " AS FCScdQtyB4 "
            tSql = tSql & vbCrLf & "," & pcDisQty & " AS FCScdQty"
            tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtQtyAllDisAvi - " & pcAlwDisQty & " < 0 THEN 0 ELSE FCSdtQtyAllDisAvi - " & pcAlwDisQty & " END AS FCScdQtyAF"
            tSql = tSql & vbCrLf & "," & "FCSdtSaleAmt" & " AS FCScdAmtB4"
            tSql = tSql & vbCrLf & "," & "FCSdtAmtDisAvi" & " AS FCScdAmtB4Avi"
            tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(pcDisAmt, True, False) & " AS FCScdAmt"
            tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtSaleAmt - " & SP_FMTcCurPoint(pcDisAmt, True, False) & " < 0 THEN 0 ELSE FCSdtSaleAmt - " & SP_FMTcCurPoint(pcDisAmt, True, False) & " END AS FCScdAmtAF"
            tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtAmtDisAvi -" & SP_FMTcCurPoint(pcAlwDis, True, False) & " <0 THEN 0 ELSE FCSdtAmtDisAvi -" & SP_FMTcCurPoint(pcAlwDis, True, False) & " END AS FCScdAmtAFAvi"
        End If
        
        Select Case nC_ENDiscType
            Case nEN_Disc09CoopCpn, nEN_Disc10CoopBrand '*KT 53-02-09 Communication Sheet PH3-ML-SF-083 Record Corperate Coupon rate wrong when SKU allow lower % than Corporate Coupon % input.xls
                tSql = tSql & vbCrLf & "," & cC_DiscPer & " AS FCScdDiscPer"
            Case Else
                tSql = tSql & vbCrLf & "," & cC_DiscPerOrg & " AS FCScdDiscPer" '*TW 51-06-05 แก้จาก cC_DiscPer เป็น cC_DiscPerOrg
        End Select
        
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscAmtOrg, True, False) & " AS FCScdDiscAmt" '*TW 51-06-05 แก้จาก cC_DiscAmt เป็น cC_DiscAmtOrg
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscNewOrg, True, False) & " AS FCScdDiscNew" '*TW 51-06-05 แก้จาก cC_DiscNew เป็น cC_DiscNewOrg
        tSql = tSql & vbCrLf & "," & nC_ENLnDisType & " AS FTScdDiscType"
        tSql = tSql & vbCrLf & "," & nC_ENDiscType & " AS FNDctNo"
        tSql = tSql & vbCrLf & ",'" & SP_SQLtFunction("", "FTDctDiscType", "TCNMDiscountType", "WHERE FNDctNo = " & nC_ENDiscType, oDB.ocnOnLine, True, True) & "' AS FTDctName" 'read master
        tSql = tSql & vbCrLf & ",'" & tC_RefCode & "' AS FTDctRefCode"
        tSql = tSql & vbCrLf & ",'" & ptEventID & "' AS FTScdEventID"
        tSql = tSql & vbCrLf & ",'" & tC_CardID & "' AS FTScdCardID"
        tSql = tSql & vbCrLf & ",'" & SP_DTEtStrSTD(dC_CardExpDate) & "' AS FDScdExpDate"
        tSql = tSql & vbCrLf & ",'" & tC_CardType & "' AS FTScdCardType"
        
        Select Case nC_ENDiscType
            Case nEN_Disc12EmpCrd
                If UCase(tVB_DiscEmpChkLimit) = "Y" Then
                    tSql = tSql & vbCrLf & ",'1'  AS FTScdStaTrans"
                Else
                    tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
                End If
            Case Else
                tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
        End Select
'------------------------------------------------------------------------------------------------------------------------------------
        
        tSql = tSql & vbCrLf & ",'" & tC_CusID & "' AS FTScdCstID"
        tSql = tSql & vbCrLf & ",'" & tC_CardType & "' AS FTScdCardKey"
        tSql = tSql & vbCrLf & ",'" & tC_StaChkExpire & "' AS FTScdStaChkExpire"
        tSql = tSql & vbCrLf & ",'" & tC_StaChkDepart & "' AS FTScdStaChkDepart"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_AlwDis, True, False) & " AS FCScdMaxAlwDisc"
        tSql = tSql & vbCrLf & "," & nC_ActSeqNo & " AS FNScdActSeqNo"
        tSql = tSql & vbCrLf & "," & C_GETnGroup & " AS FNScdOrder"
        
        If nC_ENDiscType = nEN_Disc02PmtTime Then
            cDisPmt = pcDisAmt
        Else
            cDisPmt = 0
        End If
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cDisPmt, True, False) & " AS FCScdAmtTime"
        If nC_ENDiscType = nEN_Disc03Pmt Then
            cDisPmt = pcDisAmt
        Else
            cDisPmt = 0
        End If
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cDisPmt, True, False) & " AS FCScdAmtPmt"
        tSql = tSql & vbCrLf & "," & nC_DealGroup & " AS FNScdDealGroup"
        tSql = tSql & vbCrLf & "," & cC_DiscPerOrg & " AS FCScdDiscPerOrg"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscAmtOrg, True, False) & " AS FCScdDiscAmtOrg"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscNewOrg, True, False) & " AS FCScdDiscNewOrg"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(pcDiscPro, True, False) & " AS FCScdDiscSubPro"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(pcDiscReg, True, False) & " AS FCScdDiscSubReg"
        tSql = tSql & vbCrLf & "," & pcMallCrdRate & " AS FCSkuMallCardRate" '*KT 51-09-25 เก็บ Rate ของ Mall Card ที่ตั้งค่าที่ SKU
        tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tC_CardName) & "' AS FTScdCardName"  '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount '*KT 53-02-12 PH3-ML-SF-079Can't read Mall card which customer name or surname included
        tSql = tSql & vbCrLf & ",'" & IIf(bVB_CNModePreviewDay, "Y", "N") & "' AS FTScdStaPreview" '*KT 53-11-01 Preview Day
        Select Case nC_ENDiscType
            Case nEN_Disc12EmpCrd
'                tSql = tSql & vbCrLf & ",'" & tC_EmpRemark & "," & tC_Emplimit & "' AS FTRemark"                 '*Ao 2013/03/06     TK Emp EDC
                tSql = tSql & vbCrLf & ",'" & tC_Emplimit & "' AS FTRemark"                  '*Ao 2013/03/06     TK Emp EDC
            Case Else
                tSql = tSql & vbCrLf & ",''" & " AS FTRemark"                 '*Ao 2013/03/06     TK Emp EDC
        End Select
        
        tSql = tSql & vbCrLf & Left(SP_GETtWhoInsertValue, Len(SP_GETtWhoInsertValue) - 1)
        tSql = tSql & vbCrLf & " FROM " & tVB_CNTblDT & ""
        tSql = tSql & vbCrLf & tWhe
    Else
        bW_ChkMallCard = False
        tSql = "INSERT INTO " & tVB_CNTblCD & " WITH(ROWLOCK) (" & tField & SP_GETtWhoInsertField
        tSql = tSql & vbCrLf & " SELECT "
        tSql = tSql & vbCrLf & "FTStmCode , FTTmnNum, FTShdTransNo,0 AS FNSdtSeqNo"
        tSql = tSql & vbCrLf & "," & nMaxSeqNo & " AS FNScdSeqNo "
        tSql = tSql & vbCrLf & ",'' AS FTSkuCode,FTShdTransType,FDShdTransDate,'' AS FTSdtBarCode"
        '*KT 52-11-25 Deal&Promo
        If bC_PromoAftDeal Then
            tSql = tSql & vbCrLf & ",'1' AS FTSdtStkCode "
        Else
            tSql = tSql & vbCrLf & ",'' AS FTSdtStkCode "
        End If
        tSql = tSql & vbCrLf & ",'' AS FCSdtStkFac,'' AS FTSkuDepCode,'' AS FTSkuBrdCode,'' AS FTPunCode"
        tSql = tSql & vbCrLf & "," & "FCSdtQtyAllDisAvi" & " AS FCScdQtyB4 "
        tSql = tSql & vbCrLf & "," & 1 & " AS FCScdQty"
        tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtQtyAllDisAvi - " & pcAlwDisQty & " < 0 THEN 0 ELSE FCSdtQtyAllDisAvi - " & pcAlwDisQty & " END AS FCScdQtyAF"
        tSql = tSql & vbCrLf & ",0 AS FCScdAmtB4"
        tSql = tSql & vbCrLf & ",0 AS FCScdAmtB4Avi"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(pcDisAmt, True, False) & " AS FCScdAmt"
        tSql = tSql & vbCrLf & ",0 AS FCScdAmtAF"
        tSql = tSql & vbCrLf & ",0 AS FCScdAmtAFAvi"
        tSql = tSql & vbCrLf & "," & cC_DiscPer & " AS FCScdDiscPer"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscAmt, True, False) & " AS FCScdDiscAmt"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscNew, True, False) & " AS FCScdDiscNew"
        tSql = tSql & vbCrLf & "," & nC_ENLnDisType & " AS FTScdDiscType"
        tSql = tSql & vbCrLf & "," & nC_ENDiscType & " AS FNDctNo"
        tSql = tSql & vbCrLf & ",'" & SP_SQLtFunction("", "FTDctDiscType", "TCNMDiscountType", "WHERE FNDctNo = " & nC_ENDiscType, oDB.ocnOnLine, True, True) & "' AS FTDctName" 'read master
        tSql = tSql & vbCrLf & ",'" & tC_RefCode & "' AS FTDctRefCode"
        tSql = tSql & vbCrLf & ",'" & ptEventID & "' AS FTScdEventID"
        tSql = tSql & vbCrLf & ",'" & tC_CardID & "' AS FTScdCardID"
        tSql = tSql & vbCrLf & ",'" & SP_DTEtStrSTD(dC_CardExpDate) & "' AS FDScdExpDate"
        tSql = tSql & vbCrLf & ",'" & tC_CardType & "' AS FTScdCardType"
        
 '*KT 51-09-19 Communication Sheet ML-SF-558 Print  Remain Discount for Employee Card.xls   เก็บค่าไว้ใน FTScdStaTrans ว่า Emp Card ใดไม่ต้องแสดง Remain Discount
'        Select Case nC_ENDiscType
'            Case nEN_Disc01Deal, nEN_Disc02PmtTime, nEN_Disc03Pmt, nEN_Disc05LnDisc, nEN_Disc06LnMDown, nEN_Disc04CoopDisc
'                tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
'            Case Else
'                tSql = tSql & vbCrLf & ",'1'  AS FTScdStaTrans"
'        End Select

        Select Case nC_ENDiscType
            Case nEN_Disc12EmpCrd
                If UCase(tVB_DiscEmpChkLimit) = "Y" Then
                    tSql = tSql & vbCrLf & ",'1'  AS FTScdStaTrans"
                Else
                    tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
                End If
            Case Else
                tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
        End Select
'------------------------------------------------------------------------------------------------------------------------------------
        
        tSql = tSql & vbCrLf & ",'" & tC_CusID & "' AS FTScdCstID"
        tSql = tSql & vbCrLf & ",'" & tC_CardType & "' AS FTScdCardKey"
        tSql = tSql & vbCrLf & ",'" & tC_StaChkExpire & "' AS FTScdStaChkExpire"
        tSql = tSql & vbCrLf & ",'" & tC_StaChkDepart & "' AS FTScdStaChkDepart"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_AlwDis, True, False) & " AS FCScdMaxAlwDisc"
        tSql = tSql & vbCrLf & "," & nC_ActSeqNo & " AS FNScdActSeqNo"
        tSql = tSql & vbCrLf & "," & C_GETnGroup & " AS FNScdOrder"
        
        If nC_ENDiscType = nEN_Disc02PmtTime Then
            cDisPmt = pcDisAmt
        Else
            cDisPmt = 0
        End If
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cDisPmt, True, False) & " AS FCScdAmtTime"
        If nC_ENDiscType = nEN_Disc03Pmt Then
            cDisPmt = pcDisAmt
        Else
            cDisPmt = 0
        End If
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cDisPmt, True, False) & " AS FCScdAmtPmt"
        tSql = tSql & vbCrLf & "," & nC_DealGroup & " AS FNScdDealGroup"
        tSql = tSql & vbCrLf & "," & cC_DiscPerOrg & " AS FCScdDiscPerOrg"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscAmtOrg, True, False) & " AS FCScdDiscAmtOrg"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cC_DiscNewOrg, True, False) & " AS FCScdDiscNewOrg"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(pcDiscPro, True, False) & " AS FCScdDiscSubPro"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(pcDiscReg, True, False) & " AS FCScdDiscSubReg"
        tSql = tSql & vbCrLf & "," & pcMallCrdRate & " AS FCSkuMallCardRate" '*KT 51-09-25 เก็บ Rate ของ Mall Card ที่ตั้งค่าที่ SKU
        tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tC_CardName) & "' AS FTScdCardName"   '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount '*KT 53-02-12 PH3-ML-SF-079Can't read Mall card which customer name or surname included
        tSql = tSql & vbCrLf & ",'" & IIf(bVB_CNModePreviewDay, "Y", "N") & "' AS FTScdStaPreview" '*KT 53-11-01 Preview Day
        Select Case nC_ENDiscType
            Case nEN_Disc12EmpCrd
'                tSql = tSql & vbCrLf & ",'" & tC_EmpRemark & "," & tC_Emplimit & "' AS FTRemark"                 '*Ao 2013/03/06     TK Emp EDC
                tSql = tSql & vbCrLf & ",'" & tC_Emplimit & "' AS FTRemark"                  '*Ao 2013/03/06     TK Emp EDC
            Case Else
                tSql = tSql & vbCrLf & ",''" & " AS FTRemark"                 '*Ao 2013/03/06     TK Emp EDC
        End Select
        tSql = tSql & vbCrLf & Left(SP_GETtWhoInsertValue, Len(SP_GETtWhoInsertValue) - 1)
        tSql = tSql & vbCrLf & " FROM " & tVB_CNTblDT & ""
        tSql = tSql & vbCrLf & tWhe
    End If
    If SP_SQLbExecute(tSql, , False, nEffect, nErr) Then
        If nEffect <= 0 Then
            'Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, nErr, tVB_CNWheBill & "(C_DATxInsertDiscCD/1)", 2)  '*KT 52-03-17 PH1.5-ML-SF-120
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[4]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
            Call SP_MSGnShowing(tC_Msg003, tCS_CNMsgWarn)
        Else
            pbFlagInsCD = True '*KT 52-01-08 ทำการ set ค่า Flag ที่บอกว่ามีการบันทึกข้อมูล CD
        End If
    Else
        'Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, nErr, tVB_CNWheBill & "(C_DATxInsertDiscCD/2)", 2) '*KT 52-03-17 PH1.5-ML-SF-120
        Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[5]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
        Call SP_MSGnShowing(tC_Msg003, tCS_CNMsgWarn)
    End If
    Exit Sub
Err_Handler:
'    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, tVB_CNWheBill & "(C_DATxInsertDiscCD/3)", 2)  '*KT 52-03-17 PH1.5-ML-SF-120
    Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[6]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub C_DATxInsertFreeCD(ByVal ptFreeDesc$, ByVal pcFreeQty#)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 54-09-09
'DESCRIPTION       # *KT 54-09-09 V.4.0.1 ทำการแก้ไขให้สามารถใช้งานได้ปกติ
'ARGUMENT LIST #
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim tField As String
    Dim tValue As String
    Dim nMaxSeqNo As Integer
    Dim tWhe As String
    Dim nEffect As Long
    Dim TWheMax As String
    
    On Error GoTo Err_Handler
    
    tField = "FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTShdTransType"
    tField = tField & ",FNScdSeqNo,FDShdTransDate,FCScdQty,FNDctNo,FTDctName"
    tField = tField & ",FNScdOrder"
    tField = tField & ",FNScdActSeqNo"
    
    
    tWhe = " WHERE FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " And  FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " And FDShdTransDate = '" & tUT_SaleDate & "'"
    nMaxSeqNo = SP_VALcText2Double(SP_SQLtFunction("MAX", "FNScdSeqNo", "" & tVB_CNTblCD & "", tWhe))
    nMaxSeqNo = nMaxSeqNo + 1
    
    tSql = "INSERT INTO " & tVB_CNTblCD & " WITH(ROWLOCK) (" & tField & SP_GETtWhoInsertField
    tSql = tSql & vbCrLf & " VALUES('" & tVB_CNStoreCode & "'" 'FTStmCode
     tSql = tSql & vbCrLf & ",'" & tVB_CNTerminalNum & "'"                             'FTTmnNum
    tSql = tSql & vbCrLf & ",'" & tVB_Running & "'"                                                  'FTShdTransNo
    tSql = tSql & vbCrLf & ",0"                                                                                     'FNSdtSeqNo
    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "'"                                                          'FTShdTransType
    tSql = tSql & vbCrLf & "," & nMaxSeqNo                                                          'FNScdSeqNo
    tSql = tSql & vbCrLf & ",'" & tUT_SaleDate & "'"          'FDShdTransDate
    tSql = tSql & vbCrLf & "," & pcFreeQty                                                                 'FCScdQty
    tSql = tSql & vbCrLf & "," & nC_ENDiscType                                                     'FNDctNo
    tSql = tSql & vbCrLf & ",'" & ptFreeDesc & "'"                                                  'FTDctName
    tSql = tSql & vbCrLf & "," & C_GETnGroup
    tSql = tSql & vbCrLf & "," & nC_ActSeqNo                                                        'FNScdActSeqNo
    tSql = tSql & vbCrLf & SP_GETtWhoInsertValue
    If SP_SQLbExecute(tSql, , False, nEffect) Then
        If nEffect <= 0 Then
            Call SP_MSGnShowing(tC_Msg003, tCS_CNMsgWarn)
        End If
    Else
        Call SP_MSGnShowing(tC_Msg003, tCS_CNMsgWarn)
    End If
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function C_GETnGroup() As Integer
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Select Case nC_ENDiscType
        Case nEN_Disc01Deal
            C_GETnGroup = 1
        Case nEN_Disc02PmtTime, nEN_Disc03Pmt, nEN_Disc04CoopDisc
            C_GETnGroup = 2
        Case nEN_Disc05LnDisc, nEN_Disc06LnMDown
            C_GETnGroup = 3
        Case nEN_Disc09CoopCpn
            C_GETnGroup = 4
        Case nEN_Disc10CoopBrand
            C_GETnGroup = 5
        Case nEN_Disc07AutoTran
            C_GETnGroup = 6
        Case nEN_Disc08MallCpn
            C_GETnGroup = 7
        Case nEN_Disc11MallCrd
            C_GETnGroup = 8
        Case nEN_Disc12EmpCrd
            C_GETnGroup = 9
        Case nEN_Disc13TranDisc
            C_GETnGroup = 10
        Case nEN_Disc14TranMDown
            C_GETnGroup = 11
        Case nEN_Disc16FreeGift
            C_GETnGroup = 12
        Case nEN_Disc15Ontop
            C_GETnGroup = 13
        Case nEN_Disc23TranMDown1, nEN_Disc24TranMDown2, nEN_Disc25TranMDown3, nEN_Disc26TranMDown4, nEN_Disc27TranMDown5  '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
            C_GETnGroup = 11
    End Select
End Function

Private Function C_GETcDiscount() As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsTemp As ADODB.Recordset
    On Error GoTo Err_Handler
    Call C_DATxGetDataDT(orsTemp)
    If orsTemp Is Nothing Then Exit Function
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        Select Case nC_ENDiscType
            Case nEN_Disc01Deal
                'none
            Case nEN_Disc02PmtTime
                C_GETcDiscount = C_GETcDiscPmtTime(orsTemp)
            
            Case nEN_Disc03Pmt
                C_GETcDiscount = C_GETcDiscPmtPmt(orsTemp)
            
            Case nEN_Disc04CoopDisc
                C_GETcDiscount = C_GETcDiscCoopDisc(orsTemp)
                
            Case nEN_Disc05LnDisc
                C_GETcDiscount = C_GETcDiscLnDisc(orsTemp)
                
            Case nEN_Disc06LnMDown
                C_GETcDiscount = C_GETcDiscLnDisc(orsTemp)
                
            Case nEN_Disc09CoopCpn
                C_GETcDiscount = C_GETcDiscCoopCpn(orsTemp)
                
            Case nEN_Disc10CoopBrand
                C_GETcDiscount = C_GETcDiscCoopCpn(orsTemp)
                
            Case nEN_Disc07AutoTran
                C_GETcDiscount = C_GETcDiscAutoTrans(orsTemp)
                
            Case nEN_Disc08MallCpn
                C_GETcDiscount = C_GETcDiscMallCpn(orsTemp)
                
            Case nEN_Disc11MallCrd
'                C_GETcDiscount = C_GETcDiscMallCard(True, orsTemp)
                C_GETcDiscount = C_GETcDiscMallCardEx(orsTemp)            '*Ao 2013/03/08  TK
            Case nEN_Disc12EmpCrd
'                C_GETcDiscount = C_GETcDiscEmpCard(orsTemp)           '*Ao 51-06-20
                C_GETcDiscount = C_GETcDiscEmpCardEx(orsTemp)                          '*Ao 2013/03/08  TK
            Case nEN_Disc13TranDisc
                C_GETcDiscount = C_GETcDiscDiscMrkTran(orsTemp)
            
            Case nEN_Disc14TranMDown
                C_GETcDiscount = C_GETcDiscDiscMrkTran(orsTemp)
                
            Case nEN_Disc23TranMDown1 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                C_GETcDiscount = C_GETcDiscDiscMrkTran(orsTemp)
            
            Case nEN_Disc24TranMDown2 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                C_GETcDiscount = C_GETcDiscDiscMrkTran(orsTemp)
                
            Case nEN_Disc25TranMDown3 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                C_GETcDiscount = C_GETcDiscDiscMrkTran(orsTemp)
                
            Case nEN_Disc26TranMDown4 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                C_GETcDiscount = C_GETcDiscDiscMrkTran(orsTemp)
                
            Case nEN_Disc27TranMDown5 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                C_GETcDiscount = C_GETcDiscDiscMrkTran(orsTemp)
                
        End Select
        
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Public Function C_CHKnDDP(ByVal pnDTSeqNo As Integer) As Integer '*KT 52-11-28 Deal&Promo
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls เรื่องสินค้า Promotion ซ้อนกันกับ Deal กลุ่ม Buy
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim orsTemp2 As ADODB.Recordset
    Dim tDctNo As String
    Dim nDctNo As Long
    Dim tDiscType As String
On Error GoTo Err_Handler
    tC_TheBestType = ""
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        tDctNo = ""
        tSql = "SELECT CD.FNDctNo"
        tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD & " CD WITH(NOLOCK) "
        tSql = tSql & vbCrLf & " WHERE  CD.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND CD.FNSdtSeqNo = " & pnDTSeqNo
        tSql = tSql & vbCrLf & " GROUP BY CD.FNDctNo"
        If SP_SQLvExecute(tSql, orsTemp) = 0 Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                Do While Not orsTemp.EOF
                    nDctNo = SP_FEDvChkNumeric(orsTemp, "FNDctNo")
                    Select Case nC_ENDiscType
                        Case 2, 3
                            Select Case nDctNo
                                Case 1 'Deal
                                    tSql = "SELECT CD.FNDctNo,CD.FCScdAmt,CD.FTScdStaChkDepart,CD.FTScdDiscType"
                                    tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD & " CD WITH(NOLOCK)  "
                                    tSql = tSql & vbCrLf & " WHERE  CD.FTShdTransNo = '" & tVB_Running & "'"
                                    tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
                                    tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & tUT_SaleDate & "'"
                                    tSql = tSql & vbCrLf & " AND CD.FNSdtSeqNo = " & pnDTSeqNo
                                    tSql = tSql & vbCrLf & " AND CD.FNDctNo = " & 1
                                    tSql = tSql & vbCrLf & " GROUP BY CD.FNDctNo,CD.FCScdAmt,CD.FTScdStaChkDepart,CD.FTScdDiscType"
                                    If SP_SQLvExecute(tSql, orsTemp2) = 0 Then
                                        tDiscType = SP_FEDtChkString(orsTemp2, "FTScdDiscType")
                                        If SP_FEDvChkNumeric(orsTemp2, "FCScdAmt") <= 0 And SP_FEDtChkString(orsTemp2, "FTScdStaChkDepart") = "1" And (tDiscType = "0" Or tDiscType = "4") Then '*KT 53-01-12 ทำการเปลี่ยน FTScdStaChkDepart ของกลุ่ม Buy ของ Free เป็น 1
                                            tDctNo = tDctNo & ""
                                        ElseIf SP_FEDvChkNumeric(orsTemp2, "FCScdAmt") <= 0 And SP_FEDtChkString(orsTemp2, "FTScdStaChkDepart") = "1" Then 'Deal กลุ่ม Buy
                                            tDctNo = tDctNo & ""
                                        Else
                                            tDctNo = tDctNo & CStr(nDctNo)
                                        End If
                                    Else
                                        tDctNo = tDctNo & ""
                                    End If
                                Case Else
                                     tDctNo = tDctNo & CStr(nDctNo)
                            End Select
                        Case Else
                            tDctNo = tDctNo & CStr(nDctNo)
                    End Select
                    orsTemp.MoveNext
                    If Not orsTemp.EOF Then tDctNo = tDctNo & ","
                Loop
            End If
            If tDctNo <> "" Then
                tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy"
                tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO  WITH(NOLOCK)  "
                tSql = tSql & vbCrLf & "  WHERE PO.FNDdpNo2 = " & nC_ENDiscType
                tSql = tSql & vbCrLf & "  AND PO.FNDdpNo1 IN (" & tDctNo & ")"
                If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                    If Not (orsTemp.BOF And orsTemp.EOF) Then
                            orsTemp.Filter = "FTDdpStaPolicy = 'N'"
                            If Not (orsTemp.BOF And orsTemp.EOF) Then
                                C_CHKnDDP = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
                            Else
                                orsTemp.Filter = "FTDdpStaPolicy = 'B'"
                                If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
                                    Do While Not orsTemp.EOF
                                        tC_TheBestType = tC_TheBestType & IIf(tC_TheBestType <> "", ",", "") & SP_FEDvChkNumeric(orsTemp, "FNDdpNo1") 'เก็บเลขที่การลดไว้ เพื่อเทียบดูว่าสินค้ารายการปัจจุบันเป็น The Best กับตัวไหนบ้าง
                                        orsTemp.MoveNext
                                    Loop
                                    C_CHKnDDP = 2
                                Else
                                    C_CHKnDDP = 1
                                End If
                            End If
                    Else
                        C_CHKnDDP = 1
                    End If
                Else
                    C_CHKnDDP = 0 'Error ไม่ผ่าน
                End If
            Else
                C_CHKnDDP = 1
            End If
        Else
            C_CHKnDDP = 0
        End If
    Else
        tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy"
        tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO WITH(NOLOCK)  LEFT JOIN " & tVB_CNTblCD & " CD WITH(NOLOCK)  ON PO.FNDdpNo1 = CD.FNDctNo"
        tSql = tSql & vbCrLf & "  WHERE PO.FNDdpNo2 = " & nC_ENDiscType
        tSql = tSql & vbCrLf & " AND CD.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND CD.FNSdtSeqNo = " & pnDTSeqNo
        tSql = tSql & vbCrLf & " AND CD.FCScdAmt >0" '*KT 53-01-12 ทำการเปลี่ยน FTScdStaChkDepart ของกลุ่ม Buy ของ Free เป็น 1
        tSql = tSql & vbCrLf & " AND CD.FTScdStaChkDepart <> '1'" '*KT 53-01-12 ทำการเปลี่ยน FTScdStaChkDepart ของกลุ่ม Buy ของ Free เป็น 1
        tSql = tSql & vbCrLf & " ORDER BY CD.FNScdSeqNo "
        If SP_SQLbExecute(tSql, orsTemp) Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                orsTemp.Filter = "FTDdpStaPolicy = 'N'"
                If Not (orsTemp.BOF And orsTemp.EOF) Then
                    C_CHKnDDP = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
                Else
                    orsTemp.Filter = "FTDdpStaPolicy = 'B'"
                    If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
                        Do While Not orsTemp.EOF
                            tC_TheBestType = tC_TheBestType & IIf(tC_TheBestType <> "", ",", "") & SP_FEDvChkNumeric(orsTemp, "FNDdpNo1") 'เก็บเลขที่การลดไว้ เพื่อเทียบดูว่าสินค้ารายการปัจจุบันเป็น The Best กับตัวไหนบ้าง
                            orsTemp.MoveNext
                        Loop
                        C_CHKnDDP = 2
                    Else
                        C_CHKnDDP = 1
                    End If
                End If
            Else
                C_CHKnDDP = 1 'ยังไม่มีส่วนลดอื่น ให้ผ่าน
            End If
        Else
            C_CHKnDDP = 0 'Error ไม่ผ่าน
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsTemp2)
    Exit Function
Err_Handler:
    C_CHKnDDP = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsTemp2)
End Function

Private Function C_CHKnDDPEmp() As Integer
'CREATE BY            # Kitti                             LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 52-02-11
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 52-02-02  PH2-ML-SF-014 แก้ไขใน POSFront 1.5.6 '*KT 52-02-11 PH2-ML-SF-090โ ทำให้ Sub all ผิด เมื่อมีส่วนลดแบบ Employee discount
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tDctNo$
On Error GoTo Err_Handler
    tC_TheBestType = ""
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        tDctNo = ""
        tSql = "SELECT CD.FNDctNo"
        tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD & " CD WITH(NOLOCK)  "
        tSql = tSql & vbCrLf & " WHERE  CD.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " GROUP BY CD.FNDctNo"
        If SP_SQLvExecute(tSql, orsTemp) = 0 Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                Do While Not orsTemp.EOF
                    tDctNo = tDctNo & SP_FEDtChkString(orsTemp, "FNDctNo")
                    orsTemp.MoveNext
                    If Not orsTemp.EOF Then tDctNo = tDctNo & ","
                Loop
            End If
            If tDctNo <> "" Then
                tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy"
                tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO WITH(NOLOCK)  "
                tSql = tSql & vbCrLf & "  WHERE PO.FNDdpNo1 = " & nEN_Disc12EmpCrd
                tSql = tSql & vbCrLf & "  AND PO.FNDdpNo2 = " & nC_ENDiscType
                tSql = tSql & vbCrLf & "  AND PO.FNDdpNo1 IN (" & tDctNo & ")"
                If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                    If Not (orsTemp.BOF And orsTemp.EOF) Then
                            orsTemp.Filter = "FTDdpStaPolicy = 'N'"
                            If Not (orsTemp.BOF And orsTemp.EOF) Then
                                C_CHKnDDPEmp = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
                            Else
                                orsTemp.Filter = "FTDdpStaPolicy = 'B'"
                                If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
                                    Do While Not orsTemp.EOF
                                        tC_TheBestType = tC_TheBestType & IIf(tC_TheBestType <> "", ",", "") & SP_FEDvChkNumeric(orsTemp, "FNDdpNo1") 'เก็บเลขที่การลดไว้ เพื่อเทียบดูว่าสินค้ารายการปัจจุบันเป็น The Best กับตัวไหนบ้าง
                                        orsTemp.MoveNext
                                    Loop
                                    C_CHKnDDPEmp = 2
                                Else
                                    C_CHKnDDPEmp = 1
                                End If
                            End If
                    Else
                        C_CHKnDDPEmp = 1
                    End If
                Else
                    C_CHKnDDPEmp = 0
                End If
            Else
                C_CHKnDDPEmp = 1
            End If
        Else
            C_CHKnDDPEmp = 0
        End If
    Else
        tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy"
        tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO WITH(NOLOCK)  INNER JOIN " & tVB_CNTblCD & " CD WITH(NOLOCK) ON PO.FNDdpNo1 = CD.FNDctNo"
        tSql = tSql & vbCrLf & " WHERE PO.FNDdpNo1 = " & nEN_Disc12EmpCrd
        tSql = tSql & vbCrLf & " AND PO.FNDdpNo2 = " & nC_ENDiscType
        tSql = tSql & vbCrLf & " AND CD.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " ORDER BY CD.FNScdSeqNo "
        If SP_SQLbExecute(tSql, orsTemp) Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                orsTemp.Filter = "FTDdpStaPolicy = 'N'"
                If Not (orsTemp.BOF And orsTemp.EOF) Then
                    C_CHKnDDPEmp = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
                Else
                    orsTemp.Filter = "FTDdpStaPolicy = 'B'"
                    If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
                        Do While Not orsTemp.EOF
                            tC_TheBestType = tC_TheBestType & IIf(tC_TheBestType <> "", ",", "") & SP_FEDvChkNumeric(orsTemp, "FNDdpNo1") 'เก็บเลขที่การลดไว้ เพื่อเทียบดูว่าสินค้ารายการปัจจุบันเป็น The Best กับตัวไหนบ้าง
                            orsTemp.MoveNext
                        Loop
                        C_CHKnDDPEmp = 2
                    Else
                        C_CHKnDDPEmp = 1
                    End If
                End If
            Else
                C_CHKnDDPEmp = 1 'ยังไม่มีส่วนลดอื่น ให้ผ่าน
            End If
        Else
            C_CHKnDDPEmp = 0
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    C_CHKnDDPEmp = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function


Public Function C_CHKnDDPMallCard(ByVal pnDTSeqNo As Integer) As Integer
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tDctNo$
On Error GoTo Err_Handler
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        tDctNo = ""
        tSql = "SELECT CD.FNDctNo"
        tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD & " CD WITH(NOLOCK)  "
        tSql = tSql & vbCrLf & " WHERE  CD.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND CD.FNSdtSeqNo = " & pnDTSeqNo
        tSql = tSql & vbCrLf & " GROUP BY CD.FNDctNo"
        If SP_SQLvExecute(tSql, orsTemp) = 0 Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                Do While Not orsTemp.EOF
                    tDctNo = tDctNo & SP_FEDtChkString(orsTemp, "FNDctNo")
                    orsTemp.MoveNext
                    If Not orsTemp.EOF Then tDctNo = tDctNo & ","
                Loop
            End If
            If tDctNo <> "" Then
                tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy"
                tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO WITH(NOLOCK)  "
                tSql = tSql & vbCrLf & "  WHERE PO.FNDdpNo1 = " & nC_ENDiscType
                tSql = tSql & vbCrLf & "  AND PO.FNDdpNo2 = " & nEN_Disc11MallCrd
                tSql = tSql & vbCrLf & "  AND PO.FNDdpNo1 IN (" & tDctNo & ")"
                If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                    If Not (orsTemp.BOF And orsTemp.EOF) Then
                            orsTemp.Filter = "FTDdpStaPolicy = 'N'"
                            If Not (orsTemp.BOF And orsTemp.EOF) Then
                                C_CHKnDDPMallCard = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
                            Else
                                orsTemp.Filter = "FTDdpStaPolicy = 'B'"
                                If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
                                    Do While Not orsTemp.EOF
                                        tC_TheBestType = tC_TheBestType & IIf(tC_TheBestType <> "", ",", "") & SP_FEDvChkNumeric(orsTemp, "FNDdpNo1") 'เก็บเลขที่การลดไว้ เพื่อเทียบดูว่าสินค้ารายการปัจจุบันเป็น The Best กับตัวไหนบ้าง
                                        orsTemp.MoveNext
                                    Loop
                                    C_CHKnDDPMallCard = 2
                                Else
                                    C_CHKnDDPMallCard = 1
                                End If
                            End If
                    Else
                        C_CHKnDDPMallCard = 1
                    End If
                Else
                    C_CHKnDDPMallCard = 0
                End If
            Else
                C_CHKnDDPMallCard = 1
            End If
        Else
            C_CHKnDDPMallCard = 0
        End If
    Else
        tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy"
        tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO WITH(NOLOCK)  LEFT JOIN " & tVB_CNTblCD & " CD WITH(NOLOCK)  ON PO.FNDdpNo1 = CD.FNDctNo"
        tSql = tSql & vbCrLf & " WHERE PO.FNDdpNo1 = " & nC_ENDiscType
        tSql = tSql & vbCrLf & " AND PO.FNDdpNo2 = " & nEN_Disc11MallCrd
        tSql = tSql & vbCrLf & " AND CD.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND CD.FNSdtSeqNo = " & pnDTSeqNo
        tSql = tSql & vbCrLf & " ORDER BY CD.FNScdSeqNo "
        If SP_SQLbExecute(tSql, orsTemp) Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                orsTemp.Filter = "FTDdpStaPolicy = 'N'"
                If Not (orsTemp.BOF And orsTemp.EOF) Then
                    C_CHKnDDPMallCard = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
                Else
                    orsTemp.Filter = "FTDdpStaPolicy = 'B'"
                    If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
                        Do While Not orsTemp.EOF
                            orsTemp.MoveNext
                        Loop
                        C_CHKnDDPMallCard = 2
                    Else
                        C_CHKnDDPMallCard = 1
                    End If
                End If
            Else
                C_CHKnDDPMallCard = 1 'ยังไม่มีส่วนลดอื่น ให้ผ่าน
            End If
        Else
            C_CHKnDDPMallCard = 0 'Error ไม่ผ่าน
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    C_CHKnDDPMallCard = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Public Function C_CHKnDDPDeal(ByVal pnDTSeqNo As Integer) As Integer
'CREATE BY            # Kitti                        LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 52-11-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 52-11-25  Deal&Promo
'*KT 52-12-29 แก้ไขตาม DCR PRJ09010TheMall-R3-DCR_12002(Deal)-01.00.00.xls เรื่องสินค้า Promotion ซ้อนกันกับ Deal กลุ่ม Buy
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tDctNo$
    Dim tDiscType As String
    Dim nTmpDctNo As String
On Error GoTo Err_Handler
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        tDctNo = ""
        tSql = "SELECT CD.FNDctNo,CD.FCScdAmt,CD.FTScdStaChkDepart,CD.FTScdDiscType"
        tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD & " CD WITH(NOLOCK)  "
        tSql = tSql & vbCrLf & " WHERE  CD.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND CD.FNSdtSeqNo = " & pnDTSeqNo
        tSql = tSql & vbCrLf & " GROUP BY CD.FNDctNo,CD.FCScdAmt,CD.FTScdStaChkDepart,CD.FTScdDiscType"
        If SP_SQLvExecute(tSql, orsTemp) = 0 Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                Do While Not orsTemp.EOF
                    Select Case SP_FEDvChkNumeric(orsTemp, "FNDctNo")
                        Case 1
                            tDiscType = SP_FEDtChkString(orsTemp, "FTScdDiscType")
                            nTmpDctNo = SP_FEDvChkNumeric(orsTemp, "FNDctNo")
                            If SP_FEDvChkNumeric(orsTemp, "FCScdAmt") <= 0 And SP_FEDtChkString(orsTemp, "FTScdStaChkDepart") = "1" And (tDiscType = "0" Or tDiscType = "4") Then '*KT 53-01-12 ทำการเปลี่ยน FTScdStaChkDepart ของกลุ่ม Buy ของ Free เป็น 1
                                tDctNo = tDctNo & ""
                            ElseIf SP_FEDvChkNumeric(orsTemp, "FCScdAmt") <= 0 And SP_FEDtChkString(orsTemp, "FTScdStaChkDepart") = "1" Then 'Deal กลุ่ม Buy
                                tDctNo = tDctNo & ""
                            Else
                                tDctNo = tDctNo & CStr(nTmpDctNo)
                            End If
                        Case Else
                            tDctNo = tDctNo & CStr(nTmpDctNo)
                    End Select
                    orsTemp.MoveNext
                    If Not orsTemp.EOF Then tDctNo = tDctNo & ","
                Loop
            End If
            If tDctNo <> "" Then
                tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy"
                tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO WITH(NOLOCK)  "
                tSql = tSql & vbCrLf & "  WHERE PO.FNDdpNo1 = " & nEN_Disc01Deal 'nC_ENDiscType
                tSql = tSql & vbCrLf & "  AND PO.FNDdpNo2 = " & nC_ENDiscType 'nEN_Disc01Deal
                'tSql = tSql & vbCrLf & "  AND PO.FNDdpNo1 IN (" & tDctNo & ")"
                If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                    If Not (orsTemp.BOF And orsTemp.EOF) Then
                            orsTemp.Filter = "FTDdpStaPolicy = 'N'"
                            If Not (orsTemp.BOF And orsTemp.EOF) Then
                                C_CHKnDDPDeal = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
                            Else
                                orsTemp.Filter = "FTDdpStaPolicy = 'B'"
                                If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
                                    Do While Not orsTemp.EOF
                                        tC_TheBestType = tC_TheBestType & IIf(tC_TheBestType <> "", ",", "") & SP_FEDvChkNumeric(orsTemp, "FNDdpNo1") 'เก็บเลขที่การลดไว้ เพื่อเทียบดูว่าสินค้ารายการปัจจุบันเป็น The Best กับตัวไหนบ้าง
                                        orsTemp.MoveNext
                                    Loop
                                    C_CHKnDDPDeal = 2
                                Else
                                    C_CHKnDDPDeal = 1
                                End If
                            End If
                    Else
                        C_CHKnDDPDeal = 1
                    End If
                Else
                    C_CHKnDDPDeal = 0
                End If
            Else
                C_CHKnDDPDeal = 1
            End If
        Else
            C_CHKnDDPDeal = 0
        End If
    Else
        tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy,CD.FCScdAmt,CD.FTScdStaChkDepart,CD.FTScdDiscType"
        tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO WITH(NOLOCK)  LEFT JOIN " & tVB_CNTblCD & " CD WITH(NOLOCK)  ON PO.FNDdpNo1 = CD.FNDctNo"
        tSql = tSql & vbCrLf & " WHERE PO.FNDdpNo1 = " & nEN_Disc01Deal 'nC_ENDiscType
        tSql = tSql & vbCrLf & " AND PO.FNDdpNo2 = " & nC_ENDiscType 'nEN_Disc01Deal
        tSql = tSql & vbCrLf & " AND CD.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND CD.FNSdtSeqNo = " & pnDTSeqNo
        tSql = tSql & vbCrLf & " AND CD.FCScdAmt >0" '*KT 53-01-12 ทำการเปลี่ยน FTScdStaChkDepart ของกลุ่ม Buy ของ Free เป็น 1
        tSql = tSql & vbCrLf & " AND CD.FTScdStaChkDepart <> '1'" '*KT 53-01-12 ทำการเปลี่ยน FTScdStaChkDepart ของกลุ่ม Buy ของ Free เป็น 1
        tSql = tSql & vbCrLf & " ORDER BY CD.FNScdSeqNo "
        If SP_SQLbExecute(tSql, orsTemp) Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                orsTemp.Filter = "FTDdpStaPolicy = 'N'"
                If Not (orsTemp.BOF And orsTemp.EOF) Then
                    C_CHKnDDPDeal = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
                Else
                    orsTemp.Filter = "FTDdpStaPolicy = 'B'"
                    If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
                        Do While Not orsTemp.EOF
                            orsTemp.MoveNext
                        Loop
                        C_CHKnDDPDeal = 2
                    Else
                        C_CHKnDDPDeal = 1
                    End If
                End If
            Else
                C_CHKnDDPDeal = 1 'ยังไม่มีส่วนลดอื่น ให้ผ่าน
            End If
        Else
            C_CHKnDDPDeal = 0 'Error ไม่ผ่าน
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    C_CHKnDDPDeal = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function


'Private Function C_CHKnDDPMPrice() As Integer
''CREATE BY            # Kitti                             LAST MODIFIED BY           # Kitti
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 52-04-22
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  #
''ต้องกลับไปแก้ไขที่ SP_CHKnDDPMPrice
''*KT 52-04-22 ทำการตรวจสอบ Policy ของส่วนลดที่ทำ ณ ปัจจุบันกับราคาสินค้า Member Price
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim tSql As String
'    Dim orsTemp As ADODB.Recordset
'    On Error GoTo Err_Handler
'    tC_TheBestType = ""
'    tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy"
'    tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO "
'    tSql = tSql & vbCrLf & " WHERE PO.FNDdpNo1 IN (" & nEN_Disc21MPrice & "," & nEN_Disc22MDisc & ")" '*KT 52-09-07 Member Price + Consignment
'    tSql = tSql & vbCrLf & " AND PO.FNDdpNo2 = " & nC_ENDiscType
'    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then 'Read master
'    If Not (orsTemp.BOF And orsTemp.EOF) Then
'            orsTemp.Filter = "FTDdpStaPolicy = 'N'"
'            If Not (orsTemp.BOF And orsTemp.EOF) Then
'                C_CHKnDDPMPrice = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
'            Else
'                orsTemp.Filter = "FTDdpStaPolicy = 'B'"
'                If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
'                    Do While Not orsTemp.EOF
'                        tC_TheBestType = tC_TheBestType & IIf(tC_TheBestType <> "", ",", "") & SP_FEDvChkNumeric(orsTemp, "FNDdpNo1") 'เก็บเลขที่การลดไว้ เพื่อเทียบดูว่าสินค้ารายการปัจจุบันเป็น The Best กับตัวไหนบ้าง
'                        orsTemp.MoveNext
'                    Loop
'                    C_CHKnDDPMPrice = 2
'                Else
'                    C_CHKnDDPMPrice = 1
'                End If
'            End If
'        Else
'            C_CHKnDDPMPrice = 1 'ยังไม่มีส่วนลดอื่น ให้ผ่าน
'        End If
'    Else
'        C_CHKnDDPMPrice = 3 'ให้ผ่าน
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'    Exit Function
'Err_Handler:
'    C_CHKnDDPMPrice = 0
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'    Call SP_DATxRsNothing(orsTemp)
'End Function

Private Function C_GETcB4DisDDP(ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cTotalAviY As Double
    Dim tWhe As String
    Dim nCount As Integer
    Dim nFNSdtSeqNo As Long
    'Dim tFTSkuCode As String
    Dim cFCSdtSaleAmt As Double
    Dim cFCSdtAmtDisAvi As Double
    On Error GoTo Err_Handler
    
    cC_QtyAll = 0               '*Ao 51-06-19   เอาไว้เช็ค  NewPrice ตั้งเกินราคา
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
        'tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
        cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
        cFCSdtAmtDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAvi")
        tWhe = "WHERE FTSkuCode = '" & SP_FEDtChkString(porsTemp, "FTSkuCode") & "'"
        Select Case nC_ENDiscType
            Case nEN_Disc13TranDisc
                tWhe = tWhe & " AND FTSkuStaDisAlwTran = 'Y'"
                tWhe = tWhe & " AND FDSkuDisAlwEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
            Case nEN_Disc14TranMDown
                tWhe = tWhe & " AND FTSkuStaMrkAlwTran = 'Y'"
                tWhe = tWhe & " AND FDSkuMrkAlwEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
            Case nEN_Disc23TranMDown1 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                 tWhe = tWhe & " AND FTSkuStaMrkAlwTran1 = 'Y'"
                tWhe = tWhe & " AND FDSkuMrkAlwEndDate1 >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
            Case nEN_Disc24TranMDown2
                 tWhe = tWhe & " AND FTSkuStaMrkAlwTran2 = 'Y'"
                tWhe = tWhe & " AND FDSkuMrkAlwEndDate2 >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
            Case nEN_Disc25TranMDown3
                 tWhe = tWhe & " AND FTSkuStaMrkAlwTran3 = 'Y'"
                tWhe = tWhe & " AND FDSkuMrkAlwEndDate3 >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
            Case nEN_Disc26TranMDown4
                 tWhe = tWhe & " AND FTSkuStaMrkAlwTran4 = 'Y'"
                tWhe = tWhe & " AND FDSkuMrkAlwEndDate4 >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
            Case nEN_Disc27TranMDown5
                 tWhe = tWhe & " AND FTSkuStaMrkAlwTran5 = 'Y'"
                tWhe = tWhe & " AND FDSkuMrkAlwEndDate5 >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
        End Select
        nCount = SP_VALnText2Double(SP_SQLtFunction("COUNT", "FTSkuCode", tVB_CNTblSku, tWhe, oDB.ocnOnLine, True, True)) 'read master
        If nCount >= 1 Then
            Select Case C_CHKnDDP(nFNSdtSeqNo) 'Check Double Policy
                Case 0 'ไม่อนุญาต
                    'คิดเฉพาะส่วนที่เหลือ
                    cTotalAviY = cTotalAviY + SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAvi")
                Case 1 'อนุญาต
                    cTotalAviY = cTotalAviY + SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
                    cC_QtyAll = cC_QtyAll + SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAll")               '*Ao 51-06-19
                Case 2 'หาผลที่ดีที่สุด
                    'ไม่มี
            End Select
        End If
        porsTemp.MoveNext
    Loop
    C_GETcB4DisDDP = cTotalAviY
    Exit Function
Err_Handler:
    C_GETcB4DisDDP = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_GETbYellowLabel(ByRef porsTemp As ADODB.Recordset, Optional ByVal pbCoupon As Boolean = True) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim cTotalAviN As Double
    Dim cTotalAviY As Double
    Dim cTotalAviB As Double
    Dim tWhe As String
    Dim nCount As Integer
On Error GoTo Err_Handler
    tSql = "SELECT DT.FNSdtSeqNo, DT.FCSdtSaleAmt, DT.FCSdtAmtDisAvi, DT.FTSkuStaMallEmpCard,DT.FCSdtQtyAllDisAvi,DT.FCSdtBluePrice,DT.FCSdtYellowPrice"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " DT WITH(NOLOCK)  LEFT JOIN " & tVB_CNTblSku & " SKU WITH(NOLOCK)  ON DT.FTSkuCode = SKU.FTSkuCode"
    tSql = tSql & vbCrLf & "WHERE  DT.FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND DT.FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND DT.FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND ISNULL(DT.FTSdtCancelFlag,'0') <> '1'"
    If pbCoupon Then
        tSql = tSql & vbCrLf & " AND SKU.FTSkuStaMallCpn = 'Y'"
    End If
    If SP_SQLbExecute(tSql, porsTemp) Then
        C_GETbYellowLabel = True
    Else
        C_GETbYellowLabel = False
    End If
    Exit Function
Err_Handler:
    C_GETbYellowLabel = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_GETcB4DisDDPSpecial(ByVal pbIsFindBlue As Boolean, ByRef porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim cTotalAviN As Double
    Dim cTotalAviY As Double
    Dim cTotalAviB As Double
    Dim tWhe As String
    Dim nCount As Integer

    
On Error GoTo Err_Handler
    
    tSql = "SELECT DT.FNSdtSeqNo,DT.FTSkuCode"
    If pbIsFindBlue Then
        tSql = tSql & vbCrLf & ",DT.FCSdtBluePrice AS FCSdtAmtDisAviXXX"
    Else
        tSql = tSql & vbCrLf & ",DT.FCSdtYellowPrice AS FCSdtAmtDisAviXXX"
    End If
    tSql = tSql & vbCrLf & ",DT.FCSdtSalePrice,DT.FCSdtB4DisChg,DT.FCSdtSaleAmt,DT.FCSdtAmtDisAvi,DT.FCSdtQtyAllDisAvi"
    tSql = tSql & vbCrLf & " FROM " & tVB_CNTblDT & " DT WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & "WHERE  DT.FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND DT.FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND DT.FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND ISNULL(DT.FTSdtCancelFlag,'0') <> '1'"
'*TW 51-08-15
    If pbIsFindBlue Then
        tSql = tSql & vbCrLf & " AND DT.FCSdtBluePrice > 0"
    Else
        tSql = tSql & vbCrLf & " AND DT.FCSdtYellowPrice > 0"
    End If
    
    If SP_SQLbExecute(tSql, porsTemp) Then
        If Not (porsTemp.BOF And porsTemp.EOF) Then
            Do While Not porsTemp.EOF
                cTotalAviY = cTotalAviY + SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAviXXX")
                porsTemp.MoveNext
            Loop
        Else
            C_GETcB4DisDDPSpecial = 0 'ยังไม่มีส่วนลดอื่น ให้ผ่าน
        End If
    Else
        C_GETcB4DisDDPSpecial = 0 'Error ไม่ผ่าน
    End If
    C_GETcB4DisDDPSpecial = cTotalAviY
    Exit Function
Err_Handler:
    C_GETcB4DisDDPSpecial = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function


Private Sub C_DISxGetNextAct()
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tWhe As String
    On Error GoTo Err_Handler
    tWhe = " WHERE FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " And  FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " And FDShdTransDate = '" & tUT_SaleDate & "'"
    nC_ActSeqNo = CInt(SP_VALcText2Double(SP_SQLtFunction("MAX", "FNScdActSeqNo", "" & tVB_CNTblCD & "", tWhe)))
    nC_ActSeqNo = nC_ActSeqNo + 1
    Exit Sub
Err_Handler:
    nC_ActSeqNo = 1
End Sub

Private Function C_GETcDiscPmtTime(ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'ขั้นตอน
'1 : เช็คช่วงวันเวลาที่ได้ Promotion Time Sale ตามวันที่ปัจจุบันกับเวลาตามเอกสาร
'2 : ตรวจสอบจำนวน Max ของ Promotion Time Sale ที่จำกัดไว้ ถ้าจำนวนที่นำมาคำนวณมากกว่าจำนวน Max จะทำการคิดส่วนลดได้เต็มที่ที่ค่า Max
'3 : แยกการคำนวณได้ 3 แบบ คือ % Off ,Amount Off,New Price
'4 : ตรวจสอบว่าเป็น Double Policy กับส่วนที่ได้มาก่อนหน้าหรือไม่ ดังนี้
'    ไม่อนุญาต : นำส่วนที่เหลือมาคำนวณหาส่วนลด
'     อนุญาต : นำทั้งหมดมาคำนวณหาส่วนลด
'     the bast : รอดำเนินการ
'5 : คำนวณหาส่วนลดตาม
'   % Off  :  ส่วนลดที่ได้ = % * จำนวนที่คิดส่วนลด/100
'   Amount Off :
'   Tax Except Sale : ส่วนลดที่ได้ =  (Amount Off  * 100 / (100 + Tax Rate ของสินค้า)) *  จำนวนชิ้นที่ลดได้
'    Other : ส่วนลดที่ได้ =  Amount Off  * จำนวนชิ้นที่ลดได้
'    New Price :
'   Tax Except Sale : ส่วนลดที่ได้ = จำนวนที่คิดส่วนลด - (New Price * 100/(100 + Tax Rate ของสินค้า))  * จำนวนชิ้นที่ลดได้)
'    Other : ส่วนลดที่ได้ =  จำนวนที่คิดส่วนลด - (New Price * จำนวนชิ้นที่ลดได้)
'6 : ตรวจสอบความถูกต้อง เช่น % มากว่า 100 ,ส่วนลดที่ได้ติดลบ
'7 : ตรวจสอบกับ Mall Card ว่าเป็น Double Policy กันหรือไม่เพื่อทำการ Flag ค่าเป็น *
'8 : Insert Data ลง CD

    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tEventID As String
    Dim tFTSkuCode As String
    Dim cFCSdtSalePrice As Double
    Dim cFCSdtSaleAmt As Double
    Dim cFCSdtAmtDisAvi As Double
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cSourceAmt As Double
    
    Dim tStaPmtType As String
    Dim nMaxQrty As Long
    Dim tDocTime As String
    Dim tWhe As String
    
    Dim cFCSdtTax As String
    
    '*KT 52-02-16 member price
    Dim tStaMemPriceAlw As String
    Dim cMemberPrice As Double
    
On Error GoTo Err_Handler
    If nC_DTSeqNo < 1 Then Exit Function
    
    porsTemp.Filter = "FNSdtSeqNo = " & nC_DTSeqNo
    If (porsTemp.BOF And porsTemp.EOF) Then Exit Function
    
    tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
    cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
    cFCSdtAmtDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAvi")
    cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
    cFCSdtTax = SP_FEDvChkNumeric(porsTemp, "FCSdtTax")             '*Ao 51-05-08
    tStaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd")
    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice")
    
    tWhe = "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " AND FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    'วันที่ปัจจุบันเวลาตามเอกสาร
    tDocTime = SP_SQLtFunction("", "FTShdTransTime", "" & tVB_CNTblHD & "", tWhe)
    tSql = "SELECT FTEvnCode,FTPptStaPromoType,FCPptSkuAmt,FNPptSkuMaxQty,FTPptSkuStartTime,FTPptSkuEndTime,FDPptSkuEndDate "
    tSql = tSql & vbCrLf & "FROM  " & tVB_CNTblPromoTime & " WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE FTSkuCode = '" & tFTSkuCode & "'"
    tSql = tSql & vbCrLf & "AND FDPptSkuEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*TW 51-04-19 '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    tSql = tSql & vbCrLf & "AND '" & tDocTime & "' BETWEEN FTPptSkuStartTime AND FTPptSkuEndTime"
    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            tEventID = SP_FEDtChkString(orsTemp, "FTEvnCode")
            tStaPmtType = SP_FEDtChkString(orsTemp, "FTPptStaPromoType")
            cC_DiscValue = SP_FEDvChkNumeric(orsTemp, "FCPptSkuAmt")
            nMaxQrty = SP_FEDvChkNumeric(orsTemp, "FNPptSkuMaxQty")
            If cFCSdtQtyAllDisAvi > nMaxQrty Then cFCSdtQtyAllDisAvi = nMaxQrty
            Select Case tStaPmtType
                Case "1" '% Off
                    nC_ENLnDisType = nEN_LnDis1Per
                    cC_DiscPer = cC_DiscValue
                Case "2" 'Amount Off
                    nC_ENLnDisType = nEN_LnDis2Amt
                    cC_DiscAmt = cC_DiscValue
                Case "3" 'New Price
                    nC_ENLnDisType = nEN_LnDis3NewPrice
                    cC_DiscNew = cC_DiscValue
                Case "4" 'No Promotion
                    SP_DATxRsNothing (orsTemp)
                    Exit Function
                Case Else '*KT 51-12-16 V 1.0.3 Phase 2 ข้อ 16 Communication Sheet PH2-ML-SF-003 Coop. Dis is not shown on Screen when Copy Sale.xls
                    SP_DATxRsNothing (orsTemp)
                    Exit Function
            End Select
            cC_DiscPerOrg = cC_DiscPer
            cC_DiscAmtOrg = cC_DiscAmt
            cC_DiscNewOrg = cC_DiscNew

            Select Case C_CHKnDDP(nC_DTSeqNo) 'Check Double Policy
                Case 0 'ไม่อนุญาต
                    'คิดเฉพาะส่วนที่เหลือ
                    cSourceAmt = cFCSdtAmtDisAvi
                Case 1 'อนุญาต
                    cSourceAmt = cFCSdtSaleAmt
                    bC_PromoAftDeal = False '*KT 52-11-25 Deal&Promo
                Case 2 'หาผลที่ดีที่สุด
                    'รอดำเนินการ
                    bC_PromoAftDeal = False '*KT 52-11-25 Deal&Promo
            End Select
        
            If cSourceAmt > 0 Then
                Select Case nC_ENLnDisType
                    Case nEN_LnDis1Per 'Percent Discount
                        cC_DiscPer = cC_DiscValue
                        cC_DiscAmt = cC_DiscPer * cSourceAmt / 100
                    Case nEN_LnDis2Amt  ' Amount Discount               '*AO 51-05-08
                        Select Case Val(Format(tVB_TransactionSub))
                                Case tEN_TrnTaxExceptSale
                                    cC_DiscAmt = (cC_DiscValue * 100 / (100 + cFCSdtTax)) * cFCSdtQtyAllDisAvi
                                Case Else
                                    cC_DiscAmt = cC_DiscValue * cFCSdtQtyAllDisAvi
                            End Select
                            cC_DiscPer = 0
                    Case nEN_LnDis3NewPrice 'New Price
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnTaxExceptSale
                                cC_DiscAmt = cSourceAmt - ((cC_DiscValue * 100 / (100 + cFCSdtTax)) * cFCSdtQtyAllDisAvi)
                            Case Else
                                cC_DiscAmt = cSourceAmt - (cC_DiscValue * cFCSdtQtyAllDisAvi)
                            End Select
                            cC_DiscPer = 0
                End Select
            Else
                cC_DiscPer = 0
                cC_DiscAmt = 0
            End If
            
            If cC_DiscPer > 100 Then
                Call SP_MSGnShowing(tC_Msg004, tCS_CNMsgWarn)
                Call SP_DATxRsNothing(orsTemp)
                Exit Function
            End If
            If cSourceAmt - cC_DiscAmt < 0 Then
                Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
                Call SP_DATxRsNothing(orsTemp)
                Exit Function
            End If
            
            '*KT 52-02-13 Member price
            '---------------------------------------------------------------------------------------------
'            If bVB_CNMemberPrice Then
'                If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" And tStaMemPriceAlw = "Y" Then
'                    If (cFCSdtSaleAmt - cC_DiscAmt) < cMemberPrice Then
'                         'none
'                         If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
'                    Else
''                         If C_DATxSumDTMemberPrice(nC_DTSeqNo, tFTSkuCode) Then
'                            cC_DiscAmt = 0
'                            C_GETcDiscPmtTime = cC_DiscAmt
'                            tVB_StaMemPriceAlw = "Y"
''                            tVB_StaMemPricePromoAlw = true
''                        Else
''                            If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
''                        End If
'                    End If
'                Else
'                     'none
'                       If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
'                End If
'            Else
'                    'none
'                      If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
'            End If
            '*Ao 2011/10/13  CommSheet ML-PRV-050-SF Wrong Member Price when Copy Sale.xls
            If bVB_CNMemberPrice Then
                If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" And tStaMemPriceAlw = "Y" Then
                    If cMemberPrice = 0 Then
                        'none
                        If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
                    Else
                        If (cFCSdtSaleAmt - cC_DiscAmt) < cMemberPrice Then
                             'none
                              If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
                        Else
    '                        If C_DATxSumDTMemberPrice(nC_DTSeqNo, tFTSkuCode) Then
                                cC_DiscAmt = 0
                                C_GETcDiscPmtTime = cC_DiscAmt
                                tVB_StaMemPriceAlw = "Y"
    '                             tVB_StaMemPricePromoAlw = True
    '                        Else
    '                            If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
    '                        End If
                        End If
                    End If
                Else
                     'none
                       If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
                End If
            Else
                    'none
                     If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
            End If
            '---------------------------------------------------------------------------------------------
            C_GETcDiscPmtTime = cC_DiscAmt
            cSourceAmt = cC_DiscAmt                 '*Ao 51-07-25     'เหลือที่ลดได้อีก
            If cC_DiscAmt > 0 Then
                ''*KT 51-06-03 ใส่เครื่องหมาย * หน้าสินค้าที่ไม่ยอมให้ทำ Double Policy กับ Mall Card
                '--------------------------------------------------------
                Select Case C_CHKnDDPMallCard(nC_DTSeqNo)
                    Case 0 'Not Allow
                        tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) "
                        tSql = tSql & vbCrLf & " SET FTSkuStaMallEmpCard='N'"
                        tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
                        tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
                        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
                        tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nC_DTSeqNo
                        Call SP_SQLbExecute(tSql)
                    Case 1 'allow
                        'none
                    Case 2 'Thebest
                        'none
                End Select
                '--------------------------------------------------------
                Call C_DATxInsertDiscCD(cC_DiscAmt, nC_DTSeqNo, tEventID, cFCSdtQtyAllDisAvi, cFCSdtAmtDisAvi, cFCSdtQtyAllDisAvi)
            End If
        End If
    End If
    
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'*KT 51-10-03 3.0.128/3.0.200 Communication Sheet ML-SF-572 Price Check is incorrect.xls
Public Function C_GETcChkDiscPmtTime(ByVal ptFTSkuCode$) As Double
'CREATE BY            # Kitti                             LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-10-03
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim orsSKU As ADODB.Recordset
    Dim tEventID As String
    Dim tFTSkuCode As String
    Dim cFCSdtSaleAmt As Double
    Dim cFCSdtAmtDisAvi As Double
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cSourceAmt As Double
    
    Dim tStaPmtType As String
    Dim nMaxQrty As Long
    Dim tDocTime As String
    Dim tWhe As String
    Dim cFCSdtTax As String
    Dim cFCSdtQtyAll As Double
    
On Error GoTo Err_Handler
        
    tWhe = "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " AND FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    'วันที่ปัจจุบันเวลาตามเอกสาร
    tDocTime = SP_SQLtFunction("", "FTShdTransTime", "" & tVB_CNTblHD & "", tWhe)
    
    tSql = "SELECT    TA.FCTaxRate, TA.FTTaxCode,SKU.FCSkuRegPrice"
    tSql = tSql & vbCrLf & "FROM    " & tVB_CNTblSku & " SKU WITH(NOLOCK)  INNER JOIN"
    tSql = tSql & vbCrLf & "TCNMTaxMtn TA  WITH(NOLOCK)  ON SKU.FTTaxCode = TA.FTTaxCode"
    tSql = tSql & vbCrLf & "WHERE SKU.FTSkuCode ='" & ptFTSkuCode & "'"
    Call SP_SQLvExecute(tSql, orsSKU, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsSKU.BOF Or orsSKU.EOF) Then
        '*Ao 2011/04/22   CommSheet ST-FL-016 Error on Price Check for Instore Bar.xls
        If oSKU.GETInStoreBar Then
            cFCSdtAmtDisAvi = oSKU.GETSKUSaleWeight
            cFCSdtSaleAmt = oSKU.GETSKUSaleWeight
        Else
            cFCSdtAmtDisAvi = SP_FEDvChkNumeric(orsSKU, "FCSkuRegPrice")
            cFCSdtSaleAmt = SP_FEDvChkNumeric(orsSKU, "FCSkuRegPrice")
        End If
        '______________________________________________________________________
'        cFCSdtAmtDisAvi = SP_FEDvChkNumeric(orsSKU, "FCSkuRegPrice")
'        cFCSdtSaleAmt = SP_FEDvChkNumeric(orsSKU, "FCSkuRegPrice")
        cFCSdtQtyAll = 1
        cFCSdtQtyAllDisAvi = 1
        cFCSdtTax = SP_FEDtChkString(orsSKU, "FCTaxRate")
    Else
        cFCSdtAmtDisAvi = 0
        cFCSdtSaleAmt = 0
        cFCSdtQtyAll = 1
        cFCSdtQtyAllDisAvi = 0
        cFCSdtTax = 0
    End If
    
    tSql = "SELECT FTEvnCode,FTPptStaPromoType,FCPptSkuAmt,FNPptSkuMaxQty,FTPptSkuStartTime,FTPptSkuEndTime,FDPptSkuEndDate "
    tSql = tSql & vbCrLf & "FROM  " & tVB_CNTblPromoTime & " WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE FTSkuCode = '" & ptFTSkuCode & "'"
    tSql = tSql & vbCrLf & "AND FDPptSkuEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*TW 51-04-19 '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    tSql = tSql & vbCrLf & "AND '" & tDocTime & "' BETWEEN FTPptSkuStartTime AND FTPptSkuEndTime"
    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            tEventID = SP_FEDtChkString(orsTemp, "FTEvnCode")
            tStaPmtType = SP_FEDtChkString(orsTemp, "FTPptStaPromoType")
            cC_DiscValue = SP_FEDvChkNumeric(orsTemp, "FCPptSkuAmt")
            nMaxQrty = SP_FEDvChkNumeric(orsTemp, "FNPptSkuMaxQty")
            If cFCSdtQtyAllDisAvi > nMaxQrty Then cFCSdtQtyAllDisAvi = nMaxQrty
            Select Case tStaPmtType
                Case "1" '% Off
                    nC_ENLnDisType = nEN_LnDis1Per
                    cC_DiscPer = cC_DiscValue
                Case "2" 'Amount Off
                    nC_ENLnDisType = nEN_LnDis2Amt
                    cC_DiscAmt = cC_DiscValue
                Case "3" 'New Price
                    nC_ENLnDisType = nEN_LnDis3NewPrice
                    cC_DiscNew = cC_DiscValue
                Case "0" 'No Promotion
                    Exit Function
            End Select
            cC_DiscPerOrg = cC_DiscPer
            cC_DiscAmtOrg = cC_DiscAmt
            cC_DiscNewOrg = cC_DiscNew

            Select Case C_CHKnDDP(nC_DTSeqNo) 'Check Double Policy
                Case 0 'ไม่อนุญาต
                    'คิดเฉพาะส่วนที่เหลือ
                    cSourceAmt = cFCSdtAmtDisAvi
                Case 1 'อนุญาต
                    cSourceAmt = cFCSdtSaleAmt
                Case 2 'หาผลที่ดีที่สุด
                    'รอดำเนินการ
            End Select
        
            If cSourceAmt > 0 Then
                Select Case nC_ENLnDisType
                    Case nEN_LnDis1Per 'Percent Discount
                        cC_DiscPer = cC_DiscValue
                        cC_DiscAmt = cC_DiscPer * cSourceAmt / 100
                    Case nEN_LnDis2Amt  ' Amount Discount               '*AO 51-05-08
                        Select Case Val(Format(tVB_TransactionSub))
                                Case tEN_TrnTaxExceptSale
                                    'cC_DiscAmt = (cC_DiscValue * 100 / (100 + cFCSdtTax)) * cFCSdtQtyAllDisAvi
                                    cC_DiscAmt = cC_DiscValue * cFCSdtQtyAllDisAvi  '*KT 51-10-06 V 3.0.128/V 3.0.200
                                Case Else
                                    cC_DiscAmt = cC_DiscValue * cFCSdtQtyAllDisAvi
                            End Select
                            cC_DiscPer = 0
                    Case nEN_LnDis3NewPrice 'New Price
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnTaxExceptSale
                                'cC_DiscAmt = cSourceAmt - ((cC_DiscValue * 100 / (100 + cFCSdtTax)) * cFCSdtQtyAllDisAvi) '*KT 51-10-06 V 3.0.128/V 3.0.200
                                cC_DiscAmt = cSourceAmt - (cC_DiscValue * cFCSdtQtyAllDisAvi)
                            Case Else
                                cC_DiscAmt = cSourceAmt - (cC_DiscValue * cFCSdtQtyAllDisAvi)
                            End Select
                            cC_DiscPer = 0
                End Select
            Else
                cC_DiscPer = 0
                cC_DiscAmt = 0
            End If
            
            If cC_DiscPer > 100 Then
                Call SP_MSGnShowing(tC_Msg004, tCS_CNMsgWarn)
                Call SP_DATxRsNothing(orsSKU)
                Call SP_DATxRsNothing(orsTemp)
                Exit Function
            End If
            If cSourceAmt - cC_DiscAmt < 0 Then
                Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
                Call SP_DATxRsNothing(orsSKU)
                Call SP_DATxRsNothing(orsTemp)
                Exit Function
            End If
            C_GETcChkDiscPmtTime = cSourceAmt - cC_DiscAmt
        Else
            C_GETcChkDiscPmtTime = 0
        End If
    Else
        C_GETcChkDiscPmtTime = 0
    End If
    
    If C_GETcChkDiscPmtTime = 0 Then
        tSql = "SELECT FTEvnCode,FTPpmStaPromoType,FCPpmSkuAmt,FNPpmSkuMaxQty,FDPpmSkuEndDate "
        tSql = tSql & vbCrLf & "FROM  " & tVB_CNTblPromo & " WITH(NOLOCK)  "
        tSql = tSql & vbCrLf & "WHERE FTSkuCode = '" & ptFTSkuCode & "'"
        tSql = tSql & vbCrLf & "AND FDPpmSkuEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
        If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                tEventID = SP_FEDtChkString(orsTemp, "FTEvnCode")
                tStaPmtType = SP_FEDtChkString(orsTemp, "FTPpmStaPromoType")
                cC_DiscValue = SP_FEDvChkNumeric(orsTemp, "FCPpmSkuAmt")
                nMaxQrty = SP_FEDvChkNumeric(orsTemp, "FNPpmSkuMaxQty")
                If cFCSdtQtyAllDisAvi > nMaxQrty Then cFCSdtQtyAllDisAvi = nMaxQrty
                Select Case tStaPmtType
                    Case "1"                           '%
                        nC_ENLnDisType = nEN_LnDis1Per
                        cC_DiscPer = cC_DiscValue
                    Case "2"                    'Amt
                        nC_ENLnDisType = nEN_LnDis2Amt
                        cC_DiscAmt = cC_DiscValue
                    Case "3"            'NewPrice
                        nC_ENLnDisType = nEN_LnDis3NewPrice
                        cC_DiscNew = cC_DiscValue
                    Case "0"
                        Call SP_DATxRsNothing(orsSKU)
                        Call SP_DATxRsNothing(orsTemp)
                        Exit Function
                End Select
                cC_DiscPerOrg = cC_DiscPer
                cC_DiscAmtOrg = cC_DiscAmt
                cC_DiscNewOrg = cC_DiscNew
    
                Select Case C_CHKnDDP(nC_DTSeqNo) 'Check Double Policy
                    Case 0 'ไม่อนุญาต
                        'คิดเฉพาะส่วนที่เหลือ
                        'If cFCSdtQtyAllDisAvi <= 0 Then Exit Function
                        cSourceAmt = cFCSdtAmtDisAvi
                    Case 1 'อนุญาต
                        cSourceAmt = cFCSdtSaleAmt
                    Case 2 'หาผลที่ดีที่สุด
                        'รอดำเนินการ
                End Select
            
                If cSourceAmt > 0 Then
                    Select Case nC_ENLnDisType
                        Case nEN_LnDis1Per 'Percent Discount
                            cC_DiscPer = cC_DiscValue
                            cC_DiscAmt = cC_DiscValue * cSourceAmt / 100
                        Case nEN_LnDis2Amt  ' Amount Discount
                            Select Case Val(Format(tVB_TransactionSub))                      '*AO 51-05-08
                                Case tEN_TrnTaxExceptSale
                                    'cC_DiscAmt = (cC_DiscValue * 100 / (100 + cFCSdtTax)) * cFCSdtQtyAllDisAvi
                                    cC_DiscAmt = cC_DiscValue * cFCSdtQtyAllDisAvi '*KT 51-10-06 V 3.0.128/V 3.0.200
                                Case Else
                                    cC_DiscAmt = cC_DiscValue * cFCSdtQtyAllDisAvi
                            End Select
                            cC_DiscPer = 0
                        Case nEN_LnDis3NewPrice 'New Price
                            Select Case Val(Format(tVB_TransactionSub))               '*AO 51-05-08
                                Case tEN_TrnTaxExceptSale
                                    'cC_DiscAmt = cSourceAmt - ((cC_DiscValue * 100 / (100 + cFCSdtTax)) * cFCSdtQtyAll)
                                     cC_DiscAmt = cC_DiscValue * cFCSdtQtyAllDisAvi ' '*KT 51-10-06 V 3.0.128/V 3.0.200
                                Case Else
                                    cC_DiscAmt = cSourceAmt - (cC_DiscValue * cFCSdtQtyAll)
                            End Select
                            cC_DiscPer = 0
                    End Select
                Else
                    cC_DiscPer = 0
                    cC_DiscAmt = 0
                End If
                If cC_DiscPer > 100 Then
                    Call SP_MSGnShowing(tC_Msg004, tCS_CNMsgWarn)
                    Call SP_DATxRsNothing(orsSKU)
                    Call SP_DATxRsNothing(orsTemp)
                    Exit Function
                End If
                If cSourceAmt - cC_DiscAmt < 0 Then
                    Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
                    Call SP_DATxRsNothing(orsSKU)
                    Call SP_DATxRsNothing(orsTemp)
                    Exit Function
                End If
                C_GETcChkDiscPmtTime = cSourceAmt - cC_DiscAmt
            Else
                C_GETcChkDiscPmtTime = cFCSdtAmtDisAvi
            End If
        Else
            C_GETcChkDiscPmtTime = cFCSdtAmtDisAvi
        End If
    End If
    
    Call SP_DATxRsNothing(orsSKU)
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    C_GETcChkDiscPmtTime = 0
    Call SP_DATxRsNothing(orsSKU)
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_GETcDiscPmtPmt(ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'ขั้นตอน
'1 : เช็คสินค้าว่าอยู่ใน Promotion
'2 : ตรวจสอบจำนวน Max ของ Promotion ที่จำกัดไว้ ถ้าจำนวนที่นำมาคำนวณมากกว่าจำนวน Max จะทำการคิดส่วนลดได้เต็มที่ที่ค่า Max
'3 : แยกการคำนวณได้ 3 แบบ คือ % Off ,Amount Off,New Price
'4 : ตรวจสอบว่าเป็น Double Policy กับส่วนที่ได้มาก่อนหน้าหรือไม่ ดังนี้
'    ไม่อนุญาต : นำส่วนที่เหลือมาคำนวณหาส่วนลด
'     อนุญาต : นำทั้งหมดมาคำนวณหาส่วนลด
'     the bast : รอดำเนินการ
'5 : คำนวณหาส่วนลดตาม
'   % Off  :  ส่วนลดที่ได้ = % * จำนวนที่คิดส่วนลด/100
'   Amount Off :
'   Tax Except Sale : ส่วนลดที่ได้ =  (Amount Off  * 100 / (100 + Tax Rate ของสินค้า)) *  จำนวนชิ้นที่ลดได้
'    Other : ส่วนลดที่ได้ =  Amount Off  * จำนวนชิ้นที่ลดได้
'    New Price :
'   Tax Except Sale : ส่วนลดที่ได้ = จำนวนที่คิดส่วนลด - (New Price * 100/(100 + Tax Rate ของสินค้า))  * จำนวนชิ้นที่ลดได้)
'    Other : ส่วนลดที่ได้ =  จำนวนที่คิดส่วนลด - (New Price * จำนวนชิ้นที่ลดได้)
'6 : ตรวจสอบความถูกต้อง เช่น % มากว่า 100 ,ส่วนลดที่ได้ติดลบ
'7 : ตรวจสอบกับ Mall Card ว่าเป็น Double Policy กันหรือไม่เพื่อทำการ Flag ค่าเป็น *
'8 : Insert Data ลง CD

    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tEventID As String
    Dim tFTSkuCode As String
    Dim cFCSdtSalePrice As Double
    Dim cFCSdtSaleAmt As Double
    Dim cFCSdtAmtDisAvi As Double
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cSourceAmt As Double
    Dim cFCSdtQtyAll As Double
    Dim tStaPmtType As String
    Dim nMaxQrty As Long
    Dim tWhe As String
    
    Dim cFCSdtTax As String
    
    '*KT 52-02-16 member price
    Dim tStaMemPriceAlw As String
    Dim cMemberPrice As Double
    Dim bFlagInsCD As Boolean
    
On Error GoTo Err_Handler
    If nC_DTSeqNo < 1 Then Exit Function
    
    porsTemp.Filter = "FNSdtSeqNo = " & nC_DTSeqNo
    If (porsTemp.BOF And porsTemp.EOF) Then Exit Function
    
    tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
    cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
    cFCSdtAmtDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAvi")
    cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
    cFCSdtQtyAll = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAll")
    cFCSdtTax = SP_FEDvChkNumeric(porsTemp, "FCSdtTax")                         '*Ao 51-05-08
    tStaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") '*KT 52-03-25 แก้ไข POSFront 1.8.1
    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice")
     
    tSql = "SELECT FTEvnCode,FTPpmStaPromoType,FCPpmSkuAmt,FNPpmSkuMaxQty,FDPpmSkuEndDate "
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblPromo & " WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE FTSkuCode = '" & tFTSkuCode & "'"
    tSql = tSql & vbCrLf & "AND FDPpmSkuEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            tEventID = SP_FEDtChkString(orsTemp, "FTEvnCode")
            tStaPmtType = SP_FEDtChkString(orsTemp, "FTPpmStaPromoType")
            cC_DiscValue = SP_FEDvChkNumeric(orsTemp, "FCPpmSkuAmt")
            nMaxQrty = SP_FEDvChkNumeric(orsTemp, "FNPpmSkuMaxQty")
            If cFCSdtQtyAllDisAvi > nMaxQrty Then cFCSdtQtyAllDisAvi = nMaxQrty
            Select Case tStaPmtType
                Case "1"                           '%
                    nC_ENLnDisType = nEN_LnDis1Per
                    cC_DiscPer = cC_DiscValue
                Case "2"                    'Amt
                    nC_ENLnDisType = nEN_LnDis2Amt
                    cC_DiscAmt = cC_DiscValue
                Case "3"            'NewPrice
                    nC_ENLnDisType = nEN_LnDis3NewPrice
                    cC_DiscNew = cC_DiscValue
                Case "4"    'No Promotion
                    Exit Function
                Case Else '*KT 51-12-16 V 1.0.3 Phase 2 ข้อ 16 Communication Sheet PH2-ML-SF-003 Coop. Dis is not shown on Screen when Copy Sale.xls
                    Exit Function
            End Select
            cC_DiscPerOrg = cC_DiscPer
            cC_DiscAmtOrg = cC_DiscAmt
            cC_DiscNewOrg = cC_DiscNew

            Select Case C_CHKnDDP(nC_DTSeqNo) 'Check Double Policy
                Case 0 'ไม่อนุญาต
                    'คิดเฉพาะส่วนที่เหลือ
                    'If cFCSdtQtyAllDisAvi <= 0 Then Exit Function
                    cSourceAmt = cFCSdtAmtDisAvi
                Case 1 'อนุญาต
                    cSourceAmt = cFCSdtSaleAmt
                    bC_PromoAftDeal = False '*KT 52-11-25 Deal&Promo
                Case 2 'หาผลที่ดีที่สุด
                    'รอดำเนินการ
                    bC_PromoAftDeal = False '*KT 52-11-25 Deal&Promo
            End Select
        
            If cSourceAmt > 0 Then
                Select Case nC_ENLnDisType
                    Case nEN_LnDis1Per 'Percent Discount
                        cC_DiscPer = cC_DiscValue
                        cC_DiscAmt = cC_DiscValue * cSourceAmt / 100
                    Case nEN_LnDis2Amt  ' Amount Discount
                        Select Case Val(Format(tVB_TransactionSub))                      '*AO 51-05-08
                            Case tEN_TrnTaxExceptSale
                                cC_DiscAmt = (cC_DiscValue * 100 / (100 + cFCSdtTax)) * cFCSdtQtyAllDisAvi
                            Case Else
                                cC_DiscAmt = cC_DiscValue * cFCSdtQtyAllDisAvi
                        End Select
                        cC_DiscPer = 0
                    Case nEN_LnDis3NewPrice 'New Price
                        Select Case Val(Format(tVB_TransactionSub))               '*AO 51-05-08
                            Case tEN_TrnTaxExceptSale
                                cC_DiscAmt = cSourceAmt - ((cC_DiscValue * 100 / (100 + cFCSdtTax)) * cFCSdtQtyAll)
                            Case Else
                                cC_DiscAmt = cSourceAmt - (cC_DiscValue * cFCSdtQtyAll)
                        End Select
                        cC_DiscPer = 0
                End Select
            Else
                cC_DiscPer = 0
                cC_DiscAmt = 0
            End If
            If cC_DiscPer > 100 Then
                Call SP_MSGnShowing(tC_Msg004, tCS_CNMsgWarn)
                Call SP_DATxRsNothing(orsTemp)
                Exit Function
            End If
            If cSourceAmt - cC_DiscAmt < 0 Then
                Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
                Call SP_DATxRsNothing(orsTemp)
                Exit Function
            End If
            
            
            '*KT 52-02-13 Member price
            '---------------------------------------------------------------------------------------------
'            If bVB_CNMemberPrice Then
'                If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" And tStaMemPriceAlw = "Y" Then
'                    If (cFCSdtSaleAmt - cC_DiscAmt) < cMemberPrice Then
'                         'none
'                          If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
'                    Else
''                        If C_DATxSumDTMemberPrice(nC_DTSeqNo, tFTSkuCode) Then
'                            cC_DiscAmt = 0
'                            C_GETcDiscPmtPmt = cC_DiscAmt
'                            tVB_StaMemPriceAlw = "Y"
''                             tVB_StaMemPricePromoAlw = True
''                        Else
''                            If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
''                        End If
'                    End If
'                Else
'                     'none
'                       If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
'                End If
'            Else
'                    'none
'                     If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
'            End If
            '*Ao 2011/10/13  CommSheet ML-PRV-050-SF Wrong Member Price when Copy Sale.xls
            If bVB_CNMemberPrice Then
                If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" And tStaMemPriceAlw = "Y" Then
                    If cMemberPrice = 0 Then
                        'none
                        If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
                    Else
                        If (cFCSdtSaleAmt - cC_DiscAmt) < cMemberPrice Then
                             'none
                              If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
                        Else
    '                        If C_DATxSumDTMemberPrice(nC_DTSeqNo, tFTSkuCode) Then
                                cC_DiscAmt = 0
                                C_GETcDiscPmtPmt = cC_DiscAmt
                                tVB_StaMemPriceAlw = "Y"
    '                             tVB_StaMemPricePromoAlw = True
    '                        Else
    '                            If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
    '                        End If
                        End If
                    End If
                Else
                     'none
                       If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
                End If
            Else
                    'none
                     If Not tVB_StaMemPricePromoAlw Then tVB_StaMemPriceAlw = "N"
            End If
            '---------------------------------------------------------------------------------------------
            
            C_GETcDiscPmtPmt = cC_DiscAmt
            cSourceAmt = cC_DiscAmt                '*Ao 51-07-25     'เหลือที่ลดได้อีก
            If cC_DiscAmt > 0 Then
                ''*KT 51-06-03 ใส่เครื่องหมาย * หน้าสินค้าที่ไม่ยอมให้ทำ Double Policy กับ Mall Card
                '--------------------------------------------------------
                Select Case C_CHKnDDPMallCard(nC_DTSeqNo)
                    Case 0 'Not Allow
                        tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) "
                        tSql = tSql & vbCrLf & " SET FTSkuStaMallEmpCard='N'"
                        tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
                        tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
                        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
                        tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nC_DTSeqNo
                        Call SP_SQLbExecute(tSql)
                    Case 1 'allow
                        'none
                    Case 2 'Thebest
                        'none
                End Select
                '--------------------------------------------------------
                Call C_DATxInsertDiscCD(cC_DiscAmt, nC_DTSeqNo, tEventID, cFCSdtQtyAllDisAvi, cFCSdtAmtDisAvi, cFCSdtQtyAllDisAvi)    '
                
            End If
        End If
    End If
    
    
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_GETcDiscCoopDisc(ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'ขั้นตอน
'1 : ตรวจสอบว่าเป็น Double Policy กับส่วนที่ได้มาก่อนหน้าหรือไม่ ดังนี้
'    ไม่อนุญาต : นำส่วนที่เหลือมาคำนวณหาส่วนลด
'     อนุญาต : นำทั้งหมดมาคำนวณหาส่วนลด
'     the bast : รอดำเนินการ
'2 : ตรวจสอบสินค้าใน Table SKU วันที่หมดอายุ/เป็น Cooperate Discount
'3 : คำนวนหาส่วนลดจาก % ที่ Get มาได้ ส่วนลดที่ได้ = จำนวนที่คำนว

    Dim nFNSdtSeqNo As Integer
    Dim tFTSkuCode As String
    Dim cFCSdtSaleAmt As Double
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tEventID As String
    Dim cFCSdtAmtDisAvi As Double
    Dim cFCSdtQtyAllDisAvi As Double
    
'ใช้คำนวณ
    Dim cDisPer As Double
    Dim cDisAmt As Double
    Dim cSourceAmt As Double
    
    Dim tStaMemPriceAlw As String
    Dim dSkuCoDisEndDate As Date
    Dim tSkuStaCoDisAlw As String
    Dim cFCSdtMPriceAmt As Double
    Dim cFCSdtQty As Double
    Dim bTran As Boolean
On Error GoTo ErrHandle
    If nC_DTSeqNo < 1 Then Exit Function
    porsTemp.Filter = "FNSdtSeqNo = " & nC_DTSeqNo
    If (porsTemp.BOF And porsTemp.EOF) Then
        Call SP_DATxRsNothing(orsTemp)
        Exit Function
    End If
    tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
    cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
    cFCSdtAmtDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAvi")
    cFCSdtMPriceAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice")
    cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
    tStaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd")
    dSkuCoDisEndDate = SP_FEDdChkDate(porsTemp, "FDSkuCoDisEndDate")
    tSkuStaCoDisAlw = SP_FEDtChkString(porsTemp, "FTSkuStaCoDisAlw")
    cFCSdtQty = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAll")
    
    Select Case C_CHKnDDP(nC_DTSeqNo) 'Check Double Policy
        Case 0 'ไม่อนุญาต
            'คิดเฉพาะส่วนที่เหลือ
            'If cFCSdtQtyAllDisAvi <= 0 Then Exit Function
            cSourceAmt = cFCSdtAmtDisAvi
        Case 1 'อนุญาต
            cSourceAmt = cFCSdtSaleAmt
        Case 2 'หาผลที่ดีที่สุด
            'รอดำเนินการ
    End Select
    
    '*KT 52-02-17 member price เอกสาร Mail วันที่  17/02/2009
    'คำถาม ในกรณีที่ป้อน SKU ที่ได้ราคา Member, ถ้ามี Discount อื่นๆ เช่น Coop.Disc จะลดได้อีกหรือไม่?
    'ถ้าลดได้ ส่วนลดอะไรบ้างที่ลดได้?, ส่วนลดอะไรที่ลดไม่ได้? และสามารถ Price Override ได้หรือไม?
    'FSBT - กรณี Cooperate Disc. & Cooperate Coupon ไม่สามารถ ลดซ้อนได้ หลังจากที่เป็น Member Price
    ' สินค้าจะมีลักษณะ เป็นป้ายเหลือง
    'ปัจจุบันให้มีการเปลี่ยนเป็น Double Policy แทน
    '--------------------------------------------------------------------------------------
'    If bVB_CNMemberPrice Then
'        If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" And tStaMemPriceAlw = "Y" And tVB_StaMemPriceAlw = "Y" Then
'            'cSourceAmt = 0
'                Select Case C_CHKnDDPMPrice() 'Check Double Policy กับ Member Price
'                    Case 0 'ไม่อนุญาต
'                        'คิดเฉพาะส่วนที่เหลือ
'                        cSourceAmt = 0
'                    Case 1 'อนุญาต
'                        cSourceAmt = cFCSdtSaleAmt 'cFCSdtMPriceAmt * cFCSdtQty *KT 52-05-04
'                    Case 2 'หาผลที่ดีที่สุด
'                        'รอดำเนินการ
'                End Select
'        End If
'    End If
    '--------------------------------------------------------------------------------------
    
    If cSourceAmt > 0 Then
        If Format(dSkuCoDisEndDate, "YYYY/MM/DD") >= Format(SP_GETdChkDateDisc, "YYYY/MM/DD") And tSkuStaCoDisAlw = "Y" Then '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
             cDisPer = SP_FEDvChkNumeric(porsTemp, "FCSkuCoDisRate")
        Else
            cDisPer = 0
        End If
        nC_ENLnDisType = nEN_LnDis1Per
        cC_DiscPer = cDisPer
        cC_DiscAmt = 0
        cC_DiscNew = 0
        cC_DiscPerOrg = cC_DiscPer
        cC_DiscAmtOrg = cC_DiscAmt
        cC_DiscNewOrg = cC_DiscNew
         tEventID = SP_FEDtChkString(porsTemp, "FTSkuEventCoDisCpn")
         '*KT 52-05-07 หาส่วนลดที่เป็นต่อ 1 หน่วยก่อนแล้วค่อยคูณ qty ที่หลังถึงจะถูกต้อง
        cDisAmt = SP_FMTcCurPoint(SP_FMTcCurPoint(cSourceAmt / cFCSdtQty, True, False) * cDisPer / 100, True, False)
        cDisAmt = cDisAmt * cFCSdtQty
        If cDisAmt > 0 Then
            bTran = False
            ocnAll.BeginTrans
            bTran = True
            nVB_CNCountBegin = nVB_CNCountBegin + 1
            bC_FlagInsCD = False
            Call C_DATxInsertDiscCD(cDisAmt, nC_DTSeqNo, tEventID, cFCSdtQtyAllDisAvi, cSourceAmt, 0, , , , bC_FlagInsCD)
            If bC_FlagInsCD Then
                C_GETcDiscCoopDisc = cDisAmt
            Else
                C_GETcDiscCoopDisc = 0
                If bTran Then
                    ocnAll.RollbackTrans
                    bTran = False
                    nVB_CNCountBegin = nVB_CNCountBegin - 1
                End If
                Call SP_DATxRsNothing(orsTemp)
                Exit Function
            End If
        Else
            C_GETcDiscCoopDisc = 0
        End If
    Else
        C_GETcDiscCoopDisc = 0
    End If
    If bTran Then
        ocnAll.CommitTrans
        bTran = False
        nVB_CNCountBegin = nVB_CNCountBegin - 1
    End If
    Call SP_DATxRsNothing(orsTemp)
    '*Eaw 56-08-26 CommSheet TK-ISS3000-317-UAT-FAPL-93.xls ถ้า Coop Disc มากกว่า 0 ให้ Update Coop flag='Y'
    If C_GETcDiscCoopDisc > 0 Then
        Call W_DATxUpdateDT(nC_DTSeqNo, tFTSkuCode)
    End If
    Exit Function
ErrHandle:
    C_GETcDiscCoopDisc = 0
     If bTran Then
        ocnAll.RollbackTrans
        bTran = False
        nVB_CNCountBegin = nVB_CNCountBegin - 1
    End If
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_GETcDiscLnDisc)", 2)
    Call SP_DATxRsNothing(orsTemp)
     Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_GETcDiscLnDisc(ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
'    Dim cDisPer As Double
'    Dim cDisAmt As Double
    Dim orsTemp As ADODB.Recordset
    Dim tEventID As String
    Dim tFTSkuCode As String
    Dim tFTShdTransType As String
    Dim tFTSkuDepCode As String
    Dim tFTSkuBrdCode As String
    Dim cFCSdtSalePrice As Double
    Dim cFCSdtSaleAmt As Double
    Dim cFCSdtAmtDisAvi As Double
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cDisQty As Double
    Dim cFCSdtQty As Double
    Dim cSourceQty As Double
    Dim cSourceAmt As Double
    Dim tStaMemPriceAlw As String
    Dim cFCSdtMPriceAmt As Double
    Dim tMemberPriceEvent As String
    Dim cB4DiscChg As Double
    Dim bTran As Boolean
'    Dim bLineOver As Boolean
    Dim cFCLineOver As Double  '*Eaw 56-03-15
    Dim cFCSdtRegPrice As Double  '*Eaw 56-03-15
    Dim bNewPrice As Boolean '*Eaw 56-09-02
    
On Error GoTo Err_Handler
    If nC_DTSeqNo < 1 Then Exit Function
    porsTemp.Filter = "FNSdtSeqNo = " & nC_DTSeqNo
    If (porsTemp.BOF And porsTemp.EOF) Then
        Call SP_DATxRsNothing(orsTemp)
        Exit Function
    End If
'    bLineOver = False           '*Ao 2012/11/29 Line Discount- (ลดเป็นบาท, ชาร์จ เป็นบาท มากกว่า /น้อยกว่า ราคาปกติ)
    tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
    cFCSdtSalePrice = SP_FEDvChkNumeric(porsTemp, "FCSdtSalePrice")
    cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
    cFCSdtAmtDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAvi")
    cFCSdtMPriceAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice")
    cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
    cFCSdtQty = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAll")
    cB4DiscChg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChg")
    tStaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd")
    tMemberPriceEvent = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
    cFCLineOver = SP_FEDvChkNumeric(porsTemp, "FCLineOver")   '*Eaw 56-03-15

    Select Case C_CHKnDDP(nC_DTSeqNo) 'Check Double Policy
        Case 0 'ไม่อนุญาต
            'คิดเฉพาะส่วนที่เหลือ
            'If cFCSdtQtyAllDisAvi <= 0 Then Exit Function
            cSourceAmt = cFCSdtAmtDisAvi
            cSourceQty = cFCSdtQtyAllDisAvi
        Case 1 'อนุญาต
            cSourceAmt = cFCSdtSaleAmt
            cSourceQty = cFCSdtQty
        Case 2 'หาผลที่ดีที่สุด
            'รอดำเนินการ
    End Select
    '*Eaw 56-03-15
'     If cFCLineOver > 0 Then                       '*Ao 2013/04/05     CommSheet TK-ISS3000-116 00.00.03 FSBT-Ao.xlsx
'        cSourceAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtRegPrice")
'     End If
     '----------------------------------------------------------------------------------------------------------
    tSql = "SELECT FTSkuCode"
    tSql = tSql & " FROM " & tVB_CNTblSku & " WITH(NOLOCK) "
    If nC_ENDiscType = nEN_Disc05LnDisc Then
        tSql = tSql & " WHERE FTSkuStaDisAlwItem = 'Y' "
        tSql = tSql & " AND FDSkuDisAlwEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    Else
        tSql = tSql & " WHERE FTSkuStaMrkAlwItem = 'Y' "
        tSql = tSql & " AND FDSkuMrkAlwEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    End If
    tSql = tSql & " AND FTSkuCode = '" & tFTSkuCode & "'"
    
    '*KT 52-02-17 member price เอกสาร Mail วันที่  17/02/2009
    'คำถาม ในกรณีที่ป้อน SKU ที่ได้ราคา Member, ถ้ามี Discount อื่นๆ เช่น Coop.Disc จะลดได้อีกหรือไม่?
    'ถ้าลดได้ ส่วนลดอะไรบ้างที่ลดได้?, ส่วนลดอะไรที่ลดไม่ได้? และสามารถ Price Override ได้หรือไม?
    'FSBT - กรณี Cooperate Disc. & Cooperate Coupon ไม่สามารถ ลดซ้อนได้ หลังจากที่เป็น Member Price
    ' สินค้าจะมีลักษณะ เป็นป้ายเหลือง
    '--------------------------------------------------------------------------------------
'    If bVB_CNMemberPrice Then
''        If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" And tStaMemPriceAlw = "Y" And cFCSdtMPriceAmt = cB4DiscChg Then
'         If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" And tStaMemPriceAlw = "Y" And tMemberPriceEvent <> "" Then  '*KT 52-09-08 Member Price + Consignment
'            'cSourceAmt = 0
'            Select Case C_CHKnDDPMPrice() 'Check Double Policy กับ Member Price
'                Case 0 'ไม่อนุญาต
'                    'คิดเฉพาะส่วนที่เหลือ
'                    cSourceAmt = 0
'                Case 1 'อนุญาต
'                    cSourceAmt = cFCSdtSaleAmt 'cFCSdtMPriceAmt * cFCSdtQty *KT 52-05-04
'                Case 2 'หาผลที่ดีที่สุด
'                    'รอดำเนินการ
'            End Select
'        End If
'    End If
    '--------------------------------------------------------------------------------------
    
    If cSourceAmt > 0 Then
        If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then 'read master
            cC_DiscPer = 0
            cC_DiscAmt = 0
            cC_DiscNew = 0
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                Select Case nC_ENLnDisType
                    Case nEN_LnDis1Per 'Percent Discount
                        cC_DiscPerOrg = cC_DiscValue
                        cC_DiscPer = cC_DiscPerOrg
                         '*KT 52-05-07 หาส่วนลดที่เป็นต่อ 1 หน่วยก่อนแล้วค่อยคูณ qty ที่หลังถึงจะถูกต้อง
                        cC_DiscAmt = SP_FMTcCurPoint(cC_DiscValue * SP_FMTcCurPoint(cSourceAmt / cSourceQty, True, False) / 100, True, False)
                        cC_DiscAmt = cC_DiscAmt * cSourceQty
                    Case nEN_LnDis2Amt  ' Amount Discount
                        cC_DiscAmtOrg = cC_DiscValue
                        cC_DiscAmt = cC_DiscValue * cSourceQty
                    Case nEN_LnDis3NewPrice 'New Price
                        cC_DiscNewOrg = cC_DiscValue
                        cC_DiscAmt = cSourceAmt - (cC_DiscValue * cSourceQty)
                End Select
            Else
                cC_DiscPerOrg = 0
                cC_DiscAmtOrg = 0
            End If
            If cC_DiscPer > 100 Then
                Call SP_MSGnShowing(tC_Msg004, tCS_CNMsgWarn)
                Call SP_DATxRsNothing(orsTemp)
                Exit Function
            End If
            If cSourceAmt - cC_DiscAmt < 0 Then
            '*Eaw 56-03-05
'                If SP_MSGnShowing(tC_Msg009, tCS_CNMsgYesNo) = vbYes Then
'                    cC_DiscAmt = cSourceAmt
'                Else
'                    Call SP_DATxRsNothing(orsTemp)
'                    Exit Function
'                End If
                Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
                Call SP_DATxRsNothing(orsTemp)
                Exit Function
            '------------------------------------------------------------------------------
            ElseIf cC_DiscAmt < 0 Then
                    '___________________________________________________'*Ao 2012/11/29 Line Discount- (ลดเป็นบาท, ชาร์จ เป็นบาท มากกว่า /น้อยกว่า ราคาปกติ)
                    If nC_ENLnDisType = nEN_LnDis3NewPrice Then
                        If cC_DiscNewOrg * cSourceQty > cSourceAmt Then      'If cC_DiscNewOrg  > cSourceAmt Then     '*Ao 2013/05/08  CommSheet TK-ISS3000-178-ST 00 00 01.xls
                            'UPDATE FCLineOver
                            Call wMain.W_PRCxUpdateLineOver(cC_DiscNewOrg, cSourceAmt, nC_DTSeqNo)
'                            cC_DiscAmt = 0              '*Eaw 56-03-05
                            cC_DiscAmt = Abs(cC_DiscAmt)    '*Eaw 56-03-05
'                            bLineOver = True
                        Else
                            Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
                            Call SP_DATxRsNothing(orsTemp)
                        End If
                    Else
                        Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
                        Call SP_DATxRsNothing(orsTemp)
                    End If
'                    Exit Function      '*Eaw 56-03-05
                    '______________________________________________________
'                Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
'                Call SP_DATxRsNothing(orsTemp)
            End If
            tEventID = ""
            If cC_DiscAmt > 0 Then
'*KT 51-06-03 ใส่เครื่องหมาย * หน้าสินค้าที่ไม่ยอมให้ทำ Double Policy กับ Mall Card *KT 52-02-19 PH2-ML-SF-102 โปรแกรมไม่แสดง *  ในกรณีที่ สินค้ามีส่วนลดแล้วไม่สามารถใช้ Mall Card ได้  (EX  Set Double Discount
'Line Dis - > Mall Card = NO) แต่เมื่อ Cancel สินค้านั้น จะมี * ขึ้นมา
'--------------------------------------------------------
'                Select Case C_CHKnDDPMallCard(nC_DTSeqNo)
'                    Case 0 'Not Allow
'                        tSql = "UPDATE " & tVB_CNTblDT
'                        tSql = tSql & vbCrLf & " SET FTSkuStaMallEmpCard='N'"
'                        tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                        tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
'                        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                        tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nC_DTSeqNo
'                        Call SP_SQLbExecute(tSql)
'                    Case 1 'allow
'                        'none
'                    Case 2 'Thebest
'                        'none
'                End Select
'--------------------------------------------------------
                bTran = False
                ocnAll.BeginTrans
                bTran = True
                nVB_CNCountBegin = nVB_CNCountBegin + 1
                bC_FlagInsCD = False
                 '*Eaw 56-03-05
                If nC_ENLnDisType = nEN_LnDis3NewPrice Then
                    If cC_DiscNewOrg * cSourceQty > cSourceAmt Then
                        cC_DiscAmt = 0
                        bNewPrice = True '*Eaw 56-09-02
                    End If
                End If
                '-----------------------------------------------------
                Call C_DATxInsertDiscCD(cC_DiscAmt, nC_DTSeqNo, tEventID, cFCSdtQtyAllDisAvi, cSourceAmt, 0, , , , bC_FlagInsCD)
                 If bC_FlagInsCD Then
                    C_GETcDiscLnDisc = cC_DiscAmt
                Else
                    If bTran Then
                        ocnAll.RollbackTrans
                        bTran = False
                        nVB_CNCountBegin = nVB_CNCountBegin - 1
                    End If
                    C_GETcDiscLnDisc = 0
                    Call SP_DATxRsNothing(orsTemp)
                    Exit Function
                End If
            Else
                C_GETcDiscLnDisc = 0
            End If
        Else
            C_GETcDiscLnDisc = 0
        End If
    Else
        C_GETcDiscLnDisc = 0
    End If
     If bTran Then
        ocnAll.CommitTrans
        bTran = False
        nVB_CNCountBegin = nVB_CNCountBegin - 1
    End If
    Call SP_DATxRsNothing(orsTemp)
    '*Eaw 56-09-02 CommSheet TK-ISS3000-329-UAT.xls ถ้า Line มากกว่า 0 ให้ Update FTSdtTaxInv='Y'
    If C_GETcDiscLnDisc > 0 Or bNewPrice Then
        Call W_DATxUpdateDT(nC_DTSeqNo, tFTSkuCode)
    End If
    Exit Function
Err_Handler:
    C_GETcDiscLnDisc = 0
    If bTran Then
        ocnAll.RollbackTrans
        bTran = False
        nVB_CNCountBegin = nVB_CNCountBegin - 1
    End If
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_GETcDiscLnDisc)", 2)
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function
            
Private Function C_CHKbCoopFromMaster(ByVal pcDisRate#, ByRef pbLimited As Boolean, ByRef pcLimitAmt As Double) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       # Find limit flag and lmit value of Cooperate Coupon (Brand)
'ARGUMENT LIST # pcDisRate = Cooperate Coupon (Brand) Rate
'                                         pbLimited = Ret --> True if Check Limit, False if Do Not Check Limit
'                                         pcLimitAmt = Ret --> Limit Value
'
'RETURN VALUE  # True if Allow ,  False if Not Allow
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    
    On Error GoTo Err_Handler
    If nC_ENDiscType = nEN_Disc09CoopCpn Then
        tSql = "SELECT FTCpcStaExpDate AS StaExp, FDCpcExpDate AS ExpDate, FTCpcStaLimitAmt AS StaLimit, FCCpcLimitAmt AS LimitAmt"
        tSql = tSql & vbCrLf & " FROM TCNMCoopCpMtn WITH(NOLOCK)  "
        tSql = tSql & vbCrLf & " WHERE FCCpcRate = " & pcDisRate
    Else
        tSql = "SELECT FTCpbStaExpDate AS StaExp, FDCpbExpDate AS ExpDate, FTCpbStaLimitAmt AS StaLimit, FCCpbLimitAmt AS LimitAmt"
        tSql = tSql & vbCrLf & " FROM TCNMCoopBrnMtn WITH(NOLOCK) "
        tSql = tSql & vbCrLf & " WHERE FCCpbRate = " & pcDisRate
    End If
    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then 'read master
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            If UCase(SP_FEDtChkString(orsTemp, "StaExp")) = "Y" Then
                If SP_FEDdChkDate(orsTemp, "ExpDate") < Format(SP_GETdChkDateDisc, "YYYY/MM/DD") Then  '*KT 51-10-17 ตรวจสอบวันที่หมดอายุการใช้ Co.Cpn '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
                    Call SP_MSGnShowing(tC_Msg008, tCS_CNMsgWarn)
                    C_CHKbCoopFromMaster = False
                    Call SP_DATxRsNothing(orsTemp)
                    Exit Function
                End If
            End If
            If UCase(SP_FEDtChkString(orsTemp, "StaLimit")) = "Y" Then
                pcLimitAmt = SP_FEDvChkNumeric(orsTemp, "LimitAmt")
                pbLimited = True
            End If
            C_CHKbCoopFromMaster = True
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    C_CHKbCoopFromMaster = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function
            
Private Function C_MCPcDiscFull(ByVal porsTemp As ADODB.Recordset, ByVal ptRefField As String) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cCal As Double
    Dim tEventID As String
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cAlwSource As Double
    Dim cC_DiscAmt As Double
    Dim cSumDiscount As Double
    Dim nFNSdtSeqNo As Integer
    With porsTemp
        .MoveFirst
        Do While Not .EOF 'ดำเนินการ
            nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
            If C_CHKnDDP(nFNSdtSeqNo) = 1 Then
                cCal = SP_FEDvChkNumeric(porsTemp, ptRefField)
                tEventID = ""
                cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
                cAlwSource = cCal
                cC_DiscAmt = cCal
                Call C_DATxInsertDiscCD(cC_DiscAmt, nFNSdtSeqNo, tEventID, cFCSdtQtyAllDisAvi, cAlwSource, 0, , , , bC_FlagInsCD)
                cSumDiscount = cSumDiscount + IIf(bC_FlagInsCD, cC_DiscAmt, 0) '*KT 52-01-08 ทำการ set ค่า Flag ที่บอกว่ามีการบันทึกข้อมูล CD
            End If
            .MoveNext
        Loop
        C_MCPcDiscFull = cSumDiscount
    End With
    Exit Function
Err_Handler:
     Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_MCPcDiscFull)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function
            
Private Function C_MCPcDiscProRate(ByVal porsTemp As ADODB.Recordset, ByVal ptRefField As String, ByVal pcAlwCal As Double, ByVal pcDiscAlw As Double) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cCal As Double
    Dim tEventID As String
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cAlwSource As Double
    Dim cC_DiscAmt As Double
    Dim cSumDiscount As Double
    Dim nFNSdtSeqNo As Integer
    Dim cSumDiscAmt As Double
    cSumDiscAmt = 0
    With porsTemp
        .MoveFirst
        Do While Not .EOF 'ดำเนินการ
            nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
            If C_CHKnDDP(nFNSdtSeqNo) = 1 Then
                cCal = SP_FEDvChkNumeric(porsTemp, ptRefField)
                tEventID = ""
                cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
                'cAlwSource = cCal '*KT 51-10-01
                cC_DiscAmt = SP_FMTcCurPoint(cCal / pcAlwCal * pcDiscAlw, True, False)
                
                '*KT 51-09-30 ทำการ Pro Rate ส่วนลดที่เป็น Mall Cpn
                '----------------------------------------------------------------------------------------
                .MoveNext
                If .EOF Then  'ดำเนินการ
                    If cSumDiscAmt + cC_DiscAmt <> pcDiscAlw Then
                        cC_DiscAmt = pcDiscAlw - cSumDiscAmt
                    End If
                End If
                cSumDiscAmt = cSumDiscAmt + cC_DiscAmt
                .MovePrevious
                '------------------------------------------------------------------------------------------
                cAlwSource = cC_DiscAmt '*KT 51-10-01
                Call C_DATxInsertDiscCD(cC_DiscAmt, nFNSdtSeqNo, tEventID, cFCSdtQtyAllDisAvi, cAlwSource, 0, , , , bC_FlagInsCD)
                cSumDiscount = cSumDiscount + IIf(bC_FlagInsCD, cC_DiscAmt, 0) '*KT 52-01-08 ทำการ set ค่า Flag ที่บอกว่ามีการบันทึกข้อมูล CD
            End If
            .MoveNext
        Loop
        C_MCPcDiscProRate = cSumDiscount
    End With
    Exit Function
Err_Handler:
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_MCPcDiscProRate)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function
            
Private Function C_GETcDiscMallCpn(ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 52-10-06
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 52-03-02 member price
'Mall Coupon ต้องทำการ Prorate ให้กับสินค้าป้ายเหลืองก่อนแล้วค่อย  Prorate สินค้าปกติ ซึ่งสินค้าที่ได้ Member Price ก้อเป็นสินค้าป้ายเหลือง
'และในการ prorate ใช้ FCSdtYellowPrice แทนตัวแปรอื่น
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tFilter As String
    Dim cYellowPrice As Double
    Dim cBluePrice As Double
    Dim cCurrentCpn As Double
    Dim cSumDiscAll As Double
    Dim orsTarget As ADODB.Recordset
    Dim cFCSdtSaleAmt As Double
    Dim cSourceAmt As Double
    Dim nI As Long
    
    nC_ENLnDisType = nEN_LnDis2Amt
    cC_DiscPer = 0
    cC_DiscAmt = cC_DiscValue
    cC_DiscNew = 0
    cC_DiscPerOrg = cC_DiscPer
    cC_DiscAmtOrg = cC_DiscAmt
    cC_DiscNewOrg = cC_DiscNew
    cCurrentCpn = cC_DiscValue
    
    tFilter = "FTSkuStaMallCpn = 'Y'"
    porsTemp.Filter = tFilter
    
    Set orsTarget = New ADODB.Recordset
    For nI = 1 To porsTemp.Fields.Count
        Call orsTarget.Fields.Append(porsTemp.Fields(nI - 1).Name, porsTemp.Fields(nI - 1).Type, porsTemp.Fields(nI - 1).DefinedSize)
    Next nI
    orsTarget.CursorType = adOpenStatic
    orsTarget.Open
    
    Do While Not porsTemp.EOF
        cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
        cSourceAmt = cFCSdtSaleAmt
        If cSourceAmt > 0 Then
            cYellowPrice = cYellowPrice + SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
            cBluePrice = cBluePrice + SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
            orsTarget.AddNew
            For nI = 1 To porsTemp.Fields.Count
                Select Case UCase(Mid(porsTemp.Fields(nI - 1).Name, 2, 1))
                    Case "T"
                        orsTarget.Fields(nI - 1).Value = porsTemp.Fields(nI - 1).Value & vbNullString
                    Case "C"
                        orsTarget.Fields(nI - 1).Value = SP_VALcText2Double(porsTemp.Fields(nI - 1).Value & vbNullString)
                    Case "N"
                        orsTarget.Fields(nI - 1).Value = SP_VALnText2Double(porsTemp.Fields(nI - 1).Value & vbNullString)
                    Case "D"
                        If IsNull(porsTemp.Fields(nI - 1).Value) Then
                            orsTarget.Fields(nI - 1).Value = "Null"
                        Else
                           orsTarget.Fields(nI - 1).Value = porsTemp.Fields(nI - 1).Value
                        End If
                End Select
            Next nI
        End If
        porsTemp.MoveNext
    Loop
    
    If DISbChkMallCpnMinAlw(cYellowPrice + cBluePrice, bC_MallCpnInStoreBar) Then '*KT 54-07-28 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon
        If cYellowPrice >= cCurrentCpn Then
            If cYellowPrice > cCurrentCpn Then            'Disc = Yellow Prorate
                orsTarget.Filter = "FCSdtYellowPrice  > 0 AND " & tFilter
                If Not (orsTarget.BOF And orsTarget.EOF) Then
                    cSumDiscAll = cSumDiscAll + C_MCPcDiscProRate(orsTarget, "FCSdtYellowPrice", cYellowPrice, cCurrentCpn)
                    '*KT 52-04-22 กรณีที่ยังไม่หมด
                    '----------------------------------------------------------------------------------
                    cCurrentCpn = cCurrentCpn - cSumDiscAll
                    If cCurrentCpn > 0 Then
                            If cBluePrice >= cCurrentCpn Then
                                If cBluePrice > cCurrentCpn Then                'Disc = Blue Prorate
                                    orsTarget.Filter = "FCSdtBluePrice  > 0 AND " & tFilter
                                    If Not (orsTarget.BOF And orsTarget.EOF) Then
                                        cSumDiscAll = cSumDiscAll + C_MCPcDiscProRate(orsTarget, "FCSdtBluePrice", cBluePrice, cCurrentCpn)
                                    End If
                                Else                'Disc = Blue
                                    orsTarget.Filter = "FCSdtBluePrice  > 0 AND " & tFilter
                                    If Not (orsTarget.BOF And orsTarget.EOF) Then
                                        cSumDiscAll = cSumDiscAll + C_MCPcDiscFull(orsTarget, "FCSdtBluePrice")
                                    End If
                                End If
                            Else            'Disc = Blue
                                orsTarget.Filter = "FCSdtBluePrice  > 0 AND " & tFilter
                                If Not (orsTarget.BOF And orsTarget.EOF) Then
                                    cSumDiscAll = cSumDiscAll + C_MCPcDiscFull(orsTarget, "FCSdtBluePrice")
                                End If
                            End If
                    End If
                    '---------------------------------------------------------------------------------
                End If
            Else     'Disc = Yellow
                orsTarget.Filter = "FCSdtYellowPrice  > 0 AND " & tFilter
                If Not (orsTarget.BOF And orsTarget.EOF) Then
                    cSumDiscAll = cSumDiscAll + C_MCPcDiscFull(orsTarget, "FCSdtYellowPrice")
                    '*KT 52-04-22 กรณีที่ยังไม่หมด
                    '----------------------------------------------------------------------------------
                    cCurrentCpn = cCurrentCpn - cSumDiscAll
                    If cCurrentCpn > 0 Then
                            If cBluePrice >= cCurrentCpn Then
                                If cBluePrice > cCurrentCpn Then                'Disc = Blue Prorate
                                    orsTarget.Filter = "FCSdtBluePrice  > 0 AND " & tFilter
                                    If Not (orsTarget.BOF And orsTarget.EOF) Then
                                        cSumDiscAll = cSumDiscAll + C_MCPcDiscProRate(orsTarget, "FCSdtBluePrice", cBluePrice, cCurrentCpn)
                                    End If
                                Else                'Disc = Blue
                                    orsTarget.Filter = "FCSdtBluePrice  > 0 AND " & tFilter
                                    If Not (orsTarget.BOF And orsTarget.EOF) Then
                                        cSumDiscAll = cSumDiscAll + C_MCPcDiscFull(orsTarget, "FCSdtBluePrice")
                                    End If
                                End If
                            Else            'Disc = Blue
                                orsTarget.Filter = "FCSdtBluePrice  > 0 AND " & tFilter
                                If Not (orsTarget.BOF And orsTarget.EOF) Then
                                    cSumDiscAll = cSumDiscAll + C_MCPcDiscFull(orsTarget, "FCSdtBluePrice")
                                End If
                            End If
                    End If
                    '---------------------------------------------------------------------------------
                End If
            End If
        Else        'Disc = Yellow
            orsTarget.Filter = "FCSdtYellowPrice  > 0 AND " & tFilter
            If Not (orsTarget.BOF And orsTarget.EOF) Then
                cSumDiscAll = cSumDiscAll + C_MCPcDiscFull(orsTarget, "FCSdtYellowPrice")
                cCurrentCpn = cCurrentCpn - cYellowPrice
            End If
            If cBluePrice >= cCurrentCpn Then
                If cBluePrice > cCurrentCpn Then                'Disc = Blue Prorate
                    orsTarget.Filter = "FCSdtBluePrice  > 0 AND " & tFilter
                    If Not (orsTarget.BOF And orsTarget.EOF) Then
                        cSumDiscAll = cSumDiscAll + C_MCPcDiscProRate(orsTarget, "FCSdtBluePrice", cBluePrice, cCurrentCpn)
                    End If
                Else                'Disc = Blue
                    orsTarget.Filter = "FCSdtBluePrice  > 0 AND " & tFilter
                    If Not (orsTarget.BOF And orsTarget.EOF) Then
                        cSumDiscAll = cSumDiscAll + C_MCPcDiscFull(orsTarget, "FCSdtBluePrice")
                    End If
                End If
            Else            'Disc = Blue
                orsTarget.Filter = "FCSdtBluePrice  > 0 AND " & tFilter
                If Not (orsTarget.BOF And orsTarget.EOF) Then
                    cSumDiscAll = cSumDiscAll + C_MCPcDiscFull(orsTarget, "FCSdtBluePrice")
                End If
            End If
        End If
        C_GETcDiscMallCpn = cSumDiscAll
    Else
        Call SP_MSGnShowing("tms_CN005225", nCS_Warning)
        C_GETcDiscMallCpn = 0
    End If
    Call SP_DATxRsNothing(orsTarget)
    Exit Function
Err_Handler:
    C_GETcDiscMallCpn = 0
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_GETcDiscMallCpn)", 2)
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTarget)
End Function

Private Sub C_MALxInsMall(ByVal pcDisc As Double, ByVal pnSeqNo As Integer, ByVal ptEventID As String, ByVal pcQtyAvi As Double, ByVal pcAlwCal As Double, ByVal pcDiscPer As Double, ByRef pcDisSumAmt As Double, Optional ByVal pcMallCrdRate As Double = 0, Optional ByVal pcDiscMallCrdHD As Double = 0, Optional ByVal pbSeqNoLast As Boolean = False)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tEventNo As String
    Dim cDicsTemp As Double
On Error GoTo ErrHandle
    cDicsTemp = pcDisSumAmt
    cC_DiscAmt = pcDisc
    tEventNo = ptEventID
    If cC_DiscAmt > 0 Then
        bW_ChkMallCard = False
        'MsgBox "STEP1:" & pcDisSumAmt & "+" & cC_DiscAmt
        pcDisSumAmt = pcDisSumAmt + cC_DiscAmt
        '*KT 51-11-04 PRORATE
        If pbSeqNoLast Then
            'MsgBox "STEP2:"
'            If pcDisSumAmt > pcDiscMallCrdHD Then                                      '*Ao 2011/03/21        ถ้าไม่คลุมด้วย Function ทำให้ที่เป็นจุดทศนิยมเปรียบเทียบกันผิด เช่น    36.36  > 36.36  จะเป็น True
            If SP_FMTcCurPoint(pcDisSumAmt) > SP_FMTcCurPoint(pcDiscMallCrdHD) Then      '*Ao 2011/03/21  Communication Sheet PH3-ML-SB-134 Mall Card  discount amount minus in last item.xls
                'MsgBox "STEP3:"
                cC_DiscAmt = cC_DiscAmt - (pcDisSumAmt - pcDiscMallCrdHD)
                 'MsgBox "STEP4:"
                pcDisSumAmt = pcDiscMallCrdHD
            ElseIf SP_FMTcCurPoint(pcDisSumAmt) < SP_FMTcCurPoint(pcDiscMallCrdHD) Then
                'MsgBox "STEP5:"
                cC_DiscAmt = cC_DiscAmt + (pcDiscMallCrdHD - pcDisSumAmt)
                 'MsgBox "STEP6:"
                pcDisSumAmt = pcDiscMallCrdHD
            End If
        End If
        cC_DiscPer = pcDiscPer
        bC_FlagInsCD = False '*KT 52-03-17 PH1.5-ML-SF-120
        'MsgBox "STEP7:" & cC_DiscAmt
        Call C_DATxInsertDiscCD(cC_DiscAmt, pnSeqNo, tEventNo, pcQtyAvi, pcAlwCal, 0, , , pcDiscPer, bC_FlagInsCD)  '*KT 52-01-16 ถ้า Insert ไม่ผ่านต้อง set ค่ากลับเหมือนเดิม
        If Not bC_FlagInsCD Then pcDisSumAmt = cDicsTemp '*KT 52-01-16 ถ้า Insert ไม่ผ่านต้อง set ค่ากลับเหมือนเดิม
    End If
    Exit Sub
ErrHandle:
    'Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, tVB_CNWheBill & "(C_MALxInsMall)", 2) '*KT 52-03-17 PH1.5-ML-SF-120
    Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[7]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
End Sub

Private Function C_MALcSpec() As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    
    tSql = "SELECT SUM(FCSrcAmt) AS FCSrcAmt"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblRC & " WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE FTTdmCode IN ('T005','T007','T014')" '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher"
    tSql = tSql & vbCrLf & "AND  FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    If SP_SQLbExecute(tSql, orsTemp) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            C_MALcSpec = SP_FEDvChkNumeric(orsTemp, "FCSrcAmt")
        End If
    End If
    If bC_StaRecal Then
        C_MALcSpec = C_MALcSpec + cC_SpecialPrice
    End If
End Function

Private Function C_GETcDiscMallCard(ByVal pbMallCard As Boolean, ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
' *KT 52-10-27 PH2-ML-SF-160 Mall Card ไม่คำนวนส่วนลดในกรณีที่ Dept Code เดียวกันแต่ Brand Code ต่างกัน (Allow Dept & SKU)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String

    Dim cDisSumAmt As Double
    
    Dim nFNSdtSeqNo As Integer
    Dim tFTSkuCode As String
    Dim tFTSkuDepCode As String
    Dim tFTSkuBrdCode As String
    Dim cFCSdtQtyAllDisAvi As Double
    
    Dim cDptRate As Double
    Dim dCurrentExpDate As Date
    
    Dim aEmpCard() As UDT_EmpCard
    Dim nI As Integer
    Dim cBluePrice As Double
    Dim cYellowPrice As Double

    Dim cSumCalBalance As Double
    Dim cSumLimit As Double
    Dim tStaChkLimit As String
    Dim tStaChkAllowBuy As String
    Dim tEmpNoFrm As String
    Dim tEmpNoTo As String
    Dim bOnline As Boolean
    Dim tWhere As String
    Dim cSkuRate As Double
    '*TW 51-08-15
    Dim cSumSpecial As Double
    Dim cSumSpecialBlue As Double
    Dim cSumSpecialReg As Double
    Dim cSumSpecialRegOrg As Double
    Dim bRate As Boolean
    Dim nMall  As Long
    '*KT 51-09-25
    Dim cFTSkuMallCrdRate As Double
    Dim cDiscHD As Double
    Dim bSeqLast  As Boolean
    '*KT 52-02-18 mall card discount dpt/sku allowed
    Dim orsMallSKUAlw As ADODB.Recordset
    Dim orsMallDptAlw As ADODB.Recordset
    Dim orsSum  As ADODB.Recordset
    Dim bMallSKUAlw As Boolean
    Dim bMallDptAlw As Boolean
    Dim bMallAlw  As Boolean
    Dim cBlueYellowAmt As Double
    Dim tSKUBlue As String
    Dim cTempSum As Double
    Dim cTempQty  As Double
    Dim cTempRegSum As Double
    Dim cTempProSum As Double
    Dim cTotalAmt As Double
    Dim cChkTotalAmt As Double
    Dim nStep As Long
    Dim cTotalAmt2  As Double
    Dim nSeqNoLast As Long
    Dim cSumArray As Double
    Dim nJ As Long
    '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
    Dim tSkuPointAlw As String
    Dim tSkuPointAlwDiscountMall As String
    Dim tSkuStaMallEmpCard As String
    Dim nEff As Long
    '----------------------------
    Dim bYellow As Boolean  '*KT 54-11-11 V.4.0.4 DCR_5411002
    Dim cSaleAmt As Double '*KT 55-06-27 CommSheet ML-PH4-R3-026 Sometine Mall Card not apply after Mall Coupon
    
On Error GoTo Err_Handler
    bYellow = False  '*KT 54-11-11 V.4.0.4 DCR_5411002
    cTotalAmt = 0
    cChkTotalAmt = 0
    bOnline = Not (ocnMember Is Nothing)
    If bOnline = False Then
        '*KT 52-05-26 PH 1.8 ML-SF-075 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ได้ ขอให้ POS แสดง Warning Message เพียงครั้งเดียว
       ' If ocnAll.ConnectionString <> oDB.ocnOffLine.ConnectionString And bC_ReConnectMember Then 'PH 2.0.0 OLD
       If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) And bC_ReConnectMember Then
            If SP_CONbMemberServer(, , False) Then
                bOnline = True
            End If
        End If
    End If
    
    If UCase(tC_StaChkExpire) = "Y" Then
        dCurrentExpDate = DateAdd("m", 1, dC_CardExpDate) - 1
'        If dCurrentExpDate < Format(dVB_CurrentDate, "mm/dd/yyyy") Then Exit Function '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  ใส่ Format mm/dd/yyyyy เพราะ dVB_CurrentDate มีเวลาเข้ามาด้วย Communication Sheet PH2-ML-SF-027 Credit Expire.xls
        '*Ao 52-02-05  เนื่องจาก dCurrentExpDate Format เป็น dd/mm/yyyy  ทำให้การเปรียบเทียบผิด
        'เช่น วันนี้ 05/02/2009 -> Format(05/02/2009, "mm/dd/yyyy")=02/05/2009
        '01/04/2009 <  02/05/2009
'        If CDate(Format(dCurrentExpDate, "YYYY/MM/DD")) < CDate(Format(SP_GETdChkDateDisc, "YYYY/MM/DD")) Then Exit Function  '*KT 52-06-04 ตรวจสอบวันที่หมดอายุของ point ตามที่กำหนดใน parameter
         If CDate(Format(dCurrentExpDate, "YYYY/MM/DD")) < CDate(Format(dVB_CurrentDate, "YYYY/MM/DD")) Then
            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[1]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
            
            Exit Function    'ทาง FSBT ให้ทำการตรวจสอบวันที่ Expire Date ตาม System Date
        End If
    End If
   
    nC_ENLnDisType = nEN_LnDis1Per
    cC_DiscPer = cC_DiscValue
    cC_DiscAmt = 0
    cC_DiscNew = 0
    cC_DiscPerOrg = cC_DiscPer
    cC_DiscAmtOrg = cC_DiscAmt
    cC_DiscNewOrg = cC_DiscNew
    
    
    ReDim aW_MallCard(0)    'clear array
    aW_MallCard(0).MALLAMT = 0
    aW_MallCard(0).MALLRATE = 0
    If Not (porsTemp.BOF And porsTemp.EOF) Then
        porsTemp.Sort = "FCSdtSaleAmt ASC"           '*Ao 2011/03/21  Communication Sheet PH3-ML-SB-134 Mall Card  discount amount minus in last item.xls
        porsTemp.MoveFirst
    End If
    
    
'    porsTemp.MoveFirst
    ReDim aEmpCard(0) As UDT_EmpCard
    nI = 0
    '*TW 51-08-15
    cSumSpecialRegOrg = 1
    cSumSpecialReg = 1
    cSumSpecial = C_MALcSpec
    If cSumSpecial > 0 Then
        cSumSpecialBlue = 0
        cSumSpecialRegOrg = 0
        porsTemp.MoveFirst
        Do While Not porsTemp.EOF
            cBluePrice = SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
            cYellowPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
            cSumSpecialRegOrg = SP_FMTcCurPoint(cSumSpecialRegOrg + cBluePrice, True)
            cSumSpecialBlue = SP_FMTcCurPoint(cSumSpecialBlue + SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt") - cYellowPrice, True)
            porsTemp.MoveNext
        Loop
        cSumSpecialReg = cSumSpecialRegOrg - cSumSpecial
    End If
    
    '*KT 52-02-18 mall card discount dpt/sku allowed
    If UCase(tC_StaChkDepart) = "Y" Then
        tSql = "SELECT     MM.FTMcdCode, MK.FTSkuCode, MK.FTMcsAlw,MK.FDMcsStartDate,MK.FDMcsEndDate,MK.FCMcsRate,MK.FCMcsAmtMin,MK.FCMcsAmtMax"
        tSql = tSql & vbCrLf & "FROM  TCNMMallCardMtn MM WITH(NOLOCK) INNER JOIN"
        tSql = tSql & vbCrLf & "TCNMMallCardSKU MK WITH(NOLOCK) ON MM.FTMcdCode = MK.FTMcdCode"
        tSql = tSql & vbCrLf & "WHERE MM.FTMcdCode = '" & tVB_DiscMallCardKey & "'"
        tSql = tSql & vbCrLf & "AND MK.FTMcsAlwPoint = '0'"
        Call SP_SQLvExecute(tSql, orsMallSKUAlw, oDB.ocnOnLine, , , , , True, True) 'read master
        If Not (orsMallSKUAlw.BOF Or orsMallSKUAlw.EOF) Then
            bMallSKUAlw = True
        Else
            bMallSKUAlw = False
        End If
        tSql = "SELECT     MM.FTMcdCode, MD.FTMpcCode,MD.FTMpcBrdCode,MD.FTMpcAlw,MD.FDMpcStartDate,MD.FDMpcEndDate,MD.FCMpcRate,MD.FTMpcReg,MD.FTMpcPro,MD.FCMpcAmtMin,MD.FCMpcAmtMax"
        tSql = tSql & vbCrLf & "FROM    TCNMMallCardMtn MM WITH(NOLOCK) INNER JOIN"
        tSql = tSql & vbCrLf & "TCNMMallCardDep MD WITH(NOLOCK) ON MM.FTMcdCode = MD.FTMcdCode"
        tSql = tSql & vbCrLf & "WHERE MM.FTMcdCode = '" & tVB_DiscMallCardKey & "'"
        tSql = tSql & vbCrLf & "AND MD.FTMpcAlwPoint = '0'"
        Call SP_SQLvExecute(tSql, orsMallDptAlw, oDB.ocnOnLine, , , , , True, True) 'read master
        If Not (orsMallDptAlw.BOF Or orsMallDptAlw.EOF) Then
            bMallDptAlw = True
        Else
            bMallDptAlw = False
        End If
    End If
    '---------------------------------------
    
    cTotalAmt2 = 0
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
           cBluePrice = SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
           cYellowPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
           cSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt") '*KT 55-06-27 CommSheet ML-PH4-R3-026 Sometine Mall Card not apply after Mall Coupon ผลกระทบจาก Req.13 Phase 4 ทำให้ขอมูล cBluePrice,cYellowPrice ไม่สามารถนำมารวมเป็น SubTotal-All
            If tC_MallFrmAmt = "Y" Then
                'cTotalAmt2 = cTotalAmt2 + CDbl(SP_FMTcCurPoint(cBluePrice + cYellowPrice, True, False))
                cTotalAmt2 = cTotalAmt2 + CDbl(SP_FMTcCurPoint(cSaleAmt, True, False)) '*KT 55-06-27 CommSheet ML-PH4-R3-026 Sometine Mall Card not apply after Mall Coupon ผลกระทบจาก Req.13 Phase 4 ทำให้ขอมูล cBluePrice,cYellowPrice ไม่สามารถนำมารวมเป็น SubTotal-All
            Else
                cTotalAmt2 = cTotalAmt2 + CDbl(SP_FMTcCurPoint(cBluePrice, True, False))  'CDbl(SP_FMTcCurPoint(IIf(tC_MallRegAlw = "Y", cBluePrice, 0) + IIf(tC_MallProAlw = "Y", cYellowPrice, 0), True, False))
            End If
           '*KT 54-11-11 V.4.0.4 DCR_5411002
           'หาว่ามีสินค้า SKU Point Active  = YES    , Allow Discount (SKU Point Maintenance)= NO     Mall Card/Employee Card(SKU Maintenance) = YES หรือไม่ ถ้ามีให้ bYellow = True
           If Not bYellow Then
                If SP_FEDtChkString(porsTemp, "FTSdtMemberPointAlw") = "Y" And SP_FEDtChkString(porsTemp, "FTSdtStaPdt14") <> "1" And SP_FEDtChkString(porsTemp, "FTSkuStaMallEmpCardEx") = "Y" Then
                    bYellow = True
                End If
           End If
            porsTemp.MoveNext
    Loop
    
    nStep = 1
    porsTemp.MoveFirst
    Set orsSum = porsTemp.Clone  '*KT 52-02-18 mall card discount dpt/sku allowed
    Do While Not porsTemp.EOF
        nI = nI + 1
        ReDim Preserve aEmpCard(nI) As UDT_EmpCard
        nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
        tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
        tFTSkuDepCode = SP_FEDtChkString(porsTemp, "FTSkuDepCode")
        tFTSkuBrdCode = SP_FEDtChkString(porsTemp, "FTSkuBrdCode")
        cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
        cBluePrice = SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
        cYellowPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
        cSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt") '*KT 55-06-27 CommSheet ML-PH4-R3-026 Sometine Mall Card not apply after Mall Coupon ผลกระทบจาก Req.13 Phase 4 ทำให้ขอมูล cBluePrice,cYellowPrice ไม่สามารถนำมารวมเป็น SubTotal-All
        aEmpCard(nI).udt_nSeqNo = nFNSdtSeqNo
        aEmpCard(nI).udt_tSkuCode = tFTSkuCode
        aEmpCard(nI).udt_tDptCode = tFTSkuDepCode
        aEmpCard(nI).udt_cQtyAllAvi = cFCSdtQtyAllDisAvi
        
        cFTSkuMallCrdRate = CDbl(SP_FEDvChkNumeric(porsTemp, "FCSkuMallCardRate"))
        tSKUBlue = SP_FEDtChkString(porsTemp, "FTSkuStaMallEmpCardEx")
        If tSKUBlue = "Y" Then
            cSkuRate = SP_FEDvChkNumeric(porsTemp, "FCSkuMallCardRate")
        Else
            cSkuRate = 0
        End If

        '*KT 52-02-18 mall card discount
        'udt_cB4CalAmt = MALLAMT ,udt_cCalRate = MALLRATE
        '1 not allow sku
        '2 allow sku
        '3 not allow dpt
        '4 allow dpt
        '-------------------------------------------------------------------------------------------------------------------------
        
        If UCase(tC_StaChkDepart) = "Y" Then 'ตรวจสอบระดับ Allow และ Not Allow ของ Dept Code/Sku Code
            If bMallSKUAlw Then 'ระดับ  Sku Code
                orsMallSKUAlw.Filter = adFilterNone
                orsMallSKUAlw.Filter = "FTSkuCode = '" & tFTSkuCode & "'"
                If Not (orsMallSKUAlw.BOF Or orsMallSKUAlw.EOF) Then
                    'not allow sku/allow sku
                    
                    orsSum.Filter = "FTSkuCode = '" & tFTSkuCode & "'"
                    If Not (orsSum.BOF And orsSum.EOF) Then
                        orsSum.MoveFirst
                        Call C_GETxSumSaleAmt(orsSum, cTempSum, cTempQty)
                    Else
                        cTempSum = 0
                        cTempQty = 0
                    End If
                    orsSum.Filter = adFilterNone
                    If C_EMPbGetSKURateEx(orsMallSKUAlw, cBluePrice, cYellowPrice, cTempSum, cTempQty, cDptRate, cBlueYellowAmt) Then '*KT 52-02-18 ต้องมีการแก้ไขเรื่อง allow sku,not allow sku,allow dpt,not allow dpt
                        aEmpCard(nI).udt_cCalRate = cDptRate
                        aEmpCard(nI).udt_cB4CalAmt = cBlueYellowAmt
                        aEmpCard(nI).udt_cB4UnCalAmt = 0
                        aEmpCard(nI).udt_tStaDisMallCrd = "0"
                        aEmpCard(nI).udt_tStaDisSKUOrDpt = "1" 'By SKU '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                        bMallAlw = False
                    Else
                        If bMallDptAlw Then
                            orsMallDptAlw.Filter = adFilterNone
                            orsMallDptAlw.Filter = "(FTMpcCode = '" & tFTSkuDepCode & "' AND FTMpcBrdCode ='" & tFTSkuBrdCode & "') OR (FTMpcCode = '" & tFTSkuDepCode & "' AND FTMpcBrdCode ='') OR (FTMpcCode = '' AND FTMpcBrdCode = '" & tFTSkuBrdCode & "') OR (FTMpcCode = '" & tFTSkuDepCode & "')"
                            orsMallDptAlw.Sort = "FTMpcCode DESC,FTMpcBrdCode DESC"
                            If Not (orsMallDptAlw.BOF Or orsMallDptAlw.EOF) Then
                                'not allow dpt/allow dpt
                                orsSum.Filter = "(FTSkuDepCode = '" & tFTSkuDepCode & "' AND FTSkuBrdCode ='" & tFTSkuBrdCode & "') OR (FTSkuDepCode = '" & tFTSkuDepCode & "' AND FTSkuBrdCode ='') OR (FTSkuDepCode = ''  AND FTSkuBrdCode = '" & tFTSkuBrdCode & "') OR (FTSkuDepCode = '" & tFTSkuDepCode & "')"
                                If Not (orsSum.BOF And orsSum.EOF) Then
                                    orsSum.MoveFirst
                                    Call C_GETxSumSaleAmtEx(orsSum, cTempRegSum, cTempProSum, cTempQty)
                                Else
                                    cTempRegSum = 0
                                    cTempProSum = 0
                                    cTempQty = 0
                                End If
                                orsSum.Filter = adFilterNone
                                '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                                cSumArray = 0
                                For nJ = 1 To UBound(aEmpCard)
                                    If tFTSkuDepCode = aEmpCard(nJ).udt_tDptCode And aEmpCard(nJ).udt_tStaDisSKUOrDpt = "2" Then
                                        cSumArray = cSumArray + aEmpCard(nJ).udt_cB4CalAmt
                                    End If
                                Next nJ
                                '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                                If C_EMPbGetDptRateEx(orsMallDptAlw, tSKUBlue, cBluePrice, cYellowPrice, cTempRegSum, cTempProSum, cTempQty, cSumArray, cDptRate, cBlueYellowAmt) Then '*KT 52-02-18 ต้องมีการแก้ไขเรื่อง allow sku,not allow sku,allow dpt,not allow dpt
                                    aEmpCard(nI).udt_cCalRate = cDptRate
                                    aEmpCard(nI).udt_cB4CalAmt = cBlueYellowAmt
                                    aEmpCard(nI).udt_cB4UnCalAmt = 0
                                    aEmpCard(nI).udt_tStaDisMallCrd = "0"
                                    aEmpCard(nI).udt_tStaDisSKUOrDpt = "2"  'By Dpt '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                                    bMallAlw = False
                                Else
                                    bMallAlw = True
                                End If
                            Else
                                bMallAlw = True
                            End If
                        Else
                            bMallAlw = True
                        End If
                    End If
                Else
                    If bMallDptAlw Then 'ระดับ  Dept Code
                        orsMallDptAlw.Filter = adFilterNone
                        orsMallDptAlw.Filter = "(FTMpcCode = '" & tFTSkuDepCode & "' AND FTMpcBrdCode ='" & tFTSkuBrdCode & "') OR (FTMpcCode = '" & tFTSkuDepCode & "' AND FTMpcBrdCode ='') OR (FTMpcCode = '' AND FTMpcBrdCode ='" & tFTSkuBrdCode & "') OR (FTMpcCode = '" & tFTSkuDepCode & "')"
                        orsMallDptAlw.Sort = "FTMpcCode DESC,FTMpcBrdCode DESC"
                        If Not (orsMallDptAlw.BOF Or orsMallDptAlw.EOF) Then
                            'not allow dpt/allow dpt
                            orsSum.Filter = "(FTSkuDepCode = '" & tFTSkuDepCode & "' AND FTSkuBrdCode ='" & tFTSkuBrdCode & "') OR (FTSkuDepCode = '" & tFTSkuDepCode & "' AND FTSkuBrdCode ='') OR (FTSkuDepCode = '' AND FTSkuBrdCode ='" & tFTSkuBrdCode & "') OR (FTSkuDepCode = '" & tFTSkuDepCode & "')"
                            If Not (orsSum.BOF And orsSum.EOF) Then
                                orsSum.MoveFirst
                                Call C_GETxSumSaleAmtEx(orsSum, cTempRegSum, cTempProSum, cTempQty)
                            Else
                                cTempRegSum = 0
                                cTempProSum = 0
                                cTempQty = 0
                            End If
                            orsSum.Filter = adFilterNone
                            '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                            cSumArray = 0
                            For nJ = 1 To UBound(aEmpCard)
                                If tFTSkuDepCode = aEmpCard(nJ).udt_tDptCode And aEmpCard(nJ).udt_tStaDisSKUOrDpt = "2" Then
                                    cSumArray = cSumArray + aEmpCard(nJ).udt_cB4CalAmt
                                End If
                            Next nJ
                            '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                            If C_EMPbGetDptRateEx(orsMallDptAlw, tSKUBlue, cBluePrice, cYellowPrice, cTempRegSum, cTempProSum, cTempQty, cSumArray, cDptRate, cBlueYellowAmt) Then '*KT 52-02-18 ต้องมีการแก้ไขเรื่อง allow sku,not allow sku,allow dpt,not allow dpt
                                aEmpCard(nI).udt_cCalRate = cDptRate
                                aEmpCard(nI).udt_cB4CalAmt = cBlueYellowAmt
                                aEmpCard(nI).udt_cB4UnCalAmt = 0
                                aEmpCard(nI).udt_tStaDisMallCrd = "0"
                                aEmpCard(nI).udt_tStaDisSKUOrDpt = "2"  'By Dpt '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                                bMallAlw = False
                            Else
                                bMallAlw = True
                            End If
                        Else
                            bMallAlw = True
                        End If
                    Else
                        bMallAlw = True
                    End If
                End If

            Else
                If bMallDptAlw Then
                    orsMallDptAlw.Filter = adFilterNone
                    orsMallDptAlw.Filter = "(FTMpcCode = '" & tFTSkuDepCode & "' AND FTMpcBrdCode ='" & tFTSkuBrdCode & "') OR (FTMpcCode = '" & tFTSkuDepCode & "' AND FTMpcBrdCode ='') OR (FTMpcCode = '' AND FTMpcBrdCode ='" & tFTSkuBrdCode & "') OR (FTMpcCode = '" & tFTSkuDepCode & "')"
                    orsMallDptAlw.Sort = "FTMpcCode DESC,FTMpcBrdCode DESC"
                    If Not (orsMallDptAlw.BOF Or orsMallDptAlw.EOF) Then
                        orsSum.Filter = "(FTSkuDepCode = '" & tFTSkuDepCode & "' AND  FTSkuBrdCode ='" & tFTSkuBrdCode & "') OR (FTSkuDepCode = '" & tFTSkuDepCode & "' AND FTSkuBrdCode ='') OR (FTSkuDepCode = ''  AND FTSkuBrdCode ='" & tFTSkuBrdCode & "') OR (FTSkuDepCode = '" & tFTSkuDepCode & "')"
                        If Not (orsSum.BOF And orsSum.EOF) Then
                            orsSum.MoveFirst
                            Call C_GETxSumSaleAmtEx(orsSum, cTempRegSum, cTempProSum, cTempQty)
                        Else
                            cTempRegSum = 0
                            cTempProSum = 0
                            cTempQty = 0
                        End If
                        orsSum.Filter = adFilterNone
                        '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                        cSumArray = 0
                        For nJ = 1 To UBound(aEmpCard)
                            If tFTSkuDepCode = aEmpCard(nJ).udt_tDptCode And aEmpCard(nJ).udt_tStaDisSKUOrDpt = "2" Then
                                cSumArray = cSumArray + aEmpCard(nJ).udt_cB4CalAmt
                            End If
                        Next nJ
                        '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                        If C_EMPbGetDptRateEx(orsMallDptAlw, tSKUBlue, cBluePrice, cYellowPrice, cTempRegSum, cTempProSum, cTempQty, cSumArray, cDptRate, cBlueYellowAmt) Then '*KT 52-02-18 ต้องมีการแก้ไขเรื่อง allow sku,not allow sku,allow dpt,not allow dpt
                            aEmpCard(nI).udt_cCalRate = cDptRate
                            aEmpCard(nI).udt_cB4CalAmt = cBlueYellowAmt
                            aEmpCard(nI).udt_cB4UnCalAmt = 0
                            aEmpCard(nI).udt_tStaDisMallCrd = "0"
                            aEmpCard(nI).udt_tStaDisSKUOrDpt = "2"  'By Dpt '*Ao 2010/09/07   Communication Sheet PH3-ML-SF-128 Calculate Mall card discount wrong when Total.xls
                            bMallAlw = False
                        Else
                            bMallAlw = True
                        End If
                    Else
                        bMallAlw = True
                    End If
                Else
                    bMallAlw = True
                End If
            End If
            If bMallAlw Then
                '*KT 52-03-31 PH1.8-ML-SF-017
                'กรณีที่ Mall Card ไม่อนุญาติให้ลดสินค้าป้ายเหลือง Promo Rate ที่ใช้ในการคำนวณส่วนลดเหมือนกันกับเงื่อนไขสินค้าป้ายน้ำเงินคือเปรียบเทียบว่า Rate ของ SKU กับ Mall Card ฝั่งไหนน้อยกว่า
                'ยกเว้นกรณีที่สินค้าป้ายเหลืองเกิดจากไม่มีการ Flag ให้ใช้ Mall Card ที่ SKU ให้ใช้ Rate ของสินค้าป้ายเหลืองจาก Mall Card เลย
                'If tC_MallProAlw = "Y" And SP_FEDtChkString(porsTemp, "FTSkuStaMallEmpCardEx") = "Y" Then
                If (tC_MallProAlw = "Y" And tSKUBlue = "Y") Or (tC_MallProAlw = "N" And tSKUBlue = "Y") Then
                    aEmpCard(nI).udt_cCalRate = SP_FMTcCurPoint(IIf(cC_DiscValue > cSkuRate, cSkuRate, cC_DiscValue), True, False)
                Else
                    If tC_MallProAlw = "Y" Then
                        aEmpCard(nI).udt_cCalRate = SP_FMTcCurPoint(cC_DiscValue, True, False)
                    Else
                        aEmpCard(nI).udt_cCalRate = 0
                    End If
                End If
                aEmpCard(nI).udt_tStaDisMallCrd = "1"
            End If
        Else
                '*KT 52-03-31 PH1.8-ML-SF-017
                'กรณีที่ Mall Card ไม่อนุญาติให้ลดสินค้าป้ายเหลือง Promo Rate ที่ใช้ในการคำนวณส่วนลดเหมือนกันกับเงื่อนไขสินค้าป้ายน้ำเงินคือเปรียบเทียบว่า Rate ของ SKU กับ Mall Card ฝั่งไหนน้อยกว่า
                'ยกเว้นกรณีที่สินค้าป้ายเหลืองเกิดจากไม่มีการ Flag ให้ใช้ Mall Card ที่ SKU ให้ใช้ Rate ของสินค้าป้ายเหลืองจาก Mall Card เลย
                'If tC_MallProAlw = "Y" And SP_FEDtChkString(porsTemp, "FTSkuStaMallEmpCardEx") = "Y" Then
                If (tC_MallProAlw = "Y" And tSKUBlue = "Y") Or (tC_MallProAlw = "N" And tSKUBlue = "Y") Then
                    aEmpCard(nI).udt_cCalRate = SP_FMTcCurPoint(IIf(cC_DiscValue > cSkuRate, cSkuRate, cC_DiscValue), True, False)
                Else
                    If tC_MallProAlw = "Y" Then
                        aEmpCard(nI).udt_cCalRate = SP_FMTcCurPoint(cC_DiscValue, True, False)
                    Else
                        aEmpCard(nI).udt_cCalRate = 0
                    End If
                End If
                aEmpCard(nI).udt_tStaDisMallCrd = "1"
        End If
        '---------------------------------------------------------------------------------------------------------------------------------------------------
        
        If aEmpCard(nI).udt_cCalRate = 0 Then
            'กลุ่มที่ไม่ได้ mall card discount
            aEmpCard(nI).udt_cB4CalAmt = 0
            aEmpCard(nI).udt_cB4UnCalAmt = SP_FMTcCurPoint(cYellowPrice + cBluePrice, True, False)
            If tC_MallFrmAmt = "Y" Then
                'cTotalAmt = cTotalAmt + CDbl(SP_FMTcCurPoint(cBluePrice + cYellowPrice, True, False))
                cTotalAmt = cTotalAmt + CDbl(SP_FMTcCurPoint(cSaleAmt, True, False)) '*KT 55-06-27 CommSheet ML-PH4-R3-026 Sometine Mall Card not apply after Mall Coupon ผลกระทบจาก Req.13 Phase 4 ทำให้ขอมูล cBluePrice,cYellowPrice ไม่สามารถนำมารวมเป็น SubTotal-All
            Else
                cTotalAmt = cTotalAmt + CDbl(SP_FMTcCurPoint(cBluePrice, True, False))   'CDbl(SP_FMTcCurPoint(IIf(tC_MallRegAlw = "Y", cBluePrice, 0) + IIf(tC_MallProAlw = "Y", cYellowPrice, 0), True, False))
            End If
        Else
            If aEmpCard(nI).udt_cB4CalAmt = 0 Then
                If tDiscMallAlw = "N" Then
                    aEmpCard(nI).udt_cB4CalAmt = 0
                    aEmpCard(nI).udt_cB4UnCalAmt = SP_FMTcCurPoint(cYellowPrice + cBluePrice, True, False)
                Else
                    'กลุ่มที่ได้ mall card discount จาก allow mall card
                    '-------------------------------------------------------------------------------------------------------------------
                    If aEmpCard(nI).udt_tStaDisMallCrd = "1" Then
                        
                        If tC_MallFrmAmt = "Y" Then
                            'cTotalAmt = cTotalAmt + CDbl(SP_FMTcCurPoint(cBluePrice + cYellowPrice, True, False))
                            cTotalAmt = cTotalAmt + CDbl(SP_FMTcCurPoint(cSaleAmt, True, False)) '*KT 55-06-27 CommSheet ML-PH4-R3-026 Sometine Mall Card not apply after Mall Coupon ผลกระทบจาก Req.13 Phase 4 ทำให้ขอมูล cBluePrice,cYellowPrice ไม่สามารถนำมารวมเป็น SubTotal-All
                        Else
                            cTotalAmt = cTotalAmt + CDbl(SP_FMTcCurPoint(cBluePrice, True, False)) 'CDbl(SP_FMTcCurPoint(IIf(tC_MallRegAlw = "Y", cBluePrice, 0) + IIf(tC_MallProAlw = "Y", cYellowPrice, 0), True, False))
                        End If
                        
                        If cTotalAmt > cC_AmtMallMax Then
                            If cChkTotalAmt < cC_AmtMallMax Then
                                aEmpCard(nI).udt_cB4CalAmt = cC_AmtMallMax - cChkTotalAmt
                            Else
                                aEmpCard(nI).udt_cB4CalAmt = 0
                            End If
                            aEmpCard(nI).udt_cB4UnCalAmt = SP_FMTcCurPoint(cYellowPrice, True, False)
                        ElseIf cTotalAmt >= cC_AmtMallMin And cTotalAmt <= cC_AmtMallMax Then
                            aEmpCard(nI).udt_cB4CalAmt = CDbl(SP_FMTcCurPoint(IIf(tC_MallRegAlw = "Y", cBluePrice, 0) + IIf(tC_MallProAlw = "Y", cYellowPrice, 0), True, False))
                            aEmpCard(nI).udt_cB4UnCalAmt = CDbl(SP_FMTcCurPoint(IIf(tC_MallRegAlw = "Y", 0, cBluePrice) + IIf(tC_MallProAlw = "Y", 0, cYellowPrice), True, False)) 'SP_FMTcCurPoint(cYellowPrice, True, False)
                        Else
                            If (cTotalAmt2 >= cC_AmtMallMin And cTotalAmt <= cC_AmtMallMax) Or cTotalAmt2 > cC_AmtMallMax Then
                                    aEmpCard(nI).udt_cB4CalAmt = CDbl(SP_FMTcCurPoint(IIf(tC_MallRegAlw = "Y", cBluePrice, 0) + IIf(tC_MallProAlw = "Y", cYellowPrice, 0), True, False))
                                    aEmpCard(nI).udt_cB4UnCalAmt = CDbl(SP_FMTcCurPoint(IIf(tC_MallRegAlw = "Y", 0, cBluePrice) + IIf(tC_MallProAlw = "Y", 0, cYellowPrice), True, False)) 'SP_FMTcCurPoint(cYellowPrice, True, False)
                            Else
                                    aEmpCard(nI).udt_cB4CalAmt = 0
                                    aEmpCard(nI).udt_cB4UnCalAmt = SP_FMTcCurPoint(cYellowPrice + cBluePrice, True, False)
'                                    If tC_MallFrmAmt = "Y" Then
'                                        cTotalAmt = cTotalAmt + CDbl(SP_FMTcCurPoint(cBluePrice + cYellowPrice, True, False))
'                                    Else
'                                        cTotalAmt = cTotalAmt + aEmpCard(nI).udt_cB4CalAmt  'CDbl(SP_FMTcCurPoint(IIf(tC_MallRegAlw = "Y", cBluePrice, 0) + IIf(tC_MallProAlw = "Y", cYellowPrice, 0), True, False))
'                                    End If
                            End If
                        End If
                    Else
                        aEmpCard(nI).udt_cB4CalAmt = CDbl(SP_FMTcCurPoint(IIf(tC_MallRegAlw = "Y", cBluePrice, 0) + IIf(tC_MallProAlw = "Y", cYellowPrice, 0), True, False))
                        aEmpCard(nI).udt_cB4UnCalAmt = SP_FMTcCurPoint(cYellowPrice, True, False)
                    End If
                    '------------------------------------------------------------------------------------------------------------------------
                End If
            Else
                'กลุ่มที่ได้ mall card discount จาก allow sku,not allow sku,allow dpt.code,not allow dpt.code
                'ไม่ต้องระบค่า  aEmpCard(nI).udt_cB4CalAmt,aEmpCard(nI).udt_cB4UnCalAmt  เพราะก่อนหน้าทำมาแล้ว
                If tC_MallFrmAmt = "Y" Then
                    'cTotalAmt = cTotalAmt + CDbl(SP_FMTcCurPoint(cBluePrice + cYellowPrice, True, False))
                    cTotalAmt = cTotalAmt + CDbl(SP_FMTcCurPoint(cSaleAmt, True, False)) '*KT 55-06-27 CommSheet ML-PH4-R3-026 Sometine Mall Card not apply after Mall Coupon ผลกระทบจาก Req.13 Phase 4 ทำให้ขอมูล cBluePrice,cYellowPrice ไม่สามารถนำมารวมเป็น SubTotal-All
                Else
                    cTotalAmt = cTotalAmt + aEmpCard(nI).udt_cB4CalAmt  'CDbl(SP_FMTcCurPoint(IIf(tC_MallRegAlw = "Y", cBluePrice, 0) + IIf(tC_MallProAlw = "Y", cYellowPrice, 0), True, False))
                End If
            End If
        End If
        cChkTotalAmt = cChkTotalAmt + aEmpCard(nI).udt_cB4CalAmt
        
        aEmpCard(nI).udt_cB4Disc = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalAmt * aEmpCard(nI).udt_cCalRate / 100, True, False) 'Old
        aEmpCard(nI).udt_cB4CalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalAmt - aEmpCard(nI).udt_cB4Disc, True, False) 'CalBal = Cal-Disc
        aEmpCard(nI).udt_cB4AllCalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalBalance + aEmpCard(nI).udt_cB4UnCalAmt, True, False) 'All Balance
        aEmpCard(nI).udt_cAFCalAmt = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalAmt, True, False)
        aEmpCard(nI).udt_cAFUnCalAmt = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4UnCalAmt, True, False)
        aEmpCard(nI).udt_cAFDisc = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4Disc, True, False)
        aEmpCard(nI).udt_cAFCalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cB4CalBalance, True, False)
        aEmpCard(nI).udt_cAFAllCalBalance = SP_FMTcCurPoint(aEmpCard(nI).udt_cAFUnCalAmt, True, False)
        aEmpCard(nI).udt_cAllLimit = 0
        aEmpCard(nI).udt_cMallCrdSkuRate = cFTSkuMallCrdRate
        porsTemp.MoveNext
    Loop
    
    For nI = 1 To UBound(aEmpCard)
        bRate = False
        If UBound(aW_MallCard) > 0 Then
            For nMall = 0 To UBound(aW_MallCard) - 1
                If aW_MallCard(nMall).MALLRATE = aEmpCard(nI).udt_cCalRate Then
                    bRate = True
                    aW_MallCard(nMall).MALLAMT = CDbl(Val(Format(aW_MallCard(nMall).MALLAMT))) + aEmpCard(nI).udt_cB4CalAmt
                    Exit For
                End If
            Next nMall
        End If
        If Not bRate Then
            aW_MallCard(UBound(aW_MallCard)).MALLAMT = aEmpCard(nI).udt_cB4CalAmt
            aW_MallCard(UBound(aW_MallCard)).MALLRATE = aEmpCard(nI).udt_cCalRate
            ReDim Preserve aW_MallCard(UBound(aW_MallCard) + 1)
        End If
        '*KT 53-02-09 Communication Sheet PH3-ML-SF-069 Discount Amt not same.xls
        'ทำการหา FNSdtSeqNo สุดท้ายที่ได้ส่วนลด Mall Card
        If aEmpCard(nI).udt_cAFDisc > 0 Then
            nSeqNoLast = aEmpCard(nI).udt_nSeqNo
        End If
    Next nI
    
    nStep = 2
    cDiscHD = 0
    'ตรวจสอบกรณีที่วันที่เริ่มต้นและสันที่สินค้าการให้ส่วนลด mall card
    '*KT 52-06-04 ตรวจสอบวันที่หมดอายุจากที่กำหนดใน parameter
    If Format(SP_GETdChkDateDisc, "YYYY/MM/DD") >= Format(dC_MallStartDate, "YYYY/MM/DD") And Format(SP_GETdChkDateDisc, "YYYY/MM/DD") <= Format(dC_MallEndDate, "YYYY/MM/DD") Then
        For nMall = 0 To UBound(aW_MallCard) - 1
                cDiscHD = cDiscHD + SP_FMTcCurPoint(CDbl(Val(Format(aW_MallCard(nMall).MALLAMT))) * CDbl(Val(Format(aW_MallCard(nMall).MALLRATE))) / 100, True, False)
        Next nMall
    End If
    
    If cDiscHD = 0 And tC_DiscMallAlw = "N" And tC_StaChkDepart = "N" Then
        Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[2]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
        Exit Function
    End If
    
    nStep = 3
    bW_ChkMallCard = True
    If cDiscHD > 0 Then
        Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[3]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
        For nI = 1 To UBound(aEmpCard)
'            '*KT 51-09-25
'            If nI = UBound(aEmpCard) Then
'                bSeqLast = True
'            Else
'                bSeqLast = False
'            End If
'*KT 53-02-09 Communication Sheet PH3-ML-SF-069 Discount Amt not same.xls
'ตรวจสอบ Array เป็น FNSdtSeqNo สุดท้ายที่ตรงกับ nSeqNoLast หรือไม่ เพื่อส่ง Flag bSeqLast
            If aEmpCard(nI).udt_nSeqNo = nSeqNoLast Then
                bSeqLast = True
            Else
                bSeqLast = False
            End If
            Call C_MALxInsMall(aEmpCard(nI).udt_cAFDisc, aEmpCard(nI).udt_nSeqNo, "", aEmpCard(nI).udt_cQtyAllAvi, aEmpCard(nI).udt_cAFCalAmt, aEmpCard(nI).udt_cCalRate, cDisSumAmt, aEmpCard(nI).udt_cMallCrdSkuRate, cDiscHD, bSeqLast)
        Next nI
    End If
    nStep = 4
    If bW_ChkMallCard = True Then
        Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[8]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
        Call C_DATxInsertDiscCD(0, 1, "", 0, 0, 0, , , 0)
    End If
    
     '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
'    If tC_MallAlwPoint = "Y" Then
'        porsTemp.Filter = adFilterNone
'        porsTemp.MoveFirst
'        If Not (porsTemp.BOF And porsTemp.EOF) Then
'                porsTemp.Filter = "FTSdtMemberPointAlw = 'Y' AND FTSdtStaPdt14 <> '1'"
'                If Not (porsTemp.BOF And porsTemp.EOF) Then
'                    Do While Not porsTemp.EOF
'                        tSql = "UPDATE " & tVB_CNTblDT
'                        tSql = tSql & vbCrLf & "SET FCSdtBSubPro = " & SP_FEDvChkNumeric(porsTemp, "FCSdtBSubPro")
'                        tSql = tSql & vbCrLf & ",FCSdtDiscSubPro =" & SP_FEDvChkNumeric(porsTemp, "FCSdtDiscSubPro")
'                        tSql = tSql & vbCrLf & ",FCSdtBSubReg =" & SP_FEDvChkNumeric(porsTemp, "FCSdtBSubReg")
'                        tSql = tSql & vbCrLf & ",FCSdtDiscSubReg =" & SP_FEDvChkNumeric(porsTemp, "FCSdtDiscSubReg")
'                        tSql = tSql & vbCrLf & ",FCSdtBluePrice =" & SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
'                        tSql = tSql & vbCrLf & ",FCSdtYellowPrice =" & SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
'                        tSql = tSql & vbCrLf & ",FTSkuStaMallEmpCard ='" & SP_FEDvChkNumeric(porsTemp, "FTSkuStaMallEmpCard") & "'"
'                        tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                        tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
'                        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                        tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
'                        tSql = tSql & vbCrLf & " AND FTSkuStaMallEmpCard = 'Y'"
'                        nEff = 0
'                        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine, , , nEff)
'                        If nEff = 0 Then
'                            tSql = "DELETE FROM " & tVB_CNTblCD
'                            tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                            tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
'                            tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                            tSql = tSql & vbCrLf & " AND FNDctNo = " & nC_ENDiscType
'                            tSql = tSql & vbCrLf & " AND FTScdCardID = '" & tC_CardID & "'"
'                            tSql = tSql & vbCrLf & " AND FTScdCardType = '" & tC_CardType & "'"
'                            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine, , , nEff)
'                            C_GETcDiscMallCard = 0
'                            Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[11]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
'                            Exit Function
'                        End If
'                        porsTemp.MoveNext
'                    Loop
'                End If
'                porsTemp.Filter = adFilterNone
'        End If
'    End If

    '*KT 54-11-11 V.4.0.4 DCR_5411002
    'SKU Point Active  = YES    , Allow Discount (SKU Point Maintenance)= NO     Mall Card/Employee Card(SKU Maintenance) = YES
    'ให้สินค้านั้นเป็นสินค้าป้ายเหลืองทันที มีผลต่อยอด SUBTOTAL-REG
    If tC_MallAlwPoint = "Y" And bYellow Then
           porsTemp.Filter = adFilterNone
           porsTemp.MoveFirst
           If Not (porsTemp.BOF And porsTemp.EOF) Then
                porsTemp.Filter = "FTSdtMemberPointAlw = 'Y' AND FTSdtStaPdt14 <> '1' AND FTSkuStaMallEmpCardEx = 'Y'"
                If Not (porsTemp.BOF And porsTemp.EOF) Then
                         Do While Not porsTemp.EOF
                                tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) "
                                tSql = tSql & vbCrLf & "SET FCSdtBSubPro = " & SP_FEDvChkNumeric(porsTemp, "FCSdtBSubReg")
                                tSql = tSql & vbCrLf & ",FCSdtDiscSubPro =" & SP_FEDvChkNumeric(porsTemp, "FCSdtDiscSubReg")
                                tSql = tSql & vbCrLf & ",FCSdtBSubReg =" & SP_FEDvChkNumeric(porsTemp, "FCSdtBSubPro")
                                tSql = tSql & vbCrLf & ",FCSdtDiscSubReg =" & SP_FEDvChkNumeric(porsTemp, "FCSdtDiscSubPro")
                                tSql = tSql & vbCrLf & ",FCSdtBluePrice =" & SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
                                tSql = tSql & vbCrLf & ",FCSdtYellowPrice =" & SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
                                tSql = tSql & vbCrLf & ",FTSkuStaMallEmpCard ='N'"
                                tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
                                tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
                                tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
                                tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
                                tSql = tSql & vbCrLf & " AND FTSkuStaMallEmpCard = 'Y'"
                                nEff = 0
                                Call SP_SQLvExecute(tSql, , oDB.ocnOffLine, , , nEff)
                                If nEff = 0 Then
                                     Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[11]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
                                End If
                                porsTemp.MoveNext
                         Loop
                End If
           End If
    End If
    

    C_GETcDiscMallCard = cDisSumAmt
    Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[9]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
    Exit Function
Err_Handler:
    C_GETcDiscMallCard = 0
    Call SP_CRTxLogMallCardEvent("MSR/Keyboard", "oDisc[10]", IIf(tVB_CNTrnDiscType <> "1850", "Discount", "Mall Card Register"), tC_CardID)
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, tVB_CNWheBill & "(C_GETcDiscMallCard[" & nStep & "])", 2)  '*KT 52-03-17 PH1.5-ML-SF-120
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub C_GETxSumSaleAmt(ByVal porsTemp As ADODB.Recordset, ByRef pcSumAmt As Double, ByRef pcSumQty As Double)
'CREATE BY            # kitti                              LAST MODIFIED BY           # kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 52-02-25
'DESCRIPTION       #
'ARGUMENT LIST # recordset
'
'RETURN VALUE  # ค่า sum FCSdtSaleAmt
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
On Error GoTo Err_Handler
    pcSumAmt = 0
    pcSumQty = 0
    Do While Not porsTemp.EOF
        pcSumAmt = pcSumAmt + SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
        pcSumQty = pcSumQty + SP_FEDvChkNumeric(porsTemp, "FCSdtQty")
        porsTemp.MoveNext
    Loop
    Exit Sub
Err_Handler:
    pcSumAmt = 0
    pcSumQty = 0
End Sub

Private Sub C_GETxSumSaleAmtEx(ByVal porsTemp As ADODB.Recordset, ByRef pcRegSumAmt As Double, ByRef pcProSumAmt As Double, ByRef pcSumQty As Double)
'CREATE BY            # kitti                              LAST MODIFIED BY           # kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 52-02-25
'DESCRIPTION       #
'ARGUMENT LIST # recordset
'
'RETURN VALUE  # ค่า sum FCSdtSaleAmt
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
On Error GoTo Err_Handler
    pcRegSumAmt = 0
    pcProSumAmt = 0
    pcSumQty = 0
    Do While Not porsTemp.EOF
        pcRegSumAmt = pcRegSumAmt + SP_FEDvChkNumeric(porsTemp, "FCSdtBluePrice")
        pcProSumAmt = pcProSumAmt + SP_FEDvChkNumeric(porsTemp, "FCSdtYellowPrice")
        pcSumQty = pcSumQty + SP_FEDvChkNumeric(porsTemp, "FCSdtQty")
        porsTemp.MoveNext
    Loop
    Exit Sub
Err_Handler:
    pcRegSumAmt = 0
    pcProSumAmt = 0
End Sub

Private Function C_GETcDiscDiscMrkTran(ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 512-10-06
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cAllCanDisc As Double
    Dim cDisPer As Double
    Dim cDisAmt As Double
    Dim cDisSumAmt As Double
    Dim tEventID As String
    Dim nCount As Integer
    Dim nFNSdtSeqNo As Integer
    Dim tFTSkuCode As String
    Dim cFCSdtSalePrice As Double
    Dim cFCSdtSaleAmt As Double
    Dim cFCSdtAmtDisAvi As Double
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cDisQty As Double
    Dim tWhe As String
    Dim cSourceAmt As Double
    Dim tSql As String
    Dim nI As Long
    Dim tFilter As String
    Dim bTran As Boolean
    Dim bLast As Boolean
    Dim cDisAmtHD As Double
    Dim cDisProRate As Double
        
On Error GoTo Err_Handler
    cAllCanDisc = C_GETcB4DisDDP(porsTemp)
    If cAllCanDisc <= 0 Then Exit Function
    porsTemp.MoveFirst
    cC_DiscPer = 0
    cC_DiscAmt = 0
    cC_DiscNew = 0
    nI = 0
    Select Case nC_ENLnDisType
        Case nEN_LnDis1Per
            cC_DiscPer = cC_DiscValue
        Case nEN_LnDis2Amt
            cC_DiscAmt = cC_DiscValue
        Case nEN_LnDis3NewPrice
            cC_DiscNew = cC_DiscValue
    End Select
    cC_DiscPerOrg = cC_DiscPer
    cC_DiscAmtOrg = cC_DiscAmt
    cC_DiscNewOrg = cC_DiscNew
    
    '*KT 51-10-21 V 3.0.140/V 3.0.200
    '--------------------------------------------------------------------------------------------------------------
    porsTemp.Filter = adFilterNone
    If nC_ENDiscType = nEN_Disc13TranDisc Then
'        tFilter = "FTSkuStaDisAlwTran = 'Y' AND FDSkuDisAlwEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
        '*ao 2011/01/11 CommSheet PH3-ML-SF-138 Wrong Discount amount for Tran Mkdn and Tran Dis.xls
        'สินค้าที่ได้ Free การให้ส่วนลดจะคำนวนผิด
        tFilter = "FTSkuStaDisAlwTran = 'Y' AND FDSkuDisAlwEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "' AND FCSdtSaleAmt > 0"
        porsTemp.Filter = tFilter
    Else
'        tFilter = "FTSkuStaMrkAlwTran = 'Y' AND FDSkuMrkAlwEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
        Select Case nC_ENDiscType
            Case nEN_Disc14TranMDown
                tFilter = "FTSkuStaMrkAlwTran = 'Y' AND FDSkuMrkAlwEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "' AND FCSdtSaleAmt > 0"
            Case nEN_Disc23TranMDown1  '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                tFilter = "FTSkuStaMrkAlwTran1 = 'Y' AND FDSkuMrkAlwEndDate1 >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "' AND FCSdtSaleAmt > 0"
            Case nEN_Disc24TranMDown2
                tFilter = "FTSkuStaMrkAlwTran2 = 'Y' AND FDSkuMrkAlwEndDate2 >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "' AND FCSdtSaleAmt > 0"
             Case nEN_Disc25TranMDown3
                tFilter = "FTSkuStaMrkAlwTran3 = 'Y' AND FDSkuMrkAlwEndDate3 >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "' AND FCSdtSaleAmt > 0"
             Case nEN_Disc26TranMDown4
                tFilter = "FTSkuStaMrkAlwTran4 = 'Y' AND FDSkuMrkAlwEndDate4 >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "' AND FCSdtSaleAmt > 0"
             Case nEN_Disc27TranMDown5
                tFilter = "FTSkuStaMrkAlwTran5 = 'Y' AND FDSkuMrkAlwEndDate5 >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "' AND FCSdtSaleAmt > 0"
        End Select
        porsTemp.Filter = tFilter
    End If
    '--------------------------------------------------------------------------------------------------------------
    
    bTran = False
    ocnAll.BeginTrans
    bTran = True
    nVB_CNCountBegin = nVB_CNCountBegin + 1
    Do While Not porsTemp.EOF
        nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
        tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
        cFCSdtSalePrice = SP_FEDvChkNumeric(porsTemp, "FCSdtSalePrice")
        cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
        cFCSdtAmtDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAvi")
        cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
        cDisPer = 0
        cDisAmt = 0
        nI = nI + 1
        Select Case C_CHKnDDP(nFNSdtSeqNo) 'Check Double Policy
            Case 0 'ไม่อนุญาต
                'คิดเฉพาะส่วนที่เหลือ
               ' If cFCSdtQtyAllDisAvi <= 0 Then Exit Function
                cSourceAmt = cFCSdtAmtDisAvi
            Case 1 'อนุญาต
                cSourceAmt = cFCSdtSaleAmt
            Case 2 'หาผลที่ดีที่สุด
                'รอดำเนินการ
        End Select
        
        If cSourceAmt > 0 Then
                cC_DiscPer = 0
                cC_DiscAmt = 0
                cC_DiscNew = 0
                Select Case nC_ENLnDisType
                    Case nEN_LnDis1Per 'Percent Discount
                        cC_DiscPer = cC_DiscValue
                        cDisPer = cC_DiscValue
                        cDisAmt = SP_FMTcCurPoint(cC_DiscValue * cSourceAmt / 100, True, False)
                    Case nEN_LnDis2Amt  ' Amount Discount
                        cC_DiscAmt = cC_DiscValue
                        cDisPer = 0
                        cDisAmt = SP_FMTcCurPoint(cC_DiscValue * cSourceAmt / cAllCanDisc, True, False)
                    Case nEN_LnDis3NewPrice 'New Price
                        cC_DiscNew = cC_DiscValue
                        cDisPer = 0
                        '*Ao 51-06-18   ทำคำนวณ  New Price
                        If cC_QtyAll = 1 Then
                            If cSourceAmt - cC_DiscNew < 0 Then
                                If bTran Then
                                    ocnAll.RollbackTrans
                                    bTran = False
                                    nVB_CNCountBegin = nVB_CNCountBegin - 1
                                End If
                                cDisAmt = 0
                                Call SP_MSGnShowing(tC_Msg010, tCS_CNMsgWarn)
                                Exit Function
                            Else
                                'แก้ไขสินค้า มีสินค้า=100  2 Item แล้ว New Price ใหม่เป็น 150 จะทำให้ลดไม่ได้
                                cDisAmt = SP_FMTcCurPoint(cSourceAmt / cAllCanDisc * cC_DiscNew, True, False)           '*Ao 51-06-18   ทำคำนวณ  New Price
                                cDisAmt = cSourceAmt - cDisAmt
                            End If
                        Else
                            If cAllCanDisc < cC_DiscNew Then
                                If bTran Then
                                    ocnAll.RollbackTrans
                                    bTran = False
                                    nVB_CNCountBegin = nVB_CNCountBegin - 1
                                End If
                                cDisAmt = 0
                                Call SP_MSGnShowing(tC_Msg010, tCS_CNMsgWarn)
                                Exit Function
                            Else
                                cDisAmt = SP_FMTcCurPoint(cSourceAmt / cAllCanDisc * cC_DiscNew, True, False)           '*Ao 51-06-18   ทำคำนวณ  New Price
                                cDisAmt = cSourceAmt - cDisAmt
                            End If
                        End If
                End Select
            
                    If cDisPer > 100 Then
                        If bTran Then
                            ocnAll.RollbackTrans
                            bTran = False
                            nVB_CNCountBegin = nVB_CNCountBegin - 1
                        End If
                        Call SP_MSGnShowing(tC_Msg004, tCS_CNMsgWarn)
                        Exit Function
                    End If
                    If cSourceAmt - cDisAmt < 0 Then
                        If bTran Then
                            ocnAll.RollbackTrans
                            bTran = False
                            nVB_CNCountBegin = nVB_CNCountBegin - 1
                        End If
                        Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
                        Exit Function
                    End If
                    tEventID = ""
                    
                    '*KT 51-09-2008 แก้ไขเรื่องการ Pro Rate ของ Transaction Discount และ Transaction Markdown
                    '*KT 51-09-30 ตรวจสอบการ Pro Rate เฉพาะ Amount Off
                    '------------------------------------------------------------------------------
                   Select Case nC_ENLnDisType
                        Case nEN_LnDis1Per
                           cDisSumAmt = cDisSumAmt + cDisAmt
                        Case nEN_LnDis2Amt
                            If cDisSumAmt + cDisAmt > cC_DiscValue Then cDisAmt = SP_FMTcCurPoint(cC_DiscValue - cDisSumAmt, True, False)
                            cDisSumAmt = cDisSumAmt + cDisAmt
                            If nI = porsTemp.RecordCount Then
                                If cC_DiscValue > cDisSumAmt Then
                                    cDisAmt = cDisAmt + (cC_DiscValue - cDisSumAmt)
                                                                        
                                    '*Neung 55-01-18 CommSheet ML-PRV-076  Wrong TRan Mkdn discount and Sale Amount on last item.xls
                                    If cDisAmt > cFCSdtSalePrice Then
                                        cDisProRate = SP_FMTcCurPoint(cC_DiscValue - cDisSumAmt, True, False)
                                        cDisAmt = cDisAmt - (cC_DiscValue - cDisSumAmt)
                                        Call C_CHKxProRate(cDisProRate)
                                    End If
                                    '________________________________________________________________________________________
                                    
                                    cDisSumAmt = cDisSumAmt + (cC_DiscValue - cDisSumAmt)
                                End If
                            End If
                        Case nEN_LnDis3NewPrice
                            cDisSumAmt = cDisSumAmt + cDisAmt
                    End Select
                    '----------------------------------------------------------------------------
                    
                    If cDisAmt > 0 Then
                        ''*KT 51-06-03 ใส่เครื่องหมาย * หน้าสินค้าที่ไม่ยอมให้ทำ Double Policy กับ Mall Card
                        '--------------------------------------------------------
                        Select Case C_CHKnDDPMallCard(nFNSdtSeqNo)
                            Case 0 'Not Allow
                                tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) "
                                tSql = tSql & vbCrLf & " SET FTSkuStaMallEmpCard='N'"
                                tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
                                tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
                                tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
                                tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nFNSdtSeqNo
                                Call SP_SQLbExecute(tSql)
                            Case 1 'allow
                                'none
                            Case 2 'Thebest
                                'none
                        End Select
                        '--------------------------------------------------------
                        
                        '*KT 54-08-31 V.4.0.1 แก้ไขเรื่องการคิดส่วนลดกรณีส่วนลดจาก HD กับ ส่วนลดจาก DT มีการ Diff กันเช่น
                        '------------------------------------------------------------------------------------------------------------------------------
                        '                 Amt              Disc
                        'Item 1    100.05         15.01
                        'Item 2    135.92         20.39
                        'Item 3    112.11         16.82
                        'Total     348.08         52.22
                        'TranDisc 15%
                        '                52.212
                        'ปัดทศนิยม 2 ตำแหน่ง
                        '                52.21
                        '------------------------------------------------------------------------------------------------------------------------------
                        bLast = False
                        porsTemp.MoveNext
                        If porsTemp.EOF Then
                            bLast = True
                        End If
                        If bLast Then
                                Select Case nC_ENLnDisType
                                    Case nEN_LnDis1Per 'Percent Discount
                                        cDisAmtHD = SP_FMTcCurPoint(cC_DiscValue * cAllCanDisc / 100, True, False)
                                    Case nEN_LnDis2Amt  ' Amount Discount
                                        cDisAmtHD = cC_DiscValue
                                        If cDisAmtHD > cAllCanDisc Then cDisAmtHD = cAllCanDisc
                                    Case nEN_LnDis3NewPrice 'New Price
                                        cDisAmtHD = cAllCanDisc - cC_DiscValue
                                         If cDisAmtHD > cAllCanDisc Then cDisAmtHD = cAllCanDisc
                                End Select
                                If cDisSumAmt <> cDisAmtHD Then
                                    cDisAmt = cDisAmt + (cDisAmtHD - cDisSumAmt)
                                    cDisSumAmt = cDisAmtHD
                                End If
                        End If
                        porsTemp.MovePrevious
                        
                        bC_FlagInsCD = False
                        Call C_DATxInsertDiscCD(cDisAmt, nFNSdtSeqNo, tEventID, cFCSdtQtyAllDisAvi, cSourceAmt, 0, , , , bC_FlagInsCD)
                        If Not bC_FlagInsCD Then
                            If bTran Then
                                ocnAll.RollbackTrans
                                bTran = False
                                 nVB_CNCountBegin = nVB_CNCountBegin - 1
                            End If
                            C_GETcDiscDiscMrkTran = 0
                            porsTemp.Filter = adFilterNone
                            Exit Function
                        End If
                    End If
        End If
        porsTemp.MoveNext
    Loop
    If bTran Then
        C_GETcDiscDiscMrkTran = cDisSumAmt
        ocnAll.CommitTrans
        bTran = False
        nVB_CNCountBegin = nVB_CNCountBegin - 1
    Else
        C_GETcDiscDiscMrkTran = 0
    End If
    porsTemp.Filter = adFilterNone
    Exit Function
Err_Handler:
    If bTran Then
        ocnAll.RollbackTrans
        bTran = False
        nVB_CNCountBegin = nVB_CNCountBegin - 1
    End If
    C_GETcDiscDiscMrkTran = 0
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_GETcDiscDiscMrkTran)", 2)
    porsTemp.Filter = adFilterNone
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function C_CHKxProRate(ByVal pcDisProRate As Double)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Wanchalerm
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 2012-01-18
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim orsTempSeq As ADODB.Recordset
    Dim bChkLoop As Boolean
    Dim nMax As Integer
    Dim nFNSdtSeqNo As Integer
    Dim cFCScdAmtB4 As Double
    Dim cFCScdAmt As Double
    Dim cFCScdAmtNew As Double

On Error GoTo Err_Handler

    bChkLoop = True
    nMax = 0
        
    tSql = "SELECT  MAX(FNScdSeqNo)  AS FNScdSeqNo FROM " & tVB_CNTblCD & " WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND FNDctNo='14'"
    
    If SP_SQLvExecute(tSql, orsTemp) = 0 Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            Do While Not orsTemp.EOF
                nFNSdtSeqNo = SP_FEDvChkNumeric(orsTemp, "FNScdSeqNo")
                Do While bChkLoop
                    tSql = "SELECT FCScdAmtB4,FCScdAmt FROM " & tVB_CNTblCD & " WITH(NOLOCK) "
                    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
                    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
                    tSql = tSql & vbCrLf & " AND FNDctNo='14'"
                    tSql = tSql & vbCrLf & " AND FNScdSeqNo = " & nFNSdtSeqNo
                    If SP_SQLbExecute(tSql, orsTempSeq) Then
                        If Not (orsTempSeq.BOF And orsTempSeq.EOF) Then
                            Do While Not orsTempSeq.EOF
                            cFCScdAmtB4 = SP_FEDvChkNumeric(orsTempSeq, "FCScdAmtB4")
                            cFCScdAmt = SP_FEDvChkNumeric(orsTempSeq, "FCScdAmt")
                            If SP_FMTcCurPoint(cFCScdAmtB4 - cFCScdAmt, True, False) > pcDisProRate Then
                                cFCScdAmtNew = SP_FMTcCurPoint(cFCScdAmt + pcDisProRate, True, False)
                                tSql = "UPDATE " & tVB_CNTblCD & " WITH(ROWLOCK) "
                                tSql = tSql & vbCrLf & " SET FCScdAmt=" & cFCScdAmtNew
                                tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
                                tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
                                tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
                                tSql = tSql & vbCrLf & " AND FNDctNo='14'"
                                tSql = tSql & vbCrLf & " AND FNScdSeqNo = " & nFNSdtSeqNo
                                Call SP_SQLbExecute(tSql)
                                bChkLoop = False
                            End If
                            orsTempSeq.MoveNext
                            Loop
                        End If
                    End If
                    nFNSdtSeqNo = nFNSdtSeqNo - 1
                Loop
                orsTemp.MoveNext
            Loop
        End If
    End If
    
    Call SP_DATxRsNothing(orsTempSeq)
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function


'Private Function C_GETcDiscAutoTrans22062011(ByVal porsTemp As ADODB.Recordset) As Double
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Kitti
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 54-06-22
''DESCRIPTION       #
''ARGUMENT LIST #
''RETURN VALUE  #
''*KT 54-06-22 V.4.0.0 Ph.4.0 Req.9 Auto Transaction Discount เพิ่มเรื่อง Tier,Allow/Not Allow Dept/Brand Code,Allow/Not Allow SKU Code
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim tSql As String
'    Dim cDisPer As Double
'    Dim cDisPerFromMaster As Double
'    Dim cDisAmt As Double
'    Dim orsTemp As ADODB.Recordset
'    Dim tEventID As String
'    Dim nFNSdtSeqNo As Integer
'    Dim tFTSkuCode As String
'    Dim tFTShdTransType As String
'    Dim tFTSkuDepCode As String
'    Dim tFTSkuBrdCode As String
'    Dim cFCSdtSalePrice As Double
'    Dim cFCSdtSaleAmt As Double
'    Dim cFCSdtAmtDisAvi As Double
'    Dim cFCSdtQtyAllDisAvi As Double
'    Dim cDisQty As Double
'    Dim cSumDisAvi As Double
'    Dim cSumTranAvi As Double
'    Dim cSumAllAvi As Double
'    Dim cDisAllAvi As Double
'    Dim cSumAllDis As Double
'    Dim cDisSumAmt As Double
'    Dim cSourceAmt As Double
'    Dim cFCSdtQty As Double
'    Dim bTran As Boolean
'On Error GoTo Err_Handler
'    With porsTemp
'        .Filter = "FTSkuStaAutoTran = 'Y'"
'        If Not (.BOF And .EOF) Then
'            cSumDisAvi = 0
'            Do While Not .EOF
'                 cSumDisAvi = cSumDisAvi + SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
'                .MoveNext
'            Loop
'        Else
'            cSumDisAvi = 0
'        End If
'    End With
'
'    If cSumDisAvi <= 0 Then
'        Call SP_DATxRsNothing(orsTemp)
'        Exit Function
'    End If
'
'    tSql = "SELECT TOP 1 FCAtdAmt,FCAtdRate,FTAtdStaAlw"
'    tSql = tSql & vbCrLf & " FROM TCNMAutoTranDis"
'    tSql = tSql & vbCrLf & " WHERE FTAtdStaAlw = 'Y'" '*KT 51-05-27 Comsheet ML-SF-295
'    tSql = tSql & vbCrLf & " AND FCAtdAmt <= " & cSumDisAvi
'    tSql = tSql & vbCrLf & " ORDER BY FCAtdAmt DESC"
'    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then 'read master
'        If Not (orsTemp.BOF And orsTemp.EOF) Then
'            cSumTranAvi = SP_FEDvChkNumeric(orsTemp, "FCAtdAmt") 'หายอดของราคาที่เริ่มต้นอนุญาตให้ลด
'            cDisPerFromMaster = SP_FEDvChkNumeric(orsTemp, "FCAtdRate") 'หา%
'        Else
'            Call SP_DATxRsNothing(orsTemp)
'            Exit Function
'        End If
'    Else
'        Exit Function
'    End If
'    If cSumTranAvi > cSumDisAvi Then
'        Call SP_DATxRsClose(orsTemp)
'        Exit Function
'    End If
'    If cSumTranAvi <= 0 Then
'        Call SP_DATxRsNothing(orsTemp)
'        Exit Function
'    End If
'    cDisAllAvi = cDisPerFromMaster * cSumDisAvi / 100
'    cC_DiscValue = cDisPerFromMaster
'    cC_DiscPer = cDisPerFromMaster
'    cC_DiscAmt = 0
'    cC_DiscNew = 0
'    nC_ENLnDisType = nEN_LnDis1Per
'    cC_DiscPerOrg = cC_DiscPer
'    cC_DiscAmtOrg = 0
'    cC_DiscNewOrg = 0
'
'    bTran = False
'    ocnAll.BeginTrans
'    bTran = True
'    nVB_CNCountBegin = nVB_CNCountBegin + 1
'    porsTemp.MoveFirst
'    Do While Not porsTemp.EOF
'        nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
'        tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
'        tFTSkuDepCode = SP_FEDtChkString(porsTemp, "FTSkuDepCode")
'        cFCSdtSalePrice = SP_FEDvChkNumeric(porsTemp, "FCSdtSalePrice")
'        cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
'        cFCSdtAmtDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAvi")
'        cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
'        cFCSdtQty = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAll")
'        Select Case C_CHKnDDP(nFNSdtSeqNo) 'Check Double Policy
'            Case 0 'ไม่อนุญาต
'                'คิดเฉพาะส่วนที่เหลือ
'               ' If cFCSdtQtyAllDisAvi <= 0 Then Exit Function
'                cSourceAmt = cFCSdtAmtDisAvi
'            Case 1 'อนุญาต
'                cSourceAmt = cFCSdtSaleAmt
'            Case 2 'หาผลที่ดีที่สุด
'                'รอดำเนินการ
'        End Select
'        If cSourceAmt > 0 Then
'            If UCase(SP_FEDtChkString(porsTemp, "FTSkuStaAutoTran")) = "Y" Then
'                cDisPer = cC_DiscValue
'                cDisAmt = cC_DiscValue * cSourceAmt / 100
'            Else
'                cDisPer = 0
'                cDisAmt = 0
'            End If
'        Else
'            cDisPer = 0
'            cDisAmt = 0
'        End If
'        If cDisPer > 100 Then
'            If bTran Then
'                ocnAll.RollbackTrans
'                bTran = False
'                nVB_CNCountBegin = nVB_CNCountBegin - 1
'            End If
'            Call SP_MSGnShowing(tC_Msg004, tCS_CNMsgWarn)
'            Call SP_DATxRsNothing(orsTemp)
'            Exit Function
'        End If
'        If cSourceAmt - cDisAmt < 0 Then
'            If bTran Then
'                ocnAll.RollbackTrans
'                bTran = False
'                nVB_CNCountBegin = nVB_CNCountBegin - 1
'            End If
'            Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
'            Call SP_DATxRsNothing(orsTemp)
'            Exit Function
'        End If
'        tEventID = ""
'        cSumAllDis = cSumAllDis + cDisAmt
'        If cDisAmt > 0 Then
'            '*KT 51-06-03 ใส่เครื่องหมาย * หน้าสินค้าที่ไม่ยอมให้ทำ Double Policy กับ Mall Card
'            '--------------------------------------------------------
'            Select Case C_CHKnDDPMallCard(nFNSdtSeqNo)
'                Case 0 'Not Allow
'                    tSql = "UPDATE " & tVB_CNTblDT
'                    tSql = tSql & vbCrLf & " SET FTSkuStaMallEmpCard='N'"
'                    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
'                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nFNSdtSeqNo
'                    Call SP_SQLbExecute(tSql)
'                Case 1 'allow
'                    'none
'                Case 2 'Thebest
'                    'none
'            End Select
'            '--------------------------------------------------------
'            bC_FlagInsCD = False
'            Call C_DATxInsertDiscCD(cDisAmt, nFNSdtSeqNo, tEventID, cFCSdtQtyAllDisAvi, cSourceAmt, 0, , , , bC_FlagInsCD)
'            If Not bC_FlagInsCD Then
'                If bTran Then
'                    ocnAll.RollbackTrans
'                    bTran = False
'                    nVB_CNCountBegin = nVB_CNCountBegin - 1
'                End If
'                C_GETcDiscAutoTrans = 0
'                Call SP_DATxRsNothing(orsTemp)
'                Exit Function
'            End If
'        End If
'        porsTemp.MoveNext
'    Loop
'    If bTran Then
'        ocnAll.CommitTrans
'        bTran = False
'        nVB_CNCountBegin = nVB_CNCountBegin - 1
'    End If
'    C_GETcDiscAutoTrans = cSumAllDis
'    Call SP_DATxRsNothing(orsTemp)
'    Exit Function
'Err_Handler:
'    C_GETcDiscAutoTrans = 0
'    If bTran Then
'        ocnAll.RollbackTrans
'        bTran = False
'        nVB_CNCountBegin = nVB_CNCountBegin - 1
'    End If
'    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_GETcDiscAutoTrans)", 2)
'    Call SP_DATxRsNothing(orsTemp)
'End Function

Private Function C_GETcDiscAutoTrans(ByVal porsTemp As ADODB.Recordset) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 54-06-22
'DESCRIPTION       #
'ARGUMENT LIST #
'RETURN VALUE  #
'*KT 54-06-22 V.4.0.0 Ph.4.0 Req.9 Auto Transaction Discount เพิ่มเรื่อง Tier,Allow/Not Allow Dept/Brand Code,Allow/Not Allow SKU Code
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim cDisPer As Double
    Dim cDisAmt As Double
    Dim orsCon As ADODB.Recordset
    Dim orsSKU As ADODB.Recordset
    Dim orsDept As ADODB.Recordset
    Dim orsTier As ADODB.Recordset
    Dim tEventID As String
    Dim nFNSdtSeqNo As Integer
    Dim tFTSkuCode As String
    Dim tFTShdTransType As String
    Dim tFTSkuDepCode As String
    Dim tFTSkuBrandCode As String
    Dim cFCSdtSalePrice As Double
    Dim cFCSdtSaleAmt As Double
    Dim cFCSdtAmtDisAvi As Double
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cSumAllDis As Double
    Dim cSourceAmt As Double
    Dim cSumSourceAmt As Double
    Dim cFCSdtQty As Double
    Dim bTran As Boolean
    Dim nJ As Long
    Dim nI As Long
    Dim nSeqNoLast As Long
    Dim cDisAmtTemp As Double
    Dim bNext As Boolean
    Dim bIsSKU As Boolean
    Dim bIsDept  As Boolean
    Dim bNew As Boolean
    Dim cDisAmtByItem As Double
    Dim cDiff As Double
    Dim tTypeTier As String
    Dim tStaCons As String
    Dim tNonStaCons As String
    Dim tConsign  As String
    Dim tAtdCode As String
On Error GoTo Err_Handler
    With porsTemp
        .Filter = "FTSkuStaAutoTran = 'Y'"
        If (.BOF And .EOF) Then
            Call SP_DATxRsNothing(orsCon)
            Call SP_DATxRsNothing(orsSKU)
            Call SP_DATxRsNothing(orsDept)
            Call SP_DATxRsNothing(orsTier)
            Exit Function
        End If
    End With

    tSql = "SELECT     dbo.TCNMAutoTranID.FTAttlStaAlw, dbo.TCNMAutoTranID.FDAttStartDate, dbo.TCNMAutoTranID.FDAttEndDate, dbo.TCNMAutoTranDis.FTAtdCode,"
    tSql = tSql & vbCrLf & "dbo.TCNMAutoTranDis.FTAtdName, dbo.TCNMAutoTranDis.FDAtdEndDate, dbo.TCNMAutoTranDis.FTAtdAlwAll, dbo.TCNMAutoTranID.FTAttlStaType,"
    tSql = tSql & vbCrLf & "dbo.TCNMAutoTranID.FTAttName , dbo.TCNMAutoTranID.FTAttTierDiscountID,dbo.TCNMAutoTranDis.FTEvnCode"
    tSql = tSql & vbCrLf & "FROM         dbo.TCNMAutoTranDis WITH(NOLOCK) INNER JOIN"
    tSql = tSql & vbCrLf & "                      dbo.TCNMAutoTranID WITH(NOLOCK) ON dbo.TCNMAutoTranDis.FTAttTierDiscountID = dbo.TCNMAutoTranID.FTAttTierDiscountID"
    tSql = tSql & vbCrLf & "WHERE dbo.TCNMAutoTranID.FTAttlStaAlw = 'Y'"
    tSql = tSql & vbCrLf & "AND '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "' BETWEEN dbo.TCNMAutoTranID.FDAttStartDate AND dbo.TCNMAutoTranDis.FDAtdEndDate"
    tSql = tSql & vbCrLf & "ORDER BY dbo.TCNMAutoTranDis.FTAtdCode"
    If SP_SQLbExecute(tSql, orsCon, , , , oDB.ocnOnLine, True, True) Then 'read master
        If orsCon.BOF And orsCon.EOF Then
            Call SP_DATxRsNothing(orsCon)
            Call SP_DATxRsNothing(orsSKU)
            Call SP_DATxRsNothing(orsDept)
            Call SP_DATxRsNothing(orsTier)
            Exit Function
        End If
    Else
        Call SP_DATxRsNothing(orsCon)
        Call SP_DATxRsNothing(orsSKU)
        Call SP_DATxRsNothing(orsDept)
        Call SP_DATxRsNothing(orsTier)
        Exit Function
    End If

    nI = 0
    ReDim aW_AutoTran(0)
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
    
        nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
        tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
        tFTSkuDepCode = SP_FEDtChkString(porsTemp, "FTSkuDepCode")
        tFTSkuBrandCode = SP_FEDtChkString(porsTemp, "FTSkuBrdCode")
        cFCSdtSalePrice = SP_FEDvChkNumeric(porsTemp, "FCSdtSalePrice")
        cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
        cFCSdtAmtDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAvi")
        cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
        cFCSdtQty = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAll")
        tConsign = SP_FEDtChkString(porsTemp, "FTSdtConsignFlag")
        Select Case C_CHKnDDP(nFNSdtSeqNo) 'Check Double Policy
            Case 0 'ไม่อนุญาต
                'คิดเฉพาะส่วนที่เหลือ
               ' If cFCSdtQtyAllDisAvi <= 0 Then Exit Function
                cSourceAmt = cFCSdtAmtDisAvi
            Case 1 'อนุญาต
                cSourceAmt = cFCSdtSaleAmt
            Case 2 'หาผลที่ดีที่สุด
                'รอดำเนินการ
        End Select
        
        If cSourceAmt > 0 Then
            bNext = True
            
            'Not Allow SKU,Allow SKU
            tSql = "SELECT     FTAtdCode, FTAtsSKU, FTAtsName, FTAtsAlw"
            tSql = tSql & vbCrLf & "FROM dbo.TCNMAutoTranSKU WITH(NOLOCK) "
            If SP_SQLbExecute(tSql, orsSKU, , , , oDB.ocnOnLine, True, True) Then 'read master
                If Not (orsSKU.BOF And orsSKU.EOF) Then
                    bIsSKU = True
                End If
            End If
            If bIsSKU Then
                '--Not Allow SKU
                orsCon.MoveFirst
                Do While Not orsCon.EOF
                    bIsSKU = False
                    tAtdCode = SP_FEDvChkNumeric(orsCon, "FTAtdCode")
                    orsSKU.Filter = adFilterNone
                    orsSKU.Filter = "FTAtdCode = '" & tAtdCode & "' AND FTAtsAlw = '1'"
                    If Not (orsSKU.BOF And orsSKU.EOF) Then
                        bIsSKU = True
                    End If
                    If bIsSKU Then
                        With orsSKU
                            .Filter = "FTAtdCode = '" & tAtdCode & "'  AND FTAtsSKU ='" & tFTSkuCode & "' AND FTAtsAlw = '1'"
                            If Not (.BOF And .EOF) Then
                                bNext = False
                                Exit Do
                            End If
                            .Filter = adFilterNone
                        End With
                    End If
                    orsCon.MoveNext
                Loop
                '- Allow SKU
                If bNext Then
                    orsCon.MoveFirst
                    Do While Not orsCon.EOF
                        bIsSKU = False
                        tAtdCode = SP_FEDvChkNumeric(orsCon, "FTAtdCode")
                        orsSKU.Filter = adFilterNone
                        orsSKU.Filter = "FTAtdCode = '" & tAtdCode & "' AND FTAtsAlw = '0'"
                        If Not (orsSKU.BOF And orsSKU.EOF) Then
                            bIsSKU = True
                        End If
                        If bIsSKU Then
                            With orsSKU
                                .Filter = adFilterNone
                                .Filter = "FTAtdCode = '" & tAtdCode & "'  AND FTAtsSKU ='" & tFTSkuCode & "' AND FTAtsAlw = '0'"
                                If Not (.BOF And .EOF) Then
                                    ReDim Preserve aW_AutoTran(nI)
                                    aW_AutoTran(nI).tSKUCode = tFTSkuCode
                                    aW_AutoTran(nI).tDeptCode = tFTSkuDepCode
                                    aW_AutoTran(nI).tBndCode = tFTSkuBrandCode
                                    aW_AutoTran(nI).nSeqNo = nFNSdtSeqNo
                                    aW_AutoTran(nI).cAmt = cSourceAmt
                                    aW_AutoTran(nI).tDiscountID = SP_FEDvChkNumeric(orsCon, "FTAttTierDiscountID")
                                    aW_AutoTran(nI).tAutoTranID = SP_FEDvChkNumeric(orsCon, "FTAtdCode")
                                    aW_AutoTran(nI).cQtyAllDisAvi = cFCSdtQtyAllDisAvi
                                    aW_AutoTran(nI).tEventID = SP_FEDvChkNumeric(orsCon, "FTEvnCode")
                                    nI = nI + 1
                                    bNext = False
                                    Exit Do
                                End If
                            End With
                        End If
                        orsCon.MoveNext
                    Loop
                End If
            End If
            
            
             'Not Allow Dept/Brand,Allow  Dept/Brand
             If bNext Then
                bIsDept = False
                tSql = "SELECT     FTAtdCode, FTAtcDepCode, FTAtcBrdCode, FTAtcName, FTAtcSeq, FTAtcAlw, FTAtcStaCons, FTAtcStaNonCons"
                tSql = tSql & vbCrLf & "FROM dbo.TCNMAutoTranDep WITH(NOLOCK) "
                If SP_SQLbExecute(tSql, orsDept, , , , oDB.ocnOnLine, True, True) Then 'read master
                    If Not (orsDept.BOF And orsDept.EOF) Then
                        bIsDept = True
                    End If
                End If
                If bIsDept Then
                    '--Not Allow Dept/Brand Code
                    orsCon.MoveFirst
                    Do While Not orsCon.EOF
                        bIsDept = False
                        tAtdCode = SP_FEDvChkNumeric(orsCon, "FTAtdCode")
                        orsDept.Filter = adFilterNone
                        orsDept.Filter = "FTAtdCode = '" & tAtdCode & "' AND FTAtcAlw = '1'"
                        If Not (orsDept.BOF And orsDept.EOF) Then
                            bIsDept = True
                        End If
                        If bIsDept Then
                            With orsDept
                                .Filter = "FTAtdCode = '" & tAtdCode & "' AND FTAtcDepCode ='" & tFTSkuDepCode & "' AND FTAtcBrdCode = '" & tFTSkuBrandCode & "' AND FTAtcAlw = '1'"
                                If Not (.BOF And .EOF) Then
                                    bNext = False
                                    Exit Do
                                End If
                                .Filter = adFilterNone
                                .Filter = "FTAtdCode = '" & tAtdCode & "' AND FTAtcDepCode ='" & tFTSkuDepCode & "' AND FTAtcBrdCode = '' AND FTAtcAlw = '1'"
                                If Not (.BOF And .EOF) Then
                                    bNext = False
                                    Exit Do
                                End If
                                .Filter = adFilterNone
                                .Filter = "FTAtdCode = '" & tAtdCode & "' AND FTAtcDepCode ='' AND FTAtcBrdCode = '" & tFTSkuBrandCode & "' AND FTAtcAlw = '1'"
                                If Not (.BOF And .EOF) Then
                                    bNext = False
                                    Exit Do
                                End If
                                .Filter = adFilterNone
                            End With
                        End If
                        orsCon.MoveNext
                    Loop
                    
                    If bNext Then
                        '- Allow Dept/Brand Code
                        orsCon.MoveFirst
                        Do While Not orsCon.EOF
                            bIsDept = False
                            tAtdCode = SP_FEDvChkNumeric(orsCon, "FTAtdCode")
                            orsDept.Filter = adFilterNone
                            orsDept.Filter = "FTAtdCode = '" & tAtdCode & "' AND FTAtcAlw = '0'"
                            If Not (orsDept.BOF And orsDept.EOF) Then
                                bIsDept = True
                            End If
                            If bIsDept Then
                                With orsDept
                                    .Filter = "FTAtdCode = '" & tAtdCode & "' AND FTAtcDepCode ='" & tFTSkuDepCode & "' AND FTAtcBrdCode = '" & tFTSkuBrandCode & "' AND FTAtcAlw = '0'"
                                    .Sort = "FTAtcStaCons DESC,FTAtcStaNonCons DESC"
                                    If Not (.BOF And .EOF) Then
                                        tStaCons = SP_FEDtChkString(orsDept, "FTAtcStaCons")
                                        tNonStaCons = SP_FEDtChkString(orsDept, "FTAtcStaNonCons")
                                        If tStaCons = "Y" And tNonStaCons = "Y" Then
                                            bNew = True
                                        ElseIf tStaCons = "Y" And tConsign = "Y" Then
                                            bNew = True
                                        ElseIf tNonStaCons = "Y" And tConsign = "N" Then
                                            bNew = True
                                        Else
                                            bNew = False
                                        End If
                                        If bNew Then
                                            ReDim Preserve aW_AutoTran(nI)
                                            aW_AutoTran(nI).tSKUCode = tFTSkuCode
                                            aW_AutoTran(nI).tDeptCode = tFTSkuDepCode
                                            aW_AutoTran(nI).tBndCode = tFTSkuBrandCode
                                            aW_AutoTran(nI).nSeqNo = nFNSdtSeqNo
                                            aW_AutoTran(nI).cAmt = cSourceAmt
                                            aW_AutoTran(nI).tDiscountID = SP_FEDvChkNumeric(orsCon, "FTAttTierDiscountID")
                                            aW_AutoTran(nI).tAutoTranID = SP_FEDvChkNumeric(orsCon, "FTAtdCode")
                                            aW_AutoTran(nI).cQtyAllDisAvi = cFCSdtQtyAllDisAvi
                                             aW_AutoTran(nI).tEventID = SP_FEDvChkNumeric(orsCon, "FTEvnCode")
                                             nI = nI + 1
                                             bNext = False
                                             Exit Do
                                        End If
                                    End If
                                    .Filter = adFilterNone
                                    .Filter = "FTAtdCode = '" & tAtdCode & "' AND FTAtcDepCode ='" & tFTSkuDepCode & "' AND FTAtcBrdCode = '' AND FTAtcAlw = '0'"
                                    .Sort = "FTAtcStaCons DESC,FTAtcStaNonCons DESC"
                                    If Not (.BOF And .EOF) Then
                                        tStaCons = SP_FEDtChkString(orsDept, "FTAtcStaCons")
                                        tNonStaCons = SP_FEDtChkString(orsDept, "FTAtcStaNonCons")
                                        If tStaCons = "Y" And tNonStaCons = "Y" Then
                                            bNew = True
                                        ElseIf tStaCons = "Y" And tConsign = "Y" Then
                                            bNew = True
                                        ElseIf tNonStaCons = "Y" And tConsign = "N" Then
                                            bNew = True
                                        Else
                                            bNew = False
                                        End If
                                        If bNew Then
                                            ReDim Preserve aW_AutoTran(nI)
                                            aW_AutoTran(nI).tSKUCode = tFTSkuCode
                                            aW_AutoTran(nI).tDeptCode = tFTSkuDepCode
                                            aW_AutoTran(nI).tBndCode = tFTSkuBrandCode
                                            aW_AutoTran(nI).nSeqNo = nFNSdtSeqNo
                                            aW_AutoTran(nI).cAmt = cSourceAmt
                                            aW_AutoTran(nI).tDiscountID = SP_FEDvChkNumeric(orsCon, "FTAttTierDiscountID")
                                            aW_AutoTran(nI).tAutoTranID = SP_FEDvChkNumeric(orsCon, "FTAtdCode")
                                            aW_AutoTran(nI).cQtyAllDisAvi = cFCSdtQtyAllDisAvi
                                             aW_AutoTran(nI).tEventID = SP_FEDvChkNumeric(orsCon, "FTEvnCode")
                                             nI = nI + 1
                                             bNext = False
                                             Exit Do
                                         End If
                                    End If
                                    .Filter = adFilterNone
                                    .Filter = "FTAtdCode = '" & tAtdCode & "' AND FTAtcDepCode ='' AND FTAtcBrdCode = '" & tFTSkuBrandCode & "' AND FTAtcAlw = '0'"
                                    .Sort = "FTAtcStaCons DESC,FTAtcStaNonCons DESC"
                                    If Not (.BOF And .EOF) Then
                                        tStaCons = SP_FEDtChkString(orsDept, "FTAtcStaCons")
                                        tNonStaCons = SP_FEDtChkString(orsDept, "FTAtcStaNonCons")
                                        If tStaCons = "Y" And tNonStaCons = "Y" Then
                                            bNew = True
                                        ElseIf tStaCons = "Y" And tConsign = "Y" Then
                                            bNew = True
                                        ElseIf tNonStaCons = "Y" And tConsign = "N" Then
                                            bNew = True
                                        Else
                                            bNew = False
                                        End If
                                        If bNew Then
                                            ReDim Preserve aW_AutoTran(nI)
                                            aW_AutoTran(nI).tSKUCode = tFTSkuCode
                                            aW_AutoTran(nI).tDeptCode = tFTSkuDepCode
                                            aW_AutoTran(nI).tBndCode = tFTSkuBrandCode
                                            aW_AutoTran(nI).nSeqNo = nFNSdtSeqNo
                                            aW_AutoTran(nI).cAmt = cSourceAmt
                                            aW_AutoTran(nI).tDiscountID = SP_FEDvChkNumeric(orsCon, "FTAttTierDiscountID")
                                            aW_AutoTran(nI).tAutoTranID = SP_FEDvChkNumeric(orsCon, "FTAtdCode")
                                            aW_AutoTran(nI).cQtyAllDisAvi = cFCSdtQtyAllDisAvi
                                             aW_AutoTran(nI).tEventID = SP_FEDvChkNumeric(orsCon, "FTEvnCode")
                                             nI = nI + 1
                                             bNext = False
                                             Exit Do
                                        End If
                                    End If
                                    .Filter = adFilterNone
                                End With
                            End If
                            orsCon.MoveNext
                        Loop
                    End If
                End If
            End If
    
            'Allow All
            If bNext Then
                    orsCon.Filter = "FTAtdAlwAll = 'Y'"
                    orsCon.Sort = "FTAtdCode ASC"
                    With orsCon
                        If Not (.BOF And .EOF) Then
                                ReDim Preserve aW_AutoTran(nI)
                                aW_AutoTran(nI).tSKUCode = tFTSkuCode
                                aW_AutoTran(nI).tDeptCode = tFTSkuDepCode
                                aW_AutoTran(nI).tBndCode = tFTSkuBrandCode
                                aW_AutoTran(nI).nSeqNo = nFNSdtSeqNo
                                aW_AutoTran(nI).cAmt = cSourceAmt
                                aW_AutoTran(nI).tDiscountID = SP_FEDvChkNumeric(orsCon, "FTAttTierDiscountID")
                                aW_AutoTran(nI).tAutoTranID = SP_FEDvChkNumeric(orsCon, "FTAtdCode")
                                aW_AutoTran(nI).cQtyAllDisAvi = cFCSdtQtyAllDisAvi
                                 aW_AutoTran(nI).tEventID = SP_FEDvChkNumeric(orsCon, "FTEvnCode")
                                 nI = nI + 1
                                 bNext = False
                        End If
                        .Filter = adFilterNone
                    End With
            End If
        End If
        porsTemp.MoveNext
    Loop
        
    bTran = False
    ocnAll.BeginTrans
    bTran = True
    nVB_CNCountBegin = nVB_CNCountBegin + 1
    
    'ทำการหาเลขที่ AutoTran ID ที่ไม่ซ้ำกันจากสินค้าที่ได้ AutoTran แบบต่างๆ
    ReDim aW_AutoTranTemp(0)
    For nI = 0 To UBound(aW_AutoTran)
        bNew = True
        For nJ = 0 To UBound(aW_AutoTranTemp)
            If aW_AutoTranTemp(nJ) = aW_AutoTran(nI).tDiscountID Or aW_AutoTran(nI).tDiscountID = "" Then
                bNew = False
                Exit For
            End If
        Next nJ
        If bNew Then
            If UBound(aW_AutoTranTemp) > 0 Or aW_AutoTranTemp(0) <> "" Then
                ReDim Preserve aW_AutoTranTemp(UBound(aW_AutoTranTemp) + 1)
            End If
            aW_AutoTranTemp(UBound(aW_AutoTranTemp)) = aW_AutoTran(nI).tDiscountID
        End If
    Next nI
    
    'วนทำการ Insert discount แต่ละกลุ่ม Discount ID
    For nI = 0 To UBound(aW_AutoTranTemp)
        nSeqNoLast = 0
        cSumSourceAmt = 0
        cDisAmtByItem = 0
        cDisAmtTemp = 0
        If aW_AutoTranTemp(nI) <> "" Then
            'หายอดเงินของแต่ละกลุ่ม AutoTran ID เพื่อไปหา Tier
            For nJ = 0 To UBound(aW_AutoTran)
                If aW_AutoTran(nJ).tDiscountID = aW_AutoTranTemp(nI) Then
                    cSumSourceAmt = cSumSourceAmt + aW_AutoTran(nJ).cAmt
                    nSeqNoLast = aW_AutoTran(nJ).nSeqNo
                End If
             Next nJ
             'ถ้ามีสินค้าที่ได้รับ AutoTran ID แต่ละกลุ่ม
             If nSeqNoLast > 0 Then
                'หา Tier
                bNext = False
                tSql = "SELECT     FTAttTierDiscountID, FTAttSeq, FCAttAmt, FTAttlStaType, FNAttDisc, FCAttPer"
                tSql = tSql & vbCrLf & "FROM dbo.TCNMAutoTranIDTier WITH(NOLOCK)  "
                tSql = tSql & vbCrLf & "WHERE FTAttTierDiscountID = '" & aW_AutoTranTemp(nI) & "'"
                'tSql = tSql & vbCrLf & "ORDER BY FTAttSeq DESC" '*Tao 58-12-15 ของเดิม
                tSql = tSql & vbCrLf & "ORDER BY CONVERT(INT, FTAttSeq) DESC" '*Tao 58-12-15 แปลงเป็นตัวเลข จากมากไปหาน้อย
                'tSql = tSql & vbCrLf & "ORDER BY FCAttAmt DESC" '*Tao 58-12-15 เรียงโดยจำนวนเงิน จากมากไปหาน้อย
                If SP_SQLbExecute(tSql, orsTier, , , , oDB.ocnOnLine, True, True) Then 'read master
                   With orsTier
                       If Not (.BOF And .EOF) Then
                            If SP_FEDtChkString(orsTier, "FTAttlStaType") = "0" Then '0 : Amount 1 : %
                               '---Tier Amount
                                tTypeTier = "0"
                               cDisPer = 0
                               cDisAmt = 0
                               Do While Not .EOF
                                   If cSumSourceAmt >= SP_FEDvChkNumeric(orsTier, "FCAttAmt") Then
                                       cDisPer = 0
                                       cDisAmt = SP_FEDvChkNumeric(orsTier, "FNAttDisc")
                                       bNext = True
                                       Exit Do
                                   End If
                                   .MoveNext
                               Loop
                            Else
                               '---Tier %
                               tTypeTier = "1"
                               cDisPer = 0
                               cDisAmt = 0
                               Do While Not .EOF
                                   If cSumSourceAmt >= SP_FEDvChkNumeric(orsTier, "FCAttAmt") Then
                                       cDisPer = SP_FEDvChkNumeric(orsTier, "FCAttPer")
                                       cDisAmt = SP_FMTcCurPoint((cSumSourceAmt * cDisPer / 100), , False)
                                       bNext = True
                                       Exit Do
                                   End If
                                   .MoveNext
                               Loop
                               If cDisPer > 100 Then
                                   If bTran Then
                                       ocnAll.RollbackTrans
                                       bTran = False
                                       nVB_CNCountBegin = nVB_CNCountBegin - 1
                                   End If
                                   Call SP_MSGnShowing(tC_Msg004, tCS_CNMsgWarn)
                                   Call SP_DATxRsNothing(orsCon)
                                   Call SP_DATxRsNothing(orsSKU)
                                   Call SP_DATxRsNothing(orsDept)
                                   Call SP_DATxRsNothing(orsTier)
                                   Exit Function
                               End If
                            End If
                       End If
                   End With
                End If
                If bNext Then
                    cSumAllDis = cSumAllDis + cDisAmt
                   tEventID = ""
                   For nJ = 0 To UBound(aW_AutoTran)
                           If aW_AutoTran(nJ).tDiscountID = aW_AutoTranTemp(nI) Then
                               tEventID = aW_AutoTran(nJ).tEventID
                               
                               If cDisAmt > 0 Then
                                   '*KT 51-06-03 ใส่เครื่องหมาย * หน้าสินค้าที่ไม่ยอมให้ทำ Double Policy กับ Mall Card
                                   '--------------------------------------------------------
                                   Select Case C_CHKnDDPMallCard(aW_AutoTran(nJ).nSeqNo)
                                       Case 0 'Not Allow
                                           tSql = "UPDATE " & tVB_CNTblDT & " WITH(ROWLOCK) "
                                           tSql = tSql & vbCrLf & " SET FTSkuStaMallEmpCard='N'"
                                           tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
                                           tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
                                           tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
                                           tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & aW_AutoTran(nJ).nSeqNo
                                           Call SP_SQLbExecute(tSql)
                                       Case 1 'allow
                                           'none
                                       Case 2 'Thebest
                                           'none
                                   End Select
                                   '--------------------------------------------------------
                                    If cSumSourceAmt <> 0 Then
                                        cDisAmtByItem = SP_FMTcCurPoint((aW_AutoTran(nJ).cAmt * cDisAmt / cSumSourceAmt), , False)
                                    Else
                                        cDisAmtByItem = 0
                                    End If
                                   cDisAmtTemp = cDisAmtTemp + cDisAmtByItem
                                   If nSeqNoLast = aW_AutoTran(nJ).nSeqNo Then
                                        If cDisAmtTemp <> cDisAmt Then
                                            cDiff = cDisAmtTemp - cDisAmt
                                            cDisAmtByItem = cDisAmtByItem - cDiff
                                        End If
                                   End If
                                   bC_FlagInsCD = False
                                    If tTypeTier = "1" Then
                                        cC_DiscValue = cDisPer
                                        cC_DiscPer = cDisPer
                                        cC_DiscAmt = 0
                                        cC_DiscNew = 0
                                        nC_ENLnDisType = nEN_LnDis1Per
                                        cC_DiscPerOrg = cC_DiscPer
                                        cC_DiscAmtOrg = 0
                                        cC_DiscNewOrg = 0
                                    Else
                                        cC_DiscValue = cDisAmtByItem
                                        cC_DiscPer = 0
                                        cC_DiscAmt = cDisAmtByItem
                                        cC_DiscNew = 0
                                        nC_ENLnDisType = nEN_LnDis2Amt
                                        cC_DiscPerOrg = 0
                                        cC_DiscAmtOrg = cDisAmt
                                        cC_DiscNewOrg = 0
                                    End If
                                    tC_CardID = aW_AutoTran(nJ).tAutoTranID
                                    Call C_DATxInsertDiscCD(cDisAmtByItem, aW_AutoTran(nJ).nSeqNo, tEventID, aW_AutoTran(nJ).cQtyAllDisAvi, aW_AutoTran(nJ).cAmt, 0, , , , bC_FlagInsCD)
                                   If Not bC_FlagInsCD Then
                                       If bTran Then
                                           ocnAll.RollbackTrans
                                           bTran = False
                                           nVB_CNCountBegin = nVB_CNCountBegin - 1
                                       End If
                                       C_GETcDiscAutoTrans = 0
                                       Call SP_DATxRsNothing(orsCon)
                                       Call SP_DATxRsNothing(orsSKU)
                                       Call SP_DATxRsNothing(orsDept)
                                       Call SP_DATxRsNothing(orsTier)
                                       Exit Function
                                   End If
                               End If
                           End If
                   Next nJ
                End If
            End If
        End If
    Next nI
        
    If bTran Then
        ocnAll.CommitTrans
        bTran = False
        nVB_CNCountBegin = nVB_CNCountBegin - 1
    End If
    C_GETcDiscAutoTrans = cSumAllDis
    Call SP_DATxRsNothing(orsCon)
    Call SP_DATxRsNothing(orsSKU)
    Call SP_DATxRsNothing(orsDept)
    Call SP_DATxRsNothing(orsTier)
    Exit Function
Err_Handler:
    C_GETcDiscAutoTrans = 0
    If bTran Then
        ocnAll.RollbackTrans
        bTran = False
        nVB_CNCountBegin = nVB_CNCountBegin - 1
    End If
    Call SP_LOGxTableUpdate("DiscType:" & nC_ENDiscType, Err.Number, Err.Description & "(C_GETcDiscAutoTrans)", 2)
    Call SP_DATxRsNothing(orsCon)
    Call SP_DATxRsNothing(orsSKU)
    Call SP_DATxRsNothing(orsDept)
    Call SP_DATxRsNothing(orsTier)
End Function


Private Sub C_GETxFreeGift()
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim cSourceAmt As Double
    Dim tWhe As String
    Dim tFreeDesc As String
    Dim cFreeAmt As Double
    Dim cFreeQty As Double
'    Dim tSKUCode As String
On Error GoTo Err_Handler
    
'    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        tSql = "SELECT SUM(FCSdtSaleAmt) AS FCSdtSaleAmt"
        tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK) "
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND SUBSTRING(ISNULL(FTSdtStaPdt,''),12,1) = '1'"
        tSql = tSql & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
        If SP_SQLvExecute(tSql, orsTemp) = 0 Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                cSourceAmt = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt")
            Else
                cSourceAmt = 0
            End If
        Else
            cSourceAmt = 0
        End If
'    Else
'        tWhe = "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'        tWhe = tWhe & " AND FTShdTransNo = '" & tVB_Running & "'"
'        tWhe = tWhe & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'        tWhe = tWhe & " AND FTSkuStaFreeGift = 'Y'"
'        tWhe = tWhe & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
'        cSourceAmt = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSdtSaleAmt", "" & tVB_CNTblDT & "  LEFT JOIN " & tVB_CNTblSku & " ON " & tVB_CNTblDT & ".FTSkuCode = " & tVB_CNTblSku & ".FTSkuCode", tWhe))
'    End If

    Call SP_DATxRsClose(orsTemp)
    
    tSql = "SELECT FTFgcDesc2,FCFgcAmt,FNFgcQty"
    tSql = tSql & vbCrLf & " FROM TCNMFreeGift WITH(NOLOCK) "
    tSql = tSql & vbCrLf & " WHERE FTFgcStaAlw = 'Y'"
    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            Do While Not orsTemp.EOF
                tFreeDesc = SP_FEDtChkString(orsTemp, "FTFgcDesc2")
                cFreeAmt = SP_FEDvChkNumeric(orsTemp, "FCFgcAmt")
                cFreeQty = SP_FEDvChkNumeric(orsTemp, "FNFgcQty")
                If cFreeAmt <= cSourceAmt Then
                    '*KT 54-09-09 V.4.0.1 แก้ไขให้ใช้ Free Gift ได้
                    Call C_DATxInsertFreeCD(tFreeDesc, cFreeQty)
                End If
                orsTemp.MoveNext
            Loop
        End If
    Else
        Exit Sub
    End If
    
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
Err_Handler:
    Call SP_DATxRsNothing(orsTemp)
End Sub

Private Function C_CHKbAlwOntop(ByRef pnFNCdcOntopRate%, ByRef ptFTCdcMallCard$, ByRef ptFTCdcOnTop$, ByRef pcFCCdcAmtMax#, ByRef pcFCCdcAmtMin#, Optional ByVal pbCredit As Boolean = False) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 53-11-12
'DESCRIPTION       #
'ARGUMENT LIST #
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim tTmp As String
    
'*KT 53-11-12 ปรับปรุง
'    Dim orsTemp As ADODB.Recordset
'    tSql = "SELECT FNCdcOntopRate,FTCdcMallCard,FTCdcOnTop,FCCdcAmtMax,FCCdcAmtMin "
'    tSql = tSql & vbCrLf & ",FTCdcDepCode,FTCdcReg , FTCdcPro"
'    tSql = tSql & vbCrLf & "FROM TCNMCreditCard"
'    tSql = tSql & vbCrLf & "WHERE FTTdmCode = '" & tC_CardKey & "'"
'    tSql = tSql & vbCrLf & "AND FTCdcCreditType = '" & tC_CardType & "'"
'    If Not pbCredit Then
'        tSql = tSql & vbCrLf & "AND  FTCdcOnTop = 'Y'"
'        tSql = tSql & vbCrLf & "AND  '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "' BETWEEN FDCdcStartDate AND FDCdcEndDate" '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
'    End If
'    tTmp = "(FTCdcOnTop = 'Y') AND ('" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "' >= FDCdcStartDate) AND ('" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "' <= FDCdcEndDate)"
     tTmp = "(FTCdcOnTop = 'Y') AND (FDCdcStartDate <= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "' ) AND (FDCdcEndDate >= '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "')"
    If Not oC_RsTempOnTop Is Nothing Then
        If Not pbCredit Then
                '*KT 54-03-02 แก้ไขเรื่องการคิด Ontop ยังไม่ตรวจสอบช่วงวันที่หมดอายุของการให้ Ontop
                oC_RsTempOnTop.Filter = tTmp
        End If
    End If
'    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then 'read master
    If Not oC_RsTempOnTop Is Nothing Then
        If Not (oC_RsTempOnTop.BOF And oC_RsTempOnTop.EOF) Then
            pnFNCdcOntopRate = 0 'SP_FEDvChkNumeric(oC_RsTempOnTop, "FNCdcOntopRate")
            ptFTCdcMallCard = SP_FEDtChkString(oC_RsTempOnTop, "FTCdcMallCard")
            ptFTCdcOnTop = SP_FEDtChkString(oC_RsTempOnTop, "FTCdcOnTop")
            pcFCCdcAmtMax = SP_FEDvChkNumeric(oC_RsTempOnTop, "FCCdcAmtMax")
            pcFCCdcAmtMin = SP_FEDvChkNumeric(oC_RsTempOnTop, "FCCdcAmtMin")
            bC_OnTopUseReg = IIf(SP_FEDtChkString(oC_RsTempOnTop, "FTCdcReg") = "1", True, False)
            bC_OnTopUsePro = IIf(SP_FEDtChkString(oC_RsTempOnTop, "FTCdcPro") = "1", True, False)
            bC_OnTopDepAlw = IIf(UCase(SP_FEDtChkString(oC_RsTempOnTop, "FTCdcDepCode")) = "Y", True, False)
            C_CHKbAlwOntop = True
        Else
            bC_OnTopUseReg = False
            bC_OnTopUsePro = False
            bC_OnTopDepAlw = False
        End If
    Else
        bC_OnTopUseReg = False
        bC_OnTopUsePro = False
        bC_OnTopDepAlw = False
    End If
'    Call SP_DATxRsNothing(orsTemp)
    If Not oC_RsTempOnTop Is Nothing Then    '*Eaw  56-01-26
        oC_RsTempOnTop.Filter = adFilterNone
    End If
'    oC_RsTempOnTop.Filter = adFilterNone
    Exit Function
Err_Handler:
'    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function C_GETbRateOntopByTierAmt(ByRef pcSUMCal As Double, ByVal pcTotalAmt As Double, ByVal ptTdmCode As String, ByVal ptCardType As String) As Boolean
'CREATE BY            # Kitti                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  54-06-03                            LAST MODIFIED DATE     # 54-06-03
'DESCRIPTION       #
'ARGUMENT LIST #
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsRate As ADODB.Recordset
    Dim cCtoAmt As Double
On Error GoTo ErrHandle
    pcSUMCal = 0
    C_GETbRateOntopByTierAmt = False
    tSql = "SELECT FTTdmCode,FTCdcCreditType,FCCtoAmt, FCCtoAmtOff"
    tSql = tSql & vbCrLf & "FROM TCNMCreditCardTierOnTop WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE FTTdmCode='" & ptTdmCode & "'"
    tSql = tSql & vbCrLf & "AND FTCdcCreditType='" & ptCardType & "'"
    tSql = tSql & vbCrLf & "AND FTCtoType='0'" 'Tier Amount
    tSql = tSql & vbCrLf & "ORDER BY FTCtoSeq DESC"
    If SP_SQLvExecute(tSql, orsRate, oDB.ocnOnLine, , , , , True, True) = 0 Then
        If Not (orsRate.BOF And orsRate.EOF) Then
            Do While Not orsRate.EOF
                cCtoAmt = SP_FEDvChkNumeric(orsRate, "FCCtoAmt")
                If pcTotalAmt >= cCtoAmt Then
                    pcSUMCal = pcSUMCal + SP_FEDvChkNumeric(orsRate, "FCCtoAmtOff")
                    pcTotalAmt = pcTotalAmt - cCtoAmt
                    If tC_StaOnTopAlwTierAmtLast = "Y" Then 'Y : Last N : Loop
                        Exit Do
                    End If
                Else
                    orsRate.MoveNext
                End If
                
            Loop
        End If
    End If
    C_GETbRateOntopByTierAmt = True
    Exit Function
ErrHandle:
     pcSUMCal = 0
    C_GETbRateOntopByTierAmt = False
    Call SP_DATxRsNothing(orsRate)
End Function

Private Function C_GETbRateOntopByTier(ByRef pnFNCdcOntopRate As Integer, ByVal pcTotalAmt As Double, ByVal ptTdmCode As String, ByVal ptCardType As String) As Boolean
'CREATE BY            # Kitti                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  54-06-03                            LAST MODIFIED DATE     # 54-06-03
'DESCRIPTION       #
'ARGUMENT LIST #
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsRate As ADODB.Recordset
    
On Error GoTo ErrHandle
    pnFNCdcOntopRate = 0
    C_GETbRateOntopByTier = False
    'tSql = "SELECT FTTdmCode,FTCdcCreditType,FCCtoAmt, FNCtoPerOff"
    tSql = "SELECT FTTdmCode,FTCdcCreditType,FCCtoAmt, FCCtoPerOff" '*Bump 57-08-28 เปลี่ยน ชื่อฟิลด์จาก  FNCtoPerOff เป็น FCCtoPerOff  ให้ตรงตาม DB
    tSql = tSql & vbCrLf & "FROM TCNMCreditCardTierOnTop WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE FTTdmCode='" & ptTdmCode & "'"
    tSql = tSql & vbCrLf & "AND FTCdcCreditType='" & ptCardType & "'"
    tSql = tSql & vbCrLf & "AND FTCtoType='1'" 'Tier %
    tSql = tSql & vbCrLf & "ORDER BY FTCtoSeq DESC"
    If SP_SQLvExecute(tSql, orsRate, oDB.ocnOnLine, , , , , True, True) = 0 Then
        If Not (orsRate.BOF And orsRate.EOF) Then
            Do While Not orsRate.EOF
                If pcTotalAmt >= SP_FEDvChkNumeric(orsRate, "FCCtoAmt") Then
                    'pnFNCdcOntopRate = SP_FEDvChkNumeric(orsRate, "FNCtoPerOff")
                    pnFNCdcOntopRate = SP_FEDvChkNumeric(orsRate, "FCCtoPerOff")
                    Exit Do
                End If
                orsRate.MoveNext
            Loop
        End If
    End If
    C_GETbRateOntopByTier = True
    Exit Function
ErrHandle:
     pnFNCdcOntopRate = 0
    C_GETbRateOntopByTier = False
    Call SP_DATxRsNothing(orsRate)
End Function

Private Function C_CHKbAlwOntopByPOSNo(ByVal ptTmnNo As String, ByVal ptTmnFloor As String, ByVal ptCardType As String, ByVal ptTdmCode As String) As Boolean
'CREATE BY            # Kitti                            LAST MODIFIED BY           # Kitti
'CREATE DATE      #  54-06-01                             LAST MODIFIED DATE     # 54-06-01
'DESCRIPTION       #
'ARGUMENT LIST #
'RETURN VALUE  #
'*KT 54-06-01 V.4.0.0 Req.19 OnTop Specify POSNO
    Dim orsPOSNo As ADODB.Recordset
    Dim tSql As String
    Dim nEff As Long
    Dim aFloor() As String
    Dim nI As Long
    Dim bIsFloor As Boolean
On Error GoTo ErrHandle

    C_CHKbAlwOntopByPOSNo = False
    If tC_POSNoOnTopAlwType = "0" Then 'By Floor
        If tC_FloorAlwOnTop <> "" Then
             bIsFloor = False
             aFloor = Split(tC_FloorAlwOnTop, ",")
             For nI = 0 To UBound(aFloor)
                 If Trim(aFloor(nI)) = ptTmnFloor Then
                    bIsFloor = True
                 End If
             Next nI
             If tC_FloorAlwOnTop = ptTmnFloor Then bIsFloor = True
             If tVB_CNTerminalFloor = "" Then bIsFloor = False
             If bIsFloor Then
                    C_CHKbAlwOntopByPOSNo = True  'Filter Not Allow POS NO
                    tSql = "SELECT FTCapPOSNo "
                    tSql = tSql & vbCrLf & "FROM TCNMCreditCardOnTopAlwPOS WITH(NOLOCK) "
                    tSql = tSql & vbCrLf & "WHERE FTCapPOSNo='" & ptTmnNo & "'"
                    tSql = tSql & vbCrLf & "AND FTTdmCode='" & ptTdmCode & "'"
                    tSql = tSql & vbCrLf & "AND FTCdcCreditType='" & ptCardType & "'"
                    tSql = tSql & vbCrLf & "AND FTCdcPOSNoAlwType='0'"
                    tSql = tSql & vbCrLf & "AND FTCapAlw='N'"
                    If SP_SQLvExecute(tSql, orsPOSNo, oDB.ocnOnLine, , , nEff, , True, True) = 0 Then
                        If Not (orsPOSNo.BOF And orsPOSNo.EOF) Then
                            C_CHKbAlwOntopByPOSNo = False
                        End If
                    End If
            End If
        End If
    ElseIf tC_POSNoOnTopAlwType = "1" Then  'By POSNo.
        tSql = "SELECT FTCapPOSNo "
        tSql = tSql & vbCrLf & "FROM TCNMCreditCardOnTopAlwPOS WITH(NOLOCK) "
        tSql = tSql & vbCrLf & "WHERE FTCapPOSNo='" & ptTmnNo & "'"
        tSql = tSql & vbCrLf & "AND FTTdmCode='" & ptTdmCode & "'"
        tSql = tSql & vbCrLf & "AND FTCdcCreditType='" & ptCardType & "'"
        tSql = tSql & vbCrLf & "AND FTCdcPOSNoAlwType='1'"
        tSql = tSql & vbCrLf & "AND FTCapAlw='Y'"
        nEff = 0
        If SP_SQLvExecute(tSql, orsPOSNo, oDB.ocnOnLine, , , nEff, , True, True) = 0 Then
            If Not (orsPOSNo.BOF And orsPOSNo.EOF) Then
                C_CHKbAlwOntopByPOSNo = True
            End If
        End If
    Else
        C_CHKbAlwOntopByPOSNo = True
    End If
    Call SP_DATxRsNothing(orsPOSNo)
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsPOSNo)
End Function

Private Sub C_DATxDelMallCrdFromOntop(ByVal pnFNSdtSeqNo%)    ''*Ao 51-05-13   ใหม่
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    
    tSql = "UPDATE " & tVB_CNTblCD & " WITH(ROWLOCK) "
    tSql = tSql & vbCrLf & "SET FCScdAmt = 0"
    tSql = tSql & vbCrLf & ",FCScdAmtAF = FCScdAmtB4"
    tSql = tSql & vbCrLf & " WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & pnFNSdtSeqNo
    tSql = tSql & vbCrLf & " AND FNDctNo = " & nEN_Disc11MallCrd
    Call SP_SQLbExecute(tSql)
End Sub

Private Function C_GETcTotalDiscMallCpn() As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    tSql = "SELECT SUM(FCScdAmt) AS FCScdAmt"
    tSql = tSql & vbCrLf & "From " & tVB_CNTblCD & " WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND FNDctNo IN (" & nEN_Disc08MallCpn & "," & nEN_Disc09CoopCpn & "," & nEN_Disc10CoopBrand & ")"
    Call SP_SQLbExecute(tSql, orsTemp)
    C_GETcTotalDiscMallCpn = 0
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        C_GETcTotalDiscMallCpn = SP_FEDvChkNumeric(orsTemp, "FCScdAmt")
    End If
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function C_GETcSpecialWithPmt(ByVal porsTemp As ADODB.Recordset, ByVal pcAllDiscValue#, ByVal pcAllDiscAvi#, ByVal pcAddReg As Double) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nFNSdtSeqNo As Integer
    Dim tFTSkuCode As String
    Dim tFTSkuDepCode As String
    Dim cFCSdtSalePrice As Double
    Dim cFCSdtSaleAmt As Double
    Dim cFCSdtAmtDisAvi As Double
    Dim cFCSdtQtyAllDisAvi As Double
    Dim cFCSdtRegPrice As Double
    Dim cDisQty As Double
    Dim cDisPer As Double
    Dim cDisAmt As Double
    Dim tEventID As String
    Dim cSourceAmt As Double
    Dim cYellowPrice As Double
    If Not (porsTemp.BOF And porsTemp.EOF) Then
        porsTemp.MoveFirst
        Do While Not porsTemp.EOF
            nFNSdtSeqNo = SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
            tFTSkuCode = SP_FEDtChkString(porsTemp, "FTSkuCode")
            cFCSdtSalePrice = SP_FEDvChkNumeric(porsTemp, "FCSdtSalePrice")
            cFCSdtSaleAmt = SP_FEDvChkNumeric(porsTemp, "FCSdtSaleAmt")
            cFCSdtAmtDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAvi")
            cFCSdtQtyAllDisAvi = SP_FEDvChkNumeric(porsTemp, "FCSdtQtyAllDisAvi")
            cYellowPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtAmtDisAviXXX")
            cDisPer = 0
            cDisAmt = 0
            cSourceAmt = cYellowPrice '*TW 51-09-01
            If cSourceAmt > 0 Then
                cDisAmt = pcAllDiscValue * cSourceAmt / pcAllDiscAvi
            Else
                cDisAmt = 0
            End If
            If cSourceAmt - cDisAmt < 0 Then
                Call SP_MSGnShowing(tC_Msg005, tCS_CNMsgWarn)
                Exit Function
            End If
            tEventID = ""
            C_GETcSpecialWithPmt = C_GETcSpecialWithPmt + cDisAmt
            If cDisAmt > 0 Then
                Call C_DATxInsertDiscCD(cDisAmt, nFNSdtSeqNo, tEventID, cFCSdtQtyAllDisAvi, cDisAmt, 0)
'                Call C_DATxSumDT '*KT 52-04-09 ทำการ ' ไว้ก่อน เพราะจะทำให้ SubReg ผิด
                '*TW 51-09-12
'                Call C_UPDxLabel(True, pcAddReg, 0, True) '*KT 52-04-09 ทำการ ' ไว้ก่อน เพราะจะทำให้ SubReg ผิด
                'Call C_UPDxLabel(True, pcAddReg)'Old
            End If
            porsTemp.MoveNext
        Loop
    End If
    
    '*KT 52-04-09  ต้องให้ทำการ C_DATxInsertDiscCD ให้ทุกรายการสินค้าก่อนแล้วทำ C_DATxSumDT กับ C_UPDxLabel
    Call C_DATxSumDT
    Call C_UPDxLabel(True, pcAddReg, 0, True)
End Function

Public Function DIScDiscount(ByVal pnENDiscType As eEN_DiscType, Optional ByVal pnByType As eEN_VerifyDiscType = nEN_Type3None) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตั้งต้น 3.0.19
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim orsTemp As ADODB.Recordset
    Dim cDiscount As Double
    nC_ENDiscType = pnENDiscType
    If (IsMissing(tC_TmnNum) Or tC_TmnNum = "") Or IsMissing(dC_TranDate) Or (IsMissing(tC_TransNo) Or tC_TransNo = "") Then
        Call SP_MSGnShowing(tC_Msg001, tCS_CNMsgWarn)
        DIScDiscount = 0
        Exit Function
    End If
    If bC_RecalByItem Then
        nC_ActSeqNo = nC_RecalActSeqNo
    Else
        Call C_DISxGetNextAct
    End If
    Select Case pnENDiscType
        Case nEN_Disc16FreeGift
            Call C_GETxFreeGift
        Case nEN_Disc01Deal
            Call C_DISxDeal
            If cC_DiscValue > 0 Then
'                Call C_DATxSumDT
                 Call C_DATxSumDT(nEN_Type1ByItem)  '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls
            End If
        Case Else
            DIScDiscount = C_GETcDiscount
            If Not bC_PromoAftDeal Then  '*KT 52-1125 Deal&Promo
                If DIScDiscount <> 0 Then
                    Select Case nC_ENDiscType
                        Case nEN_Disc08MallCpn To nEN_Disc12EmpCrd
                            If C_CHKbUnseq(nC_ENDiscType) And Not bC_StaRecal Then
                                Call DISxPrcRecal
                            Else
                                Call C_DATxSumDT(pnByType) '*KT 52-09-23 Analyse code
                            End If
                        Case Else
                            Call C_DATxSumDT(pnByType) '*KT 52-09-23 Analyse code
                    End Select
                Else
'                    Select Case nC_ENDiscType                  '*Aoo 2012/10/09    New Price สามารถตั้งราคา ให้มากกว่าราคาเดิมได้
'                        Case nEN_Disc05LnDisc
'                            Call SP_MSGnShowing(tC_Msg006, tCS_CNMsgInfo)
'                        Case nEN_Disc06LnMDown ', nEN_Disc09CoopCpn, nEN_Disc10CoopBrand, nEN_Disc08MallCpn, nEN_Disc13TranDisc, nEN_Disc14TranMDown, nEN_Disc11MallCrd, nEN_Disc12EmpCrd
'                            Call SP_MSGnShowing(tC_Msg007, tCS_CNMsgInfo)
'                        Case Else
'                            'nothing
'                    End Select
                End If
            End If
    End Select
    If Not bC_PromoAftDeal Then '*KT 52-1125 Deal&Promo
        '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls
        If pnENDiscType = nEN_Disc01Deal Then
            If bC_IsLastDeal Then
                Call C_UPDxLabel
            End If
        Else
            Call C_UPDxLabel
        End If
    End If
    DIScDiscount = SP_FMTcCurPoint(DIScDiscount, True, False)
    cC_LastDiscount = DIScDiscount
End Function

Public Function DIScGetCredit() As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 52-07-09
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim nFNCdcOntopRate As Integer
    Dim tFTCdcMallCard As String
    Dim tFTCdcOnTop As String
    Dim cFCCdcAmtMax As Double
    Dim cFCCdcAmtMin As Double

On Error GoTo Err_Handler
    DIScGetCredit = 0
    
    If InStr(1, tCS_CNNotAllowTenderMall, tVB_Transaction) > 0 Then Exit Function '*TW 51-08-19
    cVB_CNSumOntopMaxCal = 0
    'Move data to temp
    Screen.MousePointer = 11
    If C_CHKbAlwOntop(nFNCdcOntopRate, tFTCdcMallCard, tFTCdcOnTop, cFCCdcAmtMax, cFCCdcAmtMin, True) Then
            If UCase(tFTCdcMallCard) <> "Y" Then 'if not mall then
'*KT 52-04-21 ทำการคำนวนส่วนลด Mall Card ผิดในกรณีที่ ชำระด้วย Credit Card ที่ Not Allow Mall Card  จะต้องคำนวนเช่นเดียวกับ Special Coupon / Special Gift Voucher  ตามเอกสาร phase 1.8 UAT status on 20-04-2009
'                 DIScGetCredit = SP_FMTcCurPoint(oDisc.DIScSpecial("T002"), True, False)   '*Eaw 56-01-31 แก้เรื่อง Point Hire Purchase
                DIScGetCredit = SP_FMTcCurPoint(oDisc.DIScSpecial(tVB_TenderCode), True, False)
            End If
    End If
    Screen.MousePointer = 0
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    DIScGetCredit = 0
    Screen.MousePointer = 0
End Function

Public Function DIScCalCredit() As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim nFNCdcOntopRate As Integer
    Dim tFTCdcMallCard As String
    Dim tFTCdcOnTop As String
    Dim cFCCdcAmtMax As Double
    Dim cFCCdcAmtMin As Double
    Dim tOrgCD As String
    Dim tOrgDT As String
    Dim tTmpCD As String
    Dim tTmpDT As String
On Error GoTo Err_Handler
    If InStr(1, tCS_CNNotAllowTenderMall, tVB_Transaction) > 0 Then Exit Function '*TW 51-08-19
    cVB_CNSumOntopMaxCal = 0
    'Move data to temp
    Screen.MousePointer = 11
    tOrgCD = tVB_CNTblCD
    tOrgDT = tVB_CNTblDT
    tTmpCD = tVB_CNTblCD & "Ontop"
    tTmpDT = tVB_CNTblDT & "Ontop"
    
    If C_TMPbCreate(tOrgCD, tTmpCD) And C_TMPbCreate(tOrgDT, tTmpDT) Then
        tVB_CNTblCD = tTmpCD
        tVB_CNTblDT = tTmpDT
        If C_CHKbAlwOntop(nFNCdcOntopRate, tFTCdcMallCard, tFTCdcOnTop, cFCCdcAmtMax, cFCCdcAmtMin, True) Then
                If UCase(tFTCdcMallCard) <> "Y" Then 'if not mall then
'*KT 52-04-21 ทำการคำนวนส่วนลด Mall Card ผิดในกรณีที่ ชำระด้วย Credit Card ที่ Not Allow Mall Card  จะต้องคำนวนเช่นเดียวกับ Special Coupon / Special Gift Voucher  ตามเอกสาร phase 1.8 UAT status on 20-04-2009
'                    DIScCalCredit = SP_FMTcCurPoint(oDisc.DIScSpecial("T002"), True, False)     '*Eaw 56-01-31 แก้เรื่อง Point Hire Purchase
                    DIScCalCredit = SP_FMTcCurPoint(oDisc.DIScSpecial(tVB_TenderCode), True, False)
                End If
        End If
    End If
    tVB_CNTblCD = tOrgCD
    tVB_CNTblDT = tOrgDT
    If SP_TBLbChkExist(tTmpCD) Then
        tSql = "DROP TABLE " & tTmpCD
        Call SP_SQLvExecute(tSql)
    End If
    If SP_TBLbChkExist(tTmpDT) Then
        tSql = "DROP TABLE " & tTmpDT
        Call SP_SQLvExecute(tSql)
    End If
    Screen.MousePointer = 0
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    tVB_CNTblCD = tOrgCD
    tVB_CNTblDT = tOrgDT
    DIScCalCredit = 0
    If SP_TBLbChkExist(tTmpCD) Then
        tSql = "DROP TABLE " & tTmpCD
        Call SP_SQLvExecute(tSql)
    End If
    If SP_TBLbChkExist(tTmpDT) Then
        tSql = "DROP TABLE " & tTmpDT
        Call SP_SQLvExecute(tSql)
    End If
    Screen.MousePointer = 0
End Function

Private Sub C_ONTxGetDT(ByRef porsDT As ADODB.Recordset)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
'    Dim tSKUCode As String
'    Dim nI As Long
'    Dim orsTemp As ADODB.Recordset
'    Dim orsSKU As ADODB.Recordset
On Error GoTo Err_Handler
'    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
'        tSql = "SELECT DT.FTSkuCode,DT.FNSdtSeqNo,DT.FCSdtYellowPrice,DT.FCSdtBluePrice,FTSkuDepCode AS FTSkuDepCode,FTSkuBrdCode AS FTSkuBrdCode"
'        tSql = tSql & vbCrLf & " FROM    " & tVB_CNTblDT & " DT "
'        tSql = tSql & vbCrLf & " WHERE DT.FTTmnNum = '" & tVB_CNTerminalNum & "'"
'        tSql = tSql & vbCrLf & " AND DT.FTShdTransNo = '" & tVB_Running & "'"
'        tSql = tSql & vbCrLf & " AND DT.FDShdTransDate = '" & tUT_SaleDate & "'"
'        tSql = tSql & vbCrLf & " AND ISNULL(DT.FTSdtCancelFlag,'0') <> '1'"
'        tSql = tSql & vbCrLf & " ORDER BY DT.FNSdtSeqNo"
'        If SP_SQLvExecute(tSql, orsTemp) = 0 Then
'            If Not (orsTemp.BOF And orsTemp.EOF) Then
'                Set porsDT = New ADODB.Recordset
'                For nI = 1 To orsTemp.Fields.Count
'                    Call porsDT.Fields.Append(orsTemp.Fields(nI - 1).Name, orsTemp.Fields(nI - 1).Type, orsTemp.Fields(nI - 1).DefinedSize)
'                Next nI
'                porsDT.CursorType = adOpenStatic
'                porsDT.Open
'                tSKUCode = ""
'                Do While Not orsTemp.EOF
'                    tSKUCode = tSKUCode & "'" & SP_FEDtChkString(orsTemp, "FTSkuCode") & "'"
'                    orsTemp.MoveNext
'                    If Not orsTemp.EOF Then tSKUCode = tSKUCode & ","
'                Loop
'                If tSKUCode <> "" Then
'                    tSql = "SELECT FTSkuCode,FTSkuDepCode,FTSkuBrdCode,FTSkuStaOnTop "
'                    tSql = tSql & vbCrLf & " FROM  " & tVB_CNTblSku
'                    tSql = tSql & vbCrLf & " WHERE FTSKUCode IN (" & tSKUCode & ")"
'                    tSql = tSql & vbCrLf & " AND FTSkuStaOnTop = 'Y'"
'                    If SP_SQLvExecute(tSql, orsSKU, oDB.ocnOnLine, , , , , True, True) = 0 Then
'                         If Not (orsSKU.BOF And orsSKU.EOF) Then
'                                orsTemp.MoveFirst
'                                Do While Not orsTemp.EOF
'                                    tSKUCode = SP_FEDtChkString(orsTemp, "FTSkuCode")
'                                    orsSKU.Filter = "FTSkuCode = '" & tSKUCode & "'"
'                                     If Not (orsSKU.BOF And orsSKU.EOF) Then
'                                            porsDT.AddNew
'                                            porsDT.Fields("FTSkuCode").Value = SP_FEDtChkString(orsTemp, "FTSkuCode")
'                                            porsDT.Fields("FNSdtSeqNo").Value = SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
'                                            porsDT.Fields("FCSdtYellowPrice").Value = SP_FEDvChkNumeric(orsTemp, "FCSdtYellowPrice")
'                                            porsDT.Fields("FCSdtBluePrice").Value = SP_FEDvChkNumeric(orsTemp, "FCSdtBluePrice")
'                                            porsDT.Fields("FTSkuDepCode").Value = SP_FEDtChkString(orsTemp, "FTSkuDepCode")
'                                            porsDT.Fields("FTSkuBrdCode").Value = SP_FEDtChkString(orsTemp, "FTSkuBrdCode")
'                                     End If
'                                     orsSKU.Filter = adFilterNone
'                                     orsTemp.MoveNext
'                                Loop
'                         End If
'                    End If
'                End If
'            End If
'        End If
'    Else
'       ' tSql = "SELECT DT.FNSdtSeqNo,DT.FCSdtYellowPrice,DT.FCSdtBluePrice,SKU.FTSkuDepCode,SKU.FTSkuBrdCode"
'        tSql = "SELECT DT.FNSdtSeqNo,DT.FCSdtYellowPrice,DT.FCSdtBluePrice,DT.FTSkuDepCode,DT.FTSkuBrdCode"
'        tSql = tSql & vbCrLf & " FROM    " & tVB_CNTblDT & " DT LEFT JOIN "
'        tSql = tSql & vbCrLf & tVB_CNTblSku & " SKU ON  DT.FTSkuCode = SKU.FTSkuCode"
'        tSql = tSql & vbCrLf & " WHERE DT.FTTmnNum = '" & tVB_CNTerminalNum & "'"
'        tSql = tSql & vbCrLf & " AND DT.FTShdTransNo = '" & tVB_Running & "'"
'        tSql = tSql & vbCrLf & " AND DT.FDShdTransDate = '" & tUT_SaleDate & "'"
'        tSql = tSql & vbCrLf & " AND ISNULL(DT.FTSdtCancelFlag,'0') <> '1'"
'        tSql = tSql & vbCrLf & " AND SKU.FTSkuStaOnTop = 'Y'"
'        tSql = tSql & vbCrLf & " ORDER BY DT.FNSdtSeqNo"
'        Call SP_SQLbExecute(tSql, porsDT)
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'    Call SP_DATxRsNothing(orsSKU)

    tSql = "SELECT DT.FNSdtSeqNo,DT.FCSdtYellowPrice,DT.FCSdtBluePrice,DT.FTSkuDepCode,DT.FTSkuBrdCode,DT.FTSkuCode"
    tSql = tSql & vbCrLf & " FROM    " & tVB_CNTblDT & " DT WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & " WHERE DT.FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND DT.FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND DT.FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND ISNULL(DT.FTSdtCancelFlag,'0') <> '1'"
    tSql = tSql & vbCrLf & " AND SUBSTRING(ISNULL(DT.FTSdtStaPdt,''),11,1) = '1'"
    tSql = tSql & vbCrLf & " ORDER BY DT.FNSdtSeqNo"
    Call SP_SQLbExecute(tSql, porsDT)
    
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
'    Call SP_DATxRsNothing(orsTemp)
'    Call SP_DATxRsNothing(orsSKU)
End Sub

'Private Function C_ONTbExceptOntop(ByVal ptDepCode As String, ByVal ptBrdCode As String, ByRef pbReg As Boolean, ByRef pbPro As Boolean) As Boolean
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  #
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim tSql As String
'    Dim orsTemp As ADODB.Recordset
'    On Error GoTo Err_Handler
'    pbReg = False
'    pbPro = False
'    tSql = "SELECT FTCpcPro,FTCpcReg"
'    tSql = tSql & vbCrLf & "FROM  TCNMCreditCardDep"
'    tSql = tSql & vbCrLf & "WHERE FTCpcCode = '" & ptDepCode & "' AND (ISNULL(FTCpcBrdCode,'') = '" & ptBrdCode & "' OR ISNULL(FTCpcBrdCode,'')='')"
'    tSql = tSql & vbCrLf & "AND FTCdcCreditType = '" & tC_CardType & "'"
'    tSql = tSql & vbCrLf & "AND FTTdmCode = '" & tC_CardKey & "'"
'    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then 'read master
'        If Not (orsTemp.BOF And orsTemp.EOF) Then
'            pbReg = (SP_FEDtChkString(orsTemp, "FTCpcReg") = "1")
'            pbPro = (SP_FEDtChkString(orsTemp, "FTCpcPro") = "1")
'            C_ONTbExceptOntop = True
''            If pbPro And pbReg Then
''                C_ONTbExceptOntop = False
''            Else
''                C_ONTbExceptOntop = True
''            End If
'        End If
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'    Call SP_DATxRsNothing(orsTemp)
'End Function

Private Function C_ONTbExceptOntop(ByVal porsDept As ADODB.Recordset, ByVal ptDepCode As String, ByVal ptBrdCode As String, ByRef pbReg As Boolean, ByRef pbPro As Boolean, ByRef pbNext As Boolean) As Boolean
'CREATE BY            # Kitti                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  54-06-02                             LAST MODIFIED DATE     # 54-07-28
'DESCRIPTION       #  Process Allow or Not Allow by SKU Code /Allow or Not Allow by Dept/Brand Code
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
On Error GoTo Err_Handler
    pbNext = True
    C_ONTbExceptOntop = True
    pbReg = True
    pbPro = True
    If Not porsDept Is Nothing Then
        'Not Allow By Dept/Brand Code
        '---Filter Dept/Brand Code
        porsDept.Filter = "(FTCpcCode='" & ptDepCode & "' AND FTCpcBrdCode='" & ptBrdCode & "' AND FTCpcAlw='1')"
        If Not (porsDept.BOF And porsDept.EOF) Then
                pbReg = Not (SP_FEDtChkString(porsDept, "FTCpcReg") = "Y")
                pbPro = Not (SP_FEDtChkString(porsDept, "FTCpcPro") = "Y")
                If pbReg Or pbPro Then
                    C_ONTbExceptOntop = True
                Else
                    C_ONTbExceptOntop = False
                End If
                porsDept.Filter = adFilterNone
                pbNext = False
                Exit Function
        End If
        '---Filter Dept Code
        porsDept.Filter = "(FTCpcCode='" & ptDepCode & "' AND FTCpcBrdCode='' AND FTCpcAlw='1')"
        If Not (porsDept.BOF And porsDept.EOF) Then
                pbReg = Not (SP_FEDtChkString(porsDept, "FTCpcReg") = "Y")
                pbPro = Not (SP_FEDtChkString(porsDept, "FTCpcPro") = "Y")
                If pbReg Or pbPro Then
                    C_ONTbExceptOntop = True
                Else
                    C_ONTbExceptOntop = False
                End If
                porsDept.Filter = adFilterNone
                pbNext = False
                Exit Function
        End If
        '---Filter Brand Code
        If ptDepCode = "" Then
            porsDept.Filter = "(FTCpcCode='' AND FTCpcBrdCode='" & ptBrdCode & "' AND FTCpcAlw='1')"
            If Not (porsDept.BOF And porsDept.EOF) Then
                    pbReg = Not (SP_FEDtChkString(porsDept, "FTCpcReg") = "Y")
                    pbPro = Not (SP_FEDtChkString(porsDept, "FTCpcPro") = "Y")
                    If pbReg Or pbPro Then
                        C_ONTbExceptOntop = True
                    Else
                        C_ONTbExceptOntop = False
                    End If
                    porsDept.Filter = adFilterNone
                    pbNext = False
                    Exit Function
            End If
        Else
            porsDept.Filter = "(FTCpcCode='' AND FTCpcBrdCode='" & ptBrdCode & "' AND FTCpcAlw='1')"
            If Not (porsDept.BOF And porsDept.EOF) Then
                    pbReg = Not (SP_FEDtChkString(porsDept, "FTCpcReg") = "Y")
                    pbPro = Not (SP_FEDtChkString(porsDept, "FTCpcPro") = "Y")
                    If pbReg Or pbPro Then
                        C_ONTbExceptOntop = True
                    Else
                        C_ONTbExceptOntop = False
                    End If
                    porsDept.Filter = adFilterNone
                    pbNext = False
                    Exit Function
            End If
        End If

        'Allow By Dept/Brand Code
        porsDept.Filter = "FTCpcAlw='0'"
        If Not (porsDept.BOF And porsDept.EOF) Then
            C_ONTbExceptOntop = False
            pbReg = False
            pbPro = False
        End If
        porsDept.Filter = adFilterNone
        '---Filter Dept/Brand Code
        porsDept.Filter = "(FTCpcCode='" & ptDepCode & "' AND FTCpcBrdCode='" & ptBrdCode & "' AND FTCpcAlw='0')"
        If Not (porsDept.BOF And porsDept.EOF) Then
                pbReg = (SP_FEDtChkString(porsDept, "FTCpcReg") = "Y")
                pbPro = (SP_FEDtChkString(porsDept, "FTCpcPro") = "Y")
                If pbReg Or pbPro Then
                    C_ONTbExceptOntop = True
                Else
                    C_ONTbExceptOntop = False
                End If
                porsDept.Filter = adFilterNone
                pbNext = False
                Exit Function
        End If
        '---Filter Dept Code
        porsDept.Filter = "(FTCpcCode='" & ptDepCode & "' AND FTCpcBrdCode='' AND FTCpcAlw='0')"
        If Not (porsDept.BOF And porsDept.EOF) Then
                pbReg = (SP_FEDtChkString(porsDept, "FTCpcReg") = "Y")
                pbPro = (SP_FEDtChkString(porsDept, "FTCpcPro") = "Y")
                If pbReg Or pbPro Then
                    C_ONTbExceptOntop = True
                Else
                    C_ONTbExceptOntop = False
                End If
                porsDept.Filter = adFilterNone
                pbNext = False
                Exit Function
        End If
        '---Filter Brand Code
        If ptDepCode = "" Then
            porsDept.Filter = "(FTCpcCode='' AND FTCpcBrdCode='" & ptBrdCode & "' AND FTCpcAlw='0')"
            If Not (porsDept.BOF And porsDept.EOF) Then
                    pbReg = (SP_FEDtChkString(porsDept, "FTCpcReg") = "Y")
                    pbPro = (SP_FEDtChkString(porsDept, "FTCpcPro") = "Y")
                    If pbReg Or pbPro Then
                        C_ONTbExceptOntop = True
                    Else
                        C_ONTbExceptOntop = False
                    End If
                    porsDept.Filter = adFilterNone
                    pbNext = False
                    Exit Function
            End If
        Else
            porsDept.Filter = "(FTCpcCode='' AND FTCpcBrdCode='" & ptBrdCode & "' AND FTCpcAlw='0')"
            If Not (porsDept.BOF And porsDept.EOF) Then
                    pbReg = (SP_FEDtChkString(porsDept, "FTCpcReg") = "Y")
                    pbPro = (SP_FEDtChkString(porsDept, "FTCpcPro") = "Y")
                    If pbReg Or pbPro Then
                        C_ONTbExceptOntop = True
                    Else
                        C_ONTbExceptOntop = False
                    End If
                    porsDept.Filter = adFilterNone
                    pbNext = False
                    Exit Function
            End If
        End If
    End If
    porsDept.Filter = adFilterNone
    pbNext = False
    Exit Function
Err_Handler:
    C_ONTbExceptOntop = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_ONTbExceptOntopSKU(ByVal porsSKU As ADODB.Recordset, ptSKUCode As String, ByRef pbNext As Boolean) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String

On Error GoTo Err_Handler
    pbNext = True
    If Not porsSKU Is Nothing Then
        porsSKU.Filter = "FTCpsSKU='" & ptSKUCode & "' AND FTCpsAlw='1'"
        If Not (porsSKU.BOF And porsSKU.EOF) Then
            C_ONTbExceptOntopSKU = False
            pbNext = False
            porsSKU.Filter = adFilterNone
            Exit Function
        Else
            porsSKU.Filter = "FTCpsSKU='" & ptSKUCode & "' AND FTCpsAlw='0'"
            If Not (porsSKU.BOF And porsSKU.EOF) Then
                C_ONTbExceptOntopSKU = True
                pbNext = False
            Else
                C_ONTbExceptOntopSKU = False
                pbNext = True
            End If
        End If
        porsSKU.Filter = adFilterNone
    End If
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function C_GETxSKUOnTop(ByRef porsSKU As ADODB.Recordset, ByVal ptTdmCode As String, ByVal ptCardType As String) As Boolean
 'CREATE BY            # Kitti                           LAST MODIFIED BY           # Kitti
'CREATE DATE      #  54-06-01                             LAST MODIFIED DATE     # 54-06-01
'DESCRIPTION       #
'ARGUMENT LIST # Get data allow and not allow ontop discount of sku code code to recordset
'
'RETURN VALUE  #
'*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
On Error GoTo Err_Handler
    C_GETxSKUOnTop = False
    tSql = "SELECT * "
    tSql = tSql & vbCrLf & "FROM  TCNMCreditCardSKU WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "WHERE FTCdcCreditType = '" & ptCardType & "'"
    tSql = tSql & vbCrLf & "AND FTTdmCode = '" & ptTdmCode & "'"
    tSql = tSql & vbCrLf & "AND FTCpsAlwPoint = '0'" 'Ontop
    If SP_SQLbExecute(tSql, porsSKU, , , , oDB.ocnOnLine, True, True) Then 'read master
        C_GETxSKUOnTop = True
    End If
    Exit Function
Err_Handler:
    C_GETxSKUOnTop = False
End Function

'
Private Function C_GETxDeptOnTop(ByRef porsDept As ADODB.Recordset, ByVal ptTdmCode As String, ByVal ptCardType As String) As Boolean
 'CREATE BY            # Kitti                           LAST MODIFIED BY           # Kitti
'CREATE DATE      #  54-06-01                             LAST MODIFIED DATE     # 54-06-01
'DESCRIPTION       #  Get data allow and not allow ontop discount of dept/brand code to recordset
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
On Error GoTo Err_Handler
    C_GETxDeptOnTop = False
    tSql = "SELECT * "
    tSql = tSql & vbCrLf & "FROM  TCNMCreditCardDep WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & "WHERE FTCdcCreditType = '" & ptCardType & "'"
    tSql = tSql & vbCrLf & "AND FTTdmCode = '" & ptTdmCode & "'"
    tSql = tSql & vbCrLf & "AND FTCpcAlwPoint = '0'" 'Ontop
    If SP_SQLbExecute(tSql, porsDept, , , , oDB.ocnOnLine, True, True) Then 'read master
        C_GETxDeptOnTop = True
    End If
    Exit Function
Err_Handler:
    C_GETxDeptOnTop = False
End Function

Public Function DIScCalOntop() As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 54-06-01 V.4.0.0 Req.19 OnTop Specify POSNo.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim tWhr  As String
    Dim orsDT As ADODB.Recordset
    Dim nFNCdcOntopRate As Integer
    Dim tFTCdcMallCard As String
    Dim tFTCdcOnTop As String
    Dim cFCCdcAmtMax As Double
    Dim cFCCdcAmtMin As Double

    Dim nFNSdtSeqNo As Integer
    
    Dim tFTSkuDepCode As String
    Dim tFTSkuBrdCode As String
    Dim dCurrentExpDate As Date
    Dim cSumCal As Double
    Dim tOrgCD As String
    Dim tOrgDT As String
    Dim tTmpCD As String
    Dim tTmpDT As String
    Dim cOldDisc As Double
    Dim cNewDisc As Double
    Dim cSumOntopMaxDisc As Double
    Dim cSumOntopMaxPay As Double
    Dim bUsePro As Boolean
    Dim bUseReg As Boolean
    'Dim cKeepDiscValue As Double
    
    Dim cKeepDiscValue As Double
    Dim tKeeptCardID As String
    Dim tKeeptCardType As String
    Dim tKeeptCardKey As String
    Dim tKeepttStaChkExpire As String
    Dim dKeepdCardExpDate As Date
    Dim cKeepcDiscOntopAlw  As Double
    Dim cKeepcDiscPer  As Double
    Dim cSumOntopRegCal As Double
    Dim cSumOntopProCal As Double
    Dim cSumOntopBothCal As Double
    Dim cBlueCal As Double
    Dim cYellowCal As Double
    Dim cBlueDisc As Double
    Dim cYellowDisc As Double
    
    
    Dim orsSKUOnTop  As ADODB.Recordset
    Dim orsDeptOnTop   As ADODB.Recordset
    Dim tFTSkuCode As String
    Dim bStep As Boolean
    Dim bSubAllTier As Boolean
    Dim bCal As Boolean
    
    'เก็บไว้ เพราะ Recal แล้วจะหาย
    tKeeptCardID = tC_CardID
    tKeeptCardType = tC_CardType
    tKeeptCardKey = tC_CardKey
    cKeepDiscValue = cC_DiscValue
    tKeepttStaChkExpire = tC_StaChkExpire
    dKeepdCardExpDate = dC_CardExpDate
    cKeepcDiscOntopAlw = cC_DiscOntopAlw
    cKeepcDiscPer = cC_DiscPer
    
    'cKeepDiscValue = cC_DiscValue
    cC_AmtAllMall = 0
    If InStr(1, tCS_CNNotAllowTenderMall, tVB_Transaction) > 0 Then Exit Function '*TW 51-08-19
    On Error GoTo Err_Handler
    cVB_CNSumOntopMaxCal = 0
    'Move data to temp
    Screen.MousePointer = 11
    tOrgCD = tVB_CNTblCD
    tOrgDT = tVB_CNTblDT
    tTmpCD = tVB_CNTblCD & "Ontop"
    tTmpDT = tVB_CNTblDT & "Ontop"
    nC_ENDiscType = nEN_Disc15Ontop
    
    tWhr = "WHERE  FTShdTransNo = '" & tVB_Running & "'"
    tWhr = tWhr & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tWhr = tWhr & vbCrLf & "AND  ISNULL(FTSdtStkCode,'') <> '1'" '*KT 52-11-25 Deal&Promo
    
    cOldDisc = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCScdAmt", tOrgCD, tWhr))
    If (IsMissing(tC_TmnNum) Or tC_TmnNum = "") Or IsMissing(dC_TranDate) Or (IsMissing(tC_TransNo) Or tC_TransNo = "") Then
        Call SP_MSGnShowing(tC_Msg001, tCS_CNMsgWarn)
        DIScCalOntop = 0
        Screen.MousePointer = 0
        Exit Function
    End If
    If C_TMPbCreate(tOrgCD, tTmpCD) And C_TMPbCreate(tOrgDT, tTmpDT) Then
        tVB_CNTblCD = tTmpCD
        tVB_CNTblDT = tTmpDT
        Call C_DISxGetNextAct
            
        'Select Credit
        'Check ontop,Expire
        If UCase(tC_StaChkExpire) = "Y" Then
            dCurrentExpDate = DateAdd("m", 1, dC_CardExpDate) - 1
             '*KT 52-06-04 ตรวจสอบวันที่หมดอายุของ point ตามที่กำหนดใน parameter
            If Format(dCurrentExpDate, "YYYY/MM/DD") < Format(SP_GETdChkDateDisc, "YYYY/MM/DD") Then '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26 ใส่ Format mm/dd/yyyyy เพราะ dVB_CurrentDate มีเวลาเข้ามาด้วย Communication Sheet PH2-ML-SF-027 Credit Expire.xls
                Screen.MousePointer = 0
                DIScCalOntop = 0
                cVB_CNSumOntopMaxCal = 0
                cSumCal = 0
                cC_AmtAllMall = 0
                GoTo ExitPoint
            End If
        End If
        
        If C_CHKbAlwOntop(nFNCdcOntopRate, tFTCdcMallCard, tFTCdcOnTop, cFCCdcAmtMax, cFCCdcAmtMin) Then
        
            '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
            'ตรวจสอบว่าเครื่องจุดขายอนุญาตให้ได้รับส่วนลด OnTop หรือไม่
            If bC_POSNoOnTopFlag Then
                If Not C_CHKbAlwOntopByPOSNo(tVB_CNTerminalNum, tVB_CNTerminalFloor, tC_CardType, tVB_TenderCode) Then
                    Screen.MousePointer = 0
                    Call SP_MSGnShowing(tC_Msg011, tCS_CNMsgWarn)
                    DIScCalOntop = 0
                    cVB_CNSumOntopMaxCal = 0
                    cSumCal = 0
                    cC_AmtAllMall = 0
                    GoTo ExitPoint
                End If
            End If
        
            'Calculate from SUBTOTALALL
            bSubAllTier = IIf(SP_FEDtChkString(oC_RsTempOnTop, "FTCdcStaSubAllOnTop") = "Y", True, False)
            If bSubAllTier Then
                If cC_SubTotalAll >= cFCCdcAmtMin Then
                    bCal = True
                Else
                    bCal = False
                End If
            Else
                If cC_SubTotal >= cFCCdcAmtMin Then
                    bCal = True
                Else
                    bCal = False
                End If
            End If
            If Not bCal Then
                Screen.MousePointer = 0
                DIScCalOntop = 0
                cVB_CNSumOntopMaxCal = 0
                cSumCal = 0
                cC_AmtAllMall = 0
                GoTo ExitPoint
            End If
        
            bUsePro = bC_OnTopUsePro
            bUseReg = bC_OnTopUseReg
            Call C_ONTxGetDT(orsDT)
            'if exists
            If Not (orsDT.BOF And orsDT.EOF) Then
                If UCase(tFTCdcMallCard) <> "Y" Then 'if not mall then
'*KT 52-04-21 ทำการคำนวนส่วนลด Mall Card ผิดในกรณีที่ ชำระด้วย Credit Card ที่ Not Allow Mall Card  จะต้องคำนวนเช่นเดียวกับ Special Coupon / Special Gift Voucher  ตามเอกสาร phase 1.8 UAT status on 20-04-2009
                    Call oDisc.DIScSpecial("T002")
                    nC_ENDiscType = nEN_Disc15Ontop
                    tC_CardID = tKeeptCardID
                    tC_CardType = tKeeptCardType
                    tC_CardKey = tKeeptCardKey
                    cC_DiscValue = cKeepDiscValue
                    tC_StaChkExpire = tKeepttStaChkExpire
                    dC_CardExpDate = dKeepdCardExpDate
                    cDiscOntopAlw = cKeepcDiscOntopAlw
                    cC_DiscPer = cKeepcDiscPer
                    cC_DiscPerOrg = cKeepcDiscPer

                    Call C_ONTxGetDT(orsDT) '*KT 53-11-08 ปรับปรุงให้ Performance ดีขึ้น
                End If
                
                 If bC_OnTopDepAlw Then
                    'Record Sku Alw,Not Alw
                    Call SP_DATxRsClose(orsSKUOnTop)
                    Call C_GETxSKUOnTop(orsSKUOnTop, tC_CardKey, tC_CardType)
                    'Record Dept/Brand Code Alw,Not Alw
                     Call SP_DATxRsClose(orsDeptOnTop)
                     Call C_GETxDeptOnTop(orsDeptOnTop, tC_CardKey, tC_CardType)
                 End If
                
                orsDT.MoveFirst
                Do While Not orsDT.EOF 'while
                    bStep = True
                    nFNSdtSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
                    cBlueCal = SP_FEDvChkNumeric(orsDT, "FCSdtBluePrice")
                    cYellowCal = SP_FEDvChkNumeric(orsDT, "FCSdtYellowPrice")
                    cBlueDisc = C_ONTcDiscOldOntop(nFNSdtSeqNo, True) 'หาส่วนลดที่เป็น Ontop
                    cYellowDisc = C_ONTcDiscOldOntop(nFNSdtSeqNo, False) 'หาส่วนลดที่เป็น Ontop
                    tFTSkuCode = SP_FEDtChkString(orsDT, "FTSkuCode")
                    If bC_OnTopDepAlw Then
                        'By SKU Code
                        If C_ONTbExceptOntopSKU(orsSKUOnTop, tFTSkuCode, bStep) Then
                            'Allow SKU Code  นั้นไม่ขึ้นกับ reg ,promo ในส่วนข้างบน
                            cSumOntopBothCal = cSumOntopBothCal + (cYellowCal - cYellowDisc) + (cBlueCal - cBlueDisc)
                            cVB_CNSumOntopMaxCal = cVB_CNSumOntopMaxCal + (cYellowCal - cYellowDisc) + (cBlueCal - cBlueDisc)
                        End If
                        If bStep Then
                            'By Dept/Brand Code
                            tFTSkuDepCode = SP_FEDtChkString(orsDT, "FTSkuDepCode")
                            tFTSkuBrdCode = SP_FEDtChkString(orsDT, "FTSkuBrdCode")
                            If C_ONTbExceptOntop(orsDeptOnTop, tFTSkuDepCode, tFTSkuBrdCode, bUseReg, bUsePro, bStep) Then
                                cSumOntopBothCal = cSumOntopBothCal + IIf(bUsePro, cYellowCal - cYellowDisc, 0) + IIf(bUseReg, cBlueCal - cBlueDisc, 0)
                                cVB_CNSumOntopMaxCal = cVB_CNSumOntopMaxCal + IIf(bUsePro, cYellowCal - cYellowDisc, 0) + IIf(bUseReg, cBlueCal - cBlueDisc, 0)
                            End If
                        End If
                    Else
                        cSumOntopBothCal = cSumOntopBothCal + IIf(bC_OnTopUsePro, cYellowCal - cYellowDisc, 0) + IIf(bC_OnTopUseReg, cBlueCal - cBlueDisc, 0)
                        cVB_CNSumOntopMaxCal = cVB_CNSumOntopMaxCal + IIf(bC_OnTopUsePro, cYellowCal - cYellowDisc, 0) + IIf(bC_OnTopUseReg, cBlueCal - cBlueDisc, 0)
                    End If
                    orsDT.MoveNext
                Loop    'loop
                
                cNewDisc = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCScdAmt", tVB_CNTblCD, tWhr))
                If cVB_CNSumOntopMaxCal > cKeepDiscValue + (cOldDisc - cNewDisc) Then
                    cVB_CNSumOntopMaxCal = cKeepDiscValue + (cOldDisc - cNewDisc)
                End If
                
'                If cFCCdcAmtMin > cVB_CNSumOntopMaxCal Then
'                    DIScCalOntop = 0
'                    cVB_CNSumOntopMaxCal = 0
'                    cSumCal = 0
'                    cC_AmtAllMall = 0
'                    GoTo ExitPoint
'                End If

                cNewDisc = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCScdAmt", tVB_CNTblCD, "WHERE " & tVB_CNWheBill & " AND  ISNULL(FTSdtStkCode,'') <> '1'")) '*KT 52-11-25 Deal&Promo
            End If
            If cFCCdcAmtMax < cVB_CNSumOntopMaxCal Then
                cVB_CNSumOntopMaxCal = cFCCdcAmtMax
            End If
        End If
        
    Else
        cNewDisc = cOldDisc
    End If
    
    ''*KT 54-06-01 V.4.0.0 Req.19 OnTop Specify POSNo. *หา rate % ส่วนลดตาม Tier ใหม่
    If SP_FEDtChkString(oC_RsTempOnTop, "FTCdcStaTier") = "Y" Then
        Call C_GETbRateOntopByTier(nFNCdcOntopRate, cVB_CNSumOntopMaxCal, tC_CardKey, tC_CardType)
        cVB_CNSumOntopMaxCal = SP_FMTcCurPoint(cVB_CNSumOntopMaxCal, True, False)
        cDiscOntopAlw = cVB_CNSumOntopMaxCal
        cSumCal = SP_FMTcCurPoint(nFNCdcOntopRate * cVB_CNSumOntopMaxCal / 100, True, False)
        cDiscValue = cDiscOntopAlw - cSumCal
        cC_AmtAllMall = SP_FMTcCurPoint(cOldDisc - cNewDisc, True, False)
        DIScCalOntop = cSumCal
    Else
        Call C_GETbRateOntopByTierAmt(cSumCal, cVB_CNSumOntopMaxCal, tC_CardKey, tC_CardType)
        cVB_CNSumOntopMaxCal = SP_FMTcCurPoint(cVB_CNSumOntopMaxCal, True, False)
        cDiscOntopAlw = cVB_CNSumOntopMaxCal
        cDiscValue = cDiscOntopAlw - cSumCal
        cC_AmtAllMall = SP_FMTcCurPoint(cOldDisc - cNewDisc, True, False)
        DIScCalOntop = cSumCal
    End If
    

    Call SP_DATxRsNothing(orsDT)
ExitPoint:
    tVB_CNTblCD = tOrgCD
    tVB_CNTblDT = tOrgDT
    If SP_TBLbChkExist(tTmpCD) Then
        tSql = "DROP TABLE " & tTmpCD
        Call SP_SQLvExecute(tSql)
    End If
    If SP_TBLbChkExist(tTmpDT) Then
        tSql = "DROP TABLE " & tTmpDT
        Call SP_SQLvExecute(tSql)
    End If
    Call SP_DATxRsNothing(orsDT)
    Screen.MousePointer = 0
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    tVB_CNTblCD = tOrgCD
    tVB_CNTblDT = tOrgDT
    DIScCalOntop = 0
    If SP_TBLbChkExist(tTmpCD) Then
        tSql = "DROP TABLE " & tTmpCD
        Call SP_SQLvExecute(tSql)
    End If
    If SP_TBLbChkExist(tTmpDT) Then
        tSql = "DROP TABLE " & tTmpDT
        Call SP_SQLvExecute(tSql)
    End If
    Call SP_DATxRsNothing(orsDT)
    Screen.MousePointer = 0
End Function

Private Function C_ONTcDiscOldOntop(ByVal pnFNSdtSeqNo As Integer, ByVal pbSubReg As Boolean) As Double
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    On Error GoTo Err_Handler
    If pbSubReg Then
        tSql = "SELECT SUM(FCScdDiscSubReg) AS FCDisOntop"
    Else
        tSql = "SELECT SUM(FCScdDiscSubPro) AS FCDisOntop"
    End If
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD & " WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & "AND FNSdtSeqNo = " & pnFNSdtSeqNo
    tSql = tSql & vbCrLf & "AND FNDctNo = " & nEN_Disc15Ontop
    If SP_SQLbExecute(tSql, orsTemp) Then
        C_ONTcDiscOldOntop = SP_FEDvChkNumeric(orsTemp, "FCDisOntop")
    Else
        C_ONTcDiscOldOntop = 0
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Public Function DIScGetOntop() As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 54-06-01 V.4.0.0 Req.19 OnTop Specify POSNo.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsDT As ADODB.Recordset
    Dim nFNCdcOntopRate As Integer
    Dim tFTCdcMallCard As String
    Dim tFTCdcOnTop As String
    Dim cFCCdcAmtMax As Double
    Dim cFCCdcAmtMin As Double
    Dim nFNSdtSeqNo As Integer
    Dim tFTSkuDepCode As String
    Dim tFTSkuBrdCode As String
    Dim dCurrentExpDate As Date
    Dim cSumCal As Double
    Dim tOrgCD As String
    Dim tOrgDT As String
    Dim tTmpCD As String
    Dim tTmpDT As String
    Dim cOldDisc As Double
    Dim cNewDisc As Double
    Dim cSumOntopMaxDisc As Double
    Dim cSumOntopMaxPay As Double
    Dim cSumOntopMaxCal As Double
    Dim cSumOntopMaxCalAlw As Double
    Dim bUsePro As Boolean
    Dim bUseReg As Boolean
    Dim cSumSourceAmt As Double
    Dim cSourceAmt As Double
    Dim cDisAmt As Double
    Dim cLastDiscAmt As Double
    Dim nLastFNSdtSeqNo As Integer
    Dim cKeepDiscValue As Double
    Dim tKeeptCardID As String
    Dim tKeeptCardType As String
    Dim tKeeptCardKey As String
    Dim tKeepttStaChkExpire As String
    Dim dKeepdCardExpDate As Date
    Dim cKeepcDiscOntopAlw  As Double
    Dim cSumOntopRegCal As Double
    Dim cSumOntopProCal As Double
    Dim cSumOntopBothCal As Double
    Dim cBlueCal As Double
    Dim cYellowCal As Double
    Dim cBlueDisc As Double
    Dim cYellowDisc As Double
    Dim cKeepcDiscPer  As Double
    'เก็บไว้ เพราะ Recal แล้วจะหาย
    Dim nSeqNoLast As Long '*KT 51-10-24 V 3.0.141/V 3.0.200
    
    Dim orsSKUOnTop  As ADODB.Recordset
    Dim orsDeptOnTop   As ADODB.Recordset
    Dim tFTSkuCode As String
    Dim bSubAllTier As Boolean
    Dim bStep As Boolean
    Dim bCal As Boolean
    
    Dim cDiscTierAmt As Double
    Dim cSourceAmtAvg As Double
    
    tKeeptCardID = tC_CardID
    tKeeptCardType = tC_CardType
    tKeeptCardKey = tC_CardKey
    cKeepDiscValue = cC_DiscValue
    tKeepttStaChkExpire = tC_StaChkExpire
    dKeepdCardExpDate = dC_CardExpDate
    cKeepcDiscOntopAlw = cDiscOntopAlw
    cKeepcDiscPer = cC_DiscPer
    
    cC_AmtAllMall = 0
    If InStr(1, tCS_CNNotAllowTenderMall, tVB_Transaction) > 0 Then Exit Function '*TW 51-08-19
    On Error GoTo Err_Handler
    cVB_CNSumOntopMaxCal = 0
    'Move data to temp
    Screen.MousePointer = 11
    nC_ENDiscType = nEN_Disc15Ontop
    cOldDisc = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCScdAmt", tVB_CNTblCD, "WHERE " & tVB_CNWheBill & " AND  ISNULL(FTSdtStkCode,'') <> '1'")) '*KT 52-11-25 Deal&Promo
    cNewDisc = cOldDisc
    If (IsMissing(tC_TmnNum) Or tC_TmnNum = "") Or IsMissing(dC_TranDate) Or (IsMissing(tC_TransNo) Or tC_TransNo = "") Then
        Call SP_MSGnShowing(tC_Msg001, tCS_CNMsgWarn)
        DIScGetOntop = 0
        Screen.MousePointer = 0
        Exit Function
    End If
    Call C_DISxGetNextAct
        
    If UCase(tC_StaChkExpire) = "Y" Then
        dCurrentExpDate = DateAdd("m", 1, dC_CardExpDate) - 1
         '*KT 52-06-04 ตรวจสอบวันที่หมดอายุของ point ตามที่กำหนดใน parameter
        If Format(dCurrentExpDate, "YYYY/MM/DD") < Format(SP_GETdChkDateDisc, "YYYY/MM/DD") Then '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  ใส่ Format mm/dd/yyyyy เพราะ dVB_CurrentDate มีเวลาเข้ามาด้วย Communication Sheet PH2-ML-SF-027 Credit Expire.xls
            Screen.MousePointer = 0:
            DIScGetOntop = 0
            cVB_CNSumOntopMaxCal = 0
            cSumCal = 0
            cC_AmtAllMall = 0
            Exit Function
        End If
    End If
    
    If C_CHKbAlwOntop(nFNCdcOntopRate, tFTCdcMallCard, tFTCdcOnTop, cFCCdcAmtMax, cFCCdcAmtMin) Then
        '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
        'ตรวจสอบว่าเครื่องจุดขายอนุญาตให้ได้รับส่วนลด OnTop หรือไม่
        If bC_POSNoOnTopFlag Then
            If Not C_CHKbAlwOntopByPOSNo(tVB_CNTerminalNum, tVB_CNTerminalFloor, tC_CardType, tVB_TenderCode) Then
                Screen.MousePointer = 0:
                DIScGetOntop = 0
                cVB_CNSumOntopMaxCal = 0
                cSumCal = 0
                cC_AmtAllMall = 0
                Exit Function
            End If
        End If
    
        'Calculate from SUBTOTALALL
        bSubAllTier = IIf(SP_FEDtChkString(oC_RsTempOnTop, "FTCdcStaSubAllOnTop") = "Y", True, False)
        If bSubAllTier Then
            If cC_SubTotalAll >= cFCCdcAmtMin Then
                bCal = True
            Else
                bCal = False
            End If
        Else
            If cC_SubTotal >= cFCCdcAmtMin Then
                bCal = True
            Else
                bCal = False
            End If
        End If
        If Not bCal Then
            Screen.MousePointer = 0:
            DIScGetOntop = 0
            cVB_CNSumOntopMaxCal = 0
            cSumCal = 0
            cC_AmtAllMall = 0
            Exit Function
        End If
    
        bUsePro = bC_OnTopUsePro
        bUseReg = bC_OnTopUseReg
        Call C_ONTxGetDT(orsDT)
        If Not (orsDT.BOF And orsDT.EOF) Then
            If UCase(tFTCdcMallCard) <> "Y" Then 'if not mall then
            
'                Do While Not orsDT.EOF 'while
'                    'removeMall
'                    nFNSdtSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                    Call C_DATxDelMallCrdFromOntop(nFNSdtSeqNo) 'คุณกำพลยืนยัน
'                    orsDT.MoveNext
'                Loop    'loop

'*KT 52-04-21 ทำการคำนวนส่วนลด Mall Card ผิดในกรณีที่ ชำระด้วย Credit Card ที่ Not Allow Mall Card  จะต้องคำนวนเช่นเดียวกับ Special Coupon / Special Gift Voucher  ตามเอกสาร phase 1.8 UAT status on 20-04-2009
                Call oDisc.DIScSpecial("T002")

'                Call C_DATxSumDT
'                Call C_UPDxLabel
'                Call DISxPrcRecal("", 0, True)
                'เก็บไว้ เพราะ Recal แล้วจะหาย
                nC_ENDiscType = nEN_Disc15Ontop
                tC_CardID = tKeeptCardID
                tC_CardType = tKeeptCardType
                tC_CardKey = tKeeptCardKey
                cC_DiscValue = cKeepDiscValue
                tC_StaChkExpire = tKeepttStaChkExpire
                dC_CardExpDate = dKeepdCardExpDate
                cDiscOntopAlw = cKeepcDiscOntopAlw
                cC_DiscPer = cKeepcDiscPer
                cC_DiscPerOrg = cKeepcDiscPer
            End If
            cNewDisc = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCScdAmt", tVB_CNTblCD, "WHERE " & tVB_CNWheBill & "AND  ISNULL(FTSdtStkCode,'') <> '1'")) '*KT 52-11-25 Deal&Promo
            Call C_ONTxGetDT(orsDT)
            
            If bC_OnTopDepAlw Then
                'Record Sku Alw,Not Alw
                Call SP_DATxRsClose(orsSKUOnTop)
                Call C_GETxSKUOnTop(orsSKUOnTop, tC_CardKey, tC_CardType)
                'Record Dept/Brand Code Alw,Not Alw
                 Call SP_DATxRsClose(orsDeptOnTop)
                 Call C_GETxDeptOnTop(orsDeptOnTop, tC_CardKey, tC_CardType)
             End If
            
            orsDT.MoveFirst
            Do While Not orsDT.EOF 'while
                nFNSdtSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
                cBlueCal = SP_FEDvChkNumeric(orsDT, "FCSdtBluePrice")
                cYellowCal = SP_FEDvChkNumeric(orsDT, "FCSdtYellowPrice")
                cBlueDisc = C_ONTcDiscOldOntop(nFNSdtSeqNo, True) 'หาส่วนลดที่เป็น Ontop
                cYellowDisc = C_ONTcDiscOldOntop(nFNSdtSeqNo, False) 'หาส่วนลดที่เป็น Ontop
                If bC_OnTopDepAlw Then
                    tFTSkuDepCode = SP_FEDtChkString(orsDT, "FTSkuDepCode")
                    tFTSkuBrdCode = SP_FEDtChkString(orsDT, "FTSkuBrdCode")
                    tFTSkuCode = SP_FEDtChkString(orsDT, "FTSkuCode")
                    'By SKU Code
                    If C_ONTbExceptOntopSKU(orsSKUOnTop, tFTSkuCode, bStep) Then
                        'Allow SKU Code  นั้นไม่ขึ้นกับ reg ,promo ในส่วนข้างบน
                        cSumOntopBothCal = cSumOntopBothCal + (cYellowCal - cYellowDisc) + (cBlueCal - cBlueDisc)
                        cSumOntopMaxCal = cSumOntopMaxCal + (cYellowCal - cYellowDisc) + (cBlueCal - cBlueDisc)
                    End If
                    'By Dept/Brand Code
                    If bStep Then
                        If C_ONTbExceptOntop(orsDeptOnTop, tFTSkuDepCode, tFTSkuBrdCode, bUseReg, bUsePro, bStep) Then
                            cSumOntopBothCal = cSumOntopBothCal + IIf(bUsePro, cYellowCal - cYellowDisc, 0) + IIf(bUseReg, cBlueCal - cBlueDisc, 0)
                            cSumOntopMaxCal = cSumOntopMaxCal + IIf(bUsePro, cYellowCal - cYellowDisc, 0) + IIf(bUseReg, cBlueCal - cBlueDisc, 0)
                        End If
                    End If
                Else
                    cSumOntopBothCal = cSumOntopBothCal + IIf(bC_OnTopUsePro, cYellowCal - cYellowDisc, 0) + IIf(bC_OnTopUseReg, cBlueCal - cBlueDisc, 0)
                    cSumOntopMaxCal = cSumOntopMaxCal + IIf(bC_OnTopUsePro, cYellowCal - cYellowDisc, 0) + IIf(bC_OnTopUseReg, cBlueCal - cBlueDisc, 0)
                End If
                orsDT.MoveNext
            Loop    'loop
            
            ''*KT 54-06-01 V.4.0.0 Req.19 OnTop Specify POSNo. *หา rate % ส่วนลดตาม Tier ใหม่
            If SP_FEDtChkString(oC_RsTempOnTop, "FTCdcStaTier") = "Y" Then
                    If cC_DiscValue < cVB_CNOntopMaxCal Then
                        Call C_GETbRateOntopByTier(nFNCdcOntopRate, cC_DiscValue, tC_CardKey, tC_CardType)
                    Else
                        Call C_GETbRateOntopByTier(nFNCdcOntopRate, cSumOntopMaxCal, tC_CardKey, tC_CardType)
                    End If
                    cFCCdcAmtMax = IIf(cFCCdcAmtMax > (cKeepDiscValue * 100 / (100 - nFNCdcOntopRate)), (cKeepDiscValue * 100 / (100 - nFNCdcOntopRate)), cFCCdcAmtMax)
            Else
                    Call C_GETbRateOntopByTierAmt(cDiscTierAmt, cC_DiscValue, tC_CardKey, tC_CardType)
                    cFCCdcAmtMax = cC_DiscValue
            End If
            
            If cFCCdcAmtMax < cSumOntopMaxCal Then
                cSumOntopMaxCalAlw = cFCCdcAmtMax
            Else
                cSumOntopMaxCalAlw = cSumOntopMaxCal
            End If
            
            If cSumOntopMaxCal > 0 And cSumOntopMaxCalAlw > 0 Then
                orsDT.MoveFirst
                Do While Not orsDT.EOF 'while
                    nFNSdtSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
                    cBlueCal = SP_FEDvChkNumeric(orsDT, "FCSdtBluePrice")
                    cYellowCal = SP_FEDvChkNumeric(orsDT, "FCSdtYellowPrice")
                    cBlueDisc = C_ONTcDiscOldOntop(nFNSdtSeqNo, True) 'หาส่วนลดที่เป็น Ontop
                    cYellowDisc = C_ONTcDiscOldOntop(nFNSdtSeqNo, False) 'หาส่วนลดที่เป็น Ontop

                    If bC_OnTopDepAlw Then
                        tFTSkuDepCode = SP_FEDtChkString(orsDT, "FTSkuDepCode")
                        tFTSkuBrdCode = SP_FEDtChkString(orsDT, "FTSkuBrdCode")
                        tFTSkuCode = SP_FEDtChkString(orsDT, "FTSkuCode")
                       'By SKU Code
                        If C_ONTbExceptOntopSKU(orsSKUOnTop, tFTSkuCode, bStep) Then
                            'Allow SKU Code  นั้นไม่ขึ้นกับ reg ,promo ในส่วนข้างบน
                            cSumSourceAmt = (cYellowCal - cYellowDisc) + (cBlueCal - cBlueDisc)
                        End If
                        'By Dept/Brand Code
                        If bStep Then
                            If C_ONTbExceptOntop(orsDeptOnTop, tFTSkuDepCode, tFTSkuBrdCode, bUseReg, bUsePro, bStep) Then
                                  cSumSourceAmt = IIf(bUsePro, cYellowCal - cYellowDisc, 0) + IIf(bUseReg, cBlueCal - cBlueDisc, 0)
                            End If
                        End If
                    Else
                        cSumSourceAmt = IIf(bC_OnTopUsePro, cYellowCal - cYellowDisc, 0) + IIf(bC_OnTopUseReg, cBlueCal - cBlueDisc, 0)
                    End If
                    If cSumOntopMaxCalAlw < cSumSourceAmt Then
                        '*KT 51-10-24 V 3.0.141/V 3.0.200
                        '---------------------------------------------------------------------------------
                        'cSumOntopBothCal = cSumOntopBothCal * cSumOntopMaxCalAlw / cSumOntopMaxCal
                       ' cSumOntopProCal = cSumOntopProCal * cSumOntopMaxCalAlw / cSumOntopMaxCal
                       ' cSumOntopRegCal = cSumOntopRegCal * cSumOntopMaxCalAlw / cSumOntopMaxCal
                        'cSourceAmt = cSumSourceAmt * cSumOntopMaxCalAlw / cSumOntopMaxCal
                        '--------------------------------------------------------------------------------
                        cSourceAmt = cSumSourceAmt
                    Else
                        cSourceAmt = cSumSourceAmt
                    End If
                    If cSourceAmt > 0 Then
                        cBlueDisc = 0 '(cSumOntopRegCal + (cBlueCal / (cYellowCal + cBlueCal)) * cSumOntopBothCal) * nFNCdcOntopRate / 100 '*KT 51-10-24 V 3.0.141/V 3.0.200
                        cYellowDisc = 0 ' (cSumOntopProCal + (cYellowCal / (cYellowCal + cBlueCal)) * cSumOntopBothCal) * nFNCdcOntopRate / 100 '*KT 51-10-24V 3.0.141/V 3.0.200
                        nFNSdtSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
                        '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO ***Tier Amount or Tier Persent
                        If SP_FEDtChkString(oC_RsTempOnTop, "FTCdcStaTier") = "Y" Then
                            cDisAmt = SP_FMTcCurPoint((cSourceAmt * cSumOntopMaxCalAlw / cSumOntopMaxCal) * nFNCdcOntopRate / 100, True, False) 'cSourceAmt * nFNCdcOntopRate / 100 '*KT 51-10-13 V 3.0.134/V 3.0.200
                        Else
                            cSourceAmtAvg = SP_FMTcCurPoint(cSourceAmt * cSumOntopMaxCalAlw / cSumOntopMaxCal, True, False)
                            cDisAmt = SP_FMTcCurPoint(cSourceAmtAvg * cDiscTierAmt / cSumOntopMaxCalAlw, True, False)
                        End If
                        cC_DiscPer = nFNCdcOntopRate
                        cC_DiscPerOrg = nFNCdcOntopRate
                        cLastDiscAmt = cSumOntopMaxDisc
                        nLastFNSdtSeqNo = nFNSdtSeqNo
                        cSumOntopMaxDisc = cSumOntopMaxDisc + SP_FMTcCurPoint(cDisAmt, True, False)
                         '*KT 51-10-13 V 3.0.134/V 3.0.200
                         '--------------------------------------------------------------------
                        orsDT.MoveNext
                        If orsDT.EOF Then
                            cC_DiscOntopAlw = SP_FMTcCurPoint(cC_DiscOntopAlw, True, False)
                            cSumOntopMaxDisc = SP_FMTcCurPoint(cSumOntopMaxDisc, True, False)
                            If cC_DiscOntopAlw <> cSumOntopMaxDisc Then
                                cDisAmt = SP_FMTcCurPoint(cDisAmt + (cC_DiscOntopAlw - cSumOntopMaxDisc), True, False) '*KT 51-10-27 V 3.0.143/V 3.0.200
                                cSumOntopMaxDisc = cC_DiscOntopAlw
                            End If
                        End If
                        orsDT.MovePrevious
                         '----------------------------------------------------------------------
                        'Call C_DATxInsertDiscCD(cDisAmt, nFNSdtSeqNo, "", 0, cSourceAmt, 0, cBlueDisc, cYellowDisc)
                        Call C_DATxInsertDiscCD(cDisAmt, nFNSdtSeqNo, "", 1, cSourceAmt, 1, cBlueDisc, cYellowDisc) '*KT 51-12-15 Phase 2 ข้อ 26 Default FCScdQty = 1 เสมอ
                        nSeqNoLast = nFNSdtSeqNo
                    Else
                         '*KT 51-10-24 V 3.0.141/V 3.0.200
                         '--------------------------------------------------------------------
                        orsDT.MoveNext
                        If orsDT.EOF Then
                            cC_DiscOntopAlw = SP_FMTcCurPoint(cC_DiscOntopAlw, True, False)
                            cSumOntopMaxDisc = SP_FMTcCurPoint(cSumOntopMaxDisc, True, False)
                            If cC_DiscOntopAlw <> cSumOntopMaxDisc Then
                                cDisAmt = SP_FMTcCurPoint(cC_DiscOntopAlw - cSumOntopMaxDisc, True, False) '*KT 51-10-27 V 3.0.143/V 3.0.200
                                cSumOntopMaxDisc = cC_DiscOntopAlw
                                tSql = "UPDATE " & tVB_CNTblCD & " WITH(ROWLOCK) SET FCScdAmt =FCScdAmt + " & cDisAmt
                                tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
                                tSql = tSql & vbCrLf & " AND  FTTmnNum = '" & tVB_CNTerminalNum & "'"
                                tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
                                tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nSeqNoLast '*KT 51-10-27 V 3.0.143/V 3.0.200
                                tSql = tSql & vbCrLf & " AND FTScdCardID = '" & tC_CardID & "'"
                                tSql = tSql & vbCrLf & " AND FNScdSeqNo = " & SP_VALcText2Double(SP_SQLtFunction("MAX", "FNScdSeqNo", "" & tVB_CNTblCD & "", "WHERE " & tVB_CNWheBill & " AND FTScdCardID = '" & tC_CardID & "'"))
                                tSql = tSql & vbCrLf & " AND FNDctNo = " & nC_ENDiscType
                                Call SP_SQLbExecute(tSql)
                            End If
                        End If
                        orsDT.MovePrevious
                        '--------------------------------------------------------------------
                    End If
                    orsDT.MoveNext
                Loop    'loop
            End If
            Call C_DATxSumDT
            'Call C_UPDxLabel
        End If
    End If
    If cC_DiscOntopAlw <> cSumOntopMaxDisc Then
        If nLastFNSdtSeqNo > 0 Then
            cSumOntopMaxDisc = cC_DiscOntopAlw
            tSql = "UPDATE " & tVB_CNTblCD & " WITH(ROWLOCK) "
            tSql = tSql & vbCrLf & "SET FCScdAmt = " & SP_FMTcCurPoint(cC_DiscOntopAlw, True, False) - SP_FMTcCurPoint(cLastDiscAmt, True, False)
            tSql = tSql & vbCrLf & ",FCScdAmtAF = FCScdAmtB4 - " & SP_FMTcCurPoint(cC_DiscOntopAlw, True, False) - SP_FMTcCurPoint(cLastDiscAmt, True, False)
            tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
            tSql = tSql & vbCrLf & "AND FNDctNo = " & nEN_Disc15Ontop
            tSql = tSql & vbCrLf & "AND FTScdCardID = '" & tC_CardID & "'"
            tSql = tSql & vbCrLf & "AND FNSdtSeqNo = " & nLastFNSdtSeqNo
            tSql = tSql & vbCrLf & "AND FNScdSeqNo = " & SP_VALcText2Double(SP_SQLtFunction("MAX", "FNScdSeqNo", "" & tVB_CNTblCD & "", "WHERE " & tVB_CNWheBill & " AND FTScdCardID = '" & tC_CardID & "'"))
            Call SP_SQLbExecute(tSql)
            Call C_DATxSumDT
        End If
    End If
    cSumOntopMaxCalAlw = SP_FMTcCurPoint(cSumOntopMaxCalAlw, True, False)
    cDiscOntopAlw = cSumOntopMaxCalAlw
    cSumCal = SP_FMTcCurPoint(cSumOntopMaxDisc, True, False)
    cDiscValue = SP_FMTcCurPoint(cDiscOntopAlw - cSumCal, True, False)
    cC_AmtAllMall = SP_FMTcCurPoint(cOldDisc - cNewDisc, True, False)
    DIScGetOntop = SP_FMTcCurPoint(cSumCal, True, False)
    Call SP_DATxRsNothing(orsDT)
    Screen.MousePointer = 0
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    cC_AmtAllMall = cOldDisc - cNewDisc
    DIScGetOntop = 0
    Call SP_DATxRsNothing(orsDT)
    Screen.MousePointer = 0
End Function

Private Function C_SPCbAllMall(ByVal ptTenderType As String) As Boolean
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    On Error GoTo Err_Handler
    tSql = "SELECT FTTdmStaMallCard"
    tSql = tSql & vbCrLf & "FROM TCNMTenderMtn WITH(NOLOCK) "
    '*Eaw 56-01-30 Point ตาม tender
'    Select Case ptTenderType
'        Case "T002"
'            tSql = tSql & vbCrLf & "WHERE FTTdmCode = '002'"
'        Case "T005", "T014" '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher"
'            tSql = tSql & vbCrLf & "WHERE FTTdmCode = '007'"
'        Case "T007"
'            tSql = tSql & vbCrLf & "WHERE FTTdmCode = '007'"
'
'    End Select
    Select Case UCase(ptTenderType)
        Case UCase("T002"), UCase("T005"), UCase("T014"), UCase("T007")
            tSql = tSql & vbCrLf & "WHERE FTTdmCode = '" & ptTenderType & "'"
    End Select
    '-------------------------------------------------------------------------------------------------------------------------
    Call SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) 'read master
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        If UCase(SP_FEDtChkString(orsTemp, "FTTdmStaMallCard")) = "Y" Then
            C_SPCbAllMall = True
        Else
            C_SPCbAllMall = False
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    C_SPCbAllMall = False
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function DIScSpecial(ByVal ptTenderType As String) As Double
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 52-04-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cTotalPmt As Double
    Dim cTotalMalCpn As Double
    Dim cY1 As Double
    Dim cPerMCard As Double
    Dim tOrgCD As String
    Dim tOrgDT As String
    Dim tTmpCD As String
    Dim tTmpDT As String
    Dim cOldDisc As Double
    Dim cNewDisc As Double
    '*TW 51-07-30 ML-SF-421 Special Gift Voucher is Wrong When not Allow Mall Card กรณี จ่าย Double
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim orsLabel As ADODB.Recordset
    Dim cCalPrice As Double
    Dim cSumBlue As Double
    Dim cSumYellow As Double
    Dim b2ndRound As Boolean
    Dim cSumSpecial As Double
    Dim tWhe As String
    Dim cTmpSaleAmt As Double
    Dim cTmpBlue As Double
    Dim nSeqNo As Long
    Dim bClearMall As Boolean
    '--------------------------------------------
    Dim cAddRegPrice As Double
    Dim bAllowMall As Boolean
'    Dim orsKeepOld As ADODB.Recordset
    Dim orsOldSpec As ADODB.Recordset
    Dim cAddSpec As Double
    On Error GoTo Err_Handler
    If InStr(1, tCS_CNNotAllowTenderMall, tVB_Transaction) > 0 Then Exit Function '*TW 51-08-19
    tOrgCD = tVB_CNTblCD
    tOrgDT = tVB_CNTblDT
    tTmpCD = tVB_CNTblCD & "Ontop"
    tTmpDT = tVB_CNTblDT & "Ontop"
    nC_ENDiscType = nEN_Disc20Special
    If (IsMissing(tC_TmnNum) Or tC_TmnNum = "") Or IsMissing(dC_TranDate) Or (IsMissing(tC_TransNo) Or tC_TransNo = "") Then
        Call SP_MSGnShowing(tC_Msg001, tCS_CNMsgWarn)
        DIScSpecial = 0
        Exit Function
    End If
    cPerMCard = SP_VALcText2Double(SP_SQLtFunction("MAX", "FCScdDiscPer", tOrgCD, "WHERE FNDctNo = " & nEN_Disc11MallCrd))
    cOldDisc = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCScdAmt", tOrgCD, "WHERE " & tVB_CNWheBill))
    bAllowMall = C_SPCbAllMall(ptTenderType) 'E
    If ptTenderType = "T002" Then bAllowMall = False 'บังคับว่ากรณีที่มี Recall Mall Card ใหม่เมื่อจ่ายด้วย Credit Card นั้นให้ bAllowMall = false เมื่ออยุ่ใน Produdure DIScSpecial
    If cPerMCard > 0 And Not bAllowMall Then
    '*TW 51-07-30 ML-SF-421 Special Gift Voucher is Wrong When not Allow Mall Card กรณี จ่าย Double
        tSql = "SELECT FNSrcSeqNo,FCSrcAmt" 'E
        tSql = tSql & vbCrLf & "FROM " & tVB_CNTblRC & " WITH(NOLOCK)  "
        tSql = tSql & vbCrLf & "WHERE ((FTTdmCode IN ('T005','T007','T014')) OR (FTTdmCode ='T002' AND ISNULL(FTSrcStaPrc ,'') = '1'))" '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher"
        tSql = tSql & vbCrLf & "AND  " & tVB_CNWheBill
        tSql = tSql & vbCrLf & "ORDER BY FNSrcSeqNo"
        If orsTemp Is Nothing Then Set orsTemp = New ADODB.Recordset
        If orsTemp.State = adStateOpen Then orsTemp.Close
        Call orsTemp.Open(tSql, ocnAll, adOpenStatic, adLockOptimistic)
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            b2ndRound = True
            'ดึงเอาข้อมูลก่อนเข้ามา มาใช้
            
            'arrVB_CNKeepDT,arrVB_CNKeepCD Find ดูทั้งโปรแกรมไม่มีการใช้งานเลยทำการ Conment ไว้ก่อน
'            If Not (C_TMPbArray2DB(arrVB_CNKeepDT, tOrgDT)) Then Exit Function
'            If Not (C_TMPbArray2DB(arrVB_CNKeepCD, tOrgCD)) Then Exit Function
        Else
            b2ndRound = False
            'เอา ข้อมูลก่อนเข้ามา เก็บไว้
            
             'arrVB_CNKeepDT,arrVB_CNKeepCD Find ดูทั้งโปรแกรมไม่มีการใช้งานเลยทำการ Conment ไว้ก่อน
'            tSql = "SELECT * FROM " & tOrgCD 'tpscd08107tmp
'            Call SP_SQLbExecute(tSql, orsKeepOld)
'            If Not C_TMPbOrs2Array(orsKeepOld, arrVB_CNKeepCD) Then Exit Function
'            tSql = "SELECT * FROM " & tOrgDT 'tpsdt08107tmp
'            Call SP_SQLbExecute(tSql, orsKeepOld)
'            If Not C_TMPbOrs2Array(orsKeepOld, arrVB_CNKeepDT) Then Exit Function
'            Call SP_DATxRsClose(orsKeepOld)
        End If
        orsTemp.ActiveConnection = Nothing
        If tVB_Transaction = Format(tEN_TrnAddDeposit, "00") Then
            tSql = "WHERE ((FTTdmCode IN ('T005','T007','T014')) OR (FTTdmCode ='T002' AND ISNULL(FTSrcStaPrc ,'') = '1'))" '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher"
            tSql = tSql & vbCrLf & "AND TPSTDepositHD.FTShdDepReTransNo = '" & tVB_TransNoAddDpt & "'"
            tSql = tSql & vbCrLf & "AND TPSTDepositHD.FTShdDepRefTmnNum = '" & tVB_TerminalNoAddDpt & "'"
            tSql = tSql & vbCrLf & "AND TPSTDepositHD.FDShdDepReTransDate = '" & Format(dVB_TransDateAddDpt, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(TPSTDepositHD.FTShdStaCancel,'N') = 'N'" '*KT 52-08-11 PH2-ML-SB- 108 แก้ไขกรณีที่ Deposit Create >> Redeem >> IPV Redeem >> Redeem แล้วยอด Sub Reg ผิด(ใช้ Tender Special Gift Voucher,Special Coupon,Credit Card)
            cAddSpec = CDbl(Val(Format(SP_SQLtFunction("SUM", "FCSrcAmt", "TPSTDepositRC INNER JOIN TPSTDepositHD ON TPSTDepositRC.FTTmnNum = TPSTDepositHD.FTTmnNum AND TPSTDepositRC.FTShdTransNo = TPSTDepositHD.FTShdTransNo AND TPSTDepositRC.FDShdTransDate = TPSTDepositHD.FDShdTransDate", tSql, oDB.ocnOnLine))))
        Else
            If tVB_Transaction = Format(tEN_TrnRedeem, "00") Then 'Or tVB_Transaction = Format(tEN_TrnRedeem, "00") Then
                tSql = "WHERE ((FTTdmCode IN ('T005','T007','T014')) OR (FTTdmCode ='T002' AND ISNULL(FTSrcStaPrc ,'') = '1'))" '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher"
                tSql = tSql & vbCrLf & "AND TPSTDepositHD.FTShdDepReTransNo = '" & tVB_TransNoRedeem & "'"
                tSql = tSql & vbCrLf & "AND TPSTDepositHD.FTShdDepRefTmnNum = '" & tVB_TerminalNoRedeem & "'"
                tSql = tSql & vbCrLf & "AND TPSTDepositHD.FDShdDepReTransDate = '" & Format(dVB_TransDateRedeem, "YYYY/MM/DD") & "'"
                 tSql = tSql & vbCrLf & "AND ISNULL(TPSTDepositHD.FTShdStaCancel,'N') = 'N'" '*KT 52-08-11 PH2-ML-SB- 108 แก้ไขกรณีที่ Deposit Create >> Redeem >> IPV Redeem >> Redeem แล้วยอด Sub Reg ผิด(ใช้ Tender Special Gift Voucher,Special Coupon,Credit Card)
                cAddSpec = CDbl(Val(Format(SP_SQLtFunction("SUM", "FCSrcAmt", "TPSTDepositRC INNER JOIN TPSTDepositHD ON TPSTDepositRC.FTTmnNum = TPSTDepositHD.FTTmnNum AND TPSTDepositRC.FTShdTransNo = TPSTDepositHD.FTShdTransNo AND TPSTDepositRC.FDShdTransDate = TPSTDepositHD.FDShdTransDate", tSql, oDB.ocnOnLine))))
            Else
                cAddSpec = 0
            End If
        End If
        orsTemp.AddNew
        orsTemp.Fields("FNSrcSeqNo") = 999
        orsTemp.Fields("FCSrcAmt") = cC_DiscValue + cAddSpec
        orsTemp.MoveFirst
        If C_TMPbCreate(tOrgCD, tTmpCD) And C_TMPbCreate(tOrgDT, tTmpDT) Then
            tVB_CNTblCD = tTmpCD
            tVB_CNTblDT = tTmpDT
            cC_DiscValue = 0
            Do While Not orsTemp.EOF
                cC_DiscValue = cC_DiscValue + SP_FEDvChkNumeric(orsTemp, "FCSrcAmt")
                orsTemp.MoveNext
            Loop

            If C_GETbYellowLabel(orsLabel, False) Then
                cSumBlue = 0
                cSumYellow = 0
                Do While Not orsLabel.EOF
                    'If UCase(SP_FEDtChkString(orsLabel, "FTSkuStaMallEmpCard")) = "Y" Then
                        cCalPrice = SP_FEDvChkNumeric(orsLabel, "FCSdtBluePrice")
                        cSumBlue = cSumBlue + cCalPrice
                    'Else
                        cCalPrice = SP_FEDvChkNumeric(orsLabel, "FCSdtYellowPrice")
                        cSumYellow = cSumYellow + cCalPrice
                    'End If
                    orsLabel.MoveNext
                Loop
            End If
            
            
            cTotalPmt = cSumYellow
            cAddRegPrice = IIf(cC_DiscValue - cTotalPmt > 0, cC_DiscValue - cTotalPmt, 0)
            cY1 = cTotalPmt ' - cTotalMalCpn'*TW 51-09-01
            If cC_DiscValue > cY1 Then
                Call C_CLRxMCrdSpecial(False) 'E ?
                Call C_DATxSumDT
                Call C_UPDxLabel
                bClearMall = True
                Call C_DISxGetNextAct
                Call C_DISxSpecial
            End If
            
            If bClearMall Then
                Call DISxPrcRecal(tOrgCD, cAddRegPrice)
            End If

            Call C_CLRxSpecial
            If C_TMPbRet2Org(tOrgCD, tTmpCD) And C_TMPbRet2Org(tOrgDT, tTmpDT) Then
                tVB_CNTblCD = tOrgCD
                tVB_CNTblDT = tOrgDT
            Else
                'ควรมีการฟ้องอะไรตรงนี้บ้าง
                tVB_CNTblCD = tOrgCD
                tVB_CNTblDT = tOrgDT
            End If
            
            cNewDisc = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCScdAmt", tOrgCD, "WHERE " & tVB_CNWheBill))
            DIScSpecial = SP_FMTcCurPoint(cOldDisc - cNewDisc, True, False)
             ''*KT 51-09-15 แก้ไขเรื่องจ่ายด้วย Special Gift Voucher/Special Coupon หลังหักส่วนลด Mall Cardด้วยสินค้าป้ายน้ำเงินอย่างเดียว
             '--------------------------------------------------------------------------------
            If bClearMall Then
                Call C_UPDxLabel
            Else
                Call C_UPDxLabel(True, IIf(cAddRegPrice < 0, 0, cAddRegPrice))
            End If
            '---------------------------------------------------------------------------------
        End If
    Else
        tSql = "SELECT FNSrcSeqNo,FCSrcAmt"
        tSql = tSql & vbCrLf & "FROM " & tVB_CNTblRC & " WITH(NOLOCK)  "
        tSql = tSql & vbCrLf & "WHERE FTTdmCode IN ('T005','T007','T014')" '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher"
        tSql = tSql & vbCrLf & "AND  FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & "ORDER BY FNSrcSeqNo"
        If orsTemp Is Nothing Then Set orsTemp = New ADODB.Recordset
        If orsTemp.State = adStateOpen Then orsTemp.Close
        Call orsTemp.Open(tSql, ocnAll, adOpenStatic, adLockOptimistic)
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            b2ndRound = True
        Else
            b2ndRound = False
        End If
        orsTemp.ActiveConnection = Nothing
        If tVB_Transaction = Format(tEN_TrnAddDeposit, "00") Then
            tSql = "WHERE FTTdmCode IN ('T005','T007','T014')" '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher"
            tSql = tSql & vbCrLf & "AND TPSTDepositHD.FTShdDepReTransNo = '" & tVB_TransNoAddDpt & "'"
            tSql = tSql & vbCrLf & "AND TPSTDepositHD.FTShdDepRefTmnNum = '" & tVB_TerminalNoAddDpt & "'"
            tSql = tSql & vbCrLf & "AND TPSTDepositHD.FDShdDepReTransDate = '" & Format(dVB_TransDateAddDpt, "YYYY/MM/DD") & "'"
            cAddSpec = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSrcAmt", "TPSTDepositRC INNER JOIN TPSTDepositHD ON TPSTDepositRC.FTTmnNum = TPSTDepositHD.FTTmnNum AND TPSTDepositRC.FTShdTransNo = TPSTDepositHD.FTShdTransNo AND TPSTDepositRC.FDShdTransDate = TPSTDepositHD.FDShdTransDate", tSql))
        Else
            If tVB_Transaction = Format(tEN_TrnRedeem, "00") Then 'Or tVB_Transaction = Format(tEN_TrnRedeem, "00") Then
                tSql = "WHERE FTTdmCode IN ('T005','T007','T014')" '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher"
                tSql = tSql & vbCrLf & "AND TPSTDepositHD.FTShdDepReTransNo = '" & tVB_TransNoRedeem & "'"
                tSql = tSql & vbCrLf & "AND TPSTDepositHD.FTShdDepRefTmnNum = '" & tVB_TerminalNoRedeem & "'"
                tSql = tSql & vbCrLf & "AND TPSTDepositHD.FDShdDepReTransDate = '" & Format(dVB_TransDateRedeem, "YYYY/MM/DD") & "'"
                cAddSpec = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSrcAmt", "TPSTDepositRC INNER JOIN TPSTDepositHD ON TPSTDepositRC.FTTmnNum = TPSTDepositHD.FTTmnNum AND TPSTDepositRC.FTShdTransNo = TPSTDepositHD.FTShdTransNo AND TPSTDepositRC.FDShdTransDate = TPSTDepositHD.FDShdTransDate", tSql))
            Else
                cAddSpec = 0
            End If
        End If
        orsTemp.AddNew
        orsTemp.Fields("FNSrcSeqNo") = 999
        orsTemp.Fields("FCSrcAmt") = cC_DiscValue + cAddSpec
        orsTemp.MoveFirst
        cC_DiscValue = 0
        Do While Not orsTemp.EOF
            cC_DiscValue = cC_DiscValue + SP_FEDvChkNumeric(orsTemp, "FCSrcAmt")
            orsTemp.MoveNext
        Loop
        If C_GETbYellowLabel(orsLabel, False) Then
            cSumBlue = 0
            cSumYellow = 0
            Do While Not orsLabel.EOF
                    'If UCase(SP_FEDtChkString(orsLabel, "FTSkuStaMallEmpCard")) = "Y" Then
                        cCalPrice = SP_FEDvChkNumeric(orsLabel, "FCSdtBluePrice")
                        cSumBlue = cSumBlue + cCalPrice
                    'Else
                        cCalPrice = SP_FEDvChkNumeric(orsLabel, "FCSdtYellowPrice")
                        cSumYellow = cSumYellow + cCalPrice
                    'End If
                orsLabel.MoveNext
            Loop
        End If
        cTotalPmt = cSumYellow
        cAddRegPrice = SP_FMTcCurPoint(cC_DiscValue - cTotalPmt, True, False)
        cAddRegPrice = IIf(cAddRegPrice > cSumBlue, cSumBlue, cAddRegPrice)
        '*KT 52-02-23 Communication Sheet PH2-ML-SF-073 Special gift voucher & Special coupon when tic Allow mall card wrong.xls
        'กรณีที่อนุญาติให้การชำระด้วย Special Gift Voucher และ Special Coupon นั้นแสดงว่าค่า DIScSpecial กลับไปต้องเป็น 0 จึงไม่ต้องปรับยอด Subtotal Reg ที่ Function C_UPDxLabel
        'PH 2.0-ML-SF-001 การคำนวนส่วนลด Mall Card ผิดในกรณีที่ ชำระด้วย Credit Card ที่ Not Allow Mall Card  จะต้องคำนวนเช่นเดียวกับ Special Coupon / Special Gift Voucher ของ Mail วันที่ 27/04/2009
        '*KT 52-06-08 กรณีที่ Special Gift Voucher และ Special Coupon  และ Credit Card ที่มี Flag Not Allow Mall Card ค่าของ Sub Reg ต้องลดด้วยแม้ว่าส่วนลดลดของ Mall Card จะไม่ลดก็ตาม
        'If DIScSpecial > 0 Then
        If Not bAllowMall Then
            Call C_UPDxLabel(True, IIf(cAddRegPrice < 0, 0, cAddRegPrice))
        End If
    End If
    Call SP_DATxRsNothing(orsTemp) '*TW 51-07-30 ML-SF-421 Special Gift Voucher is Wrong When not Allow Mall Card กรณี จ่าย Double
    Exit Function
Err_Handler:
    tVB_CNTblCD = tOrgCD
    tVB_CNTblDT = tOrgDT
    Call SP_DATxRsNothing(orsTemp) '*TW 51-07-30 ML-SF-421 Special Gift Voucher is Wrong When not Allow Mall Card กรณี จ่าย Double
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function


Public Sub DISxClearAll()
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tWhe As String
    cC_LastDiscount = 0
    cC_LastDiscountPerUnit = 0
    If Not bC_StaRecal Then
        tC_TmnNum = ""
        tC_TransNo = ""
        dC_TranDate = Date
        nC_ENLnDisType = nEN_LnDis0NoDis
        cC_DiscValue = 0
        cC_DiscPer = 0
        cC_DiscAmt = 0
        cC_DiscNew = 0
        nC_DTSeqNo = 0
    End If
    tC_TheBestType = ""
    cC_AlwDis = 0
    cC_AlwDisQty = 0
    nC_ENDiscType = nEN_Disc01Deal
    tC_RefCode = ""
    tC_CardID = ""
    tC_CardType = ""
    tC_CardKey = ""
    tC_CusID = ""
    dC_CardExpDate = "1980/01/01"
    nC_ActSeqNo = 0
    
    tC_StaChkDepart = ""
    tC_StaChkExpire = ""
    '*Ao 09-11-18
    cC_DiscPerOrg = 0
    cC_DiscAmtOrg = 0
    cC_DiscNewOrg = 0
    bC_PromoAftDeal = False '*KT 52-11-25 Deal&Promo
    tC_CardName = "" '*KT 52-11-26 PH2-ML-SF-170 ใน Ticket ไม่พิมพ์ชื่อพนักงานที่ใช้ส่วนลด Employee discount
    '_________________________
 End Sub

Public Sub DISxClearTrans()
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    If tVB_Running = "" Then Exit Sub
    tSql = "DELETE FROM " & tVB_CNTblCD & " "
    tSql = tSql & vbCrLf & " WHERE FNDctNo IN ("
    tSql = tSql & nEN_Disc01Deal
'    tSql = tSql & "," & nEN_Disc04CoopDisc
    tSql = tSql & "," & nEN_Disc07AutoTran
    tSql = tSql & "," & nEN_Disc08MallCpn
    tSql = tSql & "," & nEN_Disc09CoopCpn
    tSql = tSql & "," & nEN_Disc10CoopBrand
    tSql = tSql & "," & nEN_Disc11MallCrd
    tSql = tSql & "," & nEN_Disc12EmpCrd
    tSql = tSql & "," & nEN_Disc13TranDisc
    tSql = tSql & "," & nEN_Disc14TranMDown
    tSql = tSql & "," & nEN_Disc16FreeGift
    tSql = tSql & "," & nEN_Disc15Ontop
    tSql = tSql & "," & nEN_Disc23TranMDown1 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
    tSql = tSql & "," & nEN_Disc24TranMDown2
    tSql = tSql & "," & nEN_Disc25TranMDown3
    tSql = tSql & "," & nEN_Disc26TranMDown4
    tSql = tSql & "," & nEN_Disc27TranMDown5
    tSql = tSql & ") "
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    Call SP_SQLbExecute(tSql)
    Call C_DATxSumDT
End Sub

Public Sub DISxClearTender()
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    If tVB_Running = "" Then Exit Sub
    tSql = "DELETE FROM " & tVB_CNTblCD & " "
    tSql = tSql & vbCrLf & " WHERE FNDctNo IN ("
    tSql = tSql & "" & nEN_Disc16FreeGift
    tSql = tSql & "," & nEN_Disc15Ontop
    tSql = tSql & ") "
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    Call SP_SQLbExecute(tSql)
    Call C_DATxSumDT
End Sub

Public Sub DISxClearAutoTran()
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-21
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    If tVB_Running = "" Then Exit Sub
    tSql = "DELETE FROM " & tVB_CNTblCD & " "
    tSql = tSql & vbCrLf & " WHERE FNDctNo IN ("
    tSql = tSql & "" & nEN_Disc07AutoTran
    tSql = tSql & ") "
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    Call SP_SQLbExecute(tSql)
    Call C_DATxSumDT
End Sub

Private Function C_SETcDiscTypeRecal(ByVal porsTemp As ADODB.Recordset) As Double
    nC_ENLnDisType = SP_VALcText2Double(SP_FEDtChkString(porsTemp, "FTScdDiscType"))
    cC_DiscAmt = 0
    cC_DiscNew = 0
    cC_DiscPer = 0
    Select Case nC_ENLnDisType
        Case nEN_LnDis1Per
            cC_DiscPer = SP_FEDvChkNumeric(porsTemp, "FCScdDiscPerOrg")
            cC_DiscValue = cC_DiscPer
        Case nEN_LnDis2Amt
            cC_DiscAmt = SP_FEDvChkNumeric(porsTemp, "FCScdDiscAmtOrg")
            cC_DiscValue = cC_DiscAmt
            C_SETcDiscTypeRecal = cC_DiscAmt
        Case nEN_LnDis3NewPrice
            cC_DiscNew = SP_FEDvChkNumeric(porsTemp, "FCScdDiscNewOrg")
            cC_DiscValue = cC_DiscNew
            C_SETcDiscTypeRecal = SP_FEDvChkNumeric(porsTemp, "FCScdAmt")
        Case nEN_LnDis0NoDis
    End Select
End Function
                

Public Sub DISxPrcRecal(Optional ByVal ptOrgName$ = "", Optional ByVal pcSpecial As Double = 0, Optional ByVal pbOntop As Boolean = False, Optional ByVal pnSeqNo As Long = 0)
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  #
'*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตั้งต้น 3.0.19
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim tTerminalNo As String
    Dim tTransNo As String
    Dim dTransDate As Date
    Dim cCurrentDiscAmt As Double
    Dim cLastDisc As Double
    Dim orsRecal As ADODB.Recordset
    Dim nOldActSeqNo As Integer
    Dim orsMall As ADODB.Recordset
    Dim orsDT As ADODB.Recordset
    Dim bExit As Boolean
    
    'Return Mall Card
    Dim cDisAmt As Double
    Dim tEvtID As String
    Dim nDiscType As Long
    Dim tScdCardID As String
    Dim nSeqNo As Long
    Dim cDisQty As Double
    Dim cSdtSaleAmt As Double
    Dim cSdtQtyAllDisAvi As Double
    Dim tCrdType As String
    
    Dim cScdQty As Double
    Dim cScdDiscAmtByOne As Double
    
'    Dim orsDTTmp  As ADODB.Recordset
'    Dim cSdtQty As Double
    
On Error GoTo Err_Handler
    If (IsMissing(tC_TmnNum) Or tC_TmnNum = "") Or IsMissing(dC_TranDate) Or (IsMissing(tC_TransNo) Or tC_TransNo = "") Then
        Call SP_MSGnShowing(tC_Msg001, tCS_CNMsgWarn)
        Exit Sub
    End If
    bC_StaRecal = True
        
    If nC_ENDiscType <> nEN_Disc20Special Then   'FCScdDiscPerOrg
        tSql = "SELECT FNSdtSeqNo,FTScdDiscType,FCScdDiscPerOrg,FCScdDiscAmtOrg,FCScdDiscNewOrg,FDScdExpDate,FNScdActSeqNo,FNDctNo"
        tSql = tSql & vbCrLf & ", FTScdCardID, FTScdCardType, FTScdCardKey, FTScdStaChkExpire,FTScdStaChkDepart,FCScdMaxAlwDisc,FNScdSptSeqNo"
        tSql = tSql & vbCrLf & ", FCScdDiscPer,FCScdDiscAmt,FCScdDiscNew,FCScdAmt"
        tSql = tSql & vbCrLf & ", FCScdDiscPerOrg,FCScdDiscAmtOrg,FCScdDiscNewOrg" '*TW 51-06-05 เพื่อเรียกใช้ให้ได้ตัวเดิมเสมอ
        tSql = tSql & vbCrLf & ", FTDctRefCode ,FTScdCstID,FNScdSeqNo,FCScdQty"
        tSql = tSql & vbCrLf & ", FTSdtStkCode" '*KT 52-11-25 Deal&Promo
        tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD & " WITH(NOLOCK) "
        tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND FNDctNo <> " & nEN_Disc01Deal
        If pbOntop Then '*TW 51-08-28
            tSql = tSql & vbCrLf & " AND FNDctNo > " & nEN_Disc11MallCrd
           ' tSql = tSql & vbCrLf & " AND FNDctNo <> " & nEN_Disc15Ontop
            tSql = tSql & vbCrLf & " AND FNDctNo NOT IN (" & nEN_Disc15Ontop & "," & nEN_Disc21MPrice & "," & nEN_Disc22MDisc & ")" '*KT 52-09-24 Member Price + Consignment
        End If
        If pnSeqNo > 0 Then
            tSql = tSql & vbCrLf & " AND FNSdtSeqNo =" & pnSeqNo
        End If
        tSql = tSql & vbCrLf & " ORDER BY FNScdOrder,FNScdActSeqNo "
        If Not SP_DATbRsOpen(orsRecal, adUseClient, tSql, ocnAll, adOpenStatic, adLockOptimistic, -1) Then
            MsgBox "Error On Query", vbCritical 'tms_CN005249
            Exit Sub
        End If
        tSql = "DELETE "
        tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD & " WITH(ROWLOCK) "
        tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND FNDctNo <> " & nEN_Disc01Deal '*TW 51-07-04
        If pbOntop Then '*TW 51-08-28
            tSql = tSql & vbCrLf & " AND FNDctNo > " & nEN_Disc11MallCrd
'            tSql = tSql & vbCrLf & " AND FNDctNo <> " & nEN_Disc15Ontop
             tSql = tSql & vbCrLf & " AND FNDctNo NOT IN (" & nEN_Disc15Ontop & "," & nEN_Disc21MPrice & "," & nEN_Disc22MDisc & ")" '*KT 52-09-24 Member Price + Consignment
        End If
        If pnSeqNo > 0 Then
            tSql = tSql & vbCrLf & " AND FNSdtSeqNo =" & pnSeqNo
        End If
        Call SP_SQLbExecute(tSql)
        'Call DISbClearDiscDT '*KT 51-09-17  กรณีที่ใช้บัตรเครดิตที่เป็น ไม่ Flag Mall Card และ On Top
        Set orsRecal.ActiveConnection = Nothing
    Else 'For Special Coupon
        tSql = "SELECT FNSdtSeqNo,FTScdDiscType,FCScdDiscPerOrg,FCScdDiscAmtOrg,FCScdDiscNewOrg,FDScdExpDate,FNScdActSeqNo,FNDctNo"
        tSql = tSql & vbCrLf & ", FTScdCardID, FTScdCardType, FTScdCardKey, FTScdStaChkExpire,FTScdStaChkDepart,FCScdMaxAlwDisc,FNScdSptSeqNo"
        tSql = tSql & vbCrLf & ", FCScdDiscPer,FCScdDiscAmt,FCScdDiscNew,FCScdAmt"
        tSql = tSql & vbCrLf & ", FCScdDiscPerOrg,FCScdDiscAmtOrg,FCScdDiscNewOrg" '*TW 51-06-05 เพื่อเรียกใช้ให้ได้ตัวเดิมเสมอ
        tSql = tSql & vbCrLf & ",FTDctRefCode,FTScdCstID,FNScdSeqNo"
        tSql = tSql & vbCrLf & ", FTSdtStkCode" '*KT 52-11-25 Deal&Promo
        tSql = tSql & vbCrLf & " FROM " & ptOrgName & " WITH(NOLOCK) "
        tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND FNDctNo >= " & nEN_Disc11MallCrd
        tSql = tSql & vbCrLf & " AND FNDctNo NOT IN (" & nEN_Disc21MPrice & "," & nEN_Disc22MDisc & ")" '*KT 52-09-24 Member Price + Consignment
        tSql = tSql & vbCrLf & " ORDER BY FNScdOrder,FNScdActSeqNo "
        If Not SP_DATbRsOpen(orsRecal, adUseClient, tSql, ocnAll, adOpenStatic, adLockOptimistic, -1) Then
            MsgBox "Error On Query", vbCritical 'tms_CN005249
            Exit Sub
        End If
        Set orsRecal.ActiveConnection = Nothing
    End If
    If Not (orsRecal.BOF And orsRecal.EOF) Then
        'Call C_DATxSumDT '*TW 51-08-28
        'Call C_UPDxLabel '*TW 51-08-28
        Do While Not orsRecal.EOF
            If nOldActSeqNo <> SP_FEDvChkNumeric(orsRecal, "FNScdActSeqNo") Then
                nOldActSeqNo = SP_FEDvChkNumeric(orsRecal, "FNScdActSeqNo")
                nC_RecalActSeqNo = nOldActSeqNo
                If pnSeqNo > 0 Then
                     bC_RecalByItem = True
                     nC_ScdSeqNo = SP_FEDvChkNumeric(orsRecal, "FNScdSeqNo")
                Else
                    bC_RecalByItem = False
                    nC_ScdSeqNo = SP_FEDvChkNumeric(orsRecal, "FNScdSeqNo")
                End If
                bC_PromoAftDeal = IIf(SP_FEDtChkString(orsRecal, "FTSdtStkCode") = "1", True, False)  '*KT 52-11-25 Deal&Promo
                Select Case SP_FEDvChkNumeric(orsRecal, "FNDctNo")
                    Case nEN_Disc01Deal
                    
                    Case nEN_Disc02PmtTime
                        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด Promo Time จะมีการตรวจสอบกลับไปตาราง Mater ของ Promo Time ว่ามี SKU นั้นร่วมหรือไม่
                        'แต่การระบุส่วนลด Promo Time  ของ Return แบบ ไม่ Refer นั้นทำ Manual ไม่ตรวจสอบ Master จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
'*KT 66-03-24 TTK-ISS3000-821 recal discount
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK) "
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            cScdDiscAmtByOne = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsRecal, "FCScdAmt") / SP_FEDvChkNumeric(orsRecal, "FCScdQty"))
'                                            cDisAmt = SP_FMTcCurPoint(cScdDiscAmtByOne * cDisQty)
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType, , cScdDiscAmtByOne) '*KT 66-03-24 TTK-ISS3000-821 recal discount for qty change '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                             Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
'==========================================
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Select Case nC_ENLnDisType
                                        Case nEN_LnDis1Per
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc02PmtTime, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                        Case Else
                                            cC_DiscAmt = cC_DiscValue
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc02PmtTime, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    End Select
                                Else 'แบบ Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Select Case nC_ENLnDisType
                                        Case nEN_LnDis1Per
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc02PmtTime, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                        Case Else
                                            cC_DiscAmt = cC_DiscValue
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc02PmtTime, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    End Select
                                End If
                            Case Else
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                Select Case nC_ENLnDisType
                                    Case nEN_LnDis1Per
                                        nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                        If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc02PmtTime, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    Case Else
                                        cC_DiscAmt = cC_DiscValue
                                        nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                        If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc02PmtTime, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                End Select
                         End Select
                    Case nEN_Disc03Pmt
                        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด Promo  จะมีการตรวจสอบกลับไปตาราง Mater ของ Promo  ว่ามี SKU นั้นร่วมหรือไม่
                        'แต่การระบุส่วนลด Promo   ของ Return แบบ ไม่ Refer นั้นทำ Manual ไม่ตรวจสอบ Master จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
'*KT 66-03-24 TTK-ISS3000-821 recal discount
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            cScdDiscAmtByOne = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsRecal, "FCScdAmt") / SP_FEDvChkNumeric(orsRecal, "FCScdQty"))
'                                            cDisAmt = SP_FMTcCurPoint(cScdDiscAmtByOne * cDisQty)
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType, , cScdDiscAmtByOne) '*KT 66-03-24 TTK-ISS3000-821 recal discount for qty change '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                             Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
'==============================================
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Select Case nC_ENLnDisType
                                        Case nEN_LnDis1Per
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc03Pmt, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                        Case Else
                                            cC_DiscAmt = cC_DiscValue
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc03Pmt, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    End Select
                                Else  'แบบ Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Select Case nC_ENLnDisType
                                        Case nEN_LnDis1Per
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc03Pmt, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                        Case Else
                                            cC_DiscAmt = cC_DiscValue
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc03Pmt, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    End Select
                                End If
                            Case Else
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                Select Case nC_ENLnDisType
                                    Case nEN_LnDis1Per
                                        nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                        If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc03Pmt, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    Case Else
                                        cC_DiscAmt = cC_DiscValue
                                        nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                        If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc03Pmt, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                End Select
                         End Select
                    Case nEN_Disc04CoopDisc
                         '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด CoopDisc  จะมีเป็นแบบ % off อย่างเดียว แต่การ Return แบบ ไม่ Refer  เป็นการใส่ Amount off
                        'จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType) '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                            Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
                                    '----------------------------------------------------------------------------------------------------------------'*Ao 2013/05/09        ไม่อ้างบิล   ทำเหมือนขาย
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                    If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc04CoopDisc, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                Else 'แบบ Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                    If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc04CoopDisc, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                End If
                            Case Else
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc04CoopDisc, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                        End Select
                    
                    Case nEN_Disc05LnDisc
                        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด LnDisc  จะมีการตรวจสอบกลับไปตาราง Mater ของ SKU  ว่ามี SKU นั้นร่วมหรือไม่
                        'แต่การระบุส่วนลด LnDisc   ของ Return แบบ ไม่ Refer นั้นทำ Manual ไม่ตรวจสอบ Master จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType) '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                            Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
                                    '----------------------------------------------------------------------------------------------------------------'*Ao 2013/05/09    ไม่อ้างบิล   ทำเหมือนขาย
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Select Case nC_ENLnDisType
                                        Case nEN_LnDis1Per
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc05LnDisc, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                        Case Else
                                            cC_DiscAmt = cC_DiscValue
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc05LnDisc, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    End Select
                                Else 'แบบ  Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Select Case nC_ENLnDisType
                                        Case nEN_LnDis1Per
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc05LnDisc, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                        Case Else
                                            cC_DiscAmt = cC_DiscValue
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc05LnDisc, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    End Select
                                End If
                            Case Else
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                Select Case nC_ENLnDisType
                                    Case nEN_LnDis1Per
                                        nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                        If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc05LnDisc, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    Case Else
                                        cC_DiscAmt = cC_DiscValue
                                        nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                        If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc05LnDisc, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                End Select
                        End Select
                    Case nEN_Disc06LnMDown
                        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด LnMDown  จะมีการตรวจสอบกลับไปตาราง Mater ของ SKU  ว่ามี SKU นั้นร่วมหรือไม่
                        'แต่การระบุส่วนลด LnMDown   ของ Return แบบ ไม่ Refer นั้นทำ Manual ไม่ตรวจสอบ Master จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType) '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                            Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
                                    '----------------------------------------------------------------------------------------------------------------'*Ao 2013/05/09    ไม่อ้างบิล   ทำเหมือนขาย
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Select Case nC_ENLnDisType
                                        Case nEN_LnDis1Per
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc06LnMDown, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                        Case Else
                                            cC_DiscAmt = cC_DiscValue
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc06LnMDown, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    End Select
                                    '----------------------------------------------------------------------------------------------------------------'*Ao 2013/05/09
                                Else 'แบบ  Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Select Case nC_ENLnDisType
                                        Case nEN_LnDis1Per
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc06LnMDown, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                        Case Else
                                            cC_DiscAmt = cC_DiscValue
                                            nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                            If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc06LnMDown, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    End Select
                                End If
                            Case Else
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                Select Case nC_ENLnDisType
                                    Case nEN_LnDis1Per
                                        nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                        If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc06LnMDown, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                    Case Else
                                        cC_DiscAmt = cC_DiscValue
                                        nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                                        If nC_DTSeqNo > 0 Then Call DIScDiscount(nEN_Disc06LnMDown, IIf(pnSeqNo > 0, eEN_VerifyDiscType.nEN_Type1ByItem, eEN_VerifyDiscType.nEN_Type3None)) '*KT 52-09-23 Analyse Code
                                End Select
                        End Select
                    Case nEN_Disc09CoopCpn
                        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด CoopCpn  จะมีเป็นแบบ % off อย่างเดียว แต่การ Return แบบ ไม่ Refer  เป็นการใส่ Amount off
                        'จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
 '*KT 66-03-24 TTK-ISS3000-821 recal discount
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            cScdDiscAmtByOne = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsRecal, "FCScdAmt") / SP_FEDvChkNumeric(orsRecal, "FCScdQty"))
'                                            cDisAmt = SP_FMTcCurPoint(cScdDiscAmtByOne * cDisQty)
'                                            'cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType, , cScdDiscAmtByOne) '*KT 66-03-24 TTK-ISS3000-821 recal discount for qty change '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                             Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
'=====================================
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                    tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                    Call DIScDiscount(nEN_Disc09CoopCpn, eEN_VerifyDiscType.nEN_Type2ByDcn)


                                Else 'แบบ  Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                    tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                    Call DIScDiscount(nEN_Disc09CoopCpn, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                End If
                            Case Else
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                Call DIScDiscount(nEN_Disc09CoopCpn, eEN_VerifyDiscType.nEN_Type2ByDcn)
                        End Select
                        
                    Case nEN_Disc10CoopBrand
                         '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด CoopBrand  จะมีเป็นแบบ % off อย่างเดียว แต่การ Return แบบ ไม่ Refer  เป็นการใส่ Amount off
                        'จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
'*KT 66-03-24 TTK-ISS3000-821 recal discount
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            'cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            cScdDiscAmtByOne = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsRecal, "FCScdAmt") / SP_FEDvChkNumeric(orsRecal, "FCScdQty"))
'                                            cDisAmt = SP_FMTcCurPoint(cScdDiscAmtByOne * cDisQty)
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType, , cScdDiscAmtByOne) '*KT 66-03-24 TTK-ISS3000-821 recal discount for qty change '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                             Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
'==========================================
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                    tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                    Call DIScDiscount(nEN_Disc10CoopBrand, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                Else 'แบบ  Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                    tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                    Call DIScDiscount(nEN_Disc10CoopBrand, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                End If
                            Case Else
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                Call DIScDiscount(nEN_Disc10CoopBrand, eEN_VerifyDiscType.nEN_Type2ByDcn)
                        End Select
                        
                    Case nEN_Disc08MallCpn
                         '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด MallCpn  จะมีเป็นแบบ % off อย่างเดียว แต่การ Return แบบ ไม่ Refer  เป็นการใส่ Amount off
                        'จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
'*KT 66-03-24 TTK-ISS3000-821 recal discount
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            'cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            cScdDiscAmtByOne = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsRecal, "FCScdAmt") / SP_FEDvChkNumeric(orsRecal, "FCScdQty"))
'                                            cDisAmt = SP_FMTcCurPoint(cScdDiscAmtByOne * cDisQty)
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType, , cScdDiscAmtByOne) '*KT 66-03-24 TTK-ISS3000-821 recal discount for qty change '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                             Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
'====================================
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    bMallCpnInStoreBar = False '*KT 54-07-28 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon
                                    Call DIScDiscount(nEN_Disc08MallCpn, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                Else 'แบบ  Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    bMallCpnInStoreBar = False '*KT 54-07-28 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon
                                    Call DIScDiscount(nEN_Disc08MallCpn, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                End If
                            Case Else
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                bMallCpnInStoreBar = False '*KT 54-07-28 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon
                                Call DIScDiscount(nEN_Disc08MallCpn, eEN_VerifyDiscType.nEN_Type2ByDcn)
                        End Select
                    
                    Case nEN_Disc11MallCrd
                        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด MallCrd  จะมีเป็นแบบ % off อย่างเดียว แต่การ Return แบบ ไม่ Refer  เป็นการใส่ Amount off
                        'จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
 '*KT 66-03-24 TTK-ISS3000-821 recal discount
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                           ' cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            cScdDiscAmtByOne = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsRecal, "FCScdAmt") / SP_FEDvChkNumeric(orsRecal, "FCScdQty"))
'                                            cDisAmt = SP_FMTcCurPoint(cScdDiscAmtByOne * cDisQty)
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType, , cScdDiscAmtByOne)  '*KT 66-03-24 TTK-ISS3000-821 recal discount for qty change'*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                            Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'
'                                        End If
'                                    End If
'=====================================
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                    tCardType = SP_FEDtChkString(orsRecal, "FTScdCardType")
                                    tCardKey = SP_FEDtChkString(orsRecal, "FTScdCardKey")
                                    dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                    tStaChkExpire = SP_FEDtChkString(orsRecal, "FTScdStaChkExpire")
                                    tStaChkDepart = SP_FEDtChkString(orsRecal, "FTScdStaChkDepart")
                                    cAlwDis = SP_FEDvChkNumeric(orsRecal, "FCScdMaxAlwDisc")
                                     tSql = "STP_PRCxFindMallCardByID @ptAccNum= '" & tC_CardID & "'"
                                    Call SP_SQLvExecute(tSql, orsMall, oDB.ocnOnLine, , , , , True, True) 'read master
                                    cAmtMallMax = SP_FEDvChkNumeric(orsMall, "FCMcdAmtMax")
                                    cAmtMallMin = SP_FEDvChkNumeric(orsMall, "FCMcdAmtMin")
                                    tMallRegAlw = SP_FEDtChkString(orsMall, "FTMcdReg")
                                    tMallProAlw = SP_FEDtChkString(orsMall, "FTMcdPromo")
                                    dMallStartDate = SP_FEDdChkDate(orsMall, "FDMcdStartDate")
                                    dMallEndDate = SP_FEDdChkDate(orsMall, "FDMcdEndDate")
                                    tMallFrmAmt = SP_FEDtChkString(orsMall, "FTMcdStaAmtDisc")
                                    tDiscMallAlw = SP_FEDtChkString(orsMall, "FTMcdStaRang")
                                    cC_SpecialPrice = pcSpecial
                                    Call DIScDiscount(nEN_Disc11MallCrd, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                    cC_SpecialPrice = 0
                                Else 'แบบ Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                    tCardType = SP_FEDtChkString(orsRecal, "FTScdCardType")
                                    tCardKey = SP_FEDtChkString(orsRecal, "FTScdCardKey")
                                    dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                    tStaChkExpire = SP_FEDtChkString(orsRecal, "FTScdStaChkExpire")
                                    tStaChkDepart = SP_FEDtChkString(orsRecal, "FTScdStaChkDepart")
                                    cAlwDis = SP_FEDvChkNumeric(orsRecal, "FCScdMaxAlwDisc")
                                     tSql = "STP_PRCxFindMallCardByID @ptAccNum= '" & tC_CardID & "'"
                                    Call SP_SQLvExecute(tSql, orsMall, oDB.ocnOnLine, , , , , True, True) 'read master
                                    cAmtMallMax = SP_FEDvChkNumeric(orsMall, "FCMcdAmtMax")
                                    cAmtMallMin = SP_FEDvChkNumeric(orsMall, "FCMcdAmtMin")
                                    tMallRegAlw = SP_FEDtChkString(orsMall, "FTMcdReg")
                                    tMallProAlw = SP_FEDtChkString(orsMall, "FTMcdPromo")
                                    dMallStartDate = SP_FEDdChkDate(orsMall, "FDMcdStartDate")
                                    dMallEndDate = SP_FEDdChkDate(orsMall, "FDMcdEndDate")
                                    tMallFrmAmt = SP_FEDtChkString(orsMall, "FTMcdStaAmtDisc")
                                    tDiscMallAlw = SP_FEDtChkString(orsMall, "FTMcdStaRang")
                                    cC_SpecialPrice = pcSpecial
                                    Call DIScDiscount(nEN_Disc11MallCrd, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                    cC_SpecialPrice = 0
                                End If
                            Case Else
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                tCardType = SP_FEDtChkString(orsRecal, "FTScdCardType")
                                tCardKey = SP_FEDtChkString(orsRecal, "FTScdCardKey")
                                dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                tStaChkExpire = SP_FEDtChkString(orsRecal, "FTScdStaChkExpire")
                                tStaChkDepart = SP_FEDtChkString(orsRecal, "FTScdStaChkDepart")
                                cAlwDis = SP_FEDvChkNumeric(orsRecal, "FCScdMaxAlwDisc")
                                 tSql = "STP_PRCxFindMallCardByID @ptAccNum= '" & tC_CardID & "'"
                                Call SP_SQLvExecute(tSql, orsMall, oDB.ocnOnLine, , , , , True, True) 'read master
                                cAmtMallMax = SP_FEDvChkNumeric(orsMall, "FCMcdAmtMax")
                                cAmtMallMin = SP_FEDvChkNumeric(orsMall, "FCMcdAmtMin")
                                tMallRegAlw = SP_FEDtChkString(orsMall, "FTMcdReg")
                                tMallProAlw = SP_FEDtChkString(orsMall, "FTMcdPromo")
                                dMallStartDate = SP_FEDdChkDate(orsMall, "FDMcdStartDate")
                                dMallEndDate = SP_FEDdChkDate(orsMall, "FDMcdEndDate")
                                tMallFrmAmt = SP_FEDtChkString(orsMall, "FTMcdStaAmtDisc")
                                tDiscMallAlw = SP_FEDtChkString(orsMall, "FTMcdStaRang")
                                cC_SpecialPrice = pcSpecial
                                Call DIScDiscount(nEN_Disc11MallCrd, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                cC_SpecialPrice = 0
                        End Select
                    Case nEN_Disc12EmpCrd
                        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด EmpCrd  จะมีเป็นแบบ % off อย่างเดียว แต่การ Return แบบ ไม่ Refer  เป็นการใส่ Amount off
                        'จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
  '*KT 66-03-24 TTK-ISS3000-821 recal discount
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            'cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            cScdDiscAmtByOne = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsRecal, "FCScdAmt") / SP_FEDvChkNumeric(orsRecal, "FCScdQty"))
'                                            cDisAmt = SP_FMTcCurPoint(cScdDiscAmtByOne * cDisQty)
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            tVB_DiscEmpChkLimit = W_GETtEmpChkLimit(tScdCardID)
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType, , cScdDiscAmtByOne)  '*KT 66-03-24 TTK-ISS3000-821 recal discount for qty change '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                             Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
'=====================================
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                    tCardType = SP_FEDtChkString(orsRecal, "FTScdCardType")
                                    tCardKey = SP_FEDtChkString(orsRecal, "FTScdCardKey")
                                    dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                    tStaChkExpire = SP_FEDtChkString(orsRecal, "FTScdStaChkExpire")
                                    tStaChkDepart = SP_FEDtChkString(orsRecal, "FTScdStaChkDepart")
                                    cAlwDis = SP_FEDvChkNumeric(orsRecal, "FCScdMaxAlwDisc")
                                    tC_RefCode = SP_FEDtChkString(orsRecal, "FTDctRefCode")
                                    tC_CusID = SP_FEDtChkString(orsRecal, "FTScdCstID")
                                    cC_SpecialPrice = pcSpecial
                                    Call DIScDiscount(nEN_Disc12EmpCrd, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                    cC_SpecialPrice = 0
                                Else 'แบบ Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                    tCardType = SP_FEDtChkString(orsRecal, "FTScdCardType")
                                    tCardKey = SP_FEDtChkString(orsRecal, "FTScdCardKey")
                                    dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                    tStaChkExpire = SP_FEDtChkString(orsRecal, "FTScdStaChkExpire")
                                    tStaChkDepart = SP_FEDtChkString(orsRecal, "FTScdStaChkDepart")
                                    cAlwDis = SP_FEDvChkNumeric(orsRecal, "FCScdMaxAlwDisc")
                                    tC_RefCode = SP_FEDtChkString(orsRecal, "FTDctRefCode")
                                    tC_CusID = SP_FEDtChkString(orsRecal, "FTScdCstID")
                                    cC_SpecialPrice = pcSpecial
                                    Call DIScDiscount(nEN_Disc12EmpCrd, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                    cC_SpecialPrice = 0
                                End If
                            Case Else 'แบบ No Refer
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                tCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                tCardType = SP_FEDtChkString(orsRecal, "FTScdCardType")
                                tCardKey = SP_FEDtChkString(orsRecal, "FTScdCardKey")
                                dCardExpDate = SP_FEDdChkDate(orsRecal, "FDScdExpDate")
                                tStaChkExpire = SP_FEDtChkString(orsRecal, "FTScdStaChkExpire")
                                tStaChkDepart = SP_FEDtChkString(orsRecal, "FTScdStaChkDepart")
                                cAlwDis = SP_FEDvChkNumeric(orsRecal, "FCScdMaxAlwDisc")
                                tC_RefCode = SP_FEDtChkString(orsRecal, "FTDctRefCode")
                                tC_CusID = SP_FEDtChkString(orsRecal, "FTScdCstID")
                                cC_SpecialPrice = pcSpecial
                                Call DIScDiscount(nEN_Disc12EmpCrd, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                cC_SpecialPrice = 0
                        End Select
                    Case nEN_Disc13TranDisc
                        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด TranDisc  จะมีการตรวจสอบกลับไปตาราง Mater ของ SKU  ว่ามี SKU นั้นร่วมหรือไม่
                        'แต่การระบุส่วนลด TranDisc   ของ Return แบบ ไม่ Refer นั้นทำ Manual ไม่ตรวจสอบ Master จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
   '*KT 66-03-24 TTK-ISS3000-821 recal discount
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            'cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            cScdDiscAmtByOne = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsRecal, "FCScdAmt") / SP_FEDvChkNumeric(orsRecal, "FCScdQty"))
'                                            cDisAmt = SP_FMTcCurPoint(cScdDiscAmtByOne * cDisQty)
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType, , cScdDiscAmtByOne) '*KT 66-03-24 TTK-ISS3000-821 recal discount for qty change '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                             Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
'===========================
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Call DIScDiscount(nEN_Disc13TranDisc, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                Else 'แบบ Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Call DIScDiscount(nEN_Disc13TranDisc, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                End If
                            Case Else 'แบบ No Refer
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                Call DIScDiscount(nEN_Disc13TranDisc, eEN_VerifyDiscType.nEN_Type2ByDcn)
                        End Select
                    Case nEN_Disc14TranMDown, nEN_Disc23TranMDown1, nEN_Disc24TranMDown2, nEN_Disc25TranMDown3, nEN_Disc26TranMDown4, nEN_Disc27TranMDown5
                        '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
                        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด TranMDown  จะมีการตรวจสอบกลับไปตาราง Mater ของ SKU  ว่ามี SKU นั้นร่วมหรือไม่
                        'แต่การระบุส่วนลด TranMDown   ของ Return แบบ ไม่ Refer นั้นทำ Manual ไม่ตรวจสอบ Master จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
 '*KT 66-03-24 TTK-ISS3000-821 recal discount
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            'cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            cScdDiscAmtByOne = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsRecal, "FCScdAmt") / SP_FEDvChkNumeric(orsRecal, "FCScdQty"))
'                                            cDisAmt = SP_FMTcCurPoint(cScdDiscAmtByOne * cDisQty)
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType, , cScdDiscAmtByOne)  '*KT 66-03-24 TTK-ISS3000-821 recal discount for qty change '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                             Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
'==========================
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    
                                    Call DIScDiscount(SP_FEDvChkNumeric(orsRecal, "FNDctNo"), eEN_VerifyDiscType.nEN_Type2ByDcn)
                                Else 'แบบ Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    
                                    Call DIScDiscount(SP_FEDvChkNumeric(orsRecal, "FNDctNo"), eEN_VerifyDiscType.nEN_Type2ByDcn)
                                End If
                            Case Else 'แบบ No Refer
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                Call DIScDiscount(SP_FEDvChkNumeric(orsRecal, "FNDctNo"), eEN_VerifyDiscType.nEN_Type2ByDcn)
                        End Select
                    
                    Case nEN_Disc07AutoTran
                        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
                        'กรณีที่มีการ Return แบบ ไม่ Refer นั้น เมื่อมีการ Cancel Item แล้วมีการ Recal ถ้าเป็นส่วนลด AutoTran  จะมีการตรวจสอบกลับไปตาราง Mater ของ AutoTran  ว่ามี SKU นั้นร่วมหรือไม่
                        'แต่การระบุส่วนลด AutoTran   ของ Return แบบ ไม่ Refer นั้นทำ Manual ไม่ตรวจสอบ Master จำเป็นต้อง insert ค่าโดยไม่ใช้ Function  DIScDiscount ของ Class cPSDisc
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                If Not bVB_Reference Then 'แบบ No Refer
 '*KT 66-03-24 TTK-ISS3000-821 recal discount
'                                    If Not orsDT Is Nothing Then
'                                        If orsDT.State = adStateOpen Then orsDT.Close
'                                    End If
'                                    tSql = "SELECT FNSdtSeqNo,FCSdtSaleAmt,FCSdtQty"
'                                    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
'                                    tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
'                                    tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'                                    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
'                                    tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
'                                    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
'                                    If SP_SQLvExecute(tSql, orsDT) = 0 Then
'                                        If Not (orsDT.BOF And orsDT.EOF) Then
'                                            cDisQty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
'                                            nSeqNo = SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo")
'                                            cSdtSaleAmt = SP_FEDvChkNumeric(orsDT, "FCSdtSaleAmt")
'                                            cSdtQtyAllDisAvi = 0
'                                            nDiscType = SP_FEDvChkNumeric(orsRecal, "FNDctNo")
'                                            tEvtID = "" 'SP_FEDtChkString(orsRecal, "FTScdEventID")
'                                            tScdCardID = SP_FEDtChkString(orsRecal, "FTScdCardID")
'                                            'cDisAmt = SP_FEDvChkNumeric(orsRecal, "FCScdAmt")
'                                            cScdDiscAmtByOne = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsRecal, "FCScdAmt") / SP_FEDvChkNumeric(orsRecal, "FCScdQty"))
'                                            cDisAmt = SP_FMTcCurPoint(cScdDiscAmtByOne * cDisQty)
'                                            tCrdType = SP_FEDtChkString(orsRecal, "FTScdCardType")
'                                            tVB_DiscEmpChkLimit = W_GETtEmpChkLimit(tScdCardID)
'                                            Call W_DATxInsertCD(cDisAmt, nSeqNo, tEvtID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCrdType, nDiscType, , cScdDiscAmtByOne) '*KT 66-03-24 TTK-ISS3000-821 recal discount for qty change '*KT 53-02-11 เหมือนกันกับ wCNReturnDisc/W_DATxInsertCD หากมีการเปลี่ยนแปลงควรเปลี่ยนแปลงคู่กัน
'                                             Call DISxClearAll
'                                            tTerminalNo = tVB_CNTerminalNum
'                                            dTranDate = tUT_SaleDate
'                                            tTransNo = tVB_Running
'                                            nDTSeqNo = nSeqNo
'                                            Call C_DATxSumDT(eEN_VerifyDiscType.nEN_Type1ByItem)
'                                        End If
'                                    End If
'=============================
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Call DIScDiscount(nEN_Disc07AutoTran, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                Else 'แบบ Refer
                                    Call DISxClearAll
                                    cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                    Call DIScDiscount(nEN_Disc07AutoTran, eEN_VerifyDiscType.nEN_Type2ByDcn)
                                End If
                            Case Else 'แบบ No Refer
                                Call DISxClearAll
                                cCurrentDiscAmt = C_SETcDiscTypeRecal(orsRecal)
                                Call DIScDiscount(nEN_Disc07AutoTran, eEN_VerifyDiscType.nEN_Type2ByDcn)
                        End Select

                    Case nEN_Disc16FreeGift
                        Call DISxClearAll
                        Call DIScDiscount(nEN_Disc16FreeGift, eEN_VerifyDiscType.nEN_Type3None)
    
                    Case nEN_Disc15Ontop
                    
                    Case nEN_Disc21MPrice, nEN_Disc22MDisc
                        nDTSeqNo = SP_FEDvChkNumeric(orsRecal, "FNSdtSeqNo")
                         If Not orsDT Is Nothing Then
                             If orsDT.State = adStateOpen Then orsDT.Close
                        End If
                        tSql = "SELECT FTSKUCode,FCSdtQty,FTSdtEventMemPrice,FCSdtRegPrice,FCSdtMemberPrice"
                        tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " WITH(NOLOCK)  "
                        tSql = tSql & vbCrLf & " WHERE FTShdTransNo = '" & tVB_Running & "'"
                        tSql = tSql & vbCrLf & " AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
                        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
                        tSql = tSql & vbCrLf & " AND FTShdTransType = '" & tVB_Transaction & "'"
                        tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nC_DTSeqNo
                        If SP_SQLvExecute(tSql, orsDT) = 0 Then
                            If Not (orsDT.BOF And orsDT.EOF) Then
                                 xVB_CNMemberP.ID = SP_FEDtChkString(orsRecal, "FTScdCardID")
                                 xVB_CNMemberP.MallCardType = SP_FEDtChkString(orsRecal, "FTScdCardType")
                                 tVB_StaMemPriceAlw = "Y"
                                 tVB_StaMemPricePromoAlw = False
                                 nVB_CNMemPriceMaxQty = 99999999
                                 tVB_SKUConsignFlag = IIf(SP_FEDvChkNumeric(orsRecal, "FNDctNo") = nEN_Disc22MDisc, "Y", "N")
                                 tVB_CNMemPriceType = SP_FEDtChkString(orsRecal, "FTScdDiscType")
                                 bExit = False
                                 Select Case tVB_CNMemPriceType
                                     Case "1" '% Off
                                          cVB_CNMemPriceAmt = SP_FEDvChkNumeric(orsRecal, "FCScdDiscPer")
                                     Case "2" 'Amount Off
                                         cVB_CNMemPriceAmt = SP_FEDvChkNumeric(orsRecal, "FCScdDiscAmt")
                                     Case "3" 'New Price
                                         cVB_CNMemPriceAmt = SP_FEDvChkNumeric(orsRecal, "FCScdDiscAmt")
                                     Case "4" 'No Promotion
                                         bExit = True
                                     Case Else
                                         bExit = True
                                 End Select
                                 If bExit = False Then
                                    cUT_Qty = SP_FEDvChkNumeric(orsDT, "FCSdtQty")
                                    tVB_CNEventMemPrice = SP_FEDvChkNumeric(orsDT, "FTSdtEventMemPrice")
                                    If cUT_Qty > 1 Then
                                       cVB_CNMemPriceAmtDis = (SP_FEDvChkNumeric(orsDT, "FCSdtRegPrice") - SP_FEDvChkNumeric(orsDT, "FCSdtMemberPrice")) * cUT_Qty
                                    Else
                                       cVB_CNMemPriceAmtDis = (SP_FEDvChkNumeric(orsDT, "FCSdtRegPrice") - SP_FEDvChkNumeric(orsDT, "FCSdtMemberPrice"))
                                    End If
                                    cUT_RegPrice = SP_FEDvChkNumeric(orsDT, "FCSdtRegPrice")
                                    wFunctionKB.W_PRCxMemberPrice nC_DTSeqNo, SP_FEDtChkString(orsDT, "FTSKUCode"), 0, nC_ScdSeqNo
                                    cUT_Qty = 1         '*Ao 2011/03/31  CommSheet ML-PRV-054-SF Wrong qty on last item.xls
                                End If
                            End If
                        End If
                End Select
            End If
            orsRecal.MoveNext
        Loop
        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
        If Val(Format(tVB_TransactionSub)) = tEN_TrnReturn Or Val(Format(tVB_TransactionSub)) = tEN_TrnReturnTaxInc Or Val(Format(tVB_TransactionSub)) = tEN_TrnReturnManualTax Then
            If Not bVB_Reference Then 'แบบ Refer
                Call C_UPDxLabel
            End If
        End If
    End If
    cUT_Qty = 1         '*Ao 2011/03/31  CommSheet ML-PRV-054-SF Wrong qty on last item.xls
    bC_StaRecal = False
    bC_RecalByItem = False
    Call SP_DATxRsNothing(orsDT)
    Call SP_DATxRsNothing(orsMall)
    Call SP_DATxRsNothing(orsRecal)
    Exit Sub
Err_Handler:
    cUT_Qty = 1         '*Ao 2011/03/31  CommSheet ML-PRV-054-SF Wrong qty on last item.xls
    bC_StaRecal = False
    bC_RecalByItem = False
    Call SP_DATxRsNothing(orsDT)
    Call SP_DATxRsNothing(orsMall)
    Call SP_DATxRsNothing(orsRecal)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function W_GETtEmpChkLimit(ByVal ptEmpID As String) As String
    Dim orsTemp As ADODB.Recordset
    Dim tSql$
On Error GoTo ErrHandle
    tSql = "SELECT  FTEcdStaCreditLimit "
    tSql = tSql & vbCrLf & " FROM   TCNMEmpCardMtn WITH(NOLOCK)  "
    tSql = tSql & vbCrLf & " WHERE '" & ptEmpID & "' BETWEEN FTEcdFrom AND FTEcdTo"
    tSql = tSql & vbCrLf & "  AND  LEN(" & ptEmpID & ") <= LEN(FTEcdTo)"
    tSql = tSql & vbCrLf & " AND FTEcdStaRang = 'Y'"
    tSql = tSql & vbCrLf & " ORDER BY LEN(FTEcdFrom)"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
                W_GETtEmpChkLimit = SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")
        Else
                W_GETtEmpChkLimit = "N"
        End If
    Else
        W_GETtEmpChkLimit = "N"
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_GETtEmpChkLimit = "N"
     Call SP_DATxRsNothing(orsTemp)
End Function

Private Function W_GETnGroup(ByVal pnDiscType As Integer) As Integer
    Select Case pnDiscType
        Case nEN_Disc01Deal
            W_GETnGroup = 1
        Case nEN_Disc02PmtTime, nEN_Disc03Pmt, nEN_Disc04CoopDisc
            W_GETnGroup = 2
        Case nEN_Disc05LnDisc, nEN_Disc06LnMDown
            W_GETnGroup = 3
        Case nEN_Disc09CoopCpn
            W_GETnGroup = 4
        Case nEN_Disc10CoopBrand
            W_GETnGroup = 5
        Case nEN_Disc07AutoTran
            W_GETnGroup = 6
        Case nEN_Disc08MallCpn
            W_GETnGroup = 7
        Case nEN_Disc11MallCrd
            W_GETnGroup = 8
        Case nEN_Disc12EmpCrd
            W_GETnGroup = 9
        Case nEN_Disc13TranDisc
            W_GETnGroup = 10
        Case nEN_Disc14TranMDown
            W_GETnGroup = 11
        Case nEN_Disc16FreeGift
            W_GETnGroup = 12
        Case nEN_Disc15Ontop
            W_GETnGroup = 13
        Case nEN_Disc23TranMDown1, nEN_Disc24TranMDown2, nEN_Disc25TranMDown3, nEN_Disc26TranMDown4, nEN_Disc27TranMDown5 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
            W_GETnGroup = 11
    End Select
End Function

Private Function W_DISnGetNextAct() As Integer
Dim tWhe As String
Dim nActSeqNo As Integer

    On Error GoTo ErrHandle
    tWhe = " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " And FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " And FDShdTransDate = '" & tUT_SaleDate & "'"
    nActSeqNo = CInt(SP_VALcText2Double(SP_SQLtFunction("MAX", "FNScdActSeqNo", "" & tVB_CNTblCD & "", tWhe)))
    nActSeqNo = nActSeqNo + 1
    W_DISnGetNextAct = nActSeqNo
    Exit Function
ErrHandle:
    nActSeqNo = 1
    W_DISnGetNextAct = nActSeqNo
End Function

Private Sub W_DATxInsertCD(ByVal pcDisAmt#, ByVal pnFNSdtSeqNo%, ByVal ptEventID$, ByVal pcDisQty#, ByVal pcAlwDis# _
, ByVal pcAlwDisQty#, ByVal ptScdCardID$, ByVal ptCardType$, ByVal pnDiscType%, Optional ByVal ptStaChkDepart$ = "", Optional ByVal pcDiscAmtByOne As Double = 0)
'--------------------------------------------------
 'Insert Discount data to TPST  SalCD
 '*KT 53-01-26 Communication Sheet PH3-ML-SB-073 Deal & Free  in Return Transaction.xls
 '--------------------------------------------------
Dim tSql As String
Dim tField As String
Dim nMaxSeqNo As Integer
Dim tWhe As String
Dim nEffect As Long
Dim TWheMax As String
Dim nLoop As Integer, nMall As Integer, nTop As Integer
Dim cDisPmt As Double
Dim nScdOrder As Integer, nActSeqNo As Integer

On Error GoTo ErrHandle
   
    nScdOrder = W_GETnGroup(pnDiscType)
    nActSeqNo = W_DISnGetNextAct()
    tField = "FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FNScdSeqNo"
    tField = tField & ",FTSkuCode,FTShdTransType,FDShdTransDate,FTSdtBarCode,FTSdtStkCode"
    tField = tField & ",FCSdtStkFac,FTSkuDepCode,FTSkuBrdCode,FTPunCode,FCScdQtyB4"
    tField = tField & ",FCScdQty,FCScdQtyAF,FCScdAmtB4,FCScdAmtB4Avi,FCScdAmt"
    tField = tField & ",FCScdAmtAF,FCScdAmtAFAvi,FCScdDiscPer,FCScdDiscAmt,FCScdDiscNew"
    tField = tField & ",FTScdDiscType,FNDctNo,FTDctName,FTDctRefCode,FTScdEventID"
    tField = tField & ",FTScdCardID,FDScdExpDate,FTScdCardType,FTScdStaTrans"
    tField = tField & ",FTScdCstID,FTScdCardKey,FTScdStaChkExpire,FTScdStaChkDepart,FNScdSptSeqNo,FCScdMaxAlwDisc"
    tField = tField & ",FNScdActSeqNo"
    tField = tField & ",FNScdOrder,FCScdAmtTime,FCScdAmtPmt,FNScdDealGroup"
    tField = tField & ",FCScdDiscPerOrg,FCScdDiscAmtOrg,FCScdDiscNewOrg"
    
    tWhe = " WHERE FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " And  FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " And FDShdTransDate = '" & tUT_SaleDate & "'"
    TWheMax = tWhe
    tWhe = tWhe & " And FNSdtSeqNo = " & pnFNSdtSeqNo
    
    nMaxSeqNo = SP_VALcText2Double(SP_SQLtFunction("MAX", "FNScdSeqNo", "" & tVB_CNTblCD & "", TWheMax))
    nMaxSeqNo = nMaxSeqNo + 1
                
    tSql = "INSERT INTO " & tVB_CNTblCD & " WITH(ROWLOCK) (" & tField & SP_GETtWhoInsertField
    tSql = tSql & vbCrLf & " SELECT "
    tSql = tSql & vbCrLf & "FTStmCode,FTTmnNum, FTShdTransNo, FNSdtSeqNo"
    tSql = tSql & vbCrLf & "," & nMaxSeqNo & " AS FNScdSeqNo "
    tSql = tSql & vbCrLf & ",FTSkuCode,FTShdTransType,FDShdTransDate,FTSdtBarCode,FTSdtStkCode"
    tSql = tSql & vbCrLf & ",FCSdtStkFac,FTSkuDepCode,FTSkuBrdCode,FTPunCode"
    tSql = tSql & vbCrLf & "," & "FCSdtQtyAllDisAvi" & " AS FCScdQtyB4 "
    tSql = tSql & vbCrLf & "," & "FCSdtQty" & " AS FCScdQty"
    tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtQtyAllDisAvi - " & pcAlwDisQty & " < 0 THEN 0 ELSE FCSdtQtyAllDisAvi - " & pcAlwDisQty & " END AS FCScdQtyAF"
    tSql = tSql & vbCrLf & "," & "FCSdtSaleAmt" & " AS FCScdAmtB4"
    tSql = tSql & vbCrLf & "," & "FCSdtAmtDisAvi" & " AS FCScdAmtB4Avi"
    tSql = tSql & vbCrLf & "," & pcDisAmt & " AS FCScdAmt"
    tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtSaleAmt - " & pcDisAmt & " < 0 THEN 0 ELSE FCSdtSaleAmt - " & pcDisAmt & " END AS FCScdAmtAF"
    tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtAmtDisAvi -" & pcAlwDis & " <0 THEN 0 ELSE FCSdtAmtDisAvi -" & pcAlwDis & " END AS FCScdAmtAFAvi"
    tSql = tSql & vbCrLf & "," & 0 & " AS FCScdDiscPer"
    tSql = tSql & vbCrLf & "," & pcDiscAmtByOne & " AS FCScdDiscAmt"
    tSql = tSql & vbCrLf & "," & 0 & " AS FCScdDiscNew"
    tSql = tSql & vbCrLf & "," & "2" & " AS FTScdDiscType"
    tSql = tSql & vbCrLf & "," & pnDiscType & " AS FNDctNo"
    tSql = tSql & vbCrLf & ",'" & SP_SQLtFunction("", "FTDctDiscType", "TCNMDiscountType", "WHERE FNDctNo = " & pnDiscType, oDB.ocnOnLine, True, True) & "' AS FTDctName" 'read master
    tSql = tSql & vbCrLf & ",'' AS FTDctRefCode"
    tSql = tSql & vbCrLf & ",'" & ptEventID & "' AS FTScdEventID"
    tSql = tSql & vbCrLf & ",'" & ptScdCardID & "' AS FTScdCardID"
    tSql = tSql & vbCrLf & ",'" & tUT_SaleDate & "' AS FDScdExpDate"
    tSql = tSql & vbCrLf & ",'" & ptCardType & "' AS FTScdCardType"

    
'*KT 51-09-19 Communication Sheet ML-SF-558 Print  Remain Discount for Employee Card.xls   เก็บค่าไว้ใน FTScdStaTrans ว่า Emp Card ใดไม่ต้องแสดง Remain Discount
'    Select Case nW_DiscType
'        Case nEN_Disc01Deal, nEN_Disc02PmtTime, nEN_Disc03Pmt, nEN_Disc05LnDisc, nEN_Disc06LnMDown, nEN_Disc04CoopDisc
'            tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
'        Case Else
'            tSql = tSql & vbCrLf & ",'1'  AS FTScdStaTrans"
'    End Select

    Select Case pnDiscType
        Case nEN_Disc12EmpCrd
            If UCase(tVB_DiscEmpChkLimit) = "Y" Then
                tSql = tSql & vbCrLf & ",'1'  AS FTScdStaTrans"
            Else
                tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
            End If
        Case Else
            tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
    End Select
'------------------------------------------------------------------------------------------------------------------------------------

    tSql = tSql & vbCrLf & ",'' AS FTScdCstID"
    tSql = tSql & vbCrLf & ",'" & ptCardType & "' AS FTScdCardKey"
    tSql = tSql & vbCrLf & ",'' AS FTScdStaChkExpire"
    tSql = tSql & vbCrLf & ",'" & ptStaChkDepart & "' AS FTScdStaChkDepart"
    tSql = tSql & vbCrLf & ",'' AS FNScdSptSeqNo"
    tSql = tSql & vbCrLf & ",0 AS FCScdMaxAlwDisc"
    tSql = tSql & vbCrLf & "," & nActSeqNo & " AS FNScdActSeqNo"
    tSql = tSql & vbCrLf & "," & nScdOrder & " AS FNScdOrder"

    If pnDiscType = nEN_Disc02PmtTime Then
        cDisPmt = pcDisAmt
    Else
        cDisPmt = 0
    End If
    tSql = tSql & vbCrLf & "," & cDisPmt & " AS FCScdAmtTime"
    If pnDiscType = nEN_Disc03Pmt Then
        cDisPmt = pcDisAmt
    Else
        cDisPmt = 0
    End If
    tSql = tSql & vbCrLf & "," & cDisPmt & " AS FCScdAmtPmt"
    If pnDiscType = nEN_Disc01Deal Then
        tSql = tSql & vbCrLf & ",'" & nMaxSeqNo & "' AS FNScdDealGroup"
    Else
        tSql = tSql & vbCrLf & ",'' AS FNScdDealGroup"
    End If
    tSql = tSql & vbCrLf & ",0 AS FCScdDiscPerOrg"
    tSql = tSql & vbCrLf & "," & pcDisAmt & " AS FCScdDiscAmtOrg"
    tSql = tSql & vbCrLf & ",0 AS FCScdDiscNewOrg"
    
    tSql = tSql & vbCrLf & Left(SP_GETtWhoInsertValue, Len(SP_GETtWhoInsertValue) - 1)
    tSql = tSql & vbCrLf & " FROM " & tVB_CNTblDT & " WITH(NOLOCK) "
    tSql = tSql & vbCrLf & tWhe
    If SP_SQLbExecute(tSql, , False, nEffect) Then
        If nEffect <= 0 Then
            Call SP_MSGnShowing(tC_Msg003, tCS_CNMsgWarn)
        End If
    Else
        Call SP_MSGnShowing(tC_Msg003, tCS_CNMsgWarn)
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Public Function DISbChkMallCpnMinAlw(ByVal pcNetAmt As Double, ByVal pbInStoreBar As Boolean) As Boolean
'CREATE BY            # Kitti                              LAST MODIFIED BY           # Kitti
'CREATE DATE      # 54-07-28                             LAST MODIFIED DATE     # 54-07-28
'DESCRIPTION       # ตรวจสอบว่า Mininum Amount Req.6 Phase4.0(Member Point) Promotion Coupon
'ARGUMENT LIST # pcNetAmt = Net Amount of ticket

'RETURN VALUE  # True if Allow
'                                         False if Not Allow
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
On Error GoTo Err_Handler
    If pbInStoreBar Then
        tSql = "SELECT FCMcpMinAmt FROM TCNMMallCpn WITH(NOLOCK)  "
        tSql = tSql & vbCrLf & "WHERE FNDctNo = " & nEN_Disc08MallCpn
        If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , True, True) = 0 Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                If pcNetAmt >= SP_FEDvChkNumeric(orsTemp, "FCMcpMinAmt") Then
                    DISbChkMallCpnMinAlw = True
                Else
                    DISbChkMallCpnMinAlw = False
                End If
            Else
                DISbChkMallCpnMinAlw = True
            End If
        Else
            DISbChkMallCpnMinAlw = False
        End If
    Else
         DISbChkMallCpnMinAlw = True
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    DISbChkMallCpnMinAlw = False
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function DISbChkAutoTransAct() As Boolean
'CREATE BY            # Kitti                              LAST MODIFIED BY           # Kitti
'CREATE DATE      # 54-11-23                            LAST MODIFIED DATE     # 54-11-23
'DESCRIPTION       # ตรวจสอบว่ามีการกำหนดส่วนลด Auto Trans ที่ Active หรือไม่
'ARGUMENT LIST #
'RETURN VALUE  # True if Active
'                                         False if Not Active
'VERSION                #V.4.0.5
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
On Error GoTo Err_Handler
    tSql = "SELECT     dbo.TCNMAutoTranID.FTAttTierDiscountID,dbo.TCNMAutoTranDis.FTAtdCode"
    tSql = tSql & vbCrLf & "FROM         dbo.TCNMAutoTranDis WITH(NOLOCK) INNER JOIN"
    tSql = tSql & vbCrLf & "                      dbo.TCNMAutoTranID WITH(NOLOCK) ON dbo.TCNMAutoTranDis.FTAttTierDiscountID = dbo.TCNMAutoTranID.FTAttTierDiscountID"
    tSql = tSql & vbCrLf & "WHERE dbo.TCNMAutoTranID.FTAttlStaAlw = 'Y'"
    tSql = tSql & vbCrLf & "AND '" & Format(SP_GETdChkDateDisc, "YYYY/MM/DD") & "' BETWEEN dbo.TCNMAutoTranID.FDAttStartDate AND dbo.TCNMAutoTranDis.FDAtdEndDate"
    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then 'read master
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            DISbChkAutoTransAct = True
        Else
            DISbChkAutoTransAct = False
        End If
    Else
        DISbChkAutoTransAct = False
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    DISbChkAutoTransAct = False
    Call SP_DATxRsNothing(orsTemp)
    'Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function DISbChkFreeGiftAct() As Boolean
'CREATE BY            # Kitti                              LAST MODIFIED BY           # Kitti
'CREATE DATE      # 54-11-23                            LAST MODIFIED DATE     # 54-11-23
'DESCRIPTION       # ตรวจสอบว่ามีการกำหนดส่วนลด FreeGift ที่ Active หรือไม่
'ARGUMENT LIST #
'RETURN VALUE  # True if Active
'                                         False if Not Active
'VERSION                #V.4.0.5
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
On Error GoTo Err_Handler
    tSql = "SELECT FTFgcCode "
    tSql = tSql & vbCrLf & " FROM TCNMFreeGift WITH(NOLOCK) "
    tSql = tSql & vbCrLf & " WHERE ISNULL(FTFgcStaAlw,'N') = 'Y'"
    If SP_SQLbExecute(tSql, orsTemp, , , , oDB.ocnOnLine, True, True) Then 'read master
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            DISbChkFreeGiftAct = True
        Else
            DISbChkFreeGiftAct = False
        End If
    Else
        DISbChkFreeGiftAct = False
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    DISbChkFreeGiftAct = False
    Call SP_DATxRsNothing(orsTemp)
    'Call SP_MSGxShowError(Err.Number, Err.Description)
End Function
                
'Property
'-------------------------------------------------------------------------------------------------------------------------------------
Property Get cDisEMP() As Double   '*Ao 2013/10/29
    cDisEMP = cC_DisEMP
End Property

Property Let bChkEMP(New_bCTmnNum As Boolean)       '*Ao 2013/10/29
    bC_ChkEMP = New_bCTmnNum
End Property

Property Let tTerminalNo(New_tCTmnNum As String)
    tC_TmnNum = New_tCTmnNum
End Property

Property Let dTranDate(New_dCTranDate As Date)
    dC_TranDate = New_dCTranDate
End Property

Property Let tTransNo(New_tTransNo As String)
    tC_TransNo = New_tTransNo
End Property

Property Let tCardID(New_tCardID As String) 'หมายเลขบัตร
    tC_CardID = New_tCardID
End Property

Property Let tCusID(New_tCusID As String) 'หมายเลขสมาชิก
    tC_CusID = New_tCusID
End Property

Property Let tCardType(New_tCardType As String) 'ประเภทบัตร
    tC_CardType = New_tCardType
End Property

Property Let tCardKey(New_tCardKey As String) 'Key ของ Format บัตร
    tC_CardKey = New_tCardKey
End Property

Property Let cDiscValue(New_cDiscValue As Double) 'ค่าที่เป็นจำนวนส่วนลด
    cC_DiscValue = New_cDiscValue
End Property

Property Let nENLnDisType(New_nENLnDisType As eEN_LineDiscType) 'ประเภทการลด
    nC_ENLnDisType = New_nENLnDisType
End Property

Property Let dCardExpDate(New_dCardExpDate As Date) 'วันที่หมดอายุ
    dC_CardExpDate = New_dCardExpDate
End Property
    
Property Let nDTSeqNo(New_nDTSeqNo As Integer) 'ลำดับรายการสินค้า
    nC_DTSeqNo = New_nDTSeqNo
End Property

Property Get nDTSeqNo() As Integer 'ลำดับรายการสินค้า
    nDTSeqNo = nC_DTSeqNo
End Property
    
Property Let nSptSeqNo(New_nSptSeqNo As Integer) 'ลำดับรายการสินค้าใน DT ที่ถูก Split
    nC_SptSeqNo = New_nSptSeqNo
End Property
    
Property Let tStaChkExpire(New_tStaChkExpire As String) 'สถานะการตรวจสอบวันหมดอายุ
    tC_StaChkExpire = New_tStaChkExpire
End Property
    
Property Let tStaChkDepart(New_tStaChkDepart As String) 'สถานะการตรวจสอบ Depart ของ SKU
    tC_StaChkDepart = New_tStaChkDepart
End Property
    
Property Let cAlwDis(New_cAlwDis As Double) 'มูลค่าที่อนุญาตให้คำนวณส่วนลด
    cC_AlwDis = New_cAlwDis
End Property
    
Property Let cAlwDisQty(New_cAlwDisQty As Double) 'จำนวนที่อนุญาตให้คำนวณส่วนลด
    cC_AlwDisQty = New_cAlwDisQty
End Property
    
Property Let cTotalDisAmt(New_cTotalDisAmt As Double) 'มูลค่าที่ลด
    cC_TotalDisAmt = New_cTotalDisAmt
End Property
    
Property Let cTotalDisQty(New_cTotalDisQty As Double) 'จำนวนที่ลด
    cC_TotalDisQty = New_cTotalDisQty
End Property
    
Property Let cOntopFactor(New_cOntopFactor As Double) 'อัตราส่วนที่คำนวณ Ontop ใหม่
    cC_OntopFactor = New_cOntopFactor
End Property
    
Property Let cBalance(New_cBalance As Double) 'ยอดที่ต้องชำระ
    cC_Balance = New_cBalance
End Property
    
Property Let tRefCode(New_tRefCode As String)
    tC_RefCode = New_tRefCode
End Property
    
Property Let tEventID(New_tEventID As String)
    tC_EventID = New_tEventID
End Property
    
Property Get cLastDiscount() As Double
    cLastDiscount = cC_LastDiscount
End Property

'Property Get cLastDiscountPerUnit() As Double
'    cLastDiscountPerUnit = cC_LastDiscountPerUnit
'End Property

Property Get eLastDiscType() As eEN_DiscType
    eLastDiscType = nC_ENDiscType
End Property

Property Get cDiscValue() As Double
    cDiscValue = cC_DiscValue
End Property

Property Let cDiscOntopAlw(New_DiscOntopAlw As Double)
    cC_DiscOntopAlw = New_DiscOntopAlw
End Property

Property Get cDiscOntopAlw() As Double
    cDiscOntopAlw = cC_DiscOntopAlw
End Property

Property Get cAlwDis() As Double  'มูลค่าที่อนุญาตให้คำนวณส่วนลด
    cAlwDis = cC_AlwDis
End Property

Property Get nENLnDisType() As eEN_LineDiscType
    nENLnDisType = nC_ENLnDisType
End Property

Property Let cEmpMaxDis(New_EmpMaxDis As Double)
    cC_EmpMaxDis = New_EmpMaxDis
End Property

Property Get cAmtAllMall() As Double  '*Ao 51-07-30   ยอดที่เอาไปคิด Mall
    cAmtAllMall = cC_AmtAllMall
End Property

Property Get cAmtMallMin() As Double
    cAmtMallMin = cC_AmtMallMin
End Property

Property Let cAmtMallMin(New_AmtMallMin As Double)
    cC_AmtMallMin = New_AmtMallMin
End Property

Property Get cAmtMallMax() As Double
    cAmtMallMax = cC_AmtMallMax
End Property

Property Let cAmtMallMax(New_AmtMallMax As Double)
    cC_AmtMallMax = New_AmtMallMax
End Property

Property Get tMallRegAlw() As String
    tMallRegAlw = tC_MallRegAlw
End Property

Property Let tMallRegAlw(New_MallRegAlw As String)
    tC_MallRegAlw = New_MallRegAlw
End Property

Property Get tMallProAlw() As String
    tMallProAlw = tC_MallRegAlw
End Property

Property Let tMallProAlw(New_MallProAlw As String)
    tC_MallProAlw = New_MallProAlw
End Property

Property Get dMallStartDate() As Date
    dMallStartDate = dC_MallStartDate
End Property

Property Let dMallStartDate(New_MallStartDate As Date)
    dC_MallStartDate = New_MallStartDate
End Property

Property Get dMallEndDate() As Date
    dMallEndDate = dC_MallEndDate
End Property

Property Let dMallEndDate(New_MallEndDate As Date)
    dC_MallEndDate = New_MallEndDate
End Property

Property Get tMallFrmAmt() As String
    tMallFrmAmt = tC_MallFrmAmt
End Property

Property Let tMallFrmAmt(New_MallFrmAmt As String)
    tC_MallFrmAmt = New_MallFrmAmt
End Property

Property Get tDiscMallAlw() As String
    tDiscMallAlw = tC_DiscMallAlw
End Property

Property Let tDiscMallAlw(New_DiscMallAlw As String)
    tC_DiscMallAlw = New_DiscMallAlw
End Property

Property Let nENDiscType(New_DiscType As eEN_DiscType)
    nC_ENDiscType = New_DiscType
End Property

Property Get nENDiscType() As Long
    nENDiscType = nC_ENDiscType
End Property

Property Let bRecalByItem(New_RecalByItem As Boolean)
    bC_RecalByItem = New_RecalByItem
End Property

Property Let nScdSeqNo(New_ScdSeqNo As Long)
    nC_ScdSeqNo = New_ScdSeqNo
End Property

Property Let nRecalActSeqNo(New_RecalActSeqNo As Long)
    nC_RecalActSeqNo = New_RecalActSeqNo
End Property

Property Let bReConnectMember(New_ReReConnectMember As Boolean)
    bC_ReConnectMember = New_ReReConnectMember
End Property

Property Let cDiscAmt(New_cDiscAmt As Double)
    cC_DiscAmt = New_cDiscAmt
End Property

Property Get cDiscAmt() As Double
    cDiscAmt = cC_DiscAmt
End Property

Property Let cDiscPer(New_cDiscPer As Double)
    cC_DiscPer = New_cDiscPer
End Property

Property Get cDiscPer() As Double
    cDiscPer = cC_DiscPer
End Property

Property Let cDiscNew(New_cDiscNew As Double)
    cC_DiscNew = New_cDiscNew
End Property

Property Get cDiscNew() As Double
    cDiscNew = cC_DiscNew
End Property

Property Let cDiscAmtOrg(New_cDiscAmtOrg As Double)
    cC_DiscAmtOrg = New_cDiscAmtOrg
End Property

Property Let cDiscPerOrg(New_cDiscPerOrg As Double)
    cC_DiscPerOrg = New_cDiscPerOrg
End Property

Property Let cDiscNewOrg(New_cDiscNewOrg As Double)
    cC_DiscNewOrg = New_cDiscNewOrg
End Property

Property Let nActSeqNo(New_nActSeqNo As Double)
    nC_ActSeqNo = New_nActSeqNo
End Property

Property Let bPromoAftDeal(New_bPromoAftDeal As Boolean) '*KT 52-11-25 Deal&Promo
    bC_PromoAftDeal = New_bPromoAftDeal
End Property

Property Let tCardName(New_tCardName As String)
    tC_CardName = New_tCardName
End Property

Property Get tCardName() As String
    tCardName = tC_CardName
End Property

Property Let bIsLastDeal(New_bIsLastDeal As Boolean)  '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls '*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตั้งต้น 3.0.19
    bC_IsLastDeal = New_bIsLastDeal
End Property

Property Get bIsLastDeal() As Boolean  '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls '*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตั้งต้น 3.0.19
    bIsLastDeal = bC_IsLastDeal
End Property

Property Let bIsPreviewDay(New_bIsPreviewDay As Boolean) '*KT 53-10-29 Preview Day
    bC_IsPreviewDay = New_bIsPreviewDay
End Property

Property Get bIsPreviewDay() As Boolean  '*KT 53-10-29 Preview Day
    bIsPreviewDay = bC_IsPreviewDay
End Property

Property Let orsOntop(New_oRsOnTop As ADODB.Recordset) '*KT 53-11-12 ปรับปรุง
    Set oC_RsTempOnTop = New_oRsOnTop.Clone
End Property

Property Let oRsOnTopNothing(New_bNoting As Boolean) '*KT 53-11-12 ปรับปรุง
    If New_bNoting Then
        If Not oC_RsTempOnTop Is Nothing Then
            If oC_RsTempOnTop.State = adStateOpen Then oC_RsTempOnTop.Close
            Set oC_RsTempOnTop = Nothing
        End If
    End If
End Property

'*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
Property Let bPOSNoOnTopFlag(New_bPOSNoOnTopFlag As Boolean) '
    bC_POSNoOnTopFlag = New_bPOSNoOnTopFlag
End Property

Property Let tPOSNoOnTopAlwType(New_tPOSNoOnTopAlwType As String) '
    tC_POSNoOnTopAlwType = New_tPOSNoOnTopAlwType
End Property

Property Let tFloorAlwOnTop(New_tFloorAlwOnTop As String) '
    tC_FloorAlwOnTop = New_tFloorAlwOnTop
End Property

Property Let cSubTotalAll(New_cSubTotalAll As Double)  '
    cC_SubTotalAll = New_cSubTotalAll
End Property

Property Let cSubTotal(New_cSubTotal As Double)  '
    cC_SubTotal = New_cSubTotal
End Property

Property Let tStaOnTopAlwTierAmtLast(New_tStaOnTopAlwTierAmtLast As String) '
    tC_StaOnTopAlwTierAmtLast = New_tStaOnTopAlwTierAmtLast
End Property

Property Let tMallAlwPoint(New_MallAlwPoint As String) '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
    tC_MallAlwPoint = New_MallAlwPoint
End Property

Property Let bMallCpnInStoreBar(New_MallCpnInStoreBar As Boolean) '*KT 54-07-28 V.4.0.1 Ph.4.0 Req.6 (Member Point) Promotion Coupon
    bC_MallCpnInStoreBar = New_MallCpnInStoreBar
End Property

Property Let tEmplimit(New_tEmplimit As String)             '*Ao 2013/03/06
    tC_Emplimit = New_tEmplimit
End Property

Property Let tEmpRemark(New_tEmpRemark As String)             '*Ao 2013/03/06
    tC_EmpRemark = New_tEmpRemark
End Property
