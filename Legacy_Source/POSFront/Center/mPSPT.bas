Attribute VB_Name = "mPSPT"
Option Explicit
Dim oGridDeal As VSFlexGrid
Dim orsWDeal As ADODB.Recordset
'Dim orsWDisDeal As ADODB.Recordset
Dim tW_TBLDealAct As String
Dim tW_TBLDealSum As String
Dim tW_TBLDealDT As String

Private Function PT_TMPbCreate(ByVal ptOrgName$, ByVal ptTmpName$) As Boolean
'
'
'*TW 51-04-08 Create Temp by Terminal
'------------------------------------------------------------------------------------------------------
    Dim tSql As String
    On Error GoTo Err_Handler
    If SP_TBLbChkExist(ptTmpName) Then
        tSql = "DROP TABLE " & ptTmpName
        If Not SP_SQLbExecute(tSql) Then GoTo Err_Handler
    End If
    tSql = "EXEC STP_CNxCloneTable @SourceTable ='" & ptOrgName & "',@DestinationTable ='" & ptTmpName & "'"
    If Not SP_SQLbExecute(tSql) Then GoTo Err_Handler
    
    PT_TMPbCreate = True
    Exit Function
Err_Handler:
    If Err.Number <> 0 Then Call SP_MSGxShowError(Err.Number, Err.Description)
    PT_TMPbCreate = False
End Function


'Const tW_TBLDealAct = "aTPSTDealAct"
'Const tW_TBLDealSum = "TPTDS001"
'Const tW_TBLDealDT = "TPSDT001"

Private Sub PT_CALxSumDis2DT(ByVal ptDealNo, ByVal ptDealType$)
'-----------------------------------------------------------
'   Parameter :
'   Cmt :
'   File Refer :
'   Developer : SOMJAI
'   Date Create : 17/03/2008 :  16:49:17
'-----------------------------------------------------------
    Dim cDis As Double
    Dim nDTSeqNo As Integer
    Dim nSptSeqNo As Integer      'DT New Row
    Dim cAlwDis As Double        'Price for calculate
    Dim cAlwDisQty As Double 'Qty for cal
    Dim cTotalDisAmt  As Double
    Dim cTotalDisQty  As Double
    Dim tCardKey As String
    Dim tCardType As String
    Dim nENLnDisType As eEN_LineDiscType
    Dim nDealGroup As Integer
    Dim tDacSku As String
    Dim tEventID As String
    Dim cDisAmt As Double
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    
    tSql = "SELECT * "
    tSql = tSql & vbCrLf & "FROM " & tW_TBLDealSum
    If SP_SQLbExecute(tSql, orsTemp) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            Do While Not orsTemp.EOF
                nDTSeqNo = SP_FEDvChkNumeric(orsTemp, "FNDacSkuSeq")
                nSptSeqNo = 0      'DT New Row
                tDacSku = SP_FEDtChkString(orsTemp, "FTDacSku")
                If SP_FEDvChkNumeric(orsTemp, "FCDacQtyDiv") >= 1 Then
                    If SP_FEDvChkNumeric(orsTemp, "FNDacQtyGet") >= 1 Then
                        cAlwDisQty = SP_FEDvChkNumeric(orsTemp, "FNDacQtyGet")
                        cAlwDis = SP_FEDvChkNumeric(orsTemp, "FCDacAmtKey")
                    Else
                        cAlwDisQty = SP_FEDvChkNumeric(orsTemp, "FNDacQtyGet")
                        cAlwDis = 0
                    End If
                End If
                tCardKey = ptDealNo
                tCardType = ptDealType
                nDealGroup = SP_FEDvChkNumeric(orsTemp, "FNDacSkuGrp")
                tEventID = SP_FEDtChkString(orsTemp, "FTEvnCode")
                cDisAmt = SP_FEDvChkNumeric(orsTemp, "FCDacDisAmt")
                nENLnDisType = SP_VALcText2Double(SP_FEDtChkString(orsTemp, "FTDacSku"))
                
                Select Case IIf(SP_FEDvChkNumeric(orsTemp, "FNDacQtyBuy") > 0, "1", "0") & IIf(SP_FEDvChkNumeric(orsTemp, "FNDacQtyGet") > 0, "1", "0")
                    Case "00"
                        tCardKey = "0"
                    Case "01"
                        tCardKey = "2"
                    Case "10"
                        tCardKey = "1"
                    Case "11"
                        tCardKey = "3"
                End Select
                
                With oDisc
                    .cDiscValue = cDisAmt 'ยอดส่วนลดที่เกิดจาก Deal
                    .nENLnDisType = nENLnDisType 'ประเภทการลด Per, Amt, New, Free
                    .nDTSeqNo = nDTSeqNo 'SeqNo ของ DT ที่ลด
                    .nSptSeqNo = nSptSeqNo 'SeqNo ของ DT ที่แยก
                    .dTranDate = tUT_SaleDate 'Sale Date
                    .tTerminalNo = tVB_CNTerminalNum 'Terminal Number
                    .tTransNo = tVB_Running 'Tran No
                    .tRefCode = ptDealNo 'Deal No
                    .tEventID = tEventID 'Event ID
                    .tCardType = ptDealType 'Deal Type '001','002'
                    .tCardKey = tCardKey 'Free Flag
                    .cBalance = cAlwDis * cAlwDisQty 'ยอดที่นำไปคำนวณส่วนลด
                    .cAlwDis = cAlwDis 'ยอดที่นำมาคำนวณ
                    .cAlwDisQty = cAlwDisQty 'จำนวนที่นำมาคำนวณ
            '        .tCusID = nDealGroup
                    cDis = .DIScDiscount(nEN_Disc01Deal)
                End With
                orsTemp.MoveNext
            Loop
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Sub

Public Sub PT_CALxSaveResult(ByVal ptDealType As String)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-11-2008 :  15:56:42
'-----------------------------------------------------------
Dim orsTemp As ADODB.Recordset
Dim tSql$, tSqlF$, tSqlV$, tFedName$
Dim tChkSKU$
Dim nRowDat As Long
Dim nCol%

Dim nQtyCalDiv%, nQtyDiv%, nQtyMod%, nQtyPmtBuy%, nQtyPmtGet%, nQtyDivAll%
Dim cAmtDiv As Double, cAmtMod As Double, cDisAvg As Double, cAmtTotal As Double
On Error GoTo ErrHandle:

With oGridDeal
    For nRowDat = .rows - 1 To .FixedRows Step -1
        If .IsSubtotal(nRowDat) Then            'Store data
            nQtyCalDiv = .ValueMatrix(nRowDat, .ColIndex("FCDalQtyCalDiv"))
            nQtyDiv = .ValueMatrix(nRowDat, .ColIndex("FCDacQtyDiv"))
            nQtyMod = .ValueMatrix(nRowDat, .ColIndex("FCDacQtyMod"))
            cAmtDiv = .ValueMatrix(nRowDat, .ColIndex("FCDacAmtDiv"))
            cAmtMod = .ValueMatrix(nRowDat, .ColIndex("FCDacAmtMod"))
            cDisAvg = .ValueMatrix(nRowDat, .ColIndex("FCDacAmtTotal"))
            cAmtTotal = .ValueMatrix(nRowDat, .ColIndex("FCDacDisAvg"))

            'nQtyPmtBuy = .ValueMatrix(nRowDat, .ColIndex("FNDacQtyBuy"))
            'nQtyPmtGet = .ValueMatrix(nRowDat, .ColIndex("FNDacQtyGet"))
            .RemoveItem (nRowDat)
        Else
            'set value to grid
            If tChkSKU = .TextMatrix(nRowDat, .ColIndex("FTDacSKU")) Then
                Select Case ptDealType
                    Case "009", "010"   'remove dup sku row
                        tChkSKU = ""
                        .RemoveItem (nRowDat)
                    Case Else
                End Select
            Else
'                If .ValueMatrix(nRowDat, .ColIndex("FCDacQtyGet")) > 0 Then
'                    Select Case ptDealType
'                        Case "001", "003", "007", "009", "010" 'PRO RATE ALL
'                        Case "002", "004"      'PRORATE LAST ITEM
'                        Case "005", "006", "008"       'PRORATE GET GROUP
'                    End Select
                tChkSKU = .TextMatrix(nRowDat, .ColIndex("FTDacSKU"))
                .TextMatrix(nRowDat, .ColIndex("FCDalQtyCalDiv")) = nQtyCalDiv
                .TextMatrix(nRowDat, .ColIndex("FCDacQtyDiv")) = nQtyDiv
                .TextMatrix(nRowDat, .ColIndex("FCDacQtyMod")) = nQtyMod
                .TextMatrix(nRowDat, .ColIndex("FCDacAmtDiv")) = cAmtDiv
                .TextMatrix(nRowDat, .ColIndex("FCDacAmtMod")) = cAmtMod
                .TextMatrix(nRowDat, .ColIndex("FCDacAmtTotal")) = cDisAvg
                .TextMatrix(nRowDat, .ColIndex("FCDacDisAvg")) = cAmtTotal
'                End If
            End If
        End If
    Next nRowDat

'    'delete item in temp calculation specific  relate pdt key(check by all pmh code)
'    tSql = "DELETE  "
'    tSql = tSql & vbCr & " FROM " & tW_TBLDealSum      'ลบทั้ง"
'    Call SP_SQLvExecute(tSql)       '*Tee 49-09-20
'        'generate sql script for save record after calculate promoton
        For nRowDat = .FixedRows To .rows - 1
            tSqlF = ""
            tSqlV = ""
            For nCol = 0 To .Cols - 1
                tFedName = .ColKey(nCol)
                tSqlF = tSqlF & tFedName & ","
                Select Case Mid(tFedName, 2, 1)
                    Case "T"
                        tSqlV = tSqlV & "'" & SP_STRtSiggleCode(.TextMatrix(nRowDat, nCol)) & "',"
                    Case "N"
                        tSqlV = tSqlV & SP_VALnText2Double(.ValueMatrix(nRowDat, nCol)) & ","
                    Case "C"
                        tSqlV = tSqlV & SP_VALcText2Double(.ValueMatrix(nRowDat, nCol)) & ","
                    Case "D"
'                        If tFedName = "FDSpmTimeIns" Then
'                            If .ValueMatrix(nRowDat, .ColIndex("FCSpmAmtDiscZ")) > 0 Then
'                                tSqlV = tSqlV & "'" & SP_DTEtStrSTD(.TextMatrix(nRowDat, nCol)) & " " & Format(.TextMatrix(nRowDat, nCol), "HH:MM:SS") & "',"
'                            Else
'                                tSqlV = tSqlV & "'" & SP_DTEtStrSTD(.TextMatrix(nRowDat, nCol)) & " 23:59:59',"
'                            End If
'                        Else
                            tSqlV = tSqlV & "'" & SP_DTEtStrSTD(.TextMatrix(nRowDat, nCol)) & " " & Format(.TextMatrix(nRowDat, nCol), "HH:MM:SS") & "',"
'                        End If
                End Select
            Next nCol
            tSqlF = Left(tSqlF, Len(tSqlF) - 1)
            tSqlV = Left(tSqlV, Len(tSqlV) - 1)

            'บันทึกการคำนวนทุกรายการกลุ่มโปรโมชั่น
            tSql = "INSERT INTO " & tW_TBLDealSum & "(" & tSqlF & ")"        '*Tee 49-11-09
            tSql = tSql & " VALUES(" & tSqlV & ")"
            Call SP_SQLvExecute(tSql, , ocnAll)
        Next nRowDat
    End With
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub
'
'Private Sub PT_RECxDealFillData(ByVal pcDisCount As Double, ByVal pnQtyDiv%, ByVal pcSumAmtDis As Double)
''-----------------------------------------------------------
''   Call :
''   Cmt :
''   Developer : SOMJAI
''   Date Create : 03-07-2008 :  16:08:24
''-----------------------------------------------------------
'    On Error GoTo ErrHandle
'    Dim nRecCount%, nPos%, nQtyKey%
'    Dim tSKU$
'    Dim orsSKU As New ADODB.Recordset
'    Dim cDisCal As Double, cAmtKey As Double, cSumAvg As Double
'
'    orsSKU.Fields.Append "FTDacSKU", adVarChar, 20
'     orsSKU.Open
'
'    'Update record set before save into ATPSTDealTemp
'    For nRecCount = 1 To orsWDeal.RecordCount - 1
'        tSKU = SP_FEDvChkNumeric(orsWDeal, "FTDacSku")
'        orsSKU.Find "FTDacSku ='" & tSKU & "'", , adSearchForward
'
'        ',  adSeekFirstEQ  '  .Find(tSKU, 1, adSearchForward, 1)
'        If orsSKU.EOF Then
'            orsSKU.AddNew
'            orsSKU.Fields("FTDacSKU") = tSKU
'            orsSKU.Update
'
'            If SP_FEDtChkString(orsWDeal, "FNDacQtyGet") = "0" Then
'                cDisCal = 0
'             Else
'                cAmtKey = SP_FEDvChkNumeric(orsWDeal, "FCDacAmtKey")
'                cDisCal = Format((pcDisCount * cAmtKey) / pcSumAmtDis, SP_FMTtNum2Set(True, True))
'                cSumAvg = cSumAvg + cDisCal
'            End If
'
'            orsWDeal.Fields("FCDacQtyDiv") = pnQtyDiv      'CalQty
'            orsWDeal.Fields("FCDacDisAvg") = cDisCal 'DisAvg/Item
'        End If
'        orsWDeal.MoveNext
'    Next nRecCount
'    orsWDeal.MoveLast
'    orsWDeal.Fields("FCDacDisAvg") = pcDisCount - cSumAvg
'    orsWDeal.Fields("FCDacQtyDiv") = pnQtyDiv
'    'เอา ATPSTDealAct  Join DT เพื่อ เอา Seq จริง มาเฉลี่ย
'    'loop ทั้ง หมด
'    Set orsSKU = Nothing
'    Exit Sub
'ErrHandle:
'    Call SP_MSGShowError
'End Sub
''
''Private Sub PT_OBJxAddDiscount(ByVal poRsFrm As ADODB.Recordset, ByRef poRsTo As ADODB.Recordset)
'''-----------------------------------------------------------
'''   Call :
'''   Cmt :
'''   Developer : SOMJAI
'''   Date Create : 03-06-2008 :  15:19:52
'''-----------------------------------------------------------
''    On Error GoTo ErrHandle
''    Dim tSKU$
''    Dim nRec%
''
''
''
''    Do While Not poRsFrm.EOF
''        poRsTo.AddNew
''        poRsTo("dTranDate") = tUT_SaleDate      'DocDate
''        poRsTo("tTerminalNo") = tVB_CNTerminalNum        'PosNo
''        poRsTo("tTransNo") = tVB_Running        'DocNo
''        poRsTo("nDTSeqNo") = SP_FEDvChkNumeric(poRsFrm, "FNDacSkuSeq")    'DTSeq
''        poRsTo("nSptSeqNo") = SP_FEDvChkNumeric(poRsFrm, "FNDacSkuSeq")   'DTSeqNew
''        poRsTo("cAlwDis") = SP_FEDvChkNumeric(poRsFrm, "FCDacSalePri")      'CalPrice
''        poRsTo("cAlwDisQty") = SP_FEDvChkNumeric(poRsFrm, "FNDacQtyKey")       'CalQty
''        poRsTo("cTotalDisAmt") = SP_FEDvChkNumeric(poRsFrm, "FCDacDisAvg") 'DisAvg/Item
''
''            nGetQty = SP_FEDvChkNumeric(poRsFrm, "FNDacQtyGet")
''            nPmtQty = SP_FEDvChkNumeric(poRsFrm, "FNDacQtyBuy")
''            nDivQty = SP_FEDvChkNumeric(poRsFrm, "FCDacQtyDiv")
''            nDivQtyAll = nGetQty * nPmtQty * nDivQty
'''
'''            If nDivQtyAll > SP_FEDvChkNumeric(poRsFrm, .ColIndex("FNDacQtyKey")) Then nDivQtyAll = .SP_FEDvChkNumeric(poRsFrm, .ColIndex("FNDacQtyKey"))
''
''        poRsTo("cTotalDisQty") = SP_FEDvChkNumeric(poRsFrm, "FCDacQtyDiv")    'TotalDisQty
''        poRsTo("tCardKey") = SP_FEDtChkString(poRsFrm, "FTPdhDealNo")        'DealNo
''        poRsTo("tCardType") = SP_FEDtChkString(poRsFrm, "FTDacDealType")         'DealType
''        poRsTo.Update
''        poRsFrm.MoveNext
''    Loop
''
''    Exit Sub
''ErrHandle:
''    Call SP_MSGShowError
''End Sub

Private Function PT_OBJoDiscount() As ADODB.Recordset
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-06-2008 :  14:44:52
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim orsTemp As New ADODB.Recordset

    With orsTemp
        .Fields.Append "dTranDate", adDate      'DocDate
        .Fields.Append "tTerminalNo", adVarChar, 20         'PosNo
        .Fields.Append "tTransNo", adVarChar, 20       'DocNo
        .Fields.Append "nDTSeqNo", adInteger    'DTSeq
        .Fields.Append "nSptSeqNo", adInteger   'DTSeqNew
        .Fields.Append "cAlwDis", adDouble      'CalPrice
        .Fields.Append "cAlwDisQty", adDouble       'CalQty
        .Fields.Append "cTotalDisAmt", adDouble 'DisAvg/Item
        .Fields.Append "cTotalDisQty", adDouble     'TotalDisQty
        .Fields.Append "tCardKey", adVarChar, 20       'DealNo
        .Fields.Append "tCardType", adVarChar, 20       'DealType
        .Open
    End With
    Set PT_OBJoDiscount = orsTemp
    Set orsTemp = Nothing
    Exit Function
ErrHandle:
    Call SP_MSGShowError
End Function

Public Sub PT_PRCxDealVerify()
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:33:04
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim orsDealAll As ADODB.Recordset
    Dim orsDealDT As ADODB.Recordset
    Dim tDealType$, tSql$, tTblPmt$, tDealNo$
    
    tW_TBLDealAct = tVB_CNTblDealAct '"TPSTDealAct"
    tW_TBLDealSum = "TPTDS" & tVB_CNTerminalNum
    tW_TBLDealDT = tVB_CNTblDT
    
    'Check exist deal
    'ต้องเพิ่มการสร้าง Temp ของการคำนวณโปรโมชั่นที่นี่ง*TW 51-03-27
    If Not PT_TMPbCreate("TPSTDealSum", tW_TBLDealSum) Then
        Exit Sub
    End If

    
    If Not PT_CHKbIsDealActive Then Exit Sub  '"DT Join Active Deal"
    If Not PT_CHKbIsDealAllow Then Exit Sub
    oDisc.DISbClearDiscDT       'Clear DT ก่อนทำ
    'delete item in temp result calculation
    tSql = "DELETE  "
    tSql = tSql & " FROM " & tW_TBLDealSum       'ลบทั้ง"
    Call SP_SQLvExecute(tSql)       '*Tee 49-09-20
    '---------------------------------------------------------------------------------------------------------------------
    'หาลำดับ The Best
    '---------------------------------------------------------------------------------------------------------------------
    'tSql = "SELECT DISTINCT DEAL.FTDacDealType"
    tSql = "SELECT DISTINCT DEAL.FTPdhDealNo,FTDacDealType"
    tSql = tSql & " FROM " & tW_TBLDealDT & " AS DT INNER JOIN " & tW_TBLDealAct & " AS DEAL "
    tSql = tSql & " ON DT.FTSkuCode = DEAL.FTDacSku"
    tSql = tSql & " WHERE DT.FTShdTransNo='" & tVB_Running & "'"
    Call SP_SQLvExecute(tSql, orsDealAll, , , , , , True) '"record set of deal group bydeal"
    'ทำเพื่อหาลำดับการคำนวน จะได้ส่วนลดตาม the best
    Do While Not orsDealAll.EOF
        tDealType = SP_FEDtChkString(orsDealAll, "FTDacDealType")
        tDealNo = SP_FEDtChkString(orsDealAll, "FTPdhDealNo")
        Call PT_CHKxDeal(tDealNo, tDealType)
        orsDealAll.MoveNext
    Loop
    Call SP_DATxRsClose(orsDealAll)

    '---------------------------------------------------------------------------------------------------------------------
    'คำนวนตามลำดับ The Best
    '---------------------------------------------------------------------------------------------------------------------
    tSql = "SELECT FTPdhDealNo,FTDacDealType, Max(FCDacDisAvg) AS FCDacDisAvg"
    tSql = tSql & " From " & tW_TBLDealSum
    tSql = tSql & " GROUP BY FTPdhDealNo,FTDacDealType"
    tSql = tSql & " ORDER BY  FCDacDisAvg DESC"
    Call SP_SQLvExecute(tSql, orsDealAll)

    'delete item in temp result calculation
    tSql = "DELETE  "
    tSql = tSql & " FROM " & tW_TBLDealSum       'ลบทั้ง"
    Call SP_SQLvExecute(tSql)       '*Tee 49-09-20

    'ส่วนลดที่จะได้รับจริงต่อ Deal
    Do While Not orsDealAll.EOF
        tDealType = SP_FEDtChkString(orsDealAll, "FTDacDealType")
        tDealNo = SP_FEDtChkString(orsDealAll, "FTPdhDealNo")
        Call PT_CHKxDeal(tDealNo, tDealType)
        'Call PT_CALxSumDis2DT(tDealNo, tDealType)       'บันทึกและ update ลง CD and DT ทีละ Deal   ยังไม่ Split รายการ
        orsDealAll.MoveNext
    Loop
    Call PT_CALxSumDis2DT(tDealNo, tDealType)       'บันทึกและ update ลง CD and DT ทีละ Deal   ยังไม่ Split รายการ'*TW 51-04-19
    Call SP_DATxRsNothing(orsDealAll)
    Call SP_DATxRsNothing(orsDealDT)
    '---------------------------------------------------------------------------------------------------------------------


'    Dim tDealNo$
'    Dim nDealGet%, nRow%, nDealBuy%, nQtyDivAll%, nQtyDiv%
'    Dim n
'    Dim cDisAvg As Double
'    Dim cDisSum As Double
'    'selec join dt &  sumDeal
'    tSql = "SELECT DISTINCT DEAL.FNDacSkuGrp,DEAL.FTPdhDealNo, DT.*"
'    tSql = tSql & " FROM aTPSTDealSum as DEAL RIGHT JOIN TPSDT001 as DT ON DEAL.FTDacSku = DT.FTSkuCode"
'    tSql = tSql & " ORDER BY DT.FNSdtSeqNo DESC"
'    Call SP_SQLvExecute(tSql, orsDealDT)       '"record set of dealsum and DT all
'
'    tSql = "SELECT DISTINCT  FTPdhDealNo,FNDacSkuGrp,FNDacQtyBuy,FNDacQtyGet,FCDacQtyDiv,FCDacDisAvg "
'    tSql = tSql & " FROM aTPSTDealSum"
'    Call SP_SQLvExecute(tSql, orsDealAll)        'deal all get
'
'    Do While Not orsDealAll.EOF
'        tDealNo = SP_FEDtChkString(orsDealAll, "FTPdhDealNo")
'        nDealBuy = SP_FEDvChkNumeric(orsDealAll, "FNDacQtyBuy")
'        nDealGet = SP_FEDvChkNumeric(orsDealAll, "FNDacQtyGet")
'        nQtyDiv = SP_FEDvChkNumeric(orsDealAll, "FCDacQtyDiv")
'        cDisAvg = SP_FEDvChkNumeric(orsDealAll, "FCDacDisAvg")
'        nQtyDivAll = nDealBuy * nQtyDiv * nDealGet
'        orsDealDT.Filter = "FTPdhDealNo ='" & tDealNo & "'"
'
'        For nRow = 1 To orsDealDT.RecordCount - 1
'       '     cDisAvg
'            Select Case tDealType
'                Case "001", "003", "007", "009", "010"       'PRORATE ALL
'                          'นำ จำนวนไปตัดออก และ average ตามมูลค่า
'                Case "002", "004"       'PRORATE LAST ITEM
'                          'split จำนวน เท่า กับ Qty ที่ได้รับ ที่เหลือ เป็น รายการแยกออกมา
'                Case "005", "006", "008"       'PRORATE GET GROUP
'                           'เฉลี่ยทั้งหมดของ Group ที่ได้รับ
'                    If nDealGet > 0 Then orsDealDT.Filter = "FNDacQtyGet =" & nDealGet & ""
'            End Select
'            orsDealDT.Update "", ""
'        Next nRow
'        orsDealAll.MoveNext
'    Loop

    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsDealAll)
    Call SP_DATxRsNothing(orsDealDT)
End Sub

'Public Function PT_CALcDeal01_06(ByRef pnQtyDiv As Integer, ByRef pcAmtGrpAvg As Double) As Double
Public Function PT_CALcDeal01_06() As Double
'-----------------------------------------------------------
'   Call :
'   Cmt :
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim nQtyKey%, nQtyDiv%, nQtyMod%, nQtyPmt%, nRow%, nRowGet%, nQtyGet%
    Dim cAmtDiv As Double, cAmtMod As Double, nCalDiv As Double, cPmtDisAmt As Double
    Dim cAmtKey As Double, cTotal As Double, cCalValue As Double, cAmtPmt As Double, cUnitPri As Double
    Dim nMinCalDiv As Double
    Dim tPmtDisType$
    Dim nStartPro%, nStopPro%

    With oGridDeal

         .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FNDacQtyBuy"), , vbYellow, vbBlue
         .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FNDacQtyGet"), , vbYellow, vbBlue
         .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FCDacSalePri"), , vbYellow, vbBlue
         .Subtotal flexSTSum, .ColIndex("FNDacSkuGrp"), .ColIndex("FNDacQtyKey"), , vbYellow, vbBlue

        nStartPro = .FixedRows
        For nRow = .FixedRows To .rows - 1
            If .IsSubtotal(nRow) Then
                 nQtyKey = .ValueMatrix(nRow, .ColIndex("FNDacQtyKey"))
                 nQtyPmt = .ValueMatrix(nRow, .ColIndex("FNDacQtyBuy"))
                 cUnitPri = .ValueMatrix(nRow, .ColIndex("FCDacSalePri"))

                 If nMinCalDiv = nCalDiv And nMinCalDiv = 0 Then
                     nCalDiv = SP_VALcText2Double(nQtyKey \ nQtyPmt)
                     nMinCalDiv = nCalDiv
                 Else
                     nCalDiv = SP_VALcText2Double(nQtyKey \ nQtyPmt)
                     If nCalDiv < nMinCalDiv Then nMinCalDiv = nCalDiv
                 End If

                 .TextMatrix(nRow, .ColIndex("FCDalQtyCalDiv")) = nCalDiv

'                 If .ValueMatrix(nRow, .ColIndex("FNDacQtyGet")) = "1" Then
'                     nStopPro = nRow - 1
'                     pcAmtGrpAvg = pcAmtGrpAvg + .Aggregate(flexSTSum, nStartPro, .ColIndex("FCDacAmtKey"), nStopPro, .ColIndex("FCDacAmtKey"))
'                 Else
'                     nStartPro = nRow + 1
'                 End If
            End If
        Next nRow

        nQtyDiv = nMinCalDiv
'        pnQtyDiv = nQtyDiv
        For nRow = .FixedRows To .rows - 1
            If .IsSubtotal(nRow) Then
                nQtyKey = .ValueMatrix(nRow, .ColIndex("FNDacQtyKey"))
                nQtyPmt = .ValueMatrix(nRow, .ColIndex("FNDacQtyBuy"))
                cUnitPri = .ValueMatrix(nRow, .ColIndex("FCDacSalePri"))
                nQtyGet = .ValueMatrix(nRow, .ColIndex("FNDacQtyGet"))

                cAmtDiv = (nQtyDiv * nQtyPmt * cUnitPri)
                nQtyMod = nQtyKey - (nQtyDiv * nQtyPmt)
                cAmtMod = nQtyMod * cUnitPri

                .TextMatrix(nRow, .ColIndex("FCDacQtyDiv")) = nQtyDiv
                .TextMatrix(nRow, .ColIndex("FCDacAmtDiv")) = cAmtDiv
                .TextMatrix(nRow, .ColIndex("FCDacQtyMod")) = nQtyMod
                .TextMatrix(nRow, .ColIndex("FCDacAmtMod")) = cAmtMod
                If nQtyGet > 0 Then nRowGet = nRow
            End If
        Next nRow

        tPmtDisType = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FTPddDisType"), .rows - 1, .ColIndex("FTPddDisType"))
        cPmtDisAmt = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FCDacDisAmt"), .rows - 1, .ColIndex("FCDacDisAmt"))
        cAmtMod = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FCDacAmtMod"), .rows - 1, .ColIndex("FCDacAmtMod"))
        cAmtKey = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FCDacAmtKey"), .rows - 1, .ColIndex("FCDacAmtKey"))

        If nRow > 0 Then cUnitPri = .ValueMatrix(nRowGet, .ColIndex("FCDacSalePri"))

        Select Case tPmtDisType
            Case "1"    'Free
                cCalValue = (cPmtDisAmt * cUnitPri) / 100
                cTotal = cAmtKey - (nQtyDiv * cCalValue)

            Case "2"    'new price
                cTotal = cAmtKey - (nQtyDiv * cPmtDisAmt)
                
            Case "3"    'Disc Amount
                cTotal = (cPmtDisAmt * nQtyDiv) + cAmtMod
            Case "4"    'Disc Percent
                cTotal = cAmtKey - (nQtyDiv * cUnitPri)
        End Select
        .TextMatrix(.rows - 1, .ColIndex("FCDacAmtTotal")) = cTotal
        .TextMatrix(.rows - 1, .ColIndex("FCDacDisAvg")) = (cAmtKey - cTotal) * nQtyGet

        PT_CALcDeal01_06 = cAmtKey - cTotal

    End With

    Exit Function
ErrHandle:
    PT_CALcDeal01_06 = 0
End Function

'Public Function PT_CALcDeal07(ByRef pnQtyDiv As Integer, ByRef pcAmtGrpAvg As Double) As Double
Public Function PT_CALcDeal07() As Double
'-----------------------------------------------------------
'   Call :
'   Cmt :
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim nQtyKey%, nQtyPmt%, nQtyGet%, nQtyDiv%, nQtyMod%, nRow%, nMinCalDiv%, nCalDiv%
    Dim cAmtPmt As Double, cPmtDisAmt As Double
    Dim cUnitPri As Double, cAmtDiv As Double, cAmtMod As Double
    Dim cAmtKey As Double, cTotal As Double, cCalValue As Double
    Dim tPmtDisType$
    Dim nStartPro%, nStopPro%

    With oGridDeal
         .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FCDacAmtBuy"), , vbYellow, vbBlue
         .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FNDacQtyGet"), , vbYellow, vbBlue
         .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FCDacSalePri"), , vbYellow, vbBlue
         .Subtotal flexSTSum, .ColIndex("FNDacSkuGrp"), .ColIndex("FCDacAmtKey"), , vbYellow, vbBlue

        For nRow = .FixedRows To .rows - 1
            If .IsSubtotal(nRow) Then
                cAmtKey = .ValueMatrix(nRow, .ColIndex("FCDacAmtKey"))
                cAmtPmt = .ValueMatrix(nRow, .ColIndex("FCDacAmtBuy"))
                cUnitPri = .ValueMatrix(nRow, .ColIndex("FCDacSalePri"))

                If cAmtPmt = 0 Then cAmtPmt = 1
                If nMinCalDiv = nCalDiv And nMinCalDiv = 0 Then
                    nCalDiv = SP_VALcText2Double(cAmtKey \ cAmtPmt)
                    nMinCalDiv = nCalDiv
                Else
                    nCalDiv = SP_VALcText2Double(cAmtKey \ cAmtPmt)
                    If nCalDiv < nMinCalDiv Then nMinCalDiv = nCalDiv
                End If

                .TextMatrix(nRow, .ColIndex("FCDalQtyCalDiv")) = nCalDiv
'                 If .ValueMatrix(nRow, .ColIndex("FNDacQtyGet")) = "1" Then
'                     nStopPro = nRow - 1
'                     pcAmtGrpAvg = pcAmtGrpAvg + .Aggregate(flexSTSum, nStartPro, .ColIndex("FCDacAmtKey"), nStopPro, .ColIndex("FCDacAmtKey"))
'                 Else
'                     nStartPro = nRow + 1
'                 End If
            End If
        Next nRow

        nQtyDiv = nMinCalDiv
'        pnQtyDiv = nQtyDiv
        For nRow = .FixedRows To .rows - 1
            If .IsSubtotal(nRow) Then
                cAmtKey = .ValueMatrix(nRow, .ColIndex("FCDacAmtKey"))
                cAmtPmt = .ValueMatrix(nRow, .ColIndex("FCDacAmtBuy"))
                cUnitPri = .ValueMatrix(nRow, .ColIndex("FCDacSalePri"))
                nQtyGet = .ValueMatrix(nRow, .ColIndex("FNDacQtyGet"))

                cAmtDiv = (nQtyDiv * cAmtPmt)
                nQtyMod = 0
                cAmtMod = cAmtKey - cAmtDiv

                .TextMatrix(nRow, .ColIndex("FCDacQtyDiv")) = nQtyDiv
                .TextMatrix(nRow, .ColIndex("FCDacAmtDiv")) = cAmtDiv
                .TextMatrix(nRow, .ColIndex("FCDacQtyMod")) = nQtyMod
                .TextMatrix(nRow, .ColIndex("FCDacAmtMod")) = cAmtMod
            End If
        Next nRow

        tPmtDisType = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FTPddDisType"), .rows - 1, .ColIndex("FTPddDisType"))
        cPmtDisAmt = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FCDacDisAmt"), .rows - 1, .ColIndex("FCDacDisAmt"))
        Select Case tPmtDisType
            Case "1"    'Free
            Case "2"    'new price
            Case "3"    'Disc Amount
                cTotal = cAmtKey - (nQtyDiv * cPmtDisAmt)
            Case "4"    'Disc Percent
                cCalValue = (cPmtDisAmt * cAmtPmt) / 100
                cTotal = cAmtKey - (nQtyDiv * cCalValue)
        End Select
        .TextMatrix(.rows - 1, .ColIndex("FCDacAmtTotal")) = cTotal
        .TextMatrix(.rows - 1, .ColIndex("FCDacDisAvg")) = cAmtKey - cTotal

        PT_CALcDeal07 = cAmtKey - cTotal
    End With
    Exit Function
ErrHandle:
    PT_CALcDeal07 = 0
End Function

'Public Function PT_CALcDeal08(ByRef pnQtyDiv As Integer, ByRef pcAmtGrpAvg As Double) As Double
Public Function PT_CALcDeal08() As Double
'-----------------------------------------------------------
'   Call :
'   Cmt :
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim nQtyKey%, nQtyPmt%, nQtyGet%, nQtyDiv%, nQtyMod%, nRow%, nMinCalDiv%, nCalDiv%
    Dim cAmtPmt As Double, cPmtDisAmt As Double
    Dim cUnitPri As Double, cAmtDiv As Double, cAmtMod As Double
    Dim cAmtKey As Double, cTotal As Double, cCalValue As Double
    Dim tPmtDisType$
    Dim nStartPro%, nStopPro%

    With oGridDeal
        .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FCDacAmtBuy"), , vbYellow, vbBlue
        .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FNDacQtyGet"), , vbYellow, vbBlue
        .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FCDacSalePri"), , vbYellow, vbBlue
        .Subtotal flexSTSum, .ColIndex("FNDacSkuGrp"), .ColIndex("FNDacQtyKey"), , vbYellow, vbBlue
        .Subtotal flexSTSum, .ColIndex("FNDacSkuGrp"), .ColIndex("FCDacAmtKey"), , vbYellow, vbBlue

        For nRow = .FixedRows To .rows - 1
            If .IsSubtotal(nRow) Then
                cAmtKey = .ValueMatrix(nRow, .ColIndex("FCDacAmtKey"))
                cAmtPmt = .ValueMatrix(nRow, .ColIndex("FCDacAmtBuy"))
                cUnitPri = .ValueMatrix(nRow, .ColIndex("FCDacSalePri"))

                If cAmtPmt = 0 Then cAmtPmt = 1
                If nMinCalDiv = nCalDiv And nMinCalDiv = 0 Then
                    nCalDiv = SP_VALcText2Double(cAmtKey \ cAmtPmt)
                    nMinCalDiv = nCalDiv
                Else
                    nCalDiv = SP_VALcText2Double(cAmtKey \ cAmtPmt)
                    If nCalDiv < nMinCalDiv Then nMinCalDiv = nCalDiv
                End If
                .TextMatrix(nRow, .ColIndex("FCDalQtyCalDiv")) = nCalDiv
'                 If .ValueMatrix(nRow, .ColIndex("FNDacQtyGet")) = "1" Then
'                     nStopPro = nRow - 1
'                     pcAmtGrpAvg = pcAmtGrpAvg + .Aggregate(flexSTSum, nStartPro, .ColIndex("FCDacAmtKey"), nStopPro, .ColIndex("FCDacAmtKey"))
'                 Else
'                     nStartPro = nRow + 1
'                 End If

            End If

        Next nRow

        nQtyDiv = nMinCalDiv
'        pnQtyDiv = nQtyDiv
        For nRow = .FixedRows To .rows - 1
            If .IsSubtotal(nRow) Then
                cAmtKey = .ValueMatrix(nRow, .ColIndex("FCDacAmtKey"))
                cAmtPmt = .ValueMatrix(nRow, .ColIndex("FCDacAmtBuy"))
                cUnitPri = .ValueMatrix(nRow, .ColIndex("FCDacSalePri"))
                nQtyGet = .ValueMatrix(nRow, .ColIndex("FNDacQtyGet"))

                If nQtyGet > 0 Then
                    nQtyKey = .ValueMatrix(nRow, .ColIndex("FNDacQtyKey"))
                    nMinCalDiv = SP_VALcText2Double(nQtyKey / nQtyGet)
                    If nMinCalDiv < nQtyDiv Then nQtyDiv = nMinCalDiv
                    cAmtDiv = (nQtyDiv * nQtyGet * cUnitPri)
                    nQtyMod = nQtyKey - (nQtyDiv * nQtyGet)
                    cAmtMod = nQtyMod * cUnitPri
                Else
                    cAmtDiv = (nQtyDiv * cAmtPmt)
                    nQtyMod = 0
                    cAmtMod = cAmtKey - cAmtDiv
                End If

                .TextMatrix(nRow, .ColIndex("FCDacQtyDiv")) = nQtyDiv
                .TextMatrix(nRow, .ColIndex("FCDacAmtDiv")) = cAmtDiv
                .TextMatrix(nRow, .ColIndex("FCDacQtyMod")) = nQtyMod
                .TextMatrix(nRow, .ColIndex("FCDacAmtMod")) = cAmtMod
            End If
        Next nRow

        tPmtDisType = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FTPddDisType"), .rows - 1, .ColIndex("FTPddDisType"))
        cPmtDisAmt = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FCDacDisAmt"), .rows - 1, .ColIndex("FCDacDisAmt"))
            Select Case tPmtDisType
                Case "1"    'Free
                Case "2"    'new price
                Case "3"    'Disc Amount
                cTotal = cAmtKey - (nQtyGet * nQtyDiv * cPmtDisAmt)
                Case "4"    'Disc Percent
                cCalValue = (cPmtDisAmt * cAmtDiv) / 100
                cTotal = cAmtKey - (nQtyGet * nQtyDiv * cCalValue)
            End Select
            .TextMatrix(.rows - 1, .ColIndex("FCDacAmtTotal")) = cTotal
            .TextMatrix(.rows - 1, .ColIndex("FCDacDisAvg")) = cAmtKey - cTotal

        PT_CALcDeal08 = cAmtKey - cTotal
    End With
    Exit Function
ErrHandle:
    PT_CALcDeal08 = 0
End Function

'Public Function PT_CALcDeal09(ByRef pnQtyDiv As Integer, ByRef pcAmtGrpAvg As Double) As Double
Public Function PT_CALcDeal09() As Double
'-----------------------------------------------------------
'   Call :
'   Cmt :
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim nQtyKey%, nQtyDiv%, nQtyMod%, nQtyBuy%, nRow%, nRowGet%
    Dim cAmtDiv As Double, cAmtMod As Double, nCalDiv As Double, cPmtDisAmt As Double
    Dim cAmtKey As Double, cTotal As Double, cCalValue As Double, cAmtPmt As Double
    Dim nMinCalDiv As Double
    Dim tPmtDisType$
    Dim nStartPro%, nStopPro%

    With oGridDeal
           nRow = .rows - 1
           Do While nRow > .FixedRows 'remove row
                nQtyBuy = .ValueMatrix(nRow, .ColIndex("FNDacQtyBuy"))
                nQtyKey = .ValueMatrix(nRow, .ColIndex("FNDacQtyKey"))
                If nQtyBuy > nQtyKey Then .RemoveItem (nRow)
                nRow = nRow - 1
           Loop

         .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FCDacAmtKey"), , vbYellow, vbBlue
         .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FNDacQtyKey"), , vbYellow, vbBlue

        For nRow = .FixedRows To .rows - 1
            If .IsSubtotal(nRow) Then
                cAmtKey = .ValueMatrix(nRow, .ColIndex("FCDacAmtKey"))
                nQtyKey = .ValueMatrix(nRow, .ColIndex("FNDacQtyKey"))
                tPmtDisType = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FTPddDisType"), .rows - 1, .ColIndex("FTPddDisType"))
                cPmtDisAmt = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FCDacDisAmt"), .rows - 1, .ColIndex("FCDacDisAmt"))

                nQtyDiv = nQtyKey
'                pnQtyDiv = nQtyKey
                cAmtDiv = cAmtKey
                nQtyMod = 0
                cAmtMod = 0
                Select Case tPmtDisType
                    Case "1"    'Free
                    Case "2"    'new price
                    Case "3"    'Disc Amount
                    Case "4"    'Disc Percent
                        cCalValue = (cPmtDisAmt * cAmtKey) / 100
                        cTotal = cAmtKey - (cCalValue)
                End Select

                .TextMatrix(nRow, .ColIndex("FCDalQtyCalDiv")) = nQtyDiv
                .TextMatrix(nRow, .ColIndex("FCDacQtyDiv")) = nQtyDiv
                .TextMatrix(nRow, .ColIndex("FCDacAmtDiv")) = cAmtDiv
                .TextMatrix(nRow, .ColIndex("FCDacQtyMod")) = nQtyMod
                .TextMatrix(nRow, .ColIndex("FCDacAmtMod")) = cAmtMod
                .TextMatrix(nRow, .ColIndex("FCDacAmtTotal")) = cTotal
                .TextMatrix(nRow, .ColIndex("FCDacDisAvg")) = cAmtKey - cTotal
                'If .ValueMatrix(nRow - 1, .ColIndex("FNDacQtyGet")) = "1" Then
                 '   nStopPro = nRow - 1
'                    pcAmtGrpAvg = pcAmtGrpAvg + cAmtKey ' .Aggregate(flexSTSum, nStartPro, .ColIndex("FCDacAmtKey"), nStopPro, .ColIndex("FCDacAmtKey"))
                'Else
                 '   nStartPro = nRow + 1
                'End If
            End If
        Next nRow

        PT_CALcDeal09 = cAmtKey - cTotal
    End With
    Exit Function
ErrHandle:
    PT_CALcDeal09 = 0
End Function

'Public Function PT_CALcDeal10(ByRef pnQtyDiv As Integer, ByRef pcAmtGrpAvg As Double) As Double
Public Function PT_CALcDeal10() As Double
'-----------------------------------------------------------
'   Call :
'   Cmt :
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim nQtyKey%, nQtyPmt%, nQtyGet%, nQtyDiv%, nQtyMod%, nRow%, nRowGet%
    Dim cUnitPri As Double, cAmtDiv As Double, cAmtMod As Double, cPmtDisAmt As Double
    Dim cAmtKey As Double, cTotal As Double, cCalValue As Double, cAmtPmt As Double
    Dim nMinCalDiv As Double, nCalDiv As Double
    Dim tPmtDisType$
    Dim nStartPro%, nStopPro%

    With oGridDeal
           nRow = .rows - 1
           Do While nRow > .FixedRows 'remove row
                cAmtPmt = .ValueMatrix(nRow, .ColIndex("FCDacAmtBuy"))
                cAmtKey = .ValueMatrix(nRow, .ColIndex("FCDacAmtKey"))
                If cAmtKey < cAmtPmt Then .RemoveItem (nRow)
                nRow = nRow - 1
           Loop

         .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FCDacAmtKey"), , vbYellow, vbBlue
         .Subtotal flexSTMax, .ColIndex("FNDacSkuGrp"), .ColIndex("FNDacQtyKey"), , vbYellow, vbBlue

        For nRow = .FixedRows To .rows - 1
            If .IsSubtotal(nRow) Then
                cAmtKey = .ValueMatrix(nRow, .ColIndex("FCDacAmtKey"))
                nQtyKey = .ValueMatrix(nRow, .ColIndex("FNDacQtyKey"))
                cAmtPmt = .ValueMatrix(nRow, .ColIndex("FCDacAmtBuy"))
                tPmtDisType = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FTPddDisType"), .rows - 1, .ColIndex("FTPddDisType"))
                cPmtDisAmt = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FCDacDisAmt"), .rows - 1, .ColIndex("FCDacDisAmt"))

                nQtyDiv = nQtyKey
'                pnQtyDiv = nQtyKey
                cAmtDiv = cAmtKey
                nQtyMod = 0
                cAmtMod = 0

                Select Case tPmtDisType
                    Case "1"    'Free
                    Case "2"    'new price
                    Case "3"    'Disc Amount
                    Case "4"    'Disc Percent
                        cCalValue = (cPmtDisAmt * cAmtKey) / 100
                        cTotal = cAmtKey - (cCalValue)
                End Select
                .TextMatrix(nRow, .ColIndex("FCDalQtyCalDiv")) = nQtyDiv
                .TextMatrix(nRow, .ColIndex("FCDacQtyDiv")) = nQtyDiv
                .TextMatrix(nRow, .ColIndex("FCDacAmtDiv")) = cAmtDiv
                .TextMatrix(nRow, .ColIndex("FCDacQtyMod")) = nQtyMod
                .TextMatrix(nRow, .ColIndex("FCDacAmtMod")) = cAmtMod
                .TextMatrix(nRow, .ColIndex("FCDacAmtTotal")) = cTotal
                .TextMatrix(nRow, .ColIndex("FCDacDisAvg")) = cAmtKey - cTotal
'                 If .ValueMatrix(nRow, .ColIndex("FNDacQtyGet")) = "1" Then
'                     nStopPro = nRow - 1
'                     pcAmtGrpAvg = pcAmtGrpAvg + .Aggregate(flexSTSum, nStartPro, .ColIndex("FCDacAmtKey"), nStopPro, .ColIndex("FCDacAmtKey"))
'                 Else
'                     nStartPro = nRow + 1
'                 End If
            End If
        Next nRow

        PT_CALcDeal10 = cAmtKey - cTotal
    End With
    Exit Function
ErrHandle:
    PT_CALcDeal10 = 0
End Function

Private Sub PT_CHKxDeal01(ByVal ptDealType$)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim cTotalAmtGrp As Double, cDisPmtAll As Double
    Dim nQtyDiv%

    'รอว่า จะต้องการ Return มาใน รูปแบบใด

    Set oGridDeal.DataSource = orsWDeal
    Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
    'wPSDealCal.Show 1
    'cDisPmtAll = PT_CALcDeal01_06(nQtyDiv, cTotalAmtGrp)
    'Call PT_CALxSaveResult(ptDealType)

    'Call PT_RECxDealFillData(cDisPmtAll, nQtyDiv, cTotalAmtGrp)

'    orsWDeal.Save "C:\Deal" & Format(Time, "HHmmss") & ".xml", adPersistXML
'    Call PT_OBJxAddDiscount(orsWDeal, orsWDisDeal)
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub PT_CHKxDeal02(ByVal ptDealType$)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim cTotalAmtGrp As Double, cDisPmtAll As Double
    Dim nQtyDiv%

    'รอว่า จะต้องการ Return มาใน รูปแบบใด

    Set oGridDeal.DataSource = orsWDeal
    Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
    'wPSDealCal.Show 1
    'cDisPmtAll = PT_CALcDeal01_06(nQtyDiv, cTotalAmtGrp)
    'Call PT_CALxSaveResult(ptDealType)

'    Call PT_RECxDealFillData(cDisPmtAll, nQtyDiv, cTotalAmtGrp)
'    Call PT_OBJxAddDiscount(orsWDeal, orsWDisDeal)
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub PT_CHKxDeal03(ByVal ptDealType$)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim cTotalAmtGrp As Double, cDisPmtAll As Double
    Dim nQtyDiv%

    'รอว่า จะต้องการ Return มาใน รูปแบบใด

    Set oGridDeal.DataSource = orsWDeal
    Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
    'wPSDealCal.Show 1
    'cDisPmtAll = PT_CALcDeal01_06(nQtyDiv, cTotalAmtGrp)
    'Call PT_CALxSaveResult(ptDealType)

'    Call PT_RECxDealFillData(cDisPmtAll, nQtyDiv, cTotalAmtGrp)
'    Call PT_OBJxAddDiscount(orsWDeal, orsWDisDeal)
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub PT_CHKxDeal04(ByVal ptDealType$)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim cTotalAmtGrp As Double, cDisPmtAll As Double
    Dim nQtyDiv%

    'รอว่า จะต้องการ Return มาใน รูปแบบใด

    Set oGridDeal.DataSource = orsWDeal
    Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
    'wPSDealCal.Show 1
    'cDisPmtAll = PT_CALcDeal01_06(nQtyDiv, cTotalAmtGrp)
    'Call PT_CALxSaveResult(ptDealType)

'    Call PT_RECxDealFillData(cDisPmtAll, nQtyDiv, cTotalAmtGrp)
'    Call PT_OBJxAddDiscount(orsWDeal, orsWDisDeal)
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub PT_CHKxDeal05(ByVal ptDealType$)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim cTotalAmtGrp As Double, cDisPmtAll As Double
    Dim nQtyDiv%

    'รอว่า จะต้องการ Return มาใน รูปแบบใด

    Set oGridDeal.DataSource = orsWDeal
    Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
    'wPSDealCal.Show 1
    'cDisPmtAll = PT_CALcDeal01_06(nQtyDiv, cTotalAmtGrp)
    'Call PT_CALxSaveResult(ptDealType)

'    Call PT_RECxDealFillData(cDisPmtAll, nQtyDiv, cTotalAmtGrp)
'    Call PT_OBJxAddDiscount(orsWDeal, orsWDisDeal)
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub PT_CHKxDeal06(ByVal ptDealType$)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim cTotalAmtGrp As Double, cDisPmtAll As Double
    Dim nQtyDiv%

    'รอว่า จะต้องการ Return มาใน รูปแบบใด

    Set oGridDeal.DataSource = orsWDeal
    Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
    'wPSDealCal.Show 1
    'cDisPmtAll = PT_CALcDeal01_06(nQtyDiv, cTotalAmtGrp)
    'Call PT_CALxSaveResult(ptDealType)

'    Call PT_RECxDealFillData(cDisPmtAll, nQtyDiv, cTotalAmtGrp)
'    Call PT_OBJxAddDiscount(orsWDeal, orsWDisDeal)
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub PT_CHKxDeal07(ByVal ptDealType$)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim cTotalAmtGrp As Double, cDisPmtAll As Double
    Dim nQtyDiv%

    'รอว่า จะต้องการ Return มาใน รูปแบบใด

    Set oGridDeal.DataSource = orsWDeal
    Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
    'wPSDealCal.Show 1
    'cDisPmtAll = PT_CALcDeal07(nQtyDiv, cTotalAmtGrp)
    'Call PT_CALxSaveResult(ptDealType)

'    Call PT_RECxDealFillData(cDisPmtAll, nQtyDiv, cTotalAmtGrp)
'    Call PT_OBJxAddDiscount(orsWDeal, orsWDisDeal)
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub PT_CHKxDeal08(ByVal ptDealType$)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim cTotalAmtGrp As Double, cDisPmtAll As Double
    Dim nQtyDiv%

    'รอว่า จะต้องการ Return มาใน รูปแบบใด

    Set oGridDeal.DataSource = orsWDeal
    Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
    'wPSDealCal.Show 1
    'cDisPmtAll = PT_CALcDeal08(nQtyDiv, cTotalAmtGrp)
    'Call PT_CALxSaveResult(ptDealType)

'    Call PT_RECxDealFillData(cDisPmtAll, nQtyDiv, cTotalAmtGrp)
'    Call PT_OBJxAddDiscount(orsWDeal, orsWDisDeal)
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub PT_CHKxDeal09(ByVal ptDealType$)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim cTotalAmtGrp As Double, cDisPmtAll As Double
    Dim nQtyDiv%

    'รอว่า จะต้องการ Return มาใน รูปแบบใด

    Set oGridDeal.DataSource = orsWDeal
    Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
    'wPSDealCal.Show 1
    'cDisPmtAll = PT_CALcDeal09(nQtyDiv, cTotalAmtGrp)
     'Call PT_CALxSaveResult(ptDealType)

'    orsWDeal.Filter = adFilterNone
'    Call PT_RECxDealFillData(cDisPmtAll, nQtyDiv, cTotalAmtGrp)
'
'    orsWDeal.Save "C:\Deal" & Format(Time, "HHmmss") & ".xml", adPersistXML
'    Call PT_OBJxAddDiscount(orsWDeal, orsWDisDeal)
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub PT_CHKxDeal10(ByVal ptDealType$)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim cTotalAmtGrp As Double, cDisPmtAll As Double
    Dim nQtyDiv%

    'รอว่า จะต้องการ Return มาใน รูปแบบใด

    Set oGridDeal.DataSource = orsWDeal
    Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
    'wPSDealCal.Show 1
    'cDisPmtAll = PT_CALcDeal10(nQtyDiv, cTotalAmtGrp)
    'Call PT_CALxSaveResult(ptDealType)

'    Call PT_RECxDealFillData(cDisPmtAll, nQtyDiv, cTotalAmtGrp)
'    Call PT_OBJxAddDiscount(orsWDeal, orsWDisDeal)
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Public Function PT_CHKxDeal(ByVal ptDealNo As String, ByVal ptDealType As String) As VSFlexGrid
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    Dim tSql$
    Dim nCmdTimeout As Long
    Dim nCountRound As Long
    Dim tRepSta As Long
    
        'เอา DT  Join Deal Act ตาม ptDealType
        'tSql = "SELECT  * FROM " & tW_TBLDealAct & "  WHERE FTDacDealType='" & ptDealType & "'"
        tSql = "SELECT DISTINCT DEAL.FTPdhDealNo, DEAL.FTEvnCode, DEAL.FTPdhPndAct, DEAL.FTDacDealType, DEAL.FNDacSkuGrp, DEAL.FTDacSku"
        tSql = tSql & vbCrLf & " , DT.FNSdtSeqNo AS FNDacSkuSeq, DEAL.FCDacAmtBuy, DEAL.FNDacQtyBuy, DEAL.FNDacQtyGet, DT.FCSdtRegPrice AS FCDacSalePri"
        tSql = tSql & vbCrLf & " , DT.FCSdtQtyAll AS FNDacQtyKey, DT.FCSdtQtyAll*DT.FCSdtRegPrice AS FCDacAmtKey, DEAL.FTPddDisType"
        tSql = tSql & " , DEAL.FCDacDisAmt, DEAL.FCDacNewPriAll, DEAL.FCDalQtyCalDiv, DEAL.FCDacQtyDiv, DEAL.FCDacAmtDiv"
'        tSql = tSql & vbCrLf & " , CASE DEAL.FTPddDisType WHEN '0' THEN 0"
'        tSql = tSql & vbCrLf & "                                                             WHEN '1' THEN DEAL.FCDacSalePri * DEAL.FCDacDisAmt /100"
'        tSql = tSql & vbCrLf & "                                                             WHEN '2' THEN DEAL.FCDacSalePri * DEAL.FCDacDisAmt /100"
'        tSql = tSql & vbCrLf & "                                                             WHEN '3' THEN DEAL.FCDacSalePri - DEAL.FCDacDisAmt"
'        tSql = tSql & vbCrLf & "                                                             WHEN '4' THEN DEAL.FCDacSalePri END AS FCDacDisAmt"
'        tSql = tSql & vbCrLf & ", DEAL.FCDacNewPriAll, DEAL.FCDalQtyCalDiv, DEAL.FCDacQtyDiv, DEAL.FCDacAmtDiv"
        
        tSql = tSql & vbCrLf & " , DEAL.FCDacQtyMod, DEAL.FCDacAmtMod, DEAL.FCDacAmtTotal, DEAL.FCDacDisAvg"
        tSql = tSql & vbCrLf & " FROM " & tW_TBLDealDT & " AS DT INNER JOIN " & tW_TBLDealAct & " AS DEAL "
        tSql = tSql & vbCrLf & " ON DT.FTSkuCode = DEAL.FTDacSku"
        tSql = tSql & vbCrLf & " WHERE DEAL.FTPdhDealNo='" & ptDealNo & "'"
        tSql = tSql & vbCrLf & " AND DT.FTShdTransNo='" & tVB_Running & "'"
        
        '*KT 52-11-28 กรณีที่มีการทำงานกับ TPSTDealAct เกิด Replicate ขึ้นได้ให้รอจนกว่าจะ  Replicate สำเร็จ หรือ เวลาเกิน Timeoff หรือเป็น Offline
        nCmdTimeout = ocnAll.CommandTimeout
        tRepSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
        Do While tRepSta = "0" And nCountRound <= nCmdTimeout And oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP)
            nCountRound = nCountRound + 1
            SleepEx 1000, 0
            tRepSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
        Loop

        Call SP_DATbRsOpen(orsWDeal, adUseClient, tSql, ocnAll, adOpenDynamic, adLockBatchOptimistic, adCmdText)
        If (orsWDeal.BOF And orsWDeal.EOF) Then
            Call SP_DATxRsClose(orsWDeal)
            Exit Function
        End If
'*TW Test-------------------------------------
'        Set oGridDeal = wPSDealCal.ogdDeal      'ขอใช้ object
'        Set oGridDeal.DataSource = orsWDeal
'        Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
          'wPSDealCal.Show 1
'*TW Test-----------------------------------
        Set oGridDeal = wPSDealCal.ogdDeal      'ขอใช้ object
        Set oGridDeal.DataSource = orsWDeal
        Call SP_GRDxSetColumnKey(oGridDeal, orsWDeal)
        

        Select Case ptDealType 'รอว่า จะต้องการ Return มาใน รูปแบบใด
            Case "001", "002", "003", "004", "005", "006"
                 'Call PT_CHKxDeal06(ptDealType)
                 Call PT_CALcDeal01_06
            Case "007"
                 'Call PT_CHKxDeal07(ptDealType)
                 Call PT_CALcDeal07
            Case "008"
                 'Call PT_CHKxDeal08(ptDealType)
                Call PT_CALcDeal08
            Case "009"
                 'Call PT_CHKxDeal09(ptDealType)
                 Call PT_CALcDeal09
            Case "010"
                 'Call PT_CHKxDeal10(ptDealType)
                 Call PT_CALcDeal10
        End Select
        Call PT_CALxSaveResult(ptDealType)
        Call SP_DATxRsNothing(orsWDeal)
        'Set oGridDeal = Nothing'*TW
    Exit Function
ErrHandle:
    Call SP_MSGShowError
    Call SP_DATxRsNothing(orsWDeal)
End Function

Public Function PT_CHKbIsDealActive() As Boolean
'-----------------------------------------------------------
'   Call :
'   Cmt :DT Join Active Deal"
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
PT_CHKbIsDealActive = True
    Exit Function
ErrHandle:
    Call SP_MSGShowError
End Function

Public Function PT_CHKbIsDealAllow() As Boolean
'-----------------------------------------------------------
'   Call :
'   Cmt :DT Join Active Deal"
'   Developer : SOMJAI
'   Date Create : 03-05-2008 :  15:58:56
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    PT_CHKbIsDealAllow = True
'    If SP_TBLbChkExist(tTblPmt) Then
'        'TRUNCATE TABLE "TPT" & ptPosNo
'    Else
'        Call SP_TBLxCopyTable("SOURCE", "TARGET")       'Copy Scheme from Active Master
'        Call SP_TBLxAddIndex(tTblPmt, "FTPmhDocNo")
'    End
'    If PT_DATbGenData2Temp Then PT_CHKbIsDealAllow = True
    Exit Function
ErrHandle:
    PT_CHKbIsDealAllow = False
End Function
'
'Public Sub PT_SETxDealDefault(ByVal poGd As VSFlexGrid)
''-----------------------------------------------------------
''   Call :call after set data to grid
''   Cmt :For set colkey grid
''-----------------------------------------------------------
'Dim nCol%
'    On Error GoTo ErrHandle
'    With poGd
'        .Rows = 3 '.FixedRows
'        .Cols = 18
'        .ColKey(0) = "FNDacSkuGrp"
'        .ColKey(1) = "FCDacAmtBuy"
'        .ColKey(2) = "FNDacQtyBuy"
'        .ColKey(3) = "FNDacQtyGet"
'        .ColKey(4) = "FTDacSku"
'        .ColKey(5) = "FCDacSalePri"
'        .ColKey(6) = "FNDacQtyKey"
'        .ColKey(7) = "FCDacAmtKey"
'        .ColKey(8) = "FTPddDisType"
'        .ColKey(9) = "FCDacDisAmt"
'        .ColKey(10) = "FCDacNewPriAll"
'        .ColKey(11) = "FCDalQtyCalDiv"
'        .ColKey(12) = "FCDacQtyDiv"
'        .ColKey(13) = "FCDacAmtDiv"
'        .ColKey(14) = "FCDacQtyMod"
'        .ColKey(15) = "FCDacAmtMod"
'        .ColKey(16) = "FCDacAmtTotal"
'        .ColKey(17) = "FCDacDisAvg"
'
'        For nCol = 0 To .Cols - 1
'            .TextMatrix(0, nCol) = .ColKey(nCol)
'            .AutoSize nCol
'        Next nCol
'
'    End With
'    Exit Sub
'ErrHandle:
'    'Call SP_MSGShowError
'End Sub



