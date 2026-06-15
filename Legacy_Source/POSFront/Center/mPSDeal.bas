Attribute VB_Name = "mPSDeal"
Option Explicit
Public nDE_Group As Long
'Private cTotalDisAll As Double
Dim orsDE_AllDeal As ADODB.Recordset 'เก็บ Deal และรายละเอียด DT
Dim orsDE_AFSack As ADODB.Recordset 'เก็บ Deal และรายละเอียด DT ที่เกิดได้ทั้งหมดที่มีโอกาสเกิดได้จริง
Dim orsDE_DealSingle As ADODB.Recordset 'เก็บ Deal ที่จะเกิดเดี่ยว
Dim orsDE_DealMulti As ADODB.Recordset 'เก็บ Deal ที่จะเกิดร่วม
Dim nDE_MaxElement As Integer 'เก็บ จำนวนรอบที่จะเกิดร่วม

Dim atDE_AllDeal() As String 'Use
Dim acDE_AllDisc() As Double 'Maybe
Dim atDE_AllType() As String
Dim acDE_AllRegPrice() As Double
Dim atDE_AllGrp() As String

Dim acDE_SumDisc() As Double

Dim atDE_SkuSeqUsed() As String
Dim atDE_SkuQtyUsed() As String
Dim atDE_SkuStaUsed() As String 'สถานะของ Sku ที่ใช้ 0:ไม่ลด, 1:ลด

Dim tDE_Whe As String
Private Type UDT_MultiDeal
    tDealNo As String
    tFreeFlag As String
    tDealType As String
    tEventID As String
    tDiscType As String
    cDiscPer As Double
    cDiscAmt As Double
    cAlwDisc As Double
    tDTSeqNo As String 'ลำดับสินค้าที่ได้รับ Deal
    tDiscQty As String 'จำนวนที่ได้รับ Deal
    tSkuFlag As String '0 = ไม่ลด   1= ลด
    tRegPrice As String
End Type

Private Type UDT_KeepCD
    FNSdtSeqNo As Integer
    FTScdDiscType As String
    FNDctNo As Integer
    FTDctName As String
    FTDctRefCode As String
    FTScdEventID As String
    FTScdCardID As String
    FDScdExpDate As Date
    FTScdCardType As String
    FTScdStaTrans As String
    FNScdActSeqNo As Integer
    FTScdCstID As String
    FTScdCardKey As String
    FTScdStaChkExpire As String
    FTScdStaChkDepart As String
    FNScdSptSeqNo As Integer
    FCScdMaxAlwDisc As Double
    FNScdOrder As Integer
    FCScdAmtTime As Double
    FCScdAmtPmt As Double
    FNScdDealGroup As Integer
    FCScdDiscPerOrg As Double
    FCScdDiscAmtOrg As Double
    FCScdDiscNewOrg As Double
    FCScdAmt As Double
    FCScdDiscAmt As Double
    FCScdQty As Double
    FTSdtStaPreview As String
End Type

Private Type UDT_KeepDeal
    FTPdhDealNo As String
    FNDacSkuGrp As Integer
    FNSdtSeqNo As Integer
    FCSdtRegPrice As Double
    FCSdtQtyAll As Double
    FTDacSku As String
    FNDacQtyBuy As Integer
    FNDacQtyGet As Integer
    FTDacDealType As String
    FCDacDis As Double
    FTPddDisType As String
    FCDacDisAmt  As Double
    FCDacAmtBuy As Double
    FNPddLevel As Integer
    FTEvnCode As String
    FTFreeFlag As String
End Type

Private Type UDT_KeepMemDeal
    FTSkuCode As String
    FNSdtSeqNo As Long
End Type

Dim typeDE_KeepOld() As UDT_KeepCD
'Dim typeDE_KeepDeal() As UDT_KeepDeal '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls
Dim typeDE_MutiDeal() As UDT_MultiDeal
Dim nDE_KeepTime As Integer

'member price
Private bDE_MemberPrice As Boolean
Private bDE_MemberPriceAct As Boolean

Dim acDE_SumBlue() As Double
Dim acDE_Round() As Double

Private Function DE_PRCnChkRun0020Ex(ByVal porsTemp As ADODB.Recordset, ByVal ptDealType As String) As Integer
'___________________________________
'   Call:
'   Cmt:    '*Ao 2011/05/11 CommSheet ML-PRV-055 Not free item on Free Type2.xls
'___________________________________
Dim tFilter As String, tSort As String, tSKU As String
Dim tTemp As String, tTempEx As String

Dim nSeqNo As Integer, nPos As Integer, nQtyGet As Integer, nQtyCal As Integer, nCountEx As Integer
Dim nGrpMax As Integer, nCount As Integer, nLoop As Integer, nGrp As Integer, nQtyBuy As Integer

Dim bChk As Boolean, bRound As Boolean, bChk0020 As Boolean

On Error GoTo Err_Handler
    tVB_Free0020 = ""
    tTempEx = ""
    
    With porsTemp
        nGrpMax = 0
        tFilter = "(FNDacQtyBuy > 0)"               'GET
        .Filter = tFilter
        tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
        .Sort = tSort
        If .EOF = False Then
            Do While Not .EOF
                If nGrpMax < .Fields("FNDacSkuGrp").Value Then
                    nGrpMax = .Fields("FNDacSkuGrp").Value
                End If
                .MoveNext
            Loop
        End If
    End With
    
    With porsTemp
        .Filter = adFilterNone
        tFilter = "(FNDacQtyGet > 0)"               'GET
        .Filter = tFilter
        tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
        .Sort = tSort
        If .EOF = False Then
            .MoveFirst
            Do While Not .EOF
                tSKU = .Fields("FTDacSku").Value
                nSeqNo = .Fields("FNSdtSeqNo").Value
                nPos = .AbsolutePosition
                .Filter = adFilterNone
                tFilter = "FNDacQtyBuy > 0"                 'BUY
                .Filter = tFilter
                tSort = "FNDacSkuGrp"
                .Sort = tSort
                If .EOF = False Then
                    .MoveFirst
                    Do While Not .EOF
                        If (tSKU = .Fields("FTDacSku").Value) And (nSeqNo = .Fields("FNSdtSeqNo").Value) Then
                            tTemp = tTemp & "," & nSeqNo
                            Exit Do
                        End If
                        .MoveNext
                    Loop
                Else
                    Exit Function
                End If
                tFilter = "(FNDacQtyGet > 0)"               'GET
                .Filter = tFilter
                tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
                .Sort = tSort
                .AbsolutePosition = nPos
                .MoveNext
            Loop
        Else
            Exit Function
        End If
        If tTemp <> "" Then
            tTemp = Right(tTemp, Len(tTemp) - 1)
        End If
        
        bChk0020 = True
        
        Do While bChk0020
            bChk = True
            nCount = 0
            tTempEx = ""
            Do While bChk
                tFilter = "(FNDacQtyGet > 0)"               'GET
                .Filter = tFilter
                tSort = "FNSdtSeqNo DESC,FNDacSkuGrp"
                .Sort = tSort
                nCount = nCount + 1
                
                If .EOF = False Then
                    .MoveFirst
                    If nCount = 1 Then
                        nQtyGet = .Fields("FNDacQtyGet").Value
                    End If
                    Do While Not .EOF
                        nSeqNo = .Fields("FNSdtSeqNo").Value
                        nGrp = .Fields("FNDacSkuGrp").Value
                        If nCount = 1 Then
                            If InStr(1, tTemp, nSeqNo, vbBinaryCompare) > 0 Then
                                bRound = False
                            Else
                                bRound = True
                            End If
                        Else
                            bRound = True
                        End If
                        
                        If bRound Then
                            If .Fields("FCSdtQtyAll").Value > 0 Then
                                If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                                    nQtyCal = nQtyGet
                                    nQtyGet = 0
                                Else
                                    nQtyCal = .Fields("FCSdtQtyAll").Value
                                    nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                                End If
                                tTempEx = tTempEx & "," & .Fields("FNSdtSeqNo").Value
                                Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                            End If
                        End If
                        If nQtyGet = 0 Then Exit Do
                        .MoveNext
                    Loop
                    
                    If nQtyGet = 0 Then
                        bChk = False
                        bRound = True
                        bChk0020 = True
                    Else
                        If nCount = 1 Then
                            bChk = True
                            bRound = False
                            bChk0020 = False
                        Else
                            bChk = False
                            bRound = False
                            bChk0020 = False
                        End If
                    End If
                Else
                    bChk = False
                    bRound = False
                    bChk0020 = False
                End If
            Loop
            
            If bRound Then
                tFilter = "(FNDacQtyBuy > 0)"
                .Filter = tFilter
                tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
                .Sort = tSort
                If .EOF = False Then
                    .MoveFirst
                    For nLoop = 1 To nGrpMax
                        .Filter = adFilterNone
                        tFilter = "(FNDacQtyBuy > 0) AND (FNDacSkuGrp = " & nLoop & ")"
                        .Filter = tFilter
                        tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
                        .Sort = tSort
                        If .EOF = False Then
                            .MoveFirst
                            nGrp = .Fields("FNDacSkuGrp").Value
                            nQtyBuy = .Fields("FNDacQtyBuy").Value
                            nSeqNo = .Fields("FNSdtSeqNo").Value
                            Do While Not .EOF
                                
                                If .Fields("FCSdtQtyAll").Value > 0 Then
                                    If nQtyBuy > 0 Then
                                        If .Fields("FCSdtQtyAll").Value >= nQtyBuy Then
                                            nQtyCal = nQtyBuy
                                            nQtyBuy = 0
                                        Else
                                            nQtyCal = .Fields("FCSdtQtyAll").Value
                                            nQtyBuy = nQtyBuy - .Fields("FCSdtQtyAll").Value
                                        End If
                                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                                    End If
                                End If
                                .MoveNext
                            Loop
                            If nQtyBuy = 0 Then
                                bRound = True
                            Else
                                bRound = False
                                bChk0020 = False
                                Exit For
                            End If
                        Else
                            bRound = False
                            bChk0020 = False
                            Exit For
                        End If
                    Next nLoop
                        
                    If bRound Then
                        If tTempEx <> "" Then
                            tTempEx = Right(tTempEx, Len(tTempEx) - 1)
                            tVB_Free0020 = tVB_Free0020 & "," & tTempEx
                        End If
                        bChk0020 = True
                        nCountEx = nCountEx + 1
                    Else
                        bChk0020 = False
                    End If
                Else
                    bChk0020 = False                    'If .EOF = False Then
                End If
            Else                                                    'If bRound Then
                bChk0020 = False
            End If
        
        Loop
        
    End With
    If tVB_Free0020 <> "" Then
        tVB_Free0020 = Right(tVB_Free0020, Len(tVB_Free0020) - 1)
    End If
    DE_PRCnChkRun0020Ex = nCountEx
    
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_RSTxUpdateBySku0020(ByRef porsTemp As ADODB.Recordset, ByVal pnSeqNo As Integer, ByVal ptFieldName As String, ByVal pcValue As Double, ByVal ptFilter As String, ByVal ptSort As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nPos As Long
    Dim nI As Integer
    On Error GoTo Err_Handler
    With porsTemp
        .Filter = adFilterNone
        Do While Not .EOF
            If (.Fields("FNSdtSeqNo").Value = pnSeqNo) Then
                .Fields(ptFieldName).Value = pcValue
            End If
            .MoveNext
        Loop
        
        .Sort = ptSort
        
    End With
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function DE_PRCnChkRun0020(ByVal porsTemp As ADODB.Recordset, ByVal ptDealType As String) As Integer
'___________________________________
'   Call:
'   Cmt:    '*Ao 2011/05/11 CommSheet ML-PRV-055 Not free item on Free Type2.xls
'___________________________________
Dim nSeqNo As Integer
Dim tFilter As String, tSort As String, tSKU As String, tTemp As String, tTempEx As String
Dim nGrp As Integer, nGrpOld As Integer, nPos As Integer, nUncalGrp As Integer
Dim tSeqNo As String, tSeqNoEx As String
Dim cQtyGet  As Double, cRound As Double, cQtySum As Double, cQtyBuySum As Double, cQtyBuy As Double, cQtyCal As Double
Dim cRoundEx As Double
Dim bChk, bChkGetBuy As Boolean

On Error GoTo Err_Handler

    With porsTemp
        cRound = 0
        bChk = True
        Do While bChk
            .Filter = adFilterNone
            tSort = "FNDacSkuGrp DESC,FNSdtSeqNo"
            .Sort = tSort
            .MoveFirst
            nUncalGrp = .Fields("FNDacSkuGrp").Value
            nGrpOld = 0
            
            Do While Not .EOF
                nGrp = .Fields("FNDacSkuGrp").Value
                If nGrp = nUncalGrp Then
                    If nGrp <> nGrpOld Then
                        If nGrp = 10 Then
                            cQtyBuy = .Fields("FNDacQtyGet").Value
                        Else
                            cQtyBuy = .Fields("FNDacQtyBuy").Value
                        End If
                        nGrpOld = nGrp
                    End If
                    
                Else
                    If (cQtyBuy = 0) Then
                        bChk = True
                    Else
                        bChk = False
                        Exit Do
                    End If
                    If nGrp <> nGrpOld Then
                        If nGrp = 10 Then
                            cQtyBuy = .Fields("FNDacQtyGet").Value
                        Else
                            cQtyBuy = .Fields("FNDacQtyBuy").Value
                        End If
                        nGrpOld = nGrp
                    End If
                    nUncalGrp = .Fields("FNDacSkuGrp").Value
                End If
                
                If .Fields("FCSdtQtyAll").Value > 0 Then
                    If cQtyBuy > 0 Then
                        If .Fields("FCSdtQtyAll").Value >= cQtyBuy Then
                            cQtyCal = cQtyBuy
                            cQtyBuy = 0
                        Else
                            cQtyCal = .Fields("FCSdtQtyAll").Value
                            cQtyBuy = cQtyBuy - cQtyCal
                        End If
                        nPos = .AbsolutePosition
                        Call DE_RSTxUpdateBySku0020(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cQtyCal, "", tSort)
                        .AbsolutePosition = nPos
                    End If
                End If
                .MoveNext
            Loop
            
            If bChk = False Then
                bChk = False
            Else
                If (cQtyBuy = 0) And (nGrp = 1) Then
                    cRound = cRound + 1
                    bChk = True
                Else
                    bChk = False
                End If
            End If
            
        Loop
    End With
    
    DE_PRCnChkRun0020 = cRound
   
    Exit Function
Err_Handler:
    DE_PRCnChkRun0020 = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_RSTbDuplicate(pnSeqNo As Integer, ptTemp As String) As Boolean
'___________________________________________________
'   Call:
'   Cmt:    '*Ao 2011/05/13
'___________________________________________________
    Dim aSeqNo() As String
    Dim nI As Integer
    Dim bChk As Boolean
    On Error GoTo Err_Handler
    
    DE_RSTbDuplicate = True
    aSeqNo = Split(ptTemp, ",")
    For nI = 0 To UBound(aSeqNo)
        If CStr(pnSeqNo) = aSeqNo(nI) Then
            bChk = False
            Exit For
        Else
            bChk = True
        End If
        
    Next nI

    DE_RSTbDuplicate = bChk
    Exit Function
Err_Handler:
    DE_RSTbDuplicate = True
'    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_RSTxUpdateBySkuEx(ByRef porsTemp As ADODB.Recordset, ByVal pnSeqNo As Integer, ByVal ptFieldName As String, ByVal pcValue As Double, ByVal ptFilter As String, ByVal ptSort As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:  '*Ao 2011/05/11 CommSheet ML-PRV-055 Not free item on Free Type2.xls
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nPos As Long
    Dim nI As Integer
    On Error GoTo Err_Handler
    With porsTemp
        .Filter = adFilterNone
        Do While Not .EOF
            If .Fields("FNSdtSeqNo").Value = pnSeqNo Then
                .Fields(ptFieldName).Value = pcValue
            End If
            .MoveNext
        Loop
'        .Filter = ptFilter
'        .Sort = ptSort
        
    End With
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Private Function DE_PRCnChkRound0020(ByVal porsTemp As ADODB.Recordset, ByVal ptDealType As String) As Integer
''___________________________________
''   Call:
''   Cmt:    '*Ao 2011/05/11 CommSheet ML-PRV-055 Not free item on Free Type2.xls
''___________________________________
'Dim nCountRound As Integer, nSeqNo As Integer
'Dim nQtyBuy As Integer, nQtyGet As Integer, nQtyCal As Integer, nUncalGrp As Integer
'Dim tFilter As String, tSort As String, tSKU As String, tTemp As String, tTempEx As String
''Dim nBuy As Integer, nGet As Integer
'Dim bChk As Boolean, bRound As Boolean, bPOS As Boolean
'Dim bDuplicate  As Boolean
'', bRound As Boolean
''Dim nOldGrp As Integer
'Dim nGrp As Integer, nGrpOld As Integer, nPos As Integer
'
'On Error GoTo Err_Handler
'
'    nUncalGrp = 0
'    nGrpOld = 0
'    nCountRound = 0
''    tSort = "FNDacSkuGrp,FNSdtSeqNo"
'    With porsTemp
'        tFilter = "(FNDacQtyGet > 0)"               'GET
'        .Filter = tFilter
'        .Sort = tSort
'        If .EOF = False Then
'            .MoveFirst
'            tSKU = .Fields("FTDacSku").Value
'            nGet = .Fields("FNDacQtyGet").Value
'            bChk = True                         'bChk = True  = เอาตัว BUY มาเป็นตัว GET ด้วย และซื้อแต่ตัว GET มา
'            .MoveNext
'            Do While Not .EOF
'                If tSKU = .Fields("FTDacSku").Value Then
'                    nQtyGet = nQtyGet + .Fields("FCSdtQtyAll").Value                        'มีกลุ่มเดียว GROUP 10
'                    bChk = True
'                Else
'                    bChk = False
'                    Exit Do
'                End If
'                .MoveNext
'            Loop
'            If bChk = True Then
'                .Filter = adFilterNone
'                tFilter = "FNDacQtyBuy > 0"                 'BUY
''                tSort = "FNDacSkuGrp,FNSdtSeqNo"
'                .Filter = tFilter
'                .Sort = tSort
'                                                                        'ยังขาด CASE ที่ตั้ง GET มากกว่า 1 สินค้า แล้ว ซื้อมา BUY&GET
'                If .EOF = False Then
'                    .MoveFirst
'                    nBuy = .Fields("FNDacQtyBuy").Value
'                    nOldGrp = .Fields("FNDacSkuGrp").Value
'                    bChk = True                         'bChk = True  = เอาตัว BUY มาเป็นตัว GET ด้วย และซื้อแต่ตัว GET มา
'                    .MoveNext
'                    Do While Not .EOF
'                        If tSKU = .Fields("FTDacSku").Value Then
'                            nQtyBuy = nQtyBuy + .Fields("FCSdtQtyAll").Value
'                            If nOldGrp <> .Fields("FNDacSkuGrp").Value Then
'                                nBuy = nBuy + .Fields("FNDacQtyBuy").Value                          '+เก็บจำนวนที่ต้องซื้อ          'มีได้ 9 GROUP
'                                nOldGrp = .Fields("FNDacSkuGrp").Value
'                            End If
'                            bChk = True
'                        Else
'                            bChk = False
'                            Exit Do
'                        End If
'                    Loop
'                End If
'            End If
'
'            If bChk = True Then
'                nCountRound = (nQtyGet + nQtyBuy) \ (nBuy + nGet)
'            End If
'        Else
'            DE_PRCnChkRound0020 = 0
'            Exit Function
'        End If
'
'        '______________________________________________________________________
'        nQtyGet = 0
'
'        If bChk = False Then
'            .Filter = adFilterNone
'            tFilter = "FNDacQtyGet > 0"                 'GET
''            tSort = "FNDacSkuGrp,FNSdtSeqNo"
'            .Filter = tFilter
'            .Sort = tSort
'            If .EOF = False Then
'                .MoveFirst
'                nGet = .Fields("FNDacQtyGet").Value
'                Do While Not .EOF
'                    nQtyGet = nQtyGet + .Fields("FCSdtQtyAll").Value
'                    Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
'                    .MoveNext
'                Loop
'            Else
'                DE_PRCnChkRound0020 = 0
'                Exit Function
'            End If
'
'            nCountRound = nQtyGet \ nGet
'            nQtyBuy = 0
'
'            .Filter = adFilterNone
'            tFilter = "FNDacQtyBuy > 0"
''            tSort = "FNDacSkuGrp,FNSdtSeqNo"
'            .Filter = tFilter
'            .Sort = tSort
'            If .EOF = False Then
'                .MoveFirst
''                nBuy = .Fields("FNDacQtyBuy").Value
'                nOldGrp = .Fields("FNDacSkuGrp").Value
'                Do While Not .EOF
'                    If nOldGrp = .Fields("FNDacSkuGrp").Value Then
'                        nQtyBuy = nQtyBuy + .Fields("FCSdtQtyAll").Value
'                        nBuy = .Fields("FNDacQtyBuy").Value
'                    Else
'                        If nCountRound > (nQtyBuy \ nBuy) Then
'                            nCountRound = (nQtyBuy \ nBuy)
'                        End If
'                        nOldGrp = .Fields("FNDacSkuGrp").Value
'                        nQtyBuy = 0
'                    End If
'                    .MoveNext
'                Loop
'                If nCountRound > (nQtyBuy \ nBuy) Then              'GRPUO สุดท้าย
'                    nCountRound = (nQtyBuy \ nBuy)
'                End If
'            Else
'                nCountRound = 0
'            End If
'        Else
''            nCountRound = 0            'ได้มาจากตัวบนแล้ว
''            Exit Function           'ซื้อแต่ตัว BUY&GET ตัวเดียวกันทั้งหมด
'        End If
'    End With
'
'    DE_PRCnChkRound0020 = nCountRound
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_PRCnChkRound0020(ByVal porsTemp As ADODB.Recordset, ByVal ptDealType As String) As Integer
'___________________________________
'   Call:
'   Cmt:    '*Ao 2011/05/11 CommSheet ML-PRV-055 Not free item on Free Type2.xls
'___________________________________
Dim nSeqNo As Integer
Dim nQtyBuy As Integer, nQtyGet As Integer, nQtyCal As Integer, nUncalGrp As Integer
Dim tFilter As String, tSort As String, tSKU As String, tTemp As String, tTempEx As String
'Dim nBuy As Integer, nGet As Integer
Dim bChk As Boolean, bRound As Boolean, bPOS As Boolean
Dim bDuplicate  As Boolean
', bRound As Boolean
'Dim nOldGrp As Integer
Dim nGrp As Integer, nGrpOld As Integer, nPos As Integer
Dim tFilterBuy As String
Dim tSeqNo As String
Dim nLoop As Integer
Dim tQtyAll As String
Dim tQtyAllEx As String
Dim nCountRound As Integer
Dim nGrpMax As Integer

On Error GoTo Err_Handler

    nUncalGrp = 0
    nGrpOld = 0
    nCountRound = 0
'    tSort = "FNDacSkuGrp,FNSdtSeqNo"
    With porsTemp
        nGrpMax = 0
        tFilter = "(FNDacQtyBuy > 0)"               'GET
        .Filter = tFilter
        tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
        .Sort = tSort
        If .EOF = False Then
            Do While Not .EOF
                If nGrpMax < .Fields("FNDacSkuGrp").Value Then
                    nGrpMax = .Fields("FNDacSkuGrp").Value
                End If
                .MoveNext
            Loop
        End If
    End With
    
    With porsTemp
        .Filter = adFilterNone
        tFilter = "(FNDacQtyGet > 0)"               'GET
        .Filter = tFilter
        tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
        .Sort = tSort
        If .EOF = False Then
            .MoveFirst
            Do While Not .EOF
                tSKU = .Fields("FTDacSku").Value
                nSeqNo = .Fields("FNSdtSeqNo").Value
                nPos = .AbsolutePosition
                .Filter = adFilterNone
                tFilter = "FNDacQtyBuy > 0"                 'BUY
                .Filter = tFilter
                tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
                .Sort = tSort
                If .EOF = False Then
                    .MoveFirst
                    Do While Not .EOF
                        If (tSKU = .Fields("FTDacSku").Value) And (nSeqNo = .Fields("FNSdtSeqNo").Value) Then
                            tTemp = tTemp & "," & nSeqNo
                            Exit Do
                        End If
                        .MoveNext
                    Loop
                Else
                    Exit Function
                End If
                tFilter = "(FNDacQtyGet > 0)"               'GET
                .Filter = tFilter
                tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
                .Sort = tSort
                .AbsolutePosition = nPos
                .MoveNext
            Loop

        Else
            Exit Function
        End If
        If tTemp <> "" Then
            tTemp = Right(tTemp, Len(tTemp) - 1)
'            tTempEx = tTemp
        End If
        tFilterBuy = "(FNDacQtyBuy > 0)"
        Do While tTemp <> ""
            tSeqNo = cUT.UT_STRtToken(tTemp, ",", 1)
            tFilterBuy = tFilterBuy & " AND (FNSdtSeqNo <> " & tSeqNo & ")"
        Loop
                    
        bChk = True

        Do While bChk
            tFilter = "(FNDacQtyGet > 0)"               'GET
            .Filter = tFilter
            tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
            .Sort = tSort
            If .EOF = False Then
                .MoveFirst
                nQtyGet = .Fields("FNDacQtyGet").Value
                Do While Not .EOF
'                    nGrp = .Fields("FNDacSkuGrp").Value
                    If .Fields("FCSdtQtyAll").Value > 0 Then
                        If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                            nQtyCal = nQtyGet
                            nQtyGet = 0
                        Else
                            nQtyCal = .Fields("FCSdtQtyAll").Value
                            nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                        End If
                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                        If nQtyGet = 0 Then Exit Do
                    End If
                    .MoveNext
                Loop
                If nQtyGet = 0 Then
                    bRound = True
                    bChk = True
                Else
                    bRound = False
                    bChk = False
                End If
                
                If bRound Then
                    .Filter = adFilterNone
'                    tFilter = "FNDacQtyBuy > 0"                 'BUY
'                    tFilter = "(FNDacQtyBuy > 0) AND (FNSdtSeqNo NOT IN (" & tTemp & "))"
'                    tFilter = "(FNDacQtyBuy > 0)"
'                    Do While tTemp <> ""
'                        tSeqNo = cUT.UT_STRtToken(tTemp, ",", 1)
'                        tFilter = tFilter & " AND (FNSdtSeqNo <> " & tSeqNo & ")"
'                    Loop
                    .Filter = tFilterBuy
                    tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
                    .Sort = tSort
                    If .EOF = True Then                                                                         'ซื้อสินค้า ชนิดเดียว เป็นทั้ง BUY&GET
                        .Filter = "(FNDacQtyBuy > 0)"
                        tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
                        .Sort = tSort
                    End If
                    If .EOF = False Then
                        .MoveFirst
                        nUncalGrp = 0
                        nGrpOld = 0
                        nQtyBuy = 0
                        tTempEx = ""
                        tQtyAll = ""
                        Do While Not .EOF
                            nGrp = .Fields("FNDacSkuGrp").Value
                            If .Fields("FCSdtQtyAll").Value > 0 Then
                                If nGrp <> nUncalGrp Then
                                    If nGrpOld <> nGrp Then
                                        nQtyBuy = .Fields("FNDacQtyBuy").Value
                                        nGrpOld = nGrp
                                    End If
                                    
                                    If .Fields("FCSdtQtyAll").Value >= nQtyBuy Then
                                        nQtyCal = nQtyBuy
                                        nQtyBuy = 0
                                    Else
                                        nQtyCal = .Fields("FCSdtQtyAll").Value
                                        nQtyBuy = nQtyBuy - .Fields("FCSdtQtyAll").Value
                                    End If
                                    tTempEx = tTempEx & "," & .Fields("FNSdtSeqNo").Value
                                    tQtyAll = tQtyAll & "," & .Fields("FCSdtQtyAll").Value - nQtyCal
'                                    Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilterBuy, tSort)
                                    If nQtyBuy < 1 Then nUncalGrp = nGrp
                                End If
                            End If
                            .MoveNext
                        Loop

                        If (nQtyBuy = 0) And (nUncalGrp <> 0) Then
                            If tTempEx <> "" Then
                                tTempEx = Right(tTempEx, Len(tTempEx) - 1)
                                tQtyAll = Right(tQtyAll, Len(tQtyAll) - 1)
                            End If
                            Do While tTempEx <> ""
                                tSeqNo = cUT.UT_STRtToken(tTempEx, ",", 1)
                                tQtyAllEx = cUT.UT_STRtToken(tQtyAll, ",", 1)
'                                Call DE_RSTxUpdateBySku(porsTemp, CInt(tSeqNo), "FCSdtQtyAll", CInt(tQtyAllEx), tFilterBuy, tSort)
                                Call DE_RSTxUpdateBySkuEx(porsTemp, CInt(tSeqNo), "FCSdtQtyAll", CDbl(tQtyAllEx), tFilterBuy, tSort)        '***********************************
                            Loop
                            If nGrpMax = nGrp Then                          '*Aoo 2012/08/21  CommSheet ML-PH4-R3-035 Wrong Free Item for Free Type2.xls
                                nCountRound = nCountRound + 1
                            End If
'                            nCountRound = nCountRound + 1
                            bChk = True
                        Else
'                            tFilterBuy = "(FNDacQtyBuy > 0)"
                            .Filter = "(FNDacQtyBuy > 0)"
                            tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
                            .Sort = tSort
                            If .EOF = False Then
                                .MoveFirst
                                nUncalGrp = 0
                                nGrpOld = 0
                                nQtyBuy = 0
                                Do While Not .EOF
                                    nGrp = .Fields("FNDacSkuGrp").Value
                                    If .Fields("FCSdtQtyAll").Value > 0 Then
                                        If nGrp <> nUncalGrp Then
                                            If nGrpOld <> nGrp Then
                                                nQtyBuy = .Fields("FNDacQtyBuy").Value
                                                nGrpOld = nGrp
                                            End If
                                            
                                            If .Fields("FCSdtQtyAll").Value >= nQtyBuy Then
                                                nQtyCal = nQtyBuy
                                                nQtyBuy = 0
                                            Else
                                                nQtyCal = .Fields("FCSdtQtyAll").Value
                                                nQtyBuy = nQtyBuy - .Fields("FCSdtQtyAll").Value
                                            End If
                                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, .Filter, tSort)
'                                            Call DE_RSTxUpdateBySkuEx(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilterBuy, tSort)
                                            If nQtyBuy < 1 Then nUncalGrp = nGrp
                                        End If
                                    End If
                                    .MoveNext
                                Loop
                                If (nQtyBuy = 0) And (nUncalGrp <> 0) Then
                                    nCountRound = nCountRound + 1
                                    bChk = True
                                Else
                                    bChk = False
                                End If
                            Else
                                bChk = False
                            End If
                            
                        End If
                    Else
                        bChk = False
                    End If      'If .EOF = False Then   'BUY
                Else
                    bChk = False
                End If
            Else
                bChk = False
            End If          'If .EOF = False Then   'GET

        Loop
    End With

    DE_PRCnChkRound0020 = nCountRound
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Function DE_PRCnChkRound0020(ByVal porsTemp As ADODB.Recordset, ByVal ptDealType As String) As Integer
''___________________________________
''   Call:
''   Cmt:    '*Ao 2011/05/11 CommSheet ML-PRV-055 Not free item on Free Type2.xls
''___________________________________
'Dim nCountRound As Integer, nSeqNo As Integer
'Dim nQtyBuy As Integer, nQtyGet As Integer, nQtyCal As Integer, nUncalGrp As Integer
'Dim tFilter As String, tSort As String, tSKU As String, tTemp As String, tTempEx As String
''Dim nBuy As Integer, nGet As Integer
'Dim bChk As Boolean, bRound As Boolean, bPOS As Boolean
'Dim bDuplicate  As Boolean
'', bRound As Boolean
''Dim nOldGrp As Integer
'Dim nGrp As Integer, nGrpOld As Integer, nPos As Integer
'
'On Error GoTo Err_Handler
'
'    nUncalGrp = 0
'    nGrpOld = 0
'    nCountRound = 0
''    tSort = "FNDacSkuGrp,FNSdtSeqNo"
'    With porsTemp
'        tFilter = "(FNDacQtyGet > 0)"               'GET
'        .Filter = tFilter
'        tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
'        .Sort = tSort
'        If .EOF = False Then
'            .MoveFirst
'            Do While Not .EOF
'                tSKU = .Fields("FTDacSku").Value
'                nSeqNo = .Fields("FNSdtSeqNo").Value
'                nPos = .AbsolutePosition
'                .Filter = adFilterNone
'                tFilter = "FNDacQtyBuy > 0"                 'BUY
'                .Filter = tFilter
'                tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
'                .Sort = tSort
'                If .EOF = False Then
'                    .MoveFirst
'                    Do While Not .EOF
'                        If (tSKU = .Fields("FTDacSku").Value) And (nSeqNo = .Fields("FNSdtSeqNo").Value) Then
'                            tTemp = tTemp & "," & nSeqNo
'                            Exit Do
'                        End If
'                        .MoveNext
'                    Loop
'                Else
'                    Exit Function
'                End If
'                tFilter = "(FNDacQtyGet > 0)"               'GET
'                .Filter = tFilter
'                tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
'                .Sort = tSort
'                .AbsolutePosition = nPos
'                .MoveNext
'            Loop
'
'        Else
'            Exit Function
'        End If
'        If tTemp <> "" Then
'            tTemp = Right(tTemp, Len(tTemp) - 1)
'            tTempEx = tTemp
'        End If
'
'        bChk = True
'
'        Do While bChk
'            tFilter = "(FNDacQtyGet > 0)"               'GET
'            .Filter = tFilter
'            tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
'            .Sort = tSort
'            If .EOF = False Then
'                .MoveFirst
'                nQtyGet = .Fields("FNDacQtyGet").Value
'                Do While Not .EOF
'    '                nGrp = .Fields("FNDacSkuGrp").Value
'                    If .Fields("FCSdtQtyAll").Value > 0 Then
'                        If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                            nQtyCal = nQtyGet
'                            nQtyGet = 0
'                        Else
'                            nQtyCal = .Fields("FCSdtQtyAll").Value
'                            nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
'                        End If
'                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'                        If nQtyGet = 0 Then Exit Do
'                    End If
'                    .MoveNext
'                Loop
'                If nQtyGet = 0 Then
'                    bRound = True
'                Else
'                    bRound = False
'                    bChk = False
'                End If
'
'                If bRound Then
'                    .Filter = adFilterNone
'                    tFilter = "FNDacQtyBuy > 0"                 'BUY
'                    .Filter = tFilter
'                    tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
'                    .Sort = tSort
'                    If .EOF = False Then
'                        .MoveFirst
'                        nUncalGrp = 0
'                        nGrpOld = 0
'                        nQtyBuy = 0
'                        Dim nCount  As Integer
'                        Do While Not .EOF
'                            nGrp = .Fields("FNDacSkuGrp").Value
'                            If .Fields("FCSdtQtyAll").Value > 0 Then
'                                If nGrp <> nUncalGrp Then
'                                    If nGrpOld <> nGrp Then
'                                        If nQtyBuy > 0 Then                                     'สินค้าไม่ครบจำนวน
'                                            If nCount = 1 Then
'                                                .AbsolutePosition = nPos
'                                                bPOS = True                                             'Check ดพื่อเก็บตำแหน่งตัวแรกว่า ซ้ำกันไว้ก่อน
'                                                tTempEx = ""
'                                                nCount = 2
'                                            Else
'                                                Exit Do
'                                            End If
'                                        Else
'                                            nQtyBuy = .Fields("FNDacQtyBuy").Value
''                                            nGrpOld = nGrp
'                                            bPOS = True
'                                            nCount = 1
'                                            tTempEx = tTemp
'                                        End If
'                                        nGrpOld = nGrp
'                                    End If
'
'                                    If tTempEx <> "" Then
'                                        bDuplicate = DE_RSTbDuplicate(.Fields("FNSdtSeqNo").Value, tTemp)
'                                        If bDuplicate = True Then
'                                            If .Fields("FCSdtQtyAll").Value >= nQtyBuy Then
'                                                nQtyCal = nQtyBuy
'                                                nQtyBuy = 0
'                                            Else
'                                                nQtyCal = .Fields("FCSdtQtyAll").Value
'                                                nQtyBuy = nQtyBuy - .Fields("FCSdtQtyAll").Value
'                                            End If
'                                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'                                            If nQtyBuy < 1 Then nUncalGrp = nGrp
'                                            bPOS = True
'                                        Else
'                                            If bPOS = True Then
'                                                nPos = .AbsolutePosition
'                                            End If
'
'                                        End If
'                                    Else
'
'                                        If .Fields("FCSdtQtyAll").Value >= nQtyBuy Then
'                                            nQtyCal = nQtyBuy
'                                            nQtyBuy = 0
'                                        Else
'                                            nQtyCal = .Fields("FCSdtQtyAll").Value
'                                            nQtyBuy = nQtyBuy - .Fields("FCSdtQtyAll").Value
'                                        End If
'                                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'                                        If nQtyBuy < 1 Then nUncalGrp = nGrp
'                                    End If
'
'                                End If
'                            End If
'                            .MoveNext
'                        Loop
'
'                        If nQtyBuy = 0 Then
'                            nCountRound = nCountRound + 1
'                            bChk = True
'                        Else
'                            bChk = False
'                        End If
'                    Else
'                        bChk = False
'                    End If      'If .EOF = False Then   'BUY
'                End If
'            Else
'                bChk = False
'            End If          'If .EOF = False Then   'GET
'
'        Loop
'    End With
'
'    DE_PRCnChkRound0020 = nCountRound
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Sub DE_CHKxFreeTypeVendor()
'______________________________________
'   Call:
'   Cmt: '*Ao 2011/01/15
'______________________________________
Dim tSql As String, tSdtSeqNo As String
Dim orsTemp As ADODB.Recordset
        On Error GoTo Err_Handler
        
        tSql = "SELECT FNSdtSeqNo,FNScdSeqNo,FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate"
        tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD
        tSql = tSql & vbCrLf & "WHERE"
        tSql = tSql & vbCrLf & "(FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & vbCrLf & "AND  FTShdTransType = '" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "')"
        tSql = tSql & vbCrLf & "AND FTScdDiscType = '4' AND FNDctNo = '1'"
        tSql = tSql & vbCrLf & "AND FCScdAmt > 0"
        
        Call SP_SQLvExecute(tSql, orsTemp)
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            Do While Not orsTemp.EOF
                tSdtSeqNo = CStr(SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo"))
                tSql = "UPDATE " & tVB_CNTblRC
                tSql = tSql & vbCrLf & "SET FTSrcStaNoPay='1'"
                tSql = tSql & vbCrLf & "WHERE (FTSkuIndex= '" & tSdtSeqNo & "')"
                tSql = tSql & vbCrLf & "AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
                Call SP_SQLvExecute(tSql)
                orsTemp.MoveNext
            Loop
        End If
        
        Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DE_CHKxDelFreeTypeVendor()
'______________________________________
'   Call:
'   Cmt: '*Ao 2011/01/15
'______________________________________
Dim tSql As String, tSdtSeqNo As String
Dim orsTemp As ADODB.Recordset
Dim nI As Integer, nSeqNo  As Integer

    On Error GoTo Err_Handler
    tSql = "SELECT FNSdtSeqNo,FNScdSeqNo,FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & "WHERE"
    tSql = tSql & vbCrLf & "(FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND  FTShdTransType = '" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "')"
    tSql = tSql & vbCrLf & "AND FTScdDiscType = '4' AND FNDctNo = '1'"
    tSql = tSql & vbCrLf & "AND FCScdAmt > 0"

    Call SP_SQLvExecute(tSql, orsTemp)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        Do While Not orsTemp.EOF
            tSdtSeqNo = CStr(SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo"))
            tSql = "DELETE FROM " & tVB_CNTblRC
            tSql = tSql & vbCrLf & "WHERE"
            tSql = tSql & vbCrLf & "(FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
            tSql = tSql & vbCrLf & "AND  FTShdTransType = '" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "')"
            tSql = tSql & vbCrLf & "AND (FTSkuIndex = '" & tSdtSeqNo & "')"
            Call SP_SQLvExecute(tSql)
            orsTemp.MoveNext
        Loop
        tSql = "SELECT * FROM " & tVB_CNTblRC
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & "ORDER BY FNSrcSeqNo"

        Call SP_SQLvExecute(tSql, orsTemp)
         If Not (orsTemp.EOF And orsTemp.BOF) Then
            For nI = 1 To orsTemp.RecordCount
                nSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSrcSeqNo")
                tSql = "UPDATE " & tVB_CNTblRC
                tSql = tSql & vbCrLf & "SET FNSrcSeqNo=" & nI
                tSql = tSql & vbCrLf & "WHERE FNSrcSeqNo=" & nSeqNo
                tSql = tSql & vbCrLf & "AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
                Call SP_SQLvExecute(tSql)
                orsTemp.MoveNext
            Next nI
        End If
    End If
    
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function DE_GETcSumDiscByAmt(ByRef porsTemp As ADODB.Recordset, ByVal pcAmtGet As Double, Optional ByVal ptFilter As String, Optional ByVal ptSort As String) As Double
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:   Deal Type 5
'   Cmt:        '*Ao 09-10-29
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Dim cAmtBuy As Double, cDelRegPrice As Double, cQtyAll As Double, cDelSaleAmt As Double, cAmtCal As Double, cTemp As Double
Dim nQtyCal As Integer

On Error GoTo Err_Handler
    
    With porsTemp
        .Filter = adFilterNone
        If Not (ptFilter = "") Then
            .Filter = ptFilter
        End If
        If Not (ptSort = "") Then
            .Sort = ptSort
        End If
         cAmtBuy = pcAmtGet
         If .BOF And .EOF Then Exit Function
        .MoveFirst              'MoveFirst
        Do While Not .EOF
            If .Fields("FCSdtQtyAll").Value > 0 Then
                cDelRegPrice = .Fields("FCSdtRegPrice").Value
                cQtyAll = .Fields("FCSdtQtyAll").Value
                cDelSaleAmt = cDelRegPrice * cQtyAll
                
                If cDelSaleAmt >= cAmtBuy Then
                    cAmtCal = cAmtBuy
                    cAmtBuy = 0
                Else
                    cAmtCal = cDelSaleAmt
                    cAmtBuy = cAmtBuy - cDelSaleAmt
                End If
                
                nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal \ cDelRegPrice = cAmtCal / cDelRegPrice, 0, 1)
                
                cTemp = cTemp + (.Fields("FCSdtRegPrice").Value * nQtyCal)
                If cAmtBuy <= 0 Then
                    Exit Do
                End If
            End If
            .MoveNext
        Loop
        .Filter = adFilterNone

    End With

    DE_GETcSumDiscByAmt = cTemp
    Exit Function
        
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_GETcSumDiscByQtyGrp(ByRef porsTemp As ADODB.Recordset, ByVal pnQtyGet As Integer, ByVal pnGrp As Integer, ByVal ptSort As String, ByVal pnCountRound As Long) As Double
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:   Deal Type 2,Deal Type 4
'   Cmt:        '*Ao 09-10-29
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nQtyCal As Integer, nQtyGet As Integer, nMaxGrp As Integer, nLoop As Integer
    Dim cTemp As Double
    Dim nGrp As Integer, nUncalGrp As Integer
    
    On Error GoTo Err_Handler
    With porsTemp
        .Sort = ptSort
         
         If .BOF And .EOF Then Exit Function
        .MoveFirst
        nQtyGet = .Fields("FNDacQtyGet").Value * CInt(pnCountRound)
        Do While Not .EOF
            If .Fields("FCSdtQtyAll").Value > 0 Then
                nGrp = .Fields("FNDacSkuGrp").Value
                If nGrp <> nUncalGrp Then
                    If nQtyGet = 0 Then
                        nQtyGet = .Fields("FNDacQtyGet").Value * CInt(pnCountRound)
                    End If
                End If
                If nGrp <> nUncalGrp Then
                    If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                        nQtyCal = nQtyGet
                        nQtyGet = 0
                    Else
                        nQtyCal = .Fields("FCSdtQtyAll").Value
                        nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                    End If
                    
                    cTemp = cTemp + (.Fields("FCSdtRegPrice").Value * nQtyCal)
                    If nQtyGet <= 0 Then
                        nUncalGrp = nGrp
                    End If
                End If
            End If
            .MoveNext
        Loop
       
        .Filter = adFilterNone
    End With
    DE_GETcSumDiscByQtyGrp = cTemp
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_SINxDealAmt(ByVal ptLnDiscType As String, ByVal pcQtyGet As Double, ByVal pcDacDisAmt As Double, ByVal pcDiscTotalAll As Double, ByRef pcDealDiscAmt As Double, Optional ByVal pcSumCanDisc As Double = 0, Optional ByVal ptDealType As String)
'___________________________________________
'   Call:
'   Cmt:'*Ao 52-04-20
'___________________________________________
On Error GoTo Err_Handler
    Select Case ptDealType
        Case "005"
            Select Case ptLnDiscType
                Case "1"        '%
                    If pcQtyGet <= 0 Then
                        If pcDiscTotalAll = pcDacDisAmt Then Exit Sub
                        If pcDiscTotalAll > pcDacDisAmt Then
                            pcDealDiscAmt = pcDealDiscAmt - (pcDiscTotalAll - pcDacDisAmt)
                        Else
                            If pcDiscTotalAll < pcDacDisAmt Then
                                pcDealDiscAmt = pcDealDiscAmt + (pcDacDisAmt - pcDiscTotalAll)
                            End If
                        End If
                    End If
                Case "2"        'Amt
                    If pcQtyGet <= 0 Then
                        If pcDiscTotalAll = pcDacDisAmt Then Exit Sub
                        If pcDiscTotalAll > pcDacDisAmt Then
                            pcDealDiscAmt = pcDealDiscAmt - (pcDiscTotalAll - pcDacDisAmt)
                        Else
                            If pcDiscTotalAll < pcDacDisAmt Then
                                pcDealDiscAmt = pcDealDiscAmt + (pcDacDisAmt - pcDiscTotalAll)
                            End If
                        End If
                    End If
                Case "3"        'New
                    If pcQtyGet <= 0 Then
                        If pcDiscTotalAll = pcSumCanDisc - pcDacDisAmt Then Exit Sub
                            If pcDiscTotalAll > pcSumCanDisc - pcDacDisAmt Then
                                pcDealDiscAmt = pcDealDiscAmt - (pcDiscTotalAll - (pcSumCanDisc - pcDacDisAmt))
                            Else
                                If pcDiscTotalAll < pcSumCanDisc - pcDacDisAmt Then
                                    pcDealDiscAmt = pcDealDiscAmt + ((pcSumCanDisc - pcDacDisAmt) - pcDiscTotalAll)
                                End If
                            End If
                    End If
            End Select
        Case Else
            Select Case ptLnDiscType
                Case "2" 'Amt
                    If pcQtyGet <= 0 Then
                        If pcDiscTotalAll = pcDacDisAmt Then Exit Sub
                        If pcDiscTotalAll > pcDacDisAmt Then
                            pcDealDiscAmt = pcDealDiscAmt - (pcDiscTotalAll - pcDacDisAmt)
                        Else
                            If pcDiscTotalAll < pcDacDisAmt Then
                                pcDealDiscAmt = pcDealDiscAmt + (pcDacDisAmt - pcDiscTotalAll)
                            End If
                        End If
                    End If
                Case "3"
                    If pcQtyGet <= 0 Then
                        If pcDiscTotalAll = pcSumCanDisc - pcDacDisAmt Then Exit Sub
                            If pcDiscTotalAll > pcSumCanDisc - pcDacDisAmt Then
                                pcDealDiscAmt = pcDealDiscAmt - (pcDiscTotalAll - (pcSumCanDisc - pcDacDisAmt))
                            Else
                                If pcDiscTotalAll < pcSumCanDisc - pcDacDisAmt Then
                                    pcDealDiscAmt = pcDealDiscAmt + ((pcSumCanDisc - pcDacDisAmt) - pcDiscTotalAll)
                                End If
                            End If
                    End If
            End Select
    End Select
   Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub
Private Sub DE_MulxDealAmt(ByVal ptLnDiscType As String, ByVal pcQtyGet As Double, ByVal pcDacDisAmt As Double, ByVal pcDiscTotalAll As Double, ByRef pcDealDiscAmt As Double, Optional ByVal pcSumCanDisc As Double = 0, Optional ByVal ptDealType As String)
'___________________________________________
'   Call:
'   Cmt:'*Ao 52-04-20
'___________________________________________
On Error GoTo Err_Handler
    Select Case ptDealType
        Case "005"
            Select Case ptLnDiscType
                Case "1", "2", "3"      '%
                    If pcQtyGet <= 0 Then
                        If pcDiscTotalAll = pcSumCanDisc Then Exit Sub
                        If pcDiscTotalAll > pcSumCanDisc Then
                            pcDealDiscAmt = pcDealDiscAmt - (pcDiscTotalAll - pcSumCanDisc)
                        Else
                            If pcDiscTotalAll < pcSumCanDisc Then
                                pcDealDiscAmt = pcDealDiscAmt + (pcSumCanDisc - pcDiscTotalAll)
                            End If
                        End If
                    End If
'                Case "2"        'Amt
'                    If pcQtyGet <= 0 Then
'                        If pcDiscTotalAll = pcSumCanDisc Then Exit Sub
'                        If pcDiscTotalAll > pcSumCanDisc Then
'                            pcDealDiscAmt = pcDealDiscAmt - (pcDiscTotalAll - pcSumCanDisc)
'                        Else
'                            If pcDiscTotalAll < pcSumCanDisc Then
'                                pcDealDiscAmt = pcDealDiscAmt + (pcSumCanDisc - pcDiscTotalAll)
'                            End If
'                        End If
'                    End If
'                Case "3"        'New
'                    If pcQtyGet <= 0 Then
'                        If pcDiscTotalAll = pcSumCanDisc Then Exit Sub
'                            If pcDiscTotalAll > pcSumCanDisc Then
'                                pcDealDiscAmt = pcDealDiscAmt - (pcDiscTotalAll - pcSumCanDisc)
'                            Else
'                                If pcDiscTotalAll < pcSumCanDisc Then
'                                    pcDealDiscAmt = pcDealDiscAmt + (pcSumCanDisc - pcDiscTotalAll)
'                                End If
'                            End If
'                    End If
            End Select
        Case Else
                Select Case ptLnDiscType
                    Case "2"        'Amt
                        If pcQtyGet <= 0 Then
                            If pcDiscTotalAll = pcSumCanDisc Then Exit Sub
                            If pcDiscTotalAll > pcSumCanDisc Then
                                pcDealDiscAmt = pcDealDiscAmt - (pcDiscTotalAll - pcSumCanDisc)
                            Else
                                If pcDiscTotalAll < pcSumCanDisc Then
                                    pcDealDiscAmt = pcDealDiscAmt + (pcSumCanDisc - pcDiscTotalAll)
                                End If
                            End If
                        End If
                    Case "3"        'New
                        If pcQtyGet <= 0 Then
                            If pcDiscTotalAll = pcSumCanDisc Then Exit Sub
                                If pcDiscTotalAll > pcSumCanDisc Then
                                    pcDealDiscAmt = pcDealDiscAmt - (pcDiscTotalAll - pcSumCanDisc)
                                Else
                                    If pcDiscTotalAll < pcSumCanDisc Then
                                        pcDealDiscAmt = pcDealDiscAmt + (pcSumCanDisc - pcDiscTotalAll)
                                    End If
                                End If
                        End If
                End Select
    End Select
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DE_PRCxCompress(ByVal porsTemp As ADODB.Recordset)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       # ตัด Deal ที่ซ้ำกันออก
'ARGUMENT LIST # porsTemp = Deal Recordset
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    Dim tDealNo As String
    With porsTemp
        .Sort = "FTPdhDealNo"
        Do While Not .EOF
            If .Fields("FTPdhDealNo").Value <> tDealNo Then
                tDealNo = .Fields("FTPdhDealNo").Value
            Else
                .Delete
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Public Sub DE_PRCxDeal()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    'tUT_TextPrompt = Trim(wMain.otbPrompt.Text)
'    bVB_CNbHistoryHaveDael = False
    Call DE_TESxCheckDeal
    
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DE_PRCxGetDeal(ByVal ptDealSeqOrder As String, ByVal ptSeqSku As String, ByVal ptQtySku As String, ByVal ptStaSku As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nI As Integer
    Dim tSeqSku As String
    Dim tQtySku As String
    Dim tStaSku As String
    Dim nDealSeq As Integer
    Dim tDealNo As String
    Dim tSeqSkuOrg As String
    Dim tQtySkuOrg As String
    Dim tStaSkuOrg As String
    Dim tSql As String
    Dim cRegPrice As Double
    Dim cSumReg As Double
    Dim cSumRegCal As Double
    Dim cSumDisc As Double
    Dim cDisc As Double
    Dim tKeepSeqNo As String
    Dim tKeepQty As String
    Dim tKeepSta As String
    
    Dim tKeepSeqDT As String
    Dim nIndex As Integer
    On Error GoTo Err_Handler
    tSeqSkuOrg = ptSeqSku
    tQtySkuOrg = ptQtySku
    tStaSkuOrg = ptStaSku
    
    Do While Left(ptDealSeqOrder, 1) = "#"
        ptDealSeqOrder = Right(ptDealSeqOrder, Len(ptDealSeqOrder) - 1)
    Loop
    
    If ptDealSeqOrder <> "" Then
        nDealSeq = SP_VALcText2Double(UT_tToken(ptDealSeqOrder, "#", 1, False))
        tDealNo = atDE_AllDeal(nDealSeq)
        tSeqSku = UT_tToken(tSeqSkuOrg, ";", 1, False)
        tQtySku = UT_tToken(tQtySkuOrg, ";", 1, False)
        tStaSku = UT_tToken(tStaSkuOrg, ";", 1, False)
        tKeepSeqNo = tSeqSku
        tKeepQty = tQtySku
        tKeepSta = tStaSku
            
        Do While tDealNo <> ""
            nIndex = nIndex + 1
            '*Ao 09-10-17
'                If atDE_AllType(nDealSeq) = "0010" Then
'                    Call DE_CALxMultiByDealNo0010(tDealNo, atDE_AllType(nDealSeq), tSeqSku, tQtySku, tStaSku, "WHERE " & tVB_CNWheBill, nIndex)
'                Else
'                    Call DE_CALxMultiByDealNo(tDealNo, atDE_AllType(nDealSeq), tSeqSku, tQtySku, tStaSku, "WHERE " & tVB_CNWheBill, nIndex)
'                End If
            Call DE_CALxMultiByDealNo(tDealNo, atDE_AllType(nDealSeq), tSeqSku, tQtySku, tStaSku, "WHERE " & tVB_CNWheBill, nIndex)
            nDealSeq = SP_VALcText2Double(UT_tToken(ptDealSeqOrder, "#", 1, False))
            tDealNo = atDE_AllDeal(nDealSeq)
            tSeqSku = UT_tToken(tSeqSkuOrg, ";", 1, False)
            tQtySku = UT_tToken(tQtySkuOrg, ";", 1, False)
            tStaSku = UT_tToken(tStaSkuOrg, ";", 1, False)
        Loop
            
            'คืนค่า Avi ให้ DT ตรงนี้
        tSql = "UPDATE " & tVB_CNTblDT
        tSql = tSql & vbCrLf & "SET FCSdtQtyAllDisAvi = ROUND(FCSdtQtyAll,2)"
        tSql = tSql & vbCrLf & "WHERE " & tVB_CNWheBill
        Call SP_SQLbExecute(tSql)
        
        For nI = LBound(typeDE_MutiDeal) To UBound(typeDE_MutiDeal)
            If nI <> 0 Then
                Call DE_MULxSendDeal(nI)
            End If
        Next nI
    End If
    
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DE_PRCxStoreDiscount()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim nI As Integer
    On Error GoTo Err_Handler
    
    tSql = "SELECT * FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND  FTShdTransType = '" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " ORDER BY FNScdSeqNo"
    'Set orsTemp = ocnAll.Execute(tSql) '*KT 54-03-24 ป้องกันโปรแกรมหลุด
    If SP_SQLvExecute(tSql, orsTemp, ocnAll) = 0 Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            Do While Not orsTemp.EOF
                ReDim Preserve typeDE_KeepOld(nI) As UDT_KeepCD
                typeDE_KeepOld(nI).FNSdtSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                typeDE_KeepOld(nI).FTScdDiscType = SP_FEDtChkString(orsTemp, "FTScdDiscType")
                typeDE_KeepOld(nI).FNDctNo = SP_FEDvChkNumeric(orsTemp, "FNDctNo")
                typeDE_KeepOld(nI).FTDctName = SP_FEDtChkString(orsTemp, "FTDctName")
                typeDE_KeepOld(nI).FTDctRefCode = SP_FEDtChkString(orsTemp, "FTDctRefCode")
                typeDE_KeepOld(nI).FTScdEventID = SP_FEDtChkString(orsTemp, "FTScdEventID")
                typeDE_KeepOld(nI).FTScdCardID = SP_FEDtChkString(orsTemp, "FTScdCardID")
                typeDE_KeepOld(nI).FDScdExpDate = SP_FEDdChkDate(orsTemp, "FDScdExpDate")
                typeDE_KeepOld(nI).FTScdCardType = SP_FEDtChkString(orsTemp, "FTScdCardType")
                typeDE_KeepOld(nI).FTScdStaTrans = SP_FEDtChkString(orsTemp, "FTScdStaTrans")
                typeDE_KeepOld(nI).FNScdActSeqNo = SP_FEDvChkNumeric(orsTemp, "FNScdActSeqNo")
                typeDE_KeepOld(nI).FTScdCstID = SP_FEDtChkString(orsTemp, "FTScdCstID")
                typeDE_KeepOld(nI).FTScdCardKey = SP_FEDtChkString(orsTemp, "FTScdCardKey")
                typeDE_KeepOld(nI).FTScdStaChkExpire = SP_FEDtChkString(orsTemp, "FTScdStaChkExpire")
                typeDE_KeepOld(nI).FTScdStaChkDepart = SP_FEDtChkString(orsTemp, "FTScdStaChkDepart")
                typeDE_KeepOld(nI).FNScdSptSeqNo = SP_FEDvChkNumeric(orsTemp, "FNScdSptSeqNo")
                typeDE_KeepOld(nI).FCScdMaxAlwDisc = SP_FEDvChkNumeric(orsTemp, "FCScdMaxAlwDisc")
                typeDE_KeepOld(nI).FNScdOrder = SP_FEDvChkNumeric(orsTemp, "FNScdOrder")
                typeDE_KeepOld(nI).FCScdAmtTime = SP_FEDvChkNumeric(orsTemp, "FCScdAmtTime")
                typeDE_KeepOld(nI).FCScdAmtPmt = SP_FEDvChkNumeric(orsTemp, "FCScdAmtPmt")
                typeDE_KeepOld(nI).FNScdDealGroup = SP_FEDvChkNumeric(orsTemp, "FNScdDealGroup")
                typeDE_KeepOld(nI).FCScdDiscPerOrg = SP_FEDvChkNumeric(orsTemp, "FCScdDiscPerOrg")
                typeDE_KeepOld(nI).FCScdDiscAmtOrg = SP_FEDvChkNumeric(orsTemp, "FCScdDiscAmtOrg")
                typeDE_KeepOld(nI).FCScdDiscNewOrg = SP_FEDvChkNumeric(orsTemp, "FCScdDiscNewOrg")
                typeDE_KeepOld(nI).FCScdAmt = SP_FEDvChkNumeric(orsTemp, "FCScdAmt")
                typeDE_KeepOld(nI).FCScdDiscAmt = SP_FEDvChkNumeric(orsTemp, "FCScdDiscAmt")
                typeDE_KeepOld(nI).FCScdQty = SP_FEDvChkNumeric(orsTemp, "FCScdQty")
                typeDE_KeepOld(nI).FTSdtStaPreview = SP_FEDtChkString(orsTemp, "FTScdStaPreview")
                nI = nI + 1
                orsTemp.MoveNext
            Loop
        Else
            ReDim Preserve typeDE_KeepOld(nI) As UDT_KeepCD
            typeDE_KeepOld(nI).FNSdtSeqNo = 0
            typeDE_KeepOld(nI).FTScdDiscType = ""
            typeDE_KeepOld(nI).FNDctNo = 0
            typeDE_KeepOld(nI).FTDctName = ""
            typeDE_KeepOld(nI).FTDctRefCode = ""
            typeDE_KeepOld(nI).FTScdEventID = ""
            typeDE_KeepOld(nI).FTScdCardID = ""
            typeDE_KeepOld(nI).FDScdExpDate = "01/01/2010"
            typeDE_KeepOld(nI).FTScdCardType = ""
            typeDE_KeepOld(nI).FTScdStaTrans = ""
            typeDE_KeepOld(nI).FNScdActSeqNo = 0
            typeDE_KeepOld(nI).FTScdCstID = ""
            typeDE_KeepOld(nI).FTScdCardKey = ""
            typeDE_KeepOld(nI).FTScdStaChkExpire = ""
            typeDE_KeepOld(nI).FTScdStaChkDepart = ""
            typeDE_KeepOld(nI).FNScdSptSeqNo = 0
            typeDE_KeepOld(nI).FCScdMaxAlwDisc = 0
            typeDE_KeepOld(nI).FNScdOrder = 0
            typeDE_KeepOld(nI).FCScdAmtTime = 0
            typeDE_KeepOld(nI).FCScdAmtPmt = 0
            typeDE_KeepOld(nI).FNScdDealGroup = 0
            typeDE_KeepOld(nI).FCScdDiscPerOrg = 0
            typeDE_KeepOld(nI).FCScdDiscAmtOrg = 0
            typeDE_KeepOld(nI).FCScdDiscNewOrg = 0
            typeDE_KeepOld(nI).FCScdAmt = 0
            typeDE_KeepOld(nI).FCScdDiscAmt = 0
            typeDE_KeepOld(nI).FCScdQty = 0
            typeDE_KeepOld(nI).FTSdtStaPreview = "N"
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Sub

Private Sub DE_PRCxReCalDiscount()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim nI As Integer
    Dim cSumCanDisAmt As Double
    Dim nForSend As eEN_DiscType
    'Dim cSumCanDisAmt As Double
    Dim nSeqNo As Long
    Dim tSKUCode As String
    Dim orsCD  As ADODB.Recordset
    Dim bExit As Boolean
On Error GoTo Err_Handler
    
    'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
    'ให้ระวังให้ดี Not Array จะ Return ค่า -1 (ซึ่ง Computer จะเข้าใจว่าเป็น True กับ -xxxxxxxxxxxxxxxxxx (ซึ่งคอมพิวเตอร์ไม่ได้ตีความหมายว่าเป็น True)
    'แต่ คำสั่ง IF ไม่ได้เช็คที่ค่า True หรือ False แต่เช็คที่เป็น 0 หรือไม่ใช่ 0
    'ดังนั้นจึงต้องมีการเช็คให้ เท่ากับ True หรือ -1 เสมอ นะจะบอกให้ มีผลทำให้ If (Not typeDE_KeepOld) = True Then และ If (Not typeDE_KeepOld) Then ได้ค่าที่ต่างกันนะจะบอกให้อีกที
    If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
        'ยังไม่เคยถูกเรียกใช้
    Else
        If Not (typeDE_KeepOld(nI).FNSdtSeqNo = 0) Then
            For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
                tSql = "SELECT  FTSKUCode,FNSdtSeqNo,FCSdtAmtDisAvi,FCSdtQtyAll,FCSdtSaleAmt ,FCSdtRegPrice"
                tSql = tSql & vbCrLf & ",FTSdtStaPreview,FTSdtStaNotAlwPrv"
                tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                tSql = tSql & vbCrLf & " AND FDShdTransDate='" & tUT_SaleDate & "' AND FNSdtStaRef = " & typeDE_KeepOld(nI).FNSdtSeqNo
                tSql = tSql & vbCrLf & "AND ISNULL(FTSdtCancelFlag,'0') <> '1'" '*TW 51-07-17
               ' Set orsTemp = ocnAll.Execute(tSql) '*KT 54-03-24 ป้องกัน Error
               Call SP_SQLvExecute(tSql, orsTemp, ocnAll)
                If Not (orsTemp.BOF And orsTemp.EOF) Then
                    Select Case typeDE_KeepOld(nI).FNDctNo
                        Case nEN_Disc02PmtTime
                            Do While Not orsTemp.EOF
                                oDisc.DISxClearAll
                                oDisc.dTranDate = tUT_SaleDate
                                oDisc.tTerminalNo = tVB_CNTerminalNum
                                oDisc.tTransNo = tVB_Running
                                oDisc.nDTSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                                '*KT 52-11-25  Deal&Promo
                                oDisc.nENDiscType = nEN_Disc02PmtTime
                                '*KT 53-10-30 Preview Day
                                '---------------------------------------------------------------------------------------------------------------------------
                                tVB_CNSKUFlagPrv = typeDE_KeepOld(nI).FTSdtStaPreview
                                If tVB_CNSKUFlagPrv = "" Then tVB_CNSKUFlagPrv = "N"
                                tVB_CNSKUFlagNotAlwPrvUse = SP_FEDtChkString(orsTemp, "FTSdtStaNotAlwPrv")
                                If tVB_CNSKUFlagNotAlwPrvUse = "" Then tVB_CNSKUFlagNotAlwPrvUse = "N"
                                oDisc.bIsPreviewDay = IIf(tVB_CNSKUFlagPrv = "Y", True, False)
                                If bVB_CNTrStaRegPreviewDay Then
                                    If tVB_CNSKUFlagNotAlwPrvUse = "N" And tVB_CNSKUFlagPrv = "Y" Then
                                        Call SP_INIxSwitchModePreviewDay(True)
                                    Else
                                        Call SP_INIxSwitchModePreviewDay(False)
                                    End If
                                Else
                                    Call SP_INIxSwitchModePreviewDay(False)
                                End If
                                '---------------------------------------------------------------------------------------------------------------------------
                                Select Case oDisc.C_CHKnDDPDeal(oDisc.nDTSeqNo)  'Check Double Policy
                                    Case 0 'ไม่อนุญาต
                                        'คิดเฉพาะส่วนที่เหลือ
                                        oDisc.bPromoAftDeal = True
                                    Case 1 'อนุญาต
                                        oDisc.bPromoAftDeal = False
                                    Case 2 'หาผลที่ดีที่สุด
                                        'รอดำเนินการ
                                        oDisc.bPromoAftDeal = False
                                    Case Else
                                        oDisc.bPromoAftDeal = False
                                End Select
                                Call oDisc.DIScDiscount(nEN_Disc02PmtTime)
                                orsTemp.MoveNext
                            Loop
                        
                        Case nEN_Disc03Pmt
                            Do While Not orsTemp.EOF
                                oDisc.DISxClearAll
                                oDisc.dTranDate = tUT_SaleDate
                                oDisc.tTerminalNo = tVB_CNTerminalNum
                                oDisc.tTransNo = tVB_Running
                                oDisc.nDTSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                                 '*KT 52-11-25  Deal&Promo
                                oDisc.nENDiscType = nEN_Disc03Pmt
                                '*KT 53-10-30 Preview Day
                                '---------------------------------------------------------------------------------------------------------------------------
                                tVB_CNSKUFlagPrv = typeDE_KeepOld(nI).FTSdtStaPreview
                                If tVB_CNSKUFlagPrv = "" Then tVB_CNSKUFlagPrv = "N"
                                tVB_CNSKUFlagNotAlwPrvUse = SP_FEDtChkString(orsTemp, "FTSdtStaNotAlwPrv")
                                If tVB_CNSKUFlagNotAlwPrvUse = "" Then tVB_CNSKUFlagNotAlwPrvUse = "N"
                                oDisc.bIsPreviewDay = IIf(tVB_CNSKUFlagPrv = "Y", True, False)
                                If bVB_CNTrStaRegPreviewDay Then
                                    If tVB_CNSKUFlagNotAlwPrvUse = "N" And tVB_CNSKUFlagPrv = "Y" Then
                                        Call SP_INIxSwitchModePreviewDay(True)
                                    Else
                                        Call SP_INIxSwitchModePreviewDay(False)
                                    End If
                                Else
                                    Call SP_INIxSwitchModePreviewDay(False)
                                End If
                                '---------------------------------------------------------------------------------------------------------------------------
                                Select Case oDisc.C_CHKnDDPDeal(oDisc.nDTSeqNo)  'Check Double Policy
                                    Case 0 'ไม่อนุญาต
                                        'คิดเฉพาะส่วนที่เหลือ
                                        oDisc.bPromoAftDeal = True
                                    Case 1 'อนุญาต
                                        oDisc.bPromoAftDeal = False
                                    Case 2 'หาผลที่ดีที่สุด
                                        'รอดำเนินการ
                                        oDisc.bPromoAftDeal = False
                                    Case Else
                                        oDisc.bPromoAftDeal = False
                                End Select
                                Call oDisc.DIScDiscount(nEN_Disc03Pmt)
                                orsTemp.MoveNext
                            Loop
                            
                        Case nEN_Disc04CoopDisc
                            Do While Not orsTemp.EOF
                                oDisc.DISxClearAll
                                oDisc.dTranDate = tUT_SaleDate
                                oDisc.tTerminalNo = tVB_CNTerminalNum
                                oDisc.tTransNo = tVB_Running
                                oDisc.nDTSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                                '*KT 53-10-30 Preview Day
                                '---------------------------------------------------------------------------------------------------------------------------
                                tVB_CNSKUFlagPrv = typeDE_KeepOld(nI).FTSdtStaPreview
                                If tVB_CNSKUFlagPrv = "" Then tVB_CNSKUFlagPrv = "N"
                                tVB_CNSKUFlagNotAlwPrvUse = SP_FEDtChkString(orsTemp, "FTSdtStaNotAlwPrv")
                                If tVB_CNSKUFlagNotAlwPrvUse = "" Then tVB_CNSKUFlagNotAlwPrvUse = "N"
                                oDisc.bIsPreviewDay = IIf(tVB_CNSKUFlagPrv = "Y", True, False)
                                If bVB_CNTrStaRegPreviewDay Then
                                    If tVB_CNSKUFlagNotAlwPrvUse = "N" And tVB_CNSKUFlagPrv = "Y" Then
                                        Call SP_INIxSwitchModePreviewDay(True)
                                    Else
                                        Call SP_INIxSwitchModePreviewDay(False)
                                    End If
                                Else
                                    Call SP_INIxSwitchModePreviewDay(False)
                                End If
                                '---------------------------------------------------------------------------------------------------------------------------
                                Call oDisc.DIScDiscount(nEN_Disc04CoopDisc)
                                orsTemp.MoveNext
                            Loop
                        
                        Case nEN_Disc05LnDisc, nEN_Disc06LnMDown
    
                            Select Case typeDE_KeepOld(nI).FTScdDiscType
                                Case nEN_LnDis1Per
                                    Do While Not orsTemp.EOF
                                        oDisc.DISxClearAll
                                        oDisc.dTranDate = tUT_SaleDate
                                        oDisc.tTerminalNo = tVB_CNTerminalNum
                                        oDisc.tTransNo = tVB_Running
                                        oDisc.cDiscValue = typeDE_KeepOld(nI).FCScdDiscPerOrg
                                        oDisc.nDTSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                                        oDisc.nENLnDisType = typeDE_KeepOld(nI).FTScdDiscType
                                        '*KT 53-10-30 Preview Day
                                        '---------------------------------------------------------------------------------------------------------------------------
                                        tVB_CNSKUFlagPrv = typeDE_KeepOld(nI).FTSdtStaPreview
                                        If tVB_CNSKUFlagPrv = "" Then tVB_CNSKUFlagPrv = "N"
                                        tVB_CNSKUFlagNotAlwPrvUse = SP_FEDtChkString(orsTemp, "FTSdtStaNotAlwPrv")
                                        If tVB_CNSKUFlagNotAlwPrvUse = "" Then tVB_CNSKUFlagNotAlwPrvUse = "N"
                                        oDisc.bIsPreviewDay = IIf(tVB_CNSKUFlagPrv = "Y", True, False)
                                        If bVB_CNTrStaRegPreviewDay Then
                                            If tVB_CNSKUFlagNotAlwPrvUse = "N" And tVB_CNSKUFlagPrv = "Y" Then
                                                Call SP_INIxSwitchModePreviewDay(True)
                                            Else
                                                Call SP_INIxSwitchModePreviewDay(False)
                                            End If
                                        Else
                                            Call SP_INIxSwitchModePreviewDay(False)
                                        End If
                                        '---------------------------------------------------------------------------------------------------------------------------
                                        Call oDisc.DIScDiscount(typeDE_KeepOld(nI).FNDctNo)
                                        orsTemp.MoveNext
                                    Loop
                                Case nEN_LnDis2Amt
                                    orsTemp.MoveFirst
                                    Do While Not orsTemp.EOF
                                        Select Case DE_CHKnDDP(SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo"), typeDE_KeepOld(nI).FNDctNo)
                                            Case 0 'ไม่อนุญาต
                                                cSumCanDisAmt = cSumCanDisAmt + SP_FEDvChkNumeric(orsTemp, "FCSdtAmtDisAvi") / SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll")
                                            Case 1 'อนุญาต
                                                cSumCanDisAmt = cSumCanDisAmt + SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt") / SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll")
                                        End Select
                                        orsTemp.MoveNext
                                    Loop
                                    If cSumCanDisAmt > 0 Then
                                        orsTemp.MoveFirst
                                        Do While Not orsTemp.EOF
                                            oDisc.DISxClearAll
                                            oDisc.dTranDate = tUT_SaleDate
                                            oDisc.tTerminalNo = tVB_CNTerminalNum
                                            oDisc.tTransNo = tVB_Running
                                            Select Case DE_CHKnDDP(SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo"), typeDE_KeepOld(nI).FNDctNo)
                                                Case 0 'ไม่อนุญาต
                                                    If typeDE_KeepOld(nI).FCScdDiscAmtOrg < cSumCanDisAmt Then
                                                        oDisc.cDiscValue = typeDE_KeepOld(nI).FCScdDiscAmtOrg * SP_FEDvChkNumeric(orsTemp, "FCSdtAmtDisAvi") / SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll") / cSumCanDisAmt
                                                    Else
                                                        oDisc.cDiscValue = SP_FEDvChkNumeric(orsTemp, "FCSdtAmtDisAvi") / SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll")
                                                    End If
                                                Case 1 'อนุญาต
                                                    If typeDE_KeepOld(nI).FCScdDiscAmtOrg < cSumCanDisAmt Then
                                                        oDisc.cDiscValue = typeDE_KeepOld(nI).FCScdDiscAmtOrg * SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt") / SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll") / cSumCanDisAmt
                                                    Else
                                                        oDisc.cDiscValue = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt") / SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll")
                                                    End If
                                            End Select
                                            oDisc.nDTSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                                            oDisc.nENLnDisType = nVB_TypeDiscLineItem
                                            If oDisc.cDiscValue > 0 Then
                                                '*KT 53-10-30 Preview Day
                                                '---------------------------------------------------------------------------------------------------------------------------
                                                tVB_CNSKUFlagPrv = typeDE_KeepOld(nI).FTSdtStaPreview
                                                If tVB_CNSKUFlagPrv = "" Then tVB_CNSKUFlagPrv = "N"
                                                tVB_CNSKUFlagNotAlwPrvUse = SP_FEDtChkString(orsTemp, "FTSdtStaNotAlwPrv")
                                                If tVB_CNSKUFlagNotAlwPrvUse = "" Then tVB_CNSKUFlagNotAlwPrvUse = "N"
                                                oDisc.bIsPreviewDay = IIf(tVB_CNSKUFlagPrv = "Y", True, False)
                                                If bVB_CNTrStaRegPreviewDay Then
                                                    If tVB_CNSKUFlagNotAlwPrvUse = "N" And tVB_CNSKUFlagPrv = "Y" Then
                                                        Call SP_INIxSwitchModePreviewDay(True)
                                                    Else
                                                        Call SP_INIxSwitchModePreviewDay(False)
                                                    End If
                                                Else
                                                    Call SP_INIxSwitchModePreviewDay(False)
                                                End If
                                                '---------------------------------------------------------------------------------------------------------------------------
                                                Call oDisc.DIScDiscount(typeDE_KeepOld(nI).FNDctNo)
                                            End If
                                            orsTemp.MoveNext
                                        Loop
                                    End If
                                Case nEN_LnDis3NewPrice
                                    orsTemp.MoveFirst
                                    Do While Not orsTemp.EOF
                                        oDisc.DISxClearAll
                                        oDisc.dTranDate = tUT_SaleDate
                                        oDisc.tTerminalNo = tVB_CNTerminalNum
                                        oDisc.tTransNo = tVB_Running
                                        oDisc.cDiscValue = typeDE_KeepOld(nI).FCScdDiscNewOrg
                                        oDisc.nDTSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                                        oDisc.nENLnDisType = nVB_TypeDiscLineItem
                                        oDisc.nENLnDisType = nEN_LnDis3NewPrice
                                        If oDisc.cDiscValue < SP_FEDvChkNumeric(orsTemp, "FCSdtAmtDisAvi") / SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll") Then
                                            '*KT 53-10-30 Preview Day
                                            '---------------------------------------------------------------------------------------------------------------------------
                                            tVB_CNSKUFlagPrv = typeDE_KeepOld(nI).FTSdtStaPreview
                                            If tVB_CNSKUFlagPrv = "" Then tVB_CNSKUFlagPrv = "N"
                                            tVB_CNSKUFlagNotAlwPrvUse = SP_FEDtChkString(orsTemp, "FTSdtStaNotAlwPrv")
                                            If tVB_CNSKUFlagNotAlwPrvUse = "" Then tVB_CNSKUFlagNotAlwPrvUse = "N"
                                            oDisc.bIsPreviewDay = IIf(tVB_CNSKUFlagPrv = "Y", True, False)
                                            If bVB_CNTrStaRegPreviewDay Then
                                                If tVB_CNSKUFlagNotAlwPrvUse = "N" And tVB_CNSKUFlagPrv = "Y" Then
                                                    Call SP_INIxSwitchModePreviewDay(True)
                                                Else
                                                    Call SP_INIxSwitchModePreviewDay(False)
                                                End If
                                            Else
                                                Call SP_INIxSwitchModePreviewDay(False)
                                            End If
                                            '---------------------------------------------------------------------------------------------------------------------------
                                            Call oDisc.DIScDiscount(typeDE_KeepOld(nI).FNDctNo)
                                        End If
                                        orsTemp.MoveNext
                                    Loop
                            End Select
                        Case nEN_Disc21MPrice, nEN_Disc22MDisc
                            orsTemp.MoveFirst
                            Do While Not orsTemp.EOF
                                xVB_CNMemberP.ID = typeDE_KeepOld(nI).FTScdCardID
                                xVB_CNMemberP.MallCardType = typeDE_KeepOld(nI).FTScdCardType
                                tVB_StaMemPriceAlw = "Y"
                                tVB_StaMemPricePromoAlw = False
                                nVB_CNMemPriceMaxQty = 99999999
                                tVB_SKUConsignFlag = IIf(typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc, "Y", "N")
                                tVB_CNMemPriceType = typeDE_KeepOld(nI).FTScdDiscType
                                bExit = False
                                Select Case tVB_CNMemPriceType
                                    Case "1" '% Off
                                         cVB_CNMemPriceAmt = typeDE_KeepOld(nI).FCScdDiscPerOrg
                                    Case "2" 'Amount Off
                                        cVB_CNMemPriceAmt = typeDE_KeepOld(nI).FCScdDiscAmtOrg
                                    Case "3" 'New Price
                                        cVB_CNMemPriceAmt = typeDE_KeepOld(nI).FCScdDiscNewOrg
                                    Case "4" 'No Promotion
                                        'none
                                        bExit = True
                                    Case Else
                                        bExit = True
                                End Select
                                If bExit = False Then
                                    cUT_Qty = SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll")
                                    tVB_CNEventMemPrice = typeDE_KeepOld(nI).FTScdEventID
        '*KT 53-01-08 Deal&Member Price
        '                            If typeDE_KeepOld(nI).FCScdQty = 1 Then
        '                                cVB_CNMemPriceAmtDis = typeDE_KeepOld(nI).FCScdAmt * cUT_Qty
        '                            Else
        '                                cVB_CNMemPriceAmtDis = typeDE_KeepOld(nI).FCScdAmt
        '                            End If
                                    cVB_CNMemPriceAmtDis = CDbl((typeDE_KeepOld(nI).FCScdAmt / typeDE_KeepOld(nI).FCScdQty) * cUT_Qty)
                                    cUT_RegPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")
                                    '*KT 53-10-30 Preview Day
                                    '---------------------------------------------------------------------------------------------------------------------------
                                    tVB_CNSKUFlagPrv = typeDE_KeepOld(nI).FTSdtStaPreview
                                    If tVB_CNSKUFlagPrv = "" Then tVB_CNSKUFlagPrv = "N"
                                    tVB_CNSKUFlagNotAlwPrvUse = SP_FEDtChkString(orsTemp, "FTSdtStaNotAlwPrv")
                                    If tVB_CNSKUFlagNotAlwPrvUse = "" Then tVB_CNSKUFlagNotAlwPrvUse = "N"
                                    oDisc.bIsPreviewDay = IIf(tVB_CNSKUFlagPrv = "Y", True, False)
                                    If bVB_CNTrStaRegPreviewDay Then
                                        If tVB_CNSKUFlagNotAlwPrvUse = "N" And tVB_CNSKUFlagPrv = "Y" Then
                                            Call SP_INIxSwitchModePreviewDay(True)
                                        Else
                                            Call SP_INIxSwitchModePreviewDay(False)
                                        End If
                                    Else
                                        Call SP_INIxSwitchModePreviewDay(False)
                                    End If
                                    '---------------------------------------------------------------------------------------------------------------------------
                                    wFunctionKB.W_PRCxMemberPrice SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo"), SP_FEDtChkString(orsTemp, "FTSKUCode"), 0, 0, True
                                    cUT_Qty = 1     '*Ao 2011/03/31  CommSheet ML-PRV-054-SF Wrong qty on last item.xls
                                End If
                                orsTemp.MoveNext
                            Loop
                    End Select
                End If
            Next nI
        End If
    End If
    cUT_Qty = 1         '*Ao 2011/03/31  CommSheet ML-PRV-054-SF Wrong qty on last item.xls
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsCD)
    Exit Sub
Err_Handler:
    cUT_Qty = 1     '*Ao 2011/03/31  CommSheet ML-PRV-054-SF Wrong qty on last item.xls
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsCD)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DE_PRCxCalDeal(ByVal pEN_DealType As eEN_DealTypeRec)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsQtyFrame() As ADODB.Recordset
    'Dim dStart As Date
    Dim nI As Long
    Dim tSeqOrder As String
    Dim tSeqSku As String
    Dim tQtySku As String
    Dim tStaSku As String
    Dim tDealNo As String
    Dim tDocType As String
    Dim nOrgGet As Integer
    Dim nGrp As Integer
    Dim nGrpOld As Integer
    Dim nGet As Integer
    Dim nQtyDisc As Integer
    Dim cAmtDisc As Double
    Dim cSumAmt As Double
    On Error GoTo Err_Handler
    Select Case pEN_DealType
        Case nEN_Deal01Single
            Call DE_PRCxCalSingle
        Case nEN_Deal02Multiple
            'dStart = Now
            If DE_SETbOrsAllDeal Then 'สร้าง Recordset ขึ้นมาเพื่อเก็บลำดับการลด Deal ทั้งหมดที่เกิดขึ้นได้
                ReDim Preserve orsQtyFrame(0) As ADODB.Recordset
                'If SP_RSTbClone2Offline(orsDE_DealMulti, orsTemp, True) Then
                Call DE_ARYxGet4RST(orsDE_DealMulti, atDE_AllDeal, acDE_AllDisc, acDE_AllRegPrice, atDE_AllType, atDE_AllGrp)
'                Call DE_BAKxors2Array(orsDE_DealMulti, typeDE_KeepDeal) '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls
                'Set wDealTest.VSFlexGrid2.DataSource = orsDE_DealMulti '*TW XXX
                If SP_RSTbClone2OfflineEx2(orsDE_AFSack, orsQtyFrame(0), True) Then
                    nDE_MaxElement = UBound(atDE_AllDeal)
                    Call DE_PRCxRecursive(1, "", orsQtyFrame) 'Call Recursive Sub For find all posible Deal
                    If Not (orsDE_AllDeal.BOF And orsDE_AllDeal.EOF) Then
'                        orsDE_AllDeal.Sort = "FCDeaSum DESC" 'Find The bast
                        orsDE_AllDeal.Sort = "FCDeaSum DESC,FCDeaSumBlue" 'Find The bast           ส่วนลดที่มากที่สุด,ป้ายน้ำเงินน้อยที่สุด

                        tSeqOrder = orsDE_AllDeal.Fields("FTDeaSeq").Value
                        tSeqSku = orsDE_AllDeal.Fields("FTDeaSkuSeq").Value
                        tQtySku = orsDE_AllDeal.Fields("FTDeaSkuQty").Value
                        tStaSku = orsDE_AllDeal.Fields("FTDeaSkuSta").Value
                        
                        Call DE_PRCxGetDeal(tSeqOrder, tSeqSku, tQtySku, tStaSku)

                    End If
                    'MsgBox "Finish : " & DateDiff("s", dStart, Now)
                End If

                For nI = LBound(orsQtyFrame) To UBound(orsQtyFrame)
                    Call SP_DATxRsClose(orsQtyFrame(nI))
                Next nI
            End If
    End Select
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function DE_PRCbCase003(ByVal porsTemp As ADODB.Recordset, ByVal ptDealType As String, ByRef pnRound As Integer) As Boolean
'___________________________________
'   Call:
'   Cmt:    '*Ao 09-10-16
'___________________________________
Dim nLoop As Integer, nMaxGrp As Integer, nPos As Integer, nRound As Integer, nRoundTmp As Integer, nQtyBuy As Integer, nQtyGet As Integer
Dim tSKU As String, tFilter As String
Dim bChkDeal  As Boolean

On Error GoTo Err_Handler

nMaxGrp = 2
nRound = 0
bChkDeal = False

    With porsTemp
         If Not (.BOF And .EOF) Then
            For nLoop = 1 To nMaxGrp
                .Filter = adFilterNone
                If nLoop = 1 Then
                    tFilter = "FNDacQtyGet = 0"             'BUY
                Else
                    tFilter = "FNDacQtyBuy = 0"             'GET
                End If
                .Filter = tFilter
                 If nLoop = 1 Then
                    nQtyBuy = .Fields("FNDacQtyBuy").Value
                Else
                    nQtyGet = .Fields("FNDacQtyGet").Value
                End If
            Next nLoop
            .Filter = adFilterNone
            .MoveFirst
            Do While Not .EOF
                tSKU = .Fields("FTDacSku").Value
                nPos = .AbsolutePosition
                tFilter = "FTDacSku ='" & tSKU & "'"
                .Filter = tFilter
                If .RecordCount > 1 Then
                    nRoundTmp = .Fields("FCSdtQtyAll").Value \ (nQtyBuy + nQtyGet)
                    bChkDeal = True
                    If nRoundTmp < nRound Then
                        nRound = nRoundTmp
                    Else
                        If nRound = 0 Then
                            nRound = nRoundTmp
                        End If
                    End If
                End If
                .Filter = adFilterNone
                .AbsolutePosition = nPos
                .MoveNext
            Loop
            
         End If
    End With
    If bChkDeal Then
        DE_PRCbCase003 = True
    Else
        DE_PRCbCase003 = False
    End If
    pnRound = nRound
 Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_PRCnChkRoundDeal003(ByVal porsTemp As ADODB.Recordset, ByVal ptDealType As String) As Integer
'___________________________________
'   Call:
'   Cmt:    '*Ao 2011/09/02   แก้ไขให้สามารถซื้อสินค้าซ้ำกันได้ 'CommSheet ML-PRV-066  Not Apply Deal Type 3 with same SKU.xls
'___________________________________
Dim nCountRound As Integer, nLoop As Integer, nMaxGrp As Integer, nRoundBuy As Integer, nRoundGet As Integer
Dim nQtyAll As Integer, nQtyBuy As Integer, nQtyGet As Integer
Dim tFilter As String, tSort As String, tSkuGrp As String, tSkuGrpOld As String
On Error GoTo Err_Handler

    nRoundBuy = 0
    nRoundGet = 0
    nQtyAll = 0
'    nMaxGrp = 2
    tSkuGrpOld = "XXX"
    
    With porsTemp
        If Not (.BOF And .EOF) Then
            Select Case ptDealType
                Case "003"
'                    If DE_PRCbCase003(porsTemp, ptDealType, nCountRound) Then
                    .Filter = adFilterNone
                    .MoveFirst
                    Do While Not .EOF
                        tSkuGrp = .Fields("FNDacSkuGrp").Value
                        If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
                            If tSkuGrpOld <> "XXX" Then
                                If tSkuGrpOld = "1" Then
                                    nRoundBuy = nQtyAll \ nQtyBuy
                                Else
                                    nRoundGet = nQtyAll \ nQtyBuy
                                End If
                            End If
                            nQtyAll = 0
                            If tSkuGrp = "1" Then
                                nQtyBuy = .Fields("FNDacQtyBuy").Value
                            Else
                                nQtyBuy = .Fields("FNDacQtyGet").Value
                            End If
                            nQtyAll = nQtyAll + .Fields("FCSdtQtyAll").Value
                            tSkuGrpOld = tSkuGrp
                        Else
                            nQtyAll = nQtyAll + .Fields("FCSdtQtyAll").Value
                        End If
                        .MoveNext
                    Loop
                    If tSkuGrpOld = "1" Then
                        nRoundBuy = nQtyAll \ nQtyBuy
                    Else
                        nRoundGet = nQtyAll \ nQtyBuy
                    End If
                    If nRoundBuy > nRoundGet Then
                        nCountRound = nRoundGet
                    Else
                        nCountRound = nRoundBuy
                    End If
            End Select
        End If
    End With
    
    DE_PRCnChkRoundDeal003 = nCountRound
    
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_PRCnChkRoundFree(ByVal porsTemp As ADODB.Recordset, ByVal ptDealType As String) As Integer
'___________________________________
'   Call:
'   Cmt:    '*Ao 09-10-16
'___________________________________
Dim nCountRound As Integer
Dim nQtyAll As Integer, nQtyBuy As Integer, nQtyGet As Integer
On Error GoTo Err_Handler
    nCountRound = 0
    nQtyAll = 0
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
    
        If ptDealType = "0010" Then
            nQtyAll = nQtyAll + porsTemp.Fields("FCSdtQtyAll").Value
            nQtyBuy = porsTemp.Fields("FNDacQtyBuy").Value
            nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
        End If
        porsTemp.MoveNext
        
    Loop
    nCountRound = nQtyAll \ (nQtyBuy + nQtyGet)
    DE_PRCnChkRoundFree = nCountRound
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function
'Private Function DE_PRCnChkRound(ByVal porsSource As ADODB.Recordset, ByVal ptDealNo As String, ByVal ptDealType As String) As Long
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim orsTemp As ADODB.Recordset
'    Dim bDealComplete As Boolean
'    Dim nAllowCount As Integer, nRound As Integer
'    Dim nCountRound As Long
'
'    On Error GoTo Err_Handler
'    '*Ao 52-03-25
'    Call SP_RSTbClone2Offline(porsSource, orsTemp)
''    Set orsTemp = porsSource.Clone                                                 'ไม่สามารถใช้  .Clone ได้
'    '____________________________________________
'    bDealComplete = True
'    nCountRound = 0
'
'    Do While bDealComplete
'        bDealComplete = DE_CHKbDealCompleteAndCut(orsTemp, ptDealNo, ptDealType, , nRound)
'        Select Case ptDealType
'            Case "001", "003", "0010"
'                nCountRound = nRound
'                bDealComplete = False
'            Case Else
'                If bDealComplete Then
'                        nCountRound = nCountRound + 1
'                End If
'        End Select
'    Loop
'
'    DE_PRCnChkRound = nCountRound
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_PRCnChkRound(ByVal porsSource As ADODB.Recordset, ByVal ptDealNo As String, ByVal ptDealType As String) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 09-11-20
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsTemp As ADODB.Recordset
    Dim bDealComplete As Boolean
    Dim nAllowCount As Integer, nRound As Integer
    Dim nCountRound As Long
    
    On Error GoTo Err_Handler
    
    Call SP_RSTbClone2OfflineEx2(porsSource, orsTemp)
    bDealComplete = True
    nCountRound = 0
   
    Do While bDealComplete
        bDealComplete = DE_CHKbDealCompleteAndCut(orsTemp, ptDealNo, ptDealType, , nRound)
        Select Case ptDealType
            Case "001", "002", "003", "004", "0010", "0020", "0030", "0040"
                If bDealComplete Then
                    nCountRound = 1
                    bDealComplete = False
                Else
                    nCountRound = nRound
                End If
                
            Case Else
            
        End Select
    Loop
    
    DE_PRCnChkRound = nCountRound
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_PRCxCalSingle()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsTemp As ADODB.Recordset
    
    Dim nQtyGet As Integer, nQtyCal As Integer, nQtyGetOrg As Integer
    
    Dim nDTSeqNo As Integer
    '*TW For Write to CD
    Dim tDealNo As String, tFreeFlag As String, tFreeType As String, tDealType As String, tEventID As String
    Dim cDealDiscPer As Double
    Dim cDealDiscAmt As Double
    Dim tLnDiscType As String
    
    Dim nGrp As Integer, nGrpOld As Integer, nUncalGrp As Integer
    Dim nMaxGrp As Integer
    Dim bDealComplete As Boolean, bChkDeal As Boolean
    Dim nCountRound As Long
    Dim pnAllowCount As Integer
    
    On Error GoTo Err_Handler
'    Call DE_BAKxors2Array(orsDE_DealSingle, typeDE_KeepDeal) '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls
    orsDE_DealSingle.MoveFirst
'    Set wDealTest.VSFlexGrid1.DataSource = orsDE_DealSingle
    Do While Not orsDE_DealSingle.EOF
        tLnDiscType = orsDE_DealSingle.Fields("FTPddDisType").Value
        tDealNo = orsDE_DealSingle.Fields("FTPdhDealNo").Value
        tDealType = orsDE_DealSingle.Fields("FTDacDealType").Value
        tFreeType = orsDE_DealSingle.Fields("FTDacDealType").Value
        tEventID = orsDE_DealSingle.Fields("FTEvnCode").Value
        
        orsDE_AFSack.Filter = "FTPdhDealNo = '" & tDealNo & "'"
        orsDE_AFSack.Sort = "FNDacSkuGrp,FCSdtRegPrice DESC"       '*Ao 52-05-12
        Select Case tDealType
            Case "001", "002", "003", "004", "0010", "0020", "0030", "0040"
                nCountRound = DE_PRCnChkRound(orsDE_AFSack, tDealNo, tDealType)
            Case "005", "006"
                nCountRound = 1
        End Select
        Select Case tDealType
            Case "0010", "0020", "0030", "0040"
                tLnDiscType = "4"
        End Select
        Call SP_RSTbClone2OfflineEx2(orsDE_AFSack, orsTemp)
        bDealComplete = DE_CHKbDealComplete(orsTemp, tDealNo, tDealType)
        Do While bDealComplete
            nDE_Group = nDE_Group + 1 '*TW 51-08-11
            bDE_MemberPriceAct = False
            Select Case tDealType
                Case "001"
'                    Call DE_SINxCal001(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                    bChkDeal = DE_SINbCal001(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "002"
'                    Call DE_SINxCal002(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp)
                    bChkDeal = DE_SINbCal002(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "003"
                    bChkDeal = DE_SINbCal003(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "004"
                    bChkDeal = DE_SINbCal004(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "005"
'                    Call DE_SINxCal005(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp)
                    bChkDeal = DE_SINbCal005(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "006"
                    bChkDeal = DE_SINbCal006(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "007"
                    bChkDeal = DE_SINbCal007(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "008"
                    bChkDeal = DE_SINbCal008(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "0010"
                    bChkDeal = DE_SINbCal0010(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "0020"
                    bChkDeal = DE_SINbCal0020(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "0030"
                    bChkDeal = DE_SINbCal0030(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
                Case "0040"
                    bChkDeal = DE_SINbCal0040(tDealNo, tDealType, tFreeType, tEventID, tLnDiscType, orsTemp, nCountRound)
            End Select
            
            If bDE_MemberPriceAct Then
                bDealComplete = False
            Else
                Select Case tDealType
                    Case "001", "002", "004", "005", "006"
                        If bChkDeal = True Then
                            bDealComplete = DE_CHKbDealComplete(orsTemp, tDealNo, tDealType)
                        Else
                            bDealComplete = False
                        End If
                    Case "003", "007", "008", "0010", "0020", "0030", "0040"
                        bDealComplete = False
                    Case Else
                        bDealComplete = DE_CHKbDealComplete(orsTemp, tDealNo, tDealType)
                End Select
            End If
        Loop
        orsDE_DealSingle.MoveNext
    Loop
    oDisc.C_UPDxLabel '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DE_TESxCheckDeal()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Dim orsDealNo As ADODB.Recordset

On Error GoTo Err_Handler
    'หา Deal ที่มีโอกาสเกิดขึ้นได้จริง
    'tVB_CNWheBill
    
    tDE_Whe = " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tDE_Whe = tDE_Whe & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tDE_Whe = tDE_Whe & vbCrLf & " AND FDShdTransDate = '" & Format(tUT_SaleDate, "YYYY/MM/DD") & "'"
    tDE_Whe = tDE_Whe & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"

    Call DE_GETxAllDeal(orsDealNo, tVB_CNTblDT, tVB_CNTerminalNum, tVB_Running)
    
    'Set wDealTest.VSFlexGrid3.DataSource = orsDealNo '*TW XXX
    If Not (orsDealNo.BOF And orsDealNo.EOF) Then
        'Call SP_RSTbClone2Offline(orsDealNo, orsTmp) 'เอา Deal ที่อาจจะเกิดขึ้นใน บิลนี้
        If DE_FILbSackNoDeal(orsDealNo, orsDE_AFSack) Then 'เอามาเฉพาะ Deal ที่สามารถเกิดได้จริงจากบิลนี้ (องค์ประกอบไม่ครบ เอาออก)
            Call DE_FILxSepDeal(orsDE_AFSack, orsDE_DealSingle, orsDE_DealMulti) 'แยก deal ที่เป็น Multi กับ Single ออก
            'ทำส่วนที่เป็น Multi ก่อน
            Call DE_CHKxModeMemberPrice '*KT 52-02-19 member price
            Call DE_PRCxStoreDiscount
            Call DE_PRCxClearDiscount
            If Not (orsDE_DealMulti.BOF And orsDE_DealMulti.EOF) Then
                'ตัด Deal ที่ซ้ำกันออก TW 51-08-08
                    Call DE_PRCxCompress(orsDE_DealMulti)
                    
                '--------------------------
                orsDE_DealMulti.Sort = "FCDacDis DESC,FCSdtRegPrice DESC"
                
                Call DE_PRCxCalDeal(nEN_Deal02Multiple)
            End If
            
             If Not (orsDE_DealSingle.BOF And orsDE_DealSingle.EOF) Then
                Call DE_PRCxCalDeal(nEN_Deal01Single)
            End If
            Call DE_PRCxReCalDT
            Call DE_PRCxReCalDiscount
'            Call DE_CHKxDelFreeTypeVendor           '*Ao 2011/01/15
            Call DE_CHKxFreeTypeVendor           '*Ao 2011/01/15
'            bVB_CNbHistoryHaveDael = True
        End If
    End If
    Call SP_DATxRsNothing(orsDealNo)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsDealNo)
End Sub

Private Sub DE_PRCxRecursive(ByVal pnPos As Integer, ByVal ptLast As String, ByRef porsQtyFrame() As ADODB.Recordset)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nI As Integer
    Dim nTime As Long
    On Error GoTo Err_Handler
    porsQtyFrame(pnPos - 1).Filter = adFilterNone
    For nI = 1 To nDE_MaxElement
        ReDim Preserve porsQtyFrame(pnPos) As ADODB.Recordset
        ReDim Preserve acDE_SumDisc(pnPos)
        ReDim Preserve atDE_SkuSeqUsed(pnPos) As String
        ReDim Preserve atDE_SkuQtyUsed(pnPos) As String
        ReDim Preserve atDE_SkuStaUsed(pnPos) As String
        '*Ao 52-04-27
        ReDim Preserve acDE_SumBlue(pnPos)
'        ReDim Preserve acDE_Round(pnPos)
        '_________________________________
        atDE_SkuSeqUsed(pnPos) = ""
        atDE_SkuQtyUsed(pnPos) = ""
        atDE_SkuStaUsed(pnPos) = ""
        
        Call SP_RSTbClone2OfflineEx2(porsQtyFrame(pnPos - 1), porsQtyFrame(pnPos), True)
        If InStr(1, ptLast, nI) <= 0 Then 'ซ้ำไม่ทำเลขเดิม
            If DE_PRCbDiscount(pnPos, nI, acDE_SumDisc(pnPos - 1), acDE_SumBlue(pnPos - 1), porsQtyFrame, nTime) Then              'เพื่อเป็นการประหยัดเวลา ให้หาค่าก่อน ถ้าตัว Deal ปัจจุบัน หาค่าไม่ได้ ให้ออกทันที ไม่ต้องเข้า Recursive ต่อ
                If pnPos <> nDE_MaxElement Then
                    Call DE_PRCxRecursive(pnPos + 1, ptLast & DE_GETtLoopString(nTime, "#" & nI), porsQtyFrame)
                Else
                    '*Ao 52-03-17
                    If acDE_SumDisc(pnPos) > 0 Then
                        Call DE_PRCxSaveSummary(pnPos + 1, ptLast & DE_GETtLoopString(nTime, "#" & nI), atDE_SkuSeqUsed, atDE_SkuQtyUsed, atDE_SkuStaUsed)
                    Else
                        atDE_SkuSeqUsed(pnPos) = ""
                        atDE_SkuQtyUsed(pnPos) = ""
                        atDE_SkuStaUsed(pnPos) = ""
                    End If
'_________________________________________________
'                    If DE_PRCbDiscount(pnPos, nI, acDE_SumDisc(pnPos - 1), porsQtyFrame, nTime) Then
'                        Call DE_PRCxSaveSummary(pnPos + 1, ptLast & DE_GETtLoopString(nTime, "#" & nI), atDE_SkuSeqUsed, atDE_SkuQtyUsed, atDE_SkuStaUsed)
'                    Else
'                        atDE_SkuSeqUsed(pnPos) = ""
'                        atDE_SkuQtyUsed(pnPos) = ""
'                        atDE_SkuStaUsed(pnPos) = ""
'                    End If
                End If
            Else
                atDE_SkuSeqUsed(pnPos) = ""
                atDE_SkuQtyUsed(pnPos) = ""
                atDE_SkuStaUsed(pnPos) = ""
            End If
        End If
    Next nI
    
    Call DE_PRCxSaveSummary(pnPos, ptLast, atDE_SkuSeqUsed, atDE_SkuQtyUsed, atDE_SkuStaUsed)
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function DE_GETtLoopString(ByVal pnTime As Integer, ByVal ptSource As String) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nI As Integer
    On Error GoTo Err_Handler
    For nI = 1 To pnTime
        DE_GETtLoopString = DE_GETtLoopString & ptSource
    Next nI
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_PRCxSaveSummary(ByVal pnPos As Long, ByVal ptLast As String, ByRef patDE_SkuSeqUsed() As String, ByRef patDE_SkuQtyUsed() As String, ByRef patStaUsed() As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nI As Integer
    Dim tTempSeq As String
    Dim tTempQty As String
    Dim tTempSta As String
    On Error GoTo Err_Handler
    For nI = LBound(patDE_SkuSeqUsed) + 1 To UBound(patDE_SkuSeqUsed)
        tTempSeq = tTempSeq & ";" & patDE_SkuSeqUsed(nI)
        tTempQty = tTempQty & ";" & patDE_SkuQtyUsed(nI)
        tTempSta = tTempSta & ";" & patStaUsed(nI)
    Next nI
    If Left(tTempSeq, 1) = ";" Then
        tTempSeq = Right(tTempSeq, Len(tTempSeq) - 1)
        tTempQty = Right(tTempQty, Len(tTempQty) - 1)
        tTempSta = Right(tTempSta, Len(tTempSta) - 1)
    End If
    orsDE_AllDeal.AddNew
    orsDE_AllDeal.Fields("Test").Value = orsDE_AllDeal.RecordCount
    orsDE_AllDeal.Fields("FTDeaSeq").Value = ptLast
    orsDE_AllDeal.Fields("FCDeaSum").Value = acDE_SumDisc(pnPos - 1)
    orsDE_AllDeal.Fields("FTDeaSkuSeq").Value = tTempSeq
    orsDE_AllDeal.Fields("FTDeaSkuQty").Value = tTempQty
    orsDE_AllDeal.Fields("FTDeaSkuSta").Value = tTempSta
    '*Ao 52-04-27
    orsDE_AllDeal.Fields("FCDeaSumBlue").Value = acDE_SumBlue(pnPos - 1)
'    orsDE_AllDeal.Fields("FCDeaRound").Value
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function DE_PRCbDiscount(ByVal pnPos As Integer, ByVal pnKey As Long, ByVal pcOldDisc As Double, ByVal pcOldBlue As Double, ByRef porsQtyFrame() As ADODB.Recordset, ByRef pnTime As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    'pnPos :ลำดับว่าเป็นลำดับที่ 1 หรือ 2 หรือ 3 เอาไว้กำกับ Frame และ SumDis
    'pnKey : Key ที่เก็บ Index
    'pcOldDisc : รวมส่วนลดก่อนหน้า
    'porsQtyFrame เก็บจำนวนปัจจุบันที่เหลือ
    Dim cSource As Double 'เก็บจำนวนซื้อ
    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
    Dim nMaxGrp As Integer
    Dim nI As Integer
    Dim tFilter As String
    Dim tSort As String
    Dim cDisSumDisc As Double
    Dim cDisSumPrice As Double
    Dim bDealComplete As Boolean
    On Error GoTo Err_Handler
    
    '*Ao 09-11-20
    acDE_SumBlue(pnPos) = 0
'    acDE_Round(pnPos) = 1
    '_______________________________
    pnTime = 0
    acDE_SumDisc(pnPos) = 0
    atDE_SkuSeqUsed(pnPos) = ""
    atDE_SkuQtyUsed(pnPos) = ""
    atDE_SkuStaUsed(pnPos) = ""
    'atDE_AllDeal , acDE_AllDisc, acDE_AllRegPrice, atDE_AllType, atDE_AllGrp
    With porsQtyFrame(pnPos)
        Do
            pnTime = pnTime + 1
            bDealComplete = DE_CHKbDealComplete(porsQtyFrame(pnPos), atDE_AllDeal(pnKey), atDE_AllType(pnKey))
            If bDealComplete Then
                Select Case atDE_AllType(pnKey)
                    Case "001"
                        DE_PRCbDiscount = DE_DISb001(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "002"
                        DE_PRCbDiscount = DE_DISb002(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "003"
                        DE_PRCbDiscount = DE_DISb003(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)) _
                        , IIf(acDE_SumBlue(pnPos) > 0, acDE_SumBlue(pnPos), pcOldBlue + acDE_SumBlue(pnPos)))
'                        DE_PRCbDiscount = DE_DISb003(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "004"
                        DE_PRCbDiscount = DE_DISb004(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)) _
                        , IIf(acDE_SumBlue(pnPos) > 0, acDE_SumBlue(pnPos), pcOldBlue + acDE_SumBlue(pnPos)))
'                        DE_PRCbDiscount = DE_DISb004(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "005"
                        DE_PRCbDiscount = DE_DISb005(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "006"
                        DE_PRCbDiscount = DE_DISb006(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "007"
                        DE_PRCbDiscount = DE_DISb007(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "008"
                        DE_PRCbDiscount = DE_DISb008(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "0010"
                        DE_PRCbDiscount = DE_DISb0010(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "0020"
                        DE_PRCbDiscount = DE_DISb0020(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "0030"         '*Ao 09-09-23
'                        DE_PRCbDiscount = DE_DISb0030(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)) _
'                        , IIf(acDE_SumBlue(pnPos) > 0, acDE_SumBlue(pnPos), pcOldBlue + acDE_SumBlue(pnPos)))
                        DE_PRCbDiscount = DE_DISb0030(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                    Case "0040"         '*Ao 09-09-23
'                        DE_PRCbDiscount = DE_DISb0040(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)) _
'                        , IIf(acDE_SumBlue(pnPos) > 0, acDE_SumBlue(pnPos), pcOldBlue + acDE_SumBlue(pnPos)))
                        DE_PRCbDiscount = DE_DISb0040(pnTime, pnKey, porsQtyFrame(pnPos), pnPos, IIf(acDE_SumDisc(pnPos) > 0, acDE_SumDisc(pnPos), pcOldDisc + acDE_SumDisc(pnPos)))
                End Select
                'bDealComplete = DE_CHKbDealComplete(porsQtyFrame(pnPos), atDE_AllDeal(pnKey), atDE_AllType(pnKey))
                If DE_PRCbDiscount = False Then
                    bDealComplete = False
                End If
            End If
        Loop While bDealComplete
        pnTime = pnTime - 1
        DE_PRCbDiscount = pnTime > 0
    End With
'    ReDim Preserve porsQtyFrame(pnPos) As ADODB.Recordset
'    With porsQtyFrame(pnPos)
'
'    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)

End Function

'Private Function DE_DISb001(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cSource As Double 'เก็บจำนวนซื้อ
'    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
'    Dim nMaxGrp As Integer, nQtyUseRound As Integer, nI As Integer
'    Dim tFilter As String, tSort As String
'    Dim cDisSumDisc As Double
'    Dim cDisSumPrice As Double
'
'    On Error GoTo Err_Handler
'    With porsTemp
'        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
'        .Filter = tFilter
'        tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
'        .Sort = tSort
'        If Not (.BOF And .EOF) Then
'            cSource = .Fields("FNDacQtyBuy").Value
'            cDisSumDisc = 0
'            cDisSumPrice = 0
'            Do While Not .EOF
'                If .Fields("FCSdtQtyAll").Value > 0 Then
'                    If .Fields("FCSdtQtyAll").Value >= cSource Then
''                        nQtyUseRound = .Fields("FCSdtQtyAll").Value \ cSource
'                        Select Case .Fields("FTPddDisType").Value
'                            Case "1" 'Per
'                                cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * .Fields("FCDacDisAmt").Value / 100 * cSource
'                            Case "2" 'Amt
'                                cDisSumDisc = .Fields("FCDacDisAmt").Value
'                            Case "3" 'New
'                                cDisSumPrice = cDisSumPrice + .Fields("FCSdtRegPrice").Value * cSource
'                                cDisSumDisc = .Fields("FCDacDisAmt").Value
'                        End Select
'                        tFilter = .Filter
'                        tSort = .Sort
'                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
'                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)
'                        cSource = 0
'                        Exit Do
'                    Else
'                        Select Case .Fields("FTPddDisType").Value
'                            Case "1" 'Per
'                                cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * .Fields("FCDacDisAmt").Value / 100 * .Fields("FCSdtQtyAll").Value
'                            Case "2" 'Amt
'                                cDisSumDisc = .Fields("FCDacDisAmt").Value
'                            Case "3" 'New
'                                cDisSumPrice = cDisSumPrice + .Fields("FCSdtRegPrice").Value * cSource
'                                cDisSumDisc = .Fields("FCDacDisAmt").Value
'                        End Select
'                        cSource = cSource - .Fields("FCSdtQtyAll").Value
'                        tFilter = .Filter
'                        tSort = .Sort
'                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1") 'cSource
'                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
'
'                    End If
'                End If
'                .MoveNext
'            Loop
'            If cSource = 0 Then
'                DE_DISb001 = True
'                acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
'            Else
'                DE_DISb001 = False
'            End If
'        Else
'            DE_DISb001 = False
'        End If
'    End With
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_DISb001(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 09-10-26
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double, cSourceSum As Double, cDisSumDisc As Double, cDisSumPrice As Double, cRound As Double
    Dim nMaxGrp As Integer, nQtyUseRound As Integer, nI As Integer
    Dim tFilter As String, tSort As String
    
    On Error GoTo Err_Handler
    With porsTemp
        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
        .Filter = tFilter
        tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
        .Sort = tSort
'        cRound = DE_PRCnChkRound(porsTemp, atDE_AllDeal(pnKey), "001")
        
        If Not (.BOF And .EOF) Then
            .MoveFirst
            cSource = .Fields("FNDacQtyBuy").Value
            cDisSumDisc = 0
            cDisSumPrice = 0
            Do While Not .EOF
                If .Fields("FCSdtQtyAll").Value > 0 Then
                    If .Fields("FCSdtQtyAll").Value >= cSource Then
                        Select Case .Fields("FTPddDisType").Value
                            Case "1" 'Per
                                cDisSumDisc = cDisSumDisc + ((.Fields("FCSdtRegPrice").Value * (.Fields("FCDacDisAmt").Value / 100)) * cSource)
                            Case "2" 'Amt
                                cDisSumDisc = .Fields("FCDacDisAmt").Value
                            Case "3" 'New
                                cDisSumPrice = cDisSumPrice + (.Fields("FCSdtRegPrice").Value * cSource)
                                cDisSumDisc = .Fields("FCDacDisAmt").Value
                        End Select
                        tFilter = .Filter
                        tSort = .Sort
                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)
                        cSource = 0
                        Exit Do
                    Else
                        Select Case .Fields("FTPddDisType").Value
                            Case "1" 'Per
                                cDisSumDisc = cDisSumDisc + ((.Fields("FCSdtRegPrice").Value * (.Fields("FCDacDisAmt").Value / 100)) * .Fields("FCSdtQtyAll").Value)
                            Case "2" 'Amt
                                cDisSumDisc = .Fields("FCDacDisAmt").Value
                            Case "3" 'New
                                cDisSumPrice = cDisSumPrice + .Fields("FCSdtRegPrice").Value * cSource
                                cDisSumDisc = .Fields("FCDacDisAmt").Value
                        End Select
                        cSource = cSource - .Fields("FCSdtQtyAll").Value
                        tFilter = .Filter
                        tSort = .Sort
                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1") 'cSource
                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
                        
                    End If
                End If
                .MoveNext
            Loop
            
            If cSource = 0 Then
                DE_DISb001 = True
                acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
            Else
                DE_DISb001 = False
            End If
        Else
            DE_DISb001 = False
        End If
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Function DE_DISb0010(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cSource As Double 'เก็บจำนวนซื้อ
'    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
'    Dim nMaxGrp As Integer
'    Dim nI As Integer
'    Dim tFilter As String
'    Dim tSort As String
'    Dim cDisSumDisc As Double
'    Dim cDisSumPrice As Double
'    Dim nCount As Integer
'    On Error GoTo Err_Handler
'    With porsTemp
'        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
'        .Filter = tFilter
'        '*Ao 52-05-18
''        tSort = "FCSdtRegPrice DESC"
'        tSort = "FTSkuStaMallEmpCard DESC,FCSdtRegPrice"
'        .Sort = tSort
'        If Not (.BOF And .EOF) Then
'            cSource = .Fields("FNDacQtyBuy").Value
'            cDisSumDisc = 0
'            cDisSumPrice = 0
'            Do While Not .EOF
'                If .Fields("FCSdtQtyAll").Value > 0 Then
''                    If .Fields("FCSdtRegPrice").Value > cDisSumDisc Then
''                        cDisSumDisc = .Fields("FCSdtRegPrice").Value
''                    End If
'                    If .Fields("FCSdtQtyAll").Value >= cSource Then
'                        tFilter = .Filter
'                        tSort = .Sort
'                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
'                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)
'                        cSource = 0
'                        cDisSumDisc = .Fields("FCSdtRegPrice").Value
'                        Exit Do
'                    Else
'                        If .Fields("FCSdtQtyAll").Value > 0 Then
'                            cSource = cSource - .Fields("FCSdtQtyAll").Value
'                            tFilter = .Filter
'                            tSort = .Sort
''                            Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1") 'cSource
'                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
'                        End If
'                    End If
'                End If
'                .MoveNext
'            Loop
'            If cSource = 0 Then
'                DE_DISb0010 = True
'                acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
'            Else
'                DE_DISb0010 = False
'            End If
'        Else
'            DE_DISb0010 = False
'        End If
'    End With
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_DISb0010(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 09-10-16
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double, cSourceSum As Double, cDisSumDisc As Double, cDisSumPrice As Double
    Dim nMaxGrp As Integer, nI As Integer, nCount As Integer, nQtyGet As Integer, nQtyCal As Integer, nQtyBuy As Integer
    Dim tFilter As String, tSort As String
    On Error GoTo Err_Handler
    
    With porsTemp
        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
        .Filter = tFilter
        tSort = "FCSdtRegPrice,FTSkuStaMallEmpCard"                 'Buy เลือกราคาน้อยสุด,ราคาเท่ากันเลือกป้ายเหลืองก่อนป้ายน้ำเงิน
        .Sort = tSort
        nCount = DE_PRCnChkRoundFree(porsTemp, "0010")
        
        If Not (.BOF And .EOF) Then
            .MoveFirst
            nQtyBuy = .Fields("FNDacQtyBuy").Value * nCount
            nQtyGet = .Fields("FNDacQtyGet").Value * nCount
            cDisSumDisc = 0
            cDisSumPrice = 0
            'BUY
            Do While Not .EOF
                If .Fields("FCSdtQtyAll").Value > 0 Then
                    If nQtyBuy > 0 Then
                        If .Fields("FCSdtQtyAll").Value >= nQtyBuy Then
                            nQtyCal = nQtyBuy
                            nQtyBuy = 0
                        Else
                            nQtyCal = .Fields("FCSdtQtyAll").Value
                            nQtyBuy = nQtyBuy - .Fields("FCSdtQtyAll").Value
                        End If
                        tFilter = .Filter
                        tSort = .Sort
                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, nQtyCal, "0")
                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'                        cDisSumDisc = cDisSumDisc + (nQtyCal * .Fields("FCSdtRegPrice").Value)
                    Else
                        Exit Do
                    End If
                End If
                .MoveNext
            Loop
            'GET
            tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"                'Get เลือกราคามากสุด,ราคาเท่ากันเลือกป้ายเหลืองก่อนป้ายน้ำเงิน
            .Sort = tSort
            .MoveFirst
            Do While Not .EOF
                 If .Fields("FCSdtQtyAll").Value > 0 Then
                    If nQtyGet > 0 Then
                        If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                            nQtyCal = nQtyGet
                            nQtyGet = 0
                        Else
                            nQtyCal = .Fields("FCSdtQtyAll").Value
                            nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                        End If
                        tFilter = .Filter
                        tSort = .Sort
                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, nQtyCal, "1")
                        Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                        cDisSumDisc = cDisSumDisc + (nQtyCal * .Fields("FCSdtRegPrice").Value)
                    Else
                        Exit Do
                    End If
                End If
                .MoveNext
            Loop
            
            If (nQtyGet = 0) And (nQtyBuy = 0) Then
                DE_DISb0010 = True
                acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
            Else
                DE_DISb0010 = False
            End If
        Else
            DE_DISb0010 = False
        End If
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_DISb002(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double 'เก็บจำนวนซื้อ
    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
    Dim nMaxGrp As Integer
    Dim nI As Integer
    Dim tFilter As String
    Dim tSort As String
    Dim cDisSumDisc As Double
    Dim cSumDisKeep As Double
    Dim cDisSumPrice As Double
    On Error GoTo Err_Handler
    With porsTemp
                nMaxGrp = DE_GETnMaxGroup(atDE_AllDeal(pnKey), porsTemp)
                cDisSumDisc = 0
                cDisSumPrice = 0
                For nI = 1 To nMaxGrp
                    tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacSkuGrp = " & nI
                    .Filter = tFilter
                    tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
                    .Sort = tSort
                    If Not (.BOF And .EOF) Then
                        cSource = .Fields("FNDacQtyGet").Value
                        .MoveFirst
                        Do While Not .EOF
                            If .Fields("FCSdtQtyAll").Value > 0 Then
                                If .Fields("FCSdtQtyAll").Value >= cSource Then
                                    Select Case .Fields("FTPddDisType").Value
                                        Case "1" 'Per
                                            cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * .Fields("FCDacDisAmt").Value / 100 * cSource
                                        Case "2" 'Amt
                                            cDisSumDisc = .Fields("FCDacDisAmt").Value
                                        Case "3" 'New
                                            cDisSumPrice = cDisSumPrice + .Fields("FCSdtRegPrice").Value * cSource
                                            cDisSumDisc = .Fields("FCDacDisAmt").Value
                                        Case "4"
                                            cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * 1
                                    End Select
                                    Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
                                    tFilter = .Filter
                                    tSort = .Sort
                                    Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)
                                    cSource = 0
                                    Exit Do
                                Else
                                    Select Case .Fields("FTPddDisType").Value
                                        Case "1" 'Per
                                            cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * .Fields("FCDacDisAmt").Value / 100 * .Fields("FCSdtQtyAll").Value
                                        Case "2" 'Amt
                                            cDisSumDisc = .Fields("FCDacDisAmt").Value
                                        Case "3" 'New
                                            cDisSumPrice = cDisSumPrice + .Fields("FCSdtRegPrice").Value * cSource
                                            cDisSumDisc = .Fields("FCDacDisAmt").Value
                                        Case "4"
                                            cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value
                                    End Select
                                    cSource = cSource - .Fields("FCSdtQtyAll").Value
                                    Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1") 'cSource)
                                    tFilter = .Filter
                                    tSort = .Sort
'                                    Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
                                    Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)

                                End If
                            End If
                            .MoveNext
                        Loop
                        If cSource > 0 Then
                            Exit For
                        Else
'                            cSumDisKeep = cSumDisKeep + cDisSumDisc            '*Ao 09-10-21
                            cSumDisKeep = cDisSumDisc
                        End If
                    Else
                        DE_DISb002 = False
                    End If
                Next nI
                If cSource = 0 Then
                    DE_DISb002 = True
                    acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cSumDisKeep, cDisSumPrice - cSumDisKeep) + pcOldDisc
                Else
                    DE_DISb002 = False
                End If
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Function DE_DISb0020(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''   Call:
''   Cmt:    '*Ao 09-10-17
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cSource As Double, cSourceSum As Double, cRegPrice As Double, cDealDiscAmt As Double, cDisSumPrice As Double, cDisSumDisc As Double
'    Dim nRound As Integer, nGrpOld As Integer
'    Dim tFilter As String, tSort As String, tFreeFlag As String
'    Dim nGrp As Integer, nUncalGrp As Integer, nQtyBuy As Integer, nQtyCal As Integer, nQtyGet As Integer, nDTSeqNo As Integer, nQtyGetOld As Integer
'    Dim bFoundGet As Boolean
'    Dim orsTemp As ADODB.Recordset
'
'    On Error GoTo Err_Handler
'
'    nUncalGrp = 0
'    nGrpOld = 0
'    'หาตัว GET
'    With porsTemp
'        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
'        .Filter = tFilter
'        nRound = DE_PRCnChkRound(porsTemp, atDE_AllDeal(pnKey), "0020")
'        .Filter = adFilterNone
'
'        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacQtyGet =0"
'        .Filter = tFilter
'        tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard,"
'        .Sort = tSort
'        .MoveFirst
'        Do While Not .EOF
'            nGrp = .Fields("FNDacSkuGrp").Value
'            If .Fields("FCSdtQtyAll").Value > 0 Then
'                If nGrp <> nUncalGrp Then
'                    If nGrp <> nGrpOld Then
'                        If nQtyGet > 0 Then Exit Function
''                        nQtyGet = .Fields("FNDacQtyGet").Value
'                        nQtyGet = .Fields("FNDacQtyBuy").Value * nRound
'                        nGrpOld = nGrp
'                    End If
'
'                    If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                        nQtyCal = nQtyGet
'                        nQtyGet = 0
'                    Else
'                        nQtyCal = .Fields("FCSdtQtyAll").Value
'                        nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
'                    End If
'
''                    cRegPrice = .Fields("FCSdtRegPrice").Value
''                    nDTSeqNo = .Fields("FNSdtSeqNo").Value
''                    cDealDiscAmt = cRegPrice * nQtyCal
''                    cDisSumPrice = cDisSumPrice + cDealDiscAmt
''                    Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1")
''                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'
'                    cRegPrice = .Fields("FCSdtRegPrice").Value
'                    nDTSeqNo = .Fields("FNSdtSeqNo").Value
'                    Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "0")
'                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'
'                    If nQtyGet < 1 Then nUncalGrp = nGrp
'                End If
'            End If
'            .MoveNext
'        Loop
'
'        '*Ao 09-10-29
'        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacQtyBuy =0"
'        .Filter = tFilter
'        tSort = "FNDacSkuGrp DESC,FCSdtRegPrice DESC,FTSkuStaMallEmpCard,"
'        .Sort = tSort
'        .MoveFirst
'        nGrpOld = 0
'        nUncalGrp = 0
'
'        Do While Not porsTemp.EOF
'            nGrp = .Fields("FNDacSkuGrp").Value
'            If .Fields("FCSdtQtyAll").Value > 0 Then
'                If nGrp <> nUncalGrp Then
'                    If nGrp <> nGrpOld Then
'                        If nQtyGet > 0 Then Exit Do
'                        nQtyGet = .Fields("FNDacQtyGet").Value * nRound
'                        nGrpOld = nGrp
'                    End If
'
'                    If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                        nQtyCal = nQtyGet
'                        nQtyGet = 0
'                    Else
'                        nQtyCal = .Fields("FCSdtQtyAll").Value
'                        nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
'                    End If
''                   cRegPrice = .Fields("FCSdtRegPrice").Value
''                   nDTSeqNo = .Fields("FNSdtSeqNo").Value
'''                   Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1")
''                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'
'                    cRegPrice = .Fields("FCSdtRegPrice").Value
'                    nDTSeqNo = .Fields("FNSdtSeqNo").Value
'                    cDealDiscAmt = cRegPrice * nQtyCal
'                    cDisSumPrice = cDisSumPrice + cDealDiscAmt
'                    Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1")
'                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'
'                    If nQtyGet < 1 Then nUncalGrp = nGrp
'                End If
'            End If
'            .MoveNext
'        Loop
'
'        If nQtyGet < 1 Then
'            DE_DISb0020 = True
''            acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
'            acDE_SumDisc(pnPos) = cDisSumPrice + pcOldDisc
'        Else
'            acDE_SumDisc(pnPos) = 0
'            DE_DISb0020 = False
'        End If
'    End With
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_DISb0020(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 09-10-17
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double, cSourceSum As Double, cRegPrice As Double, cDealDiscAmt As Double, cDisSumPrice As Double, cDisSumDisc As Double
    Dim nRound As Integer, nGrpOld As Integer
    Dim tFilter As String, tSort As String, tFreeFlag As String
    Dim nGrp As Integer, nUncalGrp As Integer, nQtyBuy As Integer, nQtyCal As Integer, nQtyGet As Integer, nDTSeqNo As Integer, nQtyGetOld As Integer
    Dim bFoundGet As Boolean
    Dim orsTemp As ADODB.Recordset
    Dim tFree0020  As String
    
    On Error GoTo Err_Handler

    nUncalGrp = 0
    nGrpOld = 0
    'หาตัว GET
    With porsTemp
        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
        .Filter = tFilter
        nRound = DE_PRCnChkRound(porsTemp, atDE_AllDeal(pnKey), "0020")
        .Filter = adFilterNone
        
        If nRound = 0 Then Exit Function
        If tVB_Free0020 = "" Then Exit Function
        tFree0020 = tVB_Free0020
    
        Do While tFree0020 <> ""
            nDTSeqNo = CInt(Trim(UT_tToken(tFree0020, ",", 1, False)))
            tFilter = "(FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "') AND (FNDacQtyGet > 0) AND (FNSdtSeqNo = " & nDTSeqNo & ")"
            .Filter = tFilter
            tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
            .Sort = tSort
            nGrp = .Fields("FNDacSkuGrp").Value
            If Not .EOF Then
                .MoveFirst
                If .Fields("FCSdtQtyAll").Value > 0 Then
                    If nGrp <> nUncalGrp Then
                        If nGrp <> nGrpOld Then
                            nQtyGet = .Fields("FNDacQtyGet").Value * nRound
                            nGrpOld = nGrp
                        End If
    
                        If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                            nQtyCal = nQtyGet
                            nQtyGet = 0
                        Else
                            nQtyCal = .Fields("FCSdtQtyAll").Value
                            nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                        End If
    
                        cRegPrice = .Fields("FCSdtRegPrice").Value
                        nDTSeqNo = .Fields("FNSdtSeqNo").Value
                        cDealDiscAmt = cRegPrice * nQtyCal
                        cDisSumPrice = cDisSumPrice + cDealDiscAmt
'                        Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "0")
                        Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1")
                        Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
    
                        If nQtyGet < 1 Then nUncalGrp = nGrp
                    End If
                End If
            End If
        Loop

        nGrpOld = 0
        nUncalGrp = 0
        .Filter = adFilterNone
        'ทำตัว Buy
        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND (FNDacQtyBuy > 0)"
        .Filter = tFilter
        tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
        .Sort = tSort
        If Not .EOF Then
            .MoveFirst
        End If
        
        Do While Not porsTemp.EOF
            nGrp = .Fields("FNDacSkuGrp").Value
            If .Fields("FCSdtQtyAll").Value > 0 Then
                If nGrp <> nUncalGrp Then
                    If nGrp <> nGrpOld Then
                        nQtyGet = .Fields("FNDacQtyBuy").Value * nRound
                        nGrpOld = nGrp
                    End If

                    If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                        nQtyCal = nQtyGet
                        nQtyGet = 0
                    Else
                        nQtyCal = .Fields("FCSdtQtyAll").Value
                        nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                    End If

                    cRegPrice = .Fields("FCSdtRegPrice").Value
                    nDTSeqNo = .Fields("FNSdtSeqNo").Value
'                    Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1")
                    Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "0")
                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)

                    If nQtyGet < 1 Then nUncalGrp = nGrp
                End If
            End If
            .MoveNext
        Loop

        If nQtyGet < 1 Then
            DE_DISb0020 = True
            acDE_SumDisc(pnPos) = cDisSumPrice + pcOldDisc
        Else
            acDE_SumDisc(pnPos) = 0
            DE_DISb0020 = False
        End If
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_DISb003(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double, ByRef pcOldBlue As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double 'เก็บจำนวนซื้อ
    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
    Dim nMaxGrp As Integer
    Dim nI As Integer
    Dim tFilter As String
    Dim tSort As String
    Dim cDisSumDisc As Double
    Dim cDisSumPrice As Double
    Dim cBlue As Double
    
    On Error GoTo Err_Handler
    With porsTemp
        nMaxGrp = DE_GETnMaxGroup(atDE_AllDeal(pnKey), porsTemp)
        If nMaxGrp <> 2 Then Exit Function
        cDisSumDisc = 0
        cDisSumPrice = 0
        cBlue = 0
        For nI = 1 To nMaxGrp
            tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacSkuGrp = " & nI
            .Filter = tFilter
            If Not (.BOF And .EOF) Then
                .MoveFirst
                If nI = 2 Then
                    tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
                    cSource = .Fields("FNDacQtyGet").Value
                Else
                    tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
                    cSource = .Fields("FNDacQtyBuy").Value
                End If
                .Sort = tSort
                Do While Not .EOF
                    If .Fields("FCSdtQtyAll").Value > 0 Then
                        If .Fields("FCSdtQtyAll").Value >= cSource Then
                            If nI = 2 Then
                                Select Case .Fields("FTPddDisType").Value
                                    Case "1" 'Per
                                        cDisSumDisc = cDisSumDisc + ((.Fields("FCSdtRegPrice").Value * (.Fields("FCDacDisAmt").Value / 100)) * cSource)
                                    Case "2" 'Amt
                                        cDisSumDisc = .Fields("FCDacDisAmt").Value
                                    Case "3" 'New
                                        cDisSumPrice = cDisSumPrice + (.Fields("FCSdtRegPrice").Value * cSource)
                                        cDisSumDisc = .Fields("FCDacDisAmt").Value
                                    Case Else
                                        DE_DISb003 = False
                                        Exit Function
'                                    Case "4"
'                                        cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * cSource
                                End Select
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
                            Else
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "0")
                            End If
                            '*Ao 52-04-27
                            Select Case UCase(.Fields("FTSkuStaMallEmpCard").Value)     'เก็บ Sum ป้ายน้ำเงิน
                                Case "Y"
                                    cBlue = cBlue + (.Fields("FCSdtRegPrice").Value * cSource)
                            End Select
                            '__________________________________
                            tFilter = .Filter
                            tSort = .Sort
                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)
                            cSource = 0
                            Exit Do
                        Else
                            If nI = 2 Then
                                Select Case .Fields("FTPddDisType").Value
                                    Case "1" 'Per
                                        cDisSumDisc = cDisSumDisc + ((.Fields("FCSdtRegPrice").Value * (.Fields("FCDacDisAmt").Value / 100)) * .Fields("FCSdtQtyAll").Value)
                                    Case "2" 'Amt
                                        cDisSumDisc = .Fields("FCDacDisAmt").Value
                                    Case "3" 'New
'                                        cDisSumPrice = cDisSumPrice + (.Fields("FCSdtRegPrice").Value * cSource)
                                        cDisSumPrice = cDisSumPrice + (.Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value)
                                        cDisSumDisc = .Fields("FCDacDisAmt").Value
                                    Case Else
                                        DE_DISb003 = False
                                        Exit Function
'                                    Case "4"
'                                        cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value
                                End Select
                                cSource = cSource - .Fields("FCSdtQtyAll").Value
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1")
                            Else
                                cSource = cSource - .Fields("FCSdtQtyAll").Value
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "0")
                            End If
                            '*Ao 52-04-27
                            Select Case UCase(.Fields("FTSkuStaMallEmpCard").Value)     'เก็บ Sum ป้ายน้ำเงิน
                                Case "Y"
                                    cBlue = cBlue + (.Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value)
                            End Select
                            '__________________________________
                            tFilter = .Filter
                            tSort = .Sort
                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
                        End If
                    End If          'If .Fields("FCSdtQtyAll").Value > 0 Then
                    .MoveNext
                Loop
                If cSource > 0 Then
                    Exit For
                End If
            Else
                DE_DISb003 = False
                Exit Function
            End If
        Next nI
        If cSource = 0 Then
            DE_DISb003 = True
            acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
            acDE_SumBlue(pnPos) = cBlue + pcOldBlue
        Else
            DE_DISb003 = False
        End If
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Function DE_DISb0030(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''   Call:
''   Cmt:    '*Ao 09-09-24
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cSource As Double 'เก็บจำนวนซื้อ
'    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
'    Dim nMaxGrp As Integer
'    Dim nI As Integer
'    Dim tFilter As String
'    Dim tSort As String
'    Dim cDisSumDisc As Double
'    Dim cDisSumPrice As Double
'    On Error GoTo Err_Handler
'    With porsTemp
'                nMaxGrp = DE_GETnMaxGroup(atDE_AllDeal(pnKey), porsTemp)
'                If nMaxGrp <> 2 Then Exit Function
'                cDisSumDisc = 0
'                cDisSumPrice = 0
'                For nI = 1 To nMaxGrp
'                    tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacSkuGrp = " & nI
'                    .Filter = tFilter
'                    '*Ao 52-05-18
''                    tSort = "FCSdtRegPrice DESC"
'                    tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
'                    .Sort = tSort
'                    If Not (.BOF And .EOF) Then
'                        If nI = 2 Then
'                            cSource = .Fields("FNDacQtyGet").Value
'                        Else
'                            cSource = .Fields("FNDacQtyBuy").Value
'                        End If
'                        .MoveFirst
'                        Do While Not .EOF
'                            If .Fields("FCSdtQtyAll").Value > 0 Then
'                                If .Fields("FCSdtQtyAll").Value >= cSource Then
'                                    If nI = 2 Then
'                                        cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * cSource
'                                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
'                                    Else
''                                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "0")          '*Ao 09-09-24 สินค้าตัวหลักไม่ต้องมี *
'                                    End If
'                                    tFilter = .Filter
'                                    tSort = .Sort
'                                    Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)
'                                    cSource = 0
'                                    Exit Do
'                                Else
'                                    If nI = 2 Then
'                                        cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value
'                                        cSource = cSource - .Fields("FCSdtQtyAll").Value
'                                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1")
'                                    Else
'                                        cSource = cSource - .Fields("FCSdtQtyAll").Value
''                                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "0")        '*Ao 09-09-24 สินค้าตัวหลักไม่ต้องมี *
'                                    End If
'                                    tFilter = .Filter
'                                    tSort = .Sort
'                                    Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
'                                End If
'                            End If
'                            .MoveNext
'                        Loop
'                        If cSource > 0 Then
'                            Exit For
'                        End If
'                    Else
'                        DE_DISb0030 = False
'                    End If
'                Next nI
'                If cSource = 0 Then
'                    DE_DISb0030 = True
'                    acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
'                Else
'                    DE_DISb0030 = False
'                End If
'    End With
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_DISb0030(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:        '*Ao 09-10-27
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double, cDisSumDisc As Double, cDisSumPrice As Double, cRound As Double
    Dim nMaxGrp As Integer, nGrp As Integer
    Dim tFilter As String, tSort As String
    
    On Error GoTo Err_Handler
    With porsTemp
        nMaxGrp = DE_GETnMaxGroup(atDE_AllDeal(pnKey), porsTemp)
        If Not (nMaxGrp = 2) Then Exit Function
        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
        .Filter = tFilter
        cRound = DE_PRCnChkRound(porsTemp, atDE_AllDeal(pnKey), "0030")
        .Filter = adFilterNone
        
        cDisSumDisc = 0
        cDisSumPrice = 0
        For nGrp = 1 To nMaxGrp
            tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacSkuGrp = " & nGrp
            .Filter = tFilter
            If Not (.BOF And .EOF) Then
                If nGrp = 2 Then
                    cSource = .Fields("FNDacQtyGet").Value * cRound
                    tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
                Else
                    cSource = .Fields("FNDacQtyBuy").Value * cRound
                    tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
                End If
                .Sort = tSort
                .MoveFirst
                Do While Not .EOF
                    If .Fields("FCSdtQtyAll").Value > 0 Then
                        If .Fields("FCSdtQtyAll").Value >= cSource Then
                            If nGrp = 2 Then
                                cDisSumDisc = cDisSumDisc + (.Fields("FCSdtRegPrice").Value * cSource)
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
                            Else
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "0")
                            End If
                            tFilter = .Filter
                            tSort = .Sort
                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)
                            cSource = 0
                            Exit Do
                        Else
                            If nGrp = 2 Then
                                cDisSumDisc = cDisSumDisc + (.Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value)
'                                cSource = cSource - .Fields("FCSdtQtyAll").Value
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1")
                            Else
'                                cSource = cSource - .Fields("FCSdtQtyAll").Value
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "0")
                            End If
                            cSource = cSource - .Fields("FCSdtQtyAll").Value
                            tFilter = .Filter
                            tSort = .Sort
                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
                        End If
                    End If
                    .MoveNext
                Loop
                If cSource > 0 Then
                    Exit For
                End If
            Else
                DE_DISb0030 = False
            End If
        Next nGrp
        If cSource = 0 Then
            DE_DISb0030 = True
            acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
        Else
            DE_DISb0030 = False
        End If
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Function DE_DISb004(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cSource As Double 'เก็บจำนวนซื้อ
'    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
'    Dim nMaxGrp As Integer
'    Dim nI As Integer
'    Dim tFilter As String
'    Dim tSort As String
'    Dim cDisSumDisc As Double
'    Dim cDisSumPrice As Double
'    Dim cBlue As Double
'
'    On Error GoTo Err_Handler
'    With porsTemp
'        nMaxGrp = DE_GETnMaxGroup(atDE_AllDeal(pnKey), porsTemp)
'        If nMaxGrp <> 3 Then Exit Function
'        cDisSumDisc = 0
'        cDisSumPrice = 0
'        For nI = 1 To nMaxGrp
'            tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacSkuGrp = " & nI
'            .Filter = tFilter
'            If Not (.BOF And .EOF) Then
'                .MoveFirst
'                '*Ao 09-10-22
'                If nI = 3 Then
'                    tSort = "FCSdtRegPrice DESC"
'                    cSource = .Fields("FNDacQtyGet").Value
'                Else
'                    tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
'                    cSource = .Fields("FNDacQtyBuy").Value
'                End If
'                '_________________________________________________
'                .Sort = tSort
'                Do While Not .EOF
'                    If .Fields("FCSdtQtyAll").Value > 0 Then
'                        If .Fields("FCSdtQtyAll").Value >= cSource Then
'                            If nI = 3 Then
'                                Select Case .Fields("FTPddDisType").Value
'                                    Case "1" 'Per
'                                        cDisSumDisc = cDisSumDisc + ((.Fields("FCSdtRegPrice").Value * (.Fields("FCDacDisAmt").Value / 100)) * cSource)
'                                    Case "2" 'Amt
'                                        cDisSumDisc = .Fields("FCDacDisAmt").Value
'                                    Case "3" 'New
'                                        cDisSumPrice = cDisSumPrice + (.Fields("FCSdtRegPrice").Value * cSource)
'                                        cDisSumDisc = .Fields("FCDacDisAmt").Value
'                                    Case "4"
'                                        cDisSumDisc = cDisSumDisc + (.Fields("FCSdtRegPrice").Value * cSource)
'                                    Case Else
'                                        DE_DISb004 = False
'                                        Exit Function
'                                End Select
'                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
'                            Else
'                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "0")
'                            End If
'                            tFilter = .Filter
'                            tSort = .Sort
'                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)
'                            cSource = 0
'                            Exit Do
'                        Else
'                            If nI = 3 Then
'                                Select Case .Fields("FTPddDisType").Value
'                                    Case "1" 'Per
'                                        cDisSumDisc = cDisSumDisc + ((.Fields("FCSdtRegPrice").Value * (.Fields("FCDacDisAmt").Value / 100)) * .Fields("FCSdtQtyAll").Value)
'                                    Case "2" 'Amt
'                                        cDisSumDisc = .Fields("FCDacDisAmt").Value
'                                    Case "3" 'New
''                                        cDisSumPrice = cDisSumPrice + (.Fields("FCSdtRegPrice").Value * cSource)
'                                        cDisSumPrice = cDisSumPrice + (.Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value)
'                                        cDisSumDisc = .Fields("FCDacDisAmt").Value
'                                    Case "4"
'                                        cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value
'                                    Case Else
'                                        DE_DISb004 = False
'                                        Exit Function
'                                End Select
'                                cSource = cSource - .Fields("FCSdtQtyAll").Value
'                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1") 'cSource)
'                            Else
'                                cSource = cSource - .Fields("FCSdtQtyAll").Value
'                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "0") 'cSource)
'                            End If
'                            tFilter = .Filter
'                            tSort = .Sort
'                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
'                        End If
'                    End If
'                    .MoveNext
'                Loop
'                If cSource > 0 Then
'                    Exit For
'                End If
'            Else
'                DE_DISb004 = False
'            End If
'        Next nI
'        If cSource = 0 Then
'            DE_DISb004 = True
'            acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
'        Else
'            DE_DISb004 = False
'        End If
'    End With
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_DISb004(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double, ByRef pcOldBlue As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 09-11-12
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double, cSourceSum As Double, cDisSumPrice As Double, cDisSumDisc As Double
    Dim nMaxGrp As Integer, nI As Integer, nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nQtyCal As Integer
    Dim tFilter As String, tSort As String
    Dim cBlue As Double
    On Error GoTo Err_Handler
    
    With porsTemp
        nMaxGrp = 2
        cDisSumDisc = 0
        cDisSumPrice = 0
        nUncalGrp = 0
        nGrp = 0
        cBlue = 0
        For nI = 1 To nMaxGrp
            If nI = 1 Then
                tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacQtyGet = 0"             'BUY
                tSort = "FNDacSkuGrp,FTSkuStaMallEmpCard,FCSdtRegPrice"
            Else
                tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacQtyBuy = 0"             'GET
                tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
            End If
                
            .Filter = tFilter
            .Sort = tSort
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    If .Fields("FCSdtQtyAll").Value > 0 Then
                        nGrp = .Fields("FNDacSkuGrp").Value
                        If nGrp <> nUncalGrp Then
                            If nGrp <> nGrpOld Then
                                Select Case .Fields("FTFreeFlag").Value
                                    Case "1" 'Buy Only
                                        cSource = .Fields("FNDacQtyBuy").Value
                                    Case "2" 'Get or Free Only
                                        cSource = .Fields("FNDacQtyGet").Value
                                    Case "3" 'Both (Get or Free) And Buy
                                        cSource = .Fields("FNDacQtyGet").Value
                                End Select
                                
                                If .Fields("FCSdtQtyAll").Value >= cSource Then
                                    nQtyCal = cSource
                                    cSource = 0
                                Else
                                    nQtyCal = .Fields("FCSdtQtyAll").Value
                                    cSource = cSource - .Fields("FCSdtQtyAll").Value
                                End If
                                
                                If nI = 2 Then
                                    Select Case .Fields("FTPddDisType").Value
                                        Case "1" 'Per
                                            cDisSumDisc = cDisSumDisc + ((.Fields("FCSdtRegPrice").Value * (.Fields("FCDacDisAmt").Value / 100)) * nQtyCal)
                                        Case "2" 'Amt
                                            cDisSumDisc = .Fields("FCDacDisAmt").Value
                                        Case "3" 'New
                                            cDisSumPrice = cDisSumPrice + (.Fields("FCSdtRegPrice").Value * nQtyCal)
                                            cDisSumDisc = .Fields("FCDacDisAmt").Value
                                        Case "4"
                                            cDisSumDisc = cDisSumDisc + (.Fields("FCSdtRegPrice").Value * nQtyCal)
                                        Case Else
                                            cDisSumPrice = 0
                                            cDisSumDisc = 0
                                    End Select
                                    Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, nQtyCal, "1")
                                Else
                                    Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, nQtyCal, "0")
                                End If
                                '*Ao 52-04-27
                                Select Case UCase(.Fields("FTSkuStaMallEmpCard").Value)     'เก็บ Sum ป้ายน้ำเงิน
                                    Case "Y"
                                        cBlue = cBlue + (.Fields("FCSdtRegPrice").Value * nQtyCal)
                                End Select
                                '__________________________________
                                Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                                nGrpOld = nGrp
                                If cSource <= 0 Then
                                    nUncalGrp = nGrp
                                End If
                                
                            Else        'If nGrp <> nGrpOld Then
                                If .Fields("FCSdtQtyAll").Value >= cSource Then
                                    nQtyCal = cSource
                                    cSource = 0
                                Else
                                    nQtyCal = .Fields("FCSdtQtyAll").Value
                                    cSource = cSource - .Fields("FCSdtQtyAll").Value
                                End If
                                
                                If nI = 2 Then
                                    Select Case .Fields("FTPddDisType").Value
                                        Case "1" 'Per
                                            cDisSumDisc = cDisSumDisc + ((.Fields("FCSdtRegPrice").Value * (.Fields("FCDacDisAmt").Value / 100)) * nQtyCal)
                                        Case "2" 'Amt
                                            cDisSumDisc = .Fields("FCDacDisAmt").Value
                                        Case "3" 'New
                                            cDisSumPrice = cDisSumPrice + (.Fields("FCSdtRegPrice").Value * nQtyCal)
                                            cDisSumDisc = .Fields("FCDacDisAmt").Value
                                        Case "4"
                                            cDisSumDisc = cDisSumDisc + (.Fields("FCSdtRegPrice").Value * nQtyCal)
                                        Case Else
                                            cDisSumPrice = 0
                                            cDisSumDisc = 0
                                    End Select
                                    Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, nQtyCal, "1")
                                Else
                                    Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, nQtyCal, "0")
                                End If
                                '*Ao 52-04-27
                                Select Case UCase(.Fields("FTSkuStaMallEmpCard").Value)     'เก็บ Sum ป้ายน้ำเงิน
                                    Case "Y"
                                        cBlue = cBlue + (.Fields("FCSdtRegPrice").Value * nQtyCal)
                                End Select
                                '__________________________________
                                Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                                nGrpOld = nGrp
                                If cSource <= 0 Then
                                    nUncalGrp = nGrp
                                End If
                                
                            End If      'If nGrp <> nGrpOld Then
                        End If      'If nGrp <> nUncalGrp Then
                    End If      'If .Fields("FCSdtQtyAll").Value > 0 Then
                    .MoveNext
                Loop
            Else
                DE_DISb004 = False
                Exit Function
            End If      'If Not (.BOF And .EOF) Then
        Next nI
        
        If cSource = 0 Then
            DE_DISb004 = True
            acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
            acDE_SumBlue(pnPos) = cBlue + pcOldBlue
        Else
            DE_DISb004 = False
        End If
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Function DE_DISb0040(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cSource As Double 'เก็บจำนวนซื้อ
'    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
'    Dim nMaxGrp As Integer
'    Dim nI As Integer
'    Dim tFilter As String
'    Dim tSort As String
'    Dim cDisSumDisc As Double
'    Dim cDisSumPrice As Double
'    On Error GoTo Err_Handler
'    With porsTemp
'                nMaxGrp = DE_GETnMaxGroup(atDE_AllDeal(pnKey), porsTemp)
'                If nMaxGrp <> 3 Then Exit Function
'                cDisSumDisc = 0
'                cDisSumPrice = 0
'                For nI = 1 To nMaxGrp
'                    tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacSkuGrp = " & nI
'                    .Filter = tFilter
'                    '*Ao 52-05-18
''                    tSort = "FCSdtRegPrice DESC"
'                    tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
'                    .Sort = tSort
'                    If Not (.BOF And .EOF) Then
'                        If nI = 3 Then
'                            cSource = .Fields("FNDacQtyGet").Value
'                        Else
'                            cSource = .Fields("FNDacQtyBuy").Value
'                        End If
'                        .MoveFirst
'                        Do While Not .EOF
'                            If .Fields("FCSdtQtyAll").Value > 0 Then
'                                If .Fields("FCSdtQtyAll").Value >= cSource Then
'                                    If nI = 3 Then
'                                        cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * cSource
'                                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
'                                    Else
'                                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "0")
'                                    End If
'                                    tFilter = .Filter
'                                    tSort = .Sort
'                                    Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)
'                                    cSource = 0
'                                    Exit Do
'                                Else
'                                    If nI = 3 Then
'                                        cDisSumDisc = cDisSumDisc + .Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value
'                                        cSource = cSource - .Fields("FCSdtQtyAll").Value
'                                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1") 'cSource)
'                                    Else
'                                        cSource = cSource - .Fields("FCSdtQtyAll").Value
'                                        Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "0") 'cSource)
'                                    End If
'                                    tFilter = .Filter
'                                    tSort = .Sort
'                                    Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
'                                End If
'                            End If
'                            .MoveNext
'                        Loop
'                        If cSource > 0 Then
'                            Exit For
'                        End If
'                    Else
'                        DE_DISb0040 = False
'                End If
'                Next nI
'                If cSource = 0 Then
'                    DE_DISb0040 = True
'                    acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
'                Else
'                    DE_DISb0040 = False
'                End If
'    End With
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_DISb005(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double, cSourceLef As Double, cSourceSum As Double, cDisChkPrice As Double, cDisSumPrice As Double, cDisSumDisc As Double
    Dim nMaxGrp As Integer, nI As Integer
    Dim tDealNo As String, tFreeFlag As String, tFilter As String, tSort As String
    Dim cAmtGet As Double, cAmtGetOrg As Double, cQtyAll As Double, cDelRegPrice As Double, cDelSaleAmt As Double
    Dim cDelDisAmt As Double, cDealDiscPer As Double, cAmtCal As Double, cDealDiscAmt As Double, cSumCanDisc As Double
    Dim nDTSeqNo As Integer, nQtyCal As Integer, nDTSeqNoNew As Integer, nPos As Integer
    
    On Error GoTo Err_Handler
    With porsTemp
        cDisSumDisc = 0
        cDisSumPrice = 0
        tDealNo = atDE_AllDeal(pnKey)
        tFilter = "FTPdhDealNo = '" & tDealNo & "'"
        .Filter = tFilter
        tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
        .Sort = tSort
        If Not (.BOF And .EOF) Then
            nMaxGrp = DE_GETnMaxGroup(tDealNo, porsTemp)
            .MoveFirst
            If nMaxGrp = 1 Then
                cAmtGet = .Fields("FCDacAmtBuy").Value
                cAmtGetOrg = cAmtGet
                Do While Not .EOF
                    cQtyAll = .Fields("FCSdtQtyAll").Value
                    If cQtyAll > 0 Then
                        cDelRegPrice = .Fields("FCSdtRegPrice").Value
                        cDelSaleAmt = cDelRegPrice * cQtyAll
                        cDelDisAmt = .Fields("FCDacDisAmt").Value
                        nDTSeqNo = .Fields("FNSdtSeqNo").Value
                        Select Case .Fields("FTPddDisType").Value
                            Case "1" 'Per
                                cDealDiscPer = cDelDisAmt
                            Case "2", "3"
                                cDealDiscPer = 0
                            Case Else
                                cDealDiscPer = 0
                        End Select
                        tFreeFlag = "3"
                        If cDelSaleAmt >= cAmtGet Then
                            cAmtCal = cAmtGet
                            cAmtGet = 0
                        Else
                            cAmtCal = cDelSaleAmt
                            cAmtGet = cAmtGet - cDelSaleAmt
                        End If
'                        nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal Mod cDelRegPrice = 0, 0, 1)     '*Ao 09-10-20
                        nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal \ cDelRegPrice = cAmtCal / cDelRegPrice, 0, 1)
                        If cSumCanDisc <= 0 Then
                            nPos = porsTemp.AbsolutePosition
                            cSumCanDisc = DE_GETcSumDiscByAmt(porsTemp, cAmtGetOrg, tFilter, tSort)
                            .Filter = tFilter
                            .Sort = tSort
                            porsTemp.AbsolutePosition = nPos
                        End If
                        
                        Select Case porsTemp.Fields("FTPddDisType").Value
                            Case "1" 'Per
'                                cDealDiscAmt = (cDealDiscPer / 100) * cAmtCal
                                cDelDisAmt = (cDelDisAmt / 100) * cAmtGetOrg
                                cDealDiscAmt = ((cDelRegPrice * nQtyCal) * cDelDisAmt) / cSumCanDisc
                            Case "2"
'                                cDealDiscAmt = cDelDisAmt * cAmtCal / cAmtGetOrg
                                cDealDiscAmt = ((cDelRegPrice * nQtyCal) * cDelDisAmt) / cSumCanDisc
                            Case "3"
'                                cDealDiscAmt = (cDelRegPrice * nQtyCal) - (cDelDisAmt * cAmtCal / cAmtGetOrg)
                                cDealDiscAmt = ((cDelRegPrice * nQtyCal) * (cAmtGetOrg - cDelDisAmt)) / cSumCanDisc
                        End Select
                        cDisSumDisc = cDisSumDisc + cDealDiscAmt
                        tFilter = .Filter
                        tSort = .Sort
                        Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1")
                        Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                        If cAmtGet <= 0 Then
                            Exit Do
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End If
        If cDisSumDisc > 0 Then
            DE_DISb005 = True
            acDE_SumDisc(pnPos) = cDisSumDisc + pcOldDisc
        Else
            DE_DISb005 = False
        End If
    
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Function DE_DISb006(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cSource As Double 'เก็บจำนวนซื้อ
'    Dim cSourceLef As Double
'    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
'    Dim nMaxGrp As Integer
'    Dim nI As Integer
'    Dim nGrp As Integer
'    Dim tFilter As String
'    Dim tSort As String
'    Dim cDisSumDisc As Double
'    Dim cDisSumPrice As Double
'    Dim cDisChkPrice As Double
'
'    Dim nQtyGet As Integer, nUncalGrp As Integer, nDTSeqNo As Integer, nGrpOld As Integer, nQtyCal As Integer, nQtyGetOrg As Integer, nPos As Integer
'    Dim tDealNo As String, tFreeFlag As String
'    Dim cDelRegPrice As Double, cQtyAll As Double, cDelSaleAmt As Double, cDelDisAmt As Double, cDealDiscPer As Double, cAmtGet As Double
'    Dim cAmtGetOrg As Double, cAmtCal As Double, cDealDiscAmt As Double, cSumCanDisc As Double
'    On Error GoTo Err_Handler
'    nQtyGet = 1
'    With porsTemp
'        tDealNo = atDE_AllDeal(pnKey)
'        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
'        .Filter = tFilter
'        tSort = "FNDacSkuGrp,FTSkuStaMallEmpCard,FCSdtRegPrice"
'        .Sort = tSort
'        nMaxGrp = DE_GETnMaxGroup(tDealNo, porsTemp)
'        If nMaxGrp = 2 Then
'            If Not (.BOF And .EOF) Then
'                .MoveFirst
'                Do While Not porsTemp.EOF
'                    nGrp = .Fields("FNDacSkuGrp").Value
'                    cQtyAll = .Fields("FCSdtQtyAll").Value
'                    If nGrp = 1 Then 'หา Amt
'                        If cQtyAll > 0 Then
'                            cDelRegPrice = .Fields("FCSdtRegPrice").Value
'                            cQtyAll = .Fields("FCSdtQtyAll").Value
'                            cDelSaleAmt = cQtyAll * cDelRegPrice
'                            cDelDisAmt = .Fields("FCDacDisAmt").Value
'                            If nGrp <> nUncalGrp Then
'                                nDTSeqNo = .Fields("FNSdtSeqNo").Value
'                                If nGrp <> nGrpOld Then
'                                    cDealDiscPer = 0
'                                    tFreeFlag = "1"
'                                    cAmtGet = .Fields("FCDacAmtBuy").Value
'                                    cAmtGetOrg = cAmtGet
'                                    If cDelSaleAmt >= cAmtGet Then
'                                        cAmtCal = cAmtGet
'                                        cAmtGet = 0
'                                    Else
'                                        cAmtCal = cDelSaleAmt
'                                        cAmtGet = cAmtGet - cDelSaleAmt
'                                    End If
'                                    cDealDiscAmt = 0
'    '                                nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal Mod cDelRegPrice = 0, 0, 1)              '*09-10-20
'                                    nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal \ cDelRegPrice = cAmtCal / cDelRegPrice, 0, 1)
'                                    'nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                    tFilter = .Filter
'                                    tSort = .Sort
'                                    Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "0") 'cSource)
'                                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'
'                                    'Call DE_GETcCallDeal(tDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cAmtCal)
'                                    'Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'
'                                    nGrpOld = nGrp
'                                    If cAmtGet <= 0 Then
'                                        nUncalGrp = nGrp
'                                    End If
'                                Else
'                                    If cDelSaleAmt >= cAmtGet Then
'                                        cAmtCal = cAmtGet
'                                        cAmtGet = 0
'                                    Else
'                                        cAmtCal = cDelSaleAmt
'                                        cAmtGet = cAmtGet - cDelSaleAmt
'                                    End If
'                                    cDealDiscAmt = 0
'                                    cDealDiscPer = 0
'    '                                nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal Mod cDelRegPrice = 0, 0, 1)                 '*Ao 09-10-20
'                                    nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal \ cDelRegPrice = cAmtCal / cDelRegPrice, 0, 1)
'    '                                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                    tFilter = .Filter
'                                    tSort = .Sort
'                                    Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "0") 'cSource)
'                                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'
'                                    'Call DE_GETcCallDeal(tDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cAmtCal)
'                                    'Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                    If cAmtGet <= 0 Then
'                                        nUncalGrp = nGrp
'                                    End If
'                                End If
'                            End If
'                        End If
'                    Else 'Group = 2
'                        If .Fields("FCSdtQtyAll").Value > 0 Then
'                            cDelRegPrice = .Fields("FCSdtRegPrice").Value
'                            nDTSeqNo = .Fields("FNSdtSeqNo").Value
'                            If nGrp <> nGrpOld Then
'                                Select Case .Fields("FTPddDisType").Value
'                                    Case "1" 'Per
'                                        cDealDiscPer = .Fields("FCDacDisAmt").Value
'                                    Case "2", "3"
'                                        cDealDiscPer = 0
'                                    Case Else
'                                        cDealDiscPer = 0
'                                End Select
'                                tFreeFlag = .Fields("FTFreeFlag").Value
'                                nQtyGet = .Fields("FNDacQtyGet").Value
'                                nQtyGetOrg = nQtyGet
'
'                                nPos = .AbsolutePosition
'                                cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, nGrp) '*TW 51-06-27
'                                .Filter = tFilter
'                                .Sort = tSort
'                                .AbsolutePosition = nPos
'
'                                If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                                    nQtyCal = nQtyGet
'                                    nQtyGet = 0
'                                Else
'                                    nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                                    nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                                End If
'                                Select Case porsTemp.Fields("FTPddDisType").Value
'                                    Case "1" 'Per
'                                        cDealDiscAmt = cDelRegPrice * (cDealDiscPer / 100) * nQtyCal
'                                    Case "2"
'                                        cDealDiscAmt = .Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc)
'                                    Case "3"
'                                        cDealDiscAmt = (.Fields("FCSdtRegPrice").Value * nQtyCal) - (.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc))
'                                End Select
'                                cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
'                                cDisSumDisc = cDisSumDisc + cDealDiscAmt
'                                'nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                tFilter = .Filter
'                                tSort = .Sort
'                                Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1")  'cSource)
'                                Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'
''                                Call DE_GETcCallDeal(tDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cDelRegPrice * nQtyCal)
' '                               Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                nGrpOld = nGrp
'                                If nQtyGet <= 0 Then
'                                    nUncalGrp = nGrp
'                                End If
'                            Else
'                                If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                                    nQtyCal = nQtyGet
'                                    nQtyGet = 0
'                                Else
'                                    nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                                    nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
'                                End If
'                                Select Case .Fields("FTPddDisType").Value
'                                    Case "1" 'Per
'                                        cDealDiscAmt = cDelRegPrice * (cDealDiscPer / 100) * nQtyCal
'                                    Case "2"
'                                        cDealDiscAmt = .Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc)
'                                    Case "3"
'                                        cDealDiscAmt = (.Fields("FCSdtRegPrice").Value * nQtyCal) - (.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc))
'                                End Select
'                                cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
'                                cDisSumDisc = cDisSumDisc + cDealDiscAmt
'                                'nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                tFilter = .Filter
'                                tSort = .Sort
'                                Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1")  'cSource)
'                                Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'
''                                Call DE_GETcCallDeal(tDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cDelRegPrice * nQtyCal)
''                                Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                If nQtyGet <= 0 Then
'                                    nUncalGrp = nGrp
'                                End If
'                            End If
'                        End If
'                    End If
'                    .MoveNext
'                Loop
'            End If
'
'            If nQtyGet <= 0 Then
'                DE_DISb006 = True
'                acDE_SumDisc(pnPos) = cDisSumDisc + pcOldDisc
'            Else
'                DE_DISb006 = False
'            End If
'        End If
'
'    End With
'
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_DISb006(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 09-10-26
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double 'เก็บจำนวนซื้อ
    Dim cSourceLef As Double
    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
    Dim nMaxGrp As Integer
    Dim nI As Integer, nGrp As Integer
    Dim tFilter As String, tSort As String
    Dim cDisSumDisc As Double, cDisChkPrice As Double, cDisSumPrice As Double
    
    Dim nQtyGet As Integer, nUncalGrp As Integer, nDTSeqNo As Integer, nGrpOld As Integer, nQtyCal As Integer, nQtyGetOrg As Integer, nPos As Integer
    Dim tDealNo As String, tFreeFlag As String
    Dim cDelRegPrice As Double, cQtyAll As Double, cDelSaleAmt As Double, cDelDisAmt As Double, cDealDiscPer As Double, cAmtGet As Double
    Dim cAmtGetOrg As Double, cAmtCal As Double, cDealDiscAmt As Double, cSumCanDisc As Double
    Dim bChkround As Boolean
    
    On Error GoTo Err_Handler
    
    With porsTemp
        nMaxGrp = DE_GETnMaxGroup(atDE_AllDeal(pnKey), porsTemp)
        If Not (nMaxGrp = 2) Then Exit Function
        cDisSumDisc = 0
        cDisSumPrice = 0
        For nGrp = 1 To nMaxGrp
            bChkround = True
            tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacSkuGrp = " & nGrp
            .Filter = tFilter
            If Not (.BOF And .EOF) Then
                .MoveFirst
                '*Ao 09-10-22
                If nGrp = 2 Then
                    tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
                Else
'                    tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
                    tSort = "FTSkuStaMallEmpCard,FCSdtSalePricePmt"
                End If
                '_________________________________________________
                .Sort = tSort
                Do While Not .EOF
                    If nGrp = 1 Then
                        If .Fields("FCSdtQtyAll").Value > 0 Then
'                            cDelRegPrice = .Fields("FCSdtRegPrice").Value
                            cDelRegPrice = .Fields("FCSdtSalePricePmt").Value
                            cQtyAll = .Fields("FCSdtQtyAll").Value
                            cDelSaleAmt = cQtyAll * cDelRegPrice
                            cDelDisAmt = .Fields("FCDacDisAmt").Value
                            nDTSeqNo = .Fields("FNSdtSeqNo").Value
                            tFreeFlag = "1"
                            cDealDiscPer = 0
                            cDealDiscAmt = 0
                             
                            If bChkround Then
                                cAmtGet = .Fields("FCDacAmtBuy").Value
                                bChkround = False
                            End If
                            If cDelSaleAmt >= cAmtGet Then
                                cAmtCal = cAmtGet
                                cAmtGet = 0
                            Else
                                cAmtCal = cDelSaleAmt
                                cAmtGet = cAmtGet - cDelSaleAmt
                            End If
                           
                            nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal \ cDelRegPrice = cAmtCal / cDelRegPrice, 0, 1)
                            tFilter = .Filter
                            tSort = .Sort
                            Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "0") 'cSource)
                            Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                                                                
                            If cAmtGet <= 0 Then
                                Exit Do
                            End If
                        End If
                    Else        'If nGrp = 1 Then
                        If .Fields("FCSdtQtyAll").Value > 0 Then
                            cDelRegPrice = .Fields("FCSdtRegPrice").Value
                            nDTSeqNo = .Fields("FNSdtSeqNo").Value
                            tFreeFlag = .Fields("FTFreeFlag").Value
'                            Select Case .Fields("FTPddDisType").Value
'                                Case "1" 'Per
'                                    cDealDiscPer = .Fields("FCDacDisAmt").Value
'                                Case "2", "3"
'                                    cDealDiscPer = 0
'                                Case Else
'                                    cDealDiscPer = 0
'                            End Select
                            If bChkround Then
                                nQtyGet = .Fields("FNDacQtyGet").Value
                                bChkround = False
                            End If
                            Select Case .Fields("FTPddDisType").Value
                                Case "2", "3"
                                    nPos = .AbsolutePosition
                                    cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, nGrp) '*TW 51-06-27
                                    .Filter = tFilter
                                    .Sort = tSort
                                    .AbsolutePosition = nPos
                            End Select
                            If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                                nQtyCal = nQtyGet
                                nQtyGet = 0
                            Else
                                nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                                nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                            End If
                            Select Case .Fields("FTPddDisType").Value
                                Case "1" 'Per
                                    cDealDiscPer = .Fields("FCDacDisAmt").Value
                                    cDealDiscAmt = cDelRegPrice * (cDealDiscPer / 100) * nQtyCal
                                Case "2"
                                    cDealDiscPer = 0
                                    cDealDiscAmt = .Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc)
                                Case "3"
                                    cDealDiscPer = 0
                                    cDealDiscAmt = (cDelRegPrice * nQtyCal) - (.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc))
                                Case Else
                                    cDealDiscPer = 0
                                    cDealDiscAmt = 0
                            End Select
                                cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
                                cDisSumDisc = cDisSumDisc + cDealDiscAmt
                                tFilter = .Filter
                                tSort = .Sort
                                Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1")
                                Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                                If nQtyGet <= 0 Then
                                    Exit Do
                                End If
                        End If
                    End If
                    .MoveNext
                Loop
            Else
                DE_DISb006 = False
                Exit Function
            End If       'If Not (.BOF And .EOF) Then
        
        Next nGrp
        
        If nQtyGet <= 0 Then
            DE_DISb006 = True
            acDE_SumDisc(pnPos) = cDisSumDisc + pcOldDisc
        Else
            DE_DISb006 = False
        End If
    
    End With

    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_DISb007(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double 'เก็บจำนวนซื้อ
    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
    Dim nMaxGrp As Integer
    Dim nI As Integer
    Dim tFilter As String
    Dim tSort As String
    Dim cDisSumDisc As Double
    Dim cDisSumPrice As Double
    Dim nLevel As Integer
    Dim nMaxLev As Integer
    Dim cQtySumAll As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
    Dim cRegPrice As Double
    Dim nDTSeqNo As Integer
    Dim cDealDiscPer As Double
    Dim tFreeFlag As String
    Dim cDealDiscAmt As Double
    On Error GoTo Err_Handler
    With porsTemp
        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
        .Filter = tFilter
'        tSort = "FCSdtRegPrice DESC"
        tSort = "FNPddLevel "
        .Sort = tSort
        cSource = 0
        If Not (.BOF And .EOF) Then
'            nMaxLev = DE_GETnMaxGroup(atDE_AllDeal(pnKey), porsTemp, True)     '*52-05-29 ไม่ใช้ปรับเพื่อ ไม่ต้องเสียเวลาทำ
            .MoveFirst
            Do While Not .EOF
                If .Fields("FNPddLevel").Value = 1 Then
                    cQtySumAll = cQtySumAll + .Fields("FCSdtQtyAll").Value
                Else
                    Exit Do
                End If
                .MoveNext
            Loop
            .Sort = "FNPddLevel DESC"
            .MoveFirst
            Do While Not .EOF
                If .Fields("FNDacQtyBuy").Value <= cQtySumAll Then
                    nLevel = .Fields("FNPddLevel").Value
                    'nQtyGet = .Fields("FNDacQtyGet").Value
                    nQtyGet = cQtySumAll '*TW 51-08-08 'แก้ Comsheet 438
                    Exit Do
                Else
                    nLevel = 0
                End If
                .MoveNext
            Loop
            If nLevel = 0 Then
                DE_DISb007 = False
                Exit Function
            End If
            
            nQtyGetOrg = nQtyGet
            .Filter = tFilter & "AND FNPddLevel = " & nLevel
            .Sort = "FCSdtRegPrice DESC"
            .MoveFirst
            Do While Not .EOF
                If .Fields("FCSdtQtyAll").Value > 0 Then
                    cRegPrice = .Fields("FCSdtRegPrice").Value
                    nDTSeqNo = .Fields("FNSdtSeqNo").Value
                    cDealDiscPer = .Fields("FCDacDisAmt").Value
                    tFreeFlag = .Fields("FTFreeFlag").Value
                    '*TW 51-08-08
                    If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                        'nQtyCal = nQtyGet Old
                        nQtyCal = .Fields("FCSdtQtyAll").Value
                        nQtyGet = 0
                    Else
                        nQtyCal = .Fields("FCSdtQtyAll").Value
                        nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                    End If
'แก้ Comsheet 438
                    cDealDiscAmt = cRegPrice * (cDealDiscPer / 100) * nQtyCal
                    cDisSumDisc = cDisSumDisc + cDealDiscAmt
                    tFilter = .Filter
                    tSort = .Sort
                    Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1") 'cSource)
                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                    If nQtyGet <= 0 Then
                        Exit Do
                    End If
                End If
                .MoveNext
            Loop
            If nQtyGet > 0 Then
                DE_DISb007 = False
            Else
                DE_DISb007 = True
                acDE_SumDisc(pnPos) = cDisSumDisc + pcOldDisc
            End If
        End If
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_DISb008(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double 'เก็บจำนวนซื้อ
    Dim cSourceSum As Double 'เก็บจำนวนซื้อ
    Dim nMaxGrp As Integer
    Dim nI As Integer
    Dim tFilter As String
    Dim tSort As String
    Dim cDisSumDisc As Double
    Dim cDisSumPrice As Double
    Dim nLevel As Integer
    Dim nMaxLev As Integer
    Dim cQtySumAll As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
    Dim cRegPrice As Double
    Dim nDTSeqNo As Integer
    Dim cDealDiscPer As Double
    Dim tFreeFlag As String
    Dim cDealDiscAmt As Double
    Dim cAmtSumAll As Double, cAmtGet As Double, cAmtGetOrg As Double, cDelSaleAmt As Double, cAmtCal As Double
    Dim cQtyAll As Double
    On Error GoTo Err_Handler
    With porsTemp
        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
        .Filter = tFilter
        '*Ao 52-05-18
'        .Sort = "FCSdtRegPrice DESC"
        tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
        '______________________________________________
        .Sort = tSort
        .MoveFirst
        Do While Not .EOF
            If .Fields("FNPddLevel").Value = 1 Then
                cAmtSumAll = cAmtSumAll + .Fields("FCSdtQtyAll").Value * .Fields("FCSdtRegPrice").Value
            End If
            .MoveNext
        Loop
        
        .Sort = "FNPddLevel DESC"
        .MoveFirst
        Do While Not .EOF
            If .Fields("FCDacAmtBuy").Value <= cAmtSumAll Then
                nLevel = .Fields("FNPddLevel").Value
                'cAmtGet = .Fields("FCDacAmtBuy").Value
                cAmtGet = cAmtSumAll '*TW 51-08-08 'แก้ Comsheet 438
                Exit Do
            Else
                nLevel = 0
            End If
            .MoveNext
        Loop
        If nLevel = 0 Then Exit Function
        cAmtGetOrg = cAmtGet
        '*Ao 52-05-18
'        .Sort = "FCSdtRegPrice DESC"
        .Sort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
        '_______________________________________________
        .MoveFirst
        Do While Not .EOF
            If (porsTemp.Fields("FCSdtQtyAll").Value) > 0 And (porsTemp.Fields("FNPddLevel").Value = nLevel) Then
                cRegPrice = .Fields("FCSdtRegPrice").Value
                cQtyAll = .Fields("FCSdtQtyAll").Value
                cDelSaleAmt = cQtyAll * cRegPrice
                nDTSeqNo = .Fields("FNSdtSeqNo").Value
                cDealDiscPer = .Fields("FCDacDisAmt").Value
                tFreeFlag = "3"
                '*TW 51-08-08
                If cDelSaleAmt >= cAmtGet Then
                    'cAmtCal = cAmtGet 'Old
                    cAmtCal = cDelSaleAmt
                    cAmtGet = 0
                Else
                    cAmtCal = cDelSaleAmt
                    cAmtGet = cAmtGet - cDelSaleAmt
                End If
'แก้ Comsheet 438
                cDealDiscAmt = cDealDiscPer / 100 * cAmtCal
                cDisSumDisc = cDisSumDisc + cDealDiscAmt
'                nQtyCal = cAmtCal \ cRegPrice + IIf(cAmtCal Mod cRegPrice = 0, 0, 1)               '*Ao 09-10-20
                nQtyCal = cAmtCal \ cRegPrice + IIf(cAmtCal \ cRegPrice = cAmtCal / cRegPrice, 0, 1)
                tFilter = .Filter
                tSort = .Sort
                Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1") 'cSource)
                Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", cQtyAll - nQtyCal, tFilter, tSort)
                
'                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cAmtCal)
'                Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                If cAmtGet <= 0 Then
                    Exit Do
                End If
            End If
            .MoveNext
        Loop
        If cAmtGet > 0 Then
            DE_DISb008 = False
        Else
            DE_DISb008 = True
            acDE_SumDisc(pnPos) = cDisSumDisc + pcOldDisc
        End If
    
    End With

'    With porsTemp
'        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
'        .Filter = tFilter
'        tSort = "FCSdtRegPrice DESC"
'        .Sort = tSort
'        cSource = 0
'        If Not (.BOF And .EOF) Then
'            nMaxLev = DE_GETnMaxGroup(atDE_AllDeal(pnKey), porsTemp, True)
'            .MoveFirst
'            Do While Not .EOF
'                If .Fields("FNPddLevel").Value = 1 Then
'                    cAmtSumAll = cAmtSumAll + .Fields("FCSdtQtyAll").Value * .Fields("FCSdtRegPrice").Value
'                End If
'                .MoveNext
'            Loop
'            .Sort = "FNPddLevel DESC"
'            .MoveFirst
'            Do While Not .EOF
'                If .Fields("FCDacAmtBuy").Value <= cAmtSumAll Then
'                    nLevel = .Fields("FNPddLevel").Value
'                    Exit Do
'                Else
'                    nLevel = 0
'                End If
'                .MoveNext
'            Loop
'            If nLevel = 0 Then
'                DE_DISb008 = False
'                Exit Function
'            End If
'
'
'            .Filter = tFilter & " AND FNPddLevel = " & nLevel
'            .Sort = "FCSdtRegPrice DESC"
'            .MoveFirst
'            Do While Not .EOF
'                If .Fields("FCSdtQtyAll").Value > 0 Then
'                    cRegPrice = .Fields("FCSdtRegPrice").Value
'                    nDTSeqNo = .Fields("FNSdtSeqNo").Value
'                    cDealDiscPer = .Fields("FCDacDisAmt").Value
'                    tFreeFlag = .Fields("FTFreeFlag").Value
'                    cAmtGet = .Fields("FCDacAmtBuy").Value
'                    cAmtGetOrg = cAmtGet
'
'                    If cDelSaleAmt >= cAmtGet Then
'                        cAmtCal = cAmtGet
'                        cAmtGet = 0
'                    Else
'                        cAmtCal = cDelSaleAmt
'                        cAmtGet = cAmtGet - cDelSaleAmt
'                    End If
'                    cDealDiscAmt = cDealDiscPer / 100 * cAmtCal
'                    cDisSumDisc = cDisSumDisc + cDealDiscAmt
'                    nQtyCal = cAmtCal \ cRegPrice + IIf(cAmtCal Mod cRegPrice = 0, 0, 1)
'                    Call DE_SETxSkuUsed(pnPos, pnTime, nDTSeqNo, nQtyCal, "1") 'cSource)
'                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter & "AND FNPddLevel = " & nLevel, tSort)
'                    If cAmtGet <= 0 Then
'                        Exit Do
'                    End If
'                End If
'                .MoveNext
'            Loop
'            If cAmtGet > 0 Then
'                DE_DISb008 = False
'            Else
'                DE_DISb008 = True
'                acDE_SumDisc(pnPos) = cDisSumDisc + pcOldDisc
'            End If
'        End If
'    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_RSTxUpdateBySku(ByRef porsTemp As ADODB.Recordset, ByVal pnSeqNo As Integer, ByVal ptFieldName As String, ByVal pcValue As Double, ByVal ptFilter As String, ByVal ptSort As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nPos As Long
    On Error GoTo Err_Handler
    With porsTemp
        nPos = .AbsolutePosition
        .Filter = adFilterNone
        Do While Not .EOF
            If .Fields("FNSdtSeqNo").Value = pnSeqNo Then
                .Fields(ptFieldName).Value = pcValue
            End If
        '.Filter = "FNSdtSeqNo = " & pnSeqNo & ""
        'Call .Update(ptFieldName, pcValue)
            .MoveNext
        Loop
        .Filter = ptFilter
        .Sort = ptSort
        .AbsolutePosition = nPos
    End With
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function DE_GETnMaxGroup(ByVal ptDealNo As String, ByVal porsTemp As ADODB.Recordset, Optional ByVal pbLevel As Boolean = False) As Integer
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cMax As Integer
    On Error GoTo Err_Handler
    With porsTemp
        .Filter = "FTPdhDealNo = '" & ptDealNo & "'"
        If pbLevel Then
            Do While Not porsTemp.EOF
                If cMax < .Fields("FNPddLevel").Value Then
                    cMax = .Fields("FNPddLevel").Value
                End If
                .MoveNext
            Loop
        Else
            Do While Not porsTemp.EOF
                If cMax < .Fields("FNDacSkuGrp").Value Then
                    cMax = .Fields("FNDacSkuGrp").Value
                End If
                .MoveNext
            Loop
        End If
    End With
    DE_GETnMaxGroup = cMax
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_ARYxGet4RST(ByVal porsTemp As ADODB.Recordset, ByRef patArray() As String, ByRef pacDisc() As Double, ByRef pacRegPrice() As Double, ByRef patType() As String, ByRef patAllGrp() As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nCount As Long
    On Error GoTo Err_Handler
    ReDim patArray(0) As String 'DealNo
    ReDim pacDisc(0) As Double 'Disc
    ReDim pacRegPrice(0) As Double 'Reg Price
    ReDim patType(0) As String 'Deal Type
    ReDim patAllGrp(0) As String 'Deal Type
    porsTemp.Sort = "FCDacDis DESC"
    nCount = 0
    Do While Not porsTemp.EOF
        nCount = nCount + 1
        ReDim Preserve patArray(nCount)
        ReDim Preserve pacDisc(nCount)
        ReDim Preserve pacRegPrice(nCount)
        ReDim Preserve patType(nCount)
        ReDim Preserve patAllGrp(nCount)
        patArray(nCount) = porsTemp.Fields("FTPdhDealNo").Value
        pacDisc(nCount) = porsTemp.Fields("FCDacDis").Value
        pacRegPrice(nCount) = porsTemp.Fields("FCSdtRegPrice").Value
        patType(nCount) = porsTemp.Fields("FTDacDealType").Value
        patAllGrp(nCount) = porsTemp.Fields("FNDacSkuGrp").Value
        porsTemp.MoveNext
    Loop
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function DE_CHKbDelDeal(ByVal pcQtyAll As Double, ByVal pcQtyBuy As Double, ByVal ptDealType As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    DE_CHKbDelDeal = False
    
    Select Case ptDealType
        Case "001", "002", "003", "004", "005", "006", "007", "008", "0010", "0020", "0030", "0040"
            If pcQtyAll < pcQtyBuy Then
                DE_CHKbDelDeal = True
            End If

    End Select
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_GETxAllDeal(ByRef porsTemp As ADODB.Recordset, ByVal ptTblDT As String, ByVal ptTmnNo As String, ByVal ptRunning As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'*TW 51-06-20 หา Deal ที่มีโอกาสเกิดขึ้นได้จริง
'Para : porsTemp = Recordset ที่ใช้เก็บข้อมูลที่ Query ได้
'Para : ptTblDT = ตาราง DT ที่เรียกใช้ในปัจจุบัน
'Para : ptTmnNo = หมายเลขเครื่อง Terminal
'Para : ptRunning = หมายเลขบิล
'-----------------------------------------------------------------------------------------
    Dim tSql As String
    Dim nCmdTimeout As Long
    Dim tRepSta As String
    Dim nCountRound As Long
    Dim bKeep As Boolean
On Error GoTo Err_Handler
    '*KT 53-11-01 Preview Day
    If bVB_CNTrStaRegPreviewDay Then
        Call SP_INIxSwitchModePreviewDay(True)
    Else
       Call SP_INIxSwitchModePreviewDay(False)
    End If
    '*Ao 51-06-03
    tSql = "SELECT  DEL.FTPdhDealNo"
    tSql = tSql & vbCrLf & ",DEL.FNDacSkuGrp"
    tSql = tSql & vbCrLf & ",ISNULL(DT.FNSdtSeqNo,0)AS FNSdtSeqNo" 'FNSdtSeqNo
    '*KT 52-02-19 member price เปลี่ยนจากการให้ค่า FCSdtRegPrice จาก FCSdtB4DisChg มาเป็น FCSdtB4DisChgOrg
    '---------------------------------------------------------------------------------
'    If tVB_TransactionSub = tEN_TrnTaxExceptSale Then
'        tSql = tSql & vbCrLf & ",ISNULL(DT.FCSdtB4DisChg,0)*100/(100 + ISNULL(DT.FCSdtTax,0)) AS FCSdtRegPrice" 'FCSdtRegPrice'*TW 51-07-30 Comsheet 424 พี่ยูร Test เจอว่า ถ้าเป็น Tax Except Sale ต้องถอด Vat
'    Else
'        tSql = tSql & vbCrLf & ",ISNULL(DT.FCSdtB4DisChg,0) AS FCSdtRegPrice" 'FCSdtRegPrice'*TW 51-07-30 Comsheet 424 พี่ยูร Test เจอว่า ถ้าเป็น Tax Except Sale ต้องถอด Vat
'    End If
    If tVB_TransactionSub = tEN_TrnTaxExceptSale Then
        tSql = tSql & vbCrLf & ",ISNULL(DT.FCSdtB4DisChgOrg,0)*100/(100 + ISNULL(DT.FCSdtTax,0)) AS FCSdtRegPrice" 'FCSdtRegPrice'*TW 51-07-30 Comsheet 424 พี่ยูร Test เจอว่า ถ้าเป็น Tax Except Sale ต้องถอด Vat
    Else
        tSql = tSql & vbCrLf & ",ISNULL(DT.FCSdtB4DisChgOrg,0) AS FCSdtRegPrice" 'FCSdtRegPrice'*TW 51-07-30 Comsheet 424 พี่ยูร Test เจอว่า ถ้าเป็น Tax Except Sale ต้องถอด Vat
    End If
    '--------------------------------------------------------------------------------
    tSql = tSql & vbCrLf & ",ISNULL(DT.FCSdtQtyAll,0) AS FCSdtQtyAll" 'FCSdtQtyAll
    tSql = tSql & vbCrLf & ",DEL.FTDacSku" 'FTDacSku
    tSql = tSql & vbCrLf & ",ISNULL(DEL.FNDacQtyBuy,0) AS FNDacQtyBuy" 'FNDacQtyBuy
    tSql = tSql & vbCrLf & ",ISNULL(DEL.FNDacQtyGet,0) AS FNDacQtyGet" 'FNDacQtyGet
    tSql = tSql & vbCrLf & ",CASE WHEN DEL.FTPdhFreeFlag = '1' THEN DEL.FTDacDealType ELSE DEL.FTDacDealType + DEL.FTPdhFreeFlag END AS FTDacDealType" 'FTDacDealType
    tSql = tSql & vbCrLf & ",ISNULL(CASE WHEN FNDacQtyGet = 0 THEN 0 ELSE"
    tSql = tSql & vbCrLf & "    CASE DEL.FTPddDisType "
    tSql = tSql & vbCrLf & "        WHEN '1' " 'Per
    'tSql = tSql & vbCrLf & "            THEN DT.FCSdtRegPrice*DEL.FCDacDisAmt/100 " '*TW 51-07-17 Old ก่อนเปลี่ยน ใช้ Before
    If tVB_TransactionSub = tEN_TrnTaxExceptSale Then
        tSql = tSql & vbCrLf & "            THEN ISNULL(DT.FCSdtB4DisChg,0)*100/(100 + ISNULL(DT.FCSdtTax,0))*DEL.FCDacDisAmt/100 "
        tSql = tSql & vbCrLf & "        WHEN '2' " 'AMT
        tSql = tSql & vbCrLf & "            THEN ISNULL(FCDacDisAmt,0)"
        tSql = tSql & vbCrLf & "        WHEN '3' " 'New Price
        tSql = tSql & vbCrLf & "            THEN (ISNULL(DT.FCSdtB4DisChg,0)*100/(100 + ISNULL(DT.FCSdtTax,0)) * DEL.FNDacQtyGet) - ISNULL(FCDacDisAmt,0) "
        tSql = tSql & vbCrLf & "        WHEN '4' " 'Free
        tSql = tSql & vbCrLf & "            THEN (ISNULL(DT.FCSdtB4DisChg,0)*100/(100 + ISNULL(DT.FCSdtTax,0))) "
    Else
        tSql = tSql & vbCrLf & "            THEN DT.FCSdtB4DisChg*DEL.FCDacDisAmt/100 "
        tSql = tSql & vbCrLf & "        WHEN '2' " 'AMT
        tSql = tSql & vbCrLf & "            THEN ISNULL(FCDacDisAmt,0)"
        tSql = tSql & vbCrLf & "        WHEN '3' " 'New Price
        tSql = tSql & vbCrLf & "            THEN (DT.FCSdtB4DisChg * DEL.FNDacQtyGet) - ISNULL(FCDacDisAmt,0) "
        tSql = tSql & vbCrLf & "        WHEN '4' " 'Free
        tSql = tSql & vbCrLf & "            THEN (DT.FCSdtB4DisChg) "
    End If
    tSql = tSql & vbCrLf & "    END" 'New Price
    tSql = tSql & vbCrLf & "END,0) AS FCDacDis"
    tSql = tSql & vbCrLf & ",DEL.FTPddDisType AS FTPddDisType" 'FTPddDisType
    tSql = tSql & vbCrLf & ",ISNULL(DEL.FCDacDisAmt,0) AS FCDacDisAmt" 'FCDacDisAmt
    tSql = tSql & vbCrLf & ",DEL.FCDacAmtBuy,ISNULL(DEL.FNPddLevel,0) AS FNPddLevel"
    tSql = tSql & vbCrLf & ",DEL.FTEvnCode AS FTEvnCode"
    tSql = tSql & vbCrLf & ",CASE WHEN ISNULL(DEL.FNDacQtyBuy,0) > 0 THEN  "
    tSql = tSql & vbCrLf & "        CASE WHEN ISNULL(FNDacQtyGet,0) > 0 THEN "
    tSql = tSql & vbCrLf & "            '3' "
    tSql = tSql & vbCrLf & "        ELSE "
    tSql = tSql & vbCrLf & "            '1' "
    tSql = tSql & vbCrLf & "        END   "
    tSql = tSql & vbCrLf & "ELSE "
    tSql = tSql & vbCrLf & "        CASE WHEN ISNULL(FNDacQtyGet,0) > 0 THEN "
    tSql = tSql & vbCrLf & "            '2' "
    tSql = tSql & vbCrLf & "        ELSE "
    tSql = tSql & vbCrLf & "            CASE WHEN DEL.FTDacDealType = '005' THEN "
    tSql = tSql & vbCrLf & "                '3' "
    tSql = tSql & vbCrLf & "            ELSE "
    tSql = tSql & vbCrLf & "                '1' "
    tSql = tSql & vbCrLf & "            END   "
    tSql = tSql & vbCrLf & "        END "
    tSql = tSql & vbCrLf & "END  AS FTFreeFlag" 'FTFreeFlag
    tSql = tSql & vbCrLf & ",DT.FCSdtAmtDisAvi AS FCSdtAmtDisAvi"
    '*KT 52-02-19 member price
    '-----------------------------------------------------------------
    tSql = tSql & vbCrLf & ",DT.FCSdtMemberPrice AS FCSdtMemberPrice"
    tSql = tSql & vbCrLf & ",DT.FCSdtB4DisChgOrg AS FCSdtB4DisChgOrg"
    tSql = tSql & vbCrLf & ",DT.FTSdtStaRfd AS FTSdtStaRfd"
    tSql = tSql & vbCrLf & ",DT.FTSdtEventMemPrice AS FTSdtEventMemPrice"
    '----------------------------------------------------------------
    tSql = tSql & vbCrLf & ",DT.FTSkuStaMallEmpCard AS FTSkuStaMallEmpCard"
    tSql = tSql & vbCrLf & ",DT.FCSdtSalePricePmt AS FCSdtSalePricePmt"
    '___________________________________________________
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDealAct & " DEL"
    tSql = tSql & vbCrLf & " LEFT JOIN("
    tSql = tSql & vbCrLf & "SELECT FNSdtSeqNo,FCSdtB4DisChg,FCSdtQtyAll,FTSkuCode,FCSdtAmtDisAvi,FCSdtTax,FCSdtMemberPrice,FCSdtB4DisChgOrg,FTSdtStaRfd,FTSkuStaMallEmpCard,FTSdtEventMemPrice,FCSdtSalePricePmt"
    tSql = tSql & vbCrLf & " FROM " & ptTblDT 'FCSdtRegPrice'*TW 51-07-30 Comsheet 424 พี่ยูร Test เจอว่า ถ้าเป็น Tax Except Sale ต้องถอด Vat"
    tSql = tSql & vbCrLf & " WHERE " & tVB_CNWheBill
    tSql = tSql & vbCrLf & "AND ISNULL(FCOverrideAmt,0) = 0"            '*Ao 2010/12/16 Price Override  ไม่เอาสินค้าที่โดน Price Override แล้ว
    '*KT 53-11-01 Preview Day
    If bVB_CNTrStaRegPreviewDay Then
        tSql = tSql & vbCrLf & "AND ISNULL(FTSdtStaNotAlwPrv,'N') <>  'Y'"
    End If
    tSql = tSql & vbCrLf & "AND ISNULL(FTSdtCancelFlag,'0') <> '1') DT"
    tSql = tSql & vbCrLf & " ON DEL.FTDacSku = DT.FTSkuCode"
    tSql = tSql & vbCrLf & " WHERE DT.FCSdtQtyAll > 0"
    tSql = tSql & vbCrLf & " AND DEL.FDPhdEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*TW 51-07-23 MLSF-417 '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
    '*KT 54-06-07 Req.3,4 Deal & Free for Member
    If Not bVB_CNMemberPrice Then 'ถ้าเป็น Member ให้ได้ทุก Deal & Free  ถ้าลูกค้าทั่วไปจะไม่ได้ส่วนเฉพาะของ Member
         tSql = tSql & vbCrLf & "AND ISNULL(DEL.FTPdhMBFlag,'N') <>  'Y'"
    End If
    tSql = tSql & vbCrLf & " ORDER BY DEL.FTPdhDealNo,DEL.FNDacSkuGrp,FNSdtSeqNo"
    
    Do
        bKeep = False
        '*KT 52-11-28 กรณีที่มีการทำงานกับ TPSTDealAct เกิด Replicate ขึ้นได้ให้รอจนกว่าจะ  Replicate สำเร็จ หรือ เวลาเกิน Timeoff หรือเป็น Offline
        nCmdTimeout = ocnAll.CommandTimeout
        tRepSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
        Do While tRepSta = "0" And nCountRound <= nCmdTimeout And oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP)
            nCountRound = nCountRound + 1
            SleepEx 1000, 0
            tRepSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
        Loop
        tRepSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
        If tRepSta = "1" Then
            bKeep = True
        Else
            If MsgBox("Replicate is working." & vbCrLf & "To Wait click YES." & vbCrLf & "To Next click NO.", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
                bKeep = False
                nCountRound = 0
            Else
                bKeep = True
            End If
        End If
    Loop Until bKeep
    
    'Set porsTemp = ocnAll.Execute(tSql) '*KT 54-03-24 ป้องกันโปรแกรมหลุด
    Call SP_SQLvExecute(tSql, porsTemp, ocnAll)
    Exit Sub
    
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub
Private Sub DE_FILxSepDeal(ByVal porsSource As ADODB.Recordset, ByRef porsDE_DealSingle As ADODB.Recordset, ByRef orsDE_DealMulti As ADODB.Recordset)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'แยก deal ที่เป็น Multi กับ Single ออก
    Dim orsTemp As ADODB.Recordset
    Dim atGetSku() As String
    Dim atGetDealNo() As String
    Dim atGetDistSku() As String
    Dim nI As Long
    Dim nJ As Long
    Dim nCount As Long
    Dim bFound As Boolean
    On Error GoTo Err_Handler
    ReDim atGetSku(0) As String
    ReDim atGetDistSku(0) As String
    ReDim atGetDealNo(0) As String
    Dim cOldRegPrice As Double
    Call SP_RSTbClone2OfflineEx2(porsSource, orsTemp)
    Call DE_RSTxGetDistinct(orsTemp, "FTDacSku", atGetSku)
    For nI = LBound(atGetSku) To UBound(atGetSku)
        If atGetSku(nI) <> "" Then
            nCount = DE_RSTnCountRecord(orsTemp, "FTDacSku", atGetSku(nI))
            If nCount > 1 Then
                ReDim Preserve atGetDistSku(UBound(atGetDistSku) + 1)
                atGetDistSku(UBound(atGetDistSku)) = atGetSku(nI)
            End If
        End If
    Next nI
    'ได้ Array ที่เก็บ Sku ที่ซ้ำ Deal มาแล้ว atGetDistSku
    Call DE_RSTxGetDistinct(orsTemp, "FTPdhDealNo", atGetDealNo)
    'ได้ Array ที่เก็บ DealNo  มาแล้ว atGetDealNo
    Call SP_RSTbClone2OfflineEx2(orsTemp, orsDE_DealMulti, False)
    Call SP_RSTbClone2OfflineEx2(orsTemp, orsDE_DealSingle, False)
    For nI = LBound(atGetDealNo) To UBound(atGetDealNo)
        If atGetDealNo(nI) <> "" Then
            bFound = False
            For nJ = LBound(atGetDistSku) To UBound(atGetDistSku)
                If atGetDistSku(nJ) <> "" Then
                    orsTemp.Filter = "FTPdhDealNo = '" & atGetDealNo(nI) & "' AND FTDacSku = '" & atGetDistSku(nJ) & "'" ' AND FCSdtQtyAll <> 0"
                    If Not (orsTemp.BOF And orsTemp.EOF) Then
                        bFound = True
                        Exit For
                        'Multi ชัวร์
                    Else
                        bFound = False
                    End If
                End If
            Next nJ
            orsTemp.Filter = "FTPdhDealNo = '" & atGetDealNo(nI) & "'"
'            orsTemp.Sort = "FCSdtRegPrice DESC,FNSdtSeqNo"                     '*Ao 52-05-12
            orsTemp.Sort = "FNDacSkuGrp,FTSkuStaMallEmpCard,FCSdtRegPrice"
            cOldRegPrice = 0
            If bFound Then
                Do While Not orsTemp.EOF
                    If Not (cOldRegPrice = orsTemp.Fields("FCSdtRegPrice").Value) And Not (orsTemp.Fields("FCSdtRegPrice").Value = 0) Then
                        orsDE_DealMulti.AddNew
                        For nJ = 1 To orsTemp.Fields.Count
                            orsDE_DealMulti.Fields(nJ - 1) = orsTemp.Fields(nJ - 1).Value & vbNullString
                        Next nJ
                        cOldRegPrice = orsTemp.Fields("FCSdtRegPrice").Value
                    End If
                    orsTemp.MoveNext
                Loop
            Else
                orsDE_DealSingle.AddNew
                For nJ = 1 To orsTemp.Fields.Count
                    orsDE_DealSingle.Fields(nJ - 1) = orsTemp.Fields(nJ - 1).Value & vbNullString
                Next nJ
            End If
        End If
    Next nI
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Private Function DE_CHKbCheckSeqNo(ByVal porsTemp As ADODB.Recordset, ByVal pnSeqNo As String) As Boolean
''   Call:
''   Cmt:
''_______________________
'Dim orsTemp As ADODB.Recordset
'On Error GoTo Err_Handler
'    DE_CHKbCheckSeqNo = True
'    Set orsTemp = porsTemp.Clone
'    orsTemp.Filter = adFilterNone
'    If Not (orsTemp.BOF And orsTemp.EOF) Then
'        orsTemp.MoveFirst
'    End If
'    Do While Not orsTemp.EOF
'        If orsTemp.Fields("FNSdtSeqNo").Value = pnSeqNo Then
'            DE_CHKbCheckSeqNo = False
'            Exit Do
'        End If
'        orsTemp.MoveNext
'    Loop
'Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function
Private Sub DE_RSTxGetDistinct(ByVal porsTemp As ADODB.Recordset, ByVal ptFieldName As String, ByRef patTemp() As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim orsTemp As ADODB.Recordset
    Dim tOldData As String
    On Error GoTo Err_Handler
    '*Ao 52-03-17
'    Call SP_RSTbClone2Offline(porsTemp, orsTemp)
    Set orsTemp = porsTemp.Clone
    '______________________________________________
    orsTemp.Sort = ptFieldName
    orsTemp.MoveFirst
    Do While Not orsTemp.EOF
        If tOldData <> Trim(orsTemp.Fields(ptFieldName).Value) Then
            ReDim Preserve patTemp(UBound(patTemp) + 1)
            patTemp(UBound(patTemp)) = Trim(orsTemp.Fields(ptFieldName).Value)
            tOldData = Trim(orsTemp.Fields(ptFieldName).Value)
        End If
        orsTemp.MoveNext
    Loop
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function DE_RSTnCountRecord(ByVal porsTemp As ADODB.Recordset, ByVal ptFilterName As String, ByVal ptFilterValue As String) As Long
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nI As Integer
    Dim tOldDealNo As String
    Dim orsTemp As ADODB.Recordset
    On Error GoTo Err_Handler
'    Call SP_RSTbClone2Offline(porsTemp, orsTemp)
    '*Ao 52-03-16
    Set orsTemp = porsTemp.Clone
    '_____________________________
    orsTemp.Filter = ptFilterName & " = '" & ptFilterValue & "'"
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        orsTemp.Sort = "FTPdhDealNo"
        orsTemp.MoveFirst
        Do While Not orsTemp.EOF
            If tOldDealNo <> orsTemp.Fields("FTPdhDealNo").Value Then
                nI = nI + 1
                tOldDealNo = orsTemp.Fields("FTPdhDealNo").Value
            End If
            orsTemp.MoveNext
        Loop
        DE_RSTnCountRecord = nI
    Else
        DE_RSTnCountRecord = 0
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    DE_RSTnCountRecord = 0
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function DE_FILbSackNoDeal(ByVal porsTmp As ADODB.Recordset, ByRef porsDE_AFSack As ADODB.Recordset) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'เอามาเฉพาะ Deal ที่สามารถเกิดได้จริงจากบิลนี้ (องค์ประกอบไม่ครบ เอาออก)
    Dim tDealNo As String, tDealType As String, tSkuGrp As String
    Dim bDealComplete As Boolean
    Dim nPos As Long
    Dim cDisc As Double
    On Error GoTo Err_Handler
    If SP_RSTbClone2OfflineEx2(porsTmp, porsDE_AFSack) Then
        porsDE_AFSack.MoveFirst
        Do While Not porsDE_AFSack.EOF
            tDealNo = porsDE_AFSack.Fields("FTPdhDealNo").Value
            tDealType = porsDE_AFSack.Fields("FTDacDealType").Value
            cDisc = porsDE_AFSack.Fields("FCDacDis").Value
            nPos = porsDE_AFSack.AbsolutePosition
            'Set wDealTest.VSFlexGrid2.DataSource = porsDE_AFSack
            bDealComplete = DE_CHKbDealComplete(porsDE_AFSack, tDealNo, tDealType)
            porsDE_AFSack.Filter = adFilterNone                  '*Ao 52-03-17
'            porsDE_AFSack.Sort = "FTPdhDealNo"
            'Set wDealTest.VSFlexGrid2.DataSource = porsDE_AFSack
            If bDealComplete Then
                porsDE_AFSack.AbsolutePosition = nPos
                Do While Not porsDE_AFSack.EOF
                    If tDealNo = porsDE_AFSack.Fields("FTPdhDealNo").Value Then
                        porsDE_AFSack.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
            Else
                porsDE_AFSack.Filter = "FTPdhDealNo = '" & tDealNo & "'"
                Call DE_DELxDeal(porsDE_AFSack)
                porsDE_AFSack.Filter = adFilterNone
                If porsDE_AFSack.RecordCount > 0 Then
                    porsDE_AFSack.AbsolutePosition = nPos
                End If
            End If
            'Call DE_CHKxDealType(porsDE_AFSack, tDealNo, tDealType)
        Loop
         If porsDE_AFSack.RecordCount > 0 Then
            DE_FILbSackNoDeal = True
        Else
            DE_FILbSackNoDeal = False
        End If
    End If
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_DELxDeal0040(ByRef porsTemp As ADODB.Recordset, ByVal ptSkuGrpOld As Integer)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 09-09-18
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        If ptSkuGrpOld = porsTemp.Fields("FNDacSkuGrp").Value Then
            porsTemp.Delete
        End If
        porsTemp.MoveNext
    Loop
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DE_DELxDeal(ByRef porsTemp As ADODB.Recordset)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        porsTemp.Delete
        porsTemp.MoveNext
    Loop
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Private Function DE_CHKbDealCompleteAndCut(ByRef porsTmp As ADODB.Recordset, ByVal ptFilter As String, ByVal ptDealType As String, Optional ByVal ptSkuGrp As String, Optional ByRef pnRound As Integer) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim tSkuGrpOld As String, tSkuGrp As String
'    Dim cQtyAll As Double, cQtyBuy As Double, cRoundBuyTmp As Double, cRoundBuy As Double
'    Dim bChkDeal As Boolean, bCutAlready As Boolean
'    Dim nCalGrp As Integer
'
'    On Error GoTo Err_Handler
'
'    porsTmp.Filter = "FTPdhDealNo='" & ptFilter & "'"
'    porsTmp.Sort = "FNDacSkuGrp,FNSdtSeqNo"
'    tSkuGrpOld = "XXX"
'    bCutAlready = False
'    Do While Not porsTmp.EOF
'
'        tSkuGrp = porsTmp.Fields("FNDacSkuGrp").Value
'        Select Case ptDealType
'            Case "001"      '001ซื้อสินค้าในกลุ่มจำนวน มากกว่าหรือเท่ากับ n ชิ้น แล้วได้รับเงื่อนไข % Off, Amount Off, New Price
'                cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
'                If bCutAlready = False Then
'                    porsTmp.Fields("FCSdtQtyAll").Value = porsTmp.Fields("FCSdtQtyAll").Value - cQtyBuy
'                    bCutAlready = True
'                End If
'            Case "0010"
'                pnRound = DE_PRCnChkRoundFree(porsTmp, ptDealType)
'                DE_CHKbDealCompleteAndCut = False
'                Exit Do
'            Case "0020"
'                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
'                    If tSkuGrpOld <> "XXX" Then
'                        bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
'                        If bChkDeal Then
'                            Exit Function
'                        End If
'                    End If
'                    cQtyAll = 0
'                    cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
'                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                    tSkuGrpOld = tSkuGrp
'                Else
'                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                End If
'
'            Case "002"      '002ซื้อสินค้าด้วยจำนวนที่กำหนด เท่ากับ n ชิ้น แล้วได้รับเงื่อนไข   % Off, Amount Off, New Price
''                cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
''                If porsTmp.Fields("FCSdtQtyAll").Value > 0 Then
''                    If tSkuGrpOld = "XXX" Then
''                        cQtyAll = 0
''                        tSkuGrpOld = tSkuGrp
''                        If porsTmp.Fields("FCSdtQtyAll").Value > 1 Then
''                            cQtyAll = cQtyAll + cQtyBuy
''                            porsTmp.Fields("FCSdtQtyAll").Value = porsTmp.Fields("FCSdtQtyAll").Value - cQtyBuy
''                            bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
''                            If bChkDeal Then
''                                DE_CHKbDealCompleteAndCut = False
''                            Else
''                                DE_CHKbDealCompleteAndCut = True
''                            End If
''                            Exit Function
''                        Else
''                            cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
''                            porsTmp.Fields("FCSdtQtyAll").Value = porsTmp.Fields("FCSdtQtyAll").Value - 1
''                        End If
''                    Else
''                        If porsTmp.Fields("FCSdtQtyAll").Value > 1 Then
''                            cQtyAll = cQtyAll + cQtyBuy
''                            porsTmp.Fields("FCSdtQtyAll").Value = porsTmp.Fields("FCSdtQtyAll").Value - cQtyBuy
''                            bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
''                            If bChkDeal Then
''                                DE_CHKbDealCompleteAndCut = False
''                            Else
''                                DE_CHKbDealCompleteAndCut = True
''                            End If
''                            Exit Function
''                        Else
''                            cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
''                            porsTmp.Fields("FCSdtQtyAll").Value = porsTmp.Fields("FCSdtQtyAll").Value - 1
''                        End If
''                    End If
''                End If
''                If cQtyAll = cQtyBuy Then
''                    DE_CHKbDealCompleteAndCut = True
''                    Exit Function
''                End If
'                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
'                    If tSkuGrpOld <> "XXX" Then
'                        cRoundBuyTmp = cQtyAll \ cQtyBuy
'                        If cRoundBuyTmp < cRoundBuy Then
'                            cRoundBuy = cRoundBuyTmp
'                        Else
'                            If cRoundBuy = 0 Then
'                                cRoundBuy = cRoundBuyTmp
'                            End If
'                        End If
'                    End If
'                    cQtyAll = 0
'                    cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
'                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                    tSkuGrpOld = tSkuGrp
'                Else
'                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                End If
'            Case "003"
'                pnRound = DE_PRCnChkRoundDeal003(porsTmp, ptDealType)
'                DE_CHKbDealCompleteAndCut = False
'                Exit Function
'
'            Case "004", "0040"
''                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
''                    nCalGrp = nCalGrp + 1
''                    If tSkuGrpOld <> "XXX" Then
''                        bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
''                        If bChkDeal Then
''                            Exit Function
''                        End If
''                    End If
''                    cQtyAll = 0
''                    If tSkuGrp <> "3" Then
''                        cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
''                    Else
''                        cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
''                    End If
''                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
''                    tSkuGrpOld = tSkuGrp
''                Else
''                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
''                End If
'                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
'                    If tSkuGrpOld <> "XXX" Then
'                        cRoundBuyTmp = cQtyAll \ cQtyBuy
'                        If cRoundBuyTmp < cRoundBuy Then
'                            cRoundBuy = cRoundBuyTmp
'                        Else
'                            If cRoundBuy = 0 Then
'                                cRoundBuy = cRoundBuyTmp
'                            End If
'                        End If
'                    End If
'                    nQtyAll = 0
'                    If tSkuGrp = "1" Then
'                        nQtyBuy = .Fields("FNDacQtyBuy").Value
'                    Else
'                        nQtyBuy = .Fields("FNDacQtyGet").Value
'                    End If
'                    nQtyAll = nQtyAll + .Fields("FCSdtQtyAll").Value
'                    tSkuGrpOld = tSkuGrp
'                Else
'                    nQtyAll = nQtyAll + .Fields("FCSdtQtyAll").Value
'                End If
'            Case "005"          '005ซื้อสินค้าในกลุ่มที่ 1 ครบ nnn บาท แล้วได้รับส่วนลด ด้วยเงื่อนไข   % Off, Amount Off
'                cQtyBuy = porsTmp.Fields("FCDacAmtBuy").Value
'                cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
'            Case "006"          '006ซื้อสินค้าในกลุ่มที่ 1 ครบ nnn บาท แล้วได้รับสินค้ากลุ่มที่ 2 จำนวน n ชิ้น ด้วยเงื่อนไข   % Off, Amount Off
'                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
'                    nCalGrp = nCalGrp + 1
'                    If tSkuGrpOld = "XXX" Then 'ครั้งแรก
'                        If tSkuGrp = "1" Then
'                            cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
'                            cQtyBuy = porsTmp.Fields("FCDacAmtBuy").Value
'                        Else
'                            'ไม่น่าเกิดเหตุการณ์นี้
'                            Exit Function
'                        End If
'                    Else
'                        bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
'                        cQtyAll = 0
'                        If bChkDeal Then
'                            Exit Function
'                        End If
'                        If tSkuGrp <> "1" Then
'                            cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                            cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
'                        Else
'                            'ไม่น่าเกิดเหตุการณ์นี้
'                            Exit Function
'                        End If
'                    End If
'                    tSkuGrpOld = tSkuGrp
'                Else ' กลุ่มเดิม
'                    If tSkuGrp <> "1" Then
'                        cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                    Else
'                        cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
'                    End If
'                End If
'            Case "007"            '007ซื้อสินค้าในกลุ่มครบช่วงจำนวน nnn ชิ้น แล้วได้รับส่วนลดด้วยเงื่อนไข  % Off
'                If porsTmp.Fields("FNPddLevel").Value = 1 Then
'                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                    cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
'                End If
'            Case "008"
'                If porsTmp.Fields("FNPddLevel").Value = 1 Then
'                    cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
'                    cQtyBuy = porsTmp.Fields("FCDacAmtBuy").Value
'                End If
'
'        End Select
'        porsTmp.MoveNext
'    Loop
'        Select Case ptDealType
'            Case "001" '001ซื้อสินค้าในกลุ่มจำนวน มากกว่าหรือเท่ากับ n ชิ้น แล้วได้รับเงื่อนไข % Off, Amount Off, New Price
'            Case "0010"
'                Exit Function
'            Case "0020"
'            Case "002", "004"
'                cRoundBuyTmp = cQtyAll \ cQtyBuy            '002ซื้อสินค้าด้วยจำนวนที่กำหนด เท่ากับ n ชิ้น แล้วได้รับเงื่อนไข   % Off, Amount Off, New Price
'                If cRoundBuyTmp < cRoundBuy Then
'                    cRoundBuy = cRoundBuyTmp
'                Else
'                    If cRoundBuy = 0 Then
'                        cRoundBuy = cRoundBuyTmp
'                    End If
'                End If
'                pnRound = CInt(cRoundBuy)
'                DE_CHKbDealCompleteAndCut = False
'                Exit Function
'            Case "003", "0030"
'               Exit Function
''            Case "004", "0040"
''                If nCalGrp <> 3 Then Exit Function
'            Case "005"          '005ซื้อสินค้าในกลุ่มที่ 1 ครบ nnn บาท แล้วได้รับส่วนลด ด้วยเงื่อนไข   % Off, Amount Off
'            Case "006"          '006ซื้อสินค้าในกลุ่มที่ 1 ครบ nnn บาท แล้วได้รับสินค้ากลุ่มที่ 2 จำนวน n ชิ้น ด้วยเงื่อนไข   % Off, Amount Off
'                    If nCalGrp <> 2 Then Exit Function
'            Case "007"            '007ซื้อสินค้าในกลุ่มครบช่วงจำนวน nnn ชิ้น แล้วได้รับส่วนลดด้วยเงื่อนไข  % Off
'            Case "008"
'        End Select
'
'    bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
'    If bChkDeal Then
'        Exit Function
'    End If
'    DE_CHKbDealCompleteAndCut = True
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

'Private Function DE_CHKbFree0020(ByVal porsTmp As ADODB.Recordset, ByVal ptSKUCode As String) As Boolean
'    Dim orsTemp As ADODB.Recordset
'    Dim bReturn As Boolean
'    Dim nCount As Integer, nPos As Integer
'
'    On Error GoTo Err_Handler
'
'    DE_CHKbFree0020 = False
'    nPos = porsTmp.AbsolutePosition
'    Call SP_RSTbClone2OfflineEx2(porsTmp, orsTemp)
'    orsTemp.Filter = "FTDacSku ='" & ptSKUCode & "' AND FNDacQtyBuy > 0 "
'    If orsTemp.RecordCount > 0 Then
'        nCount = 1
'    Else
'        nCount = 0
'    End If
'
'    If nCount = 0 Then
'        DE_CHKbFree0020 = False
'        porsTmp.AbsolutePosition = nPos
'        Exit Function
'    End If
'
'    orsTemp.Filter = adFilterNone
'    orsTemp.Filter = "FTDacSku ='" & ptSKUCode & "' AND FNDacQtyGet > 0 "
'    If orsTemp.RecordCount > 0 Then
'        DE_CHKbFree0020 = True
'    Else
'        DE_CHKbFree0020 = False
'    End If
'    porsTmp.AbsolutePosition = nPos
'    Exit Function
'Err_Handler:
'    porsTmp.AbsolutePosition = nPos
'    DE_CHKbFree0020 = False
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_CHKbDealCompleteAndCut(ByRef porsTmp As ADODB.Recordset, ByVal ptFilter As String, ByVal ptDealType As String, Optional ByVal ptSkuGrp As String, Optional ByRef pnRound As Integer) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 09-11-19
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSkuGrpOld As String, tSkuGrp As String
    Dim cQtyAll As Double, cQtyBuy As Double, cRoundBuyTmp As Double, cRoundBuy As Double
    Dim bChkDeal As Boolean, bCutAlready As Boolean
    Dim nCalGrp As Integer
    Dim tSKUCode As String
    
    On Error GoTo Err_Handler
    
    porsTmp.Filter = "FTPdhDealNo='" & ptFilter & "'"
    porsTmp.Sort = "FNDacSkuGrp,FNSdtSeqNo"
    tSkuGrpOld = "XXX"
'    bCutAlready = False
    
    Do While Not porsTmp.EOF
        
        tSkuGrp = porsTmp.Fields("FNDacSkuGrp").Value
        Select Case ptDealType
'            Case "001"      'ซื้อสินค้ากลุ่มที่ 1(Buy) จำนวน มากกว่าหรือเท่ากับ N ชิ้น แล้วได้รับเงื่อนไขพิเศษ 1.%Off, 2.Amount Off, 3.New Price
'                cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
'                If bCutAlready = False Then
'                    porsTmp.Fields("FCSdtQtyAll").Value = porsTmp.Fields("FCSdtQtyAll").Value - cQtyBuy
'                    bCutAlready = True
'                End If
            Case "001"
                cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
            Case "0010"             'ซื้อสินค้าในกลุ่มที่ 1(Buy & Free) จำนวน มากกว่าหรือเท่ากับ N ชิ้น แล้วได้รับสินค้านั้นจำนวน N ชิ้น ฟรี
                pnRound = DE_PRCnChkRoundFree(porsTmp, ptDealType)
                Exit Do
            Case "002"      'ซื้อสินค้าด้วยจำนวนที่กำหนด เท่ากับ n ชิ้น แล้วได้รับเงื่อนไข   % Off, Amount Off, New Price
                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
                    If tSkuGrpOld <> "XXX" Then
                        cRoundBuyTmp = cQtyAll \ cQtyBuy
                        If cRoundBuyTmp < cRoundBuy Then
                            cRoundBuy = cRoundBuyTmp
                        Else
                            If cRoundBuy = 0 Then
                                cRoundBuy = cRoundBuyTmp
                            End If
                        End If
                    End If
                    cQtyAll = 0
                    cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                    tSkuGrpOld = tSkuGrp
                Else
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                End If
            Case "0020"
                'CommSheet ML-PRV-055 Not free item on Free Type2.xls
                '*Ao 2012/08/22  4.0.23 CommSheet ML-PH4-R3-035 Wrong Free Item for Free Type2.xls
'                pnRound = DE_PRCnChkRun0020(porsTmp, ptDealType)
'                pnRound = DE_PRCnChkRound0020(porsTmp, ptDealType)
                pnRound = DE_PRCnChkRun0020Ex(porsTmp, ptDealType)
                Exit Do
            Case "003"          'ซื้อสินค้าในกลุ่มที่ 1(Buy) จำนวน มากกว่าหรือเท่ากับ N ชิ้น แล้วได้รับสินค้าในกลุ่มที่ 2(Get) จำนวน N ชิ้น ด้วยเงื่อนไขพิเศษ 1.%Off, 2.Amount Off, 3.New Price
'                pnRound = DE_PRCnChkRoundDeal003(porsTmp, ptDealType)
                pnRound = DE_PRCnChkRound0020(porsTmp, ptDealType)
                Exit Do
            Case "0030"
                pnRound = DE_PRCnChkRound0020(porsTmp, ptDealType)          '*Aoo 2012/09/06  CommSheet ST-FL-040-SF POST Not responding  afetr press Total Key.xls
                Exit Do
            Case "004"         '004ซื้อสินค้าในกลุ่มที่ 1(Buy) จำนวน มากกว่าหรือเท่ากับ N ชิ้น พร้อมกับสินค้าในกลุ่มที่ 2(Buy) จำนวน มากกว่าหรือเท่ากับ N ชิ้น แล้วได้รับสินค้าในกลุ่มที่ 3(Get) จำนวน N ชิ้น ด้วยเงื่อนไขพิเศษ 1.%Off, 2.Amount Off, 3.New Price
                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
                    If tSkuGrpOld <> "XXX" Then
                        cRoundBuyTmp = cQtyAll \ cQtyBuy
                        If cRoundBuyTmp < cRoundBuy Then
                            cRoundBuy = cRoundBuyTmp
                        Else
                            If cRoundBuy = 0 Then
                                cRoundBuy = cRoundBuyTmp
                            End If
                        End If
                    End If
                    cQtyAll = 0
                    If CInt(porsTmp.Fields("FNDacQtyBuy").Value) = CInt(0) Then
                        cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
                    Else
                        cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
                    End If
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                    tSkuGrpOld = tSkuGrp
                Else
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                End If
            Case "0040"                                                                                      ''004ซื้อสินค้าในกลุ่มที่ 1(Buy) จำนวน มากกว่าหรือเท่ากับ N ชิ้น พร้อมกับสินค้าในกลุ่มที่ 2(Buy) จำนวน มากกว่าหรือเท่ากับ N ชิ้น แล้วได้รับสินค้าในกลุ่มที่ 3(Get) จำนวน N ชิ้น ด้วยเงื่อนไขพิเศษ 1.%Off, 2.Amount Off, 3.New Price
                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
                    If tSkuGrpOld <> "XXX" Then
                        cRoundBuyTmp = cQtyAll \ cQtyBuy
                        If cRoundBuyTmp < cRoundBuy Then
                            cRoundBuy = cRoundBuyTmp
                        Else
                            If cRoundBuy = 0 Then
                                cRoundBuy = cRoundBuyTmp
                            End If
                        End If
                    End If
                    cQtyAll = 0
                    If CInt(porsTmp.Fields("FNDacQtyBuy").Value) = CInt(0) Then               'Ao 10-02-04
                        cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
                    Else
                        cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
                    End If
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                    tSkuGrpOld = tSkuGrp
                Else
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                End If
            Case "005"          '005ซื้อสินค้าในกลุ่มที่ 1 ครบ nnn บาท แล้วได้รับส่วนลด ด้วยเงื่อนไข   % Off, Amount Off
                cQtyBuy = porsTmp.Fields("FCDacAmtBuy").Value
                cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
            Case "006"          '006ซื้อสินค้าในกลุ่มที่ 1 ครบ nnn บาท แล้วได้รับสินค้ากลุ่มที่ 2 จำนวน n ชิ้น ด้วยเงื่อนไข   % Off, Amount Off
                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
                    If tSkuGrpOld = "XXX" Then 'ครั้งแรก
                        If tSkuGrp = "1" Then
                            cRoundBuy = cQtyAll \ cQtyBuy
                        Else
                            cRoundBuyTmp = cQtyAll \ cQtyBuy
                        End If
                    End If
                    cQtyAll = 0
                    If tSkuGrp = "1" Then
'                        cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
                        cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtSalePricePmt").Value * porsTmp.Fields("FCSdtQtyAll").Value)
                        cQtyBuy = porsTmp.Fields("FCDacAmtBuy").Value
                    Else
                        cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                        cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
                    End If
                    tSkuGrpOld = tSkuGrp
                Else ' กลุ่มเดิม
                    If tSkuGrp = "1" Then
'                        cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
                        cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtSalePricePmt").Value * porsTmp.Fields("FCSdtQtyAll").Value)
                    Else
                        cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                    End If
                End If
            Case "007"         'ซื้อสินค้าในกลุ่มตามช่วงครบ NNN ชิ้น แล้วได้รับส่วนลดด้วยเงื่อนไขพิเษษ 1.%Off
'                If porsTmp.Fields("FNPddLevel").Value = 1 Then
'                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                    cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
'                End If
            Case "008"         'ซื้อสินค้าในกลุ่มตามช่วงครบ NNN บาท แล้วได้รับส่วนลดด้วยเงื่อนไขพิเศษ 1.%Off
'                If porsTmp.Fields("FNPddLevel").Value = 1 Then
'                    cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
'                    cQtyBuy = porsTmp.Fields("FCDacAmtBuy").Value
'                End If
        End Select
        porsTmp.MoveNext
    Loop
        Select Case ptDealType
            Case "001", "005"
                cRoundBuy = cQtyAll \ cQtyBuy
                pnRound = CInt(cRoundBuy)
'                DE_CHKbDealCompleteAndCut = False
'                Exit Function
            Case "003", "0010"
'                DE_CHKbDealCompleteAndCut = False
'                Exit Function
'            Case "002", "004", "006", "0030", "0040"
            Case "002", "004", "006", "0040"
                cRoundBuyTmp = cQtyAll \ cQtyBuy
                If cRoundBuyTmp < cRoundBuy Then
                    cRoundBuy = cRoundBuyTmp
                Else
                    If cRoundBuy = 0 Then
                        cRoundBuy = cRoundBuyTmp
                    End If
                End If
                pnRound = CInt(cRoundBuy)
'                DE_CHKbDealCompleteAndCut = False
'                Exit Function
'            Case "005"
'                cRoundBuy = cQtyAll \ cQtyBuy
'                pnRound = CInt(cRoundBuy)
'                DE_CHKbDealCompleteAndCut = False
'                Exit Function
            Case "0020"
            Case "0030"
            Case "007"
            Case "008"
        End Select
    
'    bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
'    If bChkDeal Then
'        Exit Function
'    End If
    DE_CHKbDealCompleteAndCut = False
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Function DE_CHKnDealType(ByVal porsTemp As ADODB.Recordset, ByVal ptDealType) As Integer
''__________________________________
''   Call:
''   Cmt:'*Ao 52-05-29
''__________________________________
''Dim porsTemp As ADODB.Recordset
'Dim cQtySumAll As Double
'Dim nLevel As Integer
'On Error GoTo Err_Handler
'    If ptDealType = "007" Then
''        Set porsTmp = porsClone.Clone
'        porsTemp.Filter = "FNPddLevel = 1"
'        porsTemp.MoveFirst
'        Do While Not porsTemp.EOF
'            If porsTemp.Fields("FNPddLevel").Value = 1 Then
'                cQtySumAll = cQtySumAll + porsTemp.Fields("FCSdtQtyAll").Value
'            End If
'            porsTemp.MoveNext
'        Loop
'        porsTemp.Filter = adFilterNone
'        porsTemp.Sort = "FNSdtSeqNo,FNPddLevel DESC"
'        porsTemp.MoveFirst
'        Do While Not porsTemp.EOF
'            If porsTemp.Fields("FNDacQtyGet").Value <= cQtySumAll Then
'                nLevel = porsTemp.Fields("FNPddLevel").Value
'                Exit Do
'            End If
'            porsTemp.MoveNext
'        Loop
'        DE_CHKnDealType = nLevel
'    End If
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function
'Private Function DE_CHKbDeal0040(ByVal porsDeal As ADODB.Recordset) As Boolean
''________________________________________________
''   Call
''   Cmt: '*Ao 09-09-21
''________________________________________________
'    Dim orsTmp As ADODB.Recordset
'    Dim tGrp As String, tSkuGrpOld As String
'    On Error GoTo Err_Handler
'
'    Call SP_RSTbClone2Offline(porsDeal, orsTemp)
'    tGrp = "1"
'    porsTmp.MoveFirst
'    tSkuGrp = porsTmp.Fields("FNDacSkuGrp").Value
'    Do While Not porsTmp.EOF
'
'    If (tGrp = tSkuGrp) Then
''        tSkuGrpOld = tSkuGrp
'        porsTmp.MoveNext
'        tSkuGrp = porsTmp.Fields("FNDacSkuGrp").Value
'        If Not (tGrp = tSkuGrp) Then
'             tGrp = Str(Val(tGrp) + 1)
'        End If
'    Else
'        DE_CHKbDeal0040 = True
'    End If
'    Loop
'    DE_CHKbDeal0040 = False
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'    DE_CHKbDeal0040 = True
'End Function
Private Function DE_CHKbDealComplete(ByVal porsClone As ADODB.Recordset, ByVal ptFilter As String, ByVal ptDealType As String, Optional ByVal ptSkuGrp As String) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSkuGrpOld As String, tSkuGrp As String
    Dim cQtyAll As Double, cQtyBuy As Double, cQtyGet As Double
    Dim nCalGrp As Integer, nLevel As Integer, nPos As Integer
    Dim porsTmp As ADODB.Recordset, orsMaxGrp As ADODB.Recordset
    Dim bChkDeal As Boolean, bChkGrp As Boolean
    Dim tSql As String
    Dim nMaxGrp As Integer
    Dim nCountRound As Long
    Dim tRepSta As String
    Dim nCmdTimeout As Long
    On Error GoTo Err_Handler
    '*Ao 52-03-09
    Set porsTmp = porsClone.Clone
    '________________________________
    porsTmp.Filter = "FTPdhDealNo='" & ptFilter & "'"
    porsTmp.Sort = "FNDacSkuGrp"
    tSkuGrpOld = "XXX"
    porsTmp.MoveFirst
    
    nCalGrp = 0
    
    Select Case ptDealType
        Case "002", "004", "0020", "0040"         'Deal Type 2,Deal Type 4,Free Type 2,Free Type 4
            tSql = "SELECT COUNT(DISTINCT FNDacSkuGrp) AS FNMaxGroup FROM " & tVB_CNTblDealAct & " WHERE FTPdhDealNo='" & ptFilter & "'"
            
            '*KT 52-11-28 กรณีที่มีการทำงานกับ TPSTDealAct เกิด Replicate ขึ้นได้ให้รอจนกว่าจะ  Replicate สำเร็จ หรือ เวลาเกิน Timeoff หรือเป็น Offline
            nCmdTimeout = ocnAll.CommandTimeout
            tRepSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
            Do While tRepSta = "0" And nCountRound <= nCmdTimeout And oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP)
                nCountRound = nCountRound + 1
                SleepEx 1000, 0
                tRepSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
            Loop
           Call SP_SQLvExecute(tSql, orsMaxGrp, ocnAll)
            'Set orsMaxGrp = ocnAll.Execute(tSql) '*KT 54-03-24 ป้องกันโปรแกรมหลุด
            If Not (orsMaxGrp.BOF And orsMaxGrp.EOF) Then
                nMaxGrp = orsMaxGrp.Fields("FNMaxGroup").Value
            End If
    End Select
    
    Do While Not porsTmp.EOF
        tSkuGrp = porsTmp.Fields("FNDacSkuGrp").Value
        
        Select Case ptDealType
            Case "001", "0010" '001ซื้อสินค้าในกลุ่มจำนวน มากกว่าหรือเท่ากับ n ชิ้น แล้วได้รับเงื่อนไข % Off, Amount Off, New Price             'Deal Type1,Free Type 1
                cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
                cQtyGet = porsTmp.Fields("FNDacQtyGet").Value
            Case "002"      '002ซื้อสินค้าด้วยจำนวนที่กำหนด เท่ากับ n ชิ้น แล้วได้รับเงื่อนไข   % Off, Amount Off, New Price        'Deal Type 2
                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
                    nCalGrp = nCalGrp + 1
                    If tSkuGrpOld <> "XXX" Then
                        bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
                        If bChkDeal Then
                            Exit Function
                        End If
                    End If
                    cQtyAll = 0
                    cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                    tSkuGrpOld = tSkuGrp
                Else
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                End If
            Case "0020", "0040"        'Free Type 2,'Free Type 4
                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
                    nCalGrp = nCalGrp + 1
                    If tSkuGrpOld <> "XXX" Then
                        bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
                        If bChkDeal Then
                            Exit Function
                        End If
                    End If
                    cQtyAll = 0
                    If CInt(porsTmp.Fields("FNDacQtyBuy").Value) = Int(0) Then
                        cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
                    Else
                        cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
                    End If
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                    tSkuGrpOld = tSkuGrp
                Else
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                End If
            Case "003", "0030"
                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
                    nCalGrp = nCalGrp + 1
                    If tSkuGrpOld <> "XXX" Then
                        bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
                        If bChkDeal Then
                            Exit Function
                        End If
                    End If
                    cQtyAll = 0
                    If tSkuGrp = "1" Then
                        cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
                    Else
                        cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
                    End If
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                    tSkuGrpOld = tSkuGrp
                Else
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                End If
            Case "004"
                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
                    nCalGrp = nCalGrp + 1
                    If tSkuGrpOld <> "XXX" Then
                        bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
                        If bChkDeal Then
                            Exit Function
                        End If
                    End If
                    cQtyAll = 0
'                    Select Case Trim(porsTmp.Fields("FTPddDisType").Value)
'                        Case "1", "2", "3"
'                            cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
'                        Case Else
'                            cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
'                    End Select
                    If CInt(porsTmp.Fields("FNDacQtyBuy").Value) = CInt(0) Then
                        cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
                    Else
                        cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
                    End If
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                    tSkuGrpOld = tSkuGrp
                Else
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                End If
'            Case "0040"     '*Ao 09-09-17
'                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
'                    nCalGrp = nCalGrp + 1
'                    If tSkuGrpOld <> "XXX" Then
'                        bchkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
'                        If bchkDeal Then
'                            Exit Function
'                        End If
'                    End If
'                    cQtyAll = 0
'                    If porsTmp.Fields("FNDacQtyGet").Value = 0 Then
'                        cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
'                    Else
'                        cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
'                    End If
'                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                    tSkuGrpOld = tSkuGrp
'                Else
'                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
'                End If
            Case "005"          '005ซื้อสินค้าในกลุ่มที่ 1 ครบ nnn บาท แล้วได้รับส่วนลด ด้วยเงื่อนไข   % Off, Amount Off
                cQtyBuy = porsTmp.Fields("FCDacAmtBuy").Value
                cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
            Case "006"          '006ซื้อสินค้าในกลุ่มที่ 1 ครบ nnn บาท แล้วได้รับสินค้ากลุ่มที่ 2 จำนวน n ชิ้น ด้วยเงื่อนไข   % Off, Amount Off
                If tSkuGrpOld <> tSkuGrp Then 'เริ่มกลุ่มใหม่
                    nCalGrp = nCalGrp + 1
                    If tSkuGrpOld = "XXX" Then 'ครั้งแรก
                        If tSkuGrp = "1" Then
'                            cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
                            cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtSalePricePmt").Value * porsTmp.Fields("FCSdtQtyAll").Value)       '*Ao ต้องคำนวนจากราคาที่ได้ Promo
                            cQtyBuy = porsTmp.Fields("FCDacAmtBuy").Value
                        Else
                            'ไม่น่าเกิดเหตุการณ์นี้
                            Exit Function
                        End If
                    Else
                        bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
                        cQtyAll = 0
                        If bChkDeal Then
                            Exit Function
                        End If
                        If tSkuGrp <> "1" Then
                            cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                            cQtyBuy = porsTmp.Fields("FNDacQtyGet").Value
                        Else
                            'ไม่น่าเกิดเหตุการณ์นี้
                            Exit Function
                        End If
                    End If
                    tSkuGrpOld = tSkuGrp
                Else ' กลุ่มเดิม
                    If tSkuGrp <> "1" Then
                        cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                    Else
'                        cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
                        cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtSalePricePmt").Value * porsTmp.Fields("FCSdtQtyAll").Value)
                    End If
                End If
            Case "007"            '007ซื้อสินค้าในกลุ่มครบช่วงจำนวน nnn ชิ้น แล้วได้รับส่วนลดด้วยเงื่อนไข  % Off
                If porsTmp.Fields("FNPddLevel").Value = 1 Then
                    cQtyAll = cQtyAll + porsTmp.Fields("FCSdtQtyAll").Value
                    cQtyBuy = porsTmp.Fields("FNDacQtyBuy").Value
                End If
            Case "008"
                If porsTmp.Fields("FNPddLevel").Value = 1 Then
                    cQtyAll = cQtyAll + (porsTmp.Fields("FCSdtRegPrice").Value * porsTmp.Fields("FCSdtQtyAll").Value)
                    cQtyBuy = porsTmp.Fields("FCDacAmtBuy").Value
                End If

        End Select
        porsTmp.MoveNext
    Loop
        Select Case ptDealType
            Case "001"  '001ซื้อสินค้าในกลุ่มจำนวน มากกว่าหรือเท่ากับ n ชิ้น แล้วได้รับเงื่อนไข % Off, Amount Off, New Price
            Case "0010"
                cQtyBuy = cQtyBuy + cQtyGet
            Case "002", "0020", "0040"
                If Not (nCalGrp = nMaxGrp) Then Exit Function
            Case "003", "0030"
                If nCalGrp <> 2 Then Exit Function
            Case "004"
'                If nCalGrp <> 3 Then Exit Function
                If Not (nCalGrp = nMaxGrp) Then Exit Function
            Case "005"          '005ซื้อสินค้าในกลุ่มที่ 1 ครบ nnn บาท แล้วได้รับส่วนลด ด้วยเงื่อนไข   % Off, Amount Off
            Case "006"          '006ซื้อสินค้าในกลุ่มที่ 1 ครบ nnn บาท แล้วได้รับสินค้ากลุ่มที่ 2 จำนวน n ชิ้น ด้วยเงื่อนไข   % Off, Amount Off
                If nCalGrp <> 2 Then Exit Function
            Case "007"            '007ซื้อสินค้าในกลุ่มครบช่วงจำนวน nnn ชิ้น แล้วได้รับส่วนลดด้วยเงื่อนไข  % Off
            Case "008"
        End Select
    
    bChkDeal = DE_CHKbDelDeal(cQtyAll, cQtyBuy, ptDealType)
    
'    If bchkDeal Then
'        If ptDealType = "0040" Then             '*Ao 09-09-18       Free Type 4
'            If pbChkDeal0040 = True Then
'                Call DE_DELxDeal0040(porsTmp, tSkuGrpOld)    'Get ไม่ครบลบทิ้ง
'            End If
'            If bChk0040 = True Then
'                DE_CHKbDealComplete = True
'            Else
'                DE_CHKbDealComplete = False
'            End If
'            Exit Function
'        Else
'            Exit Function
'        End If
'    End If
    
    If bChkDeal Then
        Exit Function
    End If
    
    DE_CHKbDealComplete = True
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_DELtCheckType(ByVal poRs As ADODB.Recordset, ByVal ptField As String, Optional pvRetOld As Variant) As String
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'________________________________
'   Call:
'   Cmt:    Ao51-06-12
'________________________________
    Dim tTemp As String
    Dim cTemp As Double
    On Error GoTo Err_Handler
    Select Case UCase(Mid(ptField, 2, 1))
        Case "C", "N"
            cTemp = SP_FEDvChkNumeric(poRs, ptField)
            DE_DELtCheckType = "," & cTemp
        Case "T"
            tTemp = SP_STRtSiggleCode(SP_FEDtChkString(poRs, ptField))              '*Ao 2010/07/08 PH3-ML-SF-124 Wrong Sale Amount and QTY of Free Item.xls
            DE_DELtCheckType = ",'" & tTemp & "'"
        Case "D"
            tTemp = UT_tStringDate(CDate(SP_FEDtChkString(poRs, ptField)))
            DE_DELtCheckType = ",'" & tTemp & "'"
    End Select
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_DELbSplitDT(ByVal pnSeqNo As Integer, ByVal pcQty As Double, ByVal ptTblDT As String, ByVal ptTmnNo As String, ByVal ptRunning As String, Optional ByRef pnNewSeqNo As Integer) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'________________________________
'   Call:
'   Cmt:    Ao51-06-12
'________________________________
    Dim tSql As String, tSqlV As String, tWhr As String
    Dim orsTemp As ADODB.Recordset
    Dim cQtyOld As Double, cRegPrice As Double
    Dim nLoop As Integer, nSeqNoNew As Integer
    Dim tSKUCode As String
    Dim bPass As Boolean
On Error GoTo Err_Handler
    If pcQty = 0 Then Exit Function
    
    DE_DELbSplitDT = False
    
    tSql = "SELECT * FROM " & ptTblDT
    tSql = tSql & vbCrLf & " WHERE (FNSdtSeqNo=" & pnSeqNo & ""
    tSql = tSql & vbCrLf & " AND FTTmnNum='" & ptTmnNo & "' AND FTShdTransNo ='" & ptRunning & "'" & " AND ISNULL(FTSdtCancelFlag,'0') <> '1')"
    
    'Set orsTemp = ocnAll.Execute(tSql) '*KT 54-03-24 ป้องกันโปรแกรมหลุด
    Call SP_SQLvExecute(tSql, orsTemp, ocnAll)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        cQtyOld = SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll")
        If pcQty > cQtyOld Then
            DE_DELbSplitDT = False
'            If bPass Then                          '*Ao 53-01-20           Cannot start more transactions on this session
'                ocnAll.RollbackTrans
'                bPass = False
'                nVB_CNCountBegin = nVB_CNCountBegin - 1
'             End If
            Exit Function
        Else
            cQtyOld = cQtyOld - pcQty
        End If
        If cQtyOld > 0 Then
            bPass = False
            ocnAll.BeginTrans
            bPass = True
            nVB_CNCountBegin = nVB_CNCountBegin + 1
            cRegPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtB4DisChg")
            '*Ao 52-03-12
'            tSKUCode = SP_FEDvChkNumeric(orsTemp, "FCSkuCode")
            tSKUCode = SP_FEDtChkString(orsTemp, "FTSkuCode")
            '________________________________________________________
            tWhr = " WHERE (FTTmnNum='" & ptTmnNo & "' AND FTShdTransNo ='" & ptRunning & "')"
            nSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("MAX", "FNSdtSeqNo", ptTblDT, tWhr)) + 1
        
            tSql = "UPDATE " & ptTblDT & " SET FCSdtQtyAll = " & pcQty
            
'            tSql = tSql & vbCrLf & ",FCSdtSaleAmt = FCSdtRegPrice * " & pcQty'*TW 51-07-17
'            tSql = tSql & vbCrLf & ",FCSdtAmtDisAvi = FCSdtRegPrice * " & pcQty'*TW 51-07-17
            
            tSql = tSql & vbCrLf & ",FCSdtSaleAmt = FCSdtB4DisChg * " & pcQty
            tSql = tSql & vbCrLf & ",FCSdtAmtDisAvi = FCSdtB4DisChg * " & pcQty
            
            tSql = tSql & vbCrLf & ",FCSdtBSubPro = CASE WHEN FCSdtBSubPro > 0 THEN FCSdtB4DisChg * " & pcQty & "  ELSE 0 END "
            tSql = tSql & vbCrLf & ",FCSdtBSubReg = CASE WHEN FCSdtBSubReg > 0 THEN FCSdtB4DisChg * " & pcQty & "  ELSE 0 END "
            
            tSql = tSql & vbCrLf & ",FCSdtQtyAllDisAvi = 0"
            tSql = tSql & vbCrLf & ",FNShdSign = 1"
            tSql = tSql & vbCrLf & ",FCSdtQty = " & pcQty
            tSql = tSql & vbCrLf & " WHERE (FNSdtSeqNo=" & pnSeqNo & ""
            tSql = tSql & vbCrLf & " AND FTTmnNum='" & ptTmnNo & "' AND FTShdTransNo ='" & ptRunning & "'" & " AND ISNULL(FTSdtCancelFlag,'0') <> '1')"
            'Call ocnAll.Execute(tSql)'*KT 54-03-24 ป้องกันโปรแกรมหลุด
           Call SP_SQLvExecute(tSql, , ocnAll)
        
            pnNewSeqNo = nSeqNoNew
            For nLoop = 0 To orsTemp.Fields.Count - 1
                Select Case UCase(orsTemp.Fields(nLoop).Name)
                    Case UCase("FNSdtSeqNo")
                        tSqlV = tSqlV & "," & nSeqNoNew
                    Case UCase("FCSdtQtyAll"), UCase("FCSdtQty"), UCase("FCSdtQtyAllDisAvi")
                        tSqlV = tSqlV & "," & cQtyOld
                    Case UCase("FNShdSign")
                        tSqlV = tSqlV & ",0"
                    Case UCase("FCSdtSaleAmt"), UCase("FCSdtAmtDisAvi")
                        tSqlV = tSqlV & "," & cRegPrice * cQtyOld
                    Case UCase("FTSkuIndex")            '*Neung 54-10-27 ML-PRV-072 V 3.0.69 Wrong Vendor Coupon QTY and Amount  ถ้าได้  Free ให้ FTSkuIndex=" "
                        tSqlV = tSqlV & "," & "''"
                    Case Else
                        tSqlV = tSqlV & DE_DELtCheckType(orsTemp, orsTemp.Fields(nLoop).Name)
                End Select
                
            Next nLoop
            If Mid(tSqlV, 1, 1) = "," Then
                tSqlV = Right(tSqlV, Len(tSqlV) - 1)
            End If
            
            tSql = "INSERT INTO " & ptTblDT
            tSql = tSql & vbCrLf & " VALUES (" & tSqlV & ")"
            
            'ocnAll.Execute (tSql) '*KT 54-03-24 ป้องกันโปรแกรมหลุด
            Call SP_SQLvExecute(tSql, , ocnAll)
            
            If bPass Then               '*Ao 53-01-20 Cannot start more transactions on this session
                ocnAll.CommitTrans
                bPass = False
                nVB_CNCountBegin = nVB_CNCountBegin - 1
            End If
            '*KT 52-02-19 member price ทำการ recal ของของการ split dt ที่ไม่ได้ deal ออกมา
            '-----------------------------------------------------------------------
           If bDE_MemberPrice Then
                If tVB_StaMemPriceAlw = "Y" Then
                    'Call oDisc.C_DATxSumDTMemberPrice(nSeqNoNew, tSKUCode, cQtyOld)
'                    Call DE_PRCxAddStoreDiscount(pnSeqNo, nSeqNoNew, cRegPrice * cQtyOld, cQtyOld)
                    bDE_MemberPriceAct = True
                End If
            End If
            '-------------------------------------------------------------------------
            
        End If
        DE_DELbSplitDT = True
    End If
    
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    If bPass Then
        ocnAll.RollbackTrans
        bPass = False
        nVB_CNCountBegin = nVB_CNCountBegin - 1
    End If
    DE_DELbSplitDT = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_SETxSkuUsed(ByVal pnPos As Integer, ByVal pnTime As Integer, ByVal pnSeqNo As Integer, ByVal pnQty As Integer, ByVal ptStaSku As String, Optional ByVal pcRound As Double = 1)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    If pnTime > 1 Then
        If nDE_KeepTime = pnTime Then
            atDE_SkuSeqUsed(pnPos) = atDE_SkuSeqUsed(pnPos) & "," & pnSeqNo
            atDE_SkuQtyUsed(pnPos) = atDE_SkuQtyUsed(pnPos) & "," & pnQty
            atDE_SkuStaUsed(pnPos) = atDE_SkuStaUsed(pnPos) & "," & ptStaSku
'            acDE_Round(pnPos) = pcRound
        Else
            atDE_SkuSeqUsed(pnPos) = atDE_SkuSeqUsed(pnPos) & ";" & pnSeqNo
            atDE_SkuQtyUsed(pnPos) = atDE_SkuQtyUsed(pnPos) & ";" & pnQty
            atDE_SkuStaUsed(pnPos) = atDE_SkuStaUsed(pnPos) & ";" & ptStaSku
'            acDE_Round(pnPos) = pcRound
        End If
    Else
        atDE_SkuSeqUsed(pnPos) = IIf(atDE_SkuSeqUsed(pnPos) <> "", atDE_SkuSeqUsed(pnPos) & ",", "") & pnSeqNo
        atDE_SkuQtyUsed(pnPos) = IIf(atDE_SkuQtyUsed(pnPos) <> "", atDE_SkuQtyUsed(pnPos) & ",", "") & pnQty
        atDE_SkuStaUsed(pnPos) = IIf(atDE_SkuStaUsed(pnPos) <> "", atDE_SkuStaUsed(pnPos) & ",", "") & ptStaSku
'        acDE_Round(pnPos) = pcRound
    End If
    nDE_KeepTime = pnTime
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function DE_SETbOrsAllDeal() As Boolean 'สร้าง Recordset ขึ้นมาเพื่อเก็บลำดับการลด Deal ทั้งหมดที่เกิดขึ้นได้
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    If Not orsDE_AllDeal Is Nothing Then
        If orsDE_AllDeal.State = adStateOpen Then orsDE_AllDeal.Close          'ไม่ใช้  State เพราะกรณี recordset ยังไม่ได้เปิดจะ error
        Set orsDE_AllDeal = Nothing
    End If
    Set orsDE_AllDeal = New ADODB.Recordset
    Call orsDE_AllDeal.Fields.Append("Test", adBigInt, 4)
    Call orsDE_AllDeal.Fields.Append("FTDeaSeq", adVarChar, 8000)
    Call orsDE_AllDeal.Fields.Append("FCDeaSum", adDouble, 8)
    Call orsDE_AllDeal.Fields.Append("FTDeaSkuSeq", adVarChar, 8000)
    Call orsDE_AllDeal.Fields.Append("FTDeaSkuQty", adVarChar, 8000)
    Call orsDE_AllDeal.Fields.Append("FTDeaSkuSta", adVarChar, 8000)
    '*Ao 52-0427
    Call orsDE_AllDeal.Fields.Append("FCDeaSumBlue", adDouble, 8)            'ป้ายน้ำเงินที่น้อย
'    Call orsDE_AllDeal.Fields.Append("FCDeaRound", adDouble, 8)
    '_______________________________________
    orsDE_AllDeal.CursorType = adOpenStatic
    orsDE_AllDeal.Open
    DE_SETbOrsAllDeal = True
    Exit Function
Err_Handler:
    DE_SETbOrsAllDeal = False
End Function

Private Sub DE_CALxMultiByDealNo(ByVal ptDealNo As String, ByVal ptDealType As String, ByVal ptSeqSku As String, ByVal ptQtySku As String, ByVal ptChkSta As String, ByVal ptWhe As String, ByRef pnIndex As Integer)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tChkSeqNo As String, tChkQty As String, tChkSta As String, tKeepSeqNo As String, tKeepQty As String, tKeepSta As String, tSKU As String, tQty As String, tSta As String
    Dim tDiscType As String, tSumRegPrice As String, tKeepGrp As String, tCalGrp As String

    Dim cRegPrice As Double, cDiscAmt As Double, cQtyGet As Double, cQtyBuy As Double, cAmtBuy As Double
    Dim cSumAlwDisc As Double, cSumDisc As Double, cSumDiscount As Double, cDiscAmtOld As Double, cSumAlwDiscOld As Double
    Dim cDiscount As Double, cDisPer7Amt As Double
    Dim cSumAmt8 As Double, cAmtGet8 As Double
    Dim nSeqDTOrg As Integer, nSumQty7 As Integer, nLevel As Integer
    
    On Error GoTo Err_Handler
    
    Select Case ptDealType
        Case "007"
            tKeepQty = ptQtySku
            tChkQty = UT_tToken(tKeepQty, ",", 1, False)
            Do While tChkQty <> ""
                nSumQty7 = nSumQty7 + Val(tChkQty)
                tChkQty = UT_tToken(tKeepQty, ",", 1, False)
            Loop
            orsDE_AFSack.Filter = "FTPdhDealNo = '" & ptDealNo & "'"
            orsDE_AFSack.Sort = "FNPddLevel DESC"
            orsDE_AFSack.MoveFirst
            Do While Not orsDE_AFSack.EOF
                If orsDE_AFSack.Fields("FNDacQtyBuy").Value <= nSumQty7 Then
                    nLevel = orsDE_AFSack.Fields("FNPddLevel").Value
                    Exit Do
                Else
                    nLevel = 0
                End If
                orsDE_AFSack.MoveNext
            Loop
        Case "008"
            tKeepQty = ptQtySku
            tKeepSeqNo = ptSeqSku
            tChkQty = UT_tToken(tKeepQty, ",", 1, False)
            tChkSeqNo = UT_tToken(tKeepSeqNo, ",", 1, False)
            Do While tChkSeqNo <> ""
                orsDE_AFSack.Filter = "FNSdtSeqNo = " & tChkSeqNo & " AND FTPdhDealNo = '" & ptDealNo & "'"
                cSumAmt8 = cSumAmt8 + Val(tChkQty) * orsDE_AFSack.Fields("FCSdtRegPrice").Value
                
                tChkQty = UT_tToken(tKeepQty, ",", 1, False)
                tChkSeqNo = UT_tToken(tKeepSeqNo, ",", 1, False)
            Loop
            orsDE_AFSack.Filter = "FTPdhDealNo = '" & ptDealNo & "'"
            orsDE_AFSack.Sort = "FNPddLevel DESC"
            orsDE_AFSack.MoveFirst
            Do While Not orsDE_AFSack.EOF
                If orsDE_AFSack.Fields("FCDacAmtBuy").Value <= cSumAmt8 Then
                    nLevel = orsDE_AFSack.Fields("FNPddLevel").Value
                    cAmtGet8 = orsDE_AFSack.Fields("FCDacAmtBuy").Value
                    Exit Do
                Else
                    nLevel = 0
                End If
                orsDE_AFSack.MoveNext
            Loop
    End Select
    
    
    tKeepSeqNo = ptSeqSku
    tKeepQty = ptQtySku
    tKeepSta = ptChkSta
    
    tChkSeqNo = UT_tToken(tKeepSeqNo, ",", 1, False)
    tChkQty = UT_tToken(tKeepQty, ",", 1, False)
    tChkSta = UT_tToken(tKeepSta, ",", 1, False)
    '*Ao 52-03-10
'    tKeepGrp = "#1"      'เริ่มด้วยค่า Group 1
    '____________________________
    ReDim Preserve typeDE_MutiDeal(pnIndex) As UDT_MultiDeal
    Do While tChkSeqNo <> ""
        
        nSeqDTOrg = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & tChkSeqNo)) 'หา Seq หลัง Split
        If DE_DELbSplitDT(nSeqDTOrg, Val(tChkQty), tVB_CNTblDT, tVB_CNTerminalNum, tVB_Running) Then 'สั่ง Split
            If ptDealType = "007" Or ptDealType = "008" Then
                orsDE_AFSack.Filter = "FTPdhDealNo = '" & ptDealNo & "' AND FNSdtSeqNo = " & tChkSeqNo & " AND FNPddLevel = " & nLevel
            Else
                orsDE_AFSack.Filter = "FTPdhDealNo = '" & ptDealNo & "' AND FNSdtSeqNo = " & tChkSeqNo
            End If
            tCalGrp = orsDE_AFSack.Fields("FNDacSkuGrp").Value
            
            cRegPrice = orsDE_AFSack.Fields("FCSdtRegPrice").Value
            If tChkSta = "1" Then
                cQtyGet = orsDE_AFSack.Fields("FNDacQtyGet").Value
                cQtyBuy = 0
'                cDiscAmt = orsDE_AFSack.Fields("FCDacDis").Value 'Old
'                cDiscAmt = orsDE_AFSack.Fields("FCDacDisAmt").Value
'Ao 52-03-10    ใช้สำหรับ DealType 2
'                cDiscAmtOld = cDiscAmt
                cDiscAmt = orsDE_AFSack.Fields("FCDacDisAmt").Value
            Else
                cQtyGet = 0
                cQtyBuy = orsDE_AFSack.Fields("FNDacQtyBuy").Value
                cDiscAmt = 0
            End If
            cAmtBuy = orsDE_AFSack.Fields("FCDacAmtBuy").Value
            tDiscType = orsDE_AFSack.Fields("FTPddDisType").Value
            tSKU = IIf(tSKU = "", "", tSKU & ",") & nSeqDTOrg
            tQty = IIf(tQty = "", "", tQty & ",") & tChkQty
            tSta = IIf(tSta = "", "", tSta & ",") & tChkSta
            tSumRegPrice = IIf(tSumRegPrice = "", "", tSumRegPrice & ",") & Trim(Str(cRegPrice))
            
            If tChkSta = "1" Then
                Select Case ptDealType
                    Case "0010", "0020", "0030", "0040"
                        cSumAlwDisc = cSumAlwDisc + (Val(tChkQty) * cRegPrice)
                        typeDE_MutiDeal(pnIndex).tFreeFlag = "2"
                        tDiscType = "4"                                                     '*Ao 2012/12/25  แก้ไข Deal Type 0020
                    Case "001", "002", "007"
                        cSumAlwDisc = cSumAlwDisc + (Val(tChkQty) * cRegPrice)
                        typeDE_MutiDeal(pnIndex).tFreeFlag = "3"
                    Case "008"
                        cSumAlwDisc = cSumAlwDisc + (cAmtGet8 / cSumAmt8 * (Val(tChkQty) * cRegPrice))
                        typeDE_MutiDeal(pnIndex).tFreeFlag = "3"
                    Case "003", "004", "006"
                        cSumAlwDisc = cSumAlwDisc + (Val(tChkQty) * cRegPrice)
                        typeDE_MutiDeal(pnIndex).tFreeFlag = "2"
                    Case "005"
                        cSumAlwDiscOld = cSumAlwDiscOld + (Val(tChkQty) * cRegPrice)
                        typeDE_MutiDeal(pnIndex).tFreeFlag = "3"
                    Case Else
                        cSumAlwDisc = cAmtBuy
                        typeDE_MutiDeal(pnIndex).tFreeFlag = "3"
                End Select
            Else
                Select Case ptDealType
                    Case "0010", "0020", "0030", "0040"
                        typeDE_MutiDeal(pnIndex).tFreeFlag = "1"
                        tDiscType = "4"                                                             '*Ao 2012/12/25  แก้ไข Deal Type 0020
                End Select
            End If
        End If
        tChkSeqNo = UT_tToken(tKeepSeqNo, ",", 1, False)
        tChkQty = UT_tToken(tKeepQty, ",", 1, False)
        tChkSta = UT_tToken(tKeepSta, ",", 1, False)
    Loop
    If Not orsDE_AFSack.EOF Then
    
        typeDE_MutiDeal(pnIndex).cDiscPer = 0
        If ptDealType = "005" Then                 '*Ao 09-10-21
            cSumAlwDisc = cAmtBuy
        End If
        Select Case tDiscType
            Case "1"
                cSumDisc = ((cDiscAmt / 100) * cSumAlwDisc)
                typeDE_MutiDeal(pnIndex).cDiscPer = cDiscAmt
            Case "2"
                cSumDisc = cDiscAmt
            Case "3"
                cSumDisc = cSumAlwDisc - cDiscAmt
            Case "4"
                cSumDisc = cSumAlwDisc
        End Select
        If ptDealType = "005" Then                 '*Ao 09-10-21
            cSumAlwDisc = cSumAlwDiscOld
        End If
        typeDE_MutiDeal(pnIndex).cAlwDisc = cSumAlwDisc
        typeDE_MutiDeal(pnIndex).cDiscAmt = cSumDisc
        typeDE_MutiDeal(pnIndex).tDealNo = ptDealNo
        typeDE_MutiDeal(pnIndex).tDealType = ptDealType
        typeDE_MutiDeal(pnIndex).tDiscQty = tQty
        typeDE_MutiDeal(pnIndex).tDiscType = tDiscType
        typeDE_MutiDeal(pnIndex).tDTSeqNo = tSKU
        typeDE_MutiDeal(pnIndex).tEventID = orsDE_AFSack.Fields("FTEvnCode").Value
        typeDE_MutiDeal(pnIndex).tSkuFlag = tSta
        typeDE_MutiDeal(pnIndex).tRegPrice = tSumRegPrice
    End If

    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub
                
Private Function DE_CHKnDDP(ByVal pnDTSeqNo As Integer, ByVal pnENDiscType As eEN_DiscType) As Integer
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------
'Ret    0:Not Allow
'           1:Allow
'           2:Thebest
'--------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tDctNo$
On Error GoTo Err_Handler
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        tDctNo = ""
        tSql = "SELECT CD.FNDctNo"
        tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD & " CD "
        tSql = tSql & vbCrLf & " WHERE  CD.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & Format(tUT_SaleDate, "YYYY/MM/DD") & "'"
        tSql = tSql & vbCrLf & " AND CD.FNSdtSeqNo = " & pnDTSeqNo
        tSql = tSql & vbCrLf & " GROUP BY CD.FNDctNo"
        If SP_SQLvExecute(tSql, orsTemp) = 0 Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                Do While Not orsTemp.EOF
                    tDctNo = tDctNo & CStr(SP_FEDvChkNumeric(orsTemp, "FNDctNo"))
                    orsTemp.MoveNext
                    If Not orsTemp.EOF Then
                        '*KT 53-02-20 กรณีที่ tDctNo ไม่มีค่าแล้วเข้า  tDctNo = tDctNo & "," ทำให้เกิด Error ที่เป็นหมายเลข 0 ได้
                        If CStr(SP_FEDvChkNumeric(orsTemp, "FNDctNo")) <> "" Then
                            tDctNo = tDctNo & ","
                        End If
                    End If
                Loop
            End If
            If tDctNo <> "" Then
                tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy"
                tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO "
                tSql = tSql & vbCrLf & "  WHERE PO.FNDdpNo2 = " & pnENDiscType
                tSql = tSql & vbCrLf & "  AND PO.FNDdpNo1 IN (" & tDctNo & ")"
                If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                    If Not (orsTemp.BOF And orsTemp.EOF) Then
                        orsTemp.Filter = "FTDdpStaPolicy = 'N'"
                        If Not (orsTemp.BOF And orsTemp.EOF) Then
                            DE_CHKnDDP = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
                        Else
                            orsTemp.Filter = "FTDdpStaPolicy = 'B'"
                            If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
                                Do While Not orsTemp.EOF
                                    orsTemp.MoveNext
                                Loop
                                DE_CHKnDDP = 2
                            Else
                                DE_CHKnDDP = 1
                            End If
                        End If
                    Else
                        DE_CHKnDDP = 1 'ยังไม่มีส่วนลดอื่น ให้ผ่าน
                    End If
                Else
                    DE_CHKnDDP = 0 'Error ไม่ผ่าน
                End If
            Else
                DE_CHKnDDP = 1
            End If
        Else
            DE_CHKnDDP = 0
        End If
    Else
        tSql = "SELECT PO.FNDdpNo1,PO.FNDdpNo2 ,PO.FTDdpStaPolicy"
        tSql = tSql & vbCrLf & " FROM TCNMDoubleDiscPolicy PO LEFT JOIN " & tVB_CNTblCD & " CD ON PO.FNDdpNo1 = CD.FNDctNo"
        tSql = tSql & vbCrLf & " WHERE PO.FNDdpNo2 = " & pnENDiscType
        tSql = tSql & vbCrLf & " AND CD.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND CD.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND CD.FDShdTransDate = '" & Format(tUT_SaleDate, "YYYY/MM/DD") & "'"
        tSql = tSql & vbCrLf & " AND CD.FNSdtSeqNo = " & pnDTSeqNo
        tSql = tSql & vbCrLf & " ORDER BY CD.FNScdSeqNo "
        If SP_SQLbExecute(tSql, orsTemp) Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                orsTemp.Filter = "FTDdpStaPolicy = 'N'"
                If Not (orsTemp.BOF And orsTemp.EOF) Then
                    DE_CHKnDDP = 0 'มีตัวที่เป็น N ไม่ให้ผ่าน
                Else
                    orsTemp.Filter = "FTDdpStaPolicy = 'B'"
                    If Not (orsTemp.BOF And orsTemp.EOF) Then 'มีตัวที่เป็น B
                        Do While Not orsTemp.EOF
                            orsTemp.MoveNext
                        Loop
                        DE_CHKnDDP = 2
                    Else
                        DE_CHKnDDP = 1
                    End If
                End If
            Else
                DE_CHKnDDP = 1 'ยังไม่มีส่วนลดอื่น ให้ผ่าน
            End If
        Else
            DE_CHKnDDP = 0 'Error ไม่ผ่าน
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
Err_Handler:
    DE_CHKnDDP = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Sub DE_PRCxClearDiscount()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim tWhe As String
On Error GoTo Err_Handler
    tWhe = " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & vbCrLf & " AND FDShdTransDate = '" & Format(tUT_SaleDate, "YYYY/MM/DD") & "'"
    
    tSql = "DELETE "
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & tWhe
    'Call ocnAll.Execute(tSql) '*KT 54-03-24 ป้องกันโปรแกรมหลุด
    Call SP_SQLvExecute(tSql, , ocnAll)
    With oDisc
        .DISxClearAll
        .dTranDate = tUT_SaleDate
        .tTerminalNo = tVB_CNTerminalNum
        .tTransNo = tVB_Running
        .DISbClearDiscDT
    End With
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DE_CHKxModeMemberPrice()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # kitti                              LAST MODIFIED BY           # kitti
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 52-02-19
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql$
    
    '*KT 52-02-19 member price
    'ตรวจสอบวสถานะว่าบิลปัจจุบันที่ทำรายการอยู่ member price หรือไม่
    bDE_MemberPrice = False
    If bVB_CNMemberPrice Then
        If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" Then
            bDE_MemberPrice = True
            
            tSql = "UPDATE " & tVB_CNTblDT & " SET "
            tSql = tSql & vbCrLf & " FCSdtB4DisChg = FCSdtB4DisChgOrg"
            tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
            tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
            tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & Format(tUT_SaleDate, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
            'Call ocnAll.Execute(tSql) '*KT 54-03-24 ป้องกันโปรแกรมหลุด
            Call SP_SQLvExecute(tSql, , ocnAll)
            
        End If
    End If
End Sub


Private Function DE_GETcCallDeal(ByVal ptDealNo As String, ByVal ptFreeFlag As String, ByVal ptDealType As String, ByVal ptFreeType As String, ByVal ptEventID As String _
, ByVal ptDisType As String, ByVal pcDealRate As Double, ByVal pnDTSeqNo As Integer, ByVal pcQty As Double, ByVal pcDiscValue As Double, ByVal pcAlwCalDisc As Double, Optional ByVal pbLastItem As Boolean = False) As Double
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
 '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo Err_Handler
    With oDisc
        .DISxClearAll
        .dTranDate = tUT_SaleDate
        .tTerminalNo = tVB_CNTerminalNum
        .tTransNo = tVB_Running
        .tRefCode = ptDealNo 'DealNo
        .tStaChkDepart = ptFreeFlag 'FreeFlag 1=Buy 2=Get  3=Get&Buy 4=Buy Free
        .tCardType = ptDealType 'DealType
        .tCardKey = ptFreeType ''FreeType
        .tEventID = ptEventID 'EventID FTScdEventID
        .bIsLastDeal = pbLastItem
        Select Case ptDisType
            Case "1" 'DiscPer?
                .nENLnDisType = nEN_LnDis1Per
'                .cBalance = pcDealRate
                .cDiscPerOrg = pcDealRate
            Case "2"
                .nENLnDisType = nEN_LnDis2Amt
'                .cBalance = 0
                .cDiscAmtOrg = pcDealRate
            Case "3"
                .nENLnDisType = nEN_LnDis3NewPrice
'                .cBalance = 0
                .cDiscNewOrg = pcDealRate
            Case "4"
                .nENLnDisType = nEN_LnDis4Free
                .cBalance = 0
        End Select
        .cAlwDis = IIf(pcDiscValue = 0, 0, pcAlwCalDisc)
        .cAlwDisQty = pcQty 'IIf(pcDiscValue = 0, 0, pcQty)
        .nDTSeqNo = pnDTSeqNo 'DTSeqNo
        .cDiscValue = pcDiscValue 'Discount
        If ptFreeType = "0020" Then
            Call DE_DELbSplitDT(pnDTSeqNo, pcQty, tVB_CNTblDT, tVB_CNTerminalNum, tVB_Running)
        Else
            Call DE_DELbSplitDT(pnDTSeqNo, pcQty, tVB_CNTblDT, tVB_CNTerminalNum, tVB_Running)
        End If
        
        DE_GETcCallDeal = .DIScDiscount(nEN_Disc01Deal)
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_BAKxors2Array(ByVal porsTemp As ADODB.Recordset, ByRef ptype_KeepDeal() As UDT_KeepDeal)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nI As Integer
    On Error GoTo Err_Handler
    With porsTemp
        .MoveFirst
        nI = 0
        Do While Not .EOF
            ReDim Preserve ptype_KeepDeal(nI) As UDT_KeepDeal
            ptype_KeepDeal(nI).FCDacAmtBuy = .Fields("FCDacAmtBuy").Value
            ptype_KeepDeal(nI).FCDacDis = .Fields("FCDacDis").Value
            ptype_KeepDeal(nI).FCDacDisAmt = .Fields("FCDacDisAmt").Value
            ptype_KeepDeal(nI).FCSdtQtyAll = .Fields("FCSdtQtyAll").Value
            ptype_KeepDeal(nI).FCSdtRegPrice = .Fields("FCSdtRegPrice").Value
            ptype_KeepDeal(nI).FNDacQtyBuy = .Fields("FNDacQtyBuy").Value
            ptype_KeepDeal(nI).FNDacQtyGet = .Fields("FNDacQtyGet").Value
            ptype_KeepDeal(nI).FNDacSkuGrp = .Fields("FNDacSkuGrp").Value
            ptype_KeepDeal(nI).FNPddLevel = .Fields("FNPddLevel").Value
            ptype_KeepDeal(nI).FNSdtSeqNo = .Fields("FNSdtSeqNo").Value
            ptype_KeepDeal(nI).FTDacDealType = .Fields("FTDacDealType").Value
            ptype_KeepDeal(nI).FTDacSku = .Fields("FTDacSku").Value
            ptype_KeepDeal(nI).FTEvnCode = .Fields("FTEvnCode").Value
            ptype_KeepDeal(nI).FTFreeFlag = .Fields("FTFreeFlag").Value
            ptype_KeepDeal(nI).FTPddDisType = .Fields("FTPddDisType").Value
            ptype_KeepDeal(nI).FTPdhDealNo = .Fields("FTPdhDealNo").Value
            .MoveNext
            nI = nI + 1
        Loop
    End With
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Private Function DE_SINbCal001(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByRef pnCountRound As Long) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cDealDiscPer As Double, cDealDiscAmt As Double
'    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
'    Dim tFreeFlag As String, tFilter As String, tSort As String
'    Dim nDTSeqNo As Integer
'    Dim nDTSeqNoNew As Integer
'    Dim cSumCanDisc As Double
'    Dim cRegPrice As Double
'    Dim nQtyUseRound As Long
'    Dim nKeepQtyGet As Integer
'    '*KT 52-02-19 member price
'    Dim bMemberPrice As Boolean
'    Dim cB4DisChgOrg As Double
'    Dim cMemberPrice As Double
'    Dim tSKUCode As String
'    Dim cDiscTotalAll As Double
'    Dim tEventIDMemberPrice As String
'    Dim nI As Long
'On Error GoTo Err_Handler
'    DE_SINbCal001 = False
'    tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"              '*Ao 09-10-21     เลือกสินค้าแพงที่สุด,สีป้ายเหลือง-->น้ำเงิน
'    porsTemp.Sort = tSort
'    cDiscTotalAll = 0
'
'    Select Case ptLnDiscType
'        Case "1" 'Per
'            cDealDiscPer = orsDE_DealSingle.Fields("FCDacDisAmt").Value
'        Case "2", "3"
'            cDealDiscPer = 0
'        Case Else
'            cDealDiscPer = 0
'    End Select
'    nQtyGet = orsDE_DealSingle.Fields("FNDacQtyGet").Value
'    nQtyGetOrg = nQtyGet
'
'    cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, "1", tSort) '*TW 51-06-27
'    tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"              '*Ao 09-10-21     เลือกสินค้าแพงที่สุด,สีป้ายเหลือง-->น้ำเงิน
'    porsTemp.Sort = tSort
'    porsTemp.MoveFirst
'    Do While Not porsTemp.EOF
'        If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then '*TW 51-06-27
'            cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
'             tSKUCode = porsTemp.Fields("FTDacSku").Value
'            If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                nKeepQtyGet = nQtyGet
'                nQtyCal = nQtyGet
'                nQtyGet = 0
'            Else
'                nKeepQtyGet = nQtyGet
'                nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'            End If
'            tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
'            nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
'            If (orsDE_DealSingle.Fields("FNDacQtyGet").Value = nKeepQtyGet) And (orsDE_DealSingle.Fields("FNDacQtyGet").Value <= porsTemp.Fields("FCSdtQtyAll").Value) Then
'                nQtyUseRound = porsTemp.Fields("FCSdtQtyAll").Value \ orsDE_DealSingle.Fields("FNDacQtyGet").Value
'                Select Case ptLnDiscType
'                    Case "1" 'Per
'                        cDealDiscAmt = cRegPrice * (cDealDiscPer / 100) * nQtyCal
'                    Case "2"
'                        If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
'                            DE_SINbCal001 = False
'                            Exit Function
'                        End If
'                        cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc)
'                    Case "3"
'                        If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
'                            DE_SINbCal001 = False
'                            Exit Function
'                        End If
'                        cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc))
'                End Select
'
'                cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) * nQtyUseRound
'                '*Ao 52-04-20
'                cDiscTotalAll = cDiscTotalAll + (SP_FMTcCurPoint(cDealDiscAmt, True, False) * nQtyUseRound)
'                '______________________________________________________________
'                '*KT 52-02-19 member price
'                'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
'                '--------------------------------------------------------------------------------
'                bMemberPrice = False
'                If bDE_MemberPrice Then
'                    cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                    tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                    '*KT 52-09-08 Member Price + Consignment
'                    tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                    If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
'                        If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
'                            bMemberPrice = False
'                            If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                'ยังไม่เคยถูกเรียกใช้
'                            Else
'                                For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                    If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                        typeDE_KeepOld(nI).FNDctNo = 0
'                                        Exit For
'                                    End If
'                                Next nI
'                            End If
'                        Else
'                            bMemberPrice = True
'                        End If
'                    End If
'                End If
'                '*Ao 52-04-20
'                If nQtyUseRound < 2 Then
'                    Call DE_SINxDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value * nQtyUseRound, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
'                End If
'
'                If Not bMemberPrice Then
'                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal * nQtyUseRound, cDealDiscAmt, cRegPrice * nQtyCal * nQtyUseRound)
'                    Call DE_UPDxQtySingle(porsTemp, nQtyCal * nQtyUseRound, nDTSeqNo)
'                Else
'                    If tVB_StaMemPriceAlw = "Y" Then
'                        If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                            bDE_MemberPriceAct = True
'                        Else
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal * nQtyUseRound, cDealDiscAmt, cRegPrice * nQtyCal * nQtyUseRound)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal * nQtyUseRound, nDTSeqNo)
'                        End If
'                    End If
'                End If
'                '---------------------------------------------------------------------------------
'            Else
'                nQtyUseRound = 1
'                Select Case ptLnDiscType
'                    Case "1" 'Per
'                        cDealDiscAmt = cRegPrice * (cDealDiscPer / 100) * nQtyCal
'                    Case "2"
'                        If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
'                            DE_SINbCal001 = False
'                            Exit Function
'                        End If
'                        cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc)
'                    Case "3"
'                        If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
'                            DE_SINbCal001 = False
'                            Exit Function
'                        End If
'                        cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc))
'                End Select
'                '*Ao 52-04-20
'                cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
'
'                cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt)
'                '*KT 52-02-19 member price
'                'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
'                '--------------------------------------------------------------------------------
'                bMemberPrice = False
'                If bDE_MemberPrice Then
'                    cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                    tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                    '*KT 52-09-08 Member Price + Consignment
'                    tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                     If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
'                        If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
'                            bMemberPrice = False
'                            If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                'ยังไม่เคยถูกเรียกใช้
'                            Else
'                                For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                    If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                        typeDE_KeepOld(nI).FNDctNo = 0
'                                        Exit For
'                                    End If
'                                Next nI
'                            End If
'                        Else
'                            bMemberPrice = True
'                        End If
'                    End If
'                End If
'                '*Ao 52-04-20
'                If nQtyUseRound < 2 Then
'                    Call DE_SINxDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
'                End If
'
'                If Not bMemberPrice Then
'                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                Else
'                    If tVB_StaMemPriceAlw = "Y" Then
'                         If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                            bDE_MemberPriceAct = True
'                        Else
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                        End If
'                    End If
'                End If
'                '--------------------------------------------------------------------------------
'            End If
'            If nQtyGet <= 0 Then
'                pnCountRound = pnCountRound - nQtyUseRound
'                Exit Function
'            End If
'        End If
'        porsTemp.MoveNext
'    Loop
'   DE_SINbCal001 = True
'    Exit Function
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_SINbCal001(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 09-11-20
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDealDiscPer As Double, cDealDiscAmt As Double, cSumCanDisc As Double, cRegPrice As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer, nDTSeqNo As Integer, nDTSeqNoNew As Integer, nQtyUseRound As Long
    Dim tFreeFlag As String, tFilter As String, tSort As String
    Dim nKeepQtyGet As Integer
    '*KT 52-02-19 member price
    Dim bMemberPrice As Boolean
    Dim cB4DisChgOrg As Double
    Dim cMemberPrice As Double
    Dim tSKUCode As String
    Dim cDiscTotalAll As Double
    Dim tEventIDMemberPrice As String
    Dim nI As Long
On Error GoTo Err_Handler
    
    DE_SINbCal001 = False
    With porsTemp
        tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard,FNSdtSeqNo"              '*Ao 09-10-21     เลือกสินค้าแพงที่สุด,สีป้ายเหลือง-->น้ำเงิน
        .Sort = tSort
        cDiscTotalAll = 0
        
'        Select Case ptLnDiscType
'            Case "1" 'Per
'                cDealDiscPer = .Fields("FCDacDisAmt").Value
'            Case "2", "3"
'                cDealDiscPer = 0
'            Case Else
'                cDealDiscPer = 0
'        End Select
        nQtyGet = .Fields("FNDacQtyGet").Value * pnCountRound
'        nQtyGetOrg = nQtyGet
        
        cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, "1", tSort) '*TW 51-06-27
'        tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"              '*Ao 09-10-21     เลือกสินค้าแพงที่สุด,สีป้ายเหลือง-->น้ำเงิน
        .Sort = tSort
        .MoveFirst
        Do While Not .EOF
            If .Fields("FCSdtQtyAll").Value > 0 Then '*TW 51-06-27
                cRegPrice = .Fields("FCSdtRegPrice").Value
                 tSKUCode = .Fields("FTDacSku").Value
                If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                    nKeepQtyGet = nQtyGet
                    nQtyCal = nQtyGet
                    nQtyGet = 0
                Else
'                    nKeepQtyGet = nQtyGet
                    nQtyCal = .Fields("FCSdtQtyAll").Value
                    nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
                End If
                tFreeFlag = .Fields("FTFreeFlag").Value
                nDTSeqNo = .Fields("FNSdtSeqNo").Value
                Select Case ptLnDiscType
                    Case "1" 'Per
                        cDealDiscPer = .Fields("FCDacDisAmt").Value
                        cDealDiscAmt = cRegPrice * (cDealDiscPer / 100) * nQtyCal
                    Case "2"
                        If .Fields("FCDacDisAmt").Value > cSumCanDisc Then
                            DE_SINbCal001 = False
                            Exit Function
                        End If
                        cDealDiscPer = .Fields("FCDacDisAmt").Value
                        cDealDiscAmt = ((.Fields("FCDacDisAmt").Value * pnCountRound) * (cRegPrice * nQtyCal)) / cSumCanDisc
                    Case "3"
                        If .Fields("FCDacDisAmt").Value > cSumCanDisc Then
                            DE_SINbCal001 = False
                            Exit Function
                        End If
                        cDealDiscPer = .Fields("FCDacDisAmt").Value
                        cDealDiscAmt = (.Fields("FCSdtRegPrice").Value * nQtyCal) - (((.Fields("FCDacDisAmt").Value * pnCountRound) * (cRegPrice * nQtyCal)) / cSumCanDisc)
                End Select
                    
                cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt)
                '*Ao 52-04-20
                cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
                '______________________________________________________________
                '*KT 52-02-19 member price
                'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
                '--------------------------------------------------------------------------------
                bMemberPrice = False
                If bDE_MemberPrice Then
                    cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                    tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                    '*KT 52-09-08 Member Price + Consignment
                    tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                    If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                        If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                            bMemberPrice = False
'                            If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                'ยังไม่เคยถูกเรียกใช้
'                            Else
'                                For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                    If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
''                                        typeDE_KeepOld(nI).FNDctNo = 0
'                                        Exit For
'                                    End If
'                                Next nI
'                            End If
                        Else
                            bMemberPrice = True
                        End If
                    End If
                End If
                '*Ao 52-04-20
                If nQtyGet < 1 Then
                    Call DE_SINxDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value * pnCountRound, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
                End If
                    
                If Not bMemberPrice Then
                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                     If nQtyGet < 1 Then
                        Exit Do
                    End If
                Else
                    If tVB_StaMemPriceAlw = "Y" Then
                        If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                            bDE_MemberPriceAct = True
                            DE_SINbCal001 = True
                            Exit Function
                        Else
                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                            If nQtyGet < 1 Then
                                Exit Do
                            End If
                        End If
                    End If
                End If
                '---------------------------------------------------------------------------------
            End If
            porsTemp.MoveNext
        Loop
    End With
    DE_SINbCal001 = True
    Exit Function
Err_Handler:
    DE_SINbCal001 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Sub DE_SINxCal0010(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByRef pnCountRound As Long)
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cDealDiscPer As Double, cDealDiscAmt As Double
'    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
'    Dim nQtyBuy As Integer
'    Dim tFreeFlag As String
'    Dim nDTSeqNo As Integer
'    Dim nDTSeqNoNew As Integer
'    Dim cSumCanDisc As Double
'    Dim cRegPrice As Double
'    Dim nQtyUseRound As Long
'    On Error GoTo Err_Handler
'    cDealDiscPer = 0
'    nQtyBuy = orsDE_DealSingle.Fields("FNDacQtyBuy").Value * pnCountRound
'    nQtyGet = pnCountRound
'    nQtyBuy = nQtyBuy - pnCountRound
'    nQtyGetOrg = nQtyGet
'    porsTemp.MoveFirst
'    Do While Not porsTemp.EOF
'        If nQtyGet > 0 Then
'            If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
'                cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
'                If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                    nQtyCal = nQtyGet
'                    nQtyGet = 0
'                Else
'                    nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                    nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                End If
'
'                tFreeFlag = "2"
'                '*Ao 09-09-22
'                nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
'                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                cDealDiscAmt = cRegPrice * nQtyCal
'                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'
'                Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                porsTemp.MoveNext
'            Else
'                porsTemp.MoveNext
'            End If
'            If nQtyBuy <= 0 Then Exit Do
'        Else
'            If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
'                cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
'                If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyBuy Then
'                    nQtyCal = nQtyBuy
'                    nQtyBuy = 0
'                Else
'                    nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                    nQtyBuy = nQtyBuy - porsTemp.Fields("FCSdtQtyAll").Value
'                End If
'                tFreeFlag = "1"
'                nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
'                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                cDealDiscAmt = 0
'                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                If nQtyBuy <= 0 Then Exit Do
'            End If
'            porsTemp.MoveNext
'        End If
'    Loop
'    Exit Sub
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub

Private Function DE_SINbCal0010(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 15-10-09
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDealDiscPer As Double, cDealDiscAmt As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
    Dim nQtyBuy As Integer
    Dim tFreeFlag As String, tSort As String
    Dim nDTSeqNo As Integer
    Dim nDTSeqNoNew As Integer
    Dim cSumCanDisc As Double
    Dim cRegPrice As Double
    Dim nQtyUseRound As Long
    On Error GoTo Err_Handler
    
    '*Ao 09-10-15
    DE_SINbCal0010 = False
    tSort = "FCSdtRegPrice,FTSkuStaMallEmpCard"                'Buy เลือกราคาน้อยสุด,ราคาเท่ากันเลือกป้ายเหลืองก่อนป้ายน้ำเงิน
    porsTemp.Sort = tSort
    cDealDiscPer = 0
    nQtyBuy = porsTemp.Fields("FNDacQtyBuy").Value * pnCountRound
    nQtyGet = porsTemp.Fields("FNDacQtyGet").Value * pnCountRound
    
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
            cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
            If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyBuy Then
                nQtyCal = nQtyBuy
                nQtyBuy = 0
            Else
                nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                nQtyBuy = nQtyBuy - porsTemp.Fields("FCSdtQtyAll").Value
            End If
            tFreeFlag = "1"
'            tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
            nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'            cDealDiscAmt = cRegPrice * nQtyCal
            cDealDiscAmt = 0
            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
        End If
        porsTemp.MoveNext
        If nQtyBuy < 1 Then Exit Do
    Loop
    'Get
    '*Ao 09-10-15
    tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"                'Get เลือกราคามากสุด,ราคาเท่ากันเลือกป้ายเหลืองก่อนป้ายน้ำเงิน
    porsTemp.Sort = tSort
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
            cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
            If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
                nQtyCal = nQtyGet
                nQtyGet = 0
            Else
                nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
            End If
            tFreeFlag = "2"
'            tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
            nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'            cDealDiscAmt = 0
            cDealDiscAmt = cRegPrice * nQtyCal
            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
            If nQtyGet < 1 Then Exit Do
        End If
        porsTemp.MoveNext
    Loop
    DE_SINbCal0010 = True
    Exit Function
Err_Handler:
    DE_SINbCal0010 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_SINbCal002(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:        '*Ao 09-11-03
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDealDiscPer As Double, cDealDiscAmt As Double, cSumCanDisc As Double, cRegPrice As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer, nMaxGrp As Integer
    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nPos As Integer, nDTSeqNoNew As Integer, nDTSeqNo As Integer
    Dim tFreeFlag As String, tSort As String

    '*KT 52-02-19 member price
    Dim bMemberPrice As Boolean
    Dim cB4DisChgOrg As Double
    Dim cMemberPrice As Double
    Dim tSKUCode As String
    Dim cDiscTotalAll As Double
    Dim tEventIDMemberPrice As String
    Dim nI As Long
On Error GoTo Err_Handler
    DE_SINbCal002 = False
    nMaxGrp = DE_GETnMaxGroup(ptDealNo, porsTemp)
    tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"                '*Ao 09-10-21     เลือกสินค้าแพงที่สุด,สีป้ายเหลือง-->น้ำเงิน
    porsTemp.Sort = tSort
    cDiscTotalAll = 0
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
            nGrp = porsTemp.Fields("FNDacSkuGrp").Value
            cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
            tSKUCode = porsTemp.Fields("FTDacSku").Value
            If nGrp <> nUncalGrp Then
                nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
                If nGrp <> nGrpOld Then
'                    Select Case ptLnDiscType
'                        Case "1" 'Per
'                            cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
'                        Case "2", "3"
'                            cDealDiscPer = 0
'                        Case Else
'                            cDealDiscPer = 0
'                    End Select
                    tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
                    nQtyGet = porsTemp.Fields("FNDacQtyGet").Value * CInt(pnCountRound)
'                    nQtyGetOrg = nQtyGet
                    
                    'If cSumCanDisc <= 0 Then cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, nGrp)                            '*TW 51-06-27
                    If cSumCanDisc <= 0 Then
                        nPos = porsTemp.AbsolutePosition
                        cSumCanDisc = DE_GETcSumDiscByQtyGrp(porsTemp, nQtyGet, nGrp, tSort, pnCountRound)
'                        cDiscTotalAll = 0               '*Ao 09-10-15
                        porsTemp.Sort = tSort
                        porsTemp.AbsolutePosition = nPos
                    End If
                    
                    If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
                        nQtyCal = nQtyGet
                        nQtyGet = 0
                    Else
                        nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                        nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
                    End If
                    Select Case ptLnDiscType
                        Case "1" 'Per
                            cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
                            cDealDiscAmt = cRegPrice * (cDealDiscPer / 100) * nQtyCal
                        Case "2"
                            If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
                                DE_SINbCal002 = False
                                Exit Function
                            End If
                            cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
                            cDealDiscAmt = ((porsTemp.Fields("FCDacDisAmt").Value * pnCountRound) * (cRegPrice * nQtyCal)) / cSumCanDisc
                        Case "3"
                            If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
                                DE_SINbCal002 = False
                                Exit Function
                            End If
                            cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
                            cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal) - (((porsTemp.Fields("FCDacDisAmt").Value * pnCountRound) * (cRegPrice * nQtyCal)) / cSumCanDisc)
                    End Select
                    
                    cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
                    '*Ao 52-04-20
                    cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
                    '*KT 52-02-19 member price
                    'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
                    '--------------------------------------------------------------------------------
                    bMemberPrice = False
                    If bDE_MemberPrice Then
                        cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                        cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                        tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                        '*KT 52-09-08 Member Price + Consignment
                        tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                        If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                            If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                                bMemberPrice = False
'                                If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                    'ยังไม่เคยถูกเรียกใช้
'                                Else
'                                    For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                        If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
''                                            typeDE_KeepOld(nI).FNDctNo = 0
'                                            Exit For
'                                        End If
'                                    Next nI
'                                End If
                            Else
                                bMemberPrice = True
                            End If
                        End If
                    End If
                    '*Ao 52-04-20/1
                    If nGrp = nMaxGrp Then
                        If nQtyGet < 1 Then
                            Call DE_SINxDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value * pnCountRound, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
                        End If
                    End If
                    If Not bMemberPrice Then
                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                        Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                        nGrpOld = nGrp
                        If nQtyGet < 1 Then
                            nUncalGrp = nGrp
                        End If
                    Else
                        If tVB_StaMemPriceAlw = "Y" Then
                            If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                                bDE_MemberPriceAct = True
                                DE_SINbCal002 = True
                                Exit Function
                            Else
                                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                                Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                nGrpOld = nGrp
                                If nQtyGet < 1 Then
                                    nUncalGrp = nGrp
                                End If
                            End If
                        End If
                    End If
                    '------------------------------------------------------------------------------------
                Else
                    If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
                        nQtyCal = nQtyGet
                        nQtyGet = 0
                    Else
                        nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                        nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
                    End If
                    Select Case ptLnDiscType
                        Case "1" 'Per
                            cDealDiscAmt = cRegPrice * (cDealDiscPer / 100) * nQtyCal
                        Case "2"
                            If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
                                DE_SINbCal002 = False
                                Exit Function
                            End If
                            cDealDiscAmt = ((porsTemp.Fields("FCDacDisAmt").Value * pnCountRound) * (cRegPrice * nQtyCal)) / cSumCanDisc
                        Case "3"
                            If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
                                DE_SINbCal002 = False
                                Exit Function
                            End If
                            cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal) - (((porsTemp.Fields("FCDacDisAmt").Value * pnCountRound) * (cRegPrice * nQtyCal)) / cSumCanDisc)
                    End Select
                    
                    cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
                    '*Ao 52-04-20
                    cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
                    '*KT 52-02-19 member price
                    'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
                    '--------------------------------------------------------------------------------
                    bMemberPrice = False
                    If bDE_MemberPrice Then
                        cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                        cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                        tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                        '*KT 52-09-08 Member Price + Consignment
                        tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                        If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                            If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                                bMemberPrice = False
'                                If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                    'ยังไม่เคยถูกเรียกใช้
'                                Else
'                                    For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                        If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
''                                            typeDE_KeepOld(nI).FNDctNo = 0
'                                            Exit For
'                                        End If
'                                    Next nI
'                                End If
                            Else
                                bMemberPrice = True
                            End If
                        End If
                    End If
                    '*Ao 52-04-20/2
                    If nGrp = nMaxGrp Then
                         If nQtyGet < 1 Then
                            Call DE_SINxDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value * pnCountRound, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
                        End If
                    End If
                    
                    If Not bMemberPrice Then
                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                        Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                        If nQtyGet < 1 Then
                            nUncalGrp = nGrp
                        End If
                    Else
                        If tVB_StaMemPriceAlw = "Y" Then
                            If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                                bDE_MemberPriceAct = True
                                DE_SINbCal002 = True
                                Exit Function
                            Else
                                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                                Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                If nQtyGet < 1 Then
                                    nUncalGrp = nGrp
                                End If
                            End If
                        End If
                    End If
                    '--------------------------------------------------------------------------------
                End If
            End If
        End If
        porsTemp.MoveNext
    Loop
    DE_SINbCal002 = True
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    DE_SINbCal002 = False
End Function

'Private Function DE_SINbCal0020(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''   Call:
''   Cmt:    '*Ao 2011/05/11
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cDealDiscPer As Double, cDealDiscAmt As Double
'    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
'    Dim nQtyBuy As Integer
'    Dim tFreeFlag As String, tFilter As String, tSort As String
'    Dim nDTSeqNo As Integer
'    Dim nDTSeqNoNew As Integer
'    Dim cSumCanDisc As Double
'    Dim cRegPrice As Double
'    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer
'    Dim nMaxGrp As Integer
'
'    On Error GoTo Err_Handler
'
'    DE_SINbCal0020 = False
'    cDealDiscPer = 0
'    nUncalGrp = 0
'    nGrpOld = 0
'
'    If pnCountRound = 0 Then Exit Function
'
'    With porsTemp
'        'ทำตัว Free ก่อน
'         tFilter = "(FTPdhDealNo = '" & ptDealNo & "') AND (FNDacQtyGet > 0)"
'        .Filter = tFilter
'        tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
'        .Sort = tSort
'        If .EOF = False Then
'            .MoveFirst
'            nUncalGrp = 0
'            nGrpOld = 0
'            Do While Not .EOF
'                nGrp = .Fields("FNDacSkuGrp").Value
'                nDTSeqNo = .Fields("FNSdtSeqNo").Value
'                If InStr(1, tVB_Free0020, nDTSeqNo, vbBinaryCompare) > 0 Then
'                    If .Fields("FCSdtQtyAll").Value > 0 Then
'                        If nGrp <> nUncalGrp Then                                                                                                   'ตอนนี้มี FREE GROUP เดียว ทำเพื่อไว้ กรณีมีหลาย GROUP
'                            If nGrpOld <> nGrp Then
'                                nQtyGet = .Fields("FNDacQtyGet").Value * pnCountRound
'                                nGrpOld = nGrp
'                            End If
'                            If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                                nQtyCal = nQtyGet
'                                nQtyGet = 0
'                            Else
'                                nQtyCal = .Fields("FCSdtQtyAll").Value
'                                nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
'                            End If
'
'                            cRegPrice = .Fields("FCSdtRegPrice").Value
'                            ptLnDiscType = "4"
'                            tFreeFlag = "2"
'                            cDealDiscAmt = cRegPrice * nQtyCal
'                            nDTSeqNo = .Fields("FNSdtSeqNo").Value
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                            Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'                            If nQtyGet < 1 Then nUncalGrp = nGrp
'                        End If
'                    End If
'                End If
'                .MoveNext
'            Loop
'        Else
'            Exit Function
'        End If
'
''        ptLnDiscType = "1"
'        .Filter = adFilterNone
'        'ทำตัว Buy
'        tFilter = "FTPdhDealNo = '" & ptDealNo & "' AND FNDacQtyBuy >0"
'        .Filter = tFilter
'        tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
'        .Sort = tSort
'        .MoveFirst
'        Do While Not .EOF
'            nGrp = .Fields("FNDacSkuGrp").Value
'            If .Fields("FCSdtQtyAll").Value > 0 Then
'                If nGrp <> nUncalGrp Then
'                    If nGrpOld <> nGrp Then
'                        nQtyGet = .Fields("FNDacQtyBuy").Value * pnCountRound
'                        nGrpOld = nGrp
'                    End If
'                    If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                        nQtyCal = nQtyGet
'                        nQtyGet = 0
'                    Else
'                        nQtyCal = .Fields("FCSdtQtyAll").Value
'                        nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
'                    End If
'                    cRegPrice = .Fields("FCSdtRegPrice").Value
'                    cDealDiscAmt = 0
'                    tFreeFlag = "1"
''                    tFreeFlag = .Fields("FTFreeFlag").Value
'                    nDTSeqNo = .Fields("FNSdtSeqNo").Value
'                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
''                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'                    If nQtyGet < 1 Then nUncalGrp = nGrp
'                End If
'            End If
'            .MoveNext
'        Loop
'
'    End With
'
'    DE_SINbCal0020 = True
'    Exit Function
'Err_Handler:
'    DE_SINbCal0020 = False
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

Private Function DE_SINbCal0020(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 2011/05/11
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDealDiscPer As Double, cDealDiscAmt As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
    Dim nQtyBuy As Integer
    Dim tFreeFlag As String, tFilter As String, tSort As String, tFree0020 As String
    Dim nDTSeqNo As Integer
    Dim nDTSeqNoNew As Integer
    Dim cSumCanDisc As Double
    Dim cRegPrice As Double
    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer
    Dim nMaxGrp As Integer

    On Error GoTo Err_Handler

    DE_SINbCal0020 = False
    cDealDiscPer = 0
    nUncalGrp = 0
    nGrpOld = 0
    
    If pnCountRound = 0 Then Exit Function
    If tVB_Free0020 = "" Then Exit Function
    tFree0020 = tVB_Free0020
    
    With porsTemp
        'ทำตัว Free ก่อน
        nGrpOld = 0
        nUncalGrp = 0
        Do While tFree0020 <> ""
            nDTSeqNo = CInt(Trim(UT_tToken(tFree0020, ",", 1, False)))
            tFilter = "(FTPdhDealNo = '" & ptDealNo & "') AND (FNDacQtyGet > 0) AND (FNSdtSeqNo = " & nDTSeqNo & ")"
            .Filter = tFilter
            tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
            .Sort = tSort
            If .EOF = False Then
                nGrp = .Fields("FNDacSkuGrp").Value
                nDTSeqNo = .Fields("FNSdtSeqNo").Value
                If .Fields("FCSdtQtyAll").Value > 0 Then
                    If nGrp <> nUncalGrp Then                                                                                                   'ตอนนี้มี FREE GROUP เดียว ทำเพื่อไว้ กรณีมีหลาย GROUP
                        If nGrpOld <> nGrp Then
                            nQtyGet = .Fields("FNDacQtyGet").Value * pnCountRound
                            nGrpOld = nGrp
                        End If
                        If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                            nQtyCal = nQtyGet
                            nQtyGet = 0
                        Else
                            nQtyCal = .Fields("FCSdtQtyAll").Value
                            nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                        End If

                        cRegPrice = .Fields("FCSdtRegPrice").Value
                        ptLnDiscType = "4"
                        tFreeFlag = "2"
                        cDealDiscAmt = cRegPrice * nQtyCal
                        nDTSeqNo = .Fields("FNSdtSeqNo").Value
                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                        Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                        If nQtyGet < 1 Then nUncalGrp = nGrp
                    End If
                End If
            Else
                
            End If

        Loop
        
         nGrpOld = 0
        nUncalGrp = 0
        .Filter = adFilterNone
        'ทำตัว Buy
        tFilter = "FTPdhDealNo = '" & ptDealNo & "' AND (FNDacQtyBuy > 0)"
        .Filter = tFilter
        tSort = "FNDacSkuGrp,FCSdtRegPrice,FTSkuStaMallEmpCard"
        .Sort = tSort
        If Not .EOF Then
            .MoveFirst
        End If
        Do While Not .EOF
            nGrp = .Fields("FNDacSkuGrp").Value
            If .Fields("FCSdtQtyAll").Value > 0 Then
                If nGrp <> nUncalGrp Then
                    If nGrpOld <> nGrp Then
                        nQtyGet = .Fields("FNDacQtyBuy").Value * pnCountRound
                        nGrpOld = nGrp
                    End If
                    If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                        nQtyCal = nQtyGet
                        nQtyGet = 0
                    Else
                        nQtyCal = .Fields("FCSdtQtyAll").Value
                        nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                    End If
                    cRegPrice = .Fields("FCSdtRegPrice").Value
                    cDealDiscAmt = 0
                    tFreeFlag = "1"

                    nDTSeqNo = .Fields("FNSdtSeqNo").Value
                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)

                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
                    If nQtyGet < 1 Then nUncalGrp = nGrp
                End If
            End If
            .MoveNext
        Loop
        
    End With

    DE_SINbCal0020 = True
    Exit Function
Err_Handler:
    DE_SINbCal0020 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_SINbCal003(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 09-10-21
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDealDiscPer As Double, cDealDiscAmt As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer, nQtyBuy As Integer
    Dim tFreeFlag As String, tFilter As String, tSort As String
    Dim nDTSeqNo As Integer
    Dim nDTSeqNoNew As Integer
    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer
    Dim cSumCanDisc As Double
    Dim cRegPrice As Double
    Dim nPos As Integer
    Dim bChkround As Boolean
    '*KT 52-02-19 member price
    Dim bMemberPrice As Boolean
    Dim cB4DisChgOrg As Double
    Dim cMemberPrice As Double
    Dim tSKUCode As String
    Dim cDiscTotalAll As Double
    Dim tEventIDMemberPrice As String
    Dim nI As Long
On Error GoTo Err_Handler
    DE_SINbCal003 = False
    cDiscTotalAll = 0
    
    If pnCountRound = 0 Then Exit Function
    
    nMaxGrp = DE_GETnMaxGroup(ptDealNo, porsTemp)
    If Not (nMaxGrp = 2) Then Exit Function

'    For nGrp = 1 To nMaxGrp
    For nGrp = nMaxGrp To 1 Step -1                      '*Ao 2011/09/01    ให้สามารถซื้อสินค้า BUY และ GET ซ้ำกันได้
        bChkround = True
        porsTemp.Filter = adFilterNone
        tFilter = "FNDacSkuGrp = " & nGrp
        porsTemp.Filter = tFilter
        If nGrp = 1 Then
            tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
        Else
            tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
        End If
        porsTemp.Sort = tSort

        If Not (porsTemp.BOF And porsTemp.EOF) Then
            porsTemp.MoveFirst
            Do While Not porsTemp.EOF
                If nGrp = 1 Then
                    If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
'                        nGrp = porsTemp.Fields("FNDacSkuGrp").Value
                        cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
                        tSKUCode = porsTemp.Fields("FTDacSku").Value
                        nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
                        tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
                        If bChkround Then
                            nQtyBuy = porsTemp.Fields("FNDacQtyBuy").Value
                            nQtyBuy = nQtyBuy * pnCountRound
                            bChkround = False
                        End If
                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyBuy Then
                            nQtyCal = nQtyBuy
                            nQtyBuy = 0
                        Else
                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                            nQtyBuy = nQtyBuy - porsTemp.Fields("FCSdtQtyAll").Value
                        End If
'                        Select Case porsTemp.Fields("FTPddDisType").Value
'                            Case "1" 'Per
'                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
'                            Case "2", "3"
'                                cDealDiscPer = 0
'                            Case Else
'                                cDealDiscPer = 0
'                        End Select
                        cDealDiscPer = 0
                        cDealDiscAmt = 0
                        
                        bMemberPrice = False
                        If bDE_MemberPrice Then
                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                           '*KT 52-09-08 Member Price + Consignment
                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                            If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                                If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                                    bMemberPrice = False
                                    If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
                                        'ยังไม่เคยถูกเรียกใช้
                                    Else
                                        For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
                                            If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                typeDE_KeepOld(nI).FNDctNo = 0
                                                Exit For
                                            End If
                                        Next nI
                                    End If
                                Else
                                    bMemberPrice = True
                                End If
                            End If
                        End If
    
                        If Not bMemberPrice Then
                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                            Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", porsTemp.Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
    '                        nGrpOld = nGrp
                            If nQtyBuy < 1 Then
    '                            nUncalGrp = nGrp
                                Exit Do
                            End If
                        Else
                            If tVB_StaMemPriceAlw = "Y" Then
                                 If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                                    bDE_MemberPriceAct = True
                                Else
                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", porsTemp.Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
    '                                nGrpOld = nGrp
                                    If nQtyBuy < 1 Then
    '                                    nUncalGrp = nGrp
                                        Exit Do
                                    End If
                                End If
                            End If
                        End If
                        
                    End If              'If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
                Else        'If nI = 1 Then                 'Get
                    If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
'                        nGrp = porsTemp.Fields("FNDacSkuGrp").Value
                        cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
                        tSKUCode = porsTemp.Fields("FTDacSku").Value
                        nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
                        tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
                        ptLnDiscType = porsTemp.Fields("FTPddDisType").Value
                        If bChkround = True Then
                            nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
                            nQtyGet = nQtyGet * pnCountRound
'                            nQtyGetOrg = nQtyGet
                            bChkround = False
                        End If
'                        Select Case ptLnDiscType
'                            Case "1" 'Per
'                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
'                            Case "2", "3"
'                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
'                            Case Else
'                                cDealDiscPer = 0
'                        End Select
                        Select Case ptLnDiscType
                            Case "2", "3"
                                If cSumCanDisc <= 0 Then
                                    nPos = porsTemp.AbsolutePosition
                                    cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, nGrp, tSort)
                                    porsTemp.Filter = tFilter
                                    porsTemp.Sort = tSort
                                    porsTemp.AbsolutePosition = nPos
                                End If
                        End Select
                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
                            nQtyCal = nQtyGet
                            nQtyGet = 0
                        Else
                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                            nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
                        End If
                        Select Case ptLnDiscType
                            Case "1" 'Per
                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
                                cDealDiscAmt = cRegPrice * (cDealDiscPer / 100) * nQtyCal
                            Case "2"
                                If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
                                    DE_SINbCal003 = False
                                    Exit Function
                                End If
                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
                                cDealDiscAmt = (porsTemp.Fields("FCDacDisAmt").Value * pnCountRound) * ((cRegPrice * nQtyCal) / cSumCanDisc)
                            Case "3"
                                If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
                                    DE_SINbCal003 = False
                                    Exit Function
                                End If
                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
                                cDealDiscAmt = (cRegPrice * nQtyCal) - (((porsTemp.Fields("FCDacDisAmt").Value * pnCountRound) * (cRegPrice * nQtyCal)) / cSumCanDisc)
                            Case Else
                                cDealDiscPer = 0
                                cDealDiscAmt = 0
                        End Select
                        
                        cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt)
                        cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)

                        bMemberPrice = False
                        If bDE_MemberPrice Then
                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                           '*KT 52-09-08 Member Price + Consignment
                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                            If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                                If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                                    bMemberPrice = False
                                    If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
                                        'ยังไม่เคยถูกเรียกใช้
                                    Else
                                        For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
                                            If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                typeDE_KeepOld(nI).FNDctNo = 0
                                                Exit For
                                            End If
                                        Next nI
                                    End If
                                Else
                                    bMemberPrice = True
                                End If
                            End If
                        End If
                        '*Ao 52-04-20
                        If nQtyGet = 0 Then
                            Call DE_SINxDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value * pnCountRound, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
                        End If
                        If Not bMemberPrice Then
                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                            Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", porsTemp.Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'                            nGrpOld = nGrp
                            If nQtyGet < 1 Then
'                                nUncalGrp = nGrp
                                Exit Do
                            End If
                        Else
                            If tVB_StaMemPriceAlw = "Y" Then
                                 If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                                    bDE_MemberPriceAct = True
                                Else
                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                    Call DE_RSTxUpdateBySku(porsTemp, nDTSeqNo, "FCSdtQtyAll", porsTemp.Fields("FCSdtQtyAll").Value - nQtyCal, tFilter, tSort)
'                                    nGrpOld = nGrp
                                    If nQtyGet < 1 Then
'                                        nUncalGrp = nGrp
                                        Exit Do
                                    End If
                                End If
                            End If
                        End If
                        
                    End If
                End If
                porsTemp.MoveNext
            Loop
        Else
            Call SP_MSGxShowError(Err.Number, Err.Description)
            Exit Function
        End If      'If Not (.BOF And .EOF) Then
        
    Next nGrp
    DE_SINbCal003 = True
    Exit Function
Err_Handler:
    DE_SINbCal003 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Sub DE_SINxCal004(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset)
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cDealDiscPer As Double, cDealDiscAmt As Double
'    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
'    Dim tFreeFlag As String
'    Dim nDTSeqNo As Integer
'    Dim nDTSeqNoNew As Integer
'    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer
'    Dim cSumCanDisc As Double
'    Dim cRegPrice As Double
'    Dim nPos As Integer
'    '*KT 52-02-19 member price
'    Dim bMemberPrice As Boolean
'    Dim cB4DisChgOrg As Double
'    Dim cMemberPrice As Double
'    Dim tSKUCode As String
'    Dim cDiscTotalAll As Double
'    Dim tEventIDMemberPrice As String
'    Dim nI As Long
'On Error GoTo Err_Handler
'    cDiscTotalAll = 0
'    nMaxGrp = DE_GETnMaxGroup(ptDealNo, porsTemp)
'    porsTemp.MoveFirst
'    If nMaxGrp = 3 Then
'        Do While Not porsTemp.EOF
'            If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
'                nGrp = porsTemp.Fields("FNDacSkuGrp").Value
'                cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
'                tSKUCode = porsTemp.Fields("FTDacSku").Value
'                If nGrp <> nUncalGrp Then
'                    nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
'                    If nGrp <> nGrpOld Then
'                        Select Case porsTemp.Fields("FTPddDisType").Value
'                            Case "1" 'Per
'                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
'                            Case "2", "3"
'                                cDealDiscPer = 0
'                            Case Else
'                                cDealDiscPer = 0
'                        End Select
'                        tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
'                        Select Case tFreeFlag
'                            Case "1" 'Buy Only
'                                nQtyGet = porsTemp.Fields("FNDacQtyBuy").Value
'                            Case "2" 'Get or Free Only
'                                nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
'                            Case "3" 'Both (Get or Free) And Buy
'                                nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
'                        End Select
'
'                        nQtyGetOrg = nQtyGet
'                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                            nQtyCal = nQtyGet
'                            nQtyGet = 0
'                        Else
'                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                            nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                        End If
'                        If nGrp = 3 Then
'                            nPos = porsTemp.AbsolutePosition
'                            cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGetOrg, nGrp) '*TW 51-06-27
'                            porsTemp.AbsolutePosition = nPos
'                            Select Case porsTemp.Fields("FTPddDisType").Value
'                                Case "1" 'Per
'                                    cDealDiscAmt = porsTemp.Fields("FCSdtRegPrice").Value * (cDealDiscPer / 100) * nQtyCal
'                                Case "2"
'                                    cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc)
'                                Case "3"
'                                    cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc))
'                                Case "4"
'                                    cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal)
'                            End Select
'                        Else
'                            cDealDiscAmt = 0
'                        End If
'                        '*Ao 52-04-20
'                        cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
'                        cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
'                          '*KT 52-02-19 member price
'                        'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
'                        '--------------------------------------------------------------------------------
'                        bMemberPrice = False
'                        If bDE_MemberPrice Then
'                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                            '*KT 52-09-08 Member Price + Consignment
'                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                            If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
'                                If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
'                                    bMemberPrice = False
'                                    If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                        'ยังไม่เคยถูกเรียกใช้
'                                    Else
'                                        For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                            If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                typeDE_KeepOld(nI).FNDctNo = 0
'                                                Exit For
'                                            End If
'                                        Next nI
'                                    End If
'                                Else
'                                    bMemberPrice = True
'                                End If
'                            End If
'                        End If
'                        '*Ao 52-04-20
'                        Call DE_SINcDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value, cDiscTotalAll, cDealDiscAmt)
'
'                        If Not bMemberPrice Then
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                            nGrpOld = nGrp
'                            If nQtyGet <= 0 Then
'                                nUncalGrp = nGrp
'                            End If
'                        Else
'                            If tVB_StaMemPriceAlw = "Y" Then
'                                If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                                    bDE_MemberPriceAct = True
'                                Else
'                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                    nGrpOld = nGrp
'                                    If nQtyGet <= 0 Then
'                                        nUncalGrp = nGrp
'                                    End If
'                                End If
'                            End If
'                        End If
'                        '------------------------------------------------------------------------------------
'                    Else
'                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                            nQtyCal = nQtyGet
'                            nQtyGet = 0
'                        Else
'                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                            nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                        End If
'                        If nGrp = 3 Then
'                            nPos = porsTemp.AbsolutePosition
'                            'cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, nGrp) '*TW 51-06-27
'                            porsTemp.AbsolutePosition = nPos
'                            Select Case porsTemp.Fields("FTPddDisType").Value
'                                Case "1" 'Per
'                                    cDealDiscAmt = IIf(nGrp = 1, 0, porsTemp.Fields("FCSdtRegPrice").Value * (cDealDiscPer / 100) * nQtyCal)
'                                Case "2"
'                                    cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc)
'                                Case "3"
'                                    cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc))
'                                Case "4"
'                                    cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal)
'                            End Select
'                        Else
'                            cDealDiscAmt = 0
'                        End If
'                        '*Ao 52-04-20
'                        cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
'                        cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
'                        '*KT 52-02-19 member price
'                        'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
'                        '--------------------------------------------------------------------------------
'                        bMemberPrice = False
'                        If bDE_MemberPrice Then
'                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                             '*KT 52-09-08 Member Price + Consignment
'                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                            If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
'                                If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
'                                    bMemberPrice = False
'                                    If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                        'ยังไม่เคยถูกเรียกใช้
'                                    Else
'                                        For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                            If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                typeDE_KeepOld(nI).FNDctNo = 0
'                                                Exit For
'                                            End If
'                                        Next nI
'                                    End If
'                                Else
'                                    bMemberPrice = True
'                                End If
'                            End If
'                        End If
'                        '*Ao 52-04-20
'                        Call DE_SINcDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value, cDiscTotalAll, cDealDiscAmt)
'
'                        If Not bMemberPrice Then
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                            If nQtyGet <= 0 Then
'                                nUncalGrp = nGrp
'                            End If
'                        Else
'                            If tVB_StaMemPriceAlw = "Y" Then
'                                If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                                    bDE_MemberPriceAct = True
'                                Else
'                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                    If nQtyGet <= 0 Then
'                                        nUncalGrp = nGrp
'                                    End If
'                                End If
'                            End If
'                        End If
'                        '---------------------------------------------------------------------------------
'                    End If
'                End If
'            End If
'            porsTemp.MoveNext
'        Loop
'    End If
'    Exit Sub
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub

'Private Sub DE_SINxCal004(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset)
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''   Call:
''   Cmt:    '*Ao 09-10-22
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cDealDiscPer As Double, cDealDiscAmt As Double
'    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer, nQtyBuy As Integer
'    Dim tFreeFlag As String, tFilter As String, tSort As String
'    Dim nDTSeqNo As Integer
'    Dim nDTSeqNoNew As Integer
'    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer
'    Dim cSumCanDisc As Double
'    Dim cRegPrice As Double
'    Dim nPos As Integer
'    '*KT 52-02-19 member price
'    Dim bMemberPrice As Boolean
'    Dim cB4DisChgOrg As Double
'    Dim cMemberPrice As Double
'    Dim tSKUCode As String
'    Dim cDiscTotalAll As Double
'    Dim tEventIDMemberPrice As String
'    Dim nI As Long
'    Dim bChkround As Boolean
'On Error GoTo Err_Handler
'    cDiscTotalAll = 0
'    nMaxGrp = DE_GETnMaxGroup(ptDealNo, porsTemp)
'    If Not (nMaxGrp = 3) Then Exit Sub
'
'    For nGrp = 1 To nMaxGrp
'        bChkround = True
'        porsTemp.Filter = adFilterNone
'        tFilter = "FNDacSkuGrp = " & nGrp
'        porsTemp.Filter = tFilter
'        If nGrp = 3 Then
'            tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
'        Else
'            tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
'        End If
'        porsTemp.Sort = tSort
'
'        If Not (porsTemp.BOF And porsTemp.EOF) Then
'            porsTemp.MoveFirst
'            Do While Not porsTemp.EOF
'                If Not (nGrp = 3) Then                'Group 1,2
'                    If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
''                        nGrp = porsTemp.Fields("FNDacSkuGrp").Value
'                        cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
'                        tSKUCode = porsTemp.Fields("FTDacSku").Value
'                        nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
'                        tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
'                        If bChkround Then
'                            nQtyBuy = porsTemp.Fields("FNDacQtyBuy").Value
'                            bChkround = False
'                        End If
'                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyBuy Then
'                            nQtyCal = nQtyBuy
'                            nQtyBuy = 0
'                        Else
'                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                            nQtyBuy = nQtyBuy - porsTemp.Fields("FCSdtQtyAll").Value
'                        End If
'
'                        cDealDiscPer = 0
'                        cDealDiscAmt = 0
'
'                        bMemberPrice = False
'                        If bDE_MemberPrice Then
'                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                           '*KT 52-09-08 Member Price + Consignment
'                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                            If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
'                                If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
'                                    bMemberPrice = False
'                                    If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                        'ยังไม่เคยถูกเรียกใช้
'                                    Else
'                                        For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                            If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                typeDE_KeepOld(nI).FNDctNo = 0
'                                                Exit For
'                                            End If
'                                        Next nI
'                                    End If
'                                Else
'                                    bMemberPrice = True
'                                End If
'                            End If
'                        End If
'
'                        If Not bMemberPrice Then
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                            If nQtyBuy <= 0 Then    'ครบ Group
'                                Exit Do
'                            End If
'                        Else
'                            If tVB_StaMemPriceAlw = "Y" Then
'                                 If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                                    bDE_MemberPriceAct = True
'                                Else
'                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                    If nQtyBuy <= 0 Then                'ครบ Group
'                                        Exit Do
'                                    End If
'                                End If
'                            End If
'                        End If
'
'                    End If              'If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
'                Else        'If Not (nGrp = 3) Then                'Get
'                    If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
''                        nGrp = porsTemp.Fields("FNDacSkuGrp").Value
'                        cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
'                        tSKUCode = porsTemp.Fields("FTDacSku").Value
'                        nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
'                        tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
'                        ptLnDiscType = porsTemp.Fields("FTPddDisType").Value
'
'                        If bChkround = True Then
'                            nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
''                            nQtyGetOrg = nQtyGet
'                            bChkround = False
'                        End If
'                        Select Case ptLnDiscType
'                            Case "2", "3"
'                                If cSumCanDisc <= 0 Then
'                                    nPos = porsTemp.AbsolutePosition
'                                    cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, nGrp, tSort)
'                                    porsTemp.Filter = tFilter
'                                    porsTemp.Sort = tSort
'                                    porsTemp.AbsolutePosition = nPos
'                                End If
'                        End Select
'                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                            nQtyCal = nQtyGet
'                            nQtyGet = 0
'                        Else
'                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                            nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                        End If
'                        Select Case ptLnDiscType
'                            Case "1" 'Per
'                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
'                                cDealDiscAmt = (cRegPrice * (cDealDiscPer / 100)) * nQtyCal
'                            Case "2"
'                                cDealDiscPer = 0
'                                cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc)
'                            Case "3"
'                                cDealDiscPer = 0
'                                cDealDiscAmt = (cRegPrice * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc))
''                            Case "4"
''                                cDealDiscPer = 0
''                                cDealDiscAmt = (cRegPrice * nQtyCal)
'                            Case Else
'                                cDealDiscPer = 0
'                                cDealDiscAmt = 0
'                        End Select
'
'                        cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt)
'                        cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
'
'                        bMemberPrice = False
'                        If bDE_MemberPrice Then
'                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                           '*KT 52-09-08 Member Price + Consignment
'                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                            If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
'                                If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
'                                    bMemberPrice = False
'                                    If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                        'ยังไม่เคยถูกเรียกใช้
'                                    Else
'                                        For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                            If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                typeDE_KeepOld(nI).FNDctNo = 0
'                                                Exit For
'                                            End If
'                                        Next nI
'                                    End If
'                                Else
'                                    bMemberPrice = True
'                                End If
'                            End If
'                        End If
'                        '*Ao 52-04-20
'                        If nQtyGet = 0 Then
'                            Call DE_SINxDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
'                        End If
'                        If Not bMemberPrice Then
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
''                            nGrpOld = nGrp
'                            If nQtyGet <= 0 Then
''                                nUncalGrp = nGrp
'                                Exit Do
'                            End If
'                        Else
'                            If tVB_StaMemPriceAlw = "Y" Then
'                                 If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                                    bDE_MemberPriceAct = True
'                                Else
'                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
''                                    nGrpOld = nGrp
'                                    If nQtyGet <= 0 Then
''                                        nUncalGrp = nGrp
'                                        Exit Do
'                                    End If
'                                End If
'                            End If
'                        End If
'
'                    End If
'                End If
'                porsTemp.MoveNext
'            Loop
'        Else
'            Call SP_MSGxShowError(Err.Number, Err.Description)
'            Exit Sub
'        End If      'If Not (.BOF And .EOF) Then
'
'
'    Next nGrp
'    Exit Sub
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub

Private Function DE_SINbCal004(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 09-11-11
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDealDiscPer As Double, cDealDiscAmt As Double, cSumCanDisc As Double, cRegPrice As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer, nQtyBuy As Integer, nDTSeqNo As Integer, nDTSeqNoNew As Integer, nPos As Integer
    Dim tFreeFlag As String, tFilter As String, tSort As String
    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer, nLoop As Integer, nChkMaxGrp As Integer
    '*KT 52-02-19 member price
    Dim bMemberPrice As Boolean
    Dim cB4DisChgOrg As Double
    Dim cMemberPrice As Double
    Dim tSKUCode As String
    Dim cDiscTotalAll As Double
    Dim tEventIDMemberPrice As String
    Dim nI As Long
    Dim bChkround As Boolean
On Error GoTo Err_Handler
    DE_SINbCal004 = False
    cDiscTotalAll = 0
    nUncalGrp = 0
    nGrpOld = 0
    nMaxGrp = 2
'    nChkMaxGrp = DE_GETnMaxGroup(ptDealNo, porsTemp)
    With porsTemp
        For nLoop = 1 To nMaxGrp
            porsTemp.Filter = adFilterNone
            
            If nLoop = 1 Then
                tFilter = "FNDacQtyGet = 0"             'BUY
                tSort = "FNDacSkuGrp,FTSkuStaMallEmpCard,FCSdtRegPrice"
            Else
                tFilter = "FNDacQtyBuy = 0"             'GET
                tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
            End If
            .Filter = tFilter
            .Sort = tSort

            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    nGrp = .Fields("FNDacSkuGrp").Value
                    If .Fields("FCSdtQtyAll").Value > 0 Then
                        If nGrp <> nUncalGrp Then
                            cRegPrice = .Fields("FCSdtRegPrice").Value
                            tSKUCode = .Fields("FTDacSku").Value
                            nDTSeqNo = .Fields("FNSdtSeqNo").Value
                            tFreeFlag = .Fields("FTFreeFlag").Value
                            ptLnDiscType = .Fields("FTPddDisType").Value
                            If nGrp <> nGrpOld Then
                                Select Case tFreeFlag
                                    Case "1" 'Buy Only
                                        nQtyGet = .Fields("FNDacQtyBuy").Value * CInt(pnCountRound)
                                    Case "2" 'Get or Free Only
                                        nQtyGet = .Fields("FNDacQtyGet").Value * CInt(pnCountRound)
                                    Case "3" 'Both (Get or Free) And Buy
                                        nQtyGet = .Fields("FNDacQtyGet").Value * CInt(pnCountRound)
                                End Select
'                                nQtyGetOrg = nQtyGet
                                If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                                    nQtyCal = nQtyGet
                                    nQtyGet = 0
                                Else
                                    nQtyCal = .Fields("FCSdtQtyAll").Value
                                    nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                                End If
                                If nLoop = 2 Then
                                    Select Case ptLnDiscType
                                        Case "2", "3"
                                            If cSumCanDisc <= 0 Then
                                                nPos = .AbsolutePosition
    '                                            cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGetOrg, nGrp, tSort)
                                                cSumCanDisc = DE_GETcSumDiscByQtyGrp(porsTemp, nQtyGet, nGrp, tSort, pnCountRound)
                                                .Filter = tFilter
                                                .Sort = tSort
                                                .AbsolutePosition = nPos
                                            End If
                                    End Select
                                    Select Case ptLnDiscType
                                        Case "1" 'Per
                                            cDealDiscPer = .Fields("FCDacDisAmt").Value
                                            cDealDiscAmt = cRegPrice * (cDealDiscPer / 100) * nQtyCal
                                        Case "2"
                                            cDealDiscPer = .Fields("FCDacDisAmt").Value
                                            cDealDiscAmt = (.Fields("FCDacDisAmt").Value * pnCountRound) * ((cRegPrice * nQtyCal) / cSumCanDisc)
                                        Case "3"
                                            cDealDiscPer = .Fields("FCDacDisAmt").Value
                                            cDealDiscAmt = (cRegPrice * nQtyCal) - ((.Fields("FCDacDisAmt").Value * pnCountRound) * ((cRegPrice * nQtyCal) / cSumCanDisc))
                                        Case Else
                                            cDealDiscPer = 0
                                            cDealDiscAmt = 0
                                    End Select
                                    cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt)
                                    cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
                                Else
                                    cDealDiscPer = 0
                                    cDealDiscAmt = 0
                                End If
                                
                                bMemberPrice = False
                                If bDE_MemberPrice Then
                                    cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                                    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                                    tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                                   '*KT 52-09-08 Member Price + Consignment
                                    tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                                    If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                                        If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                                            bMemberPrice = False
                                            If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
                                                'ยังไม่เคยถูกเรียกใช้
                                            Else
                                                For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
                                                    If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                        typeDE_KeepOld(nI).FNDctNo = 0
                                                        Exit For
                                                    End If
                                                Next nI
                                            End If
                                        Else
                                            bMemberPrice = True
                                        End If
                                    End If
                                End If
                                '*Ao 52-04-20
                                If nGrp = nChkMaxGrp Then
                                    If nLoop = 2 Then
                                        If nQtyGet = 0 Then
                                            Call DE_SINxDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value * pnCountRound, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
                                        End If
                                    End If
                                End If
                                '__________________________________________
                                If Not bMemberPrice Then
                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, .Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                    nGrpOld = nGrp
                                    If nQtyGet < 1 Then 'ครบ Group
                                        nUncalGrp = nGrp
                                    End If
                                Else
                                    If tVB_StaMemPriceAlw = "Y" Then
                                         If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                                            bDE_MemberPriceAct = True
                                        Else
                                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, .Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                            nGrpOld = nGrp
                                            If nQtyGet < 1 Then               'ครบ Group
                                                nUncalGrp = nGrp
                                            End If
                                        End If
                                    End If
                                End If
                                
                            Else        'If nGrp <> nGrpOld Then
'                                nGrp = .Fields("FNDacSkuGrp").Value
'                                cRegPrice = .Fields("FCSdtRegPrice").Value
'                                tSKUCode = .Fields("FTDacSku").Value
'                                nDTSeqNo = .Fields("FNSdtSeqNo").Value
'                                tFreeFlag = .Fields("FTFreeFlag").Value
'                                ptLnDiscType = .Fields("FTPddDisType").Value
                                
                                If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                                    nQtyCal = nQtyGet
                                    nQtyGet = 0
                                Else
                                    nQtyCal = .Fields("FCSdtQtyAll").Value
                                    nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                                End If
                                If nLoop = 2 Then
                                    Select Case ptLnDiscType
                                        Case "1" 'Per
                                            cDealDiscPer = .Fields("FCDacDisAmt").Value
                                            cDealDiscAmt = (cRegPrice * (cDealDiscPer / 100)) * nQtyCal
                                        Case "2"
                                            If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
                                                DE_SINbCal004 = False
                                                Exit Function
                                            End If
                                            cDealDiscPer = .Fields("FCDacDisAmt").Value
                                            cDealDiscAmt = (.Fields("FCDacDisAmt").Value * pnCountRound) * ((cRegPrice * nQtyCal) / cSumCanDisc)
                                        Case "3"
                                            If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
                                                DE_SINbCal004 = False
                                                Exit Function
                                            End If
                                            cDealDiscPer = .Fields("FCDacDisAmt").Value
                                            cDealDiscAmt = (cRegPrice * nQtyCal) - ((.Fields("FCDacDisAmt").Value * pnCountRound) * ((cRegPrice * nQtyCal) / cSumCanDisc))
                                        Case Else
                                            cDealDiscPer = 0
                                            cDealDiscAmt = 0
                                    End Select
                                    cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt)
                                    cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
                                Else
                                    cDealDiscPer = 0
                                    cDealDiscAmt = 0
                                End If
                                '__________________________________________
                                bMemberPrice = False
                                If bDE_MemberPrice Then
                                    cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                                    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                                    tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                                   '*KT 52-09-08 Member Price + Consignment
                                    tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                                    If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                                        If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                                            bMemberPrice = False
                                            If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
                                                'ยังไม่เคยถูกเรียกใช้
                                            Else
                                                For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
                                                    If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                        typeDE_KeepOld(nI).FNDctNo = 0
                                                        Exit For
                                                    End If
                                                Next nI
                                            End If
                                        Else
                                            bMemberPrice = True
                                        End If
                                    End If
                                End If
                                '*Ao 52-04-20
                                If nGrp = nChkMaxGrp Then
                                    If nLoop = 2 Then
                                        If nQtyGet = 0 Then
                                            Call DE_SINxDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value * pnCountRound, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
                                        End If
                                    End If
                                End If
                                '__________________________________________
                                If Not bMemberPrice Then
                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                    nGrpOld = nGrp
                                    If nQtyGet < 1 Then
                                        nUncalGrp = nGrp
                                    End If
                                Else
                                    If tVB_StaMemPriceAlw = "Y" Then
                                         If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                                            bDE_MemberPriceAct = True
                                        Else
                                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                            nGrpOld = nGrp
                                            If nQtyGet < 1 Then
                                                nUncalGrp = nGrp
                                            End If
                                        End If
                                    End If
                                End If
                                '__________________________________________
                            End If          'If nGrp <> nGrpOld Then
                        End If      'If nGrp <> nUncalGrp Then
                    End If          'If .Fields("FCSdtQtyAll").Value > 0 Then
                    .MoveNext
                Loop
            End If      'If Not (.BOF And .EOF) Then
        Next nLoop
    End With
    DE_SINbCal004 = True
    Exit Function
Err_Handler:
    DE_SINbCal004 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_SINbCal005(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDealDiscPer As Double, cDealDiscAmt As Double, cAmtGet As Double, cAmtGetOrg As Double, cAmtCal As Double
    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer, nDTSeqNo As Integer, nDelGrp As Integer, nDelSdtSeqNo As Integer
    Dim tFreeFlag As String, tSort As String, tFilter As String
    Dim nDTSeqNoNew As Integer, nQtyCal As Integer, nPos As Integer
    Dim cDelDisAmt As Double, cDelRegPrice As Double, cDelSaleAmt As Double, cDelAmtAvi As Double, cQtyAll As Double, cSumCanDisc As Double
    '*KT 52-02-19 member price
    Dim bMemberPrice As Boolean
    Dim cB4DisChgOrg As Double
    Dim cMemberPrice As Double
    Dim tSKUCode As String
    Dim cDiscTotalAll As Double, cSumTotal As Double
    Dim tEventIDMemberPrice As String
    Dim nI As Long
On Error GoTo Err_Handler

    DE_SINbCal005 = False
    nMaxGrp = DE_GETnMaxGroup(ptDealNo, porsTemp)
    If Not (nMaxGrp = 1) Then Exit Function
'    tFilter = "FTPdhDealNo ='" & ptDealNo & "'"
'    porsTemp.Filter = tFilter
    tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
    porsTemp.Sort = tSort
    porsTemp.MoveFirst
    If nMaxGrp = 1 Then
        cAmtGet = porsTemp.Fields("FCDacAmtBuy").Value
        cAmtGetOrg = cAmtGet
        Do While Not porsTemp.EOF
            cQtyAll = porsTemp.Fields("FCSdtQtyAll").Value
            If cQtyAll > 0 Then
                cDelRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
                tSKUCode = porsTemp.Fields("FTDacSku").Value
                cDelSaleAmt = cDelRegPrice * cQtyAll
                cDelDisAmt = porsTemp.Fields("FCDacDisAmt").Value
                'nGrp = porsTemp.Fields("FNDacSkuGrp").Value
                nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
                cDealDiscPer = cDelDisAmt
'                Select Case porsTemp.Fields("FTPddDisType").Value
'                    Case "1" 'Per
'                        cDealDiscPer = cDelDisAmt
'                    Case "2", "3"
'                        cDealDiscPer = 0
'                    Case Else
'                        cDealDiscPer = 0
'                End Select
                tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
                
                If cDelSaleAmt >= cAmtGet Then
                    cAmtCal = cAmtGet
                    cAmtGet = 0
                Else
                    cAmtCal = cDelSaleAmt
                    cAmtGet = cAmtGet - cDelSaleAmt
                End If
'                nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal Mod cDelRegPrice = 0, 0, 1)             '*Ao 09-10-20
                nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal \ cDelRegPrice = cAmtCal / cDelRegPrice, 0, 1)
'                Select Case ptLnDiscType
'                    Case "2", "3"
                If cSumCanDisc <= 0 Then
                    nPos = porsTemp.AbsolutePosition
                    cSumCanDisc = DE_GETcSumDiscByAmt(porsTemp, cAmtGetOrg, tFilter, tSort)
'                            tFilter = "FNDacSkuGrp = " & nGrp
'                            porsTemp.Filter = tFilter
                    porsTemp.Sort = tSort
                    porsTemp.AbsolutePosition = nPos
                End If
'                End Select
                
                Select Case porsTemp.Fields("FTPddDisType").Value
                    Case "1" 'Per
'                        cDealDiscAmt = (cDealDiscPer / 100) * cAmtCal
                        cDelDisAmt = (cDelDisAmt / 100) * cAmtGetOrg
                        cDealDiscAmt = ((cDelRegPrice * nQtyCal) * cDelDisAmt) / cSumCanDisc
                    Case "2"
                        If cDelDisAmt > cAmtGetOrg Then
                            DE_SINbCal005 = False
                            Exit Function
                        End If
                        cDealDiscAmt = (cDelDisAmt * (cDelRegPrice * nQtyCal)) / cSumCanDisc
'                        cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc)
                    Case "3"
                        If cDelDisAmt > cAmtGetOrg Then
                            DE_SINbCal005 = False
                            Exit Function
                        End If
                        cDealDiscAmt = ((cDelRegPrice * nQtyCal) * (cAmtGetOrg - cDelDisAmt)) / cSumCanDisc     '*Ao 2011/05/11 CommSheet ML-PRV-056 Deal Type 5 not Print  Deal Amount .xls
'                        cDealDiscAmt = (cDelRegPrice * nQtyCal) - ((cAmtGetOrg - cDelDisAmt) / cSumCanDisc)
'                        cDealDiscAmt = (cRegPrice * nQtyCal) - (((porsTemp.Fields("FCDacDisAmt").Value * pnCountRound) * (cRegPrice * nQtyCal)) / cSumCanDisc)
                End Select
                '*Ao 52-04-20
                cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
'                cSumTotal = cSumTotal + (cDelRegPrice * nQtyCal)
                '*KT 52-02-19 member price
                'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
                '--------------------------------------------------------------------------------
                bMemberPrice = False
                If bDE_MemberPrice Then
                    cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                    tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                    '*KT 52-09-08 Member Price + Consignment
                    tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                    If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                        If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                            bMemberPrice = False
'                            If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                'ยังไม่เคยถูกเรียกใช้
'                            Else
'                                For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                    If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
''                                        typeDE_KeepOld(nI).FNDctNo = 0
'                                        Exit For
'                                    End If
'                                Next nI
'                            End If
                        Else
                            bMemberPrice = True
                        End If
                    End If
                End If
                '*Ao 52-04-20
                If cAmtGet <= 0 Then
                    Call DE_SINxDealAmt(ptLnDiscType, cAmtGet, cDelDisAmt, cDiscTotalAll, cDealDiscAmt, cAmtGetOrg, ptDealType)
                End If
                If Not bMemberPrice Then
                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cAmtCal)
                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                Else
                    If tVB_StaMemPriceAlw = "Y" Then
                        If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                            bDE_MemberPriceAct = True
                            DE_SINbCal005 = True
                            Exit Function
                        Else
                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cAmtCal)
                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                        End If
                    End If
                End If
                '-----------------------------------------------------------------------------------------------------------------------
                If cAmtGet <= 0 Then
                    Exit Do
                End If
            End If
            porsTemp.MoveNext
        Loop
    End If
    DE_SINbCal005 = True
    Exit Function
Err_Handler:
    DE_SINbCal005 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Sub DE_SINxCal006(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset)
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
''CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
''DESCRIPTION       #
''ARGUMENT LIST #
''
''RETURN VALUE  # None
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cDealDiscPer As Double, cDealDiscAmt As Double
'    Dim cAmtGet As Double, cAmtGetOrg As Double, cAmtCal As Double
'    Dim tFreeFlag As String
'    Dim nDTSeqNo As Integer
'    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer
'
'    '*ช่วยให้ดูง่ายเอาตัวแปรมารับก่อน
'    Dim nDelGrp As Integer
'    Dim nDelSdtSeqNo As Integer
'    Dim nDTSeqNoNew As Integer
'    Dim cDelDisAmt As Double
'    Dim cDelRegPrice As Double
'    Dim cDelSaleAmt As Double
'    Dim cQtyAll As Double
'    Dim nPos As Integer
'    Dim cSumCanDisc As Double
'    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
'    '*KT 52-02-19 member price
'    Dim bMemberPrice As Boolean
'    Dim cB4DisChgOrg As Double
'    Dim cMemberPrice As Double
'    Dim tSKUCode As String
'    Dim cDiscTotalAll As Double
'    Dim tEventIDMemberPrice As String
'    Dim nI As Long
'On Error GoTo Err_Handler
'    cDiscTotalAll = 0
'    nMaxGrp = DE_GETnMaxGroup(ptDealNo, porsTemp)
'    porsTemp.MoveFirst
'    If nMaxGrp = 2 Then
'        Do While Not porsTemp.EOF
'            nGrp = porsTemp.Fields("FNDacSkuGrp").Value
'            cQtyAll = porsTemp.Fields("FCSdtQtyAll").Value
'            If nGrp = 1 Then 'หา Amt
'                If cQtyAll > 0 Then
'                    cDelRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
'                    cQtyAll = porsTemp.Fields("FCSdtQtyAll").Value
'                    cDelSaleAmt = cQtyAll * cDelRegPrice
'                    cDelDisAmt = porsTemp.Fields("FCDacDisAmt").Value
'                    tSKUCode = porsTemp.Fields("FTDacSku").Value
'                    If nGrp <> nUncalGrp Then
'                        nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
'                        If nGrp <> nGrpOld Then
'                            cDealDiscPer = 0
'                            tFreeFlag = "1"
'                            cAmtGet = porsTemp.Fields("FCDacAmtBuy").Value
'                            cAmtGetOrg = cAmtGet
'                            If cDelSaleAmt >= cAmtGet Then
'                                cAmtCal = cAmtGet
'                                cAmtGet = 0
'                            Else
'                                cAmtCal = cDelSaleAmt
'                                cAmtGet = cAmtGet - cDelSaleAmt
'                            End If
'                            cDealDiscAmt = 0
'
'    '                        nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal Mod cDelRegPrice = 0, 0, 1)          '*Ao 09-10-20    Mod คำนวณผิด เช่น 186 Mod 46.5 แล้วได้เศษ 2 ซึ่งผิด
'                            nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal / cDelRegPrice = cAmtCal \ cDelRegPrice, 0, 1)
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cAmtCal)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'
'                            nGrpOld = nGrp
'                            If cAmtGet <= 0 Then
'                                nUncalGrp = nGrp
'                            End If
'                        Else
'                            If cDelSaleAmt >= cAmtGet Then
'                                cAmtCal = cAmtGet
'                                cAmtGet = 0
'                            Else
'                                cAmtCal = cDelSaleAmt
'                                cAmtGet = cAmtGet - cDelSaleAmt
'                            End If
'                            cDealDiscAmt = 0
'                            cDealDiscPer = 0
'    '                        nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal Mod cDelRegPrice = 0, 0, 1)          '*Ao 09-10-20    Mod คำนวณผิด เช่น 186 Mod 46.5 แล้วได้เศษ 2 ซึ่งผิด
'                            nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal / cDelRegPrice = cAmtCal \ cDelRegPrice, 0, 1)
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cAmtCal)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                            If cAmtGet <= 0 Then
'                                nUncalGrp = nGrp
'                            End If
'                        End If
'                    End If
'                End If
'            Else 'Group = 2
'                If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
'                    cDelRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
'                    nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
'                    If nGrp <> nGrpOld Then
'                        Select Case porsTemp.Fields("FTPddDisType").Value
'                            Case "1" 'Per
'                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
'                            Case "2", "3"
'                                cDealDiscPer = 0
'                            Case Else
'                                cDealDiscPer = 0
'                        End Select
'                        tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
'                        nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
'                        nQtyGetOrg = nQtyGet
'
'                        Select Case porsTemp.Fields("FTPddDisType").Value
'                            Case "2", "3"
'                                If cSumCanDisc <= 0 Then
'                                    nPos = porsTemp.AbsolutePosition
'                                    cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, nGrp) '*TW 51-06-27
''                                    tFilter = "FNDacSkuGrp = " & nRound
''                                    porsTemp.Filter = tFilter
''                                    tSort = "FCSdtRegPrice DESC"
''                                    porsTemp.Sort = tSort
'                                    porsTemp.AbsolutePosition = nPos
'                                End If
'                        End Select
'                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                            nQtyCal = nQtyGet
'                            nQtyGet = 0
'                        Else
'                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                            nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                        End If
'                        Select Case porsTemp.Fields("FTPddDisType").Value
'                            Case "1" 'Per
'                                cDealDiscAmt = cDelRegPrice * (cDealDiscPer / 100) * nQtyCal
'                            Case "2"
'                                cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc)
'                            Case "3"
'                                cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc))
'                        End Select
'                        '*Ao 52-04-20
'                        cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
'                        cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
'                        '*KT 52-02-19 member price
'                        'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
'                        '--------------------------------------------------------------------------------
'                        bMemberPrice = False
'                        If bDE_MemberPrice Then
'                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                            '*KT 52-09-08 Member Price + Consignment
'                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                            If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
'                                If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
'                                    bMemberPrice = False
'                                    If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                        'ยังไม่เคยถูกเรียกใช้
'                                    Else
'                                        For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                            If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                typeDE_KeepOld(nI).FNDctNo = 0
'                                                Exit For
'                                            End If
'                                        Next nI
'                                    End If
'                                Else
'                                    bMemberPrice = True
'                                End If
'                            End If
'                        End If
'                        '*Ao 52-04-20
'                        Call DE_SINcDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value, cDiscTotalAll, cDealDiscAmt)
'
'                        If Not bMemberPrice Then
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cDelRegPrice * nQtyCal)
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                            nGrpOld = nGrp
'                            If nQtyGet <= 0 Then
'                                nUncalGrp = nGrp
'                            End If
'                        Else
'                            If tVB_StaMemPriceAlw = "Y" Then
'                                If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                                    bDE_MemberPriceAct = True
'                                Else
'                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cDelRegPrice * nQtyCal)
'                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                    nGrpOld = nGrp
'                                    If nQtyGet <= 0 Then
'                                        nUncalGrp = nGrp
'                                    End If
'                                End If
'                            End If
'                        End If
'                        '----------------------------------------------------------------------------
'                    Else
'                        If nQtyGet > 0 Then    '*Ao 52-03-24                 แก้ไข nQtyGet=0 ไม่ให้เข้าทำงาน เพราะจะ INSERT Qty=0
'                        '________________________________________________________________
'                            If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                                nQtyCal = nQtyGet
'                                nQtyGet = 0
'                            Else
'                                nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                                nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                            End If
'                            Select Case porsTemp.Fields("FTPddDisType").Value
'                                Case "1" 'Per
'                                    cDealDiscAmt = cDelRegPrice * (cDealDiscPer / 100) * nQtyCal
'                                Case "2"
'                                    cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc)
'                                Case "3"
'                                    cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc))
'                            End Select
'                            '*Ao 52-04-20
'                            cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
'                            cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
'
'                            '*KT 52-02-19 member price
'                            'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
'                            '--------------------------------------------------------------------------------
'                            bMemberPrice = False
'                            If bDE_MemberPrice Then
'                                cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                                cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                                tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                                '*KT 52-09-08 Member Price + Consignment
'                                tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                                If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
'                                    If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
'                                        bMemberPrice = False
'                                        If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                            'ยังไม่เคยถูกเรียกใช้
'                                        Else
'                                            For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                                If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                    typeDE_KeepOld(nI).FNDctNo = 0
'                                                    Exit For
'                                                End If
'                                            Next nI
'                                        End If
'                                    Else
'                                        bMemberPrice = True
'                                    End If
'                                End If
'                            End If
'                            '*Ao 52-04-20
'                            Call DE_SINcDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value, cDiscTotalAll, cDealDiscAmt)
'
'                            If Not bMemberPrice Then
'                                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cDelRegPrice * nQtyCal)
'                                Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                If nQtyGet <= 0 Then
'                                    nUncalGrp = nGrp
'                                End If
'                            Else
'                                If tVB_StaMemPriceAlw = "Y" Then
'                                    If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                                        bDE_MemberPriceAct = True
'                                    Else
'                                         nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cDelRegPrice * nQtyCal)
'                                        Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                        If nQtyGet <= 0 Then
'                                            nUncalGrp = nGrp
'                                        End If
'                                    End If
'                                End If
'                            End If
'                        End If
'                         '----------------------------------------------------------------------------
'                    End If         '*Ao 52-03-24
'                End If
'            End If
'            porsTemp.MoveNext
'        Loop
'    End If
'
'    Exit Sub
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub

Private Function DE_SINbCal006(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 09-10-26
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDealDiscPer As Double, cDealDiscAmt As Double
    Dim cAmtGet As Double, cAmtGetOrg As Double, cAmtCal As Double
    Dim tFreeFlag As String, tFilter As String, tSort As String
    Dim nDTSeqNo As Integer
    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer
    Dim bChkround As Boolean
    '*ช่วยให้ดูง่ายเอาตัวแปรมารับก่อน
    Dim nDelGrp As Integer
    Dim nDelSdtSeqNo As Integer
    Dim nDTSeqNoNew As Integer
    Dim cDelDisAmt As Double
    Dim cDelRegPrice As Double
    Dim cDelSaleAmt As Double
    Dim cQtyAll As Double
    Dim nPos As Integer
    Dim cSumCanDisc As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
    '*KT 52-02-19 member price
    Dim bMemberPrice As Boolean
    Dim cB4DisChgOrg As Double
    Dim cMemberPrice As Double
    Dim tSKUCode As String
    Dim cDiscTotalAll As Double
    Dim tEventIDMemberPrice As String
    Dim nI As Long
On Error GoTo Err_Handler
    DE_SINbCal006 = False
    cDiscTotalAll = 0
    nMaxGrp = DE_GETnMaxGroup(ptDealNo, porsTemp)
    If Not (nMaxGrp = 2) Then Exit Function

    For nGrp = 1 To nMaxGrp
        bChkround = True
        porsTemp.Filter = adFilterNone
        tFilter = "FNDacSkuGrp = " & nGrp
        porsTemp.Filter = tFilter
        If nGrp = 2 Then
            tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
        Else
'            tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
            tSort = "FTSkuStaMallEmpCard,FCSdtSalePricePmt"
        End If
        porsTemp.Sort = tSort

        If Not (porsTemp.BOF And porsTemp.EOF) Then
            porsTemp.MoveFirst
            Do While Not porsTemp.EOF
                If nGrp = 1 Then                'Group 1
                    If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
'                        cDelRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
                        cDelRegPrice = porsTemp.Fields("FCSdtSalePricePmt").Value
                        cQtyAll = porsTemp.Fields("FCSdtQtyAll").Value
                        cDelSaleAmt = cQtyAll * cDelRegPrice
                        cDelDisAmt = porsTemp.Fields("FCDacDisAmt").Value
                        tSKUCode = porsTemp.Fields("FTDacSku").Value
                        nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
                        cDealDiscPer = 0
                        cDealDiscAmt = 0
                        tFreeFlag = "1"
                        
                        If bChkround Then
                            cAmtGet = porsTemp.Fields("FCDacAmtBuy").Value
                            bChkround = False
                        End If
                        
                        If cDelSaleAmt >= cAmtGet Then
                            cAmtCal = cAmtGet
                            cAmtGet = 0
                        Else
                            cAmtCal = cDelSaleAmt
                            cAmtGet = cAmtGet - cDelSaleAmt
                        End If
                                           
'                        nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal Mod cDelRegPrice = 0, 0, 1)          '*Ao 09-10-20    Mod คำนวณผิด เช่น 186 Mod 46.5 แล้วได้เศษ 2 ซึ่งผิด
                        nQtyCal = cAmtCal \ cDelRegPrice + IIf(cAmtCal / cDelRegPrice = cAmtCal \ cDelRegPrice, 0, 1)
                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cAmtCal)
                        Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                        
'                        nGrpOld = nGrp
                        If cAmtGet <= 0 Then
'                            nUncalGrp = nGrp
                            Exit Do
                        End If
                    End If          'If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
                Else        'Group 2
                    If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
                        cDelRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
                        nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
                        tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
                        ptLnDiscType = porsTemp.Fields("FTPddDisType").Value
'                        Select Case porsTemp.Fields("FTPddDisType").Value
'                            Case "1" 'Per
'                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
'                            Case "2", "3"
'                                cDealDiscPer = 0
'                            Case Else
'                                cDealDiscPer = 0
'                        End Select
                        If bChkround Then
                            nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
'                            nQtyGetOrg = nQtyGet
                            bChkround = False
                        End If
                        Select Case ptLnDiscType
                            Case "2", "3"
                                If cSumCanDisc <= 0 Then
                                    nPos = porsTemp.AbsolutePosition
                                    cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, nGrp, tSort) '*TW 51-06-27
                                    porsTemp.Filter = tFilter
                                    porsTemp.Sort = tSort
                                    porsTemp.AbsolutePosition = nPos
                                End If
                        End Select
                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
                            nQtyCal = nQtyGet
                            nQtyGet = 0
                        Else
                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                            nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
                        End If
                        Select Case ptLnDiscType
                            Case "1" 'Per
                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
                                cDealDiscAmt = cDelRegPrice * (cDealDiscPer / 100) * nQtyCal
                            Case "2"
                                If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
                                    DE_SINbCal006 = False
                                    Exit Function
                                End If
                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
                                cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc)
'                                cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc)
                            Case "3"
                                If porsTemp.Fields("FCDacDisAmt").Value > cSumCanDisc Then
                                    DE_SINbCal006 = False
                                    Exit Function
                                End If
                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
                                cDealDiscAmt = (cDelRegPrice * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc))
'                                cDealDiscAmt = (cDelRegPrice * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cDelRegPrice * nQtyCal) / cSumCanDisc))
                            Case Else
                                cDealDiscPer = 0
                                cDealDiscAmt = 0
                        End Select
                        
                        cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
                        '*Ao 52-04-20
                        cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
                        '*KT 52-02-19 member price
                        'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
                        '--------------------------------------------------------------------------------
                        bMemberPrice = False
                        If bDE_MemberPrice Then
                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                            '*KT 52-09-08 Member Price + Consignment
                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                            If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                                If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                                    bMemberPrice = False
                                    If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
                                        'ยังไม่เคยถูกเรียกใช้
                                    Else
                                        For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
                                            If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
'                                                typeDE_KeepOld(nI).FNDctNo = 0
                                                Exit For
                                            End If
                                        Next nI
                                    End If
                                Else
                                    bMemberPrice = True
                                End If
                            End If
                        End If
                        '*Ao 52-04-20
                        If nQtyGet < 1 Then
                            Call DE_SINxDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
                        End If
                        
                        If Not bMemberPrice Then
                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cDelRegPrice * nQtyCal)
                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                            If nQtyGet <= 0 Then
                                Exit Do
                            End If
                        Else
                            If tVB_StaMemPriceAlw = "Y" Then
                                If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                                    bDE_MemberPriceAct = True
                                Else
                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cDelRegPrice * nQtyCal)
                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                    If nQtyGet <= 0 Then
                                        Exit Do
                                    End If
                                End If
                            End If
                        End If
                    End If          'If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
                End If      'If nGrp = 1 Then
                porsTemp.MoveNext
            Loop
        Else
            Call SP_MSGxShowError(Err.Number, Err.Description)
            Exit Function
        End If
    Next nGrp
    DE_SINbCal006 = True
    Exit Function
Err_Handler:
    DE_SINbCal006 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_SINbCal007(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tFreeFlag As String
    Dim nDTSeqNo As Integer
    Dim nDTSeqNoNew As Integer
    Dim cDealDiscPer As Double, cDealDiscAmt As Double
    Dim nMaxLev As Integer
    Dim cQtySumAll As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
    Dim nLevel As Integer
    Dim cRegPrice As Double
    '*KT 52-02-19 member price
    Dim bMemberPrice As Boolean
    Dim cB4DisChgOrg As Double
    Dim cMemberPrice As Double
    Dim tSKUCode As String
    Dim tEventIDMemberPrice As String
    Dim nI As Long
On Error GoTo Err_Handler
'    nMaxLev = DE_GETnMaxGroup(ptDealNo, porsTemp, True)        '*Ao 52-05-29   ไม่ได้ใช้
    DE_SINbCal007 = False
    porsTemp.Filter = "FNPddLevel = 1"
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        If porsTemp.Fields("FNPddLevel").Value = 1 Then
            cQtySumAll = cQtySumAll + porsTemp.Fields("FCSdtQtyAll").Value
        End If
        porsTemp.MoveNext
    Loop
    '*Ao 52-05-29
    porsTemp.Filter = adFilterNone
    porsTemp.Sort = "FNSdtSeqNo,FNPddLevel DESC"
    '_________________________________________
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        '*Ao 09-10-19
'        If porsTemp.Fields("FNDacQtyGet").Value <= cQtySumAll Then
'            nLevel = porsTemp.Fields("FNPddLevel").Value
'            'nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
'            nQtyGet = cQtySumAll '*TW 51-08-08 'แก้ Comsheet 438
'            Exit Do
'        End If
        If porsTemp.Fields("FNDacQtyBuy").Value <= cQtySumAll Then
            nLevel = porsTemp.Fields("FNPddLevel").Value
            'nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
            nQtyGet = cQtySumAll '*TW 51-08-08 'แก้ Comsheet 438
            Exit Do
        End If
        porsTemp.MoveNext
    Loop
    nQtyGetOrg = nQtyGet
'    porsTemp.Sort = "FCSdtRegPrice DESC"
'    porsTemp.Sort = "FNSdtSeqNo,FNPddLevel DESC"
    porsTemp.MoveFirst
    Do While Not porsTemp.EOF
        If (porsTemp.Fields("FCSdtQtyAll").Value) > 0 And (porsTemp.Fields("FNPddLevel").Value = nLevel) Then
            cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
            nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
            cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
            tSKUCode = porsTemp.Fields("FTDacSku").Value
            tFreeFlag = "3"
'*TW 51-08-08 'แก้ Comsheet 438
            If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
                'nQtyCal = nQtyGet 'Old
                nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                nQtyGet = 0
            Else
                nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
            End If

            cDealDiscAmt = cRegPrice * (cDealDiscPer / 100) * nQtyCal
            '*KT 52-02-19 member price
            'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
            '--------------------------------------------------------------------------------
            bMemberPrice = False
            If bDE_MemberPrice Then
                cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                '*KT 52-09-08 Member Price + Consignment
                tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                    If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                        bMemberPrice = False
'                        If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                            'ยังไม่เคยถูกเรียกใช้
'                        Else
'                            For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
''                                    typeDE_KeepOld(nI).FNDctNo = 0
'                                    Exit For
'                                End If
'                            Next nI
'                        End If
                    Else
                        bMemberPrice = True
                    End If
                End If
            End If
            If Not bMemberPrice Then
                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
            Else
                If tVB_StaMemPriceAlw = "Y" Then
                    If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                        bDE_MemberPriceAct = True
                        DE_SINbCal007 = True
                        Exit Function
                    Else
                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                        Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                    End If
                End If
            End If
            '------------------------------------------------------------------------------------
            If nQtyGet <= 0 Then
                Exit Do
            End If
        End If
        porsTemp.MoveNext
    Loop
    DE_SINbCal007 = True
    Exit Function
Err_Handler:
    DE_SINbCal007 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_SINbCal008(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tFreeFlag As String
    Dim nDTSeqNo As Integer
    Dim nDTSeqNoNew As Integer
    Dim cDealDiscPer As Double, cDealDiscAmt As Double
    Dim cAmtSumAll As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
    Dim nLevel As Integer
    Dim cRegPrice As Double
    Dim cQtyAll As Double
    Dim cDelSaleAmt As Double
    Dim cAmtGet As Double, cAmtCal As Double, cAmtGetOrg As Double
    '*KT 52-02-19 member price
    Dim bMemberPrice As Boolean
    Dim cB4DisChgOrg As Double
    Dim cMemberPrice As Double
    Dim tSKUCode As String
    Dim tEventIDMemberPrice As String
    Dim nI As Long
On Error GoTo Err_Handler
'    Call SP_RSTbClone2Offline(porsTemp, orsSinDeal
    DE_SINbCal008 = False
    With porsTemp
        .MoveFirst
        Do While Not .EOF
            If .Fields("FNPddLevel").Value = 1 Then
                cAmtSumAll = cAmtSumAll + .Fields("FCSdtQtyAll").Value * .Fields("FCSdtRegPrice").Value
            End If
            .MoveNext
        Loop
        
        .Sort = "FNPddLevel DESC"
        .MoveFirst
        Do While Not .EOF
            If .Fields("FCDacAmtBuy").Value <= cAmtSumAll Then
                nLevel = .Fields("FNPddLevel").Value
                'cAmtGet = .Fields("FCDacAmtBuy").Value
                cAmtGet = cAmtSumAll '*TW 51-08-08 'แก้ Comsheet 438
                Exit Do
            End If
            .MoveNext
        Loop
        cAmtGetOrg = cAmtGet
        '*Ao 52-05-18
'        .Sort = "FCSdtRegPrice DESC"
        .Sort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
        .MoveFirst
        
        Do While Not .EOF
            If (porsTemp.Fields("FCSdtQtyAll").Value) > 0 And (porsTemp.Fields("FNPddLevel").Value = nLevel) Then
                cRegPrice = .Fields("FCSdtRegPrice").Value
                cQtyAll = .Fields("FCSdtQtyAll").Value
                cDelSaleAmt = cQtyAll * cRegPrice
                nDTSeqNo = .Fields("FNSdtSeqNo").Value
                cDealDiscPer = .Fields("FCDacDisAmt").Value
                tSKUCode = .Fields("FTDacSku").Value
                tFreeFlag = "3"
'*TW 51-08-08 'แก้ Comsheet 438
                If cDelSaleAmt >= cAmtGet Then
                    'cAmtCal = cAmtGet Old
                    cAmtCal = cDelSaleAmt
                    cAmtGet = 0
                Else
                    cAmtCal = cDelSaleAmt
                    cAmtGet = cAmtGet - cDelSaleAmt
                End If
                
                cDealDiscAmt = cDealDiscPer / 100 * cAmtCal
'                nQtyCal = cAmtCal \ cRegPrice + IIf(cAmtCal Mod cRegPrice = 0, 0, 1)           '*Ao 09-10-20
                nQtyCal = cAmtCal \ cRegPrice + IIf(cAmtCal \ cRegPrice = cAmtCal / cRegPrice, 0, 1)
                
                  '*KT 52-02-19 member price
                'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
                '--------------------------------------------------------------------------------
                bMemberPrice = False
                If bDE_MemberPrice Then
                    cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                    tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                    '*KT 52-09-08 Member Price + Consignment
                    tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                    If tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "" Then
                        If CDbl(SP_FMTcCurPoint(((cB4DisChgOrg * nQtyCal) - cDealDiscAmt), True, False)) < CDbl(SP_FMTcCurPoint(cMemberPrice * nQtyCal, True, False)) Then
                            bMemberPrice = False
'                            If (Not typeDE_KeepOld) = True Then 'เช็คว่า Array เคยถูกเรียกใช้งานหรือยัง
'                                'ยังไม่เคยถูกเรียกใช้
'                            Else
'                                For nI = LBound(typeDE_KeepOld) To UBound(typeDE_KeepOld)
'                                    If (typeDE_KeepOld(nI).FNDctNo = nEN_Disc21MPrice Or typeDE_KeepOld(nI).FNDctNo = nEN_Disc22MDisc) And typeDE_KeepOld(nI).FNSdtSeqNo = nDTSeqNo Then
''                                        typeDE_KeepOld(nI).FNDctNo = 0
'                                        Exit For
'                                    End If
'                                Next nI
'                            End If
                        Else
                            bMemberPrice = True
                        End If
                    End If
                End If
                If Not bMemberPrice Then
                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cAmtCal)
                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                Else
                    If tVB_StaMemPriceAlw = "Y" Then
                        If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                            bDE_MemberPriceAct = True
                            DE_SINbCal008 = True
                            Exit Function
                        Else
                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cAmtCal)
                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                        End If
                    End If
                End If
                '------------------------------------------------------------------------------------
                If cAmtGet <= 0 Then
                    Exit Do
                End If
            End If
            .MoveNext
        Loop
    End With
    DE_SINbCal008 = True
    Exit Function
Err_Handler:
    DE_SINbCal008 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_UPDxQtySingle(ByRef porsTemp As ADODB.Recordset, ByVal pnQtyCal As Integer, ByVal pnDTSeqNo As Integer, Optional ByVal pcAmtCal As Double)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nPos As Long
    On Error GoTo Err_Handler
    With porsTemp
            nPos = .AbsolutePosition
            Do While Not .EOF
                If .Fields("FNSdtSeqNo").Value = pnDTSeqNo Then
                    .Fields("FCSdtQtyAll").Value = .Fields("FCSdtQtyAll").Value - pnQtyCal
                End If
                .MoveNext
            Loop
            .AbsolutePosition = nPos
        
    End With
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function DE_GETcSumDiscByQty(ByRef porsTemp As ADODB.Recordset, ByVal pnQtyGet As Integer, ByVal pnGrp As Integer, Optional ByVal ptSort As String) As Double
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim nQtyCal As Integer
    Dim cTemp As Double
    Dim nQtyGet As Integer
    On Error GoTo Err_Handler
    With porsTemp
        nQtyGet = pnQtyGet
        .Filter = "FNDacSkuGrp = " & pnGrp
        If .BOF And .EOF Then Exit Function
        If Not (ptSort = "") Then                   '*Ao 09-11-05
            .Sort = ptSort
        End If
        '________________________________________
        .MoveFirst
        Do While Not .EOF
            If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                nQtyCal = nQtyGet
                nQtyGet = 0
            Else
                nQtyCal = .Fields("FCSdtQtyAll").Value
                nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
            End If
            cTemp = cTemp + (.Fields("FCSdtRegPrice").Value * nQtyCal)
            If nQtyGet <= 0 Then Exit Do
            .MoveNext
        Loop
        .Filter = adFilterNone
    End With
    DE_GETcSumDiscByQty = cTemp
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub DE_MULxSendDeal(ByVal pnI As Integer)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tDealNo As String, tFreeFlag As String, tDealType As String, tEventID As String
    Dim tDiscType As String, cDisPer As Double, nDTSeqNo As Integer, nQtyCal As Integer, cDisAmt As Double, cAmtCal As Double
    Dim cRegPrice As Double
    Dim tChkSeqNo As String
    Dim tChkQty As String
    Dim tChkSta As String
    Dim tChkReg As String
    Dim tKeepSeqNo As String
    Dim tKeepQty As String
    Dim tKeepSta As String
    Dim tKeepRegPrice As String
    Dim cDiscTotalAll As Double
    
    On Error GoTo Err_Handler
    cDiscTotalAll = 0
    nDE_Group = nDE_Group + 1 '*TW 51-08-11
    tKeepSeqNo = typeDE_MutiDeal(pnI).tDTSeqNo
    tKeepQty = typeDE_MutiDeal(pnI).tDiscQty
    tKeepSta = typeDE_MutiDeal(pnI).tSkuFlag
    tKeepRegPrice = typeDE_MutiDeal(pnI).tRegPrice
    tChkSeqNo = UT_tToken(tKeepSeqNo, ",", 1, False)
    tChkQty = UT_tToken(tKeepQty, ",", 1, False)
    tChkSta = UT_tToken(tKeepSta, ",", 1, False)
    tChkReg = UT_tToken(tKeepRegPrice, ",", 1, False)
    
    Do While tChkSeqNo <> ""
        nDTSeqNo = Val(tChkSeqNo)
        tDealNo = typeDE_MutiDeal(pnI).tDealNo
        nQtyCal = Val(tChkQty)
        cRegPrice = Val(tChkReg)
        tDealType = typeDE_MutiDeal(pnI).tDealType
        tEventID = typeDE_MutiDeal(pnI).tEventID
        tDiscType = typeDE_MutiDeal(pnI).tDiscType
        
        If tChkSta = "1" Then
            Select Case tDealType                       '*Ao 09-12-01
                Case "0010", "0020", "0030", "0040"
                    tFreeFlag = "2"
                    tDiscType = "4"                                                         '*Ao 2012/12/25
                Case Else
                    tFreeFlag = typeDE_MutiDeal(pnI).tFreeFlag
            End Select
            
'            If typeDE_MutiDeal(pnI).tDealType = "008" Then
'                cDisAmt = typeDE_MutiDeal(pnI).cDiscAmt
'            Else
'                cDisAmt = ((cRegPrice * nQtyCal) / typeDE_MutiDeal(pnI).cAlwDisc) * typeDE_MutiDeal(pnI).cDiscAmt
'            End If
            Select Case typeDE_MutiDeal(pnI).tDealType          '*Ao 2012/12/18   Deal 008 ได้มา 2 Deal
                Case "007", "008"
                    cDisAmt = typeDE_MutiDeal(pnI).cDiscPer / 100 * (cRegPrice * nQtyCal)
                Case "0020"
                    cDisAmt = (cRegPrice * nQtyCal)
                Case Else
                    cDisAmt = ((cRegPrice * nQtyCal) / typeDE_MutiDeal(pnI).cAlwDisc) * typeDE_MutiDeal(pnI).cDiscAmt
            End Select
        Else
            Select Case tDealType                       '*Ao 09-12-01
                Case "0010", "0020", "0030", "0040"
                    tFreeFlag = "1"
                    cDisAmt = 0
                    tDiscType = "0"                             '*Ao 2012/12/25
                Case Else
                    tFreeFlag = "1"
                    cDisAmt = 0
            End Select
        End If
        
        If tDiscType = "1" Then
            cDisPer = typeDE_MutiDeal(pnI).cDiscPer
        Else
            cDisPer = 0
        End If
        '*Ao 52-05-18
        cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDisAmt, True, False)
        tChkSeqNo = UT_tToken(tKeepSeqNo, ",", 1, False)
        If tChkSeqNo = "" Then
            Call DE_MulxDealAmt(tDiscType, 0, typeDE_MutiDeal(pnI).cAlwDisc, cDiscTotalAll, cDisAmt, typeDE_MutiDeal(pnI).cDiscAmt, tDealType)
        End If
        
        cAmtCal = (cRegPrice * nQtyCal)
        '*KT 53-04-17 Communication Sheet PH3-ML-SF-111 Deal performance optimizing.xls
        If pnI = UBound(typeDE_MutiDeal) And tChkSeqNo = "" Then
            Call DE_GETcCallDeal(tDealNo, tFreeFlag, tDealType, tDealType, tEventID, tDiscType, cDisPer, nDTSeqNo, nQtyCal, cDisAmt, cAmtCal, True)
        Else
            Call DE_GETcCallDeal(tDealNo, tFreeFlag, tDealType, tDealType, tEventID, tDiscType, cDisPer, nDTSeqNo, nQtyCal, cDisAmt, cAmtCal, False)
        End If
        tChkQty = UT_tToken(tKeepQty, ",", 1, False)
        tChkSta = UT_tToken(tKeepSta, ",", 1, False)
        tChkReg = UT_tToken(tKeepRegPrice, ",", 1, False)
    Loop
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DE_CALxMultiByDealNo0010(ByVal ptDealNo As String, ByVal ptDealType As String, ByVal ptSeqSku As String, ByVal ptQtySku As String, ByVal ptChkSta As String, ByVal ptWhe As String, ByRef pnIndex As Integer)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tChkSeqNo As String
    Dim tChkQty As String
    Dim tChkSta As String
    Dim tKeepSeqNo As String
    Dim tKeepQty As String
    Dim tKeepSta As String
    Dim tSKU As String
    Dim tQty As String
    Dim tSta As String
    
    Dim cRegPrice As Double
    Dim cDiscAmt As Double
    Dim cQtyGet As Double
    Dim cQtyBuy As Double
    Dim cAmtBuy As Double
    Dim tDiscType As String
    Dim cSumAlwDisc As Double
    Dim cSumDisc As Double
    Dim tSumRegPrice As String
    Dim nSeqDTOrg As Integer
    Dim nMaxSeqNoOld As Integer
    
    Dim cMaxRegPrice As Double
    Dim nMaxSeqNo As Integer
    Dim nNewSeqNo As Integer
    Dim bFound As Boolean
    On Error GoTo Err_Handler
    tKeepSeqNo = ptSeqSku
    tKeepQty = ptQtySku
    tKeepSta = ptChkSta
    
    tChkSeqNo = UT_tToken(tKeepSeqNo, ",", 1, False)
    tChkQty = UT_tToken(tKeepQty, ",", 1, False)
    ReDim Preserve typeDE_MutiDeal(pnIndex) As UDT_MultiDeal
    Do While tChkSeqNo <> ""
        nSeqDTOrg = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & tChkSeqNo)) 'หา Seq หลัง Split
        orsDE_AFSack.Filter = "FTPdhDealNo = '" & ptDealNo & "' AND FNSdtSeqNo = " & tChkSeqNo
        cRegPrice = orsDE_AFSack.Fields("FCSdtRegPrice").Value
        If cMaxRegPrice < cRegPrice Then
            cMaxRegPrice = cRegPrice
            nMaxSeqNo = nSeqDTOrg
            nMaxSeqNoOld = Val(tChkSeqNo)
        End If
        tChkSeqNo = UT_tToken(tKeepSeqNo, ",", 1, False)
        tChkQty = UT_tToken(tKeepQty, ",", 1, False)
    Loop
    
    tKeepSeqNo = ptSeqSku
    tKeepQty = ptQtySku
    tKeepSta = ptChkSta
    tChkSeqNo = UT_tToken(tKeepSeqNo, ",", 1, False)
    tChkQty = UT_tToken(tKeepQty, ",", 1, False)
    tChkSta = UT_tToken(tKeepSta, ",", 1, False)
    Do While tChkSeqNo <> ""
        nSeqDTOrg = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & tChkSeqNo)) 'หา Seq หลัง Split
        If (nMaxSeqNo = nSeqDTOrg) And (Not bFound) Then
            orsDE_AFSack.Filter = "FTPdhDealNo = '" & ptDealNo & "' AND FNSdtSeqNo = " & tChkSeqNo
            cRegPrice = orsDE_AFSack.Fields("FCSdtRegPrice").Value
            bFound = True
            If Val(tChkQty) > 1 Then
                If DE_DELbSplitDT(nSeqDTOrg, Val(tChkQty) - 1, tVB_CNTblDT, tVB_CNTerminalNum, tVB_Running, nNewSeqNo) Then 'สั่ง Split
                    tSKU = IIf(tSKU = "", "", tSKU & ",") & nSeqDTOrg
                    tSKU = IIf(tSKU = "", "", tSKU & ",") & nNewSeqNo
                    tQty = IIf(tQty = "", "", tQty & ",") & Trim(Str(Val(tChkQty) - 1))
                    tQty = IIf(tQty = "", "", tQty & ",") & "1"
                    tSta = IIf(tSta = "", "", tSta & ",") & "0"
                    tSta = IIf(tSta = "", "", tSta & ",") & "1"
                    tSumRegPrice = IIf(tSumRegPrice = "", "", tSumRegPrice & ",") & Trim(Str(cRegPrice))
                    tSumRegPrice = IIf(tSumRegPrice = "", "", tSumRegPrice & ",") & Trim(Str(cRegPrice))
                    Call DE_DELbSplitDT(nNewSeqNo, 1, tVB_CNTblDT, tVB_CNTerminalNum, tVB_Running, nNewSeqNo)
                End If
            Else
                    tSKU = IIf(tSKU = "", "", tSKU & ",") & nSeqDTOrg
                    tQty = IIf(tQty = "", "", tQty & ",") & "1"
                    tSta = IIf(tSta = "", "", tSta & ",") & "1"
                    tSumRegPrice = IIf(tSumRegPrice = "", "", tSumRegPrice & ",") & Trim(Str(cRegPrice))
                    Call DE_DELbSplitDT(nNewSeqNo, 1, tVB_CNTblDT, tVB_CNTerminalNum, tVB_Running, nNewSeqNo)
            End If
            cQtyGet = orsDE_AFSack.Fields("FNDacQtyGet").Value
            cQtyBuy = orsDE_AFSack.Fields("FNDacQtyBuy").Value
            cAmtBuy = orsDE_AFSack.Fields("FCDacAmtBuy").Value
            cDiscAmt = orsDE_AFSack.Fields("FCDacDis").Value
            tDiscType = orsDE_AFSack.Fields("FTPddDisType").Value
            cSumAlwDisc = (cRegPrice)
            typeDE_MutiDeal(pnIndex).tFreeFlag = "2"
        Else
            Call DE_DELbSplitDT(nSeqDTOrg, Val(tChkQty), tVB_CNTblDT, tVB_CNTerminalNum, tVB_Running, nNewSeqNo)
            orsDE_AFSack.Filter = "FTPdhDealNo = '" & ptDealNo & "' AND FNSdtSeqNo = " & tChkSeqNo
            cRegPrice = orsDE_AFSack.Fields("FCSdtRegPrice").Value
            cQtyGet = orsDE_AFSack.Fields("FNDacQtyGet").Value
            cQtyBuy = orsDE_AFSack.Fields("FNDacQtyBuy").Value
            cAmtBuy = orsDE_AFSack.Fields("FCDacAmtBuy").Value
            cDiscAmt = orsDE_AFSack.Fields("FCDacDisAmt").Value
            tDiscType = orsDE_AFSack.Fields("FTPddDisType").Value
            tSKU = IIf(tSKU = "", "", tSKU & ",") & nSeqDTOrg
            tQty = IIf(tQty = "", "", tQty & ",") & tChkQty
            '*Ao 09-10-16
'            tSta = IIf(tSta = "", "", tSta & ",") & "0"
            tSta = IIf(tSta = "", "", tSta & ",") & "1"
            tSumRegPrice = IIf(tSumRegPrice = "", "", tSumRegPrice & ",") & Trim(Str(cRegPrice))
        End If
        tChkSeqNo = UT_tToken(tKeepSeqNo, ",", 1, False)
        tChkQty = UT_tToken(tKeepQty, ",", 1, False)
        tChkSta = UT_tToken(tKeepSta, ",", 1, False)
    Loop
    typeDE_MutiDeal(pnIndex).cDiscPer = 0
    cSumDisc = cSumAlwDisc
    typeDE_MutiDeal(pnIndex).cAlwDisc = cSumAlwDisc
    typeDE_MutiDeal(pnIndex).cDiscAmt = cSumDisc
    typeDE_MutiDeal(pnIndex).tDealNo = ptDealNo
    typeDE_MutiDeal(pnIndex).tDealType = ptDealType
    typeDE_MutiDeal(pnIndex).tDiscQty = tQty
    typeDE_MutiDeal(pnIndex).tDiscType = tDiscType
    typeDE_MutiDeal(pnIndex).tDTSeqNo = tSKU
    typeDE_MutiDeal(pnIndex).tEventID = orsDE_AFSack.Fields("FTEvnCode").Value
    typeDE_MutiDeal(pnIndex).tSkuFlag = tSta
    typeDE_MutiDeal(pnIndex).tRegPrice = tSumRegPrice
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub DE_PRCxReCalDT()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'CREATE BY            # Thawatchai                              LAST MODIFIED BY           # Thawatchai
'CREATE DATE      #  Not Define                               LAST MODIFIED DATE     # 51-08-28
'DESCRIPTION       #
'ARGUMENT LIST #
'
'RETURN VALUE  # None
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    On Error GoTo Err_Handler
    tSql = "UPDATE " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "SET FCSdtDutyFreeAmt = CASE WHEN DT.FTShdTransType = '13' THEN"
    tSql = tSql & vbCrLf & "    ROUND(FCSdtSaleAmt,2)"
    tSql = tSql & vbCrLf & "ELSE"
    tSql = tSql & vbCrLf & "    0"
    tSql = tSql & vbCrLf & "END"
    tSql = tSql & vbCrLf & ",FCSdtDutyFreeQty = CASE WHEN DT.FTShdTransType = '13' THEN"
    tSql = tSql & vbCrLf & "    ROUND(FCSdtQtyAll,2)"
    tSql = tSql & vbCrLf & "ELSE"
    tSql = tSql & vbCrLf & "    0"
    tSql = tSql & vbCrLf & "END"
    tSql = tSql & vbCrLf & ",FCSdtAmtDisAvi = CASE FTScdStaChkDepart "
    tSql = tSql & vbCrLf & "    WHEN '1' THEN"                                                                      '*Ao 10-01-12
    tSql = tSql & vbCrLf & "        ROUND(FCSdtSaleAmt,2)"
    tSql = tSql & vbCrLf & "    WHEN '2' THEN"
    tSql = tSql & vbCrLf & "        0"
    tSql = tSql & vbCrLf & "    WHEN '3' THEN"
    tSql = tSql & vbCrLf & "        0"
    tSql = tSql & vbCrLf & "    ELSE ROUND(FCSdtSaleAmt,2)"
    tSql = tSql & vbCrLf & "END"
    tSql = tSql & vbCrLf & ",FTSkuStaMallEmpCard = CASE WHEN FTSkuStaMallEmpCard = 'Y' THEN"
    tSql = tSql & vbCrLf & "    CASE FTScdStaChkDepart "
    tSql = tSql & vbCrLf & "        WHEN '1' THEN"
    tSql = tSql & vbCrLf & "            CASE WHEN FTScdDiscType ='4' THEN"              '*Ao 10-01-12       '4 Free
    tSql = tSql & vbCrLf & "                'Y'"
    tSql = tSql & vbCrLf & "            ELSE"
    tSql = tSql & vbCrLf & "                'N'"
    tSql = tSql & vbCrLf & "            END   "
    tSql = tSql & vbCrLf & "        WHEN '2' THEN"
    tSql = tSql & vbCrLf & "            'N'"
    tSql = tSql & vbCrLf & "        WHEN '3' THEN"
    tSql = tSql & vbCrLf & "            'N'"
    tSql = tSql & vbCrLf & "        ELSE"
    tSql = tSql & vbCrLf & "            'Y'"
    tSql = tSql & vbCrLf & "END"
    tSql = tSql & vbCrLf & "ELSE"
    tSql = tSql & vbCrLf & "    'N'"
    tSql = tSql & vbCrLf & "END"
    tSql = tSql & vbCrLf & ",FCSdtQtyAllDisAvi = CASE WHEN FCSdtDis > 0 THEN 0 ELSE FCSdtQtyAll END"
'    If bVB_CNTrStaRegPreviewDay Then               '*Ao 2010-12-14    CommSheet ML-PRV-021-SF Wrong Description on Free Item.xls
'        tSql = tSql & vbCrLf & ",FTSkuAbbName = CASE WHEN ISNULL(FTSdtStaPreview,'N')='N' AND ISNULL(FTSdtStaNotAlwPrv,'N') <> 'Y' THEN  'P-'" & "+  FTSkuAbbName ELSE FTSkuAbbName END"
'        tSql = tSql & vbCrLf & ",FTSkuAbbNameSnd = CASE WHEN ISNULL(FTSdtStaPreview,'N')='N' AND ISNULL(FTSdtStaNotAlwPrv,'N') <> 'Y' THEN  'P-'" & "+  FTSkuAbbNameSnd ELSE FTSkuAbbNameSnd END"
'        tSql = tSql & vbCrLf & ",FTSkuAbbNameSrvDoc = CASE WHEN ISNULL(FTSdtStaPreview,'N')='N' AND ISNULL(FTSdtStaNotAlwPrv,'N') <> 'Y' THEN  'P-'" & "+  FTSkuAbbNameSrvDoc ELSE FTSkuAbbNameSrvDoc END"
'        tSql = tSql & vbCrLf & ",FTSkuAbbNameSndSrvDoc = CASE WHEN ISNULL(FTSdtStaPreview,'N')='N' AND ISNULL(FTSdtStaNotAlwPrv,'N') <> 'Y' THEN  'P-'" & "+  FTSkuAbbNameSndSrvDoc ELSE FTSkuAbbNameSndSrvDoc END"
'        tSql = tSql & vbCrLf & ",FTSdtStaPreview = 'Y'"
'    End If
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " DT LEFT JOIN " & tVB_CNTblCD & " CD ON DT.FNSdtSeqNo = CD.FNSdtSeqNo"
    tSql = tSql & vbCrLf & "WHERE  DT.FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND DT.FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND DT.FDShdTransDate = '" & Format(tUT_SaleDate, "YYYY/MM/DD") & "'"
    tSql = tSql & vbCrLf & "AND ISNULL(DT.FTSdtCancelFlag,'0') <> '1'" '*TW 51-07-17
    Call SP_SQLbExecute(tSql)
    'แยกส่วน UPDATE เพราะจะทำให้ไป UPDATE สินค้าที่ไม่ได้ Deal  และ Free ผิด
    If bVB_CNTrStaRegPreviewDay Then
        tSql = "UPDATE " & tVB_CNTblDT          '*Ao 2010-12-14    CommSheet ML-PRV-021-SF Wrong Description on Free Item.xls                                      LEFT('Canceled - ' + FTSkuAbbNameSrvDoc,50)"
        tSql = tSql & vbCrLf & "SET FTSkuAbbName = CASE WHEN ISNULL(FTSdtStaPreview,'N')='N' AND ISNULL(FTSdtStaNotAlwPrv,'N') <> 'Y' THEN  LEFT('P-' +  FTSkuAbbName,100) ELSE FTSkuAbbName END"
        tSql = tSql & vbCrLf & ",FTSkuAbbNameSnd = CASE WHEN ISNULL(FTSdtStaPreview,'N')='N' AND ISNULL(FTSdtStaNotAlwPrv,'N') <> 'Y' THEN  LEFT('P-' +  FTSkuAbbNameSnd,50) ELSE FTSkuAbbNameSnd END"
        tSql = tSql & vbCrLf & ",FTSkuAbbNameSrvDoc = CASE WHEN ISNULL(FTSdtStaPreview,'N')='N' AND ISNULL(FTSdtStaNotAlwPrv,'N') <> 'Y' THEN  LEFT('P-' +  FTSkuAbbNameSrvDoc,100) ELSE FTSkuAbbNameSrvDoc END"
        tSql = tSql & vbCrLf & ",FTSkuAbbNameSndSrvDoc = CASE WHEN ISNULL(FTSdtStaPreview,'N')='N' AND ISNULL(FTSdtStaNotAlwPrv,'N') <> 'Y' THEN  LEFT('P-' +  FTSkuAbbNameSndSrvDoc,50) ELSE FTSkuAbbNameSndSrvDoc END"
        tSql = tSql & vbCrLf & ",FTSdtStaPreview = 'Y'"
        tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT & " DT INNER JOIN " & tVB_CNTblCD & " CD ON DT.FNSdtSeqNo = CD.FNSdtSeqNo"
        tSql = tSql & vbCrLf & "WHERE  DT.FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & "AND DT.FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & "AND DT.FDShdTransDate = '" & Format(tUT_SaleDate, "YYYY/MM/DD") & "'"
        tSql = tSql & vbCrLf & "AND ISNULL(DT.FTSdtCancelFlag,'0') <> '1'"
        Call SP_SQLbExecute(tSql)
    End If
    
    Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Private Sub DE_SINxCal0040(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset)
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''   Call:
''   Cmt: '*Ao 09-09-17
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cDealDiscPer As Double, cDealDiscAmt As Double, cSumCanDisc As Double, cRegPrice As Double
'    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer, nPos As Integer
'    Dim tFreeFlag As String
'    Dim nDTSeqNo As Integer, nDTSeqNoNew As Integer
'    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer
'    '*KT 52-02-19 member price
'    Dim bMemberPrice As Boolean
'    Dim cB4DisChgOrg As Double
'    Dim cMemberPrice As Double
'    Dim tSKUCode As String
'    Dim cDiscTotalAll As Double
'    Dim tEventIDMemberPrice As String
'On Error GoTo Err_Handler
'    cDiscTotalAll = 0
'    porsTemp.Sort = "FNDacSkuGrp,FTSkuStaMallEmpCard,FCSdtRegPrice"
'    porsTemp.MoveFirst
'    Do While Not porsTemp.EOF
'        If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
'            nGrp = porsTemp.Fields("FNDacSkuGrp").Value
'            cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
'            tSKUCode = porsTemp.Fields("FTDacSku").Value
'
'            If nGrp <> nUncalGrp Then
'                nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
'                If nGrp <> nGrpOld Then
'                    cDealDiscPer = 0
'                    tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
'
'                    Select Case tFreeFlag
'                        Case "1" 'Buy Only
'                            nQtyGet = porsTemp.Fields("FNDacQtyBuy").Value
'                        Case "2" 'Get or Free Only
'                            nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
'                        Case "3" 'Both (Get or Free) And Buy
'                            nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
'                    End Select
'
'                    If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                        nQtyCal = nQtyGet
'                        nQtyGet = 0
'                    Else
'                        nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                        nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                    End If
'                    If Trim(porsTemp.Fields("FTPddDisType").Value) = "4" Then
'                        cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal)
'                    Else
'                        cDealDiscAmt = 0
'                    End If
'                      '*KT 52-02-19 member price
'                    'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
'                    '--------------------------------------------------------------------------------
'                    bMemberPrice = False
'                    If bDE_MemberPrice Then
'                        cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                        cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                        tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                        '*KT 52-09-08 Member Price + Consignment
'                        tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                        If Not (tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "") Then
'                            bMemberPrice = False
'                        Else
'                            bMemberPrice = True
'                        End If
'                    End If
'
'                    If Not bMemberPrice Then
'                        If Trim(porsTemp.Fields("FTPddDisType").Value) = "4" Then
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                        End If
'                        Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                        nGrpOld = nGrp
'                        If nQtyGet <= 0 Then
'                            nUncalGrp = nGrp
'                            If Trim(porsTemp.Fields("FTPddDisType").Value = "4") And (nQtyGet = 0) Then         '*Ao 09-09-22   ได้ Free กลุ่มใดกลุ่มหนึ่ง
'                                Exit Do
'                            End If
'                        End If
'                    Else
'                        If tVB_StaMemPriceAlw = "Y" Then
'                            If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                                bDE_MemberPriceAct = True
'                            Else
'                                If Trim(porsTemp.Fields("FTPddDisType").Value) = "4" Then
'                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                End If
'                                Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                nGrpOld = nGrp
'                                If nQtyGet <= 0 Then
'                                    nUncalGrp = nGrp
'                                    If Trim(porsTemp.Fields("FTPddDisType").Value = "4") And (nQtyGet = 0) Then         '*Ao 09-09-22   ได้ Free กลุ่มใดกลุ่มหนึ่ง
'                                        Exit Do
'                                    End If
'                                End If
'                            End If
'                        End If
'                    End If
'                    '------------------------------------------------------------------------------------
'                Else        'If nGrp <> nGrpOld Then   เท่ากัน
'                    If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                        nQtyCal = nQtyGet
'                        nQtyGet = 0
'                    Else
'                        nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                        nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                    End If
'
'                    If Trim(porsTemp.Fields("FTPddDisType").Value) = "4" Then
'                        cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal)
'                    Else
'                        cDealDiscAmt = 0
'                    End If
'                    '*KT 52-02-19 member price
'                    'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
'                    '--------------------------------------------------------------------------------
'                    bMemberPrice = False
'                    If bDE_MemberPrice Then
'                        cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                        cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                        tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                         '*KT 52-09-08 Member Price + Consignment
'                        tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                        If Not (tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "") Then
'                            bMemberPrice = False
'                        Else
'                            bMemberPrice = True
'                        End If
'                    End If
'
'                    If Not bMemberPrice Then
'                        If Trim(porsTemp.Fields("FTPddDisType").Value) = "4" Then
'                            nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                            Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                        End If
'                        Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                        If nQtyGet <= 0 Then
'                            nUncalGrp = nGrp
'                            If Trim(porsTemp.Fields("FTPddDisType").Value = "4") And (nQtyGet <= 0) Then         '*Ao 09-09-22   ได้ Free กลุ่มใดกลุ่มหนึ่ง
'                                Exit Do
'                            End If
'                        End If
'                    Else
'                        If tVB_StaMemPriceAlw = "Y" Then
'                            If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                                bDE_MemberPriceAct = True
'                            Else
'                                If Trim(porsTemp.Fields("FTPddDisType").Value) = "4" Then
'                                    nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                    Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                End If
'                                Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                If nQtyGet <= 0 Then
'                                    nUncalGrp = nGrp
'                                    If Trim(porsTemp.Fields("FTPddDisType").Value = "4") And (nQtyGet = 0) Then         '*Ao 09-09-22   ได้ Free กลุ่มใดกลุ่มหนึ่ง
'                                        Exit Do
'                                    End If
'                                End If
'                            End If
'                        End If
'                    End If
'                    '---------------------------------------------------------------------------------
'                End If
'            End If
'        End If
'        porsTemp.MoveNext
'    Loop
'
'    Exit Sub
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub

Private Function DE_SINbCal0040(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 09-10-27
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDealDiscPer As Double, cDealDiscAmt As Double, cSumCanDisc As Double, cRegPrice As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer, nDTSeqNo As Integer, nDTSeqNoNew As Integer
    Dim tFreeFlag As String, tFilter As String, tSort As String
    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer
    
    '*KT 52-02-19 member price
    Dim bMemberPrice As Boolean
    Dim cB4DisChgOrg As Double
    Dim cMemberPrice As Double
    Dim tSKUCode As String
    Dim cDiscTotalAll As Double
    Dim tEventIDMemberPrice As String
On Error GoTo Err_Handler

    DE_SINbCal0040 = False
    With porsTemp
        tSort = "FNDacSkuGrp,FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
        .Sort = tSort
            If Not (.BOF And .EOF) Then
                .MoveFirst
                Do While Not .EOF
                    If .Fields("FCSdtQtyAll").Value > 0 Then
                        nGrp = .Fields("FNDacSkuGrp").Value
                        cRegPrice = .Fields("FCSdtRegPrice").Value
                        tSKUCode = .Fields("FTDacSku").Value
            
                        If nGrp <> nUncalGrp Then
                            nDTSeqNo = .Fields("FNSdtSeqNo").Value
                            If nGrp <> nGrpOld Then
                                cDealDiscPer = 0
                                tFreeFlag = .Fields("FTFreeFlag").Value
                                
                                If CInt(.Fields("FNDacQtyBuy").Value) = CInt(0) Then
                                    nQtyGet = .Fields("FNDacQtyGet").Value * pnCountRound                'Get or Free Only
                                    tFreeFlag = "2"
                                Else
                                    nQtyGet = .Fields("FNDacQtyBuy").Value * pnCountRound                'Buy Only
                                    tFreeFlag = "1"
                                End If
'                                If Trim(.Fields("FTPddDisType").Value) = "4" Then
'                                    nQtyGet = .Fields("FNDacQtyGet").Value * pnCountRound                'Get or Free Only
'                                    tFreeFlag = "2"
'                                Else
'                                    nQtyGet = .Fields("FNDacQtyBuy").Value * pnCountRound                'Buy Only
'                                    tFreeFlag = "1"
'                                End If
'                                Select Case tFreeFlag
'                                    Case "1" 'Buy Only
'                                        nQtyGet = .Fields("FNDacQtyBuy").Value
'                                        tFreeFlag = "4"
'                                    Case "2" 'Get or Free Only
'                                        nQtyGet = .Fields("FNDacQtyGet").Value
'                                        tFreeFlag = "3"
'                                    Case "3" 'Both (Get or Free) And Buy
'                                        nQtyGet = .Fields("FNDacQtyGet").Value
'                                End Select
                                
                                If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                                    nQtyCal = nQtyGet
                                    nQtyGet = 0
                                Else
                                    nQtyCal = .Fields("FCSdtQtyAll").Value
                                    nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                                End If
                                
                                If CInt(.Fields("FNDacQtyBuy").Value) = CInt(0) Then
                                    cDealDiscAmt = (.Fields("FCSdtRegPrice").Value * nQtyCal)
                                Else
                                    cDealDiscAmt = 0
                                End If
'                                If Trim(.Fields("FTPddDisType").Value) = "4" Then
'                                    cDealDiscAmt = (.Fields("FCSdtRegPrice").Value * nQtyCal)
'                                Else
'                                    cDealDiscAmt = 0
'                                End If
                                '*KT 52-02-19 member price
                                'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
                                '--------------------------------------------------------------------------------
                                bMemberPrice = False
                                If bDE_MemberPrice Then
                                    cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                                    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                                    tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                                    '*KT 52-09-08 Member Price + Consignment
                                    tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                                    If Not (tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "") Then
                                        bMemberPrice = False
                                    Else
                                        bMemberPrice = True
                                    End If
                                End If
                                If Not bMemberPrice Then
'                                    If Trim(.Fields("FTPddDisType").Value) = "4" Then
                                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                    End If
                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                    nGrpOld = nGrp
                                    If nQtyGet < 1 Then
                                        nUncalGrp = nGrp
                                    End If
                                Else
                                    If tVB_StaMemPriceAlw = "Y" Then
                                        If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                                            bDE_MemberPriceAct = True
                                        Else
'                                            If Trim(.Fields("FTPddDisType").Value) = "4" Then
                                                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                            End If
                                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                            nGrpOld = nGrp
                                            If nQtyGet < 1 Then
                                                nUncalGrp = nGrp
                                            End If
                                        End If
                                    End If
                                End If
                                '--------------------------------------------------------------------------------
                            Else        'If nGrp <> nGrpOld Then                    'เท่ากัน
                                If .Fields("FCSdtQtyAll").Value >= nQtyGet Then
                                    nQtyCal = nQtyGet
                                    nQtyGet = 0
                                Else
                                    nQtyCal = .Fields("FCSdtQtyAll").Value
                                    nQtyGet = nQtyGet - .Fields("FCSdtQtyAll").Value
                                End If
                                
                                If CInt(.Fields("FNDacQtyBuy").Value) = 0 Then                                              'Get
                                    cDealDiscAmt = (.Fields("FCSdtRegPrice").Value * nQtyCal)
                                Else
                                    cDealDiscAmt = 0                        'Buy
                                End If
'                                If Trim(.Fields("FTPddDisType").Value) = "4" Then                              'Ao 10-02-04
'                                    cDealDiscAmt = (.Fields("FCSdtRegPrice").Value * nQtyCal)
'                                Else
'                                    cDealDiscAmt = 0
'                                End If
                                '*KT 52-02-19 member price
                                'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
                                '--------------------------------------------------------------------------------
                                bMemberPrice = False
                                If bDE_MemberPrice Then
                                    cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                                    cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                                    tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                                     '*KT 52-09-08 Member Price + Consignment
                                    tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                                    If Not (tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "") Then
                                        bMemberPrice = False
                                    Else
                                        bMemberPrice = True
                                    End If
                                End If      'If bDE_MemberPrice Then
                                
                                If Not bMemberPrice Then
'                                    If Trim(.Fields("FTPddDisType").Value) = "4" Then
                                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                    End If
                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                    If nQtyGet < 1 Then
                                        nUncalGrp = nGrp
                                    End If
                                Else
                                    If tVB_StaMemPriceAlw = "Y" Then
                                        If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                                            bDE_MemberPriceAct = True
                                        Else
'                                            If Trim(.Fields("FTPddDisType").Value) = "4" Then
                                                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                            End If
                                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                            If nQtyGet < 1 Then
                                                nUncalGrp = nGrp
                                            End If
                                        End If
                                    End If
                                End If      'If Not bMemberPrice Then
            
                            End If      'If nGrp <> nGrpOld Then
                        End If          'If nGrp <> nUncalGrp Then
                    End If      'If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
                    .MoveNext
                Loop
            Else
'                Call SP_MSGxShowError(Err.Number, Err.Description)
                Exit Function
            End If
'            nPos = nPos + 1
'        Loop
        
    End With
    DE_SINbCal0040 = True
    Exit Function
Err_Handler:
    DE_SINbCal0040 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

'Private Sub DE_SINxCal0030(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset)
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
''   Call:
''   Cmt: '*Ao 09-09-22
''-----------------------------------------------------------------------------------------------------------------------------------------------------------
'    Dim cDealDiscPer As Double, cDealDiscAmt As Double
'    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer
'    Dim tFreeFlag As String
'    Dim nDTSeqNo As Integer
'    Dim nDTSeqNoNew As Integer
'    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer
'    Dim cSumCanDisc As Double
'    Dim cRegPrice As Double
'    Dim nPos As Integer
'    '*KT 52-02-19 member price
'    Dim bMemberPrice As Boolean
'    Dim cB4DisChgOrg As Double
'    Dim cMemberPrice As Double
'    Dim tSKUCode As String
'    Dim cDiscTotalAll As Double
'    Dim tEventIDMemberPrice As String
'
'On Error GoTo Err_Handler
'    cDiscTotalAll = 0
'    nMaxGrp = DE_GETnMaxGroup(ptDealNo, porsTemp)
'    porsTemp.MoveFirst
'    If nMaxGrp = 2 Then
'        Do While Not porsTemp.EOF
'            If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
'                nGrp = porsTemp.Fields("FNDacSkuGrp").Value
'                cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
'                tSKUCode = porsTemp.Fields("FTDacSku").Value
'                If nGrp <> nUncalGrp Then
'                    nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
'                    If nGrp <> nGrpOld Then
''                        Select Case porsTemp.Fields("FTPddDisType").Value
''                            Case "1" 'Per
''                                cDealDiscPer = porsTemp.Fields("FCDacDisAmt").Value
''                            Case "2", "3"
''                                cDealDiscPer = 0
''                            Case Else
'                        cDealDiscPer = 0
''                        End Select
'                        tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
'                        Select Case tFreeFlag
'                            Case "1" 'Buy Only
'                                nQtyGet = porsTemp.Fields("FNDacQtyBuy").Value
'                            Case "2" 'Get or Free Only
'                                nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
'                            Case "3" 'Both (Get or Free) And Buy
'                                nQtyGet = porsTemp.Fields("FNDacQtyGet").Value
'                        End Select
'
''                        nQtyGetOrg = nQtyGet
'                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                            nQtyCal = nQtyGet
'                            nQtyGet = 0
'                        Else
'                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                            nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                        End If
'                        If nGrp = 2 Then
''                            nPos = porsTemp.AbsolutePosition
''                            If cSumCanDisc <= 0 Then cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGetOrg, nGrp)                            '*TW 51-06-27
'                            'cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGetOrg, nGrp) '*TW 51-06-27
''                            porsTemp.AbsolutePosition = nPos
''                            Select Case porsTemp.Fields("FTPddDisType").Value
''                                Case "1" 'Per
''                                    cDealDiscAmt = cRegPrice * (cDealDiscPer / 100) * nQtyCal
''                                Case "2"
''                                    cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc)
''                                Case "3"
''                                    cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc))
''                                Case "4"
'                                        cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal)
''                            End Select
'                        Else
'                            cDealDiscAmt = 0
'                        End If
'                        '*Ao 52-04-20
''                        ptLnDiscType = porsTemp.Fields("FTPddDisType").Value
''                        cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
''                        cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
'                        '*KT 52-02-19 member price
'                        'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
'                        '--------------------------------------------------------------------------------
'                        bMemberPrice = False
'                        If bDE_MemberPrice Then
'                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                           '*KT 52-09-08 Member Price + Consignment
'                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                            If Not (tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "") Then
'                                bMemberPrice = False
'                            Else
'                                bMemberPrice = True
'                            End If
'                        End If
'                        '*Ao 52-04-20
''                        Call DE_SINcDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
'
'                        If Not bMemberPrice Then
'                            If nGrp = 2 Then
'                                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                            End If
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                            nGrpOld = nGrp
'                            If nQtyGet <= 0 Then
'                                nUncalGrp = nGrp
'                            End If
'                        Else
'                            If tVB_StaMemPriceAlw = "Y" Then
'                                 If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                                    bDE_MemberPriceAct = True
'                                Else
'                                    If nGrp = 2 Then
'                                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                    End If
'                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                    nGrpOld = nGrp
'                                    If nQtyGet <= 0 Then
'                                        nUncalGrp = nGrp
'                                    End If
'                                End If
'                            End If
'                        End If
'                         '--------------------------------------------------------------------------------
'                    Else
'                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
'                            nQtyCal = nQtyGet
'                            nQtyGet = 0
'                        Else
'                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
'                            nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
'                        End If
'                        If nGrp = 2 Then
''                            nPos = porsTemp.AbsolutePosition
''                            If cSumCanDisc <= 0 Then cSumCanDisc = DE_GETcSumDiscByQty(porsTemp, nQtyGet, nGrp)                            '*TW 51-06-27
''                            porsTemp.AbsolutePosition = nPos
''                            Select Case porsTemp.Fields("FTPddDisType").Value
''                                Case "1" 'Per
''                                    cDealDiscAmt = IIf(nGrp = 1, 0, porsTemp.Fields("FCSdtRegPrice").Value * (cDealDiscPer / 100) * nQtyCal)
''                                Case "2"
''                                    cDealDiscAmt = porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc)
''                                Case "3"
''                                    cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal) - (porsTemp.Fields("FCDacDisAmt").Value * ((cRegPrice * nQtyCal) / cSumCanDisc))
''                                Case "4"
''                                    If nQtyGet = 0 Then
'                                    cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal)
''                                    Else
''                                        cDealDiscAmt = 0
''                                    End If
''                            End Select
'                        Else
'                            cDealDiscAmt = 0
'                        End If
'                        '*Ao 52-04-20
''                        ptLnDiscType = porsTemp.Fields("FTPddDisType").Value
''                        cDiscTotalAll = cDiscTotalAll + SP_FMTcCurPoint(cDealDiscAmt, True, False)
''                        cDealDiscAmt = IIf(cDealDiscAmt < 0, 0, cDealDiscAmt) '*TW 51-06-27
'                        '*KT 52-02-19 member price
'                        'ตรวจสอบว่าราคาที่ได้จาก deal ได้ส่วนลดมากกว่า member price จะให้ส่วนลด deal เป็นหลัก
'                        '--------------------------------------------------------------------------------
'                        bMemberPrice = False
'                        If bDE_MemberPrice Then
'                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
'                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
'                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
'                            '*KT 52-09-08 Member Price + Consignment
'                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
'                            If Not (tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "") Then
'                                bMemberPrice = False
'                            Else
'                                bMemberPrice = True
'                            End If
'                        End If
'                        '*Ao 52-04-20
''                        Call DE_SINcDealAmt(ptLnDiscType, nQtyGet, porsTemp.Fields("FCDacDisAmt").Value, cDiscTotalAll, cDealDiscAmt, cSumCanDisc)
'
'                         If Not bMemberPrice Then
'                            If nGrp = 2 Then
'                                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                            End If
'                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                            If nQtyGet <= 0 Then
'                                nUncalGrp = nGrp
'                            End If
'                        Else
'                            If tVB_StaMemPriceAlw = "Y" Then
'                                 If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
'                                    bDE_MemberPriceAct = True
'                                Else
'                                    If nGrp = 2 Then
'                                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
'                                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, porsTemp.Fields("FTPddDisType").Value, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
'                                    End If
'                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
'                                    If nQtyGet <= 0 Then
'                                        nUncalGrp = nGrp
'                                    End If
'                                End If
'                            End If
'                        End If
'                        '--------------------------------------------------------------------------------
'                    End If
'                End If
'            End If
'            porsTemp.MoveNext
'        Loop
'    End If
'    Exit Sub
'Err_Handler:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub

Private Function DE_SINbCal0030(ByVal ptDealNo$, ByVal ptDealType$, ByVal ptFreeType$, ByVal ptEventID$, ByVal ptLnDiscType$, ByRef porsTemp As ADODB.Recordset, ByVal pnCountRound As Long) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 09-09-22
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cDealDiscPer As Double, cDealDiscAmt As Double, cSumCanDisc As Double, cRegPrice As Double
    Dim nQtyGet As Integer, nQtyGetOrg As Integer, nQtyCal As Integer, nQtyBuy As Integer, nDTSeqNoNew As Integer, nPos As Integer, nDTSeqNo As Integer
    Dim nGrp As Integer, nUncalGrp As Integer, nGrpOld As Integer, nMaxGrp As Integer
    Dim tFreeFlag As String, tFilter As String, tSort As String
    Dim bChkround As Boolean
    
    '*KT 52-02-19 member price
    Dim bMemberPrice As Boolean
    Dim cB4DisChgOrg As Double
    Dim cMemberPrice As Double
    Dim tSKUCode As String
    Dim tEventIDMemberPrice As String
    
On Error GoTo Err_Handler

    DE_SINbCal0030 = False
    If pnCountRound = 0 Then Exit Function
    
    nMaxGrp = DE_GETnMaxGroup(ptDealNo, porsTemp)
    If Not (nMaxGrp = 2) Then Exit Function
    
    For nGrp = 1 To nMaxGrp
        bChkround = True
        porsTemp.Filter = adFilterNone
        tFilter = "FNDacSkuGrp = " & nGrp
        porsTemp.Filter = tFilter
        If nGrp = 1 Then
            tSort = "FTSkuStaMallEmpCard,FCSdtRegPrice"
        Else
            tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
        End If
        porsTemp.Sort = tSort

        If Not (porsTemp.BOF And porsTemp.EOF) Then
            porsTemp.MoveFirst
            Do While Not porsTemp.EOF
                If nGrp = 1 Then
                    If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
                        cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
                        tSKUCode = porsTemp.Fields("FTDacSku").Value
                        nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
                        cDealDiscPer = 0
                        cDealDiscAmt = 0
'                        tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
                        tFreeFlag = "1"
                        If bChkround Then
                            nQtyBuy = porsTemp.Fields("FNDacQtyBuy").Value * pnCountRound
                            bChkround = False
                        End If
                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyBuy Then
                            nQtyCal = nQtyBuy
                            nQtyBuy = 0
                        Else
                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                            nQtyBuy = nQtyBuy - porsTemp.Fields("FCSdtQtyAll").Value
                        End If
                        nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                        Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                        
                        If nQtyBuy < 1 Then
                            Exit Do
                        End If
                    End If          'If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
                Else            'If nGrp = 1 Then
                    If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
                        cRegPrice = porsTemp.Fields("FCSdtRegPrice").Value
                        tSKUCode = porsTemp.Fields("FTDacSku").Value
                        nDTSeqNo = porsTemp.Fields("FNSdtSeqNo").Value
                        cDealDiscPer = 0
'                        tFreeFlag = porsTemp.Fields("FTFreeFlag").Value
                        tFreeFlag = "2"
                        If bChkround Then
                            nQtyGet = porsTemp.Fields("FNDacQtyGet").Value * pnCountRound
                            bChkround = False
                        End If
                        If porsTemp.Fields("FCSdtQtyAll").Value >= nQtyGet Then
                            nQtyCal = nQtyGet
                            nQtyGet = 0
                        Else
                            nQtyCal = porsTemp.Fields("FCSdtQtyAll").Value
                            nQtyGet = nQtyGet - porsTemp.Fields("FCSdtQtyAll").Value
                        End If
                        cDealDiscAmt = (porsTemp.Fields("FCSdtRegPrice").Value * nQtyCal)
                        
                        bMemberPrice = False
                        If bDE_MemberPrice Then
                            cB4DisChgOrg = SP_FEDvChkNumeric(porsTemp, "FCSdtB4DisChgOrg") ' porsTemp.Fields("FCSdtB4DisChgOrg").Value
                            cMemberPrice = SP_FEDvChkNumeric(porsTemp, "FCSdtMemberPrice") ' porsTemp.Fields("FCSdtMemberPrice").Value
                            tVB_StaMemPriceAlw = SP_FEDtChkString(porsTemp, "FTSdtStaRfd") ' porsTemp.Fields("FTSdtStaRfd").Value
                           '*KT 52-09-08 Member Price + Consignment
                            tEventIDMemberPrice = SP_FEDtChkString(porsTemp, "FTSdtEventMemPrice")
                            If Not (tVB_StaMemPriceAlw = "Y" And tEventIDMemberPrice <> "") Then
                                bMemberPrice = False
                            Else
                                bMemberPrice = True
                            End If
                        End If
                        
                        If Not bMemberPrice Then
                            If nGrp = 2 Then
                                nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                            End If
                            Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                            If nQtyGet < 1 Then
                                Exit Do
                            End If
                        Else
                            If tVB_StaMemPriceAlw = "Y" Then
                                 If oDisc.C_DATxSumDTMemberPrice(nDTSeqNo, tSKUCode) Then
                                    bDE_MemberPriceAct = True
                                Else
                                    If nGrp = 2 Then
                                        nDTSeqNoNew = SP_VALcText2Double(SP_SQLtFunction("", "FNSdtSeqNo", tVB_CNTblDT, tDE_Whe & " AND ISNULL(FNShdSign,0) = 0" & " AND FNSdtStaRef = " & nDTSeqNo)) 'หา Seq หลัง Split
                                        Call DE_GETcCallDeal(ptDealNo, tFreeFlag, ptDealType, ptFreeType, ptEventID, ptLnDiscType, cDealDiscPer, nDTSeqNoNew, nQtyCal, cDealDiscAmt, cRegPrice * nQtyCal)
                                    End If
                                    Call DE_UPDxQtySingle(porsTemp, nQtyCal, nDTSeqNo)
                                    If nQtyGet < 1 Then
                                        Exit Do
                                    End If
                                End If
                            End If
                        End If          'If Not bMemberPrice Then
                    End If      'If porsTemp.Fields("FCSdtQtyAll").Value > 0 Then
                End If      'If nGrp = 1 Then
                porsTemp.MoveNext
            Loop
        Else
            Call SP_MSGxShowError(Err.Number, Err.Description)
            Exit Function
        End If
    
    Next nGrp
    DE_SINbCal0030 = True
    Exit Function
Err_Handler:
    DE_SINbCal0030 = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function DE_DISb0040(ByVal pnTime As Integer, ByRef pnKey As Long, ByRef porsTemp As ADODB.Recordset, ByRef pnPos As Integer, ByRef pcOldDisc As Double) As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 09-09-23
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim cSource As Double, cDisSumDisc As Double, cDisSumPrice As Double, cRound As Double
    Dim nMaxGrp As Integer, nI As Integer
    Dim tFilter As String, tSort As String
    
    On Error GoTo Err_Handler
    With porsTemp
        tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "'"
        .Filter = tFilter
        cRound = DE_PRCnChkRound(porsTemp, atDE_AllDeal(pnKey), "0040")
        .Filter = adFilterNone
        
        nMaxGrp = DE_GETnMaxGroup(atDE_AllDeal(pnKey), porsTemp)
        cDisSumDisc = 0
        cDisSumPrice = 0
        For nI = 1 To nMaxGrp
            tFilter = "FTPdhDealNo = '" & atDE_AllDeal(pnKey) & "' AND FNDacSkuGrp = " & nI
            .Filter = tFilter
            tSort = "FCSdtRegPrice DESC,FTSkuStaMallEmpCard"
            .Sort = tSort
            If Not (.BOF And .EOF) Then
                .MoveFirst
                If CInt(.Fields("FNDacQtyBuy").Value) = CInt(0) Then
                    cSource = .Fields("FNDacQtyGet").Value * cRound                 'Get
                Else
                    cSource = .Fields("FNDacQtyBuy").Value * cRound                 'Buy
                End If
'                If Trim(.Fields("FTPddDisType").Value = "4") Then                      'Ao 10-02-04
'                    cSource = .Fields("FNDacQtyGet").Value * cRound
'                Else
'                    cSource = .Fields("FNDacQtyBuy").Value * cRound
'                End If
                Do While Not .EOF
                    If .Fields("FCSdtQtyAll").Value > 0 Then
                        If .Fields("FCSdtQtyAll").Value >= cSource Then
                            If CInt(.Fields("FNDacQtyBuy").Value) = CInt(0) Then                                                                                        'Ao 10-02-04
                                cDisSumDisc = cDisSumDisc + (.Fields("FCSdtRegPrice").Value * cSource)
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
                            Else
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "0")
                            End If
'                            If Trim(.Fields("FTPddDisType").Value = "4") Then
'                                cDisSumDisc = cDisSumDisc + (.Fields("FCSdtRegPrice").Value * cSource)
'                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "1")
'                            Else
'                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, cSource, "0")
'                            End If
                            tFilter = .Filter
                            tSort = .Sort
                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", .Fields("FCSdtQtyAll").Value - cSource, tFilter, tSort)
                            cSource = 0

                            Exit Do
                        Else
                            If CInt(.Fields("FNDacQtyBuy").Value) = CInt(0) Then
                                cDisSumDisc = cDisSumDisc + (.Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value)
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1")
                            Else
                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "0")
                            End If
'                            If Trim(porsTemp.Fields("FTPddDisType").Value = "4") Then                                                                                                              '*Ao 10-02-04
'                                cDisSumDisc = cDisSumDisc + (.Fields("FCSdtRegPrice").Value * .Fields("FCSdtQtyAll").Value)
'                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "1") 'cSource)
'                            Else
'                                Call DE_SETxSkuUsed(pnPos, pnTime, .Fields("FNSdtSeqNo").Value, .Fields("FCSdtQtyAll").Value, "0")
'                            End If
                            cSource = cSource - .Fields("FCSdtQtyAll").Value
                            tFilter = .Filter
                            tSort = .Sort
                            Call DE_RSTxUpdateBySku(porsTemp, .Fields("FNSdtSeqNo").Value, "FCSdtQtyAll", 0, tFilter, tSort)
                        End If
                    Else
                        cDisSumDisc = 0                         '*Ao 10-02-04          ตั้งสินค้ากลุ่ม Buy ซ้ำกัน แล้วหยิบสินค้ามาไม่ครบ
                        DE_DISb0040 = False
                        Exit For
                    End If
                    .MoveNext
                Loop
            Else
                DE_DISb0040 = False
            End If
        Next nI
        If cSource = 0 Then
            DE_DISb0040 = True
'                    acDE_SumDisc(pnPos) = IIf(cDisSumPrice = 0, cDisSumDisc, cDisSumPrice - cDisSumDisc) + pcOldDisc
            acDE_SumDisc(pnPos) = cDisSumDisc + pcOldDisc
        Else
            DE_DISb0040 = False
        End If
    End With
    Exit Function
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function







