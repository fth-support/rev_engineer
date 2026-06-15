VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSKU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_cSaleWeight As Double
Private m_cWeight As Double
Private m_tSaleWeight As String
Private m_tWeight As String
Private m_nDecWeight As Long
Private m_nDecPrice As Long
Private m_bInStoreBar As Boolean
Private m_tStaWPrint  As String 'Status Print Weight on Receopt/EJ           Y: Allow Print   N: Not Allow Print
Private m_tStaChkDigit  As String
Private m_tID As String
Private m_tIDRuning As String
Private m_bResults As Boolean
'Private m_nBafItemSize As Long
'Private m_nBafPLUSize As Long
'Private m_nBafPLUSize As Long

Public Property Get GETSKUBarcode(ByVal ptTextPrompt As String) As String
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT FTBafCode,FTBafCodeType,FNBafMatchStr,FNBafItemSize,FNBafPLUSize,FNBafPLULocate,FNBafPriceSize,FNBafPriceLocate"
    tSql = tSql & ",FNBafPriceDec,FNBafWeightSize,FNBafWeightLocate,FNBafWeightDec,FTBafStaChkDigit,FTBafStaWPrint"
    tSql = tSql & " FROM TCNMSkuBarCon"
    tSql = tSql & " WHERE FNBafItemSize = " & Len(ptTextPrompt)
    tSql = tSql & " AND FTBafCodeType = '1'" '*KT 53-10-19  "1" :  Single Line(SKU) "2" : Vendor Coupon
    tSql = tSql & " ORDER BY FTBafCode"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            orsTemp.Filter = "FNBafMatchStr = '" & Mid(ptTextPrompt, 1, 1) & "'"
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                 orsTemp.MoveFirst
                 GETSKUBarcode = Mid(ptTextPrompt, Val(Format(orsTemp("FNBafPLULocate"))), Val(Format(orsTemp("FNBafPLUSize"))))
                m_cSaleWeight = Val(Format(Mid(ptTextPrompt, Val(Format(orsTemp("FNBafPriceLocate"))), Val(Format(orsTemp("FNBafPriceSize")))))) / (10 ^ Val(Format(orsTemp("FNBafPriceDec"))))
                m_cSaleWeight = Format(m_cSaleWeight, String((Val(Format(orsTemp("FNBafPLUSize"))) - Val(Format(orsTemp("FNBafPriceDec"))) - 1), "#") & "0." & String(Val(Format(orsTemp("FNBafPriceDec"))), "0"))
                
                m_cWeight = Val(Format(Mid(ptTextPrompt, Val(Format(orsTemp("FNBafWeightLocate"))), Val(Format(orsTemp("FNBafWeightSize")))))) / (10 ^ Val(Format(orsTemp("FNBafWeightDec"))))
                m_cWeight = Format(m_cWeight, String((Val(Format(orsTemp("FNBafWeightSize"))) - Val(Format(orsTemp("FNBafWeightDec"))) - 1), "#") & "0." & String(Val(Format(orsTemp("FNBafWeightDec"))), "0"))
    
                m_tSaleWeight = Mid(ptTextPrompt, Val(Format(orsTemp("FNBafPriceLocate"))), Val(Format(orsTemp("FNBafPriceSize"))))
                m_tWeight = Mid(ptTextPrompt, Val(Format(orsTemp("FNBafWeightLocate"))), Val(Format(orsTemp("FNBafWeightSize"))))
                m_nDecPrice = Val(Format(orsTemp("FNBafPriceDec")))
                m_nDecWeight = Val(Format(orsTemp("FNBafWeightDec")))
                m_bInStoreBar = True
                m_tStaWPrint = SP_FEDtChkString(orsTemp, "FTBafStaWPrint")
            Else
                orsTemp.Filter = adFilterNone
                orsTemp.MoveFirst
                GETSKUBarcode = Mid(ptTextPrompt, Val(Format(orsTemp("FNBafPLULocate"))), Val(Format(orsTemp("FNBafPLUSize"))))
                m_cSaleWeight = 0
                m_cWeight = 0
                m_tSaleWeight = ""
                m_tWeight = ""
                m_nDecPrice = 0
                m_nDecWeight = 0
                m_bInStoreBar = False
                m_tStaWPrint = "N"
            End If
        Else
            GETSKUBarcode = ptTextPrompt
            m_cSaleWeight = 0
            m_cWeight = 0
            m_tSaleWeight = ""
            m_tWeight = ""
            m_nDecPrice = 0
            m_nDecWeight = 0
            m_bInStoreBar = False
            m_tStaWPrint = "N"
        End If
    Else
        GETSKUBarcode = ptTextPrompt
        m_cSaleWeight = 0
        m_cWeight = 0
        m_tSaleWeight = ""
        m_tWeight = ""
        m_nDecPrice = 0
        m_nDecWeight = 0
        m_bInStoreBar = False
        m_tStaWPrint = "N"
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Property
ErrHandle:
    GETSKUBarcode = ptTextPrompt
    m_cSaleWeight = 0
    m_cWeight = 0
    m_tSaleWeight = ""
    m_tWeight = ""
    m_nDecPrice = 0
    m_nDecWeight = 0
    m_bInStoreBar = False
    m_tStaWPrint = "N"
    Call SP_DATxRsNothing(orsTemp)
End Property

Public Property Get GETVendorCpnBarcode(ByVal ptTextPrompt As String) As String
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT FTBafCode,FTBafCodeType,FNBafMatchStr,FNBafItemSize,FNBafPLUSize,FNBafPLULocate,FNBafPriceSize,FNBafPriceLocate"
    tSql = tSql & ",FNBafPriceDec,FNBafWeightSize,FNBafWeightLocate,FNBafWeightDec,FTBafStaChkDigit"
    tSql = tSql & " FROM TCNMSkuBarCon"
    tSql = tSql & " WHERE FNBafItemSize = " & Len(ptTextPrompt)
    tSql = tSql & " AND FTBafCodeType = '2'" '*KT 53-10-19  "1" :  Single Line(SKU) "2" : Vendor Coupon
    tSql = tSql & " ORDER BY FTBafCode"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then  'read master
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            orsTemp.Filter = "FNBafMatchStr = '" & Mid(ptTextPrompt, 1, 1) & "'"
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                 orsTemp.MoveFirst
                 GETVendorCpnBarcode = Mid(ptTextPrompt, Val(Format(orsTemp("FNBafPLULocate"))), Val(Format(orsTemp("FNBafPLUSize"))))
                 m_bInStoreBar = True
            Else
                GETVendorCpnBarcode = ptTextPrompt
                m_bInStoreBar = False
            End If
        Else
            GETVendorCpnBarcode = ptTextPrompt
            m_bInStoreBar = False
        End If
    Else
        GETVendorCpnBarcode = ptTextPrompt
        m_bInStoreBar = False
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Property
ErrHandle:
    GETVendorCpnBarcode = ptTextPrompt
    m_bInStoreBar = False
    Call SP_DATxRsNothing(orsTemp)
End Property


Public Property Get GETRateDiscountCpnBarcode(ByVal ptTextPrompt As String) As String
'--------------------------------------------------------------------------------------------------------------------------------
'Ret : Rate Cpn
'Cmt : *KT 54-06-21 V.4.0.0 Ph.4.0 Req.14  คูปองส่วนลด เพิ่มการ Config Barcode (FTBafCodeType = '3')
'--------------------------------------------------------------------------------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT FTBafCode,FTBafCodeType,FNBafMatchStr,FNBafItemSize,FNBafPLUSize,FNBafPLULocate,FNBafPriceSize,FNBafPriceLocate"
    tSql = tSql & ",FNBafPriceDec"
    tSql = tSql & " FROM TCNMSkuBarCon"
    tSql = tSql & " WHERE FNBafItemSize = " & Len(ptTextPrompt)
    tSql = tSql & " AND FTBafCodeType = '3'"
    tSql = tSql & " ORDER BY FTBafCode"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then  'read master
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            orsTemp.Filter = "FNBafMatchStr = '" & Mid(ptTextPrompt, 1, 1) & "'"
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                 orsTemp.MoveFirst
                 GETRateDiscountCpnBarcode = Mid(ptTextPrompt, Val(Format(orsTemp("FNBafPriceLocate"))), Val(Format(orsTemp("FNBafPriceSize"))))
            Else
                GETRateDiscountCpnBarcode = "0"
            End If
        Else
            GETRateDiscountCpnBarcode = "0"
        End If
    Else
        GETRateDiscountCpnBarcode = "0"
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Property
ErrHandle:
    GETRateDiscountCpnBarcode = "0"
    Call SP_DATxRsNothing(orsTemp)
End Property

Public Property Get GETPmtCpnBarcode(ByVal ptTextPrompt As String) As Double
'--------------------------------------------------------------------------------------------------------------------------------
'Cal :  ptTextPrompt : เลขที่ Barcode ที่ระบุหรือสแกนเข้ามา
'Ret : Rate Cpn
'Cmt : *KT 54-06-21 V.4.0.1 Ph.4.0(Member Point)Req.6 Promotion Coupon(Mall Coupon)
'm_bResults = True :  เป็นหมายเลขที่ใช้งานได้อาจเป็นแบบ Barcode configuration หรือ แบบที่ Casher ระบุเอง False : เป็นหมายเลขที่ใช้งานไม่ได้ เพราะเข้า Barcode configuration แต่ไม่ผ่าน MOD 10
'm_bInStoreBar = True : Barcode configuration  False : ไม่เข้า Barcode configuration
'm_tID = เลขที่ Barcode
'm_tIDRuning = เลขที่ Running ที่ตัดจาก เลขที่ Barcode ที่ส่งมา ถ้าไม่เข้า Barcode configuration  ค่าจะเท่ากับ m_tID
'--------------------------------------------------------------------------------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tAmt As String
    Dim tDec As String
    Dim nDec As Long
    Dim bStep As Boolean
On Error GoTo ErrHandle
    m_bResults = True 'Default กรณีการระบุเลขด้วยตัวเอง
    tSql = "SELECT FTBafCode,FTBafCodeType,FNBafMatchStr,FNBafItemSize,FNBafPLUSize,FNBafPLULocate,FNBafPriceSize,FNBafPriceLocate"
    tSql = tSql & ",FNBafPriceDec,FTBafStaChkDigit"
    tSql = tSql & " FROM TCNMSkuBarCon"
    tSql = tSql & " WHERE FNBafItemSize = " & Len(ptTextPrompt)
    tSql = tSql & " AND FTBafCodeType = '4'"
    tSql = tSql & " ORDER BY FTBafCode"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then  'read master
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            orsTemp.Filter = "FNBafMatchStr = '" & Mid(ptTextPrompt, 1, 1) & "'"
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                 orsTemp.MoveFirst
                 m_tStaChkDigit = SP_FEDtChkString(orsTemp, "FTBafStaChkDigit")
                 bStep = True
                 If m_tStaChkDigit = "1" Then
                    If GETPmtCpnCharEnd(ptTextPrompt) <> Right(ptTextPrompt, 1) Then
                        bStep = False
                    End If
                 End If
                 If bStep Then
                     tAmt = Mid(ptTextPrompt, Val(Format(orsTemp("FNBafPriceLocate"))), Val(Format(orsTemp("FNBafPriceSize"))))
                     nDec = Val(Format(orsTemp("FNBafPriceDec")))
                     If nDec > 0 Then
                        tDec = Right(tAmt, nDec)
                        If CDbl(tDec) > 0 Then
                            tAmt = Mid(tAmt, 1, Len(tAmt) - nDec)
                            GETPmtCpnBarcode = CDbl(tAmt) + ((CDbl(tDec)) / (10 ^ nDec))
                        Else
                             tAmt = Mid(tAmt, 1, Len(tAmt) - nDec)
                            GETPmtCpnBarcode = CDbl(tAmt)
                        End If
                    Else
                        GETPmtCpnBarcode = CDbl(tAmt)
                    End If
                    m_tID = ptTextPrompt 'Mid(ptTextPrompt, Val(Format(orsTemp("FNBafPLULocate"))), Val(Format(orsTemp("FNBafPLUSize"))))
                    m_tIDRuning = Mid(ptTextPrompt, Val(Format(orsTemp("FNBafPLULocate"))), Val(Format(orsTemp("FNBafPLUSize"))))
                    m_bInStoreBar = True
                    m_bResults = True  'กรณีเข้า Barcode configuration แต่ผ่าน MOD 10
                Else
                    GETPmtCpnBarcode = "0"
                    m_tStaChkDigit = "2" 'No Tik
                    m_tID = ptTextPrompt
                    m_tIDRuning = ptTextPrompt
                    m_bInStoreBar = False
                    m_bResults = False 'กรณีเข้า Barcode configuration แต่ไม่ผ่าน MOD 10
                End If
            Else
                GETPmtCpnBarcode = "0"
                m_tStaChkDigit = "2" 'No Tik
                m_tID = ptTextPrompt
                m_tIDRuning = ptTextPrompt
                m_bInStoreBar = False
                 m_bResults = True 'กรณีการระบุเลขด้วยตัวเอง
            End If
        Else
            GETPmtCpnBarcode = "0"
            m_tStaChkDigit = "2" 'No Tik
            m_tID = ptTextPrompt
            m_tIDRuning = ptTextPrompt
            m_bInStoreBar = False
             m_bResults = True 'กรณีการระบุเลขด้วยตัวเอง
        End If
    Else
        GETPmtCpnBarcode = "0"
        m_tStaChkDigit = "2" 'No Tik
        m_tID = ptTextPrompt
        m_tIDRuning = ptTextPrompt
        m_bInStoreBar = False
         m_bResults = True 'กรณีการระบุเลขด้วยตัวเอง
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Property
ErrHandle:
    GETPmtCpnBarcode = "0"
    m_tStaChkDigit = "2" 'No Tik
    m_tID = ptTextPrompt
    m_tIDRuning = ptTextPrompt
    m_bInStoreBar = False
     m_bResults = True 'กรณีการระบุเลขด้วยตัวเอง
    Call SP_DATxRsNothing(orsTemp)
End Property

Public Property Get GETPmtCpnBarcodeInfo(ByVal ptID As String) As Boolean
'--------------------------------------------------------------------------------------------------------------------------------
'Ret : Rate Cpn
'Cmt : *KT 54-06-21 V.4.0.1 Ph.4.0(Member Point)Req.6 Promotion Coupon(Mall Coupon)
'--------------------------------------------------------------------------------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT * "
    tSql = tSql & " FROM TCNMSkuBarCon"
    tSql = tSql & " WHERE FTBafCodeType = '4'"
    tSql = tSql & " AND FTBafCode = '" & ptID & "'"
    tSql = tSql & " ORDER BY FTBafCode"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then  'read master
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                 orsTemp.MoveFirst
                 GETPmtCpnBarcodeInfo = True
                 m_tStaChkDigit = SP_FEDtChkString(orsTemp, "FTBafStaChkDigit")
            Else
                GETPmtCpnBarcodeInfo = False
                m_tStaChkDigit = "2" 'No Tik
            End If
        Else
            GETPmtCpnBarcodeInfo = False
            m_tStaChkDigit = "2" 'No Tik
        End If
    Else
        GETPmtCpnBarcodeInfo = False
        m_tStaChkDigit = "2" 'No Tik
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Property
ErrHandle:
    GETPmtCpnBarcodeInfo = False
    m_tStaChkDigit = "2" 'No Tik
    Call SP_DATxRsNothing(orsTemp)
End Property

Public Property Get GETPmtCpnBarcodeID(ByVal ptID As String, ByVal ptRuning As String, ByVal pcVal As Double) As String
'--------------------------------------------------------------------------------------------------------------------------------
'Ret : Rate Cpn
'Cmt : *KT 54-06-21 V.4.0.1 Ph.4.0(Member Point)Req.6 Promotion Coupon(Mall Coupon)
'--------------------------------------------------------------------------------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim nO As Long
    Dim nS As Long
    Dim nL As Long
    Dim nD As Long
    Dim nVal As Long
    Dim cDecx As Double
    Dim tChrChk As String
    Dim cValMax As Double
On Error GoTo ErrHandle
    tSql = "SELECT FNBafItemSize,FNBafMatchStr,FNBafPLULocate,FNBafPLUSize,FNBafPriceLocate"
    tSql = tSql & " ,FNBafPriceSize,FNBafPriceDec,FTBafStaChkDigit"
    tSql = tSql & " FROM TCNMSkuBarCon"
    tSql = tSql & " WHERE FTBafCodeType = '4'"
    tSql = tSql & " AND FTBafCode = '" & ptID & "'"
    tSql = tSql & " ORDER BY FTBafCode"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then  'read master
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                 orsTemp.MoveFirst
                 'ตรวจสอบ
                 '1. ตรวจสอบจำนวนเต้มของมูลค่า Coupon ว่าใส่ลง Barcode ส่วนจำนวนเต็มได้หรือไม่
                 nL = SP_FEDvChkNumeric(orsTemp, "FNBafPriceSize")
                 nD = SP_FEDvChkNumeric(orsTemp, "FNBafPriceDec")
                 nS = (pcVal \ 10) * 10
                 If Len(CStr(nS)) > (nL - nD) Then
                    GETPmtCpnBarcodeID = ""
                    Call SP_DATxRsNothing(orsTemp)
                    Exit Property
                 End If
                 
                 GETPmtCpnBarcodeID = SP_FEDtChkString(orsTemp, "FNBafMatchStr")
                 nS = SP_FEDvChkNumeric(orsTemp, "FNBafPLULocate")
                 nL = SP_FEDvChkNumeric(orsTemp, "FNBafPLUSize")
                 ptRuning = GETtRunNoPmtCpn(nL)
                 If ptRuning = "" Then
                    GETPmtCpnBarcodeID = ""
                    Call SP_DATxRsNothing(orsTemp)
                    Exit Property
                 End If
                 
                 GETPmtCpnBarcodeID = GETPmtCpnBarcodeID & IIf(nS - 2 > 0, String(nS - 2, "0"), "") & Format(ptRuning, String(nL, "0"))
                 nO = nS + nL
                 
                 nS = SP_FEDvChkNumeric(orsTemp, "FNBafPriceLocate")
                 nL = SP_FEDvChkNumeric(orsTemp, "FNBafPriceSize")
                 nD = SP_FEDvChkNumeric(orsTemp, "FNBafPriceDec")
                 nVal = Int(pcVal)
                 GETPmtCpnBarcodeID = GETPmtCpnBarcodeID & IIf(nO - nS > 0, String(nO - nS, "0"), "") & Format(nVal, String(nL - nD, "0"))
                 cDecx = pcVal - CDbl(nVal)
                 GETPmtCpnBarcodeID = GETPmtCpnBarcodeID & Format(cDecx, String(nD, "0"))
                 nL = SP_FEDvChkNumeric(orsTemp, "FNBafItemSize")
                 nS = Len(GETPmtCpnBarcodeID)
                 GETPmtCpnBarcodeID = Mid(GETPmtCpnBarcodeID, 1, nL) & IIf(nL - nS > 0, String(nL - nS, "0"), "")
                 m_tStaChkDigit = SP_FEDtChkString(orsTemp, "FTBafStaChkDigit")
                 If m_tStaChkDigit = "1" Then
                    tChrChk = GETPmtCpnCharEnd(GETPmtCpnBarcodeID)
                    If tChrChk <> "" Then
                        GETPmtCpnBarcodeID = Left(GETPmtCpnBarcodeID, Len(GETPmtCpnBarcodeID) - 1) & tChrChk
                    End If
                End If
            Else
                GETPmtCpnBarcodeID = ""
                m_tStaChkDigit = "2"
            End If
        Else
            GETPmtCpnBarcodeID = ""
             m_tStaChkDigit = "2"
        End If
    Else
        GETPmtCpnBarcodeID = ""
         m_tStaChkDigit = "2"
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Property
ErrHandle:
    GETPmtCpnBarcodeID = ""
    m_tStaChkDigit = "2" 'No Tik
    Call SP_DATxRsNothing(orsTemp)
End Property

Public Function GETPmtCpnCharEnd(ByVal ptID As String) As String
'------------------------------------------------------
' '*KT 54-07-26 V.4.0.1 Phase4.0(Member Point) Promotion Coupon ตรวจสอบ barcode config
'Ret: True : ผ่านการตรวจสอบการตรวจสอบหมายเลขตามสูตรการ Mod 10 ของ THE MALL False : ไม่ผ่านการตรวจสอบการตรวจสอบหมายเลขตามสูตรการ Mod 10 ของ THE MALL
'Cmt: ทำการตรวจสอบหมาเลข Mall Card ID ตามสูตร Mod
'Mod 10  Card Check Digit Validation
'ตัวอย่าง  หมายเลขบัตร  7102 3500 1865 3098
'               Check Digit คือ  8 มาจาก Digit สุดท้ายของหมายเลข Mall Card
'ขั้นตอน
'1.  เอาหลักลำดับคี่ (หลักที่ 1,3,5,7,9,11,13,15)  X2  ส่วนหลักคู่ (หลักที่ 2,4,6,8,10,12,14) ไม่ต้องทำการคูณ
'           1           2           3           4           5           6           7           8           9           10          11          12          13          14          15
'
'           7           1           0           2           3           5           0           0           1           8             6             5           3             0             9
'     x    2                        2                         2                        2                        2                           2                          2                             2
'         14           1           0           2           6           5           0           0          2            8            12            5          6              0            18

'2.  หลักลำดับคี่ ที่ผลคูณออกมาเป็นเลข 2 หลักให้นำมาบวกกัน
'         1+4     + 1         +0         +2         +6         +5        +0         +0         +2         +8         +1+2        +5         +6          +0          +1+8
'3.  นำผลลัพท์ที่ได้ทั้งหลักคี่และหลักคู่ รวมกัน
'                       5+1+0+2+6+5+0+0+2+8+3+5+6+0+9  =   52
'4.  แล้วเอา 10 ตั้งลบด้วย หลักหน่วยที่ได้จากผลบวก
'                      Check Digit     10 -2   =  8
'------------------------------------------------------
    Dim nI As Integer
    Dim nMax As Integer
    Dim nTemp As Integer
    Dim nFac As Integer
    Dim nMod As Integer
    Dim nTotal As Integer
On Error GoTo ErrHandle
    If ptID <> "" Then
        nMax = Len(Trim(ptID)) - 1
        ReDim aW_Digit(nMax) As Integer
        For nI = 0 To nMax
            aW_Digit(nI) = CInt(Mid(ptID, nI + 1, 1))
        Next nI
        nTotal = 0
        nFac = 0
        For nI = 0 To nMax - 1
            Select Case nI + 1
                Case 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 31 'คี่
                    nFac = aW_Digit(nI) * 2
                    If nFac > 9 Then
                        nTemp = nFac
                        nFac = (nFac Mod 10) + (nTemp \ 10)
                    End If
                Case 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30 'คู่
                    nFac = aW_Digit(nI)
                Case Else
                    If (nI + 1) Mod 2 = 0 Then
                        'คู่
                        nFac = aW_Digit(nI)
                    Else
                        'คี่
                        nFac = aW_Digit(nI) * 2
                        If nFac > 9 Then
                            nTemp = nFac
                            nFac = (nFac Mod 10) + (nTemp \ 10)
                        End If
                    End If
            End Select
            nTotal = nTotal + nFac
        Next nI
       nMod = nTotal Mod 10
       GETPmtCpnCharEnd = (10 - nMod) Mod 10
    Else
        GETPmtCpnCharEnd = ""
    End If
    Exit Function
ErrHandle:
    GETPmtCpnCharEnd = ""
End Function


Public Property Get GETSKUSaleWeight() As Double
    GETSKUSaleWeight = m_cSaleWeight
End Property

Public Property Get GETSKUWeight() As Double
    GETSKUWeight = m_cWeight
End Property

Public Property Get GETSKUTextSaleWeight() As String
    GETSKUTextSaleWeight = m_tSaleWeight
End Property

Public Property Get GETSKUTextWeight() As String
    GETSKUTextWeight = m_tWeight
End Property

Public Property Get GETSKUDecWeight() As String
    GETSKUDecWeight = m_nDecWeight
End Property

Public Property Get GETSKUDecSale() As String
    GETSKUDecSale = m_nDecPrice
End Property

Public Property Get GETInStoreBar() As Boolean
    GETInStoreBar = m_bInStoreBar
End Property

Public Property Get GETInStaWPrint() As String
    GETInStaWPrint = m_tStaWPrint
End Property

Public Property Get GETInStaChkDigit() As String
    GETInStaChkDigit = m_tStaChkDigit
End Property

Public Property Get GETID() As String
    GETID = m_tID
End Property

Public Property Get GETIDRuning() As String
    GETIDRuning = m_tIDRuning
End Property

Public Property Get GETResults() As Boolean
    GETResults = m_bResults
End Property


Public Function GETtRunNoPmtCpn(ByVal pnSize As Long) As String
'------------------------------------------------------
'Cal :  ส่งขนาด Digit เลขที่ Running
'Ret:   เลขที่ Runing ของ Promotion Coupon ในรูปแบบ Format ตามที่ Config
'Cmt:  *KT 54-08-20 V.4.0.1 Req.6 Ph.4.0(Member Point) Promotion Coupon
'ข้อจำกัด :  1.การ Update เลขที่ Running เลยทำให้เครื่องจุดขายอื่นได้เลขที่ Running ล่าสุดไม่ซ้ำกัน
'                 2.มีโอกาศที่ Running จะกระโดดข้าม ถ้าพบว่าเลขที่ Running ที่ได้ไปแล้วและมีการ Update นั้น เกิดการ Error ที่จุดอื่น
'-----------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim nTemp As Long
    Dim tTemp As String
    Dim nEff As Long
On Error GoTo ErrHandle
    tSql = "SELECT FTMcpLRunning FROM TCNMMallCpn WHERE FNDctNo =8"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , , True) = 0 Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            nEff = 0
            nTemp = CInt(Format(SP_FEDtChkString(orsTemp, "FTMcpLRunning"))) + 1
            If Len(CStr(nTemp)) > pnSize Then
                nTemp = 1
            End If
            tTemp = CStr(nTemp)
            tSql = "UPDATE TCNMMallCpn SET FTMcpLRunning = '" & tTemp & "' WHERE FNDctNo = 8"
            If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , nEff, , , True) = 0 Then
                If nEff <> 1 Then
                    nEff = 0
                    tSql = "INSERT INTO  TCNMMallCpn"
                    tSql = tSql & vbCrLf & "("
                    tSql = tSql & vbCrLf & "FNDctNo,FCMcpMinAmt"
                    tSql = tSql & vbCrLf & ",FTMcpLRunning,"
                    tSql = tSql & vbCrLf & tCS_CNLastFedIns
                    tSql = tSql & vbCrLf & "VALUES ("
                    tSql = tSql & vbCrLf & "8,0"
                    tSql = tSql & vbCrLf & ",'" & tTemp & "'"
                    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , nEff, , , True)
                    If nEff = 1 Then
                        GETtRunNoPmtCpn = Format(tTemp, String(pnSize, "0"))
                    Else
                        GETtRunNoPmtCpn = ""
                    End If
                Else
                    GETtRunNoPmtCpn = tTemp
                End If
            End If
            
        Else
            GETtRunNoPmtCpn = ""
        End If
    Else
        GETtRunNoPmtCpn = ""
    End If
    Exit Function
    Call SP_DATxRsNothing(orsTemp)
ErrHandle:
    GETtRunNoPmtCpn = ""
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Function
