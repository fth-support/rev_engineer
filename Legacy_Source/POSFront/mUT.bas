Attribute VB_Name = "mUT"
Option Explicit

Public nUT_ComPile As CS_Compile
Public bUT_FontBold As Boolean
Public bUT_FontItalic As Boolean
Public bUT_FontStrikeThru As Boolean
Public bUT_FontUnderLine As Boolean
Public bUT_ShowSumSgOut As Boolean
Public bUT_PromptCustomer As Boolean
Public bUT_FuncLockKB As Boolean        'status when call Ctrl+F1
Public bUT_SignOut As Boolean
Public bUT_VoidBrows As Boolean

'NUMBER
'for color
Public nUT_ColorFGrid As Long
Public nUT_ColorBGrid As Long
Public nUT_ColorOopGotFocus As Long
Public nUT_ColorBOlaPrice As Long

'general
Public nUT_FontSize As Long
Public nUT_SeqRC As Long
Public nUT_ChgTimeDelay As Long
Public nUT_GridMaxRow As Long                               'grid show maximum for set TopRow

'CURRENCY
'Public cUT_CmpTax As Double
Public tUT_CmpTaxCode As String
Public cUT_Qty As Double                      'default Qty =1  จำนวนชิ้นสินค้า
Public cUT_StkFac As Double
Public cUT_SetPrice As Double           'ราคารายการสุดท้าย
Public cUT_SumVat As Double            'ผลรวม Vat
Public cUT_SumNoVat As Double      'ผลรวม No Vat
Public cUT_SumNoDis As Double      'ผลรวม No Discount
Public cUT_SumNoDisNoVat As Double      'ผลรวม No Discount
Public cUT_SumNoVatAfDis As Double      'ผลรวม No Vat       '*Tee 47-10-11
Public cUT_VatAmt  As Double           'ภาษีมุลค่าเพิ่ม
Public cUT_VatableAmt  As Double           'มูลค่า Vatable
Public cUT_RegPrice As Double
Public cUT_SalePrice As Double
Public cUT_MemberPrice As Double '*KT 52-02-13 ตัวแปรเก็บค่า Member Price
Public cUT_SKUCost As Double
'STRING
Public tUT_UsrLevel As String               'lever of user login  1,2,3...
Public tUT_UsrPwd As String     'password of user who login
'for browser
Public tUT_BrwTable As String
Public tUT_BrwField As String
Public tUT_BrwHead As String
Public tUT_BrwRet As String
Public tUT_BrwWhere As String
Public tUT_BrwWCaption As String
Public tUT_SaleDate As String                 'วันที่การขายจะเปลี่ยนตามการขาย (next or previos)
Public tUT_TextPrompt As String             'store valut from wmain.otbprompt.text
Public tUT_RefDocNo As String               'เลขที่บิลขายที่อ้างอิงถึงเมื่อมีการ refund
Public tUT_LoginCode As String              'login code= Year+Month+Max(Right(LastDocNumber,6))   Form TPSTLogIn
'Public tUT_SKUCode As String
Public tUT_SKUCode As String '*KT 50-12-25
Public tUT_SKUBarC As String '*KT 50-12-25
Public tUT_SKUName As String '*KT 50-12-25
Public tUT_SKUName2 As String '*KT 50-12-25
Public tUT_FuncKBName As String
Public tUT_SKUUnitC As String       'unit code
Public tUT_SKUUnitN As String       'unit name
Public cUT_SKUUnitF As Double     'unit factor
'*GP 46-12-29
Public cUT_ChgAmt As Double
Public cUT_DisAmt As Double
Public bUT_PromptSale As Boolean 'Prompt for price
'หน้าจอ Main
Public cUT_Total As Double 'ยอดรวมมูลค่าสินค้าทั้งหมด RegPrice * Qty
Public cUT_SubTotal As Double 'ยอดรวมมูลค่าสินค้าทั้งหมดที่ลดส่วนลดรายตัวแล้ว
Public cUT_SubPro As Double 'ยอดรวมมูลค่าสินค้าที่มีราคาขายหลังหักโปรโมชั่น หมายความว่ามีโปรโมชั่นในสิค้านั้นที่ลดส่วนลดรายตัวแล้ว
Public cUT_SubReg As Double 'ยอดรวมมูลค่าสินค้าที่มีราคาขายปกติที่ลดส่วนลดรายตัวแล้ว
Public nUT_SumQty As Long 'ยอดจำนวนรวมสินค้าทั้งหมด(สินค้าที่เป็นน้ำหนักให้นับเป็น 1 หน่วย)


''printer
'Public bUT_PrinterOk As Boolean
'Public nUT_HdlePrinter As Long
'Public tUT_Printer As String 'ชื่อเครื่อง Printer
'Public tUT_PrintType As String '2 : รวมพิมพ์โดยตรง 3 : รวมพิมพ์โดยวินโดว์
''print Thai
'Public bUT_T1Print As Boolean
'Public bUT_T3Print As Boolean
'Public bUT_NotJust As Boolean

Public Function UT_tToken(ptSource$, ByVal ptSep$, Optional piLoop%, Optional pbRetOld As Boolean) As String
'----------------------------------------------------------
'   Call:   ptSource source string ("Com1;Com2")
'               ptSep separator string (";")
'               piLoop is optional if exactly know the position
'               pbRetOld is optional if true source still the same, otherwise token out
'   Ret:    first token from source string ("Com1")
'               ptSource without first token ("Com2")  [depend on pbRet]
'----------------------------------------------------------
    Dim i%, T$, n%, j%
    Dim tSrc$
    
    tSrc = ptSource
    If (IsMissing(piLoop)) Or (piLoop = 0) Then
        n = 1
    Else
        n = piLoop
    End If
    
    For j = 1 To n
        i = InStr(UCase(tSrc), UCase(ptSep))
        If i = 0 Then
            T = tSrc            'not found separator (n could be 1 or >1)
            tSrc = ""
        Else
            T = Left(tSrc, i - 1)
            tSrc = Mid(tSrc, i + Len(ptSep))
        End If
    Next j
    
    If Not pbRetOld Then        'default pbRetOld = False, always token out
        ptSource = tSrc
    End If
    
    UT_tToken = T

End Function

Public Sub UT_CreateSaleTableTMP()
'---------------------------------------------------------
'   Cmt:   create ในกรณีที่ยังไม่มี 3 tables ที่ต้องการ,
'               3 tables จะถูก drop เมื่อมีการโพรเซสแต่ละวัน
'               ใช้ได้ทั้ง Access and SQL
'---------------------------------------------------------
    Dim tSql As String
    Dim tSentComplate As String
    Dim nCount As Long
    Dim orsTemp As ADODB.Recordset
    Dim tFile As String
    Dim oFSO As New FileSystemObject
    
    nCount = 0
    tSentComplate = GetSetting("AdaPosMall", "POSFront", "SentSale")
    If tSentComplate = "1" Then
        '*KT 52-11-20 ทำการไล่ Commit Tran ของ Local แสดงว่าก่อนหน้า ServiceTranfer น่าจะมีปัญหาเรื่องการส่งข้อมูลที่เกิดการ Log ของข้อมูล
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:Start Commit Transaction" & " Form:" & "mUT" & " Control:General" & " Events:UT_CreateSaleTableTMP")
        Call SP_SQLxCommitTranLocal
        Call SP_SQLxCommitTranServer
        
        tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceTranfer.exe"
        If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
                tFile = tFile + "\POS\ServiceTranfer.exe"
                If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                    Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
                    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Can not run ServiceTranfer.exe" & " Form:mDB" & " Control:General" & " Events:Main")
                    Call SP_APPxEndProgram
                End If
        End If
        nVB_lServiceTransWnd = FindWindow(vbNullString, "ServiceTranfer")
         If nVB_lServiceTransWnd = 0 Then
            Call ShellEx(nVB_lServiceTransWnd, "open", tFile, "", "", 10)
        End If
        
    End If
    
    Do While tSentComplate = "1" And nCount < nVB_CNWaitTranfer
        SleepEx 1000, 0
        nCount = nCount + 1
        tSentComplate = GetSetting("AdaPosMall", "POSFront", "SentSale")
    Loop
    
    'อาจมีปัญหาการขายนอกสถานที ???
    If tSentComplate = "1" And nCount >= nVB_CNWaitTranfer Then
        MsgBox "แสดงข้อความเตือนไม่ยอมให้ทำการขายจนกว่า ServiceTransfer จะ Set Flag ให้เท่ากับ 0", vbCritical
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "แสดงข้อความเตือนไม่ยอมให้ทำการขายจนกว่า ServiceTransfer จะ Set Flag ให้เท่ากับ 0" & " Form:mUT" & " Control:General" & " Events:UT_CreateSaleTableTMP")
        Call SP_APPxEndProgram
    End If
    '-------------------------------------------------------------------------------------------------------------
    'ตรวจสอบก่อนว่ามี Transaction ไหนบ้างใน Local ที่ยังไม่การส่งข้อมูลโดยผ่าน ServiceTransfer ถ้ามีการส่งข้อมูลหมดให้ Truncate ไม่หมด Delete
    '*KT 52-01-19 V 1.5.0
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTSalHD WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTSalHD"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTSalHD WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '*KT 52-01-19 V 1.5.0
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTSalDT WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTSalDT"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTSalDT WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '*KT 52-01-19 V 1.5.0
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTSalRC WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTSalRC"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTSalRC WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '*KT 52-01-19 V 1.5.0
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTSalCD WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTSalCD"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTSalCD WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '*KT 52-01-19 V 1.5.0
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTLogDT WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTLogDT"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTLogDT WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '*KT 54-08-23 V.4.0.1
'    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTSalePoint WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
'    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTSalePoint"
'    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTSalePoint WHERE FTStaSentOnOff = '1'"
'    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DELETE TPSTSalePoint WHERE (FTStaSentOnOff = '1' AND FTRemark = '1')"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '*KT 54-10-06 V.4.0.1Req.6 Phase4.0(Member Point) Promotion Coupon ขอทำการแก้ไขตาม  DCR_5410001
     tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTSalPmtCpn WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTSalPmtCpn"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTSalPmtCpn WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    
    '*KT 52-06-01 PH 2.0.0
    'DEPOSITHD
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTDEPOSITHD WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTDEPOSITHD"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTDEPOSITHD WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    'DEPOSITDT
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTDEPOSITDT WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTDEPOSITDT"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTDEPOSITDT WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    'DEPOSITRC
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTDEPOSITRC WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTDEPOSITRC"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTDEPOSITRC WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    'DEPOSITCD
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTDEPOSITCD WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTDEPOSITCD"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTDEPOSITCD WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    'HOLDHD
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTHOLDHD WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTHOLDHD"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTHOLDHD WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    'HOLDDT
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTHOLDDT WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTHOLDDT"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTHOLDDT WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    'HOLDCD
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTHOLDCD WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTHOLDCD"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTHOLDCD WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
   'VOUCHERHD
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTVOUCHERHD WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTVOUCHERHD"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTVOUCHERHD WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    'VOUCHERRC
    tSql = "IF NOT EXISTS(SELECT FTStaSentOnOff FROM TPSTVOUCHERRC WHERE ISNULL(FTStaSentOnOff,'0') <> '1' )"
    tSql = tSql & vbCrLf & "TRUNCATE TABLE  TPSTVOUCHERRC"
    tSql = tSql & vbCrLf & "ELSE DELETE   TPSTVOUCHERRC WHERE FTStaSentOnOff = '1'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    
    
    'for TPSHDXXX
    '*KT 52-01-16 ให้มีการตรวจสอบใน Table Temp ก่อนว่ามีข้อมูลที่ Trigger ยังไม่ส่งสู่ Table จริงหรือไม่ ถ้ามีจะไม่ Drop และ Clone ก่อน
    tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
    If SP_TBLbChkExist(tVB_CNTblHD, oDB.ocnOffLine) Then
        tSql = "SELECT  FTShdStaPrcTmp FROM " & tVB_CNTblHD & "  WHERE ISNULL(FTShdStaPrcTmp,'0') <> '1'  AND  FTShdStaDoc = '1'"
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine)
        If orsTemp.EOF And orsTemp.BOF Then
             tSql = "DROP TABLE " & tVB_CNTblHD
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
            tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalHD" & "',@DestinationTable ='" & tVB_CNTblHD & "'"
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
            tSql = "TRUNCATE TABLE  " & tVB_CNTblHD '*KT 51-10-02 3.0.127
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
    Else
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalHD" & "',@DestinationTable ='" & tVB_CNTblHD & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblHD '*KT 51-10-02 3.0.127
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    tSql = "SELECT  FTShdStaPrcTmp FROM " & tVB_CNTblHD & "  WHERE ISNULL(FTShdStaPrcTmp,'0') <> '1'  AND  FTShdStaDoc = '1'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine)
    
    'for TPSDTXXX
    If SP_TBLbChkExist(tVB_CNTblDT, oDB.ocnOffLine) Then
        If orsTemp.EOF And orsTemp.BOF Then
           tSql = "DROP TABLE " & tVB_CNTblDT
           Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
           tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalDT" & "',@DestinationTable ='" & tVB_CNTblDT & "'"
           Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
           tSql = "TRUNCATE TABLE  " & tVB_CNTblDT '*KT 51-10-02 3.0.127
           Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
    Else
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalDT" & "',@DestinationTable ='" & tVB_CNTblDT & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblDT '*KT 51-10-02 3.0.127
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If

    
    'for TPSDTXXXTmp
    If SP_TBLbChkExist(tVB_CNTblDTTmp, oDB.ocnOffLine) Then
        If orsTemp.EOF And orsTemp.BOF Then
            tSql = "DROP TABLE " & tVB_CNTblDTTmp
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
             tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalDT" & "',@DestinationTable ='" & tVB_CNTblDTTmp & "'"
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
             tSql = "TRUNCATE TABLE  " & tVB_CNTblDTTmp '*KT 51-10-02 3.0.127
             Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
    Else
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalDT" & "',@DestinationTable ='" & tVB_CNTblDTTmp & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblDTTmp '*KT 51-10-02 3.0.127
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If

    
    'for TPSRCXXX
    If SP_TBLbChkExist(tVB_CNTblRC, oDB.ocnOffLine) Then
        If orsTemp.EOF And orsTemp.BOF Then
            tSql = "DROP TABLE " & tVB_CNTblRC
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
            tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalRC" & "',@DestinationTable ='" & tVB_CNTblRC & "'"
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
            tSql = "TRUNCATE TABLE  " & tVB_CNTblRC '*KT 51-10-02 3.0.127
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
    Else
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalRC" & "',@DestinationTable ='" & tVB_CNTblRC & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblRC '*KT 51-10-02 3.0.127
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If

    
    'for TPSCDXXX
    If SP_TBLbChkExist(tVB_CNTblCD, oDB.ocnOffLine) Then
        If orsTemp.EOF And orsTemp.BOF Then
             tSql = "DROP TABLE " & tVB_CNTblCD
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
            tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalCD" & "',@DestinationTable ='" & tVB_CNTblCD & "'"
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
            tSql = "TRUNCATE TABLE  " & tVB_CNTblCD '*KT 51-10-02 3.0.127
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
    Else
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalCD" & "',@DestinationTable ='" & tVB_CNTblCD & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblCD '*KT 51-10-02 3.0.127
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If

    
    'for TPSCDXXXTmp
    If SP_TBLbChkExist(tVB_CNTblCDTmp, oDB.ocnOffLine) Then
        If orsTemp.EOF And orsTemp.BOF Then
            tSql = "DROP TABLE " & tVB_CNTblCDTmp
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
            tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalCD" & "',@DestinationTable ='" & tVB_CNTblCDTmp & "'"
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
            tSql = "TRUNCATE TABLE  " & tVB_CNTblCDTmp '*KT 51-10-02 3.0.127
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
    Else
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalCD" & "',@DestinationTable ='" & tVB_CNTblCDTmp & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblCDTmp '*KT 51-10-02 3.0.127
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    'Reindex Table Sale and Temp
    tSql = "DBCC DBREINDEX('TPSTSalHD',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTSalDT',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTSalRC',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTSalCD',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('" & tVB_CNTblHD & "',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('" & tVB_CNTblDT & "',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('" & tVB_CNTblCD & "',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('" & tVB_CNTblRC & "',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTDepositHD',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTDepositDT',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTDepositCD',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTDepositRC',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTHoldHD',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTHoldDT',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTHoldCD',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTVoucherHD',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTVoucherRC',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTLogDT',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
     '*KT 54-08-23 V.4.0.1
    tSql = "DBCC DBREINDEX('TPSTSalePoint',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "DBCC DBREINDEX('TPSTClose',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
     '*KT 54-10-06 V.4.0.1Req.6 Phase4.0(Member Point) Promotion Coupon ขอทำการแก้ไขตาม  DCR_5410001
    tSql = "DBCC DBREINDEX('TPSTSalPmtCpn',' ',90)"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
End Sub

Public Sub UT_GenSaleNewDoc(Optional ByVal pbShowChg As Boolean = False)
'----------------------------------------------------------
'   Call:
'   Cmt:   gen new sale document
'----------------------------------------------------------
    Dim tDisplay$
On Error GoTo ErrHandle
    Call UT_ClearGenNew
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnSaleTaxInc, tEN_TrnReturnTaxInc, tEN_TrnTotalReportTaxInc, tEN_TrnOverShortReportTaxInc, tEN_TrnClosingReportTaxInc
            tVB_Transaction = aVB_TransMode(37).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(37).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(37).PrintType
        Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax, tEN_TrnTotalReportManualTax, tEN_TrnOverShortReportManualTax, tEN_TrnClosingReportManualTax
            tVB_Transaction = aVB_TransMode(42).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(42).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(42).PrintType
        Case Else
            tVB_Transaction = aVB_TransMode(2).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(2).Name
            tVB_CNTypePrinter = aVB_TransMode(2).PrintType
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSale - 1).UsrLangSec
    End Select
    
    tVB_Running = SP_AUTtGenCode(tVB_CNTblHD)
    Call SP_INSxTransHD(tVB_CNTblHD)
    With wMain  '*KT 52-02-23 640 or 800
        .olaBarcode.Caption = ""
        .olaSKUName.Caption = ""
         If Not pbShowChg Then .olaPrice.Caption = "" '*KT 54-09-20 V.4.0.1 กรณีที่ไม่เป็น OPOS ในการแสดงเงินทอน
        .olaSumQty.Caption = ""
        .olaTSubReg.Caption = ""
        .olaTSubPro.Caption = ""
        .olaTSubTotal.Caption = ""
         If Not pbShowChg Then .olaChange.Visible = False '*KT 54-09-20 V.4.0.1 กรณีที่ไม่เป็น OPOS ในการแสดงเงินทอน
        .olaDocNo.Caption = tVB_Running
        .olaPos.Caption = tVB_CNTerminalNum
        .olaTransName.Caption = tVB_TransName
        .ogdForm.rows = 1
        .ogdForm.rows = nUT_GridMaxRow
        .olaRef.Caption = ""
        .ostForm.Panels(2).Text = "Ready..."
        .nW_CntSentSale = 0 '*Tao 61-09-24 เริ่มนับใหม่
    End With
'    '*Eaw 56-11-14 ถ้าไม่มีเงินทอนจะแสดงข้อความต้อนรับ
'    If cVB_Change > 0 Then
'        pbShowChg = True
'    Else
'        pbShowChg = False
'    End If
'    cVB_Change = 0 'clear ตัวแปรของเงินทอนหลังใช้ด้วยไม่งั้นค่านี้จะติดไปบิลต่อไปด้วย
    '----------------------------------------------------------------
'   If bVB_OPOSLineDisplayEnable And (Not pbShowChg) Then '*KT 54-09-20 V.4.0.1 กรณีที่ไม่เป็น OPOS ในการแสดงเงินทอน
    If bVB_OPOSLineDisplayEnable Then '*Eaw 56-11-29 Customer Display ให้แสดง WElcome ทุกครั้งที่เริ่มบิลใหม่
        If aVB_DSlipMag(0) = "" Or aVB_DSlipMag(1) = "" Then
             tDisplay = "Welcome ....;"
        Else
            tDisplay = aVB_DSlipMag(0) & ";" & Trim(Mid(aVB_DSlipMag(1), 1, 20)) & Space(20 - Len(Trim(Mid(aVB_DSlipMag(1), 1, 20))))
        End If
        Call SP_PrnPollDisplay(tDisplay)
    End If
    Exit Sub
ErrHandle:
     wMain.nW_CntSentSale = 0 '*Tao 61-09-24 เริ่มนับใหม่
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & "mUT" & " Control:mUT" & " Events:UT_GenSaleNewDoc")
End Sub

Public Sub UT_ClearGenNew()
'----------------------------------------------------------
'   Call:
'   Cmt:    clear value when gen new doc
'----------------------------------------------------------
        Dim nLoop As Long
    On Error GoTo ErrHandle
         dVB_TransDateRet = Date '*KT 67-06-07 TTK-ISS3000-873
        
         cVB_RoundChgRC = 0 '*Eaw 56-04-29  CommSheet TK-ISS3000-105-FSBT Answer.xlsx
         bVB_SalemanIDHD = False '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
         tVB_SalemanIDChg = ""
         tVB_SalemanIDOld = ""
         bVB_Saleman = True
        
        'main
        cUT_SubTotal = 0 'ยอดรวมมูลค่าสินค้าทั้งหมด
        cUT_SubPro = 0 'ยอดรวมมูลค่าสินค้าที่มีราคาขายหลังหักโปรโมชั่น หมายความว่ามีโปรโมชั่นในสิค้านั้น
        cUT_SubReg = 0 'ยอดรวมมูลค่าสินค้าที่มีราคาขายปกติ
        nUT_SumQty = 0 'ยอดจำนวนรวมสินค้าทั้งหมด(สินค้าที่เป็นน้ำหนักให้นับเป็น 1 หน่วย)
        cUT_SumNoVat = 0 'ยอดรวมมูลค่าสินค้าที่ไม่คิดภาษี
        cUT_SumVat = 0 'ยอดรวมมูลค่าสินค้าที่คิดภาษี
        cUT_VatAmt = 0 'ยอดภาษีมูลค่าเพิ่ม
        cUT_VatableAmt = 0 'ยอดยอดรวมมูลค่าสินค้าที่คิดภาษีลบด้วยยอดภาษีมูลค่าเพิ่ม
        cVB_Pay = 0 'ยอดที่ลูกระบุมา อาจรวมกับเงินทอน
        
        'payment
        cVB_Total = 0 'ยอดเงินทั้งหมด
        cVB_TotalDisc = 0 'ยอดรวมส่วนลด
        cVB_TotalDiscB = 0 'Discount Type = 1,2,3,4,5,6
        cVB_NetSale = 0 'ยอดรวมสุทธิหลังหักส่วนลด
        cVB_Balance = 0 'ยอดคงเหลือ
        tVB_TenderName = 0 'ชื่อ Tender
        tVB_TenderCode = "" 'รหัส Tender
        tVB_TenderType = "" 'ประเภท Tender
        cVB_Change = 0 'ยอดเงินทอน
        cVB_NetDutyFreeDep = 0 'ยอดเงินมัดจำของ Duty Free
        cVB_RoundChg = 0 'ยอดปัดเศษเงินทอน
        cVB_RoundRcv = 0 'ยอดปัดเศษทั้งบิล
        
        'Credit Card
        tVB_CrdApprCode = "" 'Authorization Number . (Approved Code from Bank.)
        tVB_CrdCardNo = "" '*หมายเลขอ้างอิง (เครดิต:เลขบัตร, เช็ค:เลขที่เช็ค …) (Number of  This Tender (Such as if Tender is Credit Card, it mean Credit Card Number and if Tender is Gift Voucher, It mean Gift Voucher Number. )
        tVB_CrdExpD = ""
        tVB_CrdExpM = ""
        tVB_CrdExpY = ""
        '*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
        bVB_CrdPOSNoOnTopFlag = False
        tVB_CrdPOSNoOnTopAlwType = ""
        tVB_CrdFloorAlwOnTop = ""
        tVB_CrdEvnID = "" '*KT 54-11-17 V.4.0.5 DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc Event ID ของบัตรเครดิตการ์ด
        
        'Foreign Currency
        tVB_FTRteCode = "" 'รหัสสกุลเงิน
        cVB_FCSrcRteAmt = 0 'จำนวนเงิน
        cVB_FCSrcRteFac = 0 'อัตราแลกเปลี่ยน
        tVB_FTRteCodeLocal = "" 'รหัสรหัสสกุลเงินท้องถิ่น
        
        'Credit Voucher
         tVB_CNCVNo = ""
         cVB_CNCVAmt = 0
        'Prompt sale
        bUT_PromptSale = False
        bVB_CNTrnPay = False
        
        'IPV
        bVB_IPV = False
        
        'Mall Card
        tVB_DiscMallCardID = ""
        tVB_DiscMallCardKey = ""
        tVB_DiscMallCardType = ""
        cVB_DiscMallRate = 0
        cVB_DiscMallRanking = 0
        nVB_DiscMallCount = 0
        nVB_RegMallCount = 0
        nVB_OnTopCount = 0
        tVB_StaMemPointAlw = "N" '*KT 52-02-17 member point
        tVB_MallCardGroup = ""
        tVB_MemEvnCode = "" '*KT 54-11-18 V.4.0.5  DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc
        tVB_MallCardIDRef = "" '*Tao 60-04-20
        
        'Emp Discount
        tVB_DiscEmpCardID = ""

        'member point
        tVB_MemPointID = ""
        tVB_MemPointIDTmp = "" '*KT 52-04-10
        nVB_CurMemPointCount = 0
        nVB_MemPointMnu = 0
        
        'Copy Sale
        bVB_CpySale = False
        
        bVB_SKUisPromotion = False
        
        'Special Gift voucher
        cVB_SpcCpnAmtAffect = 0
        
        'Specail Coupon
        cVB_SpcGiftVoucherAmtAffect = 0
        
        bUT_VoidBrows = False
        
        nDE_Group = 0
        
        bMainFunc = True
        
        nVB_CNScollItem = 0
        
        tVB_CNWheBill = "FTShdTransNo = '" & tVB_Running & "' AND  FTTmnNum = '" & tVB_CNTerminalNum & "' AND FDShdTransDate = '" & tUT_SaleDate & "'"

        bVB_CNReSumAmount = False
        
        bVB_CNTenderAct = False
        
         bVB_CNPayCrdNoSuccess = False
         
         bVB_CNTrnOnlineSuccess = False
         
         bVB_CNNoUpd = False
         
        '*KT 52-02-02 Voucher  PH2-ML-SF-081
        tVB_CancelVoucherNo = ""
        bVB_CancelVoucher = False
        cVB_CancelVoucherAmt = 0
        tVB_CancelTmnNumRef = ""
        tVB_CancelTransNoRef = ""
        
        '*KT 52-02-13 Member price
        bVB_CNMemberPrice = False
        xVB_CNMemberP.Expire = Date
        xVB_CNMemberP.FName = ""
        xVB_CNMemberP.ID = ""
        xVB_CNMemberP.LName = ""
        xVB_CNMemberP.MallCardExpire = Date
        xVB_CNMemberP.MallCardMemberPrice = ""
        xVB_CNMemberP.MallCardType = ""
        xVB_CNMemberP.Status = ""
        xVB_CNMemberP.Tel = ""
        
        
        bVB_ReConnectMem = True
        bVB_ReConnectEmp = True
        
        'Preview Day
        bVB_CNTrStaRegPreviewDay = False
        Call SP_INIxSwitchModePreviewDay(False)
        Call SP_INIxClearVarPrv
        
        tVB_CNSkuPointCooAlw = ""
        tVB_CNSkuPointAlw = ""
        tVB_CNSkuEventPoint = ""
        cVB_CNSkuPointAmt = 0
        cVB_CNSkuPointTime = 0
        dVB_CNSkuDatePoint = Date
        cVB_TotalSecsionTierAmt = 0
        cVB_CNTotalOntopByTicket = 0 '*KT 52-07-13 การคิด Point ของ Deposit
        cVB_CNTotalSpecialByTicket = 0 '*KT 52-07-13 การคิด Point ของ Deposit
        tVB_CNSkuPointAlwDiscountMall = "" '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด
        
        tVB_InStoreBar = ""
        
        cVB_AdjustMalCrdOfOntop = 0
        cVB_AdjustDiscMalCrdOfCrd = 0
        
        'Deposit
        bVB_DepositAll = False  '*KT 52-07-13 member point of deposit
        
        tVB_CNSkuStaAlw = ""
        tVB_SKUConsignFlag = "N" '*KT 52-08-06
        
        bVB_TicketNotSuccess = False
        
        tVB_CNUpdateEmpLimit = "" '*KT 53-01-28 Remain Discount for Employee discount
        cVB_CNRemainDiscAmt = 0  '*KT 53-01-28 Remain Discount for Employee discount
        tVB_CNRemainDiscID = "" '*KT 53-01-28 Communication Sheet PH3-ML-SB-074 Remain Discount for Employee.xls
        
        tVB_DiscEmpChkLimit = "N"
        
        tVB_CNTranMDownNo = "" '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
        
        '*KT 54-07-25 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon
        '*KT 54-10-06 V.4.0.1Req.6 Phase4.0(Member Point) Promotion Coupon ขอทำการแก้ไขตาม  DCR_5410001
'        tVB_CNPmtCpn.ID = ""
'        tVB_CNPmtCpn.Amt = 0
'        tVB_CNPmtCpn.Val = 0
'        tVB_CNPmtCpn.Qty = 0
        '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
        cVB_CNTenderAmtAlwPoint = 0
        
        '*KT 54-08-22 V.4.0.1 R4.0(Member Point) Req.1SKU Point with multiply by time
'        ReDim aVB_CNSKUPointByTime(0)
'        aVB_CNSKUPointByTime(0).tSKU = ""
'        aVB_CNSKUPointByTime(0).nSeqNo = 0
'        aVB_CNSKUPointByTime(0).cPoint = 0
'        aVB_CNSKUPointByTime(0).tType = "0"

        '*KT 54-10-10 V.4.0.1 เก็บยอดรวมสินค้าหลังหัดส่วนลดทุกชนิดเฉพาะสินค้าป้ายน้ำเงิน,ยอดรวมสินค้าหลังหัดส่วนลดทุกชนิดเฉพาะสินค้าป้ายเหลือง
        cVB_RegTotAmtAfDisc = 0
        cVB_ProTotAmtAfDisc = 0
        
        bVB_CNReIssueCancel = False
        cVB_CNCVNumRefCancelAmt = 0
        tVB_CNCVNumRefCancel = ""
        tVB_CNCVNumRefRCCancel = ""
        tVB_CNCVType = ""
        tVB_CNCVNoReIssue = ""
        cVB_CNCVAmtReIssue = 0
        oVB_CNReIssue.cEN_ReIssueAmt = 0
        oVB_CNReIssue.tEN_ReIssueNo = ""
        oVB_CNReIssue.tEN_ReIssueExp = ""
        oVB_CNReIssue.tEN_ReIssueTel = "" '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
        If (Not aVB_CNMallCard) = True Then
            'none
        Else
            For nLoop = 0 To UBound(aVB_CNMallCard)
                aVB_CNMallCard(nLoop).IDCard = ""
'                aVB_CNMallCard(nLoop).Amt = 0
            Next nLoop
        End If
        
    tVB_MemPointID = ""                 '*Ao 2013/02/27  TK POINT
    tVB_MemExp = ""
    tVB_ChkMemExp = ""
    bVB_CNCheckCV = False
'    bW_EnableRePrintCV1 = False  '*Bump 56-03-28 comment ไว้เนื่องจากถาม Pae แล้วไม่ได้ใช้งาน
'    bW_EnableRePrintCV2 = False  '*Bump 56-03-28 comment ไว้เนื่องจากถาม Pae แล้วไม่ได้ใช้งาน

    bVB_ChkEmpPoint = True '*Eaw 56-08-21
    tVB_SalemanID = "" '*Eaw 56-09-06 CommSheet TK-ISS3000-240
    tVB_CreditGroup = ""                '*Ao 2013/11/05
    '*Tao 61-08-15 Clear Return
    bVB_Reference = False
    bVB_Return = False
    '*KT 66-08-17 TTK-ISS3000-839
    'bVB_ReturnFrmBck = False
    bVB_ReturnAll = False
    bVB_ReturnAllBill = False
    'bVB_ReturnHaveDisc = False
    
    '*Amorn 65-08-09
     bVB_ChkRetGST7 = False
    '*Amorn 65-08-08 เคลียค่า eW_Deposit.tEN_DepositNo
    tVB_DepositNoGST7 = ""
    
    '*Amorn 65-08-26 STK-POS ComSheet CR803 - 004_20220824 เคลียค่า cW_RegAmt
'    cW_RegAmt = ""
    
    Exit Sub
ErrHandle:
    
 End Sub


Public Function UT_tStringFormatDate(ByVal ptDate$, ByVal ptFormat$) As String
'----------------------------------------------------------
'   Call:   pdDate as string date
'               ptFormat$ as format dd/mm/yyyy, yyyy/mm/dd
'
'   Ret:   string of date format dd/mm/yyyy, yyyy/mm/dd
'----------------------------------------------------------
    Select Case UCase(ptFormat)
        Case "DD/MM/YYYY"
            UT_tStringFormatDate = Right(ptDate, 2) & "/" & Mid(ptDate, 6, 2) & "/" & Left(ptDate, 4)
            
        Case "YYYY/MM/DD"
            UT_tStringFormatDate = ptDate
    End Select
    
End Function

Public Function UT_tStringDate(pdDate As Date) As String
'----------------------------------------------------------
'   Call:   pdDate as selected date
'
'   Ret:   string of date format "YYYY/MM/DD"
'----------------------------------------------------------
    'last update 2003/04/21 by korn
    UT_tStringDate = Format(Year(pdDate), "0000") & "/" & Format(Month(pdDate), "00") & "/" & Format(Day(pdDate), "00")
End Function

Public Function UT_tStringYYYYMM(pdDate As Date) As String
'----------------------------------------------------------
'   Call:   pdDate as selected date
'   Ret:   string of date format "YYYY/MM"
'----------------------------------------------------------
    UT_tStringYYYYMM = Format(Year(pdDate), "0000") & Format(Month(pdDate), "00")
End Function

'Public Function UT_tStringTime(pdDate As Variant) As String
''----------------------------------------------------------
''   Call:   pdDate as selected variant
''   Ret:   string of date format "hh:mm:ss"
''----------------------------------------------------------
'    UT_tStringTime = Format(pdDate, "hh:mm:ss")
'End Function

Public Function UT_tLastInsSql() As String
'----------------------------------------------------------
'   Call:
'
'   Ret:   string of query, last 3 field date/time/who
'----------------------------------------------------------
    Dim tTemp$
    Dim tName As String '*Eaw 56-11-01

'    tTemp = "'" & UT_tStringDate(Date) & "','" & Time$ & "','" & tVB_CNUserName & "','"
'    UT_tLastInsSql = tTemp & UT_tStringDate(Date) & "','" & Time$ & "','" & tVB_CNUserName & "'"

    '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    tName = SP_STRtSiggleCode(tVB_CNUserName)
    tTemp = "'" & UT_tStringDate(Date) & "','" & Time$ & "','" & tName & "','"
    UT_tLastInsSql = tTemp & UT_tStringDate(Date) & "','" & Time$ & "','" & tName & "'"
End Function

Public Function UT_tLastUpdSql() As String
'----------------------------------------------------------
'   Call:
'
'   Ret:   string of query, last 3 field date/time/who
'---------------------------------------------------------
    Dim tSql$
    Dim tName As String '*Eaw 56-11-01
    
    tName = SP_STRtSiggleCode(tVB_CNUserName)  '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    tSql = "FDDateUpd = '" & UT_tStringDate(Date) & "',"           'carefule between Date/Date$ it's different
    tSql = tSql & "FTTimeUpd = '" & Time$ & "',"
'    tSql = tSql & "FTWhoUpd = '" & tVB_CNUserName & "'"
    tSql = tSql & "FTWhoUpd = '" & tName & "'"  '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    UT_tLastUpdSql = tSql
    
End Function


Public Function UT_nPosCutTag() As Integer
'----------------------------------------------------------
'   Call:
'
'   Cmt:    set position to cut string in UT_tToken, Tha=1, Eng=2
'----------------------------------------------------------
    Dim nPos As Integer
    
    If bVB_CNShwThai Then
         nPos = 1
    Else
         nPos = 2
    End If
    UT_nPosCutTag = nPos
End Function

Public Function UT_nRoundDiff(ByVal ptDiff$, Optional nCutOut As Byte) As Double
'----------------------------------------------------------
'   Call:   ptDiff as last 2 digits of currency format ".99"
'
'   Cmt:   adjustment 2 digits of double
'----------------------------------------------------------
On Error GoTo ErrHandel
    Dim nDif As Double
    Dim nI As Long
    Dim bPass As Boolean
    
    Dim cMeaningMin As Double
    Dim cMeaningDef As Double
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    tSql = "SELECT FNMenMin,FNMenValue FROM TSysMeaning"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        cMeaningMin = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenMin"))) / 100)
        cMeaningDef = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FNMenValue"))) / 100)
    Else
        cMeaningMin = 0
        cMeaningDef = 0
    End If
    Call SP_DATxRsNothing(orsTemp)
    
    Select Case nCutOut
    Case eRndUp
        bPass = False
        For nI = LBound(aVB_CNStangRndAlw) To UBound(aVB_CNStangRndAlw) '- 1 '*KT 51-10-16  Communication Sheet ML-SF 578 Run Time Error on POST.xls
            If aVB_CNStangRndAlw(nI) > CDbl(ptDiff) Then
                nDif = aVB_CNStangRndAlw(nI) - CDbl(ptDiff)
                bPass = True
                Exit For
            End If
        Next nI
        If Not bPass Then nDif = 1 - CDbl(ptDiff)
        
    Case eRndDn
        bPass = False
        For nI = LBound(aVB_CNStangRndAlw) To UBound(aVB_CNStangRndAlw) '- 1 '*KT 51-10-16  Communication Sheet ML-SF 578 Run Time Error on POST.xls
            If aVB_CNStangRndAlw(nI) > CDbl(ptDiff) Then
                If nI = 0 Then
                    nDif = -1 * CDbl(ptDiff)
                Else
                    nDif = -1 * (CDbl(ptDiff) - aVB_CNStangRndAlw(nI - 1))
                End If
                bPass = True
                Exit For
            End If
        Next nI
        If Not bPass Then nDif = -1 * (CDbl(ptDiff) - aVB_CNStangRndAlw(nI - 1))
        
    Case eRndMeaning
        bPass = False
        For nI = LBound(aVB_CNStangRndAlw) To UBound(aVB_CNStangRndAlw) '- 1'*KT 51-10-16  Communication Sheet ML-SF 578 Run Time Error on POST.xls
            If aVB_CNStangRndAlw(nI) > CDbl(ptDiff) Then
                If nI = 0 Then
                    nDif = CDbl(ptDiff)
                Else
                    nDif = (CDbl(ptDiff) - aVB_CNStangRndAlw(nI - 1))
                End If
                bPass = True
                Exit For
            End If
        Next nI
        If Not bPass Then nDif = (CDbl(ptDiff) - aVB_CNStangRndAlw(nI - 1)) '*KT 51-10-16  Communication Sheet ML-SF 578 Run Time Error on POST.xls
        If nDif <= cMeaningDef Then '*KT 51-10-01 เดิมใช้ cMeaningMin
             bPass = False
            For nI = LBound(aVB_CNStangRndAlw) To UBound(aVB_CNStangRndAlw) '- 1 '*KT 51-10-16  Communication Sheet ML-SF 578 Run Time Error on POST.xls
                If aVB_CNStangRndAlw(nI) > CDbl(ptDiff) Then
                    If nI = 0 Then
                        nDif = -1 * CDbl(ptDiff)
                    Else
                        nDif = -1 * (CDbl(ptDiff) - aVB_CNStangRndAlw(nI - 1))
                    End If
                     bPass = True
                    Exit For
                End If
            Next nI
            If Not bPass Then nDif = -1 * (CDbl(ptDiff) - aVB_CNStangRndAlw(nI - 1))
        Else
            nDif = cMeaningMin - CDbl(ptDiff) '*KT 51-10-01cMeaningDef
        End If
    End Select
    UT_nRoundDiff = nDif
    Exit Function
ErrHandel:
    '*KT 51-10-16  Communication Sheet ML-SF 578 Run Time Error on POST.xls
    '----------------------------------------------------------------------
    MsgBox "RoundDiff Error" & vbCrLf & Err.Number & ":" & Err.Description, vbCritical
    UT_nRoundDiff = 0
    '---------------------------------------------------------------------
End Function

Public Sub UT_ClearBrwNRptDefind()
'----------------------------------------------------------
'   Call:
'
'   Ret:    clear varaiable of browser
'---------------------------------------------------------
    tUT_BrwRet = ""           '*GP 45-04-20, effect or not?
    tUT_BrwTable = ""
    tUT_BrwField = ""
    tUT_BrwHead = ""
    tUT_BrwWhere = ""
    tUT_BrwWCaption = ""
End Sub

Public Function UT_bIsDataExist(ByVal ptSql$) As Boolean
'----------------------------------------------------------
'   Call:
'   Cmt:   finding data
'----------------------------------------------------------
    Dim orsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    UT_bIsDataExist = False
            
   If SP_SQLvExecute(ptSql, orsTmp, oDB.ocnOnLine, False, , , , True) = 0 Then
        '*Tee 49-06-09      Check TimeOut
        If (orsTmp.BOF And orsTmp.EOF) Then         'not found any data
            Call SP_MSGnShowing("tms_CN005108", nCS_Warning)
            Call UT_ClearBrwNRptDefind
            Exit Function
        End If
    Else
    
    End If
    UT_bIsDataExist = True
    Call SP_DATxRsNothing(orsTmp)
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function UT_tChkFieldString(ByVal poRs As ADODB.Recordset, ByVal ptField$, Optional ByVal pvRetOld) As String
'----------------------------------------------------------
'   Call:   poRs as ADODB.Recordset
'               ptField is field to check
'               pvRetOld is optional returning old value
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
                UT_tChkFieldString = poRs.Fields(ptField)
            End If
        End If
    End If
    If bEmpty Then
        If IsMissing(pvRetOld) Then
            UT_tChkFieldString = ""
        Else
            UT_tChkFieldString = pvRetOld
        End If
    End If
End Function

Public Function UT_vChkFieldNumeric(ByVal poRs As ADODB.Recordset, ByVal ptField$, Optional ByVal pvRetOld) As Variant
'----------------------------------------------------------
'   Call:   poRs as ADODB.Recordset
'               ptField is field to check
'               pvValue is optional return value
'   Ret:    numeric value as variant(integer,long,single,double ...)
'----------------------------------------------------------
    Dim bZero As Boolean
    Dim cOld As Variant
    
    If poRs.BOF And poRs.EOF Then
        bZero = True
    Else
        If IsNull(poRs.Fields(ptField)) Then
            bZero = True
        Else
            UT_vChkFieldNumeric = IIf(Mid(ptField, 2, 1) = "C", poRs.Fields(ptField), poRs.Fields(ptField))
        End If
    End If
    If bZero Then
        If IsMissing(pvRetOld) Then
            UT_vChkFieldNumeric = 0 ' IIf(Mid(ptField, 2, 1) = "C", 0, 0)
        Else
            cOld = pvRetOld     'หลีกเลี่ยง parameter byRef
            UT_vChkFieldNumeric = cOld 'IIf(Mid(ptField, 2, 1) = "C", cOld, cOld)
        End If
    End If
End Function

Public Sub UT_ChkKeyAscii(ByRef pnKeyAscii As Integer, Optional pbIsDTPicker As Boolean)
'-----------------------------------------------------------------------------------------------
'   Call:   pnKeyAscii is any key ascii (by reference)
'               pbIsDTPicker is optional if an object is DTPicker
'
'   Cmt:   place this code in event KeyPress ของ Text box   เพื่อไม่รับอักขระ เขาเดี่ยว('), เขาคู่("), เส้นตั้ง(|)
'              and place in event KeyPress of DTPicker for except vbKeyLeft ,vbKeyRight,vbKeyUp,vbKeyDown
'-----------------------------------------------------------------------------------------------
    If pbIsDTPicker Then
        If pnKeyAscii <> vbKeyLeft Or pnKeyAscii <> vbKeyRight Or pnKeyAscii <> vbKeyUp Or pnKeyAscii <> vbKeyDown Then pnKeyAscii = 0
    Else
        If pnKeyAscii = 34 Or pnKeyAscii = 39 Or pnKeyAscii = 124 Then pnKeyAscii = 0       ' [ ', ", | ]
    End If
End Sub

Public Sub UT_OPTxGetColor()
'----------------------------------------------------------
'   Call:
'   Cmt:  set first varaiable
'----------------------------------------------------------
    Dim orsTemp As ADODB.Recordset
    Dim tSql$
    On Error GoTo ErrDB
    tSql = "SELECT FTScoDefValue,FTScoUseDef,FTScoUsrValue ,FTScoObj "
    tSql = tSql & " FROM TSysObjColor"
    tSql = tSql & " WHERE  FTScoForm ='wOption'"
    tSql = tSql & " ORDER BY FNScoSeq"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        Do While Not orsTemp.EOF
            Select Case SP_FEDtChkString(orsTemp, "FTScoObj")
                Case "nVB_CNClrBLock"
                    nVB_CNClrBLock = IIf(SP_FEDtChkString(orsTemp, "FTScoUseDef") = "0", SP_FEDtChkString(orsTemp, "FTScoUsrValue"), SP_FEDtChkString(orsTemp, "FTScoDefValue"))
                Case "nVB_CNClrBUnLock"
                    nVB_CNClrBUnLock = IIf(SP_FEDtChkString(orsTemp, "FTScoUseDef") = "0", SP_FEDtChkString(orsTemp, "FTScoUsrValue"), SP_FEDtChkString(orsTemp, "FTScoDefValue"))
                Case "nVB_CNClrBOlaTag4"
                    nVB_CNClrBOlaTag4 = IIf(SP_FEDtChkString(orsTemp, "FTScoUseDef") = "0", SP_FEDtChkString(orsTemp, "FTScoUsrValue"), SP_FEDtChkString(orsTemp, "FTScoDefValue"))
                Case "nVB_CNClrBOlaTag5"
                    nVB_CNClrBOlaTag5 = IIf(SP_FEDtChkString(orsTemp, "FTScoUseDef") = "0", SP_FEDtChkString(orsTemp, "FTScoUsrValue"), SP_FEDtChkString(orsTemp, "FTScoDefValue"))
                Case "nVB_CNClrBOtbCal"
                    nVB_CNClrBOtbCal = IIf(SP_FEDtChkString(orsTemp, "FTScoUseDef") = "0", SP_FEDtChkString(orsTemp, "FTScoUsrValue"), SP_FEDtChkString(orsTemp, "FTScoDefValue"))
                Case "nVB_CNClrBFrmMain"
                    nVB_CNClrBFrmMain = IIf(SP_FEDtChkString(orsTemp, "FTScoUseDef") = "0", SP_FEDtChkString(orsTemp, "FTScoUsrValue"), SP_FEDtChkString(orsTemp, "FTScoDefValue"))
                Case Else       '*GPxxx no good (temporary)
                    nUT_ColorFGrid = vbRed
                    nUT_ColorBGrid = vbGreen
                    nUT_ColorOopGotFocus = vbBlue
                    nUT_ColorBOlaPrice = &HC000&               '&H0000C000&
            End Select
            orsTemp.MoveNext
        Loop
    Else
        nVB_CNClrBLock = &HC0E0FF
        nVB_CNClrBUnLock = vbWhite
        nVB_CNClrBOlaTag4 = &HC0E0FF
        nVB_CNClrBOlaTag5 = &H80FF80
        nVB_CNClrBOtbCal = &HF1B16B
        nVB_CNClrBFrmMain = &H808000
    End If
    Exit Sub
ErrDB:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Public Function UT_tAPIGetDriveSN(ByVal ptDrive As String) As String
'----------------------------------------------------------
'   Call:   ptDrive as required drive
'
'   Ret:    serial number of drive
'----------------------------------------------------------
    Dim nNo As Long
    Dim tLen As String * nCS_MaxFileLen
    Dim tAdj$
    Dim tRet$
    On Error GoTo ErrHandle:
    Call GetVolumeInformation(ptDrive + ":\", tLen, nCS_MaxFileLen, nNo, 0&, 0&, tLen, nCS_MaxFileLen)
    
    tRet = Mid(Hex(nNo), 1, 4) & Mid(Hex(nNo), 5, 4)
    If Len(tRet) <> 8 Then              'if start with zero     0XXX-XXXX
        tAdj = String(8 - Len(tRet), "0") & tRet
        tRet = Mid(tAdj, 1, 4) & Mid(tAdj, 5, 4)
    End If
    UT_tAPIGetDriveSN = tRet
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function UT_tSystemDirectory() As String
'----------------------------------------------------------
'   Call:
'
'   Ret:    system path such as C:\Windows\System
'----------------------------------------------------------
    Dim tWinPath As String
    Dim nTemp As Long
    
    tWinPath = String(145, Chr(0))
    nTemp = GetSystemDirectory(tWinPath, 145)
    UT_tSystemDirectory = Left(tWinPath, InStr(tWinPath, Chr(0)) - 1)
End Function

Public Function UT_nPwdKeyCode(ptKey As String) As Integer
'----------------------------------------------------------
'   Call:   ptKey is the key used for encrypting/decrypting the string
'
'   Ret:   numeric key

'   Cmt:  decrypt string with key word,
'              use with: UT_tPwdDecrypt +UT_tPwdEncrypt+UT_nPwdKeyCode
'
'              This function creates an integer value based on the ptKey provided.
'              The principle is simple.  The result is the absolute value of the difference between the averages of
'               the odd and even characters.
'----------------------------------------------------------
    Dim aTotal(1 To 2) As Integer
    Dim aNbChars(1 To 2) As Integer
    Dim i%

    aTotal(1) = 0
    aTotal(2) = 0
    aNbChars(1) = 0
    aNbChars(2) = 0

    For i = 1 To Len(ptKey)
        Select Case i Mod 2
            Case 0
                aTotal(1) = aTotal(1) + Asc(Mid(ptKey, i, 1))       'characters in an even position
                aNbChars(1) = aNbChars(1) + 1

            Case Else
                aTotal(2) = aTotal(2) + Asc(Mid(ptKey, i, 1))       'characters in an odd position
                aNbChars(2) = aNbChars(2) + 1

        End Select
    Next i

    'a division by zero must be avoided.
    If aNbChars(1) > 0 And aNbChars(2) > 0 Then
        'this will be the new value used for encryption
        UT_nPwdKeyCode = Abs((aTotal(1) / aNbChars(1)) - (aTotal(2) / aNbChars(2)))
    Else        'if the ptKey is less than 2 characters long, the code
        'becomes 1
        UT_nPwdKeyCode = 1
    End If
End Function

Public Function UT_tPwdEncrypt(ByVal ptOrigStr As String, ByVal ptKey As String) As String
'-----------------------------------------------------------
'   Call:  ptOrigStr is the original string value before encryption, such as serial no of HD
'              ptKey is the key used for encrypting/decrypting the string
'
'   Ret:   string encrypted
'
'   Cmt:  encrypt string with key word,
'              use with: UT_tPwdDecrypt +UT_tPwdEncrypt+UT_nPwdKeyCode
'-----------------------------------------------------------
    Dim i%, nKey%
    
    'first thing done is a calculation upon the encryption ptKey
    'to determine how the original string will be encrypted.
    nKey = UT_nPwdKeyCode(ptKey)
    UT_tPwdEncrypt = ""
    
    'now the string will be changed according to the new encryption values
    For i = 1 To Len(ptOrigStr)
        Select Case i Mod 2
            Case 0
                'characters in an even position
                UT_tPwdEncrypt = UT_tPwdEncrypt + Chr(Asc(Mid(ptOrigStr, i, 1)) + nKey)
                
            Case Else
                'characters in an odd position
                UT_tPwdEncrypt = UT_tPwdEncrypt + Chr(Asc(Mid(ptOrigStr, i, 1)) - nKey)
        End Select
    Next i
End Function

Public Function UT_tPwdDecrypt(ByVal ptEncSrc As String, ByVal ptKey As String) As String
'----------------------------------------------------------
'   Call:   ptOrigStr is  The original string value before encryption.,  ptKey   is The key used for encrypting/decrypting the string
'
'   Cmt:   decrypt string with key word,   use with: UT_tPwdDecrypt +UT_tPwdEncrypt+UT_nPwdKeyCode
'
'   Ret:    string decrypt
'----------------------------------------------------------
    Dim i%, nKey%
    
    'first thing done is a calculation upon the encryption key
    'to determine how the original string will be encrypted.
    nKey = UT_nPwdKeyCode(ptKey)
    UT_tPwdDecrypt = ""
    
    'now the encrypted string will be deciphered according to the encryption ptKey
    For i = 1 To Len(ptEncSrc)
        Select Case i Mod 2
            Case 0
                'characters in an even position
                UT_tPwdDecrypt = UT_tPwdDecrypt + Chr(Asc(Mid(ptEncSrc, i, 1)) - nKey)
                
            Case Else
                'characters in an odd position
                UT_tPwdDecrypt = UT_tPwdDecrypt + Chr(Asc(Mid(ptEncSrc, i, 1)) + nKey)
                
        End Select
    Next i
End Function

Public Function UT_tPwdReadDLL(ByVal ptFile$) As String
'----------------------------------------------------------
'   Call:   ptFile as specific file
'
'   Ret:
'----------------------------------------------------------
    Dim tLine$, tFile$
    Dim tFst$, tSnd$
    Dim tNextFst$, tNextSnd$
    Dim i%, k%
    
    On Error GoTo ErrHandle
    
    UT_tPwdReadDLL = ""
    
    k = 5
    
    tFile = ptFile
    Open tFile For Input As #1              'open file
    Do While Not EOF(1)
        Line Input #1, tLine                        'read line by line, into tLine
        If Trim(tLine) <> "" Then
            i = i + 1
            Select Case i
                Case 1
                    tFst = Mid(tLine, 2, 1)
                    tSnd = Mid(tLine, 4, 1)
                    tNextFst = Mid(tLine, 2 + k, 1)
                    tNextSnd = Mid(tLine, 4 + k, 1)
                    
                Case 2
                    tFst = tFst & Mid(tLine, 3, 1)
                    tSnd = tSnd & Mid(tLine, 1, 1)
                    tNextFst = tNextFst & Mid(tLine, 3 + k, 1)
                    tNextSnd = tNextSnd & Mid(tLine, 1 + k, 1)
                
                Case 3
                    tFst = tFst & Mid(tLine, 1, 1)
                    tSnd = tSnd & Mid(tLine, 4, 1)
                    tNextFst = tNextFst & Mid(tLine, 1 + k, 1)
                    tNextSnd = tNextSnd & Mid(tLine, 4 + k, 1)
                
                Case 4
                    tFst = tFst & Mid(tLine, 3, 1)
                    tSnd = tSnd & Mid(tLine, 2, 1)
                    tNextFst = tNextFst & Mid(tLine, 3 + k, 1)
                    tNextSnd = tNextSnd & Mid(tLine, 2 + k, 1)
                    Exit Do
                
            End Select
        End If
    Loop
    Close #1                                               'close file
    
    If (tFst & tSnd) = (tNextFst & tNextSnd) Then
        UT_tPwdReadDLL = tFst & tSnd
    End If
    
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function UT_bPwdWriteDLL(ByVal ptFile$, ByVal ptEncSn$) As Boolean
'----------------------------------------------------------
'   Call:   ptFile as specific file
'
'   Ret:
'----------------------------------------------------------
    Const nCSReading = 1, nCSWriting = 2
    Dim oFSO As Object
    Dim oFile As Object
    Dim aLine() As String
    Dim i As Integer
    Dim tEncSn$
    Dim k%
    Dim tLine$, tFile$
    Dim tNewLn As String
    Dim nLine As Integer
    
    On Error GoTo ErrHandle
    
    UT_bPwdWriteDLL = False
    tEncSn = ptEncSn            'keep parameter
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    'get old value from
    Set oFile = oFSO.OpenTextFile(ptFile, nCSReading)
    Do While oFile.AtEndOfLine <> True
        nLine = nLine + 1
        ReDim Preserve aLine(nLine)
        aLine(nLine) = oFile.ReadLine
    Loop
    oFile.Close
    
    For i = 1 To nLine
          tLine = aLine(i)
          Select Case i
              Case 1
                  tNewLn = Mid(tLine, 1, 1) & Mid(tEncSn, 1, 1) & Mid(tLine, 3, 1) & Mid(tEncSn, 5, 1)
                  tNewLn = tNewLn & Mid(tLine, 5, 2) & Mid(tEncSn, 1, 1) & Mid(tLine, 8, 1) & Mid(tEncSn, 5, 1) & Right(tLine, Len(tLine) - 9)
                  aLine(i) = tNewLn
                  
              Case 2
                  tNewLn = Mid(tEncSn, 6, 1) & Mid(tLine, 2, 1) & Mid(tEncSn, 2, 1) & Mid(tLine, 4, 2)
                  tNewLn = tNewLn & Mid(tEncSn, 6, 1) & Mid(tLine, 7, 1) & Mid(tEncSn, 2, 1) & Right(tLine, Len(tLine) - 8)
                  aLine(i) = tNewLn
                  
              Case 3
                  tNewLn = Mid(tEncSn, 3, 1) & Mid(tLine, 2, 2) & Mid(tEncSn, 7, 1) & Mid(tLine, 5, 1)
                  tNewLn = tNewLn & Mid(tEncSn, 3, 1) & Mid(tLine, 7, 2) & Mid(tEncSn, 7, 1) & Right(tLine, Len(tLine) - 9)
                  aLine(i) = tNewLn
                  
              Case 4
                  tNewLn = Mid(tLine, 1, 1) & Mid(tEncSn, 8, 1) & Mid(tEncSn, 4, 1)
                  tNewLn = tNewLn & Mid(tLine, 3, 3) & Mid(tEncSn, 8, 1) & Mid(tEncSn, 4, 1) & Right(tLine, Len(tLine) - 8)
                  aLine(i) = tNewLn
          End Select
    Next i
  
    Set oFile = oFSO.OpenTextFile(ptFile, nCSWriting, True)
    For i = 1 To nLine
        oFile.WriteLine aLine(i)
    Next i
    oFile.Close
    
    UT_bPwdWriteDLL = True
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function UT_bPwdAllow2Use(ByVal ptFile$, ByVal ptEnSn$) As Boolean
'----------------------------------------------------------
'   Call:   ptFile as specific file
'               ptEnSn is encript serial no of drive C
'
'   Ret:    if can not get serial HD, check by dll file
'----------------------------------------------------------
    Dim tGet$
    
    UT_bPwdAllow2Use = False
        
    tGet = UT_tPwdReadDLL(ptFile)
        
    If tGet = ptEnSn Then
        UT_bPwdAllow2Use = True
    Else
        If tGet = tCS_PwdEncHD Then
            If UT_bPwdWriteDLL(ptFile, ptEnSn) Then
                UT_bPwdAllow2Use = True
            End If
        End If
    End If
End Function

Public Sub UT_PwdAuthorized()
'----------------------------------------------------------
'   Call:
'
'   Ret:    True if has permission to use program
'               False if do not have any permission to use program
'----------------------------------------------------------
    Dim tSn$
    Dim tSnEnc$
    Dim tSysFile$

    tSn = UT_tAPIGetDriveSN("C")                'get serial from drive C:\
    tSnEnc = UT_tPwdEncrypt(tSn, tCS_PwdEncHDKey)       'get encript serial
    
    tSysFile = UT_tSystemDirectory & "\Avidtea.dll"
    If SP_FLEbExist(tSysFile, False) Then
        If UT_bPwdAllow2Use(tSysFile, tSnEnc) Then
            Exit Sub
        End If
    End If
    
    'others user can not use this program
    Call SP_MSGnShowing("tms_CN005028", nCS_Error)
    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Open program duplicate" & " Form:mUT" & " Control:General" & " Events:UT_PwdAuthorized")
    Call SP_APPxEndProgram      'end program *GP 46-09-23
           
End Sub

Public Function UT_GetnRoundDiff(ByVal pcGrand As Double) As Double
'----------------------------------------------------------
'   Call: Comsheet ML-SF-172/ML-SF-173
'   Cmt: calculate % for Customer or cust group        *GP 45-12-06
'----------------------------------------------------------
    Dim tRight$, tSql$, tTotal$
    Dim cDif As Double
    Dim nDec As Double
    Dim nDot%
    Dim nI As Long
    Dim tCase$
    Dim cTmp As Double
    Dim bPrc As Boolean
    Dim aRnd() As String  '*KT 51-10-17 V 3.0.137/V 3.0.200
    Dim bPass As Boolean  '*KT 51-10-17 V 3.0.137/V 3.0.200
    Dim bChk As Boolean     '*Ao 2011/02/23
    Dim tTemp As String         '*Ao 2011/02/23
    
    bChk = False        '*Ao 2011/02/23
    
    If aVB_CNStangRndAlw(0) = 0 Then
        UT_GetnRoundDiff = 0
        Exit Function
    End If
    tCase = ""
    For nI = LBound(aVB_CNStangRndAlw) To UBound(aVB_CNStangRndAlw) ' - 1'*KT 51-10-16  Communication Sheet ML-SF 578 Run Time Error on POST.xls
        If IsNumeric(aVB_CNStangRndAlw(nI)) Then
            tCase = tCase & CStr(aVB_CNStangRndAlw(nI)) & ","
        End If
    Next nI
    tCase = "0," & tCase
    tCase = Mid(tCase, 1, Len(tCase) - 1) 'Cut ','ตัวสุดท้าย
    
    UT_GetnRoundDiff = 0
    tTotal = SP_FMTcCurPoint(pcGrand, , False)
    
    'if diff of total not exactly number
    nDot = InStr(tTotal, ".")
    tRight = Mid(tTotal, nDot, Len(tTotal) - nDot + 1)    'จากทศนิยมไป n ตำแหน่ง
    If CDbl(Val(Format(tRight))) = 0 Then Exit Function '*KT 51-05-2008 Comsheet ML-SF-286
    
    Select Case nVB_CNTypeRnd
        Case eRndUp
            bPrc = True
            Do While CDbl(tRight) > 0 And bPrc
                bPrc = False
                For nI = LBound(aVB_CNStangRndAlw) To UBound(aVB_CNStangRndAlw) '- 1'*KT 51-10-16  Communication Sheet ML-SF 578 Run Time Error on POST.xls
                    If CDbl(tRight) > aVB_CNStangRndAlw(nI) Then
                        tRight = CDbl(tRight) - aVB_CNStangRndAlw(nI)
                        bPrc = True
                    End If
                Next nI
            Loop
'            Select Case CDbl(tRight)                           '*Ao 2011/02/23 Error Type mismatch
'                Case tCase
'                    cDif = 0       'or do nothing
'                Case Else
'                    If bVB_CNAlwRnd Then                       'system roucDiff or not
'                        cDif = UT_nRoundDiff(tRight, eRndUp)
'                    Else
'                        cDif = 0
'                    End If
'            End Select
            If bVB_CNAlwRnd Then                                                                    '*Ao 2011/02/23
                tTemp = cUT.UT_STRtToken(tCase, ",", 1, False)
                Do While tTemp <> ""
                    If CDbl(tTemp) = CDbl(tRight) Then
                        bChk = True
                    End If
                    tTemp = cUT.UT_STRtToken(tCase, ",", 1, False)
                Loop
                If bChk = False Then
                    cDif = UT_nRoundDiff(tRight, eRndUp)
                Else
                    cDif = 0
                End If
            Else
                cDif = 0
            End If
            
        Case eRndDn
            bPrc = True
            Do While CDbl(tRight) > 0 And bPrc
                bPrc = False
                For nI = LBound(aVB_CNStangRndAlw) To UBound(aVB_CNStangRndAlw) '- 1'*KT 51-10-16  Communication Sheet ML-SF 578 Run Time Error on POST.xls
                    If CDbl(tRight) > aVB_CNStangRndAlw(nI) Then
                        tRight = CDbl(tRight) - aVB_CNStangRndAlw(nI)
                        bPrc = True
                    End If
                Next nI
            Loop
        
'            Select Case CDbl(tRight)
'                Case tCase
'                    cDif = 0       'or do nothing
'                Case Else
'                    If bVB_CNAlwRnd Then                       'system roucDiff or not
'                        cDif = UT_nRoundDiff(tRight, eRndDn)
'                    Else
'                        cDif = 0
'                    End If
'            End Select
            If bVB_CNAlwRnd Then                                                                    '*Ao 2011/02/23
                tTemp = cUT.UT_STRtToken(tCase, ",", 1, False)
                Do While tTemp <> ""
                    If CDbl(tTemp) = CDbl(tRight) Then
                        bChk = True
                    End If
                    tTemp = cUT.UT_STRtToken(tCase, ",", 1, False)
                Loop
                If bChk = False Then
                    cDif = UT_nRoundDiff(tRight, eRndDn)
                Else
                    cDif = 0
                End If
            Else
                cDif = 0
            End If
            
        Case eRndMeaning
            bPrc = True
            Do While CDbl(tRight) > 0 And bPrc
                bPrc = False
                For nI = LBound(aVB_CNStangRndAlw) To UBound(aVB_CNStangRndAlw) '- 1'*KT 51-10-16  Communication Sheet ML-SF 578 Run Time Error on POST.xls
                    If CDbl(tRight) > aVB_CNStangRndAlw(nI) Then
                        tRight = CDbl(tRight) - aVB_CNStangRndAlw(nI)
                        bPrc = True
                    End If
                Next nI
            Loop
        
'            Select Case CDbl(tRight)
'                Case tCase
'                    cDif = 0       'or do nothing
'                Case Else
'                    If bVB_CNAlwRnd Then                       'system roucDiff or not
'                        cDif = UT_nRoundDiff(tRight, eRndMeaning)
'                    Else
'                        cDif = 0
'                    End If
'            End Select
 '*KT 51-10-17 V 3.0.137/V 3.0.200
 '-------------------------------------------------------------------------------
             ReDim aRnd(0) As String
             aRnd(0) = ""
             
             bPass = True
             aRnd = Split(tCase, ",")
             If aRnd(0) <> "" Then
                For nI = LBound(aRnd) To UBound(aRnd)
                    If CDbl(tRight) = CDbl(aRnd(nI)) Then
                        cDif = 0
                        bPass = False
                        Exit For
                    End If
                Next nI
                If bPass Then
                    If bVB_CNAlwRnd Then                       'system roucDiff or not
                        cDif = UT_nRoundDiff(tRight, eRndMeaning)
                    Else
                        cDif = 0
                    End If
                End If
            Else
                cDif = 0
            End If
 '------------------------------------------------------------------------------------
 
        Case eRndDon
            cDif = 0
    End Select
    UT_GetnRoundDiff = cDif
End Function




