Attribute VB_Name = "mMain"
Option Explicit
Public Const tMS_001 = "เนื่องขณะนี้ โปรแกรมถ่ายโอนข้อมูลจากการขายแบบ Offline ไปยัง Online ทำงานไม่สมบูรณ์" & vbCrLf & "อาจมีผลทำให้ยอดขายที่ขายด้วย Offline ไม่ถูกส่งไปยัง Server. กรุณาติดต่อผู้ดูแลระบบ;Due to uncompleting program while running ServiceTransfer.May be some offline transaction can not deliver to server." & vbCrLf & "Please contact your administrator."
Public ocnOnline As ADODB.Connection
Public ocnOffline As ADODB.Connection
Public tMM_CNDbOffServer As String
Public tMM_CNDbOffUser As String
Public tMM_CNDbOffPassword As String
Public tMM_CNDbOffDB As String
Public tMM_CNDbOnServer As String
Public tMM_CNDbOnUser As String
Public tMM_CNDbOnPassword As String
Public tMM_CNDbOnDB As String
Public tVB_CNDbAdaIni As String
Public cUT As New AdaUT.cUT

Public tMM_CNDbMUser As String                  '*Ao 2013/03/18  TK
Public tMM_CNDbMPassword As String
Public tMM_CNDbMServer As String
Public tMM_CNDbMDB As String
'file
Public Enum eEN_FleManage
    nEN_Fle1Exist = 1               'check if exist
    nEN_Fle2Delete = 2           'delete file
End Enum

'folder
Public Enum eEN_FldManage
    nEN_Fld1Exist = 1               'check if exist
    nEN_Fld2Delete = 2           'delete folder
    nEN_Fld3Create = 3           'create folder
End Enum

Public tVB_CNTmnNum$
Public tVB_CNPthFileLog$
Public tVB_FileLog$
Public oVB_LogFile As TextStream

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
    
    tEN_TrnSaleManualTax = 43
    tEN_TrnReturnManualTax = 44
    tEN_TrnTotalReportManualTax = 45
    tEN_TrnOverShortReportManualTax = 46
    tEN_TrnClosingReportManualTax = 47
   
End Enum

Public Enum eEN_DBMS
    nEN_Dbs1ACS = 1
    nEN_Dbs2SQL = 2
End Enum

Public nVB_CNBase As eEN_DBMS                 'Access or SQL Sever

'*Bmp 61-02-15 Token
Public oToken As Object 'CreateObject("MSSOAP.SoapClient30")
Public bVB_AlwToken As Boolean
Public bVB_AbleToken As Boolean
Public tVB_TokenURL As String 'http://192.11.15.248:8080/axis2/services/SafeNetTokenizer?wsdl
Public tVB_TokenNAEUsr As String 'nae_user
Public tVB_TokenNAEPwd As String 'P@ssw0rd
Public tVB_TokenDB As String 'dbToken
Public tVB_TokenDBUsr As String 'tkuser
Public tVB_TokenDBPwd As String '!@#Token
Public tVB_TokenFormat As String '4: FIRST_SIX_TOKEN
Public tVB_TokenLuhCheck As String '0

Public Const tCS_CNLastFedIns = "FDDateUpd,FTTimeUpd,FTWhoUpd,FDDateIns,FTTimeIns,FTWhoIns)"
Public Const nVB_ChkDigitToken As Integer = 7 '*Tao 61-10-05 จำนวนหลักที่จำเป็นต้อง Token

Public Sub SP_DATxRsNothing(ByRef poRs As ADODB.Recordset)
'-----------------------------------------------------------
'   Cmt:close record set and nothing
'-----------------------------------------------------------
    If Not poRs Is Nothing Then
            If poRs.State = adStateOpen Then poRs.Close          'ไม่ใช้  State เพราะกรณี recordset ยังไม่ได้เปิดจะ error
            Set poRs = Nothing
    End If
End Sub

Public Sub MM_GETxTable(ByRef porsTbl As ADODB.Recordset)
'--------------------------------------
'Call
'Cmt: อ่านข้อมูลตารางที่จะส่งข้อมูลไปยัง Server
'          *KT 52-10-21 ปรับปรุง On Error GoTo ErrHandle/WITH (NOLOCK)
'-------------------------------------
    Dim tSql As String
On Error GoTo ErrHandle
    'ให้เรียงตาม Seq โดยหลักการให้ทำ RC,CD,DT,HD
    tSql = "SELECT [FTSscTable],[FTSscDocType],[FTSscTableDest],[FTSscGrp],[FNSscGrpSeq],[FTSscPrec],[FTSscStaSync],[FTSscFedReplace],[FTSscFedPK],[FTSscFedChkGrp],[FTSscCondScript],[FTSscUpdScript],[FTSscStaActive],[FTSscStaChkType],[FTSscFedDocType],CAST(FTSscDocType AS INT) AS FNSEQNO"
    tSql = tSql & vbCrLf & "FROM [TSysSync] WITH (NOLOCK)"
    tSql = tSql & vbCrLf & "WHERE [FTSscStaActive] = '1'"
    tSql = tSql & vbCrLf & "ORDER BY FNSEQNO"
    Set porsTbl = ocnOnline.Execute(tSql)
    Exit Sub
ErrHandle:

End Sub

Public Function MM_GETxData(ByVal ptTbl As String, ByVal ptCond As String, ByRef porsData As ADODB.Recordset, ByVal ptPK As String) As Boolean
'-----------------------------------------
'Call:
'Cmt:  หาข้อมูลที่ยังไม่ส่งไปที่ Server ตามเงื่อนไขที่ส่งมาให้
'           *KT 52-10-21 ใส่ WITH (NOLOCK) ป้องกันการเกิดการ Lock ของ Record นั้น
'----------------------------------------
    Dim tSql As String
    Dim tOrder As String
    Dim tTemp As String
On Error GoTo Err_Handler
    tOrder = ""
    tTemp = cUT.UT_STRtToken(ptPK, ";", 1, False)
    Do While tTemp <> ""
        tOrder = IIf(tOrder <> "", tOrder & ",", "") & tTemp & " ASC "
        tTemp = cUT.UT_STRtToken(ptPK, ";", 1, False)
    Loop
    
    tSql = "SELECT * "
    tSql = tSql & vbCrLf & " FROM " & ptTbl
    tSql = tSql & vbCrLf & " WITH (NOLOCK)"
    tSql = tSql & vbCrLf & " WHERE " & ptCond
    If tOrder <> "" Then tSql = tSql & vbCrLf & " ORDER BY " & tOrder '*52-09-15 ให้ทำการเรียงข้อมูลเพื่อป้องกันกรณีที่ Index มีปัญหา การ Order By จะช่วยได้
    Set porsData = ocnOffline.Execute(tSql)
    MM_GETxData = True
    Exit Function
Err_Handler:
    MM_GETxData = False
End Function

Public Function MM_GETxDataComplate(ByVal ptTmnNum As String, ByVal ptTrnNo As String, pdDateTran As Date, ByVal ptPosNo As String, ByRef ptSta As String) As Boolean
'-----------------------------------------
'PH 2.0
'Ret: ยังไม่เคยส่ง = True/ส่งไปแล้ว = False  ptSta = สถานะของ FTShdStaDoc
'Cmt: ตรวจสอบข้อมุลว่าเคยส่งหรือไม่จาก Table HD ทั้ง Table Temp และ Sale
'        *KT 52-10-21 Select TPSTSalHD only
'----------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    
On Error GoTo Err_Handler
    MM_GETxDataComplate = False

    tSql = "SELECT FTTmnNum,FTShdTransNo,FTShdStaDoc"
    tSql = tSql & vbCrLf & "FROM " & "TPSTSalHD" 'ให้ดูจาก Table HD Temp เป็นหลักในฝั่งของ Offline
    tSql = tSql & vbCrLf & "WITH (NOLOCK)"
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptTmnNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & ptTrnNo & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(pdDateTran, "YYYY/MM/DD") & "'"
    tSql = tSql & vbCrLf & "AND FTShdStaDoc= '1'"
    tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaPrcGL,'0') = '1'"
    Set orsTemp = ocnOffline.Execute(tSql)
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        MM_GETxDataComplate = True
        ptSta = SP_FEDtChkString(orsTemp, "FTShdStaDoc")
    Else
        tSql = "SELECT FTTmnNum,FTShdTransNo,FTShdStaDoc"
        tSql = tSql & vbCrLf & "FROM " & "TPSTSalHD" 'ให้ดูจาก Table HD Temp เป็นหลักในฝั่งของ Offline
        tSql = tSql & vbCrLf & "WITH (NOLOCK)"
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptTmnNum & "'"
        tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & ptTrnNo & "'"
        tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(pdDateTran, "YYYY/MM/DD") & "'"
        tSql = tSql & vbCrLf & "AND FTShdStaDoc= '1'"
        tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaPrcGL,'0') = '1'"
        Set orsTemp = ocnOnline.Execute(tSql)
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            MM_GETxDataComplate = True
            ptSta = SP_FEDtChkString(orsTemp, "FTShdStaDoc")
        Else
            ptSta = "2"
        End If
    End If

    If Not (orsTemp Is Nothing) Then
        If orsTemp.State = adStateOpen Then orsTemp.Close
        Set orsTemp = Nothing
    End If
    Exit Function
Err_Handler:
    If Not (orsTemp Is Nothing) Then
        If orsTemp.State = adStateOpen Then orsTemp.Close
        Set orsTemp = Nothing
    End If
    ptSta = "2"
End Function

Public Function MM_FLEbManFile(ByVal ptPath As String, ByVal pnType As eEN_FleManage) As Boolean
'----------------------------------------------------------
'   Call:   ptPath is folder name
'               pnType  is enum of folder type 1:exist, 2:delete
'   Ret :  true if completed
'----------------------------------------------------------
    Dim oFSO As Object
    On Error GoTo ErrHandle
    MM_FLEbManFile = False
    If ptPath = "" Then Exit Function
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Select Case pnType
        Case nEN_Fle1Exist
            If oFSO.FileExists(ptPath) Then
                MM_FLEbManFile = True
            End If
        Case nEN_Fle2Delete
            If oFSO.FileExists(ptPath) Then
                oFSO.DeleteFile (ptPath)
                MM_FLEbManFile = True
            End If
    End Select
    Exit Function
ErrHandle:
    MM_FLEbManFile = False
End Function



Public Function MM_MSGnShowing(ByVal ptMSG$, ByVal ptStyle$, Optional ByVal pbNoCut As Boolean, Optional ByVal ptCaption$, Optional ByVal ptNextLine$) As Byte
'----------------------------------------------------------
'   Call:   ptMSG is showing message
'               ptStyle is style of button
'               pbNoCut is optional, True if do not want to cut message
'               ptCaption is optional, if want to add on caption
'               ptNextLine is optional, if want to add on message (with vbCrLf)
'   Ret:    no. of pressed button
'----------------------------------------------------------
    Dim tMsg$, tMsg2$
    Dim tAppTitle$
    Dim nPos As Integer
    Dim nCharPos As Integer
    Dim tTemp As String
    
    nPos = 1
    ptMSG = IIf(tTemp = "", ptMSG, tTemp)
    Screen.MousePointer = vbDefault
    tAppTitle = App.Title
    If Trim(ptCaption) <> "" Then
        tAppTitle = tAppTitle & " : " & ptCaption
    End If
    tMsg2 = ptMSG
    If Not pbNoCut Then     'if false(default) parameter tMsg will be cut (Thai/Eng)
        tMsg2 = cUT.UT_STRtToken(tMsg2, ";", nPos)
    End If
    tMsg = tMsg2
    'add next line
    If Trim(ptNextLine) <> "" Then
        tMsg = tMsg & vbCrLf & ptNextLine
    End If
    Beep
    nCharPos = InStr(1, UCase(tMsg), UCase(" vbCr "))
    If nCharPos > 0 Then
        tMsg = Trim(Left(tMsg, nCharPos - 1)) & vbCr & Trim(Mid(tMsg, nCharPos + 6))
    End If
    MM_MSGnShowing = MsgBox(tMsg, ptStyle, tAppTitle)
End Function


Public Function MM_RECbAdaIniRead() As Boolean
'----------------------------------------------------------
'----------------------------------------------------------
    Dim tFile$, tSql$, tCon$
    Dim orsTemp As ADODB.Recordset
    Dim bOffLine As Boolean     '*GP 47-02-17
    Dim tSQLAuthSrvType As String
    Dim tCfgCode As String, tPath As String, tMenu As String
    Dim oFSO As New FileSystemObject
    Dim ocnCfg As ADODB.Connection
    
    On Error GoTo ErrHandle
    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\ISSTools\AdaIni.Ada"
    If Not MM_FLEbManFile(tFile, nEN_Fle1Exist) Then
            tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
            tFile = tFile & "\ISSTools\AdaIni.Ada"
    End If
    If Not MM_FLEbManFile(tFile, nEN_Fle1Exist) Then Exit Function
    Set ocnCfg = New ADODB.Connection
    tCon = "Provider=Microsoft.Jet.OLEDB.4.0;"
    tCon = tCon + "Data Source=" & tFile & ";Jet OLEDB:Database Password="
    ocnCfg.CursorLocation = adUseClient
    ocnCfg.ConnectionTimeout = 15      ' 15 วินาที
    ocnCfg.CommandTimeout = 15
    ocnCfg.Open tCon
    If ocnCfg.State <> adStateOpen Then Exit Function       'error connect, exit
    
    tSql = "SELECT *"
    tSql = tSql & " FROM TConfig "
    tSql = tSql & " ORDER BY FTCfgCode ASC"
    Set orsTemp = ocnCfg.Execute(tSql)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        Do While Not orsTemp.EOF
            tCfgCode = SP_FEDtChkString(orsTemp, "FTCfgCode")
            Select Case tCfgCode
                Case "SF_SQLDB1"
                    tMM_CNDbOnUser = SP_FEDtChkString(orsTemp, "FTCfgOn") 'User Name
                    tMM_CNDbOnPassword = SP_FEDtChkString(orsTemp, "FTCfgOff") 'User Password
                    tMM_CNDbOnServer = SP_FEDtChkString(orsTemp, "FTCfgValue") 'Server Name
                    tMM_CNDbOnDB = SP_FEDtChkString(orsTemp, "FTCfgDescTha") 'DataBase Name
                Case "SF_SQLDB2"
                    tMM_CNDbOffUser = SP_FEDtChkString(orsTemp, "FTCfgOn") 'User Name
                    tMM_CNDbOffPassword = SP_FEDtChkString(orsTemp, "FTCfgOff") 'User Password
                    tMM_CNDbOffServer = SP_FEDtChkString(orsTemp, "FTCfgValue") 'Server Name
                    tMM_CNDbOffDB = SP_FEDtChkString(orsTemp, "FTCfgDescTha") 'DataBase Name
                Case "SM_SQLDB1"
                    tMM_CNDbMUser = SP_FEDtChkString(orsTemp, "FTCfgOn") 'User Name
                    tMM_CNDbMPassword = SP_FEDtChkString(orsTemp, "FTCfgOff") 'User Password
                    tMM_CNDbMServer = SP_FEDtChkString(orsTemp, "FTCfgValue")  'Server Name
                    tMM_CNDbMDB = SP_FEDtChkString(orsTemp, "FTCfgDescTha")  'DataBase Name
            End Select
            orsTemp.MoveNext
        Loop
    End If
    If Not (orsTemp Is Nothing) Then Set orsTemp = Nothing
    Set ocnCfg = Nothing
    MM_RECbAdaIniRead = True
    Exit Function
ErrHandle:
End Function

Public Function SP_FEDvChkNumeric(ByVal poRs As ADODB.Recordset, ByVal ptField$, Optional pvRetOld) As Variant
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

Public Function SP_FEDtChkString(ByVal poRs As ADODB.Recordset, ByVal ptField$, Optional pvRetOld) As String
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

Public Function SP_FEDdChkDate(ByVal poRs As ADODB.Recordset, ByVal ptField$, Optional pvRetOld) As Date
    Dim bEmpty As Boolean
    If poRs.BOF And poRs.EOF Then
        bEmpty = True
    Else
        If IsNull(poRs.Fields(ptField)) Then
            bEmpty = True
        Else
            SP_FEDdChkDate = poRs.Fields(ptField)
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

Public Function SP_GETdSaleDate() As Date
'------------------------------------------------------
'Ret: ค่าวันที่ Operation Date(Sale Date) ของ System
'Cmt: *KT 50-12-25 Get Operation Date On Database
'          *KT 52-10-21 ปรับปรุง On Error GoTo ErrHandle
'-----------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT FDCdtDate "
    tSql = tSql & " FROM TSysChgDateTime"
    Set orsTemp = ocnOnline.Execute(tSql)
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        SP_GETdSaleDate = SP_FEDdChkDate(orsTemp, "FDCdtDate")
    Else
        SP_GETdSaleDate = Date
    End If
    If Not (orsTemp Is Nothing) Then
        If orsTemp.State = adStateOpen Then Set orsTemp = Nothing
    End If
    Exit Function
ErrHandle:
    SP_GETdSaleDate = Date
    
End Function

Public Function SP_FLEbManFile(ByVal ptPath As String, ByVal pnType As eEN_FleManage) As Boolean
'----------------------------------------------------------
'   Call:   ptPath is folder name
'               pnType  is enum of folder type 1:exist, 2:delete
'   Ret :  true if completed
'----------------------------------------------------------
    Dim oFSO As Object
    
    On Error GoTo ErrHandle
    
    SP_FLEbManFile = False
    If ptPath = "" Then Exit Function
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Select Case pnType
        Case nEN_Fle1Exist
            If oFSO.FileExists(ptPath) Then
                SP_FLEbManFile = True
            End If
        
        Case nEN_Fle2Delete
            If oFSO.FileExists(ptPath) Then
                oFSO.DeleteFile (ptPath)
                SP_FLEbManFile = True
            End If
    End Select
    
    Exit Function
ErrHandle:
    SP_FLEbManFile = False
End Function

Public Function SP_FLDbManFolder(ByVal ptPath As String, ByVal pnType As eEN_FldManage) As Boolean
'----------------------------------------------------------
'   Call:   ptPath is folder name
'               pnType  is enum of folder type 1:exist, 2:create, 3:delete
'   Ret :  true if completed
'----------------------------------------------------------
    Dim oFSO As Object
    Dim tPath$
    Dim nCol%, i%
    
    On Error GoTo ErrHandle
    
    SP_FLDbManFolder = False
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Select Case pnType
        Case nEN_Fld1Exist
            If oFSO.FolderExists(ptPath) Then
                SP_FLDbManFolder = True
            End If
        
        Case nEN_Fld2Delete
            If oFSO.FolderExists(ptPath) Then
                oFSO.DeleteFolder (ptPath)
                SP_FLDbManFolder = True
            End If
        
        Case nEN_Fld3Create
                If Not oFSO.FolderExists(ptPath) Then
                    tPath = ptPath
                    nCol = InStr(1, tPath, "\")
                    i = 0
                    Do While nCol > 0
                        i = i + 1
                        nCol = InStr(nCol + 1, tPath, "\")
                        If nCol = 0 Then
                            tPath = ptPath
                        Else
                            tPath = Left(tPath, nCol - 1)
                        End If
                        If Not oFSO.FolderExists(tPath) Then        'check server
                            If Not (Mid(tPath, 1, 2) = "\\" And i = 2) Then oFSO.CreateFolder (tPath)
                        End If
                        tPath = ptPath
                    Loop
                End If
                SP_FLDbManFolder = True
    End Select
    
    Exit Function
ErrHandle:
    SP_FLDbManFolder = False
End Function

Public Function SP_STRtSiggleCode(ByVal ptStr As String) As String
'----------------------------------------------------
'Call:
'Cmt: ทำการ replace เครื่องหมายเลข Sigle code ให้มี Sigle code ประกบเพิ่มอีก 1 ตัวเพื่อใชในการ insert single code ได้
'---------------------------------------------------
On Error GoTo ErrHandle
    SP_STRtSiggleCode = ptStr
    SP_STRtSiggleCode = Replace(ptStr, "'", "''")
    Exit Function
ErrHandle:
    SP_STRtSiggleCode = ptStr
End Function

'Public Function SP_PRCbTokenMemSrv(ByVal ptFuncName As String, ByVal ptCardNo As String, ByRef ptRefCardNoToken As String) As Boolean
'    Dim tSql As String
'    Dim orsTemp As ADODB.Recordset
'    Dim nEff  As Long
'    Dim tdbCardNo As String
'On Error GoTo ErrHandle
'   '** หา token บน Member ก่อน เพื่อไม่ต้อง token ใหม่
'    SP_PRCbTokenMemSrv = False
'    tSql = "SELECT TOP 1 FTToken FROM [TPSTTokenLst] WITH(NOLOCK)"
'    tSql = tSql & " WHERE FTValue = '" & ptCardNo & "'"
'
'    Set orsTemp = ocnMem.Execute(tSql, nEff)
'    tdbCardNo = ""
'    If Not (orsTemp.BOF And orsTemp.EOF) Then
'            tdbCardNo = SP_FEDtChkString(orsTemp, "FTToken")
'    End If
'    If tdbCardNo <> "" Then
'        ptRefCardNoToken = tdbCardNo
'        SP_PRCbTokenMemSrv = True
'        Exit Function
'    End If
'    Set orsTemp = Nothing
'    Exit Function
'ErrHandle:
'    SP_PRCbTokenMemSrv = False
'End Function
Public Function SP_PRCbToken(ByVal ptFuncName As String, ByVal ptCardNo As String, ByRef ptRefCardNoToken As String, Optional ByVal pbNoSearchDB As Boolean = False) As Boolean '*Tao 61-05-20 เพิ่ม pbNoSearchDB
'*SP_PRCbToken
'* เรียก web service Token บัตรเครดิต
'*
'*  ใช้การ Token Format = 4: FIRST_SIX_TOKEN input = 4556221560415640, output = 455622XXXXXXXXXX
'*
'* ptCardNo = หมายเลขบัตรก่อน token
'* ptRefCardNoToken = หมายเลขบัตรหลังจาก token แล้ว
'*
'* เช่น
'*  ptCardNo =                      4556221560415640
'*  ptRefCardNoToken =  4556227077358692    (6 digits แรกจะเหมือนเดิม)
    Dim bTEST As Boolean
    Dim tSql As String
    Dim tTokenNo As String
    Dim orsTemp As ADODB.Recordset
    Dim tOnline As String
    
On Error GoTo ErrHandle
'    '==========================TEST
'        bTEST = True
'
'        Select Case ptCardNo
'        '==BP Award
'        Case "4556221550000126"
'            ptRefCardNoToken = "4556226864007022" 'token
'            Case "4556221550000127"
'            ptRefCardNoToken = "4556227215149114" 'token
'            Case "4556221550000128"
'            ptRefCardNoToken = "4556226001154237" 'token
'            '==Member
'            Case "377911240000020"
'            ptRefCardNoToken = "377911675160519" 'token
'            Case "377911240000022"
'            ptRefCardNoToken = "377911606087819" 'token
'            Case "377911240000024"
'            ptRefCardNoToken = "377911903886562" 'token
'            '==TC
'            Case "4538970000000011"
'            ptRefCardNoToken = "4538971570693662" 'token
'            Case "4538970000000015"
'            ptRefCardNoToken = "4538975485678806" 'token
'            Case "4538970000000020"
'            ptRefCardNoToken = "4538977095940214" 'token
'            '==Credit
'            Case "377911240000015"
'            ptRefCardNoToken = "377911362679072" 'token
'            Case "377911240000016"
'            ptRefCardNoToken = "377911396897048" 'token
'            Case "377911240000017"
'            ptRefCardNoToken = "377911292850469" 'token
'            '== Hire Purchase
'            Case "377911200000012"
'            ptRefCardNoToken = "377911827572978" 'token
'            Case "377911200000014"
'            ptRefCardNoToken = "377911713640081" 'token
'            Case "377911200000016"
'            ptRefCardNoToken = "377911865201201" 'token
'            '==
'        Case "4556221550000024"
'            ptRefCardNoToken = "4556227077358692" 'token
'        Case "4556221550000002"
'            ptRefCardNoToken = "4556227077358456" 'token
'        Case "4556221550000022"
'            ptRefCardNoToken = "4556227077358985" 'token
'        '''Credit
'        Case "8888400000000000"
'            ptRefCardNoToken = "8888401326803225" 'token
'        Case "377911240000004"
'            ptRefCardNoToken = "377911097558885" 'token
'        Case "377911240000006"
'            ptRefCardNoToken = "377911861837252" 'token
'        '== Hire Purchase
'        Case "377911200000012"
'        Case "377911200000014"
'        Case "377911200000016"
'        Case Else
'            ptRefCardNoToken = ptCardNo
'        End Select
'
'        If bTEST Then
'            SP_PRCbToken = True
'            Exit Function
'        End If
    '==========================TEST
    '=== จำนวน digits ยาวน้อยกว่า 15 digits ไม่ต้อง Token เพราะ Web service จะ Error
    '==== Store Para บอกว่า ไม่ต้อง Token ให้ส่งเลขเดิมกลับไป
    If (Not bVB_AlwToken) Or (Len(ptCardNo) < nVB_ChkDigitToken) Then  '*Tao 61-10-05 ตรวจสอบจำนวนหลัก ก่อน Token
        ptRefCardNoToken = ptCardNo
        SP_PRCbToken = True
        Exit Function
    End If
    
    '*Bmp 61-04-03  issue 3.      Diners credit card number not tokenized  card # of 14 digits no need to check on length
    If (Len(ptCardNo) < 14) Or (Trim(ptCardNo) = "") Then
        ptRefCardNoToken = ptCardNo
        SP_PRCbToken = True
        Exit Function
    End If

'    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "SP_PRCbToken(1)"
    
    tOnline = GetSetting("AdaPosMall", "POSFront", "Online") '*Tao 61-05-09 ใช้เช็ค Service Online ก่อน จะได้ไม่ต้องรอนาน
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_PRCbToken : Online = " & tOnline    '*Tao 61-05-09 เก็บ Log เพิ่ม
    'If (tOnline = "1") Then '*Tao 61-05-21 ของเดิม
    If tOnline = "1" And (Not pbNoSearchDB) Then '*Tao 61-05-21 เพิ่มเงื่อนไขว่า จะค้าหาในฐานข้อมูลก่อนหรือไม่
        tTokenNo = ""
        
        '*Tao 61-05-21 ของเดิม
        'tSql = "SELECT TOP 1 FTToken FROM [TPSTTokenLst] WITH(NOLOCK)"
        'tSql = tSql & " WHERE FTValue = '" & ptCardNo & "'"
        
        '*Tao 61-05-21 หาข้อมูลที่ Value ไม่เท่ากับ Token
        tSql = "SELECT TOP 1 FTToken FROM TPSTTokenLst WITH(NOLOCK)"
        tSql = tSql & vbNewLine & "WHERE (ISNULL(FTValue,'')<>'' OR ISNULL(FTToken,'')<>'') AND (FTValue<>FTToken)"
        tSql = tSql & vbNewLine & "AND (FTValue='" & ptCardNo & "')"
        tSql = tSql & vbNewLine & "ORDER BY FDDateIns DESC"
        
        Set orsTemp = ocnOnline.Execute(tSql)
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            tTokenNo = SP_FEDtChkString(orsTemp, "FTToken")
        End If
    '    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "SP_PRCbToken(2):" & tTokenNo
        If tTokenNo <> "" Then
            ptRefCardNoToken = tTokenNo
            SP_PRCbToken = True
            Exit Function
        End If
    End If
    
'    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "SP_PRCbToken(4)"
    ' ถ้า initial ไม่ผ่าน ให้ initial ใหม่
'    If (bVB_AbleToken = False) Then
'        Call SP_SETxObjects
'    End If

    '*Tao 61-05-17 ปิดไว้
'    Call SP_SETxObjects '*Tao 61-05-09  Create Object  ใหม่ทุกครั้ง
'    '** Web servie ไม่ทำงาน
'    If (bVB_AbleToken = False) Then
'        ptRefCardNoToken = ptCardNo
'        SP_PRCbToken = False
'        oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   ("AbleToken = False")
'        Exit Function
'    End If
    
'    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "SP_PRCbToken(5)"
'    Call oToken.MSSoapInit(tVB_TokenURL, "", "")
'    If Err <> 0 Then
'        '* คิดก่อนว่าจะให้ servicetransfer write log ลงที่ไหน Call SP_CRTxLogMallCardExtend(ptFuncName & "Initial Token Web service Error: " & Err.Number & ":" & Err.Description)
'        'MsgBox "Initial Token Web service Error: " & Err.Number & ":" & Err.Description, vbExclamation
'        SP_PRCbToken = False
'        Exit Function
'    End If
    'Dim cardno As String
    'Dim CardFormat As Integer
    'cardno = txtCardNo.Text
    'CardFormat = CInt(lstCardFormat.Text)
    'webservice method defined
    'Screen.MousePointer = vbHourglass
    'txtGetToken.Text = oToken.InsertToken("nae_user", "Pass1234", "tkuser", "!@#Token", "CREDITCARDNO", cardno, 4, 0)
    If SP_CHKbWebservice Then
        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " InsertToken Start."
        ptRefCardNoToken = oToken.InsertToken(tVB_TokenNAEUsr, tVB_TokenNAEPwd, tVB_TokenDBUsr, tVB_TokenDBPwd, tVB_TokenDB, ptCardNo, tVB_TokenFormat, tVB_TokenLuhCheck)
        'If (Err <> 0) Then '[5.1803.4][61/03/30] *Bmp 61-03-30
        If (Err <> 0) Or (ptRefCardNoToken = "") Then '[5.1803.4][61/03/30] *Bmp 61-03-30
            SP_PRCbToken = False
            oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " InsertToken(1)" & ptFuncName & " not success, Err:" & Err.Number & ",Err Desc:" & Err.Description
            Exit Function
        End If
        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " InsertToken(2)" & ptFuncName & " success, Token Card No:[" & ptRefCardNoToken & "]"
        SP_PRCbToken = True
    Else
        SP_PRCbToken = False
    End If
    Set oToken = Nothing '*Tao 61-05-23
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " InsertToken End."
    Exit Function
ErrHandle:
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " InsertToken(3)" & ptFuncName & " not success, Err:" & Err.Number & ",Err Desc:" & Err.Description
    SP_PRCbToken = False
    Set oToken = Nothing '*Tao 61-05-23
End Function

Public Function SP_PRCbDeToken(ByVal ptTokenCardNo As String, ByRef ptDeTokenCardNo As String, Optional ByVal pbNoSearchDB As Boolean = False) As Boolean '*Tao 61-05-20 เพิ่ม pbNoSearchDB
'*Bmp 61-04-04
'*Detoken
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim nErr  As Long
    Dim tdbCardNo As String
    Dim tOnline As String
    
On Error GoTo ErrHandle
        '* Store Parameter บอกว่า ไม่ต้อง Token
       If (Not bVB_AlwToken) Or (Len(ptTokenCardNo) < nVB_ChkDigitToken) Then  '*Tao 61-10-05 ตรวจสอบจำนวนหลัก ก่อน Token
            ptDeTokenCardNo = ptTokenCardNo
            SP_PRCbDeToken = True
            Exit Function
        End If
        
        tOnline = GetSetting("AdaPosMall", "POSFront", "Online") '*Tao 61-05-09 ใช้เช็ค Service Online ก่อน จะได้ไม่ต้องรอนาน
        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_PRCbDeToken : Online = " & tOnline    '*Tao 61-05-09 เก็บ Log เพิ่ม
        'If (tOnline = "1") Then '*Tao 61-05-21 ของเดิม
        If tOnline = "1" And (Not pbNoSearchDB) Then '*Tao 61-05-21 เพิ่มเงื่อนไขว่า จะค้าหาในฐานข้อมูลก่อนหรือไม่
            '*Bmp 61-04-06 v.[5.1804.4][61/04/06]
            '*หาเลข detoken จาก DB ก่อน
            
            '*Tao 61-05-21 ของเดิม
            'tSql = "SELECT TOP 1 FTValue  FROM [TPSTTokenLst] WITH(NOLOCK)"
            'tSql = tSql & " WHERE FTToken = '" & ptTokenCardNo & "'"
            
            '*Tao 61-05-21 หาข้อมูลที่ Value ไม่เท่ากับ Token
            tSql = "SELECT TOP 1 FTValue FROM TPSTTokenLst WITH(NOLOCK)"
            tSql = tSql & vbNewLine & "WHERE (ISNULL(FTValue,'')<>'' OR ISNULL(FTToken,'')<>'') AND (FTValue<>FTToken)"
            tSql = tSql & vbNewLine & "AND (FTToken='" & ptTokenCardNo & "')"
            tSql = tSql & vbNewLine & "ORDER BY FDDateIns DESC"
            tdbCardNo = ""
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tdbCardNo = SP_FEDtChkString(orsTemp, "FTValue") '*หมายเลขที่ส่งมา มี token แล้ว
                If tdbCardNo <> "" Then
                    ptDeTokenCardNo = tdbCardNo
                    SP_PRCbDeToken = True
                    Exit Function
                End If
            End If
            '============================================================================== '*จบ หาเลข detoken จาก DB ก่อน
            Call SP_DATxRsNothing(orsTemp)
        End If
        
        '* ตรวจสอบ web service
'        If Not bVB_AbleToken Then
'            Call SP_SETxObjects
'        End If
        
        '*Tao 61-05-17 ปิดไว้
'        Call SP_SETxObjects '*Tao 61-05-09  Create Object  ใหม่ทุกครั้ง
'        If Not bVB_AbleToken Then
'            ptDeTokenCardNo = ptTokenCardNo
'            SP_PRCbDeToken = False '= Initial web service ไม่ได้ ไม่ให้ token
'            Exit Function
'        End If
        
        '* ถ้า web service ใช้ไม่ได้(bVB_AbleToken = false)  ไม่ต้อง token
        If SP_CHKbWebservice Then
            'soapClient3.GetValue("nae_user", "Pass1234", "tkuser", "!@#Token", "CREDITCARDNO", token, 0)
            ptDeTokenCardNo = oToken.GetValue(tVB_TokenNAEUsr, tVB_TokenNAEPwd, tVB_TokenDBUsr, tVB_TokenDBPwd, tVB_TokenDB, ptTokenCardNo, tVB_TokenLuhCheck)
            If (Err <> 0) Then
                ptDeTokenCardNo = ptTokenCardNo
                SP_PRCbDeToken = False
                oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_PRCbDeToken GetValue(1) not success, Err:" & Err.Number & ",Err Desc:" & Err.Description
                'Call SP_CRTxLogMallCardExtend("Detoken:" & ptDeTokenCardNo) '*[5.1804.2][61/04/05]*Bmp 61-04-05
            Else
                If ptDeTokenCardNo <> "" Then
                    SP_PRCbDeToken = True
                    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_PRCbDeToken GetValue(1) not success, Err:" & Err.Number & ",Err Desc:" & Err.Description
                End If
            End If
        Else
            ptDeTokenCardNo = ptTokenCardNo
            SP_PRCbDeToken = False
        End If
        Set oToken = Nothing '*Tao 61-05-23
        Exit Function
ErrHandle:
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_PRCbDeToken GetValue(3) not success, Err:" & Err.Number & ",Err Desc:" & Err.Description
    ptDeTokenCardNo = ptTokenCardNo
    SP_PRCbDeToken = False
    Set oToken = Nothing '*Tao 61-05-23
End Function

Public Sub SP_GETxVariableAlw()
'----------------------------------------------------------
'   Call:
'   Cmt:    *KT 50-12-16 เก็บค่าเฟก Option ของระบบ
'----------------------------------------------------------

    Dim tSql$, tTemp$, i As Long
    Dim orsCfg As ADODB.Recordset
    Dim cMinValue As Double '*Eaw 56-08-23
On Error GoTo Err_Handle
    
     tSql = "SELECT FTStpCode,FTStpSelValue,FTStpStaAlw  FROM TCNMStorePara WHERE FTStpCode = '29'"

     Set orsCfg = ocnOnline.Execute(tSql)
     Do While Not orsCfg.EOF
        Select Case SP_FEDtChkString(orsCfg, "FTStpCode")
            Case "29" '*Bmp 61-02-15 Token Yes or No
                bVB_AlwToken = IIf(SP_FEDtChkString(orsCfg, "FTStpStaAlw") = "Y", True, False)
        End Select
        orsCfg.MoveNext
     Loop

     Call SP_DATxRsNothing(orsCfg)
    
    tSql = "SELECT * FROM TSysConfig WHERE FTScfCode ='TKToken'"
    Set orsCfg = ocnOnline.Execute(tSql)
    '    '*Bmp 61-02-15 Token
'    select FTScfCode,FTScfSeq,FTScfNameLang1,FTScfUsrValue
'    From [TSysConfig]
'    where [FTScfCode]='TKToken'
    'TKToken 001 WebServices http://192.11.15.248:8080/axis2/services/SafeNetTokenizer?wsdl
    'TKToken 002 NAEUserID   nae_user
    'TKToken 003 NAEpassword P@ssw0rd
    'TKToken 004 Token DBName    dbToken
    'TKToken 005 Token DBUser    tkuser
    'TKToken 006 Token DBPassword    !@#Token
    Call SP_OPTxGetConfigValue(orsCfg, "TKToken", "001")
    If orsCfg.BOF And orsCfg.EOF Then
        tVB_TokenURL = ""
    Else
        tVB_TokenURL = SP_FEDtChkString(orsCfg, "FTScfUsrValue")
    End If
    
    Call SP_OPTxGetConfigValue(orsCfg, "TKToken", "002")
    If orsCfg.BOF And orsCfg.EOF Then
        tVB_TokenNAEUsr = ""
    Else
        tVB_TokenNAEUsr = SP_FEDtChkString(orsCfg, "FTScfUsrValue")
    End If
    
    Call SP_OPTxGetConfigValue(orsCfg, "TKToken", "003")
    If orsCfg.BOF And orsCfg.EOF Then
        tVB_TokenNAEPwd = ""
    Else
        tVB_TokenNAEPwd = SP_FEDtChkString(orsCfg, "FTScfUsrValue")
    End If
    
    Call SP_OPTxGetConfigValue(orsCfg, "TKToken", "004")
    If orsCfg.BOF And orsCfg.EOF Then
        tVB_TokenDB = ""
    Else
        tVB_TokenDB = SP_FEDtChkString(orsCfg, "FTScfUsrValue")
    End If
    
    Call SP_OPTxGetConfigValue(orsCfg, "TKToken", "005")
    If orsCfg.BOF And orsCfg.EOF Then
        tVB_TokenDBUsr = ""
    Else
        tVB_TokenDBUsr = SP_FEDtChkString(orsCfg, "FTScfUsrValue")
    End If
    
    Call SP_OPTxGetConfigValue(orsCfg, "TKToken", "006")
    If orsCfg.BOF And orsCfg.EOF Then
        tVB_TokenDBPwd = ""
    Else
        tVB_TokenDBPwd = SP_FEDvChkNumeric(orsCfg, "FTScfUsrValue")
    End If
    
    Call SP_OPTxGetConfigValue(orsCfg, "TKToken", "007")
    If orsCfg.BOF And orsCfg.EOF Then
        tVB_TokenDBPwd = "4"
    Else
        Select Case UCase(SP_FEDtChkString(orsCfg, "FTScfUsrValue"))
        Case UCase("FIRST_SIX_TOKEN")
            tVB_TokenFormat = "4"
        Case Else
            tVB_TokenFormat = "4"
        End Select
    End If
    
    Call SP_OPTxGetConfigValue(orsCfg, "TKToken", "008")
    If orsCfg.BOF And orsCfg.EOF Then
        tVB_TokenLuhCheck = "0"
    Else
        tVB_TokenLuhCheck = SP_FEDtChkString(orsCfg, "FTScfUsrValue")
    End If
    
    Exit Sub
Err_Handle:
    'Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsCfg)
    On Error GoTo 0
End Sub

Public Sub SP_OPTxGetConfigValue(ByRef porsCfg As ADODB.Recordset, ByVal ptVCode$, ByVal ptVSeq$)
'----------------------------------------------------------
'   Call:
'   Cmt:  Get first varaiable
'----------------------------------------------------------
    Dim tWhe$
    
    porsCfg.Filter = adFilterNone
    If porsCfg.BOF And porsCfg.EOF Then Exit Sub
    
     tWhe = " FTScfCode='" & ptVCode & "'"
    If ptVSeq <> "" Then tWhe = tWhe & " AND FTScfSeq ='" & ptVSeq & "'"
    porsCfg.Filter = tWhe

End Sub

Public Function SP_CHKbWebservice() As Boolean
Dim nErrorNo As Long
Dim tErrorDesc As String

On Error GoTo Err_Handle
    'SP_CHKbWebservice = True
    '** Store Parameter บอกว่า ไม่ต้อง Token
    'If Not bVB_AlwToken Then Exit Function '*Tao 61-05-09 ปิดไว้ (ถ้า OffLine ตั้งแต่แรก จะเป็น False)
    
    '*Tao 61-05-03 clear object ก่อน
    If Not (oToken Is Nothing) Then
        Set oToken = Nothing
        Call SleepEx(200, 0) '*จาก 500 เป็น 200
    End If
    '========================
    
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_CHKbWebservice Start."

    Set oToken = CreateObject("MSSOAP.SoapClient30")
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_CHKbWebservice CreateObject OK.."
    '*Tao 61-05-17 Set ClientProperty
    oToken.ClientProperty("ServerHTTPRequest") = True
    '===========================
    Call oToken.MSSoapInit(tVB_TokenURL, "", "")
    nErrorNo = Err.Number
    tErrorDesc = Err.Description
    
    If Err <> 0 Then
        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_CHKbWebservice MSSoapInit Err: " & nErrorNo & ",Err Desc:" & tErrorDesc
        SP_CHKbWebservice = False
    Else
        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_CHKbWebservice MSSoapInit OK.."
        '*Tao 61-05-17 Set ConnectorProperty
        'oToken.ConnectorProperty("UseSSL") = False
        oToken.ConnectorProperty("Timeout") = 10000
        oToken.ConnectorProperty("ConnectTimeout") = 10000
        '===============================
    End If
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_CHKbWebservice OK."
    SP_CHKbWebservice = True
    Exit Function
Err_Handle:
    Set oToken = Nothing '*Tao 61-05-03 clear object
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SP_CHKbWebservice Err: " & Err.Number & ",Err Desc:" & Err.Description
    SP_CHKbWebservice = False
End Function

Public Sub SP_SETxObjects()
On Error GoTo Err_Handle

    '*Tao 61-05-03 clear object ก่อน
    If Not (oToken Is Nothing) Then
        Set oToken = Nothing
        Call SleepEx(200, 0) '*จาก 500 เป็น 200
    End If
    '========================

    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Start CreateObject"

    '*Bmp 61-02-15 C:\Program Files (x86)\Common Files\MSSoap\Binaries\MSSOAP30.dll
    If oToken Is Nothing Then
        bVB_AbleToken = False
        Set oToken = CreateObject("MSSOAP.SoapClient30")
    End If
    
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " CreateObject OK."    '*Tao 61-05-09 เก็บ Log เพิ่ม
    '*Tao 61-05-17 Set ClientProperty
    oToken.ClientProperty("ServerHTTPRequest") = True
    '===========================
    
    'Set oToken = CreateObject("MSSOAP.SoapClient30")
'    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "tVB_TokenURL:" & tVB_TokenURL
'    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "tVB_TokenNAEUsr:" & tVB_TokenNAEUsr
'    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "tVB_TokenNAEPwd:" & tVB_TokenNAEPwd
'    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "tVB_TokenDBUsr:" & tVB_TokenDBUsr
'    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "tVB_TokenDBPwd:" & tVB_TokenDBPwd
'    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "tVB_TokenFormat:" & tVB_TokenFormat
'
'    Public tVB_TokenURL As String 'http://192.11.15.248:8080/axis2/services/SafeNetTokenizer?wsdl
'    Public tVB_TokenNAEUsr As String 'nae_user
'    Public tVB_TokenNAEPwd As String 'P@ssw0rd
'    Public tVB_TokenDB As String 'dbToken
'    Public tVB_TokenDBUsr As String 'tkuser
'    Public tVB_TokenDBPwd As String '!@#Token
'    Public tVB_TokenFormat As String '4: FIRST_SIX_TOKEN

    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Start Soap Init"    '*Tao 61-05-09 เก็บ Log เพิ่ม
    
'    If (bVB_AbleToken = False) Then '*Tao 61-05-09 ปิดไว้ (ถ้าก่อนหน้าเป็น True จะไม่ Init)
       Call oToken.MSSoapInit(tVB_TokenURL, "", "")
        If Err <> 0 Then
            '* คิดก่อนว่าจะให้ servicetransfer write log ลงที่ไหน Call SP_CRTxLogMallCardExtend(ptFuncName & "Initial Token Web service Error: " & Err.Number & ":" & Err.Description)
            'MsgBox "Initial Token Web service Error: " & Err.Number & ":" & Err.Description, vbExclamation
            oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Innitial Token Web service is not success."
            bVB_AbleToken = False
            Exit Sub
        Else
            '*Tao 61-05-17 Set ConnectorProperty
            'oToken.ConnectorProperty("UseSSL") = False
            oToken.ConnectorProperty("Timeout") = 10000
            oToken.ConnectorProperty("ConnectTimeout") = 10000
            '===============================
        End If
'    End If
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Start Soap Init OK"    '*Tao 61-05-09 เก็บ Log เพิ่ม

    bVB_AbleToken = True
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " bVB_AbleToken:" & bVB_AbleToken
    Exit Sub
Err_Handle:
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " CreateObject err: " & Err.Number & ",Err Desc:" & Err.Description
    bVB_AbleToken = False
End Sub

Public Function SP_PRCbINSTokenList(ByVal ptTblName As String, ByVal ptFldName As String, ByVal ptTmnNum As String, ByVal ptShdTransNo As String, ByVal pdShdTransDate As Date, ByVal ptFTValue As String, ByVal ptFTToken As String) As Boolean
    Dim tSql As String
    On Error GoTo Err_Handler
    
    tSql = "INSERT INTO TPSTTokenLst (FDDateIns,FTWhoIns,FTTblName,FTFldName,FTTmnNum,FTShdTransNo,FDShdTransDate,FTValue,FTToken,FNStatus,FTSrvName)"
    tSql = tSql & vbCrLf & "SELECT CONVERT(VARCHAR(20),GETDATE(),121),'ServiceTransfer','" & ptTblName & "','" & ptFldName & "','" & ptTmnNum & "','" & ptShdTransNo & "','" & Format(pdShdTransDate, "yyyy/MM/dd") & "','" & ptFTValue & "','" & ptFTToken & "','1',' " & tMM_CNDbOnServer & "'"
    'SELECT Convert(Varchar(20),Getdate(),121),'System', 'TPSTSalCD'
    ','FTScdCardID',FTTmnNum,FTShdTransNo,FDShdTransDate,'" & tCardNo & "','" & tToken & "','1','" & tVB_SrvName & "' "
    
    SP_PRCbINSTokenList = True
    Exit Function
Err_Handler:
    SP_PRCbINSTokenList = False
End Function
