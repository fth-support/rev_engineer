VERSION 5.00
Begin VB.Form wMain 
   ClientHeight    =   2985
   ClientLeft      =   6450
   ClientTop       =   1875
   ClientWidth     =   6630
   Icon            =   "wMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox otbForm 
      Height          =   2175
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "wMain.frx":058A
      Top             =   120
      Width           =   5535
   End
   Begin VB.Timer otmForm 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "wMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tW_GetStaSent As String
Dim tW_GetDate As String
Dim tW_GetTime As String
Const nMaxDay = 7
Private Function W_PRCbIINSTokenList(ByVal ptTblName As String, ByVal ptFldName As String, ByVal ptTmnNum As String, ByVal ptShdTransNo As String, ByVal pdShdTransDate As Date, ByVal ptFTValue As String, ByVal ptFTToken As String) As Boolean
    Dim tSql As String
    On Error GoTo Err_Handler
    tSql = "INSERT INTO TPSTTokenLst (FDDateIns,FTWhoIns,FTTblName,FTFldName,FTTmnNum,FTShdTransNo,FDShdTransDate,FTValue,FTToken,FNStatus,FTSrvName)"
    tSql = tSql & vbCrLf & "SELECT CONVERT(VARCHAR(20),GETDATE(),121),'ServiceTransfer','" & ptTblName & "','" & ptFldName & "','" & ptTmnNum & "','" & ptShdTransNo & "','" & Format(pdShdTransDate, "yyyy/MM/dd") & "','" & ptFTValue & "','" & ptFTToken & "','1',' " & tMM_CNDbOnServer & "'"
    'SELECT Convert(Varchar(20),Getdate(),121),'System', 'TPSTSalCD'
    ','FTScdCardID',FTTmnNum,FTShdTransNo,FDShdTransDate,'" & tCardNo & "','" & tToken & "','1','" & tVB_SrvName & "' "
    
    Call ocnOnline.Execute(tSql)
    
    W_PRCbIINSTokenList = True
    Exit Function
Err_Handler:
    W_PRCbIINSTokenList = False
End Function
Private Function W_PRCbCheckInsertHD(ByVal ptTmnNum As String, ByVal ptTrnNo As String, ByVal pdDate As Date) As Boolean
'------------------------------------------------------------
'   Call:
'   Cmt:        '*Ao 2014/08/21     V 4.0.10            Insert HD ไม่สำเร็จ ให้ปรับ DT,CD,RC FTStaSentOnOff = '0'
'------------------------------------------------------------
Dim bChk As Boolean
Dim tSql As String
Dim orsTemp As ADODB.Recordset

On Error GoTo Err_Handler
    
    bChk = False
    
    tSql = "SELECT FTTmnNum,FTShdTransNo"
    tSql = tSql & vbCrLf & "FROM TPSTSalHD" & " WITH(NOLOCK)"
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptTmnNum & "' "
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptTrnNo & "' "
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(pdDate, "yyyy/MM/dd") & "' "       'ON LINE
    Set orsTemp = ocnOnline.Execute(tSql)
    If orsTemp.BOF And orsTemp.EOF Then
    
        bChk = True
        
          tSql = "UPDATE TPSTSALEPOINT" & " WITH(ROWLOCK)"  '* [5.1803.4][61/03/30] *Bmp 61-03-30
        tSql = tSql & vbCrLf & "SET FTStaSentOnOff = '0' "
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptTmnNum & "' "
        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptTrnNo & "' "
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(pdDate, "yyyy/MM/dd") & "' "   'OFF LINE
        Call ocnOffline.Execute(tSql)
        
        tSql = "UPDATE TPSTSalDT" & " WITH(ROWLOCK)"
        tSql = tSql & vbCrLf & "SET FTStaSentOnOff = '0' "
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptTmnNum & "' "
        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptTrnNo & "' "
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(pdDate, "yyyy/MM/dd") & "' "   'OFF LINE
        Call ocnOffline.Execute(tSql)
        
        tSql = "UPDATE TPSTSalRC" & " WITH(ROWLOCK)"
        tSql = tSql & vbCrLf & "SET FTStaSentOnOff = '0' "
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptTmnNum & "' "
        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptTrnNo & "' "
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(pdDate, "yyyy/MM/dd") & "' "   'OFF LINE
        Call ocnOffline.Execute(tSql)
        
        tSql = "UPDATE TPSTSalCD" & " WITH(ROWLOCK)"
        tSql = tSql & vbCrLf & "SET FTStaSentOnOff = '0' "
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptTmnNum & "' "
        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptTrnNo & "' "
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(pdDate, "yyyy/MM/dd") & "' "   'OFF LINE
        Call ocnOffline.Execute(tSql)
        
        tSql = "UPDATE TPSTSalHD" & " WITH(ROWLOCK)"
        tSql = tSql & vbCrLf & "SET FTStaSentOnOff = '0' "
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptTmnNum & "' "
        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptTrnNo & "' "
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(pdDate, "yyyy/MM/dd") & "' "   'OFF LINE
        Call ocnOffline.Execute(tSql)
    End If
    
    If Not (orsTemp Is Nothing) Then
        If orsTemp.State = adStateOpen Then orsTemp.Close
        Set orsTemp = Nothing
    End If
    
    W_PRCbCheckInsertHD = bChk
Exit Function
Err_Handler:
    W_PRCbCheckInsertHD = True
End Function

Public Function SP_DTEtStrSTD(pdDate As Date, Optional ByVal pbChrRet As Boolean = False, Optional ByVal ptTime$, Optional ByVal pbRetWithTime As Boolean = False) As String
'----------------------------------------------------------
'   Call:   pdDate as selected date
'   Ret:   string of date format "YYYY/MM/DD"
'----------------------------------------------------------
    Dim tRet$
    
    SP_DTEtStrSTD = Format(Year(pdDate), "0000") & "/" & Format(Month(pdDate), "00") & "/" & Format(Day(pdDate), "00")
    tRet = SP_DTEtStrSTD
    If pbChrRet Then
        Select Case nVB_CNBase
            Case nEN_Dbs1ACS
                  If pbRetWithTime Then
                       SP_DTEtStrSTD = "#" & tRet & " " & ptTime & "#"
                  Else
                       SP_DTEtStrSTD = "#" & tRet & "#"
                  End If
                
            Case nEN_Dbs2SQL
                  If pbRetWithTime Then
                       SP_DTEtStrSTD = "'" & tRet & " " & ptTime & "'"
                  Else
                       SP_DTEtStrSTD = "'" & tRet & "'"
                  End If
        End Select
    Else
        If pbRetWithTime Then
            SP_DTEtStrSTD = tRet & " " & ptTime
        End If
    End If
End Function

Private Function W_DATbConnect(ByRef pocnTemp As ADODB.Connection, ByVal ptServer As String, ByVal ptUser As String, ByVal ptPassword As String, ByVal ptDB As String) As Boolean
'-------------------------------------------------------------
'Ret: Connect = True/Not Connect = False
'Cmt: Connect DB Online/Offline
'------------------------------------------------------------
    Dim tCon As String
    On Error GoTo Err_Handler
    If Not (pocnTemp Is Nothing) Then
        Set pocnTemp = Nothing
    Else
    
    End If
    Set pocnTemp = New ADODB.Connection
    tCon = "Provider=SQLOLEDB.1;Password=" & ptPassword & ";Persist Security Info=True;User ID=" & ptUser & ";Initial Catalog=" & ptDB & ";Data Source=" & ptServer
    pocnTemp.CursorLocation = adUseServer
    pocnTemp.ConnectionTimeout = 15
    pocnTemp.CommandTimeout = 90
    pocnTemp.Open tCon
    If pocnTemp.State = adStateOpen Then
        W_DATbConnect = True
    Else
        W_DATbConnect = False
    End If
    Exit Function
Err_Handler:
End Function

Private Function W_DATbConnectOrChk(ByRef pocnTemp As ADODB.Connection, ByVal ptServer As String, ByVal ptUser As String, ByVal ptPassword As String, ByVal ptDB As String, ByVal pbServer As Boolean, ByRef pnErr As Long) As Boolean
'-------------------------------------------------------------
'Ret: Connect = True/Not Connect = False
'Cmt: Connect DB Online/Offline
'          *KT 52-10-21 ปรับปรุง if pocnTemp.State = adStateOpen Then /  tOnline = GetSetting("AdaPosMall", "POSFront", "Online")
'------------------------------------------------------------
    Dim tCon As String
    Dim tSql As String
    Dim tOnline As String
    
On Error GoTo Err_Handler
    
    If Not (pocnTemp Is Nothing) Then
        If pocnTemp.State = adStateOpen Then
            tSql = "SELECT FTTmnCode  FROM TCNMTerminalMtn " '*KT 52-12-21 ใช้ Table ที่ไม่มีการ Replicate แทน
            pocnTemp.Execute tSql
            W_DATbConnectOrChk = True
            oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " W_DATbConnectOrChk = " & W_DATbConnectOrChk    '*Tao 61-05-09 เก็บ Log เพิ่ม
            pnErr = 0
            Exit Function
        End If
    End If
    
    If Not (pocnTemp Is Nothing) Then
        Set pocnTemp = Nothing
    End If
    
    tOnline = GetSetting("AdaPosMall", "POSFront", "Online")
    If pbServer Then
        If tOnline = "1" Then
            Set pocnTemp = New ADODB.Connection
            tCon = "Provider=SQLOLEDB.1;Password=" & ptPassword & ";Persist Security Info=True;User ID=" & ptUser & ";Initial Catalog=" & ptDB & ";Data Source=" & ptServer
            pocnTemp.CursorLocation = adUseClient
            pocnTemp.ConnectionTimeout = 30
            pocnTemp.CommandTimeout = 60
            pocnTemp.Open tCon
            If pocnTemp.State = adStateOpen Then
                Call SP_GETxVariableAlw '*Tao 61-05-09 Get Variable ใหม่
                W_DATbConnectOrChk = True
            Else
                W_DATbConnectOrChk = False
            End If
        Else
            W_DATbConnectOrChk = False
        End If
    Else
        Set pocnTemp = New ADODB.Connection
        tCon = "Provider=SQLOLEDB.1;Password=" & ptPassword & ";Persist Security Info=True;User ID=" & ptUser & ";Initial Catalog=" & ptDB & ";Data Source=" & ptServer
        pocnTemp.CursorLocation = adUseClient
        pocnTemp.ConnectionTimeout = 30
        pocnTemp.CommandTimeout = 60
        pocnTemp.Open tCon
        If pocnTemp.State = adStateOpen Then
            W_DATbConnectOrChk = True
        Else
            W_DATbConnectOrChk = False
        End If
    End If
    pnErr = 0
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " W_DATbConnectOrChk = " & W_DATbConnectOrChk    '*Tao 61-05-09 เก็บ Log เพิ่ม
    Exit Function
Err_Handler:
    Select Case Err.Number
        Case -2147467259, 91, 3704, -2147217871 '*KT 54-04-11
                Set pocnTemp = Nothing
                W_DATbConnectOrChk = False
            Case Else
                W_DATbConnectOrChk = False
    End Select
    pnErr = Err.Number
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " W_DATbConnectOrChk = " & W_DATbConnectOrChk    '*Tao 61-05-09 เก็บ Log เพิ่ม
End Function

Private Function W_DATbInsertData(ByVal porsTemp As ADODB.Recordset, ByVal ptTblName As String, ByRef pnErrNo As Long) As Boolean
'-------------------------------------------------------------
'Ret:
'Cmt: Transfer data to DB Online
'------------------------------------------------------------
    Dim tSql As String
    Dim tField As String
    Dim tSumField As String
    Dim tValues As String
    Dim tSumValues As String
    Dim orsTemp As ADODB.Recordset
    Dim nI As Integer
    Dim tTmnNum$
    Dim tTrnNo$
    Dim dDateTrn As Date
    Dim orsOffline As New ADODB.Recordset
    Dim nEff As Long
    Dim bNeedToken As Boolean
    Dim tCardNo As String
    Dim tTokenCardNo As String
    Dim tTransType As String    '*Bmp 61-04-03
    Dim tRC_FTTdmCode As String
    Dim tSQL_Token As String
    Dim tFTFldName As String
    Dim tCardDeToken As String
On Error GoTo Err_Handler
    '*KT 51-10-02 1.0.2
    '---------------------------------------------------------------------
    
    W_DATbInsertData = False
    bNeedToken = False
    If UCase(ptTblName) = UCase("TSysLogTableUpdate") Then       '*Ao 2013/01/09     TK
        'ไม่มีการทำงาน
    ElseIf UCase(ptTblName) = UCase("TPSTCV") Then       '*Bump 56-06-18 CommSheet TK-ISS3000-234-ST-00.00.03.xls CV Offline-->Online
        'ไม่มีการทำงาน
    Else
        tTmnNum = SP_FEDtChkString(porsTemp, "FTTmnNum")
        tTrnNo = SP_FEDtChkString(porsTemp, "FTShdTransNo")
        dDateTrn = SP_FEDdChkDate(porsTemp, "FDShdTransDate")
         tTransType = SP_FEDtChkString(porsTemp, "FTShdTransType") '*Bmp 61-04-03
    End If
    
    tRC_FTTdmCode = ""
    If UCase(ptTblName) = UCase("TPSTDepositRC") Or UCase(ptTblName) = UCase("TPSTSalRC") Then
        tRC_FTTdmCode = SP_FEDtChkString(porsTemp, "FTTdmCode") '*Bmp 61-04-04 v.1803.7
    End If
    If UCase(ptTblName) = UCase("TPSTDepositCD") Or UCase(ptTblName) = UCase("TPSTSalCD") Then
        tRC_FTTdmCode = SP_FEDtChkString(porsTemp, "FNDctNo") '*Bmp 61-04-04 v.1803.7
    End If
    
    Select Case UCase(ptTblName)
        Case UCase("TPSTSalHD")
            tSql = "SELECT FTTmnNum,FTShdTransNo"                       '*Ao 2014/03/06   CommSheet TK-ISS3000-470
            tSql = tSql & vbCrLf & "FROM TPSTSalDT WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(dDateTrn, "yyyy/MM/dd") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTStaSentOnOff,'0') <> '1'"
            Set orsOffline = ocnOffline.Execute(tSql, nEff)
            If Not (orsOffline.BOF And orsOffline.EOF) Then
                pnErrNo = 555
                Exit Function
            End If
            Call SP_DATxRsNothing(orsOffline)
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSalRC WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(dDateTrn, "yyyy/MM/dd") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTStaSentOnOff,'0') <> '1'"
            Set orsOffline = ocnOffline.Execute(tSql, nEff)
            If Not (orsOffline.BOF And orsOffline.EOF) Then
                pnErrNo = 555
                Exit Function
            End If
            Call SP_DATxRsNothing(orsOffline)
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSalCD WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(dDateTrn, "yyyy/MM/dd") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTStaSentOnOff,'0') <> '1'"
            Set orsOffline = ocnOffline.Execute(tSql, nEff)
            If Not (orsOffline.BOF And orsOffline.EOF) Then
                pnErrNo = 555
                Exit Function
            End If
            Call SP_DATxRsNothing(orsOffline)
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSalePoint WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(dDateTrn, "yyyy/MM/dd") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTStaSentOnOff,'0') <> '1'"
            Set orsOffline = ocnOffline.Execute(tSql, nEff)
            If Not (orsOffline.BOF And orsOffline.EOF) Then
                pnErrNo = 555
                Exit Function
            End If
        Case UCase("TPSTDepositHD")
            tSql = "SELECT FTTmnNum,FTShdTransNo"                           '*Ao 2014/03/06   CommSheet TK-ISS3000-470
            tSql = tSql & vbCrLf & "FROM TPSTDepositDT WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(dDateTrn, "yyyy/MM/dd") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTStaSentOnOff,'0') <> '1'"
            Set orsOffline = ocnOffline.Execute(tSql, nEff)
            If Not (orsOffline.BOF And orsOffline.EOF) Then
                pnErrNo = 555
                Exit Function
            End If
            Call SP_DATxRsNothing(orsOffline)
            tSql = "SELECT FTTmnNum,FTShdTransNo"                           '*Ao 2014/03/06   CommSheet TK-ISS3000-470
            tSql = tSql & vbCrLf & "FROM TPSTDepositRC WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(dDateTrn, "yyyy/MM/dd") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTStaSentOnOff,'0') <> '1'"
            Set orsOffline = ocnOffline.Execute(tSql, nEff)
            If Not (orsOffline.BOF And orsOffline.EOF) Then
                pnErrNo = 555
                Exit Function
            End If
            Call SP_DATxRsNothing(orsOffline)
            tSql = "SELECT FTTmnNum,FTShdTransNo"                           '*Ao 2014/03/06   CommSheet TK-ISS3000-470
            tSql = tSql & vbCrLf & "FROM TPSTDepositCD WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & Format(dDateTrn, "yyyy/MM/dd") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTStaSentOnOff,'0') <> '1'"
            Set orsOffline = ocnOffline.Execute(tSql, nEff)
            If Not (orsOffline.BOF And orsOffline.EOF) Then
                pnErrNo = 555
                Exit Function
            End If
        Case Else
        
    End Select
'    If tTrnNo = "00678" Then Stop
    Select Case UCase(ptTblName)
        Case UCase("TPSHD" & tVB_CNTmnNum)
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSHD" & tVB_CNTmnNum & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSDT" & tVB_CNTmnNum)
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSHD" & tVB_CNTmnNum & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSdtSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSRC" & tVB_CNTmnNum)
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSHD" & tVB_CNTmnNum & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSrcSeqNo= " & SP_FEDvChkNumeric(porsTemp, "FNSrcSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSCD" & tVB_CNTmnNum)
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSHD" & tVB_CNTmnNum & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSdtSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
                tSql = tSql & vbCrLf & "AND FNScdSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNScdSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTSALHD")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTSALDT")
'            tSql = "SELECT FTTmnNum,FTShdTransNo"
'            tSql = tSql & vbCrLf & "FROM TPSTSALHD"
'            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
'            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
'            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
'            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSalDT" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND FNSdtSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")       '*Ao 2014/03/25
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSdtSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTSALRC")
'            tSql = "SELECT FTTmnNum,FTShdTransNo"
'            tSql = tSql & vbCrLf & "FROM TPSTSALHD"
'            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
'            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
'            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
'            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSalRC" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND FNSrcSeqNo= " & SP_FEDvChkNumeric(porsTemp, "FNSrcSeqNo")       '*Ao 2014/03/25
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSrcSeqNo= " & SP_FEDvChkNumeric(porsTemp, "FNSrcSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTSALCD")
'            tSql = "SELECT FTTmnNum,FTShdTransNo"
'            tSql = tSql & vbCrLf & "FROM TPSTSALHD"
'            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
'            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
'            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
'            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSalCD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"           '*Ao 2014/03/25
            tSql = tSql & vbCrLf & "AND FNSdtSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
            tSql = tSql & vbCrLf & "AND FNScdSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNScdSeqNo")
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSdtSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
                tSql = tSql & vbCrLf & "AND FNScdSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNScdSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTLOGDT")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNLogSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNLogSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTHOLDHD")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTHOLDDT")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSdtSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTHOLDCD")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSdtSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
                tSql = tSql & vbCrLf & "AND FNScdSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNScdSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        'DEPOSIT
        Case UCase("TPSTDEPOSITHD")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTDEPOSITDT")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSdtSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTDEPOSITCD")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSdtSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNSdtSeqNo")
                tSql = tSql & vbCrLf & "AND FNScdSeqNo =" & SP_FEDvChkNumeric(porsTemp, "FNScdSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTDEPOSITRC")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSrcSeqNo= " & SP_FEDvChkNumeric(porsTemp, "FNSrcSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        'VOUCHER
        Case UCase("TPSTVOUCHERHD")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TPSTVOUCHERRC")
            tSql = "SELECT FTTmnNum,FTShdTransNo"
            tSql = tSql & vbCrLf & "FROM TPSTSALHD" & " WITH(NOLOCK)"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaDoc ,'0') <> '1'"
            Set orsTemp = ocnOnline.Execute(tSql)
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                tSql = "DELETE " & ptTblName & " WITH(ROWLOCK)"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tTmnNum & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tTrnNo & "'"
                tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & Format(dDateTrn, "YYYY/MM/DD") & "'"
                tSql = tSql & vbCrLf & "AND FNSrcSeqNo= " & SP_FEDvChkNumeric(porsTemp, "FNSrcSeqNo")
                Call ocnOnline.Execute(tSql)
            End If
        Case UCase("TSysLogTableUpdate")            '*Ao 2013/01/09     TK
            'ไม่มีการทำงาน
        Case UCase("TPSTCV")            '*Bump 56-06-18 CommSheet TK-ISS3000-234-ST-00.00.03.xls CV Offline-->Online
            'ไม่มีการทำงาน
    End Select
'----------------------------------------------------------------------------

    With porsTemp
        For nI = 0 To .Fields.Count - 1
            tField = .Fields(nI).Name
            tSumField = IIf(tSumField = "", "", tSumField & ",") & tField
            Select Case Mid(tField, 2, 1)
                Case "T"
                    bNeedToken = False
                    '---------------------------------------------------------------------
                   If UCase(tField) = UCase("FTSendUp") Then
                     tValues = "'" & "0" & "'"  '*Eaw 56-01-21
                   ElseIf UCase(tField) = UCase("FTRemark") Then '*[5.1804.13][61/04/24]*Bmp 61-04-24
                        tValues = "'" & SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField)) & "'"
                   Else
'                     tValues = "'" & SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField)) & "'"
'                     '*Bmp 61-02-27 เพิ่มการ Token หมายเลขบัตร
                     Select Case UCase(ptTblName)
                        Case UCase("TPSTSalHD"), UCase("TPSTDepositHD") 'FTCstCardCode
                            If UCase(tField) = UCase("FTCstCardCode") Then
                                bNeedToken = True
                                tFTFldName = "FTCstCardCode"
                            End If
                        Case UCase("TPSTSalCD"), UCase("TPSTDepositCD") 'FTScdCardID
                            If UCase(tField) = UCase("FTScdCardID") Then
                                Select Case tRC_FTTdmCode
                                Case "11", "12", "15", "28" '=FNDctNo IN(11,12,15,28)
                                    bNeedToken = True
                                    tFTFldName = "FTScdCardID"
                                Case Else
                                    bNeedToken = False
                                End Select
                            Else
                                bNeedToken = False
                            End If
                        Case UCase("TPSTSalRC"), UCase("TPSTDepositRC") 'FTSrcCardNo
                            If UCase(tField) = UCase("FTSrcCardNo") Then
                                Select Case tRC_FTTdmCode
                                Case "T002", "T003", "T017" 'T002 Credit card, T003 Hir Purchase, T017 TC Card
                                    bNeedToken = True
                                    tFTFldName = "FTSrcCardNo"
                                Case Else
                                    bNeedToken = False
                                End Select
                            Else
                                bNeedToken = False
                          End If
                        Case UCase("TPSTSalePoint") 'FTSpoMemID
                            If UCase(tField) = UCase("FTSpoMemID") Then
                                bNeedToken = True
                                tFTFldName = "FTSpoMemID"
                            End If
                        Case Else
                            bNeedToken = False
                     End Select
                    '*Bmp 61-04-03 Transaction Type ไม่ต้อง token (double-tokenized)
                    '*Bmp 61-04-05
                     Select Case tTransType
                        Case "07", "19", "20", "16" '*Bmp 61-04-11 v.9, '*Bmp 61-04-11 20.44 น. เพิ่ม '16' v.5.1804.10
                        '19 Add Layaway
                        '20 Add Layaway Cancel
                        'Case "07" '*Layaway Pickup Token  RC,salepoint    '[5.1804.4][61/04/05]*Bmp 61-04-05 16.23
                            bNeedToken = False
                                If (UCase(ptTblName) = UCase("TPSTSalRC")) Or (UCase(ptTblName) = UCase("TPSTDepositRC")) Then
                                    If UCase(tField) = UCase("FTSrcCardNo") Then
                                        Select Case tRC_FTTdmCode
                                            Case "T002", "T003", "T017"
                                                bNeedToken = True
                                                tFTFldName = "FTSrcCardNo"
                                            Case Else
                                        End Select
                                    End If
                                ElseIf (UCase(ptTblName) = UCase("TPSTSalHD")) Or (UCase(ptTblName) = UCase("TPSTDepositHD")) Then  'issue e.salepoint not token *Bmp 61-04-20 v.12 เปลี่ยนจาก TPSTSaleHD เป็น TPSTSalHD ให้ถูกต้อง
                                    If UCase(tField) = UCase("FTCstCardCode") Then
                                        bNeedToken = True '*Bmp 61-04-06 Salepoint
'                                        '*ต้องเช็คก่อนว่าถูก Token แล้วหรือยัง ถ้ายังต้อง token
'                                        '=========*Bmp 61-04-09
'                                        If SP_PRCbDeToken(SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField)), tCardDeToken) Then
'                                            'ถ้า detoken ได้แสดงว่า เคย token แล้ว ไม่ต้อง token ซ้ำ
'                                            If tCardDeToken = "" Then
'                                                bNeedToken = False 'ถ้า detoken ได้แสดงว่า เคย token แล้ว ไม่ต้อง token ซ้ำ
'                                            Else
'                                                bNeedToken = True 'ถ้า detoken ได้เป็นหมายเลขว่าง แสดงว่า ยังไม่เคย token หมายเลขนี้ จึงต้อง token
'                                            End If
'                                        Else
'                                            bNeedToken = True
'                                        End If
'                                        '=========*Bmp 61-04-09
                                        tFTFldName = "FTCstCardCode"
                                    End If
                                ElseIf (UCase(ptTblName) = UCase("TPSTSalePoint")) Then 'issue e.salepoint not token
                                    If UCase(tField) = UCase("FTSpoMemID") Then
                                        bNeedToken = True '*Bmp 61-04-06 Salepoint
'                                        '*ต้องเช็คก่อนว่าถูก Token แล้วหรือยัง ถ้ายังต้อง token
'                                        '=========*Bmp 61-04-09
'                                        If SP_PRCbDeToken(SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField)), tCardDeToken) Then
'                                            'ถ้า detoken ได้แสดงว่า เคย token แล้ว ไม่ต้อง token ซ้ำ
'                                            If tCardDeToken = "" Then
'                                                bNeedToken = False 'ถ้า detoken ได้แสดงว่า เคย token แล้ว ไม่ต้อง token ซ้ำ
'                                            Else
'                                                bNeedToken = True 'ถ้า detoken ได้เป็นหมายเลขว่าง แสดงว่า ยังไม่เคย token หมายเลขนี้ จึงต้อง token
'                                            End If
'                                        Else
'                                            bNeedToken = True
'                                        End If
'                                        '=========*Bmp 61-04-09
                                        tFTFldName = "FTSpoMemID"
                                    End If
                                End If
                        Case "01", "02", "05", "08", "09", "12", "14" '*Bmp 61-04-11 Type'16' ต้อง Token ด้วย "16" '*Bmp 61-07-10 ต้อง Token Return ด้วย"04" v.5.1804.8
                            bNeedToken = False
                        Case "17", "18", "21", "24", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37" '*Bmp 61-04-11 เอา"20" ออก
                            bNeedToken = False
                        Case "40", "41", "42", "45", "46", "47", "48", "49", "50"
                            bNeedToken = False
                     End Select
                     '==================================================
                     tCardNo = ""
                     tCardNo = SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField))
'                     If (Trim(tCardNo) <> "") And bNeedToken Then
                            If bNeedToken Then
                               'หาใน db ก่อน ถ้าไม่มี จึง token tTmnNum
                               tTokenCardNo = ""
                               'If SP_PRCbToken(ptTblName, tCardNo, tTokenCardNo) Then '*Tao 61-05-21 ของเดิม
                               If SP_PRCbToken(ptTblName, tCardNo, tTokenCardNo, True) Then '*Tao 61-05-21 ไม่ต้องค้นหาในฐานข้อมุลก่อน
                                   tValues = "'" & tTokenCardNo & "'"
                                   '*Tao 61-05-21 ของเดิม
'                                   tSQL_Token = "IF NOT EXISTS(SELECT FTValue FROM TPSTTokenLst WHERE FTTblName = '" & ptTblName & "' AND FTFldName = '" & tFTFldName & "' AND FTTmnNum='" & tTmnNum & "' AND FTShdTransNo ='" & tTrnNo & "' AND FDShdTransDate = '" & Format(dDateTrn, "YYYY/MM/DD") & "' AND FTValue = '" & tCardNo & "')"
'                                   tSQL_Token = tSQL_Token & vbCrLf & "BEGIN"
'                                   tSQL_Token = tSQL_Token & vbCrLf & "INSERT INTO TPSTTokenLst(FDDateIns,FTWhoIns,FTTblName,FTFldName,FTTmnNum,FTShdTransNo,FDShdTransDate,FTValue,FTToken,FNStatus,FTSrvName)"
'                                   tSQL_Token = tSQL_Token & vbCrLf & "VALUES(GETDATE(),'ServiceTransfer','" & ptTblName & "','" & tFTFldName & "','" & tTmnNum & "','" & tTrnNo & "','" & Format(dDateTrn, "YYYY/MM/DD") & "','" & tCardNo & "','" & tTokenCardNo & "','1','" & tMM_CNDbOnServer & "')"
'                                   tSQL_Token = tSQL_Token & vbCrLf & "END"

                                   '*Tao 61-05-22 ถ้ามีแล้ว Update FTToken ด้วย (ให้ POSFront ใช้หาข้อมูล Token)
                                   tSQL_Token = "IF NOT EXISTS(SELECT FTValue FROM TPSTTokenLst WITH(NOLOCK) WHERE FTTblName = '" & ptTblName & "' AND FTFldName = '" & tFTFldName & "' AND FTValue = '" & tCardNo & "')"
                                   tSQL_Token = tSQL_Token & vbNewLine & "BEGIN"
                                   tSQL_Token = tSQL_Token & vbNewLine & "INSERT INTO TPSTTokenLst(FDDateIns,FTWhoIns,FTTblName,FTFldName,FTTmnNum,FTShdTransNo,FDShdTransDate,FTValue,FTToken,FNStatus,FTSrvName)"
                                   tSQL_Token = tSQL_Token & vbNewLine & "VALUES(GETDATE(),'ServiceTransfer','" & ptTblName & "','" & tFTFldName & "','" & tTmnNum & "','" & tTrnNo & "','" & Format(dDateTrn, "YYYY/MM/DD") & "','" & tCardNo & "','" & tTokenCardNo & "','1','" & tMM_CNDbOnServer & "')"
                                   tSQL_Token = tSQL_Token & vbNewLine & "END"
                                   tSQL_Token = tSQL_Token & vbNewLine & "ELSE"
                                   tSQL_Token = tSQL_Token & vbNewLine & "BEGIN"
                                   tSQL_Token = tSQL_Token & vbNewLine & "UPDATE TPSTTokenLst WITH(ROWLOCK) SET FTToken='" & tTokenCardNo & "', FDDateIns=GETDATE()"
                                   tSQL_Token = tSQL_Token & vbNewLine & "WHERE FTTblName='" & ptTblName & "' AND FTFldName='" & tFTFldName & "' AND FTValue='" & tCardNo & "'"
                                   tSQL_Token = tSQL_Token & vbNewLine & "END"
                                   '====================================
                                   Call ocnOnline.Execute(tSQL_Token, nEff)
                               Else
                                  tValues = "'" & SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField)) & "'"
                                  W_DATbInsertData = False '*Tao 61-05-21 ถ้า Token ไม่ผ่าน ไม่ควร Transfer
                                  pnErrNo = 557 '*Tao 61-05-21 ถ้า Token ไม่ผ่าน ไม่ควร Transfer
                                  
                                  '[5.1803.3][61/03/23] *Bmp 61-03-28 Check web service
                                  'ถ้า token ไม่ผ่าน ไม่ต้องส่งไป POSSDB
                                  If UCase(ptTblName) = UCase("TPSTSalHD") Then           '*Ao 2014/08/21    V 4.0.10
                                       If W_PRCbCheckInsertHD(tTmnNum, tTrnNo, dDateTrn) Then
                                           W_DATbInsertData = False
                                           pnErrNo = 555
                                           'Exit Function
                                       End If
                                    End If
                                    Exit Function '*Tao 61-05-21 ถ้า Token ไม่ผ่าน ไม่ควร Transfer
                                   '==============================================
                               End If
                            Else
                               tValues = "'" & SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField)) & "'"
                            End If '===If bNeedToken Then
'                    Else
'                        tValues = "'" & SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField)) & "'"
'                    End If '==If Trim(tCardNo) <> "" Then
                   End If
                '---------------------------------------------------------------------
'                    tValues = "'" & SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField)) & "'"
                Case "C", "N"
                    tValues = SP_FEDvChkNumeric(porsTemp, tField)
                Case "D"
                    If UCase(tField) = UCase("FDCVPickupdate") Then '*Bump 56-06-19 กรณี CV ใบใหม่  จะไม่มีวันที่ ของ  FDCVPickupdate CommSheet TK-ISS3000-234-ST-00.00.03.xls CV Offline-->Online
                        tValues = "NULL"
                    Else
                        tValues = "'" & Format(SP_FEDdChkDate(porsTemp, tField), "YYYY/MM/DD") & "'"
                    End If
            End Select
            tSumValues = IIf(tSumValues = "", "", tSumValues & ",") & tValues
        Next nI
        tSql = "INSERT INTO " & ptTblName
        tSql = tSql & vbCrLf & "(" & tSumField & ")"
        tSql = tSql & vbCrLf & "VALUES(" & tSumValues & ")"
        Call ocnOnline.Execute(tSql)
        W_DATbInsertData = True
        If Not (orsTemp Is Nothing) Then
            If orsTemp.State = adStateOpen Then orsTemp.Close
            Set orsTemp = Nothing
        End If
    End With
    
    If UCase(ptTblName) = UCase("TPSTSalHD") Then           '*Ao 2014/08/21    V 4.0.10
        If W_PRCbCheckInsertHD(tTmnNum, tTrnNo, dDateTrn) Then
            W_DATbInsertData = False
            pnErrNo = 556
            Exit Function
        End If
    End If
    Exit Function
Err_Handler:
    pnErrNo = Err.Number
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " " & ("Err_Handler:" & Err.Number & "," & Err.Description)
    If Not (orsTemp Is Nothing) Then
        If orsTemp.State = adStateOpen Then orsTemp.Close
        Set orsTemp = Nothing
    End If
    If UCase(ptTblName) = UCase("TPSTSalHD") Then           '*Ao 2014/08/21    V 4.0.10
        If W_PRCbCheckInsertHD(tTmnNum, tTrnNo, dDateTrn) Then
            W_DATbInsertData = False
            pnErrNo = 556
            Exit Function
        End If
    End If
End Function

Private Function W_DATbUpdateData(ByVal porsTemp As ADODB.Recordset, ByVal ptTblName As String, ByVal ptPK As String) As Boolean
    Dim tSql As String
    Dim tField As String
    Dim tSumField As String
    Dim tValues As String
    Dim tSumValues As String
    Dim nI As Integer
    
    Dim tTemp As String
    Dim tTempValue As String
    Dim tPKString As String
    Dim tPKValues As String
    On Error GoTo Err_Handler
    tTemp = cUT.UT_STRtToken(ptPK, ";", 1, False)
    Do While tTemp <> ""
        Select Case Mid(tTemp, 2, 1)
            Case "T"
                tTempValue = "'" & SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tTemp)) & "'"
            Case "C", "N"
                tTempValue = SP_FEDvChkNumeric(porsTemp, tTemp)
            Case "D"
                tTempValue = "'" & Format(SP_FEDdChkDate(porsTemp, tTemp), "YYYY/MM/DD") & "'"
        End Select
        tPKString = IIf(tPKString = "", "", tPKString & " AND ") & tTemp & " = " & tTempValue
        tTemp = cUT.UT_STRtToken(ptPK, ";", 1, False)
    Loop
    
    With porsTemp
        For nI = 0 To .Fields.Count - 1
            tField = .Fields(nI).Name
            Select Case Mid(tField, 2, 1)
                Case "T"
                    tValues = "'" & SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField)) & "'"  '*KT 53-02-10 Communication Sheet PH3-ML-SF-079 Can't read Mall card which customer name or surname included character'.xls
                Case "C", "N"
                    tValues = SP_FEDvChkNumeric(porsTemp, tField)
                Case "D"
                    tValues = "'" & Format(SP_FEDdChkDate(porsTemp, tField), "YYYY/MM/DD") & "'"
            End Select
            tSumField = IIf(tSumField = "", "", tSumField & ",") & tField & " = " & tValues
        Next nI
        If SP_FEDtChkString(porsTemp, "FTStaSentOnOff") = "3" Then
            tSql = "UPDATE " & ptTblName & " WITH(ROWLOCK)"
            tSql = tSql & vbCrLf & "SET " & tSumField
            tSql = tSql & vbCrLf & "WHERE " & tPKString
            Call ocnOnline.Execute(tSql)
            W_DATbUpdateData = True
        End If
        
    End With
    Exit Function
Err_Handler:
End Function

Private Sub W_DATxUpdateByScript(ByVal porsTemp As ADODB.Recordset, ByVal ptTblName As String, ByVal ptPK As String, ByVal ptUpdScript As String, ByRef ptPKString As String)
    Dim tSql As String
    Dim tTemp As String
    Dim tTempValue As String
    Dim tPKString As String
On Error GoTo Err_Handler
    tTemp = cUT.UT_STRtToken(ptPK, ";", 1, False)
    Do While tTemp <> ""
        Select Case Mid(tTemp, 2, 1)
            Case "T"
                tTempValue = "'" & SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tTemp)) & "'"
            Case "C", "N"
                tTempValue = SP_FEDvChkNumeric(porsTemp, tTemp)
            Case "D"
                tTempValue = "'" & Format(SP_FEDdChkDate(porsTemp, tTemp), "YYYY/MM/DD") & "'"
        End Select
'        tPKString = IIf(tPKString = "", "", tPKString & " AND ") & tTemp & " = " & tTempValue              '*Ao 2014/03/25    มีบ้าง Field เป็นค่า ISNULL
        tPKString = IIf(tPKString = "", "", tPKString & " AND ") & "ISNULL(" & tTemp & ",'') = " & tTempValue
        tTemp = cUT.UT_STRtToken(ptPK, ";", 1, False)
    Loop

    tSql = "UPDATE " & ptTblName & " WITH(ROWLOCK)"
    tSql = tSql & vbCrLf & "SET " & ptUpdScript
    tSql = tSql & vbCrLf & "WHERE " & tPKString
    Call ocnOffline.Execute(tSql)
    ptPKString = tPKString
    Exit Sub
Err_Handler:
    If Err.Number = -2147217871 Then 'Time out
        'Reindex Table Sale and Temp
        tSql = "DBCC DBREINDEX('" & ptTblName & "',' ',90)"
        Call ocnOffline.Execute(tSql)
        'Unlock Table
        tSql = "SELECT * FROM " & ptTblName & " WITH (NOLOCK)"
        Call ocnOffline.Execute(tSql)
        'Re Execute
        tSql = "UPDATE " & ptTblName
        tSql = tSql & vbCrLf & "SET " & ptUpdScript
        tSql = tSql & vbCrLf & "WHERE " & tPKString
        Call ocnOffline.Execute(tSql)
        ptPKString = tPKString
        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & ptTblName & " TerminalNum : " & SP_FEDtChkString(porsTemp, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(porsTemp, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(porsTemp, "FTShdTransType") & " ReUpdate complete"
    End If
End Sub

Private Sub W_GETxPKString(ByVal porsTemp As ADODB.Recordset, ByVal ptPK As String, ByRef ptPKString As String)
    Dim tTemp As String
    Dim tTempValue As String
    Dim tPKString As String
    On Error GoTo Err_Handler
    tTemp = cUT.UT_STRtToken(ptPK, ";", 1, False)
    Do While tTemp <> ""
        Select Case Mid(tTemp, 2, 1)
            Case "T"
                tTempValue = "'" & SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tTemp)) & "'"
            Case "C", "N"
                tTempValue = SP_FEDvChkNumeric(porsTemp, tTemp)
            Case "D"
                tTempValue = "'" & Format(SP_FEDdChkDate(porsTemp, tTemp), "YYYY/MM/DD") & "'"
        End Select
        tPKString = IIf(tPKString = "", "", tPKString & " AND ") & tTemp & " = " & tTempValue
        tTemp = cUT.UT_STRtToken(ptPK, ";", 1, False)
    Loop
    ptPKString = tPKString
    Exit Sub
Err_Handler:
    ptPKString = ""
End Sub

Private Sub Form_Load()
    wMain.Hide
    If App.PrevInstance Then
        End
    End If
    If MM_RECbAdaIniRead Then
        Call SaveSetting("AdaPosMall", "POSFront", "SentStartDate", Format(Date, "yyyy/MM/dd"))
        Call SaveSetting("AdaPosMall", "POSFront", "SentStartTime", Format(Now, "HH:MM:SS"))
        Call SaveSetting("AdaPosMall", "POSFront", "SentStaOn", "1")
    Else
        Call SaveSetting("AdaPosMall", "POSFront", "SentStartDate", Format(Date, "yyyy/MM/dd"))
        Call SaveSetting("AdaPosMall", "POSFront", "SentStartTime", Format(Now, "HH:MM:SS"))
        Call SaveSetting("AdaPosMall", "POSFront", "SentStaOn", "0")
        End
    End If
    tW_GetStaSent = GetSetting("AdaPosMall", "POSFront", "SentSale")
    If tW_GetStaSent = "0" Then
        Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")
    End If
    Call W_DATbConnect(ocnOnline, tMM_CNDbOnServer, tMM_CNDbOnUser, tMM_CNDbOnPassword, tMM_CNDbOnDB)
    Call W_DATbConnect(ocnOffline, tMM_CNDbOffServer, tMM_CNDbOffUser, tMM_CNDbOffPassword, tMM_CNDbOffDB)
    '====================================
'    Call W_CRTxClearLog '*KT 52-06-10 ไม่ต้องใช้แล้วมี ServiceAutoClear ทำงานแทนแล้ว
    Call W_CRTxLogTran
    nVB_CNBase = nEN_Dbs2SQL        '*Ao 2013/03/18  TK
    '*Bmp 61-02-27
    Call SP_GETxVariableAlw
    Call SP_SETxObjects
    otmForm.Enabled = True '*Tao 61-05-09 Start Timer
End Sub

Private Sub otmForm_Timer()
    '-------------------------------------------------
    'Cmt : *KT 52-11-18 เพิ่มให้การเก็บ Log นั้นแสดงหมายเลขประเภท Transaction และ Version ของ EXE
    '------------------------------------------------
    Dim oFSO As New FileSystemObject
'    Dim oTxtFile As Object
    Dim oFile As File
    Dim orsTbl As ADODB.Recordset
    Dim orsData As ADODB.Recordset
    Dim tTblName As String
    Dim tCond As String
    Dim tPK As String
    Dim nGetErr As Long
    Dim bStaInsert As Boolean
    Dim tUpdScript As String
    Dim tTmnNum As String
    Dim dSaleDate As Date
    Dim bPassTxt As Boolean
    Dim bFirst As Boolean
    Dim bLast As Boolean
    Dim tPKString As String
    Dim tStaDoc As String
    Dim nErr As Long
    Dim bChk As Boolean
    Dim tTokenCardNo As String
    Dim tCardNo As String
On Error GoTo Err_Handler

    otmForm.Enabled = False '*Tao 61-05-09 Stop Timer
    
    '*KT 52-01-12 เก็บ log Transfer
    '--------------------------------------------------------------------------
    bPassTxt = False
    If Not SP_FLEbManFile(tVB_FileLog, nEN_Fle1Exist) Then
            oFSO.CreateTextFile (tVB_FileLog)
            Set oFile = oFSO.GetFile(tVB_FileLog)
            Set oVB_LogFile = oFile.OpenAsTextStream(ForWriting)
            oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Create : " & " Version " & App.Major & "." & App.Minor & "." & App.Revision
    End If
    bPassTxt = True
    '-----------------------------------------------------------------------------
    bFirst = True
    bLast = False
    'อ่าน Regetry ของ ServiceTransfer
    tW_GetStaSent = GetSetting("AdaPosMall", "POSFront", "SentSale")
    tW_GetDate = GetSetting("AdaPosMall", "POSFront", "SentStartDate")
    tW_GetTime = GetSetting("AdaPosMall", "POSFront", "SentStartTime")
    otbForm.Text = IIf(tW_GetStaSent = "1", "กำลังทำงานอยู่", "ยังไม่ทำงาน") & vbCrLf & tW_GetDate & "   " & tW_GetTime
    If tW_GetStaSent = "1" Then
        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " SentSale = " & tW_GetStaSent    '*Tao 61-05-09 เก็บ Log เพิ่ม
        tTokenCardNo = ""
        'tCardNo = "5444820011112222"
        '===== Test Web service ก่อน
        'If SP_PRCbToken("Check connection web service", tCardNo, tTokenCardNo) Then
        'If SP_CHKbWebservice() Then '*Bmp 61-04-11 ใช้แทนข้างบน '*Tao 61-09-04 ปิดไว้ก่อน ให้ Transfer ได้ ยกเว้นบิลที่มีสามชิก หรือ ใช้บัตรเครดิต
        If GetSetting("AdaPosMall", "POSFront", "Online") = "1" Then '*Tao 61-09-04 ใช้เช็ค Service Online ก่อน จะได้ไม่ต้องรอนาน
            'Connect DB Online
            If W_DATbConnectOrChk(ocnOnline, tMM_CNDbOnServer, tMM_CNDbOnUser, tMM_CNDbOnPassword, tMM_CNDbOnDB, True, nErr) Then
                '*KT 51-10-01 Store Close ต้องไม่มีการ Transafer
                If W_GETbStoreOpen Then
                    'Connect DB Offline
                    If W_DATbConnectOrChk(ocnOffline, tMM_CNDbOffServer, tMM_CNDbOffUser, tMM_CNDbOffPassword, tMM_CNDbOffDB, False, nErr) Then
                        If Not orsTbl Is Nothing Then
                                If orsTbl.State = adStateOpen Then orsTbl.Close
                        End If
                        'อ่านข้อมูล Table ที่ต้องทำการ Sync จาก Table TSysSync
                        Call MM_GETxTable(orsTbl) 'DB Online *KT 51-11-06
                        'อ่านข้อมูล Terminal จาก Computer name
                        tTmnNum = W_GETtTmnNum()
                        dSaleDate = SP_GETdSaleDate 'อ่านวันที่ Sale Date จาก DB Online
                        tVB_CNTmnNum = tTmnNum
                        If Not (orsTbl.BOF And orsTbl.EOF) Then
                            Do While Not orsTbl.EOF
                                tTblName = SP_FEDtChkString(orsTbl, "FTSscTable")
                                '*KT 51-10-01
                                '-----------------------------------------
                                Select Case UCase(tTblName)
                                    Case "TPSHD", "TPSDT", "TPSRC", "TPSCD"
                                        tTblName = tTblName & tTmnNum
                                End Select
                                '-----------------------------------------
                                tCond = SP_FEDtChkString(orsTbl, "FTSscCondScript")
                                'ตรวจสอบเงื่อนไขวันที่ขายด้วย เนื่องจากมีการขายข้ามวันเลยต้องมี <= หรือ กรณีที่เปิดขายก่อนหน้าก่อนที่ Store จะ Open การขายก่อน SaleDate ใน transaction มีการบวก FDShdTransDate ไปแล้ว
                                '_______________________________________________________
                                If UCase(tTblName) = UCase("TSysLogTableUpdate") Then               '*Ao 2013/01/09     TK
                                    tCond = Replace(tCond, "~", "'")
                                ElseIf UCase(tTblName) = UCase("TPSTCV") Then               '*Bump 56-06-18 CommSheet TK-ISS3000-234-ST-00.00.03.xls CV Offline-->Online
                                    tCond = Replace(tCond, "~", "'")
                                Else
                                    tCond = Replace(tCond, "~", "'") & " AND FDShdTransDate <= '" & Format(dSaleDate, "yyyy/mm/dd") & "'"
                                End If
    '                            tCond = Replace(tCond, "~", "'") & " AND FDShdTransDate <= '" & Format(dSaleDate, "yyyy/mm/dd") & "'"
                                '_______________________________________________________
                                tPK = SP_FEDtChkString(orsTbl, "FTSscFedPK")
                                tPK = Replace(tPK, "~", "'")
                                tUpdScript = SP_FEDtChkString(orsTbl, "FTSscUpdScript")
                                tUpdScript = Replace(tUpdScript, "~", "'")
                                If Not orsData Is Nothing Then
                                    If orsData.State = adStateOpen Then orsData.Close
                                End If
                                'อ่านข้อมูลจาก Table ที่จะ Sync
                                If MM_GETxData(tTblName, tCond, orsData, tPK) Then
                                    If Not (orsData.BOF And orsData.EOF) Then
                                        Do While Not orsData.EOF
                                            'ตรวจสอบข้อมุลว่าเคยส่งหรือไม่จาก Table HD
                                            '*KT 51-10-02 1.0.2
        '                                    If SP_FEDtChkString(orsData, "FTShdTransNo") = "45612" And SP_FEDvChkNumeric(orsData, "FNSdtSeqNo") = 5 Then Stop
                                            '__________________________________________________________________
    '                                        If MM_GETxDataComplate(SP_FEDtChkString(orsData, "FTTmnNum"), SP_FEDtChkString(orsData, "FTShdTransNo"), SP_FEDdChkDate(orsData, "FDShdTransDate"), tTmnNum, tStaDoc) Then
                                            If UCase(tTblName) = UCase("TSysLogTableUpdate") Then               '*Ao 2013/01/09     TK
                                                bChk = True
                                            Else
    '                                            bChk = MM_GETxDataComplate(SP_FEDtChkString(orsData, "FTTmnNum"), SP_FEDtChkString(orsData, "FTShdTransNo"), SP_FEDdChkDate(orsData, "FDShdTransDate"), tTmnNum, tStaDoc)
                                                bChk = True                 '*Ao 2013/03/14   ส่งข้อมูลทุก CASE
                                            End If
                                            '__________________________________________________________________
                                            If bChk Then
                                                If bFirst Then
                                                    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Start : "    '*KT 52-01-12 เก็บ log Transfer
                                                    bFirst = False
                                                    bLast = True
                                                End If
                                                bStaInsert = False
                                                'ส่งข้อมูลไปที่ DB Online
                                                nGetErr = 0
                                                If W_DATbInsertData(orsData, tTblName, nGetErr) Then
                                                    bStaInsert = True
                                                Else
                                                    If nGetErr = -2147467259 Then
                                                        Set orsTbl = Nothing
                                                        Set orsData = Nothing
                                                        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Error Number  : -2147467259"    '*KT 52-01-12 เก็บ log Transfer
                                                        'Exit Sub
                                                        GoTo StartTimer '*Tao 61-05-09 Start Timer
                                                    End If
        '*KT 51-11-06 ไม่มีการ Update ข้อมูลที่ Update ไปแล้วเพราะอาจทำให้ข้อมูลฝั่ง Server มีข้อมูลที่ผิด เพราะมีโอกาศที่ฝั่ง Local จะมีการยกข้อมูลจาก File Backup(.dat) ซึ่ง Filed  FTStaSentOnOff จะไม่เป็น 1 ทั้งที่บิลนั้นทำที่ Online แล้ว
        'แก้ไขปัจจุบันให้ FTStaSentOnOff = '3' หมายความว่าทาง Local ต้องการให้ทำซ้ำ
        '------------------------------------------------------------------------------------------------
                                                    If nGetErr = -2147217873 Then
                                                        If W_DATbUpdateData(orsData, tTblName, tPK) Then
                                                            bStaInsert = True
                                                        End If
                                                    End If
        '------------------------------------------------------------------------------------------------
                                                End If
                                                If bStaInsert Then
                                                    'set สถานะให้ข้อมูลนั้นส่งข้อมูลแล้ว
                                                    tPKString = ""
                                                    Call W_DATxUpdateByScript(orsData, tTblName, tPK, tUpdScript, tPKString)
                                                    '*Ao 2013/01/10  TK
    '                                                oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString & " is complete" '*KT 52-01-12 เก็บ log Transfer
                                                    If UCase(tTblName) = UCase("TSysLogTableUpdate") Then               '*Ao 2013/01/09     TK
                                                        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " FDSltDate : " & SP_FEDtChkString(orsData, "FDSltDate") & " FTSltTime : " & SP_FEDtChkString(orsData, "FTSltTime") & " FTSltAct : " & SP_FEDtChkString(orsData, "FTSltAct") & " PK : " & tPKString & " is complete"
                                                    ElseIf UCase(tTblName) = UCase("TPSTCV") Then '*Bump 56-06-18 CommSheet TK-ISS3000-234-ST-00.00.03.xls CV Offline-->Online
                                                        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " FDDateIns : " & SP_FEDtChkString(orsData, "FDDateIns") & " FTTimeIns : " & SP_FEDtChkString(orsData, "FTTimeIns") & " FTCVNum : " & SP_FEDtChkString(orsData, "FTCVNum") & " PK : " & tPKString & " is complete"
                                                    Else
                                                        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString & " is complete"    '*KT 52-01-12 เก็บ log Transfer
                                                    End If
                                                    '_______________________________________________________
                                                Else
                                                    '*52-09-15 แก้ไขให้กรณีที่มีข้อมูลที่ Server แล้วไม่สามารถ Insert ได้ ต้องทำการปรับปรุงข้อมุลฝั่ง Local ให้ Flag ส่งแล้ว
                                                    If nGetErr = -2147217873 Then
                                                        'set สถานะให้ข้อมูลนั้นส่งข้อมูลแล้ว
                                                        tPKString = ""
                                                        Call W_DATxUpdateByScript(orsData, tTblName, tPK, tUpdScript, tPKString)
    '                                                    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString & " is complete" '*KT 52-01-12 เก็บ log Transfer
                                                        If UCase(tTblName) = UCase("TSysLogTableUpdate") Then               '*Ao 2013/01/09     TK
                                                            oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " FDSltDate : " & SP_FEDtChkString(orsData, "FDSltDate") & " FTSltTime : " & SP_FEDtChkString(orsData, "FTSltTime") & " FTSltAct : " & SP_FEDtChkString(orsData, "FTSltAct") & " PK : " & tPKString & " is complete"
                                                        ElseIf UCase(tTblName) = UCase("TPSTCV") Then    '*Bump 56-06-18 CommSheet TK-ISS3000-234-ST-00.00.03.xls CV Offline-->Online
                                                            oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " FDDateIns : " & SP_FEDtChkString(orsData, "FDDateIns") & " FTTimeIns : " & SP_FEDtChkString(orsData, "FTTimeIns") & " FTCVNum : " & SP_FEDtChkString(orsData, "FTCVNum") & " PK : " & tPKString & " is complete"
                                                        Else
                                                            oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString & " is complete"    '*KT 52-01-12 เก็บ log Transfer
                                                        End If
                                                    '_______________________________________________________
                                                    Else
                                                        tPKString = ""
                                                        Call W_GETxPKString(orsData, tPK, tPKString)
    '                                                    oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString & " is not complete" '*KT 52-01-12 เก็บ log Transfer
                                                        If UCase(tTblName) = UCase("TSysLogTableUpdate") Then               '*Ao 2013/01/09     TK
                                                            oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " FDSltDate : " & SP_FEDtChkString(orsData, "FDSltDate") & " FTSltTime : " & SP_FEDtChkString(orsData, "FTSltTime") & " FTSltAct : " & SP_FEDtChkString(orsData, "FTSltAct") & " PK : " & tPKString & " is complete"
                                                        ElseIf UCase(tTblName) = UCase("TPSTCV") Then    '*Bump 56-06-18 CommSheet TK-ISS3000-234-ST-00.00.03.xls CV Offline-->Online
                                                            oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " FDDateIns : " & SP_FEDtChkString(orsData, "FDDateIns") & " FTTimeIns : " & SP_FEDtChkString(orsData, "FTTimeIns") & " FTCVNum : " & SP_FEDtChkString(orsData, "FTCVNum") & " PK : " & tPKString & " is complete"
                                                        Else
    '                                                        If nGetErr = 555 Then
    '                                                            oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString
    '                                                            oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "is not complete because not send TPSTSalDT,TPSTSalRC,TPSTSalCD,TPSTSalePoint"
    '                                                        Else
    '                                                            oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString & " is complete" '*KT 52-01-12 เก็บ log Transfer
    '                                                        End If
                                                            Select Case nGetErr
                                                                Case 555
                                                                    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString
                                                                    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " is not complete because not send TPSTSalDT,TPSTSalRC,TPSTSalCD,TPSTSalePoint"
                                                                Case 556
                                                                    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString
                                                                    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " is not complete because not send TPSTSalHD Clean TPSTSalDT,TPSTSalRC,TPSTSalCD.FTStaSentOnOff = 0"
                                                                Case Else
                                                                    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString & " is complete"    '*KT 52-01-12 เก็บ log Transfer
                                                            End Select
                                                            
    '                                                        oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString & " is complete" '*KT 52-01-12 เก็บ log Transfer
                                                        End If
                                                    '_______________________________________________________
                                                    End If
                                                End If
                                            Else
                                                'set สถานะให้ข้อมูลนั้นส่งข้อมูลแล้ว
                                                If tStaDoc = "1" Then
                                                    tPKString = ""
                                                    Call W_DATxUpdateByScript(orsData, tTblName, tPK, tUpdScript, tPKString)
    '                                                oVB_LogFile.WriteLine  Format(Date, "yyyy/MM/dd") & " " & Time &   "Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString & " is online" '*KT 52-01-12 เก็บ log Transfer
                                                    If UCase(tTblName) = UCase("TSysLogTableUpdate") Then               '*Ao 2013/01/09     TK
                                                    ElseIf UCase(tTblName) = UCase("TPSTCV") Then               '*Bump 56-06-18 CommSheet TK-ISS3000-234-ST-00.00.03.xls CV Offline-->Online
                                                        'ไม่มีการทำงาน
                                                    Else
                                                        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Table : " & tTblName & " TerminalNum : " & SP_FEDtChkString(orsData, "FTTmnNum") & " TransNo : " & SP_FEDtChkString(orsData, "FTShdTransNo") & " TransType : " & SP_FEDtChkString(orsData, "FTShdTransType") & " PK : " & tPKString & " is online"    '*KT 52-01-12 เก็บ log Transfer
                                                    End If
                                                    '_______________________________________________________
                                                End If
                                            End If
                                            orsData.MoveNext
                                        Loop
                                    End If
                                Else
                                    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "0")
                                    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Can not select table " & tTblName & " on database local : "    '*KT 52-10-21 ปรับปรุง
                                End If
                                orsTbl.MoveNext
                            Loop
                        End If
                        Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "0")
                    Else '*KT 53-02-03 เก็บ Log ผิด
                        Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "0")
                        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Can not connect database local[Error Number " & nErr & "] : "    '*KT 52-10-21 ปรับปรุง
                    End If
                Else
                    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "0")
                    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Store close: "    '*KT 52-10-21 ปรับปรุง
                End If
            Else
                Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "0")
                oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Can not connect database server[Error Number " & nErr & "] : "    '*KT 52-10-21 ปรับปรุง
            End If
            '*Ao 2013/03/18  TK
            Call W_UPDxUpdatePoint
        Else
            Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "0") '*Tao 61-05-23 ให้ POSFront นำงานต่อได้
            'oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Tokenize Web Service not connect: " '*Tao 61-05-22
            oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Database server not connect: " '*Tao 61-09-04 เช็คจาก Service OnOff
        End If '===== Test Web service ก่อน
    End If '===== SentSale
    Set orsTbl = Nothing
    Set orsData = Nothing
    'set flag ใน Regetry ว่าการทำการ Transfer Success
    Call SaveSetting("AdaPosMall", "POSFront", "SentStartDate", Format(Date, "yyyy/MM/dd"))
    Call SaveSetting("AdaPosMall", "POSFront", "SentStartTime", Format(Now, "HH:MM:SS"))
    Call SaveSetting("AdaPosMall", "POSFront", "SentStaOn", "1")
    If bLast Then oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " End : "    '*KT 52-01-12 เก็บ log Transfer
    otmForm.Enabled = True '*Tao 61-05-09 Start Timer
    Exit Sub
Err_Handler:
    'Call MM_MSGnShowing(tMS_001, vbExclamation)
    If bPassTxt Then oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Error Number : " & Err.Number    '*KT 52-01-12 เก็บ log Transfer
StartTimer:
    otmForm.Enabled = True '*Tao 61-05-09 Start Timer
End Sub

Private Sub W_UPDxUpdatePoint()
'----------------------------------------------------
'   Call:
'   Cmt:    '*Ao 2013/03/18
'----------------------------------------------------
Dim ocnMem As ADODB.Connection
Dim orsPoint As ADODB.Recordset
Dim orsTemp As ADODB.Recordset
Dim orsMem  As ADODB.Recordset
Dim orsMemEx  As ADODB.Recordset
Dim tFTTmnNum As String, tFTShdTransNo As String, tFTShdTransType As String, tFDShdTransDate As String, tFTShdTransTime As String
Dim tSql  As String, tCardCode As String, tExpDate As String, tChkExp As String, tMcdExpDate As String
Dim nErr As Long, nEff As Long
Dim nCardPoint As Integer
Dim cTotalPoint  As Double, cBalance As Double
Dim bChk As Boolean
Dim bChkToken As Boolean
Dim dDate As Date
Dim tCardNo, tdbValue As String
Dim tTokenCardNo As String
Dim tSQLex As String

Dim tOrgExpDate As String '*Tao 61-05-30 เก็บค่า Original ExpDate
Dim tOrgChkExp As String '*Tao 61-05-30 เก็บค่า Original ChkExp

On Error GoTo ErrHandle
    tTokenCardNo = ""
    'tCardNo = "5444820011112222"
    '===== Test Web service ก่อน
    'If SP_PRCbToken("Check connection web service)", tCardNo, tTokenCardNo) Then '*Tao 61-05-17 ปิดไว้
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Start Update Member Point" '*Tao 61-05-09 เก็บ Log เพิ่ม
    If GetSetting("AdaPosMall", "POSFront", "Online") = "1" Then '*Tao 61-09-04 ใช้เช็ค Service Online ก่อน จะได้ไม่ต้องรอนาน
        If W_DATbConnectOrChk(ocnMem, tMM_CNDbMServer, tMM_CNDbMUser, tMM_CNDbMPassword, tMM_CNDbMDB, True, nErr) Then
            If W_DATbConnectOrChk(ocnOffline, tMM_CNDbOffServer, tMM_CNDbOffUser, tMM_CNDbOffPassword, tMM_CNDbOffDB, False, nErr) Then
                tSql = "SELECT DISTINCT FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate FROM TPSTSalePoint WITH(NOLOCK)"
                tSql = tSql & vbCrLf & "WHERE ISNULL(FTRemark,'') = '0'"
                tSql = tSql & vbCrLf & "AND (ISNULL(FTSpoMemID,'') <> '')"
                tSql = tSql & vbCrLf & "ORDER BY FTShdTransNo"
                Set orsPoint = ocnOffline.Execute(tSql)
                If Not (orsPoint.BOF And orsPoint.EOF) Then
                    Do While Not orsPoint.EOF
                        tFTTmnNum = SP_FEDtChkString(orsPoint, "FTTmnNum")
                        tFTShdTransNo = SP_FEDtChkString(orsPoint, "FTShdTransNo")
                        tFTShdTransType = SP_FEDtChkString(orsPoint, "FTShdTransType")
                        tFDShdTransDate = SP_DTEtStrSTD(SP_FEDdChkDate(orsPoint, "FDShdTransDate"))
                        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Update Member Point DocNo: " & tFTShdTransNo '*Tao 61-05-09 เก็บ Log เพิ่ม
                        tSql = "SELECT SUM(FCSpoPoint) AS FCSpoPoint,FTSpoMemID,CONVERT(VARCHAR,FDDateUpd,012) AS FDDateUpd,FTWhoUpd FROM TPSTSalePoint WITH(NOLOCK)"
                        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tFTTmnNum & "'"
                        tSql = tSql & vbCrLf & "AND FTShdTransNo= '" & tFTShdTransNo & "'"
                        tSql = tSql & vbCrLf & "AND FTShdTransType= '" & tFTShdTransType & "'"
                        tSql = tSql & vbCrLf & "AND FDShdTransDate= '" & tFDShdTransDate & "'"
                        tSql = tSql & vbCrLf & "AND ISNULL(FTRemark,'') = '0'" '*KT 61-05-28
                        tSql = tSql & vbCrLf & "GROUP BY FTSpoMemID,FDDateUpd,FTWhoUpd"
                        Set orsTemp = ocnOffline.Execute(tSql, nEff)
                        If Not (orsTemp.BOF And orsTemp.EOF) Then
                            orsTemp.MoveFirst
                            Do While Not orsTemp.EOF
                                tCardCode = SP_FEDtChkString(orsTemp, "FTSpoMemID")
                                nCardPoint = SP_FEDvChkNumeric(orsTemp, "FCSpoPoint")
                                tOrgExpDate = SP_FEDtChkString(orsTemp, "FDDateUpd") '*Tao 61-05-30 ควรเก็บ Original ExpDate ไว้ด้วย (Service Transfer จำเป็นต้องใช้ ในการแยกบัตร ว่าเป็นบัตรสมาชิกคนละใบกัน)
                                tOrgChkExp = SP_FEDtChkString(orsTemp, "FTWhoUpd")  '*Tao 61-05-30 ควรเก็บ Check ExpDate ไว้เป็นเงื่อนไขในตอน Update Remark='1'
                                oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Update Member Point MemberID: " & tCardCode & " , nCardPoint=" & nCardPoint '*Tao 61-05-09 เก็บ Log เพิ่ม
                                
                                '==================================================================
                                '=== *Bmp 61-02-28 Token บัตรเครดิต
                                tCardNo = ""
                                tCardNo = tCardCode
                                tTokenCardNo = ""
                                tdbValue = ""
                                bChkToken = False
                                '*Bmp 61-04-04 ==== เช็คบน Member ก่อน
'                                tSQLex = "SELECT TOP 1 FTToken FROM [TPSTTokenLst] WITH(NOLOCK)"
'                                tSQLex = tSQLex & " WHERE FTValue = '" & tCardCode & "'"

                                '*Tao 61-05-21 ของเดิม
                                'tSQLex = "SELECT TOP 1 FTValue FROM [TPSTTokenLst] WITH(NOLOCK)"
                                'tSQLex = tSQLex & " WHERE FTToken = '" & tCardCode & "'"
                                                          
                                 '*Tao 61-05-28 ปิดไว้ ไม่ต้องตรวจสอบใน TPSTTokenLst
'                                '*Tao 61-05-22 หาข้อมูลที่ Value ไม่เท่ากับ Token
'                                tSQLex = "SELECT TOP 1 FTValue,FTToken FROM TPSTTokenLst WITH(NOLOCK)"
'                                tSQLex = tSQLex & vbNewLine & "WHERE (ISNULL(FTValue,'')<>'' OR ISNULL(FTToken,'')<>'') AND (FTValue<>FTToken)"
'                                tSQLex = tSQLex & vbNewLine & "AND (FTToken='" & tCardCode & "')"
'                                tSQLex = tSQLex & vbNewLine & "ORDER BY A.FDDateIns DESC"
'
'                                '*หาว่าเคย token ไหม
'                                Set orsMemEx = ocnMem.Execute(tSQLex, nEff)
'                                If Not (orsMemEx.BOF And orsMemEx.EOF) Then
'                                    tdbValue = SP_FEDtChkString(orsMemEx, "FTValue") '** เคย token แล้ว
'                                    tCardNo = SP_FEDtChkString(orsMemEx, "FTToken") '** เคย token แล้ว มีใน POSSDB'*Tao 61-05-22  Token
'                                    If tdbValue <> "" Then oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " W_UPDxUpdatePoint(1) Value:" & tdbValue & ",Token:" & tCardCode
'                                End If
'                                '== ถ้าบน Member ไม่มี หาบน POSS DB '*Bmp 61-04-10
'                                If tdbValue = "" Then
'                                    Set orsMemEx = Nothing
'                                    Set orsTemp = ocnOnline.Execute(tSql, nEff)
'                                    If Not (orsTemp Is Nothing) Then
'                                        If Not (orsMemEx.BOF And orsMemEx.EOF) Then
'                                            tdbValue = SP_FEDtChkString(orsMemEx, "FTValue") '** เคย token แล้ว มีใน POSSDB
'                                            tCardNo = SP_FEDtChkString(orsMemEx, "FTToken") '** เคย token แล้ว มีใน POSSDB'*Tao 61-05-22  Token
'                                            If tdbValue <> "" Then oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " W_UPDxUpdatePoint(2) Value:" & tdbValue & ",Token:" & tCardCode
'                                        End If
'                                    End If
'                                End If
                                '== ถ้าบน Member ไม่มีค่อย token
                                If tdbValue = "" Then '* ยังไม่เคย token ให้ token '*Bmp 61-04-10
                                    'If SP_PRCbToken("TPSTSalePoint", tCardNo, tTokenCardNo) Then '*Tao 61-05-22 ของเดิม
                                    If SP_PRCbToken("TPSTSalePoint", tCardCode, tTokenCardNo, True) Then '*Tao 61-05-22 ไม่ต้องค้นหาในฐานข้อมุลก่อน
                                        tCardNo = tTokenCardNo
                                        bChkToken = True
                                        If tTokenCardNo <> "" Then oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " W_UPDxUpdatePoint(3) Token:" & tTokenCardNo
                                        '****************** เก็บลง Member DB '*Bmp 61-04-04
                                        '*Tao 61-05-21 ของเดิม
'                                        tSQLex = "IF NOT EXISTS(SELECT FTValue FROM TPSTTokenLst  WHERE FTTblName = 'TCNMMallCard' AND FTFldName = 'FTMcdCode' FTValue = '" & tCardNo & "')"
'                                        tSQLex = tSQLex & vbNewLine & "BEGIN"
'                                        tSQLex = tSQLex & vbNewLine & "INSERT INTO TPSTTokenLst(FDDateIns,FTWhoIns,FTTblName,FTFldName,FTMcdCode,FDMcdExpDate,FTValue,FTToken,FNStatus,FTSrvName)"
'                                        tSQLex = tSQLex & vbNewLine & "VALUES(GETDATE(),'ServiceTransfer','TCNMMallCard','FTMcdCode','" & tCardNo & "',GETDATE(),'" & tCardCode & "','" & tTokenCardNo & "','1','" & tMM_CNDbMServer & "')"
'                                        tSQLex = tSQLex & vbNewLine & "END"
                                        
                                        '*Tao 61-05-22 ถ้ามีแล้ว Update FTToken ด้วย (ให้ POSFront ใช้หาข้อมูล Token)
                                        tSQLex = "IF NOT EXISTS(SELECT FTValue FROM TPSTTokenLst WITH(NOLOCK) WHERE FTTblName = 'TCNMMallCard' AND FTFldName = 'FTMcdCode' FTValue = '" & tCardCode & "')"
                                        tSQLex = tSQLex & vbNewLine & "BEGIN"
                                        tSQLex = tSQLex & vbNewLine & "INSERT INTO TPSTTokenLst(FDDateIns,FTWhoIns,FTTblName,FTFldName,FTMcdCode,FDMcdExpDate,FTValue,FTToken,FNStatus,FTSrvName)"
                                        tSQLex = tSQLex & vbNewLine & "VALUES(GETDATE(),'ServiceTransfer','TCNMMallCard','FTMcdCode','" & tCardCode & "',GETDATE(),'" & tCardCode & "','" & tTokenCardNo & "','1','" & tMM_CNDbMServer & "')"
                                        tSQLex = tSQLex & vbNewLine & "END"
                                        tSQLex = tSQLex & vbNewLine & "ELSE"
                                        tSQLex = tSQLex & vbNewLine & "BEGIN"
                                        tSQLex = tSQLex & vbNewLine & "UPDATE TPSTTokenLst WITH(ROWLOCK) SET FTToken='" & tTokenCardNo & "', FDDateIns=GETDATE()"
                                        tSQLex = tSQLex & vbNewLine & "WHERE FTTblName='TCNMMallCard' AND FTFldName='FTMcdCode' AND FTValue='" & tCardCode & "'"
                                        tSQLex = tSQLex & vbNewLine & "END"
                                        '====================================
                                        Call ocnMem.Execute(tSQLex, nEff)
                                    Else
                                        bChkToken = False
                                    End If
                                Else
                                    bChkToken = True
                                    '*Tao 61-05-23 ตรวจสอบว่า ข้อมูลใน TPSTTokenLst ถูกต้องหรือไม่ (ไม่สนใจว่าจะ Token ได้หรือไม่ เพราะได้ข้อมูลมาแว้ว)
                                    If SP_PRCbToken("TPSTSalePoint", tdbValue, tTokenCardNo, True) Then '*Tao 61-05-23 ไม่ต้องค้นหาในฐานข้อมุลก่อน
                                        If (tdbValue <> tTokenCardNo) And (tTokenCardNo <> "") Then
                                            'ถ้าข้อมูลใน TPSTTokenLst ไม่ตรงกับ WebService ต้อง Update ข้อมูลให้ถูกต้องด้วย
                                            tCardNo = tTokenCardNo
                                            '*Tao 61-05-23 ถ้ามีแล้ว Update FTToken ด้วย (ให้ POSFront ใช้หาข้อมูล Token)
                                            tSQLex = "IF NOT EXISTS(SELECT FTValue FROM TPSTTokenLst WITH(NOLOCK) WHERE FTTblName = 'TCNMMallCard' AND FTFldName = 'FTMcdCode' FTValue = '" & tdbValue & "')"
                                            tSQLex = tSQLex & vbNewLine & "BEGIN"
                                            tSQLex = tSQLex & vbNewLine & "INSERT INTO TPSTTokenLst(FDDateIns,FTWhoIns,FTTblName,FTFldName,FTMcdCode,FDMcdExpDate,FTValue,FTToken,FNStatus,FTSrvName)"
                                            tSQLex = tSQLex & vbNewLine & "VALUES(GETDATE(),'ServiceTransfer','TCNMMallCard','FTMcdCode','" & tdbValue & "',GETDATE(),'" & tdbValue & "','" & tTokenCardNo & "','1','" & tMM_CNDbMServer & "')"
                                            tSQLex = tSQLex & vbNewLine & "END"
                                            tSQLex = tSQLex & vbNewLine & "ELSE"
                                            tSQLex = tSQLex & vbNewLine & "BEGIN"
                                            tSQLex = tSQLex & vbNewLine & "UPDATE TPSTTokenLst WITH(ROWLOCK) SET FTToken='" & tTokenCardNo & "', FDDateIns=GETDATE()"
                                            tSQLex = tSQLex & vbNewLine & "WHERE FTTblName='TCNMMallCard' AND FTFldName='FTMcdCode' AND FTValue='" & tdbValue & "'"
                                            tSQLex = tSQLex & vbNewLine & "END"
                                            Call ocnMem.Execute(tSQLex, nEff)
                                        End If
                                    End If
                                    '==============================================
                                End If
                                oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Transfer Member(TCNMMallCard) [Token Card No:" & tTokenCardNo & "]"
                                '==================================================================
                                tExpDate = SP_FEDtChkString(orsTemp, "FDDateUpd")
                                oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " tExpDate=" & tExpDate '*Tao 61-05-28 เพิ่ม Log
                                If tExpDate = "" Then
                                    tExpDate = Right(Format(Year(Date), "0000"), 2) & "12"              'Year & Month       *1312
                                Else
                                    tExpDate = Left(tExpDate, 4)
                                End If
                                tChkExp = SP_FEDtChkString(orsTemp, "FTWhoUpd")
                                oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " tChkExp=" & tChkExp '*Tao 61-05-28 เพิ่ม Log
    '                            If tChkExp = "" Then                        '*Ao 2014/05/07             *******
    '                                tChkExp = "N"
    '                            End If
                                
                                Select Case CInt(tFTShdTransType)             'IPV,Return จะเป็นการลบ
    '                                Case tEN_TrnReturn, tEN_TrnIPV, tEN_TrnDepositCancel, tEN_TrnIPVRedeem, tEN_TrnAddDepositCancel, tEN_TrnCancelVoucher
                                    Case tEN_TrnReturn, tEN_TrnIPV, tEN_TrnCancelVoucher, tEN_TrnDepositCancel, tEN_TrnIPVRedeem, tEN_TrnAddDepositCancel, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax      '*Ao 2013/10/23   เพิ่ม tEN_TrnReturnTaxInc,tEN_TrnReturnManualTax
                                        nCardPoint = (-1) * nCardPoint
                                    Case tEN_TrnIPVVIPService, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVOtherReceive, tEN_TrnOtherDisbursement
                                        nCardPoint = (-1) * nCardPoint
                                        'tEN_TrnOtherDisbursement       - POINT
                                        'tEN_TrnIPVOtherDisbursement +POINT
                                    Case tEN_TrnSale
                                        nCardPoint = nCardPoint
                                End Select
    '                            Select Case CInt(tFTShdTransType)                  '*Ao 2014/05/07             *******
    '                                Case tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnIPVRedeem, tEN_TrnAddDepositCancel, tEN_TrnRedeem
    '                                    tChkExp = "N"
    '                            End Select
                                '=====================================================
                                If bChkToken Then  '*Bmp 61-03-28 ถ้า token ผ่าน จึง Update
                                    If UCase(tChkExp) = UCase("Y") Then
                                        tSql = "SELECT TOP 1 * FROM TCNMMallCard WITH(NOLOCK) "
        '                                tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardCode & "'"
                                        tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardNo & "'" '=== *Bmp 61-02-28 Token
                                        tSql = tSql & vbCrLf & "AND SUBSTRING(CONVERT(VARCHAR,FDMcdExpDate,012),1,4) = '" & tExpDate & "'"
                                        tSql = tSql & vbCrLf & "ORDER BY FDMcdExpDate DESC"                             'อาจจะมีมากกว่า 1 Record
                                    Else
                                        tSql = "SELECT TOP 1 * FROM TCNMMallCard WITH(NOLOCK) "
        '                                tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardCode & "'"
                                        tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardNo & "'" '=== *Bmp 61-02-28 Token
                                        tSql = tSql & vbCrLf & "ORDER BY FDMcdExpDate DESC"                             'อาจจะมีมากกว่า 1 Record
                                    End If
                                    '=== Token(1)
                                    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Get Member:" & tSql
                                    Set orsMem = ocnMem.Execute(tSql, nEff)
                                    If Not (orsMem.BOF And orsMem.EOF) Then
                                        cTotalPoint = SP_FEDvChkNumeric(orsMem, "FCMcdTotalPoint")
                                        tMcdExpDate = SP_DTEtStrSTD(SP_FEDdChkDate(orsMem, "FDMcdExpDate"))           '***
                                        cBalance = SP_FEDvChkNumeric(orsMem, "FCBalance")
        '                                tCardNo = SP_STRtSiggleCode(SP_FEDtChkString(porsTemp, tField))
        '                                If SP_PRCbToken(ptTblName, tCardNo, tTokenCardNo) Then
        '                                    tValues = "'" & tTokenCardNo & "'"
        '                                Else
        '                                   tValues = "'" & tCardNo & "'"
        '                                End If
                                        tSql = "UPDATE TCNMMallCard WITH(ROWLOCK) "
                                        tSql = tSql & vbCrLf & "SET FCEarned = FCEarned + " & nCardPoint
                                        tSql = tSql & vbCrLf & ",FCBalance = FCMcdTotalPoint + FCEarned +" & nCardPoint
                                        tSql = tSql & vbCrLf & ",FDDateUpd = '" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "'"
                                        tSql = tSql & vbCrLf & ",FTTimeUpd = '" & Format(Time, "HH:MM:SS") & "'"
                                        tSql = tSql & vbCrLf & ",FTWhoUpd = 'ServiceTranfer'" '*Tao 61-05-25 เปลี่ยนเป็น ServiceTranfer
                                        'tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardCode & "'"
                                        tSql = tSql & vbCrLf & "WHERE FTMcdCode = '" & tCardNo & "'" '=== *Bmp 61-02-28 Token
                                        tSql = tSql & vbCrLf & "AND FDMcdExpDate = '" & tMcdExpDate & "'"
                                        '=== Token(2)
                                        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Update Member:" & tSql
                                        Call ocnMem.Execute(tSql, nEff)
                                        If nEff > 0 Then
                                            bChk = True
                                        Else
                                            bChk = False
                                        End If
                                    Else
        '                                dDate = DateSerial(Left(tExpDate, 2), Right(tExpDate, 2) + 1, 0)                   '*Ao 2013/10/25
                                        dDate = DateSerial(Left(Format(Year(Date), "0000"), 2) & Left(tExpDate, 2), Right(tExpDate, 2) + 1, 0)                     'tExpDate = YYMM = 1312
                                        tMcdExpDate = SP_DTEtStrSTD(dDate)
                                        tSql = "INSERT INTO TCNMMallCard(FTMcdStaAct,FTMcdCode,FTMcdName,FDMcdBirthday,FCMcdAmountPoint,FCMcdTotalPoint,FTMcdSta"
                                        tSql = tSql & vbCrLf & ",FDMcdStaDate,FDMcdExpDate,FCEarned,FCBalance"
                                        tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                                        'tSql = tSql & vbCrLf & "VALUES('A','" & tCardCode & "','New Member','" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "',0,0,'A'"
                                        tSql = tSql & vbCrLf & "VALUES('A','" & tCardNo & "','New Member','" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "',0,0,'A'" '=== *Bmp 61-02-28 Token
                    '                    tSql = tSql & vbCrLf & ",'" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "',DATEADD(year,1,'" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "')"
                                        tSql = tSql & vbCrLf & ",'" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "','" & tMcdExpDate & "'"
                                        tSql = tSql & vbCrLf & "," & nCardPoint & "," & nCardPoint
                                        tSql = tSql & vbCrLf & "," & cUT.UT_SQLtLastIns("ServiceTranfer")
                                        '=== Token(3)
                                        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Insert Member:" & tSql
                                        Call ocnMem.Execute(tSql, nEff)
                                        If nEff > 0 Then
                                            bChk = True
                                        Else
                                            bChk = False
                                        End If
                                    End If
                                    If bChk Then
        '                                tSql = "UPDATE TPSTSalHD"
        '                                tSql = tSql & vbCrLf & "SET FTShdUpload = '1'"
        '                                tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tFTTmnNum & "' AND FTShdTransNo='" & tFTShdTransNo & "'"
        '                                tSql = tSql & vbCrLf & "AND FTShdTransType='" & tFTShdTransType & "' AND FDShdTransDate='" & tFDShdTransDate & "'"
        '                                tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
        '                                Call ocnOffline.Execute(tSql, nEff)
                                        tSql = "UPDATE TPSTSalePoint WITH(ROWLOCK)" '*Tao 61-05-28 เพิ่ม WITH(ROWLOCK)
                                        tSql = tSql & vbCrLf & "SET FTRemark = '1'"
                                        'tSql = tSql & vbCrLf & ",FDDateIns = '" & tMcdExpDate & "'" '*Tao 61-05-30 ของเดิม
                                        tSql = tSql & vbCrLf & ",FDDateUpd = '" & tMcdExpDate & "'"  '*Tao 61-05-30
                                        tSql = tSql & vbCrLf & "WHERE  FTTmnNum='" & tFTTmnNum & "' AND FTShdTransNo='" & tFTShdTransNo & "'"
                                        tSql = tSql & vbCrLf & "AND FTShdTransType='" & tFTShdTransType & "' AND FDShdTransDate='" & tFDShdTransDate & "'"
                                        tSql = tSql & vbCrLf & "AND FTSpoMemID = '" & tCardCode & "'"
                                        tSql = tSql & vbCrLf & "AND ISNULL(FTRemark,'') = '0'" '*Tao 61-05-30 Update เฉพาะข้อมูลที่จบบิลแล้วเท่านั้น
                                        tSql = tSql & vbCrLf & "AND CONVERT(VARCHAR,FDDateUpd,012) = '" & tOrgExpDate & "'" '*Tao 61-05-30 Update แยกตามที่ GROUP BY
                                        tSql = tSql & vbCrLf & "AND FTWhoUpd = '" & tOrgChkExp & "'" '*Tao 61-05-30 Update แยกตามที่ GROUP BY
                                        Call ocnOffline.Execute(tSql, nEff)
                                        oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Update TPSTSalePoint:" & tSql
                                        
                                        tSql = "INSERT INTO TPSTBPHis(FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate,FTShdTransTime,FTShdMemberID,FCCstCardPoint ,FCTotalBP"
                                        tSql = tSql & vbCrLf & ",FCBalanceBP,FDDateUpd,FTTimeUpd)"
                                        tSql = tSql & vbCrLf & "VALUES('" & tFTTmnNum & "','" & tFTShdTransNo & "','" & tFTShdTransType & "','" & tFDShdTransDate & "','" & Format(Time, "HH:MM:SS") & "'"
                                        'tSql = tSql & vbCrLf & ",'" & tCardCode & "'," & nCardPoint & "," & cTotalPoint & "," & cBalance & ",'" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "','" & Format(Time, "HH:MM:SS") & "')"
                                        tSql = tSql & vbCrLf & ",'" & tCardNo & "'," & nCardPoint & "," & cTotalPoint & "," & cBalance & ",'" & SP_DTEtStrSTD(Format(Date, "dd/MM/yyyy")) & "','" & Format(Time, "HH:MM:SS") & "')" '*Bmp 61-03-23 Token tCardNo
                                        '=== Token(4)
                                        Call ocnMem.Execute(tSql, nEff)
                                    End If
                                End If '==If bChkWebservice Then '*Bmp 61-03-28 ถ้า token ผ่าน จึง Update
                                'ยังไม่มี Process การลบข้อมูล ออกจาก Table TPSTBPHis        '*Ao 2013/01/11
                                orsTemp.MoveNext
                            Loop
                        End If
                        orsPoint.MoveNext
                    Loop
                End If
            End If
        End If
    End If
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " End Update Member Point" '*Tao 61-05-09 เก็บ Log เพิ่ม
    Exit Sub
ErrHandle:
    '*Tao 61-05-22 ตรวจสอบก่อนว่า ยังพร้อมทำงานอยู่หรือไม่
    If SP_CHKbWebservice Then
        If W_DATbConnectOrChk(ocnMem, tMM_CNDbMServer, tMM_CNDbMUser, tMM_CNDbMPassword, tMM_CNDbMDB, True, nErr) Then
            Resume Next '*Tao 61-05-22 ให้ทำต่อให้ครบ
        End If
    End If
End Sub

Private Function W_GETbStoreOpen() As Boolean
'----------------------------------------------------------------------------------
'Ret:Store Open = True /Store Close = False
'Cmt: ตรวจสอบสถานะ Store
'----------------------------------------------------------------------------------
        Dim tSql$
        Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
        tSql = "SELECT FTTmsOpnStore  FROM TsysTerminalSta "
       Set orsTemp = ocnOnline.Execute(tSql)
       If orsTemp.EOF Then Exit Function
       If SP_FEDtChkString(orsTemp, "FTTmsOpnStore") = "Y" Then
            W_GETbStoreOpen = True
       Else
            W_GETbStoreOpen = False
       End If
       Set orsTemp = Nothing
       oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " W_GETbStoreOpen = " & W_GETbStoreOpen    '*Tao 61-05-09 เก็บ Log เพิ่ม
       Exit Function
ErrHandle:
      W_GETbStoreOpen = False
      oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " W_GETbStoreOpen = " & W_GETbStoreOpen    '*Tao 61-05-09 เก็บ Log เพิ่ม
End Function

Private Function W_GETtTmnNum() As String
'----------------------------------------------------------------------------------
'
'
'----------------------------------------------------------------------------------
        Dim tSql$
        Dim orsTemp As ADODB.Recordset
        Dim tComName$
On Error GoTo ErrHandle
        tComName = cUT.UT_SYStGetComName
        tSql = "SELECT FTTmnCode,FTTmnNum"
        tSql = tSql & vbCrLf & " FROM TCNMTerminalMtn"
        tSql = tSql & vbCrLf & " WHERE FTTmnName= '" & tComName & "'"
       Set orsTemp = ocnOnline.Execute(tSql)
       If Not (orsTemp.EOF Or orsTemp.EOF) Then
            W_GETtTmnNum = SP_FEDtChkString(orsTemp, "FTTmnNum")
       Else
            W_GETtTmnNum = ""
       End If
       Set orsTemp = Nothing
       Exit Function
ErrHandle:
      W_GETtTmnNum = ""
End Function

Private Sub W_CRTxLogTran()
'----------------------------------------------------------
'   Cmt :   Create File Log for Transfer Data
'   '*KT 52-01-12 เก็บ log Transfer
'----------------------------------------------------------
    Dim oFSO  As New FileSystemObject
    Dim oFile As File
    Dim bPassTxt  As Boolean
    Dim tFld$
    
On Error GoTo ErrHandle
    bPassTxt = False
'    tVB_CNPthFileLog = App.Path & "\"
    tVB_CNPthFileLog = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\LOG\"
    tVB_FileLog = tVB_CNPthFileLog & "TransferOffline" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & ".Log"
    tFld = oFSO.GetParentFolderName(tVB_FileLog)
    If Not SP_FLDbManFolder(tFld, nEN_Fld1Exist) Then
       Call SP_FLDbManFolder(tFld, nEN_Fld3Create)
    End If
    If Not SP_FLEbManFile(tVB_FileLog, nEN_Fle1Exist) Then
        oFSO.CreateTextFile (tVB_FileLog)
    End If
    Set oFile = oFSO.GetFile(tVB_FileLog)
    Set oVB_LogFile = oFile.OpenAsTextStream(ForWriting)
    bPassTxt = True
    oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Create : " & " Version " & App.Major & "." & App.Minor & "." & App.Revision
    Exit Sub
ErrHandle:
    If bPassTxt Then oVB_LogFile.WriteLine Format(Date, "yyyy/MM/dd") & " " & Time & " Error Number : " & Err.Number & vbCrLf & _
    "Can not Create flie TransferOffline.Log"
'    Debug.Print "Error " & Err.Number & " " & Err.Description
End Sub

'Private Sub W_CRTxClearLog()
''----------------------------------------------------------
''   Cmt :   Create File Log for Transfer Data
''   '*KT 52-01-12 เก็บ log Transfer
''----------------------------------------------------------
'    Dim oFSO As New FileSystemObject
'    Dim oFolder As Folder
'    Dim oFile As File
'    Dim dDatelog As Date
'
''TransferOffline20090320120749.Log
'On Error GoTo ErrHandle
'    tVB_CNPthFileLog = App.Path & "\"
'    If SP_FLDbManFolder(tVB_CNPthFileLog, nEN_Fld1Exist) Then
'            Set oFolder = oFSO.GetFolder(tVB_CNPthFileLog)
'             For Each oFile In oFolder.Files
'                If UCase(Right(oFile.Name, 4)) = ".LOG" And UCase(Mid(oFile.Name, 1, 15)) = UCase("TransferOffline") Then
'                    dDatelog = DateSerial(Mid(oFile.Name, 16, 4), Mid(oFile.Name, 20, 2), Mid(oFile.Name, 22, 2))
'                    If Abs(DateDiff("D", dDatelog, Date)) > nMaxDay Then
'                             oFSO.DeleteFile oFile.Path
'                    End If
'                End If
'             Next
'    End If
'    Set oFSO = Nothing
'    Exit Sub
'ErrHandle:
'End Sub


