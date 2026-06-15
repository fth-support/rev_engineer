VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCNDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"OnlineDatabase"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'ADO DECLARE
Public WithEvents ocnOnLine As ADODB.Connection
Attribute ocnOnLine.VB_VarHelpID = -1
Public WithEvents ocnOffLine As ADODB.Connection        'Set Public for open when start program
Attribute ocnOffLine.VB_VarHelpID = -1
Private WithEvents ocnAllDB As ADODB.Connection     'For use ADO Event (Pass by Reference) Offline
Attribute ocnAllDB.VB_VarHelpID = -1
Private WithEvents ocnAllDB2 As ADODB.Connection     'For use ADO Event (Pass by Reference) Online 'Read Master
Attribute ocnAllDB2.VB_VarHelpID = -1
'Private WithEvents orsAll As ADODB.Recordset

Public oServer As New SQLDMO.SQLServer

Public Enum DbNetworkType       'For pass parameter
    Online
    Offline
End Enum

'PROPERTY DECLARE
'For Online
Private C_tOnServerName As String
Private C_tOnDbName As String
Private C_tOnUsrName As String
Private C_tOnUsrPwd As String
Private C_nOnCnnTimeOut As String
Private C_nOnCmdTimeOut As String
Private C_tOnComName As String
Private C_tOnComIP As String
'For Offline
Private C_tOffServerName As String
Private C_tOffDbName As String
Private C_tOffUsrName As String
Private C_tOffUsrPwd As String
Private C_nOffCnnTimeOut As String
Private C_nOffCmdTimeOut As String
Private C_tOffComName As String
Private C_tOffComIP As String
'For Check Srtatus
Private C_bIsAlwConnect As Boolean      'Allow when start transaction (Case offline to online)
Private C_bIsReExecute As Boolean   'use in  case connection fail or switch offline to online
Private C_bIsOpenOnline As Boolean  'set flag status when connect database online by ocnAll
Private C_bIsOpenOffline As Boolean     'set flag status when connect database offline by ocnOffline
Private C_bIsCurOnline As Boolean


Private Function DB_LANbPingByIP(ByVal ptIP As String) As Boolean
''-----------------------------------------------------------
''   Call :
''   Cmt :
''   Developer : SOMJAI
''   Date Create : 02-20-2008 :  08:14:03
''-----------------------------------------------------------
'On Error GoTo ErrHandle
'Dim tComputer$
'Dim oWMIService As Object
'Dim oColPing As Object
'Dim oStatus As Object
'
'    tComputer = "."
'    Set oWMIService = GetObject("winmgmts:\\" & tComputer & "\root\cimv2")
'    Set oColPing = oWMIService.ExecQuery("Select * From Win32_PingStatus where Address = '" & ptIP & "'")
'
'    For Each oStatus In oColPing
'        If IsNull(oStatus.StatusCode) Or oStatus.StatusCode <> 0 Then
'            DB_LANbPingByIP = False
'        Else
'            DB_LANbPingByIP = True
'        End If
'    Next

    If ptIP <> "" Then
        DB_LANbPingByIP = AP_Ping(ptIP)
    Else
        DB_LANbPingByIP = False
    End If
    Exit Function
ErrHandle:
    DB_LANbPingByIP = False
End Function

Private Function DB_GETxLookupIPAddress(ByVal ptHostName As String) As String
'------------------------------------------------------
'Ret: ip adress
'Cmt: ทำ get ip adress ของเครื่องจาก parameter ptHostName
'------------------------------------------------------
    Dim oWsa As WSAdata
    Dim nRet As Long
    Dim nTemp As Long
    Dim aTemp(0 To 3) As Byte
    Dim tOut As String
    Dim oHost As Hostent
    On Error GoTo ErrHandle
    
    WSAStartup &H10, oWsa
    nRet = gethostbyname(ptHostName)
     
    If nRet = 0 Then
        tOut = ""
    Else
        CopyMemory oHost, ByVal nRet, Len(oHost)
        If oHost.hLength >= 4 Then
            CopyMemory nTemp, ByVal oHost.hAddrList, 4
            CopyMemory aTemp(0), ByVal nTemp, 4
            tOut = Format(aTemp(0)) & "." & Format(aTemp(1)) & "." & Format(aTemp(2)) & "." & Format(aTemp(3))
        Else
            tOut = ""
        End If
    End If
    WSACleanup
    DB_GETxLookupIPAddress = tOut
    Exit Function
ErrHandle:
    DB_GETxLookupIPAddress = ""
End Function

Public Property Set ActiveConnection(poCn As ADODB.Connection)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 02-16-2008 :  17:27:17
'-----------------------------------------------------------
    On Error GoTo ErrHandle
        Set ocnAllDB = poCn
    Exit Property
ErrHandle:
    Call SP_MSGShowError
End Property

Public Property Set ActiveConnection2(poCn As ADODB.Connection)
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 02-16-2008 :  17:27:17
'Read Master
'-----------------------------------------------------------
    On Error GoTo ErrHandle
        Set ocnAllDB2 = poCn
    Exit Property
ErrHandle:
    Call SP_MSGShowError
End Property

Private Sub C_VARxSetDefault()
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 02-15-2008 :  14:24:04
'-----------------------------------------------------------
    On Error GoTo ErrHandle
    C_bIsAlwConnect = False
    C_bIsReExecute = False
    C_bIsOpenOnline = False
    C_bIsOpenOffline = False
    C_bIsCurOnline = False
    
    C_tOnServerName = ""
    C_tOnDbName = ""
    C_tOnUsrName = ""
    C_tOnUsrPwd = ""
    C_nOnCnnTimeOut = ""
    C_nOnCmdTimeOut = ""
    C_tOnComName = ""
    C_tOnComIP = ""
    
    C_tOffServerName = ""
    C_tOffDbName = ""
    C_tOffUsrName = ""
    C_tOffUsrPwd = ""
    C_nOffCnnTimeOut = ""
    C_nOffCmdTimeOut = ""
    C_tOffComName = ""
    C_tOffComIP = ""
    Exit Sub
ErrHandle:
    Err.Raise 1, "cCNDB", "C_VARxSetDefault"
End Sub

Public Function SQLbAvailableServer(ByVal ptSrvComName$, Optional ByVal ptComIP As String = "") As Boolean
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 02-13-2008 :  10:08:15
'-----------------------------------------------------------
'*KT 51-06-27 Comshheet ML-SF-360
'On Error GoTo ErrHandle
'    If ptComIP <> "" Then
'        If DB_LANbPingByIP(ptComIP) Then
'            SQLbAvailableServer = True
'        Else
'            SQLbAvailableServer = False
'        End If
'        If SQLbAvailableServer Then
'            DoEvents
'            If Not oDB.IsCurOnline Then Call SP_STAxSplashByAct("Verify service MS-SQL")
'            If oServer.Status = 1 Then
'                SQLbAvailableServer = True
'            Else
'                SQLbAvailableServer = False
'            End If
'        End If
'    ElseIf ptSrvComName <> "" Then
'        DoEvents
'        If oServer.Status = 1 Then
'            SQLbAvailableServer = True
'        Else
'            SQLbAvailableServer = False
'        End If
'    Else
'        SQLbAvailableServer = False
'    End If
'    oDB.IsCurOnline = SQLbAvailableServer
'    Exit Function
'ErrHandle:
'    If Not oDB.IsCurOnline Then
'        Select Case Err.Number
'            Case -2147024890, -2147024891, -2147221504
'                Call oDB.SQLxConnectService(oDB.OnServerName, oDB.OnUsrName, oDB.OnUsrPwd)
'                If oServer.Status = 1 Then
'                    SQLbAvailableServer = True
'                Else
'                    SQLbAvailableServer = False
'                End If
'        End Select
'    Else
'        SQLbAvailableServer = False
'    End If
'   oDB.IsCurOnline = SQLbAvailableServer
'Communication Sheet ML-SF-387 PosFront Is Slow
On Error GoTo ErrHandle
    SQLbAvailableServer = CBool(GetSetting("AdaPosMall", "POSFront", "Online", "1"))
    Exit Function
ErrHandle:
    SQLbAvailableServer = False
    On Error GoTo 0
End Function

Public Sub SQLxConnectService(ByVal ptOnServerName$, ByVal ptOnUsrName$, ByVal ptOnUsrPwd$)
On Error GoTo ErrorHandle
    oServer.Connect ptOnServerName, ptOnUsrName, ptOnUsrPwd       'Connect  Service
    Call SaveSetting("AdaPosMall", "POSFront", "Online", "1") '*Tao 60-04-03 ใช้แทน serviceonoff ตอนเปิดโปรแกรม
    Exit Sub
ErrorHandle:
    Call SaveSetting("AdaPosMall", "POSFront", "Online", "0") '*Tao 60-04-03
    Exit Sub
End Sub

Public Sub SQLxServiceNothing()
On Error GoTo ErrorHandle
    Set oServer = Nothing
    Exit Sub
ErrorHandle:
    Exit Sub
End Sub

Public Function SQLbConnectDB(ByRef poCn As ADODB.Connection, ByVal ptDbType As DbNetworkType) As Boolean
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 02-15-2008 :  10:36:47
'-----------------------------------------------------------
Dim tCon$
Dim nConTime As Long
Dim nCmdTime As Long
On Error GoTo ErrHandle
    
    SQLbConnectDB = False
    Select Case ptDbType
        Case Online
            If C_tOnServerName = "" Then
                Err.Raise 1, "cCNDB.SQLbConnectDB", "Not declare server name"
                Exit Function
            End If
            'tCon = "Provider=SQLNCLI.1;Password=" & C_tOnUsrPwd & ";Persist Security Info=False ;User ID=" & C_tOnUsrName & ";Initial Catalog=" & C_tOnDbName & ";Data Source=" & C_tOnServerName
            tCon = "Provider=SQLOLEDB.1;Password=" & C_tOnUsrPwd & ";Persist Security Info=True;User ID=" & C_tOnUsrName & ";Initial Catalog=" & C_tOnDbName & ";Data Source=" & C_tOnServerName
            nConTime = C_nOnCnnTimeOut
            nCmdTime = C_nOnCmdTimeOut

        Case Offline
            If C_tOffServerName = "" Then
                Err.Raise 1, "cCNDB.SQLbConnectDB", "Not declare server name"
                Exit Function
            End If
'            tCon = "Provider=SQLNCLI.1;Password=" & C_tOffUsrPwd & ";Persist Security Info=False ;User ID=" & C_tOffUsrName & ";Initial Catalog=" & C_tOffDbName & ";Data Source=" & OffServerName
            tCon = "Provider=SQLOLEDB.1;Password=" & C_tOffUsrPwd & ";Persist Security Info=True;User ID=" & C_tOffUsrName & ";Initial Catalog=" & C_tOffDbName & ";Data Source=" & OffServerName
            nConTime = C_nOffCnnTimeOut
            nCmdTime = C_nOffCmdTimeOut
    End Select
    
    '*KT 51-10-24 V 3.0.141/V 3.0.200
    '--------------------------------------------------------
    If poCn.State = adStateOpen Then
        'poCn.Close
        Set poCn = Nothing
        Set poCn = New ADODB.Connection 'poCn.Close '*KT 50-02-29 ทำการ Close ก่อน Open ครั้งใหม่
    End If
    '-------------------------------------------------------
    
    poCn.CursorLocation = adUseClient
    poCn.ConnectionTimeout = nConTime
    poCn.CommandTimeout = nCmdTime
    poCn.Open tCon
    If poCn.State = adStateOpen Then SQLbConnectDB = True
    SQLbConnectDB = True
    Exit Function
ErrHandle:
    Select Case ptDbType
        Case Online
            Call SP_MSGnShowing("tms_CN005136", nCS_Warning)
            Call SP_ExcServiceOnOff
        Case Offline
           Call SP_MSGnShowing("tms_CN005137", nCS_Warning)
    End Select
    'Call SP_MSGShowError
End Function

Public Property Get OnServerName() As String
    OnServerName = C_tOnServerName
End Property

Public Property Let OnServerName(ByVal ptNewValue As String)
    C_tOnServerName = ptNewValue
End Property

Public Property Get OnDbName() As String
    OnDbName = C_tOnDbName
End Property

Public Property Let OnDbName(ByVal ptNewValue As String)
    C_tOnDbName = ptNewValue
End Property

Public Property Get OnUsrName() As String
    OnUsrName = C_tOnUsrName
End Property

Public Property Let OnUsrName(ByVal ptNewValue As String)
    C_tOnUsrName = ptNewValue
End Property

Public Property Get OnUsrPwd() As String
    OnUsrPwd = C_tOnUsrPwd
End Property

Public Property Let OnUsrPwd(ByVal ptNewValue As String)
    C_tOnUsrPwd = ptNewValue
End Property

Public Property Get OnCnnTimeOut() As Integer
    OnCnnTimeOut = C_nOnCnnTimeOut
End Property

Public Property Let OnCnnTimeOut(ByVal pnNewValue As Integer)
    If pnNewValue = 0 Then pnNewValue = 60
    C_nOnCnnTimeOut = pnNewValue
End Property

Public Property Get OnCmdTimeOut() As Integer
    OnCmdTimeOut = C_nOnCmdTimeOut
End Property

Public Property Let OnCmdTimeOut(ByVal pnNewValue As Integer)
    If pnNewValue = 0 Then pnNewValue = 30
    C_nOnCmdTimeOut = pnNewValue
End Property

Public Property Get OnComName() As String
    OnComName = C_tOnComName
End Property

Public Property Let OnComName(ByVal ptNewValue As String)
    C_tOnComName = ptNewValue
    C_tOnComIP = DB_GETxLookupIPAddress(ptNewValue)
End Property

Public Property Get OnComIP() As String
    OnComIP = C_tOnComIP
End Property

Public Property Let OnComIP(ByVal ptNewValue As String)
    C_tOnComIP = ptNewValue
End Property

Public Property Get OffServerName() As String
    OffServerName = C_tOffServerName
End Property

Public Property Let OffServerName(ByVal ptNewValue As String)
    C_tOffServerName = ptNewValue
End Property

Public Property Get OffDbName() As String
    OffDbName = C_tOffDbName
End Property

Public Property Let OffDbName(ByVal ptNewValue As String)
    C_tOffDbName = ptNewValue
End Property

Public Property Get OffUsrName() As String
    OffUsrName = C_tOffUsrName
End Property

Public Property Let OffUsrName(ByVal ptNewValue As String)
    C_tOffUsrName = ptNewValue
End Property

Public Property Get OffUsrPwd() As String
    OffUsrPwd = C_tOffUsrPwd
End Property

Public Property Let OffUsrPwd(ByVal ptNewValue As String)
    C_tOffUsrPwd = ptNewValue
End Property

Public Property Get OffCnnTimeOut() As Integer
    OffCnnTimeOut = C_nOffCnnTimeOut
End Property

Public Property Let OffCnnTimeOut(ByVal pnNewValue As Integer)
    If pnNewValue = 0 Then pnNewValue = 40
    C_nOffCnnTimeOut = pnNewValue
End Property

Public Property Get OffCmdTimeOut() As Integer
    OffCmdTimeOut = C_nOffCmdTimeOut
End Property

Public Property Let OffCmdTimeOut(ByVal pnNewValue As Integer)
    If pnNewValue = 0 Then pnNewValue = 20
    C_nOffCmdTimeOut = pnNewValue
End Property

Public Property Get OffComName() As String
    OffComName = C_tOffComName
End Property

Public Property Let OffComName(ByVal ptNewValue As String)
    C_tOffComName = ptNewValue
End Property

Public Property Let OffComIP(ByVal ptNewValue As String)
    C_tOffComIP = ptNewValue
End Property

Public Property Get OffComIP() As String
    OffComIP = C_tOffComIP
End Property

Public Property Get IsReExecute() As Boolean
        IsReExecute = C_bIsReExecute
End Property

Public Property Let IsReExecute(ByVal pbNewValue As Boolean)
    C_bIsReExecute = pbNewValue
End Property

Public Property Get IsAlwConnect() As Boolean
        IsAlwConnect = C_bIsAlwConnect
End Property

Public Property Let IsAlwConnect(ByVal pbNewValue As Boolean)
    C_bIsAlwConnect = pbNewValue
End Property

Public Property Get IsOpenOnline() As Boolean
        IsOpenOnline = C_bIsOpenOnline
End Property

Public Property Let IsOpenOnline(ByVal pbNewValue As Boolean)
    C_bIsOpenOnline = pbNewValue
End Property

Public Property Get IsOpenOffline() As Boolean
        IsOpenOffline = C_bIsOpenOffline
End Property

Public Property Let IsOpenOffline(ByVal pbNewValue As Boolean)
    C_bIsOpenOffline = pbNewValue
End Property

Public Property Get IsCurOnline() As Boolean
        IsCurOnline = C_bIsCurOnline
End Property

Public Property Let IsCurOnline(ByVal pbNewValue As Boolean)
    C_bIsCurOnline = pbNewValue
End Property

Public Property Get CloneOffline() As ADODB.Connection
        Set CloneOffline = ocnOffLine '*TW 51-05-13
        'CloneOffline = ocnOffLine '*TW 51-05-13 Old
End Property

Private Sub Class_Initialize()
    Call C_VARxSetDefault
End Sub

Private Sub ocnAllDB_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If IsCurOnline Then
        IsOpenOnline = IIf(pError Is Nothing, True, False)
    Else
        IsOpenOffline = IIf(pError Is Nothing, True, False)
    End If
End Sub

Private Sub ocnAllDB_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
Static bModeRecovery As Boolean
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    
    Select Case adStatus
        Case adStatusOK

'            Call SP_DATxAppendQuery(pCommand.CommandText) '*Nirut 08/03/15 Add
            
        Case Else
'PH 2.0.0
'OLD___________________________________________________________
'            If pError.Number = -2147467259 Or Err.Number = 91 Or Err.Number = -2147217871 Then '*KT 52-01-12 แก้ไขให้ connect ไม่ได้หรือ Time off ให้ใช้ Mode Offline ทันที
'                Debug.Print "ocnAllDB_ExecuteComplete " & Err.Number
'                bVB_CNRecovryAct = True
'                Call SP_PRCxSwithToOffline '*Nirut 08-03-27
'                Set ActiveConnection = ocnAll
'                oRe.bW_ReQueryMode = True 'ไม่ให้เก็บ profileในการทำ SP_INSxTransHD
'                Call SP_INSxTransHD(tVB_CNTblHD) 'จำเป็นต้องให้ bW_ReQueryMode = True
'                '*KT 52-02-16 ป้องกันการ Close Terminal และ Total Report แล้วไม่เอา Dat File เข้า
'                '-------------------------------------------------------------------------------------------------
'                Select Case Val(Format(tVB_TransactionSub))
'                    Case tEN_TrnClose, tEN_TrnTotalReport
'                        MsgBox "ทำการนำเข้าข้อมูลลงฐานข้อมูล Local เพราะสถานะ offline กรุณารอสักครู่."
'                        Call SP_TBKbMov2DB        '*KT 51-08-19 Communication Sheet ML-SF-471 Backup sale data in terminal.xls
'                End Select
'                '-------------------------------------------------------------------------------------------------
'                oRe.bW_ReQueryMode = False 'หลังจากทำ SP_INSxTransHD ให้ set กลับค่าเดิม
'                Call SP_DATxRecoveryOn2Off 'recovery
'                '*KT 52-01-14 ทำการตรวจสอบ Trigger ในฐานข้อมูล Local ว่ามี Trigger นี้หรือไม่
'                '---------------------------------------------------------------------------------------------------------------------------
'                tSql = "select  name  from Sys.sysobjects where name = '" & "TRG_Tmp2Sale" & tVB_CNTerminalNum & "'"
'                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine)
'                If orsTemp.BOF And orsTemp.EOF Then
'                     '*KT 52-01-14 กรณีตรวจสอบเจอว่าไม่มี Trigger แล้วยังสร้างไม่สำเร็จทั้งฐานข้อมูล server และ local
'                    If Not SP_DATbCrtTrigger(False, True) Then
'                        'Call SP_MSGnShowing("tms_CN005168", nCS_Warning) '*KT 52-01-16 ไม่ต้องแจ้ง error
'                        'Call SP_APPxEndProgram ''*KT 52-01-16 ห้ามให้หลุดทำงานต่อไป
'                    End If
'                End If
'                Call SP_DATxRsNothing(orsTemp)
'                '-------------------------------------------------------------------------------------------------------------------------
'
'                bVB_CNRecovryAct = False
'                IsReExecute = True
'            Else
'                IsReExecute = False
'            End If
'________________________________________________________________

    End Select
End Sub

'Private Sub ocnAllDB2_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
''Read Master
'Static bModeRecovery As Boolean
'    Dim orsTemp As ADODB.Recordset
'    Dim tSql As String
'
'    Select Case adStatus
'        Case adStatusOK
'            Call SP_DATxAppendQuery(pCommand.CommandText) '*Nirut 08/03/15 Add
'        Case Else
'            'none
'    End Select
'
'End Sub

Private Sub ocnAllDB_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    tVB_CNSqlLastCmd = ConnectionString
End Sub

Private Sub ocnAllDB_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    tVB_CNSqlLastCmd = Source
End Sub

Private Sub ocnOffline_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    IsOpenOffline = IIf(pError Is Nothing, True, False)
End Sub

Private Sub ocnOffline_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
        IsOpenOffline = False
End Sub

Private Sub ocnOffline_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    tVB_CNSqlLastCmd = ConnectionString
End Sub
