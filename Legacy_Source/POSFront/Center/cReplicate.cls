VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
''Private ocnServer As ADODB.Connection
'
'Private tC_Server As String
'Private tC_DBName As String
'Private tC_User As String
'Private tC_Password As String
'
'Private tC_ServerOff As String
'Private tC_DBNameOff As String
'Private tC_UserOff As String
'Private tC_Passwordoff As String
'
'Private tC_RepName As String
'Private tC_PathSnapshot As String
'
'Private bC_Create As Boolean
'
'Private nC_CnnTimeOut As Integer
'
'Public Enum eEN_RepMode
'        nEN_CreatePublications = 1
'        nEN_CreateAddSubscript = 2
'        nEN_CreateAddSubScriber = 3
'        nEN_DropSubscript = 4
'        nEN_DropTransactional = 5
'        nEN__DropPullSubscript = 6
'        nEN_DropArticles = 7
'        nEN_AddArticles = 8
'        nEN_ChkSubscriptions = 9
'End Enum
'
'Public Enum eEN_RepOnorOff
'    nEN_RepOn = 1
'    nEN_RepOff = 2
'End Enum
'
'Public Property Let tServer(ByVal ptNewValue As String)         'SERVER
'    tC_Server = ptNewValue
'End Property
'Public Property Let tDBName(ByVal ptNewValue As String)     'DATABASE
'    tC_DBName = ptNewValue
'End Property
'Public Property Let tUsrName(ByVal ptNewValue As String)    'USERID
'    tC_User = ptNewValue
'End Property
'Public Property Let tUsrPwd(ByVal ptNewValue As String)     'PASSWORD
'    tC_Password = ptNewValue
'End Property
'
'Public Property Let tServerOff(ByVal ptNewValue As String)         'SERVER
'    tC_ServerOff = ptNewValue
'End Property
'Public Property Let tDBnameOff(ByVal ptNewValue As String)     'DATABASE
'    tC_DBNameOff = ptNewValue
'End Property
'Public Property Let tUsrNameOff(ByVal ptNewValue As String)    'USERID
'    tC_UserOff = ptNewValue
'End Property
'Public Property Let tUsrPwdOff(ByVal ptNewValue As String)     'PASSWORD
'    tC_Passwordoff = ptNewValue
'End Property
'
'Public Property Let tPublication(ByVal ptNewValue As String)    'REP NAME
'    tC_RepName = ptNewValue
'End Property
'Public Property Let tPathSnapshot(ByVal ptNewValue As String)    'PathSnapshot
'    tC_PathSnapshot = ptNewValue
'End Property
'Public Property Let nCnnTimeOut(ByVal ptNewValue As Integer) 'TIME CONNECT
'    nC_CnnTimeOut = ptNewValue
'End Property
'Public Property Get CreateCompleted() As Boolean
'    CreateCompleted = bC_Create
'End Property
'
'
'Public Function Connect(ByVal poCon As ADODB.Connection, ByVal pnRepMode As eEN_RepMode, ByVal pnRepOnorOff As eEN_RepOnorOff) As Boolean
''__________________________________________
''
''
''__________________________________________
'Dim tSql As String
'Dim tServer As String, tDBName As String, tUser As String, tPass As String
'
'On Error GoTo ErrHandler
'       Connect = False
'       Select Case pnRepOnorOff
'        Case nEN_RepOn
'            If tC_Server = "" Or tC_DBName = "" Then
'                Exit Function
'            End If
'            tServer = tC_Server
'            tDBName = tC_DBName
'            tUser = tC_User
'            tPass = tC_Password
'
'        Case nEN_RepOff
'            If tC_ServerOff = "" Or tC_DBNameOff = "" Then
'                Exit Function
'            End If
'            tServer = tC_ServerOff
'            tDBName = tC_DBNameOff
'            tUser = tC_UserOff
'            tPass = tC_Passwordoff
'       End Select
'
''*KT 51-09-17 แก้ไขเรื่อง Replication
''        Set ocnServer = New ADODB.Connection
''        Set ocnServer = poCon                    '****
'
'        If poCon.State = adStateClosed Then
'            tSql = "Provider=SQLOLEDB.1;Password=" & tPass & ";Persist Security Info=False"
'            tSql = tSql & ";User ID=" & tUser & ";Initial Catalog=" & tDBName & ";Data Source=" & tServer
'            poCon.CursorLocation = adUseClient
'            poCon.ConnectionTimeout = 90
'            poCon.CommandTimeout = 90
'            poCon.Open tSql
'        End If
'       Connect = True
'       Select Case pnRepMode
'            Case nEN_CreatePublications
'                Call C_CreatePublications(poCon)
'            Case nEN_CreateAddSubscript
'                Call C_CreateAddSubscript(poCon)
'            Case nEN_CreateAddSubScriber
'                Call C_CreateAddSubScriber(poCon)
'            Case nEN_DropSubscript
'                Call C_DropSubscript(poCon)
'            Case nEN_DropTransactional
'                Call C_DropTransactional(poCon)
'            Case nEN__DropPullSubscript
'                Call C_DropPullSubscript(poCon)
'            Case nEN_DropArticles
'                Call C_DropArticles(poCon)
'            Case nEN_AddArticles
'                Call C_AddArticles(poCon)
'            Case nEN_ChkSubscriptions
'                Call C_ChkSubscriptions(poCon)
'       End Select
'
''*KT 51-09-17 แก้ไขเรื่อง Replication
''       ocnServer.Close
''       Set ocnServer = Nothing
'
'    Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'    Call SP_MSGShowError
'    MsgBox "Connect", vbCritical
'End Function
'Public Sub C_CreatePublications(ByVal poCon As ADODB.Connection)
''__________________________________
''
''   Cmt:    Run Server
''__________________________________
'Dim tSql As String
'Dim nAff As Long
'On Error GoTo ErrHandler
'        bC_Create = False
'        tSql = "EXEC STP_RepPublications @ptServer='" & tC_Server & "',@ptDBName='" & tC_DBName & "'"
'        tSql = tSql & " ,@ptUser='" & tC_User & "', @ptPassword ='" & tC_Password & "',@ptPathSnapshot='" & tC_PathSnapshot & "'"
'        Call poCon.Execute(tSql, nAff)
'        Debug.Print nAff
'        If nAff = -1 Then
'            bC_Create = True
'        End If
'Exit Sub      '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'    Call SP_MSGShowError
'End Sub
'
'Public Sub C_CreateAddSubscript(ByVal poCon As ADODB.Connection)
''__________________________________
''
''   Cmt:    SERVER
''__________________________________
'Dim tSql As String
''Dim orsTmp As ADODB.Recordset
'
'On Error GoTo ErrHandler
'
'    tSql = "EXEC STP_RepAddSubscript @ptServer='" & tC_ServerOff & "',@ptDBName='" & tC_DBNameOff & "'"
'    Call poCon.Execute(tSql)
'
'    bC_Create = True
'Exit Sub      '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'    Call SP_MSGShowError
'    MsgBox "C_CreateAddSubscript", vbCritical
'End Sub
'Public Sub C_CreateAddSubScriber(ByVal poCon As ADODB.Connection)
''__________________________________
''
''   Cmt: SubScriber
''__________________________________
'Dim tSql As String
'
'On Error GoTo ErrHandler
'
'    tSql = "EXEC STP_RepAddSubScriber @ptServer='" & tC_Server & "',@ptDBName='" & tC_DBName & "'"
'    tSql = tSql & " ,@ptUser='" & tC_User & "', @ptPassword ='" & tC_Password & "'"
'    Call poCon.Execute(tSql)
'
'    bC_Create = True
'Exit Sub      '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'    Call SP_MSGShowError
'    MsgBox "C_CreateAddSubScriber", vbCritical
'End Sub
'Public Sub C_DropSubscript(ByVal poCon As ADODB.Connection)
''__________________________________
''
''   Cmt:
''__________________________________
'Dim tSql As String
'
'On Error GoTo ErrHandler
'    tSql = "EXEC STP_RepDropSubscript "
'    Call poCon.Execute(tSql)
'    bC_Create = True
'Exit Sub      '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'    Call SP_MSGShowError
'End Sub
'Public Sub C_DropTransactional(ByVal poCon As ADODB.Connection)
''__________________________________
''
''   Cmt: SERVER
''__________________________________
'Dim tSql As String
'
'On Error GoTo ErrHandler
'
'    tSql = "EXEC STP_RepDropTransactional @ptServerOff='" & tC_ServerOff & "',@ptDBNameOff='" & tC_DBNameOff & "'"
'    Call poCon.Execute(tSql)
'    bC_Create = True
'
'Exit Sub      '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'    Call SP_MSGShowError
'    MsgBox "C_DropTransactional", vbCritical
'End Sub
'Public Sub C_DropPullSubscript(ByVal poCon As ADODB.Connection)
''__________________________________
''
''   Cmt: SubScriber
''__________________________________
'Dim tSql As String
'
'On Error GoTo ErrHandler
'
'    tSql = "EXEC STP_RepDropPullSubscript @ptServer='" & tC_Server & "',@ptDBName='" & tC_DBName & "'"
'    Call poCon.Execute(tSql)
'
'    bC_Create = True
'Exit Sub      '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'    Call SP_MSGShowError
'    MsgBox "C_DropPullSubscript", vbCritical
'End Sub
'Public Sub C_DropArticles(ByVal poCon As ADODB.Connection)
''__________________________________
''
''   Cmt:
''__________________________________
'Dim tSql As String
'Dim nAff As Long
'On Error GoTo ErrHandler
'
'    tSql = "EXEC STP_RepDropArticles"
'    Call poCon.Execute(tSql, nAff)
'    Debug.Print nAff
'    bC_Create = True
'Exit Sub      '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'    Call SP_MSGShowError
'End Sub
'Public Sub C_AddArticles(ByVal poCon As ADODB.Connection)
''__________________________________
''
''   Cmt:d
''__________________________________
'Dim tSql As String
'Dim nAff As Long
'On Error GoTo ErrHandler
'
'    tSql = "EXEC STP_RepAddArticle"
'    Call poCon.Execute(tSql, nAff)
'    Debug.Print nAff
'    bC_Create = True
'Exit Sub      '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'    Call SP_MSGShowError
'End Sub
'
'Public Sub C_ChkSubscriptions(ByVal poCon As ADODB.Connection)
''__________________________________
''
''   Cmt:
''__________________________________
'Dim tSql As String
'Dim nAff As Long
'Dim orsTmp As ADODB.Recordset
'On Error GoTo ErrHandler
'
'    tSql = "SELECT DISTINCT srvname FROM syssubscriptions"
'    tSql = tSql & " WHERE srvname='" & tC_ServerOff & "'"
'    Set orsTmp = poCon.Execute(tSql)
'    If orsTmp.BOF And orsTmp.EOF Then
'        Call C_CreateAddSubscript(poCon)
'    End If
'    bC_Create = True
'Exit Sub      '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'    bC_Create = False
''    Call SP_MSGShowError
'End Sub
'
'Public Function C_CHKbSubscriptionsServer(ByVal poCon As ADODB.Connection, ByVal pnRepOnorOff As eEN_RepOnorOff) As Boolean
''__________________________________
''
''   Cmt:
''__________________________________
'Dim tSql As String
'Dim nAff As Long
'Dim orsTmp As ADODB.Recordset
'Dim tServer$
'Dim tDBName$
'Dim tUser$
'Dim tPass$
'
'On Error GoTo ErrHandler
'
'    Select Case pnRepOnorOff
'     Case nEN_RepOn
'         If tC_Server = "" Or tC_DBName = "" Then
'             Exit Function
'         End If
'         tServer = tC_Server
'         tDBName = tC_DBName
'         tUser = tC_User
'         tPass = tC_Password
'
'     Case nEN_RepOff
'         If tC_ServerOff = "" Or tC_DBNameOff = "" Then
'             Exit Function
'         End If
'         tServer = tC_ServerOff
'         tDBName = tC_DBNameOff
'         tUser = tC_UserOff
'         tPass = tC_Passwordoff
'    End Select
'
'     If poCon.State = adStateClosed Then
'         tSql = "Provider=SQLOLEDB.1;Password=" & tPass & ";Persist Security Info=False"
'         tSql = tSql & ";User ID=" & tUser & ";Initial Catalog=" & tDBName & ";Data Source=" & tServer
'         poCon.CursorLocation = adUseClient
'         poCon.ConnectionTimeout = 90
'         poCon.CommandTimeout = 90
'         poCon.Open tSql
'     End If
'
'    tSql = "SELECT DISTINCT srvname FROM syssubscriptions"
'    tSql = tSql & " WHERE srvname='" & tC_ServerOff & "'"
'    Set orsTmp = poCon.Execute(tSql)
'    If (orsTmp.BOF And orsTmp.EOF) Then
'        C_CHKbSubscriptionsServer = False
'    Else
'        C_CHKbSubscriptionsServer = True
'    End If
'    If Not orsTmp Is Nothing Then
'            If orsTmp.State = adStateOpen Then orsTmp.Close          'ไม่ใช้  State เพราะกรณี recordset ยังไม่ได้เปิดจะ error
'            Set orsTmp = Nothing
'    End If
'    Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
'ErrHandler:
'    C_CHKbSubscriptionsServer = True
'    If Not orsTmp Is Nothing Then
'            If orsTmp.State = adStateOpen Then orsTmp.Close          'ไม่ใช้  State เพราะกรณี recordset ยังไม่ได้เปิดจะ error
'            Set orsTmp = Nothing
'    End If
'End Function
