VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'
'Private bW_StaCon As Boolean
'
'Public bW_ReQueryMode As Boolean
'
'
'Private Const tW_QuerySectionLine = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -"
'Private Const tW_QueryBreak = "GO  -- "
'Private Const tW_QueryBegin = "-- Begin --"
'Private Const tW_AlphabetAndNumber = "abcdefghijklmnopqrstuvwxyz0123456789"
'
'Private Const tCS_Err991 = "Recovery Error"
'Private Const tCS_Err992 = ""
'Private Const tCS_Err993 = ""
'
'Private ocnAll As ADODB.Connection
'
''property
'
'Private tW_FilterTable As String
'
'Private tW_FileName As String
'Private aW_FilterEvent() As Variant
'Private aW_FilterTable() As String
'
'
''file
'Public Enum eEN_ReFleManage
'    nEN_ReFle1Exist = 1               'check if exist
'    nEN_ReFle2Delete = 2           'delete file
'End Enum
'
'Public Property Let FilterTable(ByVal ptTables As String)
'
'    If Trim(ptTables) = vbNullString Then Exit Property
'
'    aW_FilterTable = Split(ptTables, ",")
'
'    tW_FilterTable = ptTables
'
'End Property
'
'Public Property Get FilterTable() As String
'    FilterTable = tW_FilterTable
'End Property
'
'
'Public Property Set ActiveConnection(poConnection As ADODB.Connection)
''Dim tServerName As String
'        Set ocnAll = poConnection
'        bW_StaCon = IIf(ocnAll.State = 1, True, False)
'       ' tServerName = ocnAll.Properties("Server Name") 'เกิด error เมื่อ offline
'End Property
'
'Public Property Get ActiveConnection() As ADODB.Connection
'    Set ActiveConnection = ocnAll
'End Property
'
'
'Public Function Create(ptFileName As String) As Boolean
'Dim oFile As File
'Dim oFSO As New FileSystemObject
'On Error GoTo ErrHandler
'        If Not bW_StaCon Then Exit Function
'        tW_FileName = ptFileName
'        If tW_FileName = vbNullString Then Exit Function
'        Set oFile = oFSO.GetFile(ptFileName)
'        oFile.Attributes = Normal
'        Open ptFileName For Output As #1
'
'        Print #1, "-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
'        Print #1, "-- Recovery Scripts File "
'        Print #1, "-- " & Format(Now, "yyyy-mm-dd")
'        Print #1, "-- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
'        Print #1, vbCrLf, vbCrLf
'
'        Print #1, tW_QueryBegin; tW_QuerySectionLine
'        Print #1, "USE " & ocnAll.DefaultDatabase & ";"
'        Print #1, vbCrLf & tW_QueryBreak & tW_QuerySectionLine
'
'        Close #1
'        Set oFSO = Nothing
''return
'Create = True
'
'Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'        Close #1
'End Function
'
'Public Function AppendQuery(ptSql As String, Optional ByVal ptFileName As String) As Boolean
'Dim nLof As Long
'Dim oFile As File
'Dim oFSO As New FileSystemObject
'On Error GoTo ErrHandler
'
'        If bW_ReQueryMode Then Exit Function
'        If Not bW_StaCon Then Exit Function
'        If tW_FilterTable = vbNullString Then Exit Function
'        If Not W_SQLbFilterQuery(ptSql) Then Exit Function  'กรองเอา query ที่จำเป็น
'
'        If ptFileName = vbNullString Then ptFileName = tW_FileName
'         If tW_FileName = vbNullString Then Exit Function
'        If Not RE_FLEbManFile(ptFileName, nEN_ReFle1Exist) Then Exit Function
'        Set oFile = oFSO.GetFile(ptFileName)
'        oFile.Attributes = Normal
'        Open ptFileName For Append As #1
'        nLof = LOF(1)
'        Seek #1, nLof + 1
'        Print #1, ptSql
'        Print #1, vbCrLf & tW_QueryBreak & tW_QuerySectionLine
'        Close #1
'        Set oFSO = Nothing
'
''return
'AppendQuery = True
'
'Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'        Close #1
'End Function
'
'Public Function Clear(Optional ByVal ptFileName As String) As Boolean
'
'On Error GoTo ErrHandler
'
'        If ptFileName = "" Then ptFileName = tW_FileName
'        If Not RE_FLEbManFile(tW_FileName, nEN_ReFle1Exist) Then Exit Function
'        Call Create(ptFileName)
'
''return
'Clear = True
'
'        Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'        Close #1
'End Function
'
'Public Function ReQuery(Optional ByVal ptFileName As String) As Boolean
''-----------------------------------------------------------------------------------
''Cmt: *KT 52-10-21
''_______________________________________________
'
'    Dim tSql As String, tTmp As String
'    Dim oFile As File
'    Dim oFSO As New FileSystemObject
'    Dim bPass As Boolean
'On Error GoTo ErrHandler
'
'        bPass = False
'
'        If ocnAll.State <> 1 Then Exit Function
'
'        If ptFileName = vbNullString Then
'            If Not RE_FLEbManFile(tW_FileName, nEN_ReFle1Exist) Then Exit Function
'            ptFileName = tW_FileName
'            If ptFileName = "" Then Exit Function
'            Set oFile = oFSO.GetFile(ptFileName)
'            If UCase(oFile.Name) = UCase("ScriptTrans.sql") And oFile.Size <= 476 Then Exit Function
'        End If
'
'            Open ptFileName For Input As #1
'
'        ocnAll.BeginTrans
'        bPass = True
'On Error GoTo ErrWithRollback
'
'         'goto begin point - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'        Do While Not EOF(1)
'                Line Input #1, tTmp
'                If Left(tTmp, Len(tW_QueryBegin)) = tW_QueryBegin Then
'                    Exit Do
'                End If
'        Loop
'
'         'exec by one script - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'        bW_ReQueryMode = True
'        Do While Not EOF(1)
'                Line Input #1, tTmp
'                If Left(tTmp, Len(tW_QueryBreak)) = tW_QueryBreak Then
'                    If UCase(Mid(tSql, 1, 8)) <> UCase("USE Mall") Then
'                        If UCase(Mid(tSql, 1, 3)) <> UCase("USE") Then
'                            If UCase(Mid(tSql, 1, Len("STP_CNxSQLMovSal2Tmp"))) = UCase("STP_CNxSQLMovSal2Tmp") Then
'                                If InStr(1, tSql, tVB_CNTblDTTmp) > 0 Then ocnAll.Execute "DELETE  " & tVB_CNTblDTTmp
'                                If InStr(1, tSql, tVB_CNTblCDTmp) > 0 Then ocnAll.Execute "DELETE  " & tVB_CNTblCDTmp
'                            End If
'                            ocnAll.Execute tSql
'                        End If
'                    End If
'                    tSql = vbNullString
'                Else
'                    tSql = tSql & tTmp & vbCrLf
'                End If
'        Loop
'        bW_ReQueryMode = False
'
'         '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'        Close #1
'        If bPass Then ocnAll.CommitTrans: bPass = False
'
'
''return
'ReQuery = True
'Exit Function
''- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrWithRollback:
'        If bPass Then ocnAll.RollbackTrans: bPass = False
'ErrHandler:
'        Close #1
'        bW_ReQueryMode = False
'        'Err.Raise 991, , tCS_Err991
'        Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function
'
'Private Function W_SQLbFilterQuery(ByVal ptSql As String) As Boolean
'Dim nTmp As Long
'Dim bChkEvent As Boolean, bChkTable As Boolean
'Dim tTbl As String, nDigit As Long
'
'
'On Error GoTo ErrHandler
'
'        If UBound(aW_FilterTable) = 0 Then Exit Function
'
'        'หา event
'        For nTmp = 0 To UBound(aW_FilterEvent)
'                If InStr(1, ptSql, aW_FilterEvent(nTmp), vbTextCompare) Then
'                    bChkEvent = True
'                    Exit For
'                End If
'        Next
'
'        If Not bChkEvent Then Exit Function
'
'        'หา table
'        For nTmp = 0 To UBound(aW_FilterTable)
'                tTbl = Trim(aW_FilterTable(nTmp))
'                nDigit = InStr(1, ptSql, tTbl, vbTextCompare)
'                If nDigit > 1 Then
'                    If InStr(1, tW_AlphabetAndNumber, Mid(ptSql, nDigit - 1, 1), vbTextCompare) = 0 And InStr(1, tW_AlphabetAndNumber, Mid(ptSql, nDigit + Len(tTbl), 1), vbTextCompare) = 0 Then
'                                bChkTable = True
'                                Exit For
'                    End If
'                End If
'        Next
'        If Not bChkTable Then Exit Function
'
''return
'W_SQLbFilterQuery = True
'
'Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'ErrHandler:
'
'
'End Function
'
'Private Sub Class_Initialize()
'
'        aW_FilterEvent = Array("insert", "update", "delete", "truncate", "STP_")
'
'End Sub
'
'Public Function RE_FLEbManFile(ByVal ptPath As String, ByVal pnType As eEN_ReFleManage) As Boolean
''----------------------------------------------------------
''   Call:   ptPath is folder name
''               pnType  is enum of folder type 1:exist, 2:delete
''   Ret :  true if completed
''----------------------------------------------------------
'    Dim oFSO As Object
'    Dim oFile As File
'    On Error GoTo ErrHandle
'
'    RE_FLEbManFile = False
'    If ptPath = "" Then Exit Function
'
'    Set oFSO = CreateObject("Scripting.FileSystemObject")
'
'    Select Case pnType
'        Case nEN_ReFle1Exist
'            If oFSO.FileExists(ptPath) Then
'                RE_FLEbManFile = True
'            End If
'
'        Case nEN_ReFle2Delete
'            If oFSO.FileExists(ptPath) Then
'                Set oFile = oFSO.GetFile(ptPath)
'                oFile.Attributes = Normal
'                oFSO.DeleteFile (ptPath)
'                RE_FLEbManFile = True
'            End If
'    End Select
'
'    Exit Function
'ErrHandle:
'    RE_FLEbManFile = False
'End Function
'
'
'
'
'
'
