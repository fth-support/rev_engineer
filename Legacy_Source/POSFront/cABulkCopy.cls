VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cABulkCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const tW_VKey As String = "_view"
Private Const tW_VApp As String = "View"
Private Const tW_VKeyTbl As String = "Table"
Private Const tW_VKeyFed As String = "Field"
Private Const tW_VKeyWhe As String = "Where"

Dim bConnected As Boolean
Dim tError As String
Dim oServer As SQLDMO.SQLServer2
Dim oBcp As SQLDMO.BulkCopy2
'Dim oFileObj As Scripting.FileSystemObject  'WINDOWS\system32\scrrun.dll
Dim WithEvents oBcpEv As SQLDMO.BulkCopy2
Attribute oBcpEv.VB_VarHelpID = -1

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Event OnImported(ByVal tMessage As String)
Public Event OnCopied(ByVal tMessage As String, ByVal nRow As Long)

Public Property Get ErrorDesc()
    ErrorDesc = tError
End Property

Private Sub Class_Initialize()
    tError = ""
    bConnected = False
    'Set oFileObj = New Scripting.FileSystemObject
    Set oServer = Nothing
    Set oBcp = Nothing
    Set oBcpEv = Nothing
End Sub

Private Sub Class_Terminate()
    If bConnected Then W_DEFxDisconnect
    'Set oFileObj = Nothing
    If Not oBcpEv Is Nothing Then Set oBcpEv = Nothing
    If Not oBcp Is Nothing Then Set oBcp = Nothing
    If Not oServer Is Nothing Then Set oServer = Nothing
End Sub

Public Sub W_DEFxConnect(ByVal ptServer$, ByVal ptUser$, ByVal ptPwd$)
    On Error GoTo ErrHandle
    
    If bConnected Then W_DEFxDisconnect
    If oServer Is Nothing Then Set oServer = New SQLDMO.SQLServer2

    oServer.LoginSecure = False
    oServer.Connect ptServer, ptUser, ptPwd
    bConnected = True
    Exit Sub
    
ErrHandle:
    tError = Err.Description
End Sub

Public Sub W_DEFxDisconnect()
    On Error GoTo ErrHandle
    
    If bConnected Then
        oServer.DisConnect
        bConnected = False
    End If
    Exit Sub
    
ErrHandle:
    tError = Err.Description
End Sub
    
Public Function W_GETnExportTbl(ByVal ptDb$, ByVal ptTbl$, ByVal ptFileName$) As Long
'----------------------------------------------------------------
'   Export ข้อมูลจาก Table เป็น text file
'----------------------------------------------------------------
    On Error GoTo ErrHandle
    Dim oDatabase As SQLDMO.Database2
    Dim oTable As SQLDMO.Table2
    
    W_GETnExportTbl = 0
    If Not bConnected Then
        MsgBox "Application not connected to server", vbCritical 'tms_CN005223
        Exit Function
    End If
    
    If oBcp Is Nothing Then Set oBcp = New SQLDMO.BulkCopy2
    If oBcpEv Is Nothing Then Set oBcpEv = oBcp
    Set oDatabase = oServer.Databases(ptDb)
    Set oTable = oDatabase.Tables(ptTbl)
    
    oBcp.DataFilePath = ptFileName
    oBcp.DataFileType = SQLDMODataFile_TabDelimitedChar 'กำหนดรูปแบบตัวคั่น Column ด้วย Tab และระหว่าง Record เป็น Return
    W_GETnExportTbl = oTable.ExportData(oBcp)
    
    Set oBcpEv = Nothing
    Set oBcp = Nothing
    Set oTable = Nothing
    Set oDatabase = Nothing
    
    Exit Function
    
ErrHandle:
    tError = Err.Description
    Exit Function
End Function

Public Function W_GETnExportQry(ByVal ptDb$, ByVal ptTbl$, ByVal ptFileName$, Optional ByVal ptFed$ = "*", Optional ByVal ptWhere$ = "") As Long
'----------------------------------------------------------------
'   Export ข้อมูลจาก Table โดยมีเงื่อนไข เป็น text file
'----------------------------------------------------------------
    On Error GoTo ErrHandle
    Dim nFile%, tFileView$
    Dim oDatabase As SQLDMO.Database2
    Dim oView As SQLDMO.View2
    Dim tSql$
    
    W_GETnExportQry = 0
    If Not bConnected Then
        MsgBox "Application not connected to server", vbCritical 'tms_CN005223
        Exit Function
    End If
    
    ptWhere = Replace$(ptWhere, "where", "", 1, 1, vbDatabaseCompare)
    tSql = "SELECT " & ptFed & " FROM " & ptTbl
    If Trim$(ptWhere) <> "" Then tSql = tSql & " WHERE " & Trim$(ptWhere)
    tSql = tSql & ";"
    
    Set oView = New SQLDMO.View2
    oView.Name = "tmp_" & Format$(Now, "yyyymmddhhnnss")
    oView.Text = "create VIEW " & oView.Name & " As " & tSql
    
    If oBcp Is Nothing Then Set oBcp = New SQLDMO.BulkCopy2
    If oBcpEv Is Nothing Then Set oBcpEv = oBcp
    Set oDatabase = oServer.Databases(ptDb)
    Call oDatabase.Views.Add(oView)
    
    oBcp.DataFilePath = ptFileName
    oBcp.DataFileType = SQLDMODataFile_TabDelimitedChar 'กำหนดรูปแบบตัวคั่น Column ด้วย Tab และระหว่าง Record เป็น Return
    W_GETnExportQry = oView.ExportData(oBcp)
    Call oView.Remove
    
    tFileView = Replace$(ptFileName, ".", tW_VKey & ".", 1, 1, vbDatabaseCompare)
    If tFileView = ptFileName Then tFileView = ptFileName & tW_VKey
    
    nFile = FreeFile
    Open tFileView For Output As nFile
    Call WritePrivateProfileString(tW_VApp, tW_VKeyTbl, Trim$(ptTbl), tFileView)
    Call WritePrivateProfileString(tW_VApp, tW_VKeyFed, Trim$(ptFed), tFileView)
    Call WritePrivateProfileString(tW_VApp, tW_VKeyWhe, Trim$(ptWhere), tFileView)
    Close #nFile
    
    Set oBcpEv = Nothing
    Set oBcp = Nothing
    Set oView = Nothing
    Set oDatabase = Nothing
    Exit Function
    
ErrHandle:
    tError = Err.Description
    Exit Function
End Function
    
Public Function W_GETnImportTbl(ByVal ptDb$, ByVal ptTbl$, ByVal ptFileName$, Optional ByVal pbAutoDel As Boolean = True) As Long
'----------------------------------------------------------------
'   Import ข้อมูลจาก text file เข้า Table
'----------------------------------------------------------------
    On Error GoTo ErrHandle
    Dim oDatabase As SQLDMO.Database2
    Dim oTable As SQLDMO.Table2
    Dim tSql$
    
    W_GETnImportTbl = 0
    If Not bConnected Then
        MsgBox "Application not connected to server", vbCritical 'tms_CN005223
        Exit Function
    End If
    
    If oBcp Is Nothing Then Set oBcp = New SQLDMO.BulkCopy2
    If oBcpEv Is Nothing Then Set oBcpEv = oBcp
    Set oDatabase = oServer.Databases(ptDb)
    Set oTable = oDatabase.Tables(ptTbl)
    
    With oBcp
        .DataFilePath = ptFileName
        .DataFileType = SQLDMODataFile_TabDelimitedChar 'กำหนดรูปแบบตัวคั่น Column ด้วย Tab และระหว่าง Record เป็น Return
        .ImportRowsPerBatch = 1000
    End With
    
    If pbAutoDel Then
        tSql = "TRUNCATE TABLE " & ptTbl
        oDatabase.ExecuteWithResults (tSql)
    End If
    W_GETnImportTbl = oTable.ImportData(oBcp)
    
    Set oBcpEv = Nothing
    Set oBcp = Nothing
    Set oTable = Nothing
    Set oDatabase = Nothing
    
    Exit Function
    
ErrHandle:
    tError = Err.Description
    Exit Function
End Function
    
Public Function W_GETnImportQry(ByVal ptDb$, ByVal ptFileName$, Optional ByVal ptTbl$ = "", Optional ByVal pbAutoDel As Boolean = True) As Long
'----------------------------------------------------------------
'   Import ข้อมูลจาก text file เข้า Table
'----------------------------------------------------------------
    On Error GoTo ErrHandle
    Dim oDatabase As SQLDMO.Database2
    Dim oTable As SQLDMO.Table2
    Dim tFileView$, tSql$, tWhere$, tTmp$, nFile%, nLen&
    
    W_GETnImportQry = 0
    If Not bConnected Then
        MsgBox "Application not connected to server", vbCritical 'tms_CN005223
        Exit Function
    End If
    
    'FileView
    tFileView = Replace$(ptFileName, ".", tW_VKey & ".", 1, 1, vbDatabaseCompare)
    If tFileView = ptFileName Then tFileView = ptFileName & tW_VKey
    
    'Table
    If ptTbl = "" Then   'Table
        nLen = GetPrivateProfileString(tW_VApp, tW_VKeyTbl, vbNullString, tTmp, 255, tFileView)
        ptTbl = Left$(tTmp, nLen)
    End If
    
    'Where
    nLen = GetPrivateProfileString(tW_VApp, tW_VKeyWhe, vbNullString, tTmp, 255, tFileView)
    tWhere = Left$(tTmp, nLen)
    
    If oBcp Is Nothing Then Set oBcp = New SQLDMO.BulkCopy2
    If oBcpEv Is Nothing Then Set oBcpEv = oBcp
    Set oDatabase = oServer.Databases(ptDb)
    Set oTable = oDatabase.Tables(ptTbl)
    
    With oBcp
        .DataFilePath = ptFileName
        .DataFileType = SQLDMODataFile_TabDelimitedChar 'กำหนดรูปแบบตัวคั่น Column ด้วย Tab และระหว่าง Record เป็น Return
        .ImportRowsPerBatch = 1000
    End With
    
    If pbAutoDel Then
        tSql = "DELETE FROM " & ptTbl
        If Trim$(tWhere) <> "" Then tSql = tSql & " WHERE " & Trim$(tWhere)
        tSql = tSql & ";"
        oDatabase.ExecuteWithResults (tSql)
    End If
    W_GETnImportQry = oTable.ImportData(oBcp)
    
    Set oBcpEv = Nothing
    Set oBcp = Nothing
    Set oTable = Nothing
    Set oDatabase = Nothing
    Exit Function
    
ErrHandle:
    tError = Err.Description
    Exit Function
End Function

Private Sub oBcpEv_BatchImported(ByVal Message As String)
    'Debug.Print "Imported(" & Message & ")"
    RaiseEvent OnImported(Message)
End Sub

Private Sub oBcpEv_RowsCopied(ByVal Message As String, ByVal rows As Long)
    'Debug.Print "Copied(" & Message & ",  & rows & ")"
    RaiseEvent OnCopied(Message, rows)
End Sub
