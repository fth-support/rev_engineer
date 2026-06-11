Attribute VB_Name = "Module2"
'Option Explicit
''file
'Public Enum eEN_FleManage
'    nEN_Fle1Exist = 1               'check if exist
'    nEN_Fle2Delete = 2           'delete file
'End Enum
'
'Public Function SP_FLEbManFile(ByVal ptPath As String, ByVal pnType As eEN_FleManage) As Boolean
''----------------------------------------------------------
''   Call:   ptPath is folder name
''               pnType  is enum of folder type 1:exist, 2:delete
''   Ret :  true if completed
''----------------------------------------------------------
'    Dim oFSO As Object
'
'    On Error GoTo ErrHandle
'
'    SP_FLEbManFile = False
'    If ptPath = "" Then Exit Function
'
'    Set oFSO = CreateObject("Scripting.FileSystemObject")
'
'    Select Case pnType
'        Case nEN_Fle1Exist
'            If oFSO.FileExists(ptPath) Then
'                SP_FLEbManFile = True
'            End If
'
'        Case nEN_Fle2Delete
'            If oFSO.FileExists(ptPath) Then
'                oFSO.DeleteFile (ptPath)
'                SP_FLEbManFile = True
'            End If
'    End Select
'
'    Exit Function
'ErrHandle:
'    SP_FLEbManFile = False
'End Function
'
''Public Function SP_FEDtChkString(ByVal poRs As ADODB.Recordset, ByVal ptField$, Optional pvRetOld) As String
'''----------------------------------------------------------
'''   Call:   poRs as ADODB.Recordset
'''               ptField is field to check
'''               pvRetOld(variant) if empty, or null want to return old
'''   Ret:    variant, only text and numeric
'''----------------------------------------------------------
''    Dim bEmpty As Boolean
''
''    If poRs.BOF And poRs.EOF Then
''        bEmpty = True
''    Else
''        If IsNull(poRs.Fields(ptField)) Then
''            bEmpty = True
''        Else
''            If Len(Trim(poRs.Fields(ptField))) = 0 Then
''                bEmpty = True
''            Else
''                SP_FEDtChkString = poRs.Fields(ptField)
''            End If
''        End If
''    End If
''    If bEmpty Then
''        If IsMissing(pvRetOld) Then
''            SP_FEDtChkString = ""
''        Else
''            SP_FEDtChkString = pvRetOld
''        End If
''    End If
''End Function
'
'Public Function SP_CHKbUnlock(ByVal ptTmnNum As String) As Boolean
''----------------------------------------------------------------------------------------
''*KT 51-07-2008
''Ret:  True : UnLock  False : Lock
''Call: Get Valuse FTEmpCode from TCNMTerminalMtn is Blank reture True , No Blank return False
''----------------------------------------------------------------------------------------
'On Error GoTo ErrHandle
'    Dim tSql$
'    Dim tEmpID$
'    SP_CHKbUnlock = False
'    tSql = "SELECT  FTEmpCode "
'    tSql = tSql & vbCrLf & "FROM TCNMTerminalMtn"
'    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptTmnNum & "'"
'    tEmpID = ocnOn.Execute(tSql).Fields(0).Value
'    SP_CHKbUnlock = IIf(tEmpID = "", True, False)
'    Exit Function
'ErrHandle:
'
'End Function
'
'
'
'
