Attribute VB_Name = "mCNSTP"
Option Explicit

Public oCmdSTP As New ADODB.Command
Public oRsSTP As ADODB.Recordset

Dim vW_ScriptProc As Variant
Public tW_AppPath$
Dim tW_PathStores$, tW_PathStoreBK$
Dim tW_FileStoreNew$, tW_FileStoreName$

Public Sub STP_CHKxScriptSQL()
'----------------------------------------------------------
'*Mik 50-10-30
'   Cmt :   Check Script Store Procedure
'----------------------------------------------------------
    Dim oFSO As New FileSystemObject
    Dim tSql As String
    Dim orsTmp As New ADODB.Recordset
    Dim tSriptName$, tVersion$
    Dim nLengthVersion%
    Dim tOldName$, tNewName$
    Dim tFileStoreName$, nLengthFile%
    Dim nI%, nJ%
    Dim tSriptNameTmp$, tVersionTmp$
    Dim nLengthVersionTmp%
    
On Error GoTo ErrHandle
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    tW_AppPath = oFSO.GetParentFolderName(App.Path)
    tW_PathStores = tW_AppPath & "\PosDB\ConfigScriptNew" '*TW 51-04-20
    
    If Not SP_FLDbManFolder(tW_PathStores, nEN_Fld1Exist) Then
        tW_AppPath = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
        tW_PathStores = tW_AppPath & "\PosDB\ConfigScriptNew" '*TW 51-04-20
    End If
    
    tW_AppPath = oFSO.GetParentFolderName(App.Path)
    tW_PathStoreBK = tW_AppPath & "\PosDB\ConfigScriptBackUp" '*TW 51-04-20
    If Not SP_FLDbManFolder(tW_PathStoreBK, nEN_Fld1Exist) Then
        tW_AppPath = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
        tW_PathStoreBK = tW_AppPath & "\PosDB\ConfigScriptBackUp" '*TW 51-04-20
    End If
    
    tW_FileStoreNew = tW_PathStores & "\*.SQL"
    tW_FileStoreName = Dir(tW_FileStoreNew, vbDirectory)
        
    '*Gon 50-11-06 ปรับวิธีการนิดหน่อยเพื่อไม่ให้เกิด Dead Lock กรณีไฟล์มีปัญหาแล้วค้างในโฟลเดอร์นั้นๆ
    ReDim aFileStoreName(0)
    Do While Trim(tW_FileStoreName) <> ""
        aFileStoreName(UBound(aFileStoreName)) = tW_FileStoreName
        ReDim Preserve aFileStoreName(UBound(aFileStoreName) + 1)
        tW_FileStoreName = Dir
    Loop
    
    For nI = LBound(aFileStoreName) To UBound(aFileStoreName) - 1   '*Gon 50-11-06
'    Do While tW_FileStoreName <> ""
        tW_FileStoreName$ = aFileStoreName(nI)  '*Gon 50-11-06
        
        vW_ScriptProc = Split(tW_FileStoreName, "_")     'STP_CNxBrwMain_1.0.SQL
        tSriptName = vW_ScriptProc(0) & "_" & vW_ScriptProc(1)
        nLengthVersion = Len(vW_ScriptProc(2))
        tVersion = Mid(vW_ScriptProc(2), 1, nLengthVersion - 4)
        
        tSql = "SELECT * FROM TSysSQLProc "
        tSql = tSql & " WHERE FTSpcName = '" & tSriptName & "' "
        tSql = tSql & " AND FTSpcVersion = '" & tVersion & "' "
        Call SP_SQLvExecute(tSql, orsTmp)
        If orsTmp.EOF Or orsTmp.BOF Then
            
            tW_FileStoreNew = tW_PathStores & "\" & tW_FileStoreName
            Call STP_CRTxStoreProcSql
            If UCase(tVB_CNAppName) = UCase("POSFront") Then
                 tSql = "INSERT INTO TSysSQLProc( FTSpcName, FTSpcVersion," & tCS_CNLastFedIns
                tSql = tSql & " VALUES( '" & tSriptName & "', '" & tVersion & "', "
                tSql = tSql & cUT.UT_SQLtLastIns(tVB_CNUserName)
                Call SP_SQLvExecute(tSql)
            Else
                 tSql = "INSERT INTO TSysSQLProc( FTSpcName, FTSpcVersion," & tCS_CNLastFedIns & ")"
                tSql = tSql & " VALUES( '" & tSriptName & "', '" & tVersion & "', "
                tSql = tSql & cUT.UT_SQLtLastIns(tVB_CNUserName) & ")"
                Call SP_SQLvExecute(tSql)
            End If
            
            'Remove File Script.SQL ไปยัง tW_PathStoreBK
            tOldName = tW_FileStoreNew
            tNewName = tW_PathStoreBK & "\" & tW_FileStoreName
            If Not (SP_FLEbManFile(tNewName, nEN_Fle1Exist)) Then            '*AO 51-02-19
                Name tOldName As tNewName
            Else
                On Error Resume Next
                Kill (tNewName)
                Name tOldName As tNewName
                Kill (tOldName)
            End If
            
        Else
            'Rename File Script.SQL -> xxx_vv_Date_Time & Remove File ไปยัง tW_PathStoreBK
            nLengthFile = Len(tW_FileStoreName)
            tW_FileStoreNew = "\" & Left(tW_FileStoreName, nLengthFile - 4)
            tOldName = tW_PathStores & tW_FileStoreNew & ".SQL"
            tNewName = tW_PathStoreBK & tW_FileStoreNew & "_" & Format(Date, "YYYYMMDD") & "_" & Format(Time, "HHMMSS") & ".SQL"  ' Define filenames.
            If Not (SP_FLEbManFile(tNewName, nEN_Fle1Exist)) Then            '*AO 51-02-19
                Name tOldName As tNewName
            Else
                On Error Resume Next
                Kill (tNewName)
                Name tOldName As tNewName
                Kill (tOldName)
            End If
        End If
'        tW_FileStoreName$ = Dir
'    Loop
    Next nI
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    '*KT 51-06-30 ConfigScriptBackUp
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    tW_FileStoreNew = tW_PathStoreBK & "\*.SQL"
    tW_FileStoreName = Dir(tW_FileStoreNew, vbDirectory)
    ReDim aFileStoreName(0)
    Do While Trim(tW_FileStoreName) <> ""
        aFileStoreName(UBound(aFileStoreName)) = tW_FileStoreName
        ReDim Preserve aFileStoreName(UBound(aFileStoreName) + 1)
        tW_FileStoreName = Dir
    Loop
    
    nI = UBound(aFileStoreName) - 1
    ReDim aFileStoreName2(0)
    Do While nI >= 0
        tW_FileStoreName$ = aFileStoreName(nI)
        aFileStoreName2(UBound(aFileStoreName2)) = tW_FileStoreName
        If tW_FileStoreName <> "" Then
            vW_ScriptProc = Split(tW_FileStoreName, "_")     'STP_CNxBrwMain_1.0.SQL
            tSriptName = vW_ScriptProc(0) & "_" & vW_ScriptProc(1)
            nLengthVersion = Len(vW_ScriptProc(2))
            If Replace(vW_ScriptProc(2), ".SQL", "") = vW_ScriptProc(2) Then
                tVersion = vW_ScriptProc(2)
            Else
                tVersion = Mid(vW_ScriptProc(2), 1, nLengthVersion - 4)
            End If
            nJ = UBound(aFileStoreName) - 1
            Do While nJ >= 0
                tW_FileStoreName$ = aFileStoreName(nJ)
                If tW_FileStoreName <> "" Then
                    vW_ScriptProc = Split(tW_FileStoreName, "_")     'STP_CNxBrwMain_1.0.SQL
                    tSriptNameTmp = vW_ScriptProc(0) & "_" & vW_ScriptProc(1)
                    nLengthVersionTmp = Len(vW_ScriptProc(2))
                    If Replace(vW_ScriptProc(2), ".SQL", "") = vW_ScriptProc(2) Then
                        tVersionTmp = vW_ScriptProc(2)
                    Else
                        tVersionTmp = Mid(vW_ScriptProc(2), 1, nLengthVersionTmp - 4)
                    End If
                    If tSriptNameTmp = tSriptName Then
                        If Val(Format(tVersionTmp)) > Val(Format(tVersion)) Then
                            aFileStoreName2(UBound(aFileStoreName2)) = aFileStoreName(nJ)
                        End If
                         nI = nI - 1
                    End If
                End If
                nJ = nJ - 1
            Loop
            ReDim Preserve aFileStoreName2(UBound(aFileStoreName2) + 1)
        End If
        nI = nI - 1
    Loop
    For nI = LBound(aFileStoreName2) To UBound(aFileStoreName2) - 1
        tW_FileStoreName$ = aFileStoreName2(nI)
        vW_ScriptProc = Split(tW_FileStoreName, "_")     'STP_CNxBrwMain_1.0.SQL
        tSriptName = vW_ScriptProc(0) & "_" & vW_ScriptProc(1)
        nLengthVersion = Len(vW_ScriptProc(2))
        If Replace(vW_ScriptProc(2), ".SQL", "") = vW_ScriptProc(2) Then
            tVersion = vW_ScriptProc(2)
        Else
            tVersion = Mid(vW_ScriptProc(2), 1, nLengthVersion - 4)
        End If
        tSql = "SELECT * FROM TSysSQLProc "
        tSql = tSql & " WHERE FTSpcName = '" & tSriptName & "' "
        tSql = tSql & " AND FTSpcVersion = '" & tVersion & "' "
        Call SP_SQLvExecute(tSql, orsTmp)
        If orsTmp.EOF And orsTmp.BOF Then
            tW_FileStoreNew = tW_PathStoreBK & "\" & tW_FileStoreName
            If UCase(tVB_CNAppName) = UCase("POSFront") Then
                 tSql = "INSERT INTO TSysSQLProc( FTSpcName, FTSpcVersion," & tCS_CNLastFedIns
                tSql = tSql & " VALUES( '" & tSriptName & "', '" & tVersion & "', "
                tSql = tSql & cUT.UT_SQLtLastIns(tVB_CNUserName)
                Call SP_SQLvExecute(tSql)
            Else
                 tSql = "INSERT INTO TSysSQLProc( FTSpcName, FTSpcVersion," & tCS_CNLastFedIns & ")"
                tSql = tSql & " VALUES( '" & tSriptName & "', '" & tVersion & "', "
                tSql = tSql & cUT.UT_SQLtLastIns(tVB_CNUserName) & ")"
                Call SP_SQLvExecute(tSql)
            End If
        End If
    Next nI
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    Exit Sub
ErrHandle:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
    MsgBox Err.Number
End Sub

Public Sub STP_CHKxScriptSQLOffLine(ByVal pcnOff As ADODB.Connection)
'----------------------------------------------------------
' Call
'  Cmt :   Check Script Store Procedure
'----------------------------------------------------------
    Dim oFSO As New FileSystemObject
    Dim tSql As String
    Dim orsTmp As New ADODB.Recordset
    Dim tSriptName$, tVersion$
    Dim nLengthVersion%
    Dim tOldName$, tNewName$
    Dim tFileStoreName$, nLengthFile%
    Dim nI%, nJ%
    Dim tSriptNameTmp$, tVersionTmp$
    Dim nLengthVersionTmp%
    
On Error GoTo ErrHandle
    
    '*KT 51-06-30 ConfigScriptNew
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    tW_AppPath = oFSO.GetParentFolderName(App.Path)
    tW_PathStores = tW_AppPath & "\POSDB\ConfigScriptNew"
    If Not SP_FLDbManFolder(tW_PathStores, nEN_Fld1Exist) Then
        tW_AppPath = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
        tW_PathStores = tW_AppPath & "\POSDB\ConfigScriptNew" '*TW 51-04-20
    End If
    
    tW_FileStoreNew = tW_PathStores & "\*.SQL"
    tW_FileStoreName = Dir(tW_FileStoreNew, vbDirectory)
    
    'ปรับวิธีการนิดหน่อยเพื่อไม่ให้เกิด Dead Lock กรณีไฟล์มีปัญหาแล้วค้างในโฟลเดอร์นั้นๆ
    ReDim aFileStoreName(0)
    Do While Trim(tW_FileStoreName) <> ""
        aFileStoreName(UBound(aFileStoreName)) = tW_FileStoreName
        ReDim Preserve aFileStoreName(UBound(aFileStoreName) + 1)
        tW_FileStoreName = Dir
    Loop
    
    For nI = LBound(aFileStoreName) To UBound(aFileStoreName) - 1
        tW_FileStoreName$ = aFileStoreName(nI)
        vW_ScriptProc = Split(tW_FileStoreName, "_")     'STP_CNxBrwMain_1.0.SQL
        tSriptName = vW_ScriptProc(0) & "_" & vW_ScriptProc(1)
        nLengthVersion = Len(vW_ScriptProc(2))
        tVersion = Mid(vW_ScriptProc(2), 1, nLengthVersion - 4)
        tSql = "SELECT * FROM TSysSQLProc "
        tSql = tSql & " WHERE FTSpcName = '" & tSriptName & "' "
        tSql = tSql & " AND FTSpcVersion = '" & tVersion & "' "
        Call SP_SQLvExecute(tSql, orsTmp, pcnOff)
        If orsTmp.EOF And orsTmp.BOF Then
            tW_FileStoreNew = tW_PathStores & "\" & tW_FileStoreName
            Call STP_CRTxStoreProcSqlOffline(pcnOff)
            tSql = "INSERT INTO TSysSQLProc( FTSpcName, FTSpcVersion," & tCS_CNLastFedIns
            tSql = tSql & " VALUES( '" & tSriptName & "', '" & tVersion & "', "
            tSql = tSql & cUT.UT_SQLtLastIns(tVB_CNUserName)
            Call SP_SQLvExecute(tSql, , pcnOff)
        End If
    Next nI
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    '*KT 51-06-30 ConfigScriptBackUp
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    tW_AppPath = oFSO.GetParentFolderName(App.Path)
    tW_PathStoreBK = tW_AppPath & "\POSDB\ConfigScriptBackUp"
    If Not SP_FLDbManFolder(tW_PathStoreBK, nEN_Fld1Exist) Then
        tW_AppPath = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
        tW_PathStoreBK = tW_AppPath & "\POSDB\ConfigScriptBackUp" '*TW 51-04-20
    End If
    
    tW_FileStoreNew = tW_PathStoreBK & "\*.SQL"
    tW_FileStoreName = Dir(tW_FileStoreNew, vbDirectory)

    ReDim aFileStoreName(0)
    Do While Trim(tW_FileStoreName) <> ""
        aFileStoreName(UBound(aFileStoreName)) = tW_FileStoreName
        ReDim Preserve aFileStoreName(UBound(aFileStoreName) + 1)
        tW_FileStoreName = Dir
    Loop
    
    nI = UBound(aFileStoreName) - 1
    ReDim aFileStoreName2(0)
    Do While nI >= 0
        tW_FileStoreName$ = aFileStoreName(nI)
        aFileStoreName2(UBound(aFileStoreName2)) = tW_FileStoreName
        'Debug.Print tW_FileStoreName
        If tW_FileStoreName <> "" Then
            vW_ScriptProc = Split(tW_FileStoreName, "_")     'STP_CNxBrwMain_1.0.SQL
            tSriptName = vW_ScriptProc(0) & "_" & vW_ScriptProc(1)
            nLengthVersion = Len(vW_ScriptProc(2))
            If Replace(vW_ScriptProc(2), ".SQL", "") = vW_ScriptProc(2) Then
                tVersion = vW_ScriptProc(2)
            Else
                tVersion = Mid(vW_ScriptProc(2), 1, nLengthVersion - 4)
            End If
            nJ = UBound(aFileStoreName) - 1
            Do While nJ >= 0
                tW_FileStoreName$ = aFileStoreName(nJ)
                If tW_FileStoreName <> "" Then
                    vW_ScriptProc = Split(tW_FileStoreName, "_")     'STP_CNxBrwMain_1.0.SQL
                    tSriptNameTmp = vW_ScriptProc(0) & "_" & vW_ScriptProc(1)
                    nLengthVersionTmp = Len(vW_ScriptProc(2))
                    If Replace(vW_ScriptProc(2), ".SQL", "") = vW_ScriptProc(2) Then
                        tVersionTmp = vW_ScriptProc(2)
                    Else
                        tVersionTmp = Mid(vW_ScriptProc(2), 1, nLengthVersionTmp - 4)
                    End If
                    If tSriptNameTmp = tSriptName Then
                        If Val(Format(tVersionTmp)) > Val(Format(tVersion)) Then
                            aFileStoreName2(UBound(aFileStoreName2)) = aFileStoreName(nJ)
                            'Debug.Print tW_FileStoreName
                        End If
                        nI = nI - 1
                    End If
                End If
                nJ = nJ - 1
            Loop
            ReDim Preserve aFileStoreName2(UBound(aFileStoreName2) + 1)
        End If
        nI = nI - 1
    Loop
    
    For nI = LBound(aFileStoreName2) To UBound(aFileStoreName2) - 1
        tW_FileStoreName$ = aFileStoreName2(nI)
        vW_ScriptProc = Split(tW_FileStoreName, "_")     'STP_CNxBrwMain_1.0.SQL
        tSriptName = vW_ScriptProc(0) & "_" & vW_ScriptProc(1)
        nLengthVersion = Len(vW_ScriptProc(2))
        If Replace(vW_ScriptProc(2), ".SQL", "") = vW_ScriptProc(2) Then
            tVersion = vW_ScriptProc(2)
        Else
            tVersion = Mid(vW_ScriptProc(2), 1, nLengthVersion - 4)
        End If
        tSql = "SELECT * FROM TSysSQLProc "
        tSql = tSql & " WHERE FTSpcName = '" & tSriptName & "' "
        tSql = tSql & " AND FTSpcVersion = '" & tVersion & "' "
        Call SP_SQLvExecute(tSql, orsTmp, pcnOff)
        If orsTmp.EOF And orsTmp.BOF Then
            tW_FileStoreNew = tW_PathStoreBK & "\" & tW_FileStoreName
            Call STP_CRTxStoreProcSqlOffline(pcnOff)
            tSql = "INSERT INTO TSysSQLProc( FTSpcName, FTSpcVersion," & tCS_CNLastFedIns
            tSql = tSql & " VALUES( '" & tSriptName & "', '" & tVersion & "', "
            tSql = tSql & cUT.UT_SQLtLastIns(tVB_CNUserName)
            Call SP_SQLvExecute(tSql, , pcnOff)
        End If
    Next nI
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
ErrHandle:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
    Select Case Err.Number
        Case 3704
           MsgBox "Error Number : " & Err.Number & vbCrLf & "Desc : " & Err.Description & vbCrLf & "SQL : " & tSql
        Case Else
          MsgBox "Error Number : " & Err.Number & vbCrLf & "Desc : " & Err.Description
    End Select
End Sub


Public Sub STP_CRTxStoreProcSql()
'----------------------------------------------------------
'   Call :  Ceart Store Procedure จากไฟล์ .SQL -> SQL Server
'   Developer : ARUNSRI
'   Date Create : 10-25-2007 :  09:00 AM
'----------------------------------------------------------
    Dim nFileNumber As Integer
    Dim tTemp As String
    Dim tConTmp As String
    
On Error GoTo ErrHandle
    
    nFileNumber = FreeFile
    tTemp = ""
    tConTmp = ""
    Open tW_FileStoreNew For Input As nFileNumber
     While Not EOF(nFileNumber)
         Line Input #nFileNumber, tTemp
         tTemp = Replace(tTemp, Chr(255), vbNullString)
         tTemp = Replace(tTemp, Chr(254), vbNullString)
         If Trim(UCase(tTemp)) = "GO" Then
             ocnAll.Execute (tConTmp)
             tConTmp = ""
         Else
             tConTmp = tConTmp & tTemp & vbNewLine
         End If
    Wend
'    Debug.Print tConTmp
    Close #nFileNumber
    If Trim(tConTmp) <> "" Then ocnAll.Execute (tConTmp)
    
    Exit Sub
ErrHandle:
    MsgBox Err.Number
'    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Public Sub STP_CRTxStoreProcSqlOffline(ByVal pcnOff As ADODB.Connection)
'----------------------------------------------------------
'   Call :  Ceart Store Procedure จากไฟล์ .SQL -> SQL Server
'   Developer :
'   Date Create :
'----------------------------------------------------------
    Dim nFileNumber As Integer
    Dim tTemp As String
    Dim tConTmp As String
    
On Error GoTo ErrHandle
    
    nFileNumber = FreeFile
    tTemp = ""
    tConTmp = ""
    Open tW_FileStoreNew For Input As nFileNumber
     While Not EOF(nFileNumber)
         Line Input #nFileNumber, tTemp
         tTemp = Replace(tTemp, Chr(255), vbNullString)
         tTemp = Replace(tTemp, Chr(254), vbNullString)
         If Trim(UCase(tTemp)) = "GO" Then
             pcnOff.Execute (tConTmp)
             tConTmp = ""
         Else
             tConTmp = tConTmp & tTemp & vbNewLine
         End If
    Wend
'    Debug.Print tConTmp
    Close #nFileNumber
    If Trim(tConTmp) <> "" Then pcnOff.Execute (tConTmp)
    
    Exit Sub
ErrHandle:
    MsgBox Err.Number
'    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Public Function STP_SQLnCommandQuery(ByVal ptCommand$, ByVal ptTBName$, _
'                ByVal ptFieldName$, ByVal ptWhere$, ByVal ptOrderBy$, ByVal ptOrderFlag$) As Integer
''----------------------------------------------------------
''   Call :  ptCommand -> Format For Query, ptTBName -> Table Name, ptFieldName -> Field Name,
''           ptWhere -> Where, ptOrderBy -> Order By, ptOrderFlag -> "0" = "ASC"
''   Cmt :   For Query SQL
''   Developer : ARUNSRI
''   Date Create : 10-29-2007 :  10:00 AM
''----------------------------------------------------------
'    Dim tError%, nI%
'    On Error GoTo ErrHandle
'
'    'Set oCmdSTP = New ADODB.Command
'    With oCmdSTP
'        .ActiveConnection = ocnAll
'        .CommandType = adCmdText
'        Select Case UCase(ptCommand)
'        Case "SELALL"
'            .CommandText = "{ ? =CALL STP_CNxSelMain;1(?,?,?,?,?) }"
'
'            For nI = 1 To 5
'                  .Parameters.Item(nI).Direction = adParamInput
'            Next
'            .Parameters.Item(1).Value = ptTBName
'            .Parameters.Item(2).Value = ptFieldName
'            .Parameters.Item(3).Value = ptWhere
'            .Parameters.Item(4).Value = ptOrderBy
'            .Parameters.Item(5).Value = ptOrderFlag
'        Case "SELBRW"
'            .CommandText = "{ ? =CALL STP_CNxBrwMain;1(?,?,?,?) }"
'
'            For nI = 1 To 4
'                  .Parameters.Item(nI).Direction = adParamInput
'            Next
'            .Parameters.Item(1).Value = ptTBName
'            .Parameters.Item(2).Value = ptFieldName
'            .Parameters.Item(3).Value = ptWhere
'            .Parameters.Item(4).Value = ptOrderBy
'        End Select
'        Set oRsSTP = .Execute
'        tError = oCmdSTP.Parameters.Item(0).Value
'        STP_SQLnCommandQuery = tError
'    End With
'
'    Exit Function
'ErrHandle:
'    tError = oCmdSTP.Parameters.Item(0).Value
'    STP_SQLnCommandQuery = tError
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function

'Public Function STP_SQLbCommandExecute(ByVal ptCommand$, ByVal ptTBName$, ByVal ptFieldName$, ByVal ptVal$, ByVal ptWhere$) As Boolean
''----------------------------------------------------------
''   Call :  pCommand -> Format For Execute, pTBName -> Table Name, pFieldName -> Field Name (Update ใช้ Parametersนี้ด้วย -> "Field1=:DataVal1: AND Field2=:DataVal2:"),
''           pVal -> Values, pWhere -> Where
''   Cmt :   For Execute SQL
''   Developer : ARUNSRI
''   Date Create : 10-29-2007 :  14:00 AM
''----------------------------------------------------------
'    Dim tError%, nI%
'    On Error GoTo ErrHandle
'    STP_SQLbCommandExecute = False
'    'Set oCmdSTP = New ADODB.Command
'    With oCmdSTP
'        .ActiveConnection = ocnAll
'        .CommandType = adCmdText
'        Select Case UCase(ptCommand)
'        Case "INS"
'            .CommandText = "{ ? =CALL STP_CNxINSMain;1(?,?,?) }"
'
'            For nI = 1 To 3
'                  .Parameters.Item(nI).Direction = adParamInput
'            Next
'            .Parameters.Item(1).Value = ptTBName
'            .Parameters.Item(2).Value = ptFieldName
'            .Parameters.Item(3).Value = ptVal
'        Case "DEL"
'            .CommandText = "{ ? =CALL STP_CNxDelMain;1(?,?) }"
'            For nI = 1 To 2
'                  .Parameters.Item(nI).Direction = adParamInput
'            Next
'            .Parameters.Item(1).Value = ptTBName
'            .Parameters.Item(2).Value = ptWhere
'        Case "UPDATE"
'            .CommandText = "{ ? =CALL STP_CNxUpdateMain;1(?,?,?) }"
'
'            For nI = 1 To 3
'                  .Parameters.Item(nI).Direction = adParamInput
'            Next
'            .Parameters.Item(1).Value = ptTBName
'            .Parameters.Item(2).Value = ptFieldName
'            .Parameters.Item(3).Value = ptWhere
'        End Select
'        .Execute
'
'        tError = .Parameters.Item(0).Value
'        STP_SQLbCommandExecute = True
'    End With
'
'    Exit Function
'ErrHandle:
'    STP_SQLbCommandExecute = False
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Function
'
