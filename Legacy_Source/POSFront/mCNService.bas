Attribute VB_Name = "mCNService"
Option Explicit


Public Sub SP_SQLxAlwConfigAdmin(ByVal poCn As ADODB.Connection)
'---------------------------------------------------------
'* Tee 52-07-16 FSBT-ADA ComSheet Support BSS-SFM 2009-018
' Cmt : Enable Config (SQL Surface Area)
' Call :
'---------------------------------------------------------
    Dim tSql As String
    Dim oRs As ADODB.Recordset
On Error GoTo ErrHandle:

    tSql = "SELECT"
    tSql = tSql & "  cfg.name AS [Name],"
    tSql = tSql & "  cfg.configuration_id AS [Number],"
    tSql = tSql & "  cfg.minimum AS [Minimum],"
    tSql = tSql & "  cfg.maximum AS [Maximum],"
    tSql = tSql & "  cfg.is_dynamic AS [Dynamic],"
    tSql = tSql & "  cfg.is_advanced AS [Advanced],"
    tSql = tSql & "  cfg.value AS [ConfigValue],"
    tSql = tSql & "  cfg.value_in_use AS [RunValue],"
    tSql = tSql & "  cfg.description AS [Description]"
    tSql = tSql & "  From"
    tSql = tSql & "  sys.configurations AS cfg"
    Set oRs = poCn.Execute(tSql)
    
    oRs.Filter = adFilterNone
    oRs.Filter = "Name ='" & "show advanced options" & "'"
    If Not (oRs.BOF And oRs.EOF) Then
        If oRs("ConfigValue") = 0 Then
            tSql = "EXEC sys.sp_configure N'show advanced options', 1"      'To update the currently configured value for advanced options.
            tSql = tSql & "  RECONFIGURE WITH OVERRIDE"
            poCn.Execute tSql
        End If
    Else
        MsgBox "Not found service show advanced options.", vbCritical
    End If
    
    oRs.Filter = adFilterNone
    oRs.Filter = "Name ='" & "Ole Automation Procedures" & "'"
    If Not (oRs.BOF And oRs.EOF) Then
        If oRs("ConfigValue") = 0 Then
            tSql = "EXEC sys.sp_configure N'Ole Automation Procedures', N'1'"       'To enable the feature.
            tSql = tSql & " RECONFIGURE WITH OVERRIDE"
            poCn.Execute tSql
        End If
    Else
        MsgBox "Not found service Ole Automation Procedures.", vbCritical
    End If
    
    oRs.Filter = adFilterNone
    oRs.Filter = "Name ='" & "xp_cmdshell" & "'"
    If Not (oRs.BOF And oRs.EOF) Then
        If oRs("ConfigValue") = 0 Then
            tSql = "EXEC sys.sp_configure N'xp_cmdshell', N'1'"     'To enable the feature.
            tSql = tSql & " RECONFIGURE WITH OVERRIDE"
            poCn.Execute tSql
        End If
    Else
        MsgBox "Not found service xp_cmdshell.", vbCritical
    End If
    
    oRs.Filter = adFilterNone
    oRs.Filter = "Name ='" & "Ad Hoc Distributed Queries" & "'"
    If Not (oRs.BOF And oRs.EOF) Then
        If oRs("ConfigValue") = 0 Then
            tSql = "EXEC sys.sp_configure N'Ad Hoc Distributed Queries', N'1'"     'To enable the feature.
            tSql = tSql & " RECONFIGURE WITH OVERRIDE"
            poCn.Execute tSql
        End If
    Else
        MsgBox "Not found service Ad Hoc Distributed Queries.", vbCritical
    End If
    
    oRs.Filter = adFilterNone
    oRs.Filter = "Name ='" & "clr enabled" & "'"
    If Not (oRs.BOF And oRs.EOF) Then
        If oRs("ConfigValue") = 0 Then
            tSql = "EXEC sys.sp_configure N'clr enabled', N'1'"     'To enable the feature.
            tSql = tSql & " RECONFIGURE WITH OVERRIDE"
            poCn.Execute tSql
        End If
    Else
        MsgBox "Not found service clr enabled.", vbCritical
    End If
    
    oRs.Filter = adFilterNone
    oRs.Filter = "Name ='" & "remote admin connections" & "'"
    If Not (oRs.BOF And oRs.EOF) Then
        If oRs("ConfigValue") = 0 Then
            tSql = "EXEC sys.sp_configure N'remote admin connections', N'1'"     'To enable the feature.
            tSql = tSql & " RECONFIGURE WITH OVERRIDE"
            poCn.Execute tSql
        End If
    Else
        MsgBox "Not found service remote admin connections.", vbCritical
    End If
    
    oRs.Filter = adFilterNone
    oRs.Filter = "Name ='" & "Replication XPs" & "'"
    If Not (oRs.BOF And oRs.EOF) Then
        If oRs("ConfigValue") = 0 Then
            tSql = "EXEC sys.sp_configure N'Replication XPs', N'1'"     'To enable the feature.
            tSql = tSql & " RECONFIGURE WITH OVERRIDE"
            poCn.Execute tSql
        End If
    Else
        MsgBox "Not found service Replication XPs.", vbCritical
    End If
    
    Exit Sub
ErrHandle:
    'Call SP_ERRbWriteFile(tUT_SaleDocNum, "mCNSP.SP_SQLxAlwConfigAdmin", Err.Number & " : " & Err.Description & "[ErrHandle]")    '*Tee 52-06-24
'    Call SP_MSGxShowError(Err.Number, Err.Description)
    MsgBox Err.Number
End Sub


