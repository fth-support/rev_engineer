Attribute VB_Name = "Module1"
'Option Explicit
'Public ocnOn As New ADODB.Connection
'Public ocnAll As New ADODB.Connection       '*Tee 51-02-19
'Public ocnCfg As ADODB.Connection
'Public oServer As New SQLDMO.SQLServer
'
'
'Public Function SP_RECbAdaIniRead(ByVal pbOffLine As Boolean) As Boolean
''----------------------------------------------------------
''*Duang 51-03-15
''   Call:   ptFile as specific file
''   Ret:    true if read ok
''----------------------------------------------------------
'    Dim tFile$, tSql$, tCon$
'    Dim orsTemp As ADODB.Recordset
'    Dim bOffLine As Boolean     '*GP 47-02-17
'    Dim tSQLAuthSrvType As String
'    Dim tCfgCode As String, tPath As String, tMenu As String
'    Dim oFSO As New FileSystemObject
'
'    On Error GoTo ErrHandle
'    SP_RECbAdaIniRead = False
'
'    'get ค่าจาก AdaIni.ada
'    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\ISSTools\AdaIni.Ada"
'    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
'            tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
'            tFile = tFile & "\ISSTools\AdaIni.Ada" '*TW 51-04-20
'    End If
'
'    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then Exit Function      'not found, exit
'
'    'connection
'    Set ocnCfg = New ADODB.Connection
'    tCon = "Provider=Microsoft.Jet.OLEDB.4.0;"
'    tCon = tCon + "Data Source=" & tFile & ";Jet OLEDB:Database Password="
'    ocnCfg.CursorLocation = adUseClient
'    ocnCfg.ConnectionTimeout = 15      ' 15 วินาที
'    ocnCfg.CommandTimeout = 15
'    ocnCfg.Open tCon
'    If ocnCfg.State <> adStateOpen Then Exit Function       'error connect, exit
'
'
'    tMenu = "AdaPosFont"
'    tPath = "SF_PTH"
'
'
'     '*Duang 51-03-15-----Get Data ---------------------------------------------------------------------------------------
'    tSql = "SELECT *"
'    tSql = tSql & " FROM TConfig "
'    tSql = tSql & " ORDER BY FTCfgCode ASC"
'    Set orsTemp = ocnCfg.Execute(tSql)
'    If Not (orsTemp.BOF And orsTemp.EOF) Then
'
'         'get first
'        orsTemp.Filter = adFilterNone
'        orsTemp.MoveFirst
'        Do While Not orsTemp.EOF
'            tCfgCode = SP_FEDtChkString(orsTemp, "FTCfgCode")
'            Select Case tCfgCode
'                Case tMenu & "1"                ', "SF_SQLDB1", "SD_SQLDB1", "SM_SQLDB1"     'ฐานข้อมูล Default,Online
'                    tVB_CNDbSQLAuthUsrC = SP_FEDtChkString(orsTemp, "FTCfgOn") 'User Name
'                    tVB_CNDbSQLAuthUsrP = SP_FEDtChkString(orsTemp, "FTCfgOff") 'User Password
'                    tVB_CNDbSrc = SP_FEDtChkString(orsTemp, "FTCfgValue") 'Server Name
'                    tVB_CNDbCurFile = SP_FEDtChkString(orsTemp, "FTCfgDescTha") 'DataBase Name
'                    tSQLAuthSrvType = SP_FEDtChkString(orsTemp, "FTCfgDescEng")
'                    If UCase(tSQLAuthSrvType) = "0" Then
'                        tVB_CNDbLoginTypeMain = "WINAUTHEN"
'                    Else
'                        tVB_CNDbLoginTypeMain = "WINUSER"
'                    End If
'                    '------------------------------------------------------------------------------------------------------------------------------
'                Case tMenu & "2"             ', "SF_SQLDB2", "SD_SQLDB2", "SM_SQLDB2"         'ฐานข้อมูลสำรอง ,Offline
'                    tVB_CNDbSQLAuthSecondC = SP_FEDtChkString(orsTemp, "FTCfgOn") 'User Name
'                    tVB_CNDbSQLAuthSecondP = SP_FEDtChkString(orsTemp, "FTCfgOff") 'User Password
'                    tVB_CNDbSecondSrc = SP_FEDtChkString(orsTemp, "FTCfgValue") 'Server Name
'                    tVB_CNDbSecondCurFile = SP_FEDtChkString(orsTemp, "FTCfgDescTha") 'DataBase Name
'                    tSQLAuthSrvType = SP_FEDtChkString(orsTemp, "FTCfgDescEng")
'                    If UCase(tSQLAuthSrvType) = "0" Then
'                        tVB_CNDbLoginTypeSecond = "WINAUTHEN"
'                    Else
'                        tVB_CNDbLoginTypeSecond = "WINUSER"
'                    End If
'
'
'                     '------------------------------------------------------------------------------------------------------------------------------
'                Case "SF_SQLDB1", "SF_SQLDB2"                   'ฐานข้อมูล Default,Online
'                    If tCfgCode = "SF_SQLDB1" Then
'                        tVB_CNDbSQLAuthUsrC = SP_FEDtChkString(orsTemp, "FTCfgOn") 'User Name
'                        tVB_CNDbSQLAuthUsrP = SP_FEDtChkString(orsTemp, "FTCfgOff") 'User Password
'                        tVB_CNDbSrc = SP_FEDtChkString(orsTemp, "FTCfgValue") 'Server Name
'                        tVB_CNDbCurFile = SP_FEDtChkString(orsTemp, "FTCfgDescTha") 'DataBase Name
'                        tSQLAuthSrvType = SP_FEDtChkString(orsTemp, "FTCfgDescEng")
'                        If UCase(tSQLAuthSrvType) = "0" Then
'                            tVB_CNDbLoginTypeSecond = "WINAUTHEN"
'                        Else
'                            tVB_CNDbLoginTypeSecond = "WINUSER"
'                        End If
'                        tVB_CNOnDbSrc = tVB_CNDbSrc
'                        tVB_CNOnDbSQLAuthUsrC = tVB_CNDbSQLAuthUsrC
'                        tVB_CNOnDbSQLAuthUsrP = tVB_CNDbSQLAuthUsrP
'                        tVB_CNOnDbCurFile = tVB_CNDbCurFile
'
'                    Else
'                        tVB_CNDbSQLAuthUsrC = SP_FEDtChkString(orsTemp, "FTCfgOn") 'User Name
'                        tVB_CNDbSQLAuthUsrP = SP_FEDtChkString(orsTemp, "FTCfgOff") 'User Password
'                        tVB_CNDbSrc = SP_FEDtChkString(orsTemp, "FTCfgValue") 'Server Name
'                        tVB_CNDbCurFile = SP_FEDtChkString(orsTemp, "FTCfgDescTha") 'DataBase Name
'                        tSQLAuthSrvType = SP_FEDtChkString(orsTemp, "FTCfgDescEng")
'
'                        If UCase(tSQLAuthSrvType) = "0" Then
'                            tVB_CNDbLoginTypeMain = "WINAUTHEN"
'                        Else
'                            tVB_CNDbLoginTypeMain = "WINUSER"
'                        End If
'                        tVB_CNOffDbSrc = tVB_CNDbSrc
'                        tVB_CNOffDbSQLAuthUsrC = tVB_CNDbSQLAuthUsrC
'                        tVB_CNOffDbSQLAuthUsrP = tVB_CNDbSQLAuthUsrP
'                        tVB_CNOffDbCurFile = tVB_CNDbCurFile
'
'                    End If
'
'                Case "SM_SQLDB1", "SM_SQLDB2"                   'ฐานข้อมูล Default,Online
'                    If tCfgCode = "SM_SQLDB1" Then
'                        tVB_CNDbSQLAuthUsrC = SP_FEDtChkString(orsTemp, "FTCfgOn") 'User Name
'                        tVB_CNDbSQLAuthUsrP = SP_FEDtChkString(orsTemp, "FTCfgOff") 'User Password
'                        tVB_CNDbSrc = SP_FEDtChkString(orsTemp, "FTCfgValue") 'Server Name
'                        tVB_CNDbCurFile = SP_FEDtChkString(orsTemp, "FTCfgDescTha") 'DataBase Name
'                        tSQLAuthSrvType = SP_FEDtChkString(orsTemp, "FTCfgDescEng")
'                        tVB_CNOnMemDbSrc = tVB_CNDbSrc
'                        tVB_CNOnMemDbSQLAuthUsrC = tVB_CNDbSQLAuthUsrC
'                        tVB_CNOnMemDbSQLAuthUsrP = tVB_CNDbSQLAuthUsrP
'                        tVB_CNOnMemDbCurFile = tVB_CNDbCurFile
'
'                    Else
'                        tVB_CNDbSQLAuthUsrC = SP_FEDtChkString(orsTemp, "FTCfgOn") 'User Name
'                        tVB_CNDbSQLAuthUsrP = SP_FEDtChkString(orsTemp, "FTCfgOff") 'User Password
'                        tVB_CNDbSrc = SP_FEDtChkString(orsTemp, "FTCfgValue") 'Server Name
'                        tVB_CNDbCurFile = SP_FEDtChkString(orsTemp, "FTCfgDescTha") 'DataBase Name
'                        tSQLAuthSrvType = SP_FEDtChkString(orsTemp, "FTCfgDescEng")
'                        tVB_CNOffMemDbSrc = tVB_CNDbSrc
'                        tVB_CNOffMemDbSQLAuthUsrC = tVB_CNDbSQLAuthUsrC
'                        tVB_CNOffMemDbSQLAuthUsrP = tVB_CNDbSQLAuthUsrP
'                        tVB_CNOffMemDbCurFile = tVB_CNDbCurFile
'
'                    End If
'                Case "SD_DBUSETYPE" 'ใช้ใน 'SD = Service Document สถานะผู้ใช้ว่าเป็น Department หรือ  Super Market
'                    tSQLAuthSrvType = SP_FEDtChkString(orsTemp, "FTCfgValue")
'                    If UCase(tSQLAuthSrvType) = "DEPARTMENT" Then
'                        bVB_CNDbSQLAuthSrvDepartment = True
'                    Else 'SUPER MARKET
'                        bVB_CNDbSQLAuthSrvDepartment = False
'                    End If
'
'                Case "SF_LASTDOC" 'เลขที่บิลสุดท้าย
'                    nVB_CNLastBillS = SP_FEDtChkString(orsTemp, "FTCfgValue")       'เลขที่บิลสุดท้าย
'                    tVB_CNTermNo = SP_FEDtChkString(orsTemp, "FTCfgOn")             'Terminal Number
'                    'tVB_CNLastBillNo = SP_FEDtChkString(orsTemp, "FTCfgOff")        'คำนำหน้า บิล
'
'                Case tPath & "REPORT"   'Path Report
'                    tVB_CNDbRptPath = SP_FEDtChkString(orsTemp, "FTCfgValue")
'
'                Case tPath & "IMPDRIVE" 'Path Copy TrumpDrive
'                   ' tVB_CNPthDrive = SP_FEDtChkString(orsTemp, "FTCfgValue")
'
'                Case tPath & "BAKTR" 'Path Backup File dat
'                   ' tVB_CNPthBnkTr = SP_FEDtChkString(orsTemp, "FTCfgValue")
'
'                Case tPath & "RESSRPT" 'Path Recovery Script file
'                    'tVB_CNPthReScrip = SP_FEDtChkString(orsTemp, "FTCfgValue")
'
'                Case tPath & "EJ"       'EJ ณ เวลาที่ทำยังไม่มี Flow เรื่อง EJ ก็เลยกำหนดให้ใช้ Path เดียวกันไปก่อน
'                   ' tVB_CNOnDbSrcEJ = SP_FEDtChkString(orsTemp, "FTCfgValue")
'                   ' tVB_CNOffDbSrcEJ = SP_FEDtChkString(orsTemp, "FTCfgValue")
''                Case "RepStatus"
''                    bVB_CNRepChk = IIf(SP_FEDtChkString(orsTemp, "FTCfgValue") = "Y", True, False)
'            End Select
'            orsTemp.MoveNext
'        Loop
'
'        If tMenu = "AdaPosFont" Then
'             If pbOffLine Then           'SF_SQLDB2'ฐานข้อมูลสำรอง ,Offline
'                 tVB_CNDbSrc = tVB_CNOffDbSrc
'                 tVB_CNDbSQLAuthUsrC = tVB_CNOffDbSQLAuthUsrC
'                 tVB_CNDbSQLAuthUsrP = tVB_CNOffDbSQLAuthUsrP
'                 tVB_CNDbCurFile = tVB_CNOffDbCurFile
'             Else            'SF_SQLDB1
'                 tVB_CNDbSrc = tVB_CNOnDbSrc
'                 tVB_CNDbSQLAuthUsrC = tVB_CNOnDbSQLAuthUsrC
'                 tVB_CNDbSQLAuthUsrP = tVB_CNOnDbSQLAuthUsrP
'                 tVB_CNDbCurFile = tVB_CNOnDbCurFile
'             End If
'         End If
'    End If
'
'    If Not (orsTemp Is Nothing) Then Set orsTemp = Nothing
'    tVB_CNDbAdaIni = tFile     'keep path of AdaIni
'    SP_RECbAdaIniRead = True
'    Exit Function
'ErrHandle:
'
'End Function
'
'Public Function SQLbAvailableServer(ByVal ptSrvComName$, Optional ByVal ptComIP As String = "") As Boolean
''-----------------------------------------------------------
''   Call :
''   Cmt :
''   Developer : SOMJAI
''   Date Create : 02-13-2008 :  10:08:15
''-----------------------------------------------------------
''*KT 51-06-27 Comshheet ML-SF-360
'On Error GoTo ErrHandle
'    If ptComIP <> "" Then
'        If DB_LANbPingByIP(ptComIP) Then
'            SQLbAvailableServer = True
'        Else
'            SQLbAvailableServer = False
'        End If
'        If SQLbAvailableServer Then
'            Call SQLxConnectService(tVB_CNOnDbSrc, tVB_CNOnDbSQLAuthUsrC, tVB_CNOnDbSQLAuthUsrP)
'            If oServer.Status = 1 Then
'                SQLbAvailableServer = True
'            Else
'                SQLbAvailableServer = False
'            End If
'        End If
'    ElseIf ptSrvComName <> "" Then
'        Call SQLxConnectService(tVB_CNOnDbSrc, tVB_CNOnDbSQLAuthUsrC, tVB_CNOnDbSQLAuthUsrP)
'        If oServer.Status = 1 Then
'            SQLbAvailableServer = True
'        Else
'            SQLbAvailableServer = False
'        End If
'    Else
'        SQLbAvailableServer = False
'    End If
'    Exit Function
'ErrHandle:
'    Select Case Err.Number
'        Case -2147024890, -2147024891, -2147221504
'            Call SQLxConnectService(tVB_CNOnDbSrc, tVB_CNOnDbSQLAuthUsrC, tVB_CNOnDbSQLAuthUsrP)
'            If oServer.Status = 1 Then
'                SQLbAvailableServer = True
'            Else
'                SQLbAvailableServer = False
'            End If
'        Case Else
'            SQLbAvailableServer = False
'    End Select
'End Function
'
'Private Function DB_LANbPingByIP(ByVal ptIP As String) As Boolean
'''-----------------------------------------------------------
'''   Call :
'''   Cmt :
'''   Developer : SOMJAI
'''   Date Create : 02-20-2008 :  08:14:03
'''-----------------------------------------------------------
'    If ptIP <> "" Then
'        DB_LANbPingByIP = AP_Ping(ptIP)
'    Else
'        DB_LANbPingByIP = False
'    End If
'    Exit Function
'ErrHandle:
'    DB_LANbPingByIP = False
'End Function
'
'Public Sub SQLxConnectService(ByVal ptOnServerName$, ByVal ptOnUsrName$, ByVal ptOnUsrPwd$)
'On Error GoTo ErrorHandle
'    oServer.Connect ptOnServerName, ptOnUsrName, ptOnUsrPwd       'Connect  Service
'    Exit Sub
'ErrorHandle:
'    Exit Sub
'End Sub
'
'
'
