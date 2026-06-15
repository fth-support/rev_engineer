Attribute VB_Name = "mDB"
Option Explicit
Sub Main()
'*KT 51-11-01
    Dim tFile$
    Dim oFSO As New FileSystemObject
    Dim LF$
    Dim nI As Long
    Dim tSta As String, tCheckerMode As String
    Dim bPass As Boolean
    Dim tMsg As String
    Dim bOnOffCur As Boolean '*Eaw 56-10-24
    Dim tStrDate As String
    Dim bRunAutoRep As Boolean '*Tao 59-05-31

On Error GoTo ErrHandle
     Call SP_CRTxLogPrg("Start EXE") '*KT 68-08-08 Debug
    'set use DLL
    '*Bump 56-04-03 Register ocx ของ OPOS เพิ่ม code 6 บรรทัด นับจากบรรทัดนี้
    If SP_FLEbManFile("C:\OPOS\FUJITSU\CashDrawer.ocx", nEN_Fle1Exist) Then Shell "RegSvr32 /s " & Chr(34) & "C:\OPOS\FUJITSU\CashDrawer.ocx" & Chr(34), vbHide
    If SP_FLEbManFile("C:\OPOS\FUJITSU\LineDisplay.ocx", nEN_Fle1Exist) Then Shell "RegSvr32 /s " & Chr(34) & "C:\OPOS\FUJITSU\LineDisplay.ocx" & Chr(34), vbHide
    If SP_FLEbManFile("C:\OPOS\FUJITSU\Msr.ocx", nEN_Fle1Exist) Then Shell "RegSvr32 /s " & Chr(34) & "C:\OPOS\FUJITSU\Msr.ocx" & Chr(34), vbHide
    If SP_FLEbManFile("C:\OPOS\FUJITSU\POSPrinter.ocx", nEN_Fle1Exist) Then Shell "RegSvr32 /s " & Chr(34) & "C:\OPOS\FUJITSU\POSPrinter.ocx" & Chr(34), vbHide
    If SP_FLEbManFile("C:\OPOS\FUJITSU\Scanner.ocx", nEN_Fle1Exist) Then Shell "RegSvr32 /s " & Chr(34) & "C:\OPOS\FUJITSU\Scanner.ocx" & Chr(34), vbHide
    Call SP_CRTxLogPrg("Register OPOS OCX") '*KT 68-08-08 Debug
    '*Eaw 56-06-07  POSFront.Exe Register DLL
    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\System\AdaOpos.dll"
    If SP_FLEbManFile(tFile, nEN_Fle1Exist) Then Shell "RegSvr32 /s " & Chr(34) & tFile & Chr(34), vbHide
    Call SP_CRTxLogPrg("Register AdaOpos.dll") '*KT 68-08-08 Debug
    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\System\AdaNETSDLL.dll"
    If SP_FLEbManFile(tFile, nEN_Fle1Exist) Then Shell "RegSvr32 /s " & Chr(34) & tFile & Chr(34), vbHide
    Call SP_CRTxLogPrg("Register AdaNETSDLL.dll") '*KT 68-08-08 Debug
    Call SP_SETxIE_SSL '*Bmp *61-04-26
    Call SP_CRTxLogPrg("SETxIE_SSL")
    
    '*Tao 59-02-26 Register AdaAES
    Call SP_REGxAdaAES
    
    '*Tao 60-08-25 Register AdaCallAPI
    Call SP_REGxAdaCallAPI
     '*KT 66-12-06 Register AdaMobileVoucher
    Call SP_REGxAdaMobileVoucher
    
    Call SP_CRTxWriteLogTerminated("POSFront Start" & "Main[0]")
    Call SaveSetting("AdaPosMall", "POSFront", "Online", 1)
    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "0")
           
    Set cUT = CreateObject("AdaUT.cUT")
    Set cUT_Print = CreateObject("AdaPrint.cPrinter")
    '    Set oRegPC = CreateObject("AdaRegPc.Verify")
    '    Set oRegHW = CreateObject("AdaRegHW.Machine")
    Set oOpos = CreateObject("AdaOpos.cFUJITSU")   '*Ao 2012/12/20          '*Ao 2013/01/24   Print FRONTTECH
    Set oOposSlip = CreateObject("AdaOpos.cFUJITSU")
    'OPOS
   tVB_CashDrawer = "DefaultCashDrawer"
   tVB_CashPrinter = "DefaultPOSPrinter"
   tVB_CashDisplay = "DefaultLineDisplay"
   tVB_CashMsr = "DefaultMSR"
   tVB_Scanner = "DefaultScanner"
   nVB_CashCodePage = 874
   tVB_CashSlipPrinter = "AlternatePOSPrinter"
   LF = Chr(10)
   'Preview Day
   bVB_CNTrStaRegPreviewDay = False
   Call SP_INIxSwitchModePreviewDay(False)
   Call SP_INIxClearVarPrv
       
    tVB_CNAppName = "POSFront" '*TW 51-04-20
    tVB_CNComName = cUT.UT_SYStGetComName
    Screen.MousePointer = 11
    bVB_CNPrgFsh = False
    bVB_FirstLoad = True '*Tao 60-01-20
    'bVb_SetResorution = True '*Tao 58-12-14 ของเดิม
    bVb_SetResorution = CBool(GetSetting("AdaPosMall", "POSFront", "SetResorution", True)) '*Tao 58-12-14 ตั้งค่าว่าจะเปลี่ยนขนาดหน้าจอหรือไม่
    
    bVB_CNTestSlipPrn = True
    '*Ao 52-02-23
    bVB_Res640X480 = False
    bVB_CNPastOffline = False
    '________________________
    
    'if this Exe is already open
    If App.PrevInstance Then
        Call SP_MSGnShowing(tCS_MSG61, nCS_Warning) 'tCS_MSG61 tms_CN999061= "โปรแกรมหน้าร้านถูกเปิดใช้อยู่แล้ว ไม่สามารถเปิดซ้ำได้อีก;This program is already opened, so can not run it again."
        End
    End If
    wBackGrand.Show
    Call SaveSetting("AdaPosMall", "POSFront", "EXEName", App.EXEName)     '*Ao 2013/09/06
    '*Bump 57-11-21 CommSheet TK-ISS3000-567(USB).xlsx
    If GetSetting("AdaPosMall", "POSFront", tCS_AutoDisableUSB) = "" Then
        '* ยังไม่มี ใน Registry
        '*  0 = Not Auto-Disable USB
        '*  1 = Auto-Disable USB
        Call SaveSetting("AdaPosMall", "POSFront", tCS_AutoDisableUSB, "1")
    End If
    
    Call W_DisableUSB(4) '*TON 58-12-01
    
    If bVB_Res640X480 Then
        Set wMain = wMainRet640
        Set wSub = wCNSubPay640
        Set wDisc = wCNSubDisc640
    Else
        Set wMain = wMainRet
        Set wSub = wCNSubPay
        Set wDisc = wCNSubDisc2
    End If
    
    'check authority
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "UT_PwdAuthorized[1]")
'    Call UT_PwdAuthorized       '*Ao 2013/06/06
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "UT_PwdAuthorized[2]")
     
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_DTExSetChrisEraAndFmtDte[1]")
    Call SP_DTExSetChrisEraAndFmtDte    'Format Type Year
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_DTExSetChrisEraAndFmtDte[2]")
    
    tVB_VirtualPosNo = GetSetting("AdaPosMall", "POSFront", "VirtualPosNo")
    'Default ตามwindow local
    tFile = GetSetting("AdaPosMall", "POSFront", "nVB_CNCutLang")
    If Trim(tFile) = "" Then
        nVB_CNCutLang = 1
    Else
        nVB_CNCutLang = 2
    End If
    '--------------------------------------
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_LNGxSetDefAftChg[1]")
    Call SP_LNGxSetDefAftChg
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_LNGxSetDefAftChg[2]")
    
    cUT_Print.UT_SETxPosLang = nVB_CNCutLang - 1
    
    'get ค่าจาก AdaIni.ada
    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\ISSTools\AdaIni.Ada" '*TW 51-04-20
    
    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
            tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
            tFile = tFile & "\ISSTools\AdaIni.Ada" '*TW 51-04-20
    End If
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_RECbAdaIniRead[1]")
    Call SP_STAxSplashByAct("Read configuration")
    'get ค่า สำหรับ new connection  ดึงค่าจาก Base Offline
    If Not SP_RECbAdaIniRead(tFile, False) Then         '*Duang 51-03-17 เปลี่ยนไฟล์ Adaini
        'Call SP_MSGnShowing("tms_CN005039", nCS_Warning) 'เกิดจากยังไม่มีการ Replicate และ ยังไม่ผ่านขั้นตอนการ Connect DB ต้องใช้ค่า Const ก่อน
        Call SP_MSGnShowing(tMS_CN535, tCS_CNMsgWarn)
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "SP_RECbAdaIniRead" & " Form:mDB" & " Control:General" & " Events:Main")
        Call SP_APPxEndProgram
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_RECbAdaIniRead[2]")
    
    wBackGrand.ostForm.Panels(3).Text = oDB.OffServerName & "." & oDB.OffDbName  'DB Offline PH 2.0.0
    
    If tVB_CNDbCurFile = "" Then tVB_CNDbCurFile = "AdaAcc"       'SQL
    tFile = tVB_CNDbCurFile
    '-------------------------------------------------------
    
'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "ocnAll.State = adStateClosed[1]")
   Call SP_STAxSplashByAct("Connect database")
    If ocnAll.State = adStateClosed Then        '*Tee 51-02-19
        If Not SP_SQLbConnect Then
           Call SP_MSGShowError
           Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "SP_SQLbConnect" & " Form:mDB" & " Control:General" & " Events:Main")
           Call SP_APPxEndProgram
        End If
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "ocnAll.State = adStateClosed[2]")

    Call SaveSetting("AdaPosMall", "POSFront", "RepTblLoss", "0") '*KT 54-10-07 V.4.0.1 ML-PRV-070 Invalid object name 'Table name'
    
'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SV_SQLxAlwConfigAdmin[1]")
   Call SP_SQLxAlwConfigAdmin(ocnAll) '*KT 52-10-20 ตรวจสอบ Service ของ SQL2005(Tricger , Replication)  ฝั่ง Local
'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SV_SQLxAlwConfigAdmin[2]")
    
'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_INItLodDB2INI[1]")
   Call SP_INItLodDB2INI
'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_INItLodDB2INI[2]")
   
   Call SP_STAxSplashByAct("Create store procedure")
   Call STP_CHKxScriptSQLOffLine(oDB.ocnOffLine) 'สร้าง Store Procedure ลง Offline
   Call STP_CHKxScriptSQL 'สร้าง Store Procedure ลง Online
   bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
   
'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "STP_CHKxScriptSQLOffLine[1]")
   Call SP_STAxSplashByAct("Execute store procedure")
   Call SP_PRCxAlterTbl  '*Eaw 56-06-21 เพิ่ม Field ให้ Table
   Call SP_CHKInsHDDelHis '*Bump 58-08-16 204_MissingEJ
   Call SP_EXECxStoreprocedure
   bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "STP_CHKxScriptSQL[2]")
   
'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_EXECxStoreprocedure[1]")
'   Call SP_PRCxTerminalRepOnToOff 'Transfer ของข้อมูล Terminal จาก server มายัง local                   '*Ao 2013/07/11 ย้ายไปทำหลัง    ServiceOnOff.exe
'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_EXECxStoreprocedure[2]")

    '*Tao 60-02-03 หาไอพี ของเครื่องตัวเอง เพื่อเอาไปเปรียบเทียบ subnet ของไอพีเครื่อง server
    If tVB_CNIP = "" And tVB_CNComName <> "" Then
        tVB_CNIP = AP_GetIPFromHostName(tVB_CNComName)
        If tVB_CNIP = tVB_CNComName Then tVB_CNIP = AP_ResolveIP(tVB_CNComName)
    End If
    '============================
    
'*KT 54-04-11 V.3.0.56 ย้ายไปหลังการ Connect DB Server และ Local ก่อนเรียก serveronoff เพราะมี Bacth File เคลียร์ช้อมูลใน AdaPosMall", "POSFront", "SrvIP"
   tVB_CNSrvIP = SP_GETtIPAddressSrv(tVB_CNIP)  'SP_GETxLookupIPAddress(oDB.OnComName)      '*Ao 2013/05/03
   If tVB_CNSrvIP <> "" Then
        oDB.OnComIP = tVB_CNSrvIP
    Else
        tVB_CNSrvIP = oDB.OnComIP
    End If
   oDB.OffComIP = tVB_CNIP
   If tVB_CNSrvIP <> "" Then Call SaveSetting("AdaPosMall", "POSFront", "SrvIP", tVB_CNSrvIP)
   'Call SaveSetting("AdaPosMall", "POSFront", "TmnNum", tVB_CNTerminalNum)

'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_PRCxTerminalRepOnToOff[1]")
    'เปิด EXE ที่ทำหน้าที่เป็น Service ในการเก็บค่า Offline Online
    '*KT 52-07-15 ต้องทำการ Run ServiceOnOff.exe ก่อน เพราะ Flag Online ใน Regedit จะต้องนำไปใช้ในส่วนต่างๆ เช่น ServiceAutoReplicate.exe
    '-------------------------------------------------------------------------------------
    Call SP_STAxSplashByAct("Run ServiceOnOff.exe")
    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceOnOff.exe"
    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
            tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
            tFile = tFile + "\POS\ServiceOnOff.exe"
            If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
                Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Can not run ServiceOnOff.exe" & " Form:mDB" & " Control:General" & " Events:Main")
                Call SP_APPxEndProgram
            End If
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "POS\ServiceOnOff.exe[2]")
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Run ServiceOnOff.exe[1]")
    '*KT 51-07-06 ทำการ Cloe ServiceOnOff.EXE
    '------------------------------------------------------------------------------
       
    nVB_lServiceOnOffWnd = FindWindow(vbNullString, "ServiceOnOff")
    'Shell tFile, vbHide
    If nVB_lServiceOnOffWnd = 0 Then
        Call ShellEx(nVB_lServiceOnOffWnd, "open", tFile, "", "", 10)
    Else
        '*KT 52-07-15 กรณีมีการสลับใช้ POSFront หลายๆ Phase เช่น 1.8 กับ 2.0 เป็นต้น
        Call PostMessage(nVB_lServiceOnOffWnd, WM_CLOSE, 0, 0)
        DoEvents
        SleepEx 500, 0
        Call ShellEx(nVB_lServiceOnOffWnd, "open", tFile, "", "", 10)
    End If
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Run ServiceOnOff.exe[1]")
'    DoEvents '*KT 53-02-03
    '-----------------------------------------------------------------------------------
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "ChangeSystemDate[1]")
    '*Eaw 56-10-24 ย้ายไปไว้ข้างล่าง
'    dVB_CurrentDate = SP_GETdCurrentDate
''   If ocnAll.ConnectionString <> oDB.ocnOffLine.ConnectionString Then 'PH 2.0.0 OLD
'    '*Ao 2011/03/24   Service เช็ค ON LINE ไม่ทัน
    Call SleepEx(5000, 0) 'ให้ Relay ประมาณ 5 วินาทีเพื่อให้ ServiceOnOff ทำงานเสร็จก่อน
'    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
'        Call SP_SYSxChangeSystemDate    '*KT 51-08-08  Communication Sheet ML-SF-434 Date On Terminal.xls เปลี่ยน System date ของเครื่อง POS ให้ตรงกันกับ Server
'        bVB_CNSysDateLocal = True
'    Else
'        bVB_CNSysDateLocal = False
'    End If
'-------------------------------------------------
    
    '*Tao 59-05-31 เวลาTest ไม่ต้อง Run ServiceAutoReplicate
    bRunAutoRep = CBool(GetSetting("AdaPosMall", "POSFront", "RunAutoRep", True))     '*Tao 58-05-31 จะ Run ServiceAutoReplicate หรือไม่
    If bRunAutoRep Then

        'เปิด EXE ที่ทำหน้าที่เป็น Service ในการ Claer File
        'PH 2.0.0
        '-------------------------------------------------------------------------------------
        
        Call SP_STAxSplashByAct("Run ServiceAutoReplicate.exe")
        Call SaveSetting("AdaPosMall", "POSFront", "RepSta", "0") '*KT 52-12-19 เริ่มค่า initial
        Call SaveSetting("AdaPosMall", "POSFront", "RepStop", "0") '*KT 52-07-27 เริ่มให้ ServiceAutoReplicate.EXE ทำงาน
        SleepEx 5000, 0 'ให้ Relay ประมาณ 5 วินาทีเพื่อให้ ServiceOnOff ทำงานเสร็จก่อน
        tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceAutoReplicate.exe"
        If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
                tFile = tFile + "\POS\ServiceAutoReplicate.exe"
                If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                    Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
                    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Can not run ServiceAutoReplicate.exe" & " Form:mDB" & " Control:General" & " Events:Main")
                    Call SP_APPxEndProgram 'ถ้าตัวที่ทำหน้าที่ Replicate ทำงานไม่ได้ต้องไม่ให้ทำการขาย
                End If
        End If
        bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    '    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "POS\ServiceAutoReplicate.exe[2]")
        
    '    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "ServiceAutoReplicate.exe[1]")
        '*KT 52-05-22 ทำการ Cloe ServiceAutoReplicate.EXE
        '------------------------------------------------------------------------------------
        nVB_lServiceAutoReplicateWnd = FindWindow(vbNullString, "ServiceAutoReplicate")
        If nVB_lServiceAutoReplicateWnd = 0 Then
            Call ShellEx(nVB_lServiceAutoReplicateWnd, "open", tFile, "", "", 10)
        Else
            '*KT 52-07-15  กรณีมีการสลับใช้ POSFront หลายๆ Phase เช่น 1.8 กับ 2.0 เป็นต้น
            If GetSetting("AdaPosMall", "POSFront", "RepSta") = "1" Then
                Call PostMessage(nVB_lServiceAutoReplicateWnd, WM_CLOSE, 0, 0)
                DoEvents
                SleepEx 1000, 0
                Call ShellEx(nVB_lServiceAutoReplicateWnd, "open", tFile, "", "", 10)
            End If
        End If
    '    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "ServiceAutoReplicate.exe[2]")
    '    DoEvents '*KT 53-02-03
        
    '    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Replicate data[1]")
        'ตรวจสอบว่า ServiceAutoReplicate.exe ทำงานได้สำเร็จหรือไม่
        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
    '        Call SP_STAxSplashByAct("Replicate data", 1, 700)
            Call SP_STAxSplashByAct("Replicate data", 1, 200)
        Else
            Call SP_STAxSplashByAct("Replicate data")
        End If
        tSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
        bPass = False
        Do While Not bPass
            If tSta = "" Then
                SleepEx 1000, 0
                tSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
            Else
                If tSta = "0" Then
                    bVB_CNFlagRepSta = False
                    SleepEx 1000, 0
                    tSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
                Else
                    bVB_CNFlagRepSta = True
                    bPass = True
                End If
            End If
        Loop

        bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
        '   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Replicate data[2]")
        
        '   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "GetSetting:Replicate data[1]")
        If GetSetting("AdaPosMall", "POSFront", "RepOk") = "0" Then
            Call SP_MSGnShowing("tms_CN005210", nCS_Warning)
        End If
        Call SaveSetting("AdaPosMall", "POSFront", "RepStop", "1") '*KT 52-07-27 เริ่มให้ ServiceAutoReplicate.EXE หยุดทำงาน
        '-----------------------------------------------------------------------------------
        '   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "GetSetting:Replicate data[2]")
    Else
        '*Tao 59-05-31 เวลาTest
        Call SaveSetting("AdaPosMall", "POSFront", "RepSta", "1")
        Call SaveSetting("AdaPosMall", "POSFront", "RepStop", "1")
        Call SleepEx(5000, 0) '*Tao 61-05-16 ให้ Relay ประมาณ 5 วินาที แทนการ Replicate
    End If

    Call SP_PRCxTerminalRepOnToOff 'Transfer ของข้อมูล Terminal จาก server มายัง local   '*Ao 2013/07/11  ย้ายมาจากข้างบน
    Call SP_PRCxVirtualTerminalRepOnToOff '*Eaw 56-09-30 Copy VirtualPos ทุกหมายเลขมาลงที่ Offline
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "POS\ServiceAutoClear.exe[1]")
    'เปิด EXE ที่ทำหน้าที่เป็น Service ในการ Claer File
    '-------------------------------------------------------------------------------------
    Call SP_STAxSplashByAct("Run ServiceAutoClear.exe")
    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceAutoClear.exe"
    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
            tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
            tFile = tFile + "\POS\ServiceAutoClear.exe"
            If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
            End If
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "POS\ServiceAutoClear.exe[2]")
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "ServiceAutoClear.exe[1]")
    '*KT 52-05-22 ทำการ Cloe ServiceAutoClear.EXE
    '------------------------------------------------------------------------------
    nVB_lServiceAutoClaerWnd = FindWindow(vbNullString, "ServiceAutoClear")
    If nVB_lServiceAutoClaerWnd = 0 Then
        Call ShellEx(nVB_lServiceAutoClaerWnd, "open", tFile, "", "", 10)
    End If
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "ServiceAutoClear.exe[2]")
'    DoEvents '*KT 53-02-03
    '-----------------------------------------------------------------------------------
    
   'เปิด EXE ที่ทำหน้าที่เป็น ServiceTranfer
    '-------------------------------------------------------------------------------------
    nVB_CNWaitTranfer = Val(Format(GetSetting("AdaPosMall", "POSFront", "WaitTime")))
    If nVB_CNWaitTranfer = 0 Then
        Call SaveSetting("AdaPosMall", "POSFront", "WaitTime", "300")
        nVB_CNWaitTranfer = 300
    End If
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "POS\ServiceTranfer.exe[1]")
    Call SP_STAxSplashByAct("Run ServiceTranfer.exe")
    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceTranfer.exe"
    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
            tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
            tFile = tFile + "\POS\ServiceTranfer.exe"
            If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
                Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Can not run ServiceTranfer.exe" & " Form:mDB" & " Control:General" & " Events:Main")
                Call SP_APPxEndProgram
            End If
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "POS\ServiceTranfer.exe[2]")
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "ServiceTranfer.exe[1]")
    '*KT 51-07-06 ทำการ Cloe ServiceTranfer.EXE
    '------------------------------------------------------------------------------
    nVB_lServiceTransWnd = FindWindow(vbNullString, "ServiceTranfer")
     If nVB_lServiceTransWnd = 0 Then
    'Shell tFile, vbHide
        Call ShellEx(nVB_lServiceTransWnd, "open", tFile, "", "", 10)
    Else
        '*KT 52-07-15 กรณีมีการสลับใช้ POSFront หลายๆ Phase เช่น 1.8 กับ 2.0 เป็นต้น
        If GetSetting("AdaPosMall", "POSFront", "SentSale") = "0" Then 'ระหว่างที่ทำการปิด ServiceTranfer.exe นั้นต้องดูก่อนว่าปัจจุบัน ServiceTranfer.exe  ทำงานแล้วหรือไม่ ถ้าไม่ทำงานถึงจะปิด exe แล้วเปิดใหม่
            Call PostMessage(nVB_lServiceTransWnd, WM_CLOSE, 0, 0)
            DoEvents
            SleepEx 1000, 0
            Call ShellEx(nVB_lServiceTransWnd, "open", tFile, "", "", 10)
        End If
    End If
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "ServiceTranfer.exe[2]")
'    DoEvents '*KT 53-02-03
     '-------------------------------------------------------------------------------------
     Dim nRet As Long '*TON 58-12-15
    'tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\AutoDisableUSB.exe"
    tFile = App.Path + "\AutoDisableUSB.exe"
     If SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
            nRet = Shell(tFile, vbMinimizedFocus)
    End If
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SentSale[1]")
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")
    Else
        Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "0")
    End If
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SentSale[2]")
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_PRCxSaveDataDBOnOffToTmn[1]")
    Call SP_PRCxSaveDataDBOnOffToTmn
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_PRCxSaveDataDBOnOffToTmn[2]")
    
    'Call UT_GETbOption
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_DEFxGetFmtSystem[1]")
'*KT 54-09-15 V.4.0.1 ML-PRV-068 ถ้าไม่สามารถหาค่าตัว Config สำหรับการ Running เลขที่ Ticket
    If Not SP_DEFxGetFmtSystem("TPSTSalHD", "FTShdTransNo") Then
        Call SP_APPxEndProgram
    End If
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_DEFxGetFmtSystem[2]")
    
    'bUT_TchScrn = False
    nUT_ComPile = nCS_ComEnt             'enterprise version
'    DoEvents '*KT 53-02-03
    tVB_Transaction = "99"
    tVB_TransactionSub = "99"

    tUT_SaleDate = UT_tStringDate(SP_GETdSaleDate)
    
    tVB_EJDate = tUT_SaleDate               '*Ao 2014/02/03   เก็บ Operation Date  '*Ao 2014/09/29   V 4.2.149 CommSheet TK-ISS3000-557
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_GETtStoreData[1]")
    Call SP_STAxSplashByAct("Get variable")
    'Call SP_GETtStoreData '*Tao 62-02-12 ของเดิม
    If Not SP_GETtStoreData Then '*Tao 62-02-12 ถ้าอ่าน StoreCode ไม่ได้ ให้ปิดโปรแกรม
        Call SP_MSGnShowing("tms_CN005259", tCS_CNMsgErr)
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "SP_GETtStoreData" & " Form:mDB" & " Control:General" & " Events:Main")
        Call SP_APPxEndProgram
    End If
    If Not SP_GETbTmnData Then
        Call SP_MSGnShowing("tms_CN005055", tCS_CNMsgErr)
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "SP_GETbTmnData" & " Form:mDB" & " Control:General" & " Events:Main")
        Call SP_APPxEndProgram
    End If
    '*Eaw 56-09-09 CommSheet TK-ISS3000-327 ถ้า physical POS นี้ตรงกับ Virtual Pos ระบบจะไม่ให้ใช้ Physical Pos นี้
    If SP_CHKbVirtualPOS Then
        tMsg = "Tenant POS ID " & tVB_CNTerminalNum & " is registered as Virtual POS"
        tMsg = tMsg & vbCrLf & "System not be allowed to access the server data.Please contact your Administrator."
        Call SP_MSGnShowing("tms_CN005055" & ";" & tMsg, tCS_CNMsgErr)
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "SP_CHKbVirtualPOS" & " Form:mDB" & " Control:General" & " Events:Main")
        Call SP_APPxEndProgram
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_GETbTmnData[2]")
    
    tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
    tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
    tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
    tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
    tVB_CNTblDTTmp = "TPSDT" & tVB_CNTerminalNum & "Tmp"
    tVB_CNTblCDTmp = "TPSCD" & tVB_CNTerminalNum & "Tmp"
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "UT_OPTxGetColor[1]")
    Call UT_OPTxGetColor
    Call SP_GETxVariable
    Call SP_GETxVariableAlw '*KT 51-04-22 Comsheet Comsheet ML-SF-185
    Call SP_GETxBanknoteAlw                  '*AO51-03-27
    Call SP_GETxBanknoteRndAlw
'     Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "UT_OPTxGetColor[2]")
    Call SP_GETxMainMnu
    Call SP_GETxMainMnuAll
    Call SP_GETxTrMnu
    Call SP_GETxTrMnuAll
    Call SP_GETxDiscMnu
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "UT_OPTxGetColor[3]")
    Call SP_GETxTenderMnu
    Call SP_GETxTransMode
    Call SP_GETxNameAbb
    Call SP_PRCxResetTimeOut
    Call SP_GETxFunMnu                      '*Aoo 2012/07/31 PRJ10001TK-External Specification Function Menu-00.02.00.doc
    Call SP_GETxWalletType '*Tao 60-08-31
    'Call SP_CRTxLogPrg("SP_GETxVariable success.")
    
    '*Tao 62-09-17 ปิดไว้ ไม่ต้อง Initial ก็ได้
'    Call SP_STAxSplashByAct("Initial MSSoap")
'    Call SP_CRTxLogPrg("SP_SETxTokenObjects Start.")
'    Call SP_SETxTokenObjects '*Bmp 61-03-06 Token
'    Call SP_CRTxLogPrg("SP_SETxTokenObjects End.")

    tCheckerMode = GetSetting("AdaPosMall", "POSFront", "CheckerMode")
    If tCheckerMode = "" Then
        Call SaveSetting("AdaPosMall", "POSFront", "CheckerMode", "0")
    End If
    bVB_CheckerMode = IIf(GetSetting("AdaPosMall", "POSFront", "CheckerMode") = 0, True, False)
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "UT_OPTxGetColor[4]")

    '*Nirut 08-03-28 add  tmn
    Call SP_STAxSplashByAct("Check online")
    oTerminal.SaleDate = CDate(tUT_SaleDate)
    oTerminal.Connect oDB.OnServerName, oDB.OnUsrName, oDB.OnUsrPwd, oDB.OnDbName
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    
    '*Nirut 08-03-26 edit
    'continue sale with unclose Tmn - - - - - - - - - - - - - - - - - - - - -- - - - - - - - - - - - - - - - - - - - -
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Check online[1]")
    'Call SP_PRCxTruncateDps
    '*KT 52-01-24 Communication Sheet PH2-ML-SF-061 When open store program didn't let user logon.xls
    If bVB_CNOpenTmn And oTerminal.Online And oTerminal.StoreOpen And bVB_CNOpenTmnOff Then
            oTerminal.SaleMode = oTerminal.AvailableSaleMode
            oTerminal.SaleDate = CDate(tUT_SaleDate)
            Call W_PRCxGetUsrHistory 'อยู่ขั้นตอนสุดท้ายเพราะต้องได้วันที่ขายที่ถูกต้อง
    ElseIf bVB_CNOpenTmnOff And oTerminal.Online And Not oTerminal.StoreOpen Then   'ก่อนหน้ามีการเข้าขายแบบ Offline แล้วโปรแกรมหลุด เข้าใหม่เป็นแบบ Online ยังไม่ Open Store
            Select Case Val(Format(tVB_TransactionMode))
                Case 0, 1
                    oTerminal.SaleMode = nEN_OfflineNormal
                    oTerminal.SaleDate = SP_GETdCurrentDate(True)
                    tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) '*KT 51-12-15 ต้องมีการปรับวันที่ Operation Date เมื่อต้องขายแบบ Offline(Store Close) หรือไม่
                    Set oTransPrn.ActiveConnection = oDB.ocnOffLine
                    oTransPrn.ExportFolder = tVB_CNPthBnkTr
                    oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
                Case 2
                    oTerminal.SaleMode = nEN_OfflineOpenWith
                    oTerminal.SaleDate = SP_GETdCurrentDate(True)
                    tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) 'adjust salesdate with opration date config
                    Set oTransPrn.ActiveConnection = oDB.ocnOffLine
                    oTransPrn.ExportFolder = tVB_CNPthBnkTr
                    oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
                Case Else
'กรณีเป็น Else ให้ Default เป็น nEN_OfflineNormal น่าจะให้เกิด  PH 2.0 ML-SF-052 Operation Date ในTransaction ผิด ในกรณีที่เครื่อง Offline
                    oTerminal.SaleMode = nEN_OfflineNormal
                    oTerminal.SaleDate = SP_GETdCurrentDate(True)
                    tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) '*KT 51-12-15 ต้องมีการปรับวันที่ Operation Date เมื่อต้องขายแบบ Offline(Store Close) หรือไม่
                    Set oTransPrn.ActiveConnection = oDB.ocnOffLine
                    oTransPrn.ExportFolder = tVB_CNPthBnkTr
                    oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
            End Select
            Call W_PRCxGetUsrHistory 'อยู่ขั้นตอนสุดท้ายเพราะต้องได้วันที่ขายที่ถูกต้อง
    ElseIf Not oTerminal.Online And bVB_CNOpenTmnOff Then 'ก่อนหน้ามีการเข้าขายแบบ Offline แล้วโปรแกรมหลุด เข้าใหม่เป็นแบบ Offline
            Select Case Val(Format(tVB_TransactionMode))
                Case 0, 1
                    oTerminal.SaleMode = nEN_OfflineNormal
                    oTerminal.SaleDate = SP_GETdCurrentDate(True)
                    tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) '*KT 51-12-15 ต้องมีการปรับวันที่ Operation Date เมื่อต้องขายแบบ Offline(Store Close) หรือไม่
                    Set oTransPrn.ActiveConnection = oDB.ocnOffLine
                    oTransPrn.ExportFolder = tVB_CNPthBnkTr
                    oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
                Case 2
                    oTerminal.SaleMode = nEN_OfflineOpenWith
                    oTerminal.SaleDate = SP_GETdCurrentDate(True)
                    tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) 'adjust salesdate with opration date config
                    Set oTransPrn.ActiveConnection = oDB.ocnOffLine
                    oTransPrn.ExportFolder = tVB_CNPthBnkTr
                    oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
                Case Else
'กรณีเป็น Else ให้ Default เป็น nEN_OfflineNormal น่าจะให้เกิด  PH 2.0 ML-SF-052 Operation Date ในTransaction ผิด ในกรณีที่เครื่อง Offline
                    oTerminal.SaleMode = nEN_OfflineNormal
                    oTerminal.SaleDate = SP_GETdCurrentDate(True)
                    tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) '*KT 51-12-15 ต้องมีการปรับวันที่ Operation Date เมื่อต้องขายแบบ Offline(Store Close) หรือไม่
                    Set oTransPrn.ActiveConnection = oDB.ocnOffLine
                    oTransPrn.ExportFolder = tVB_CNPthBnkTr
                    oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
            End Select
            Call W_PRCxGetUsrHistory 'อยู่ขั้นตอนสุดท้ายเพราะต้องได้วันที่ขายที่ถูกต้อง
    ElseIf oTerminal.Online And bVB_CNOpenTmnOff Then 'ก่อนหน้ามีการเข้าขายแบบ Offline แล้วโปรแกรมหลุด เข้าใหม่เป็นแบบ Online
            oTerminal.SaleMode = oTerminal.AvailableSaleMode
            oTerminal.SaleDate = CDate(tUT_SaleDate)
            Call W_PRCxGetUsrHistory 'อยู่ขั้นตอนสุดท้ายเพราะต้องได้วันที่ขายที่ถูกต้อง
    Else
            '*Nirut 08-03-26 add
            'continue sale with  open24 hour mode - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            tVB_LSaleDate = IIf(tVB_LSaleDate = vbNullString, tUT_SaleDate, tVB_LSaleDate)
            If Not bVB_CNOpenTmn And bVB_CNOpenTmnOff And IsDate(tVB_LSaleDate) And tVB_LSaleDate >= tUT_SaleDate Then
'                    Call SP_PRCxSwithToOffline 'PH 2.0.0 OLD
                    oTerminal.SaleMode = nEN_OfflineOpenWith
                    oTerminal.SaleDate = SP_GETdCurrentDate(True)
                    tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) 'adjust salesdate with opration date config
                    Set oTransPrn.ActiveConnection = oDB.ocnOffLine
                    oTransPrn.ExportFolder = tVB_CNPthBnkTr
                    oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
                    Call W_PRCxGetUsrHistory 'อยู่ขั้นตอนสุดท้ายเพราะต้องได้วันที่ขายที่ถูกต้อง
            Else
                    'none
                    Dim dDateReEntry As Date
                    Select Case Val(Format(SP_GETtLastTransType(dDateReEntry)))
                        Case tEN_TrnSaleTaxInc, tEN_TrnReturnTaxInc, tEN_TrnTotalReportTaxInc, tEN_TrnOverShortReportTaxInc, tEN_TrnOpenTaxInc
                                oTerminal.SaleMode = nEN_OnlineNormal
                                oTerminal.SaleDate = dDateReEntry
                                tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) 'adjust salesdate with opration date config
                                Set oTransPrn.ActiveConnection = oDB.ocnOffLine
                                oTransPrn.ExportFolder = tVB_CNPthBnkTr
                                oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
                                Call W_PRCxGetUsrHistoryReEntry("38")
                                Call SP_PRCxTerminalRepOnToOff(True)
                                '-----------------------------------------------------------------------------------
                                tVB_CNVirTerminalNum = tVB_CNTerminalNum  '*BG 56-10-31 CommSheet TK-ISS3000-410_Bung.xls
                                tVB_CNReEntryDate = tUT_SaleDate
                                tVB_CNMnuCurrent = ""
                                '_______________________________________________
                                
                        Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax, tEN_TrnTotalReportManualTax, tEN_TrnOverShortReportManualTax, tEN_TrnOpenManualTax
                                oTerminal.SaleMode = nEN_OnlineNormal
                                oTerminal.SaleDate = dDateReEntry
                                tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) 'adjust salesdate with opration date config
                                Call SaveSetting("AdaPosMall", "POSFront", "SaleDate", tUT_SaleDate)                    '*Ao 2014/09/29   V 4.2.149 CommSheet TK-ISS3000-557
                                Set oTransPrn.ActiveConnection = oDB.ocnOffLine
                                oTransPrn.ExportFolder = tVB_CNPthBnkTr
                                oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
                                Call W_PRCxGetUsrHistoryReEntry("43")
                                Call SP_PRCxTerminalRepOnToOff(True)
                                '------------------------------------------------------------------------------------
                                tVB_CNVirTerminalNum = tVB_LastPosNo   'tVB_CNTerminalNum '*BG 56-10-31 CommSheet TK-ISS3000-410_Bung.xls
                                tVB_CNReEntryDate = tUT_SaleDate
                                tVB_CNMnuCurrent = ""
                                '------------------------------------------------------------------------------------
                    End Select
            End If
    End If
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Check online[2]")

'    tVB_EJDate = tUT_SaleDate               '*Ao 2014/02/03   เก็บ Operation Date
    
    '*KT 52-01-21 ทำการ set ค่ารอบกรณีที่มีการเปิดโปรแกรมครั้งของวันนัน้
    If tVB_LSaleDate <> tUT_SaleDate Then
        nVB_CNLastLogin = 0
    End If
    
   tVB_CNOnCaption = SP_INItGetLang("ALL", "tcs_CN005025")
   tVB_CNOffCaption = SP_INItGetLang("ALL", "tcs_CN005026")
   tVB_CNComName = cUT.UT_SYStGetComName
   '*KT 52-05-19 PH 1.8 ML-SF-069
   'TP2000 Screen 1. เมื่อ กด X (QTY Change) รายการสินค้ารายการสุดท้ายปัจจุบันจะ เลื่อน ขึ้นไปบรรทัดแรก และรายการสินค้าก่อนหน้านั้นจะมองไม่เห็น 2. เมื่อ Scan สินค้าจนเต็มหน้าจอ หลังจกนั้น scan สินค้าตัวต่อไป
   'รายการสินค้าจะไม่ Refresh ลงมายัง Grid ด้านล่างต้อง Scan สินค้าตัวต่อไปจึงจะมาแสดงยังบรรทัดต่อไป
    If bVB_Res640X480 Then
        nUT_GridMaxRow = 11
    Else
        nUT_GridMaxRow = 12
    End If
   tVB_CNRetVATInOrEx = "1" 'ภาษีแยกนอก
   
'*KT 54-04-11 V.3.0.56  ย้ายไปหลังการ Connect DB Server และ Local ก่อนเรียก serveronoff เพราะมี Bacth File เคลียร์ช้อมูลใน AdaPosMall", "POSFront", "SrvIP"
'   tVB_CNSrvIP = SP_GETtIPAddressSrv(tVB_CNIP)  'SP_GETxLookupIPAddress(oDB.OnComName)
'   oDB.OnComIP = tVB_CNSrvIP
'   oDB.OffComIP = tVB_CNIP
'   If tVB_CNSrvIP <> "" Then Call SaveSetting("AdaPosMall", "POSFront", "SrvIP", tVB_CNSrvIP)
'   Call SaveSetting("AdaPosMall", "POSFront", "TmnNum", tVB_CNTerminalNum)
   
    'Slip Message
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_FILxSlpMagToArray[1]")
    Call SP_FILxSlpMagToArray(aVB_HSlipMag, aVB_TSlipMag, aVB_DSlipMag)
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_FILxSlpMagToArray[2]")
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "W_CLRxClearDataHD_DT[1]")
    Call W_CLRxClearDataHD_DT
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "W_CLRxClearDataHD_DT[2]")
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Check printer[1]")
    '*Nirut 08-03-25
    'set  for recovery when offline  and transaction back up
    tVB_PrnDevice = tVB_CNTerminalDrv
    Call SP_STAxSplashByAct("Check printer")
    Set oTransPrn.ActiveConnection = ocnAll
    oTransPrn.ExportFolder = tVB_CNPthBnkTr
    oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
    oTransPrn.bPrnByOPOS = bVB_CNPrnOPOS 'True : Print by OPOS    Fale : Print by OPOS                  '* Ao 51-06-16 เกิด Type mismatch  มัน Return Not Found
    oTransPrn.bPRMGenEJ = False '*KT 51-12-16 Flag ที่บอกว่ามีการเรียกใช้งาน oTransPrn จาก Tools Generate EJ
    If bVB_CNPrnOPOS Then 'Print By OPOS
        '*KT 55-03-16 การอ่านค่า Line_Chars  จาก Register Config OPOS นั้น ถ้ามี OPOS มากกว่า 1 ประเภท อาจจะไม่ได้ใช้ Key "Line_Chars" เลยเปลี่ยนให้เป็น Function เพื่อ Error จะได้ Return มาเป็น 42
        '*Ao 2012/12/20  PRINTER 360
'        MsgBox tVB_PrnDevice
'        Select Case tVB_PrnDevice                                                                                           '*Ao 2013/01/24   Print FRONTTECH
'            Case "FP360SERPRT", "FP360PARPRT", "FP360USBPRT"
'                Set oOpos = CreateObject("AdaOpos360.cFUJITSU")
'                tVB_CashPrinter = tVB_PrnDevice
''                MsgBox "AdaOpos360.cFUJITSU"
'            Case Else
'                Set oOpos = CreateObject("AdaOpos.cFUJITSU")
'        End Select
        nVB_MaxChar = W_GETnLineCharsOPOS() 'ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\OLEforRetail\ServiceOPOS\POSPrinter\FP410PRT\Receipt\", "Line_Chars") '42
        If ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\OLEforRetail\ServiceOPOS\POSPrinter\", "DefaultPOSPrinter") = "" Then
            Call SP_MSGnShowing("tms_CN005155", nCS_Warning)
            Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Printer OPOS Error" & " Form:mDB" & " Control:General" & " Events:Main")
            Call SP_APPxEndProgram
        End If
        Call SP_STAxSplashByAct("Open printer by OPOS")
        oOpos.OpenPrint tVB_CashPrinter, nVB_MaxChar, nVB_CashCodePage 'Open Printer OPOS
        bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
        Call SP_STAxSplashByAct("Check Status TurnOff by OPOS")
        If oOpos.PowerState = 2004 Then
            bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
             wCNMsgErr.Show vbModal
        ElseIf oOpos.oPrinterState = 4 Then
            bVB_CNSplash = False
            wCNMsgErr.Show vbModal
        Else
            Call SP_STAxSplashByAct("Check CoverOpen by OPOS")
            If oOpos.CoverOpen Then
                bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
                wCNMsgErr.Show vbModal
            End If
        End If
        oTransPrn.nW_StatePrn = 1
        bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    Else ''Print By Print
        nVB_MaxChar = 35 '*KT 52-05-04 Default กรณีที่ Printer ไม่ใช้ OPOS แต่ใช้ Driver เพราะใน Module การ print ต้องมีการใช้ nVB_MaxChar
        If tVB_PrnDevice = "" Then
            Call SP_MSGnShowing("tms_CN005156", tCS_CNMsgErr)
            Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Printer Driver Error1" & " Form:mDB" & " Control:General" & " Events:Main")
            Call SP_APPxEndProgram
        Else
            Set oTransPrn.ActivePrinter = oPrn.SelectPrinter(tVB_PrnDevice)
            If Not oPrn.bState Then
                    Call SP_MSGnShowing("tms_CN005156", tCS_CNMsgErr)
                    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Printer Driver Error2" & " Form:mDB" & " Control:General" & " Events:Main")
                    Call SP_APPxEndProgram
            End If
        End If
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น

  '*KT 55-07-05 Takashimaya สร้าง oOposSlip ยังไม่มี Process การพิมพ์ Slip ด้วย Driver              '*Ao 2013/01/24   Print FRONTTECH
   If bVB_CNPrnOPOS Then 'Print By OPOS
        If ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\OLEforRetail\ServiceOPOS\POSPrinter\", "AlternatePOSPrinter") = "" Then
            Call SP_MSGnShowing("tms_CN005155", nCS_Warning)
            Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Slip Printer OPOS Error" & " Form:mDB" & " Control:General" & " Events:Main")
            Call SP_APPxEndProgram
        End If
        Call SP_STAxSplashByAct("Open slip printer by OPOS")
'        nVB_MaxCharSlip = 42
        nVB_MaxCharSlip = 40                '*Ao 2013/04/30  CommSheet TK-ISS3000-095 00.00.04.xls
        oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
        oOposSlip.SlpLineChars = nVB_MaxCharSlip
        bVB_CNSplash = False
        Call SP_STAxSplashByAct("Check Status TurnOff by OPOS")
'        MsgBox oOposSlip.PowerState
'        MsgBox oOposSlip.oPrinterState
'        MsgBox oOposSlip.CoverOpen
        If oOposSlip.PowerState = 2004 Then
            bVB_CNSplash = False
            wCNMsgErrSlip.bW_Wait = False
             wCNMsgErrSlip.Show vbModal
        ElseIf oOposSlip.oPrinterState = 4 Then
            bVB_CNSplash = False
            wCNMsgErrSlip.bW_Wait = False
            wCNMsgErrSlip.Show vbModal
        Else
            Call SP_STAxSplashByAct("Check CoverOpen by OPOS")
            If oOposSlip.CoverOpen Then
                bVB_CNSplash = False
                wCNMsgErrSlip.bW_Wait = False
                wCNMsgErrSlip.Show vbModal
            End If
        End If
        oTransPrn.nW_StatePrn = 1
        bVB_CNSplash = False
    End If
    bVB_CNSplash = False
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Open Line Display by OPOS[1]")
    Call SP_STAxSplashByAct("Open Line Display by OPOS")
    If ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\OLEforRetail\ServiceOPOS\LineDisplay\", "DefaultLineDisplay") = "" Then
        '*KT 52-05-04 ให้ทำการเอาข้อความเตือนกรณีที่ OPOS ของเครื่อง TP ไม่ทำงาน ซึ่งได้ทำการแจ้งข้อเสียให้พี่บัมฟังแล้ว แต่ทาง FSBT ยังยืนยันให้เอา Message ออก ซึ่งจะมีผลทำให้ระบบไม่สามารถจับ
        'กรณีที่ Line Display ไม่ทำงาน ทาง FSBT จะเป็นผู้จัดการเอง
       ' PH1.8-ML-SF-022 (TP2000) During start application for TP2000, system's always shown woring message for Device Interface such as Customer Display, Scanner & MSR)
        'Call SP_MSGnShowing("tms_CN005163", nCS_Warning)
        bVB_OPOSLineDisplayEnable = False
    Else
        oOpos.OpenDisplay tVB_CashDisplay
        If oOpos.DisplayDeviceEnabled Then
            '106 OPOS_E_ILLEGAL,107 OPOS_E_NOHARDWARE ใช้ oOpos 1.0.15 ถึงจะใช้ oOpos.CheckHealth ได้
'            If oOpos.CheckHealth(1) = 107 Then
'                bVB_OPOSLineDisplayEnable = False
'            Else
                bVB_OPOSLineDisplayEnable = True
'            End If
        Else
             bVB_OPOSLineDisplayEnable = False
            '*KT 52-05-04 ให้ทำการเอาข้อความเตือนกรณีที่ OPOS ของเครื่อง TP ไม่ทำงาน ซึ่งได้ทำการแจ้งข้อเสียให้พี่บัมฟังแล้ว แต่ทาง FSBT ยังยืนยันให้เอา Message ออก ซึ่งจะมีผลทำให้ระบบไม่สามารถจับ
            'กรณีที่ Line Display ไม่ทำงาน ทาง FSBT จะเป็นผู้จัดการเอง
            'PH1.8-ML-SF-022 (TP2000) During start application for TP2000, system's always shown woring message for Device Interface such as Customer Display, Scanner & MSR)
             'Call SP_MSGnShowing("tms_CN005162", nCS_Warning)
        End If
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'   Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Open Line Display by OPOS[2]")
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Open CashDrawer by OPOS[1]")
    Call SP_STAxSplashByAct("Open CashDrawer by OPOS")
    If ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\OLEforRetail\ServiceOPOS\CashDrawer\", "DefaultCashDrawer") = "" Then
'        Call SP_MSGnShowing("tms_CN005164", nCS_Warning) '*BG 56-10-24 TK-ISS3000-386
        bVB_OPOSCashDrawerEnable = False
    Else
        oOpos.OpenOPOSDrawer tVB_CashDrawer, 0
        If oOpos.oCashDrawerDeviceEnabled Then
            Call SleepEx(1000, 0)
            If oOpos.DrawerOpen Then
    '*KT 51-12-18 Phase 2 แก้ไขเรื่องเกิดเหตุการณ์ลิ้นชักเปิดตั้งแต่เปิดโปรแกรมครั้งแรก
    '-----------------------------------------------------------------------------------------------------------
    '           ปัญหาที่เกิดขึ้นเกิดจากตอนเปิดโปรแกรมครั้งแรก เกิดเหตุการณ์ว่าลิ้นชักเปิดอยู่ ระบบจะตรวจสอบลิ้นชัก 2 ขั้นตอน
    '           1 ตรวจสอบใน Regedit ว่ามี OPOS ลิ้นชักหรือไม่จากค่า DefaultCashDrawer
    '           2 ตรวจสอบว่า OPOS ของ CashDrawer มีสถานะ Enable หรือไม่
    '           จากข้อที่ 2 มีการตรวจสอบ 2 Step
    '           2.1 ตรวจสอบสถานะ Enable จาก Propertise DeviceEnabled ว่าเป็น True หรือ False
    '           2.2 ตรวจสอบว่าขณะนั้นลิ้นชักเปิดหรือไม่
    '           จาก Case ที่เกิดขึ้นการตรวจสอบ CashDrawer ผ่านข้อที่ 1 ข้อที่ 2.1 แต่มาติดที่ข้อที่ 2.2 ระบบจะฟ้องข้อความเตือนผิดจากที่ควรจะเตือนว่าลิ้นชักเปิดอยู่เป็น
    '           "CashDrawer OPOS Disable" และระบบยอมปล่อยให้ผ่านได้และมีการ Set ตัวแปร bVB_OPOSCashDrawerEnable = False ซึ่งตัวแปรนี้จะเป็นเงื่อนในการ
    '           แตะลิ้นชักเมื่อมีการจบบิล จะแตะลิ้นชักเมื่อ OPOSCashDrawerEnable = True
    '            bVB_OPOSCashDrawerEnable = False
    '            Call SP_MSGnShowing("tms_CN005165", nCS_Warning)
                 Screen.MousePointer = 0
                 wMsgDrawer.Show vbModal
                 Screen.MousePointer = 11
'                DoEvents '*KT 53-02-03
                 bVB_OPOSCashDrawerEnable = True
    '------------------------------------------------------------------------------------------------------------
            Else
                bVB_OPOSCashDrawerEnable = True
            End If
        Else
            bVB_OPOSCashDrawerEnable = False
             Call SP_MSGnShowing("tms_CN005165", nCS_Warning)
        End If
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Open CashDrawer by OPOS[2]")
    
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "Open Scanner by OPOS[1]")
    Call SP_STAxSplashByAct("Open Scanner by OPOS")
    If ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\OLEforRetail\ServiceOPOS\Scanner\", "DefaultScanner") = "" Then
         Call SP_MSGnShowing("tms_CN005192", nCS_Warning)
         bVB_OPOSScannerEnable = False
    Else
        With wBackGrand.oScanner
            .Open tVB_Scanner
            .Claim 0
            .DeviceEnabled = True
             If .ResultCode <> 0 Then
                bVB_OPOSScannerEnable = False
                Call SP_MSGnShowing("tms_CN005193", nCS_Warning)
            Else
                If .DeviceEnabled Then
                    bVB_OPOSScannerEnable = True
                Else
                    bVB_OPOSScannerEnable = False
                    Call SP_MSGnShowing("tms_CN005193", nCS_Warning)
                End If
            End If
            'Scanner Close
            .DeviceEnabled = False
            .Release
            .Close
        End With
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    
    Call SP_STAxSplashByAct("Open MSR by OPOS")
    bVB_OPOSMSREnable = False
    If ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\OLEforRetail\ServiceOPOS\MSR\", "DefaultMsr") = "" Then
        Call SP_MSGnShowing("tms_CN005194", nCS_Warning)
        bVB_OPOSMSREnable = False
    Else
        oOpos.OpenMSR tVB_CashMsr, 0
       If oOpos.oMSRResultCode <> 0 Then
            bVB_OPOSMSREnable = False
            Call SP_MSGnShowing("tms_CN005195", nCS_Warning)
       Else
            If oOpos.oMSRDeviceEnable Then
                bVB_OPOSMSREnable = True
            Else
                bVB_OPOSMSREnable = False
                 Call SP_MSGnShowing("tms_CN005195", nCS_Warning)
            End If
       End If
        oOpos.CloseMSR
    End If
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    
    Call SP_STAxSplashByAct("Get  Logo")
    Call SP_GEToLogo
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    Set oTransPrn.Logo = oVB_PicLogo
    oPrn.tTmnNum = tVB_CNTerminalNum
    Call SP_STAxSplashByAct("Set Logo")
    Call oPrn.SetBitmap(oVB_PicLogo)
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น

    '*KT 51-09-2008 กรณีที่เปิด Printer ให้ทำการตัดกระดาษแล้วพิมพ์ ใหม่
    If bVB_CNPrnOPOS Then
'        oOpos.Printnormal PTR_S_RECEIPT, Chr(16) & Chr(20) '*KT 51-12-22 Phase 2 Clear Buffer          '*Ao 2013/01/24   Print FRONTTECH
'        If bVB_ImageLogo Then
'            oOpos.TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'end print by OPOS
'            oOpos.Printnormal PTR_S_RECEIPT, "" & Chr(10)
'            Call oPrn.PrintLogo
'            oOpos.TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL 'end print by OPOS
'            oOpos.CutPaper 90
'            oOpos.Printnormal PTR_S_RECEIPT, Chr(16) & Chr(20) '*KT 51-12-22 Phase 2 Clear Buffer
'        Else
'            '*KT 51-10-14 V 3.0.135/V 3.0.200 กรณีที่ไม่มีการ Set Logo
'            '--------------------------------------------------------------
'            oOpos.TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'end print by OPOS
'            oOpos.Printnormal PTR_S_RECEIPT, "" & Chr(10)
'            For nI = 1 To 8
'                oOpos.Printnormal PTR_S_RECEIPT, "" & LF
'            Next nI
'            oOpos.TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL 'end print by OPOS
'            oOpos.CutPaper 90
'            oOpos.Printnormal PTR_S_RECEIPT, Chr(16) & Chr(20) '*KT 51-12-22 Phase 2 Clear Buffer
'            '---------------------------------------------------------------
'         End If
        '-----------------------------------------------------------------
    Else
        '*KT 54-09-16 V.4.0.1 ตัดกระดาษทิ้ง
        For nI = 1 To 8
            Call oPrn.PrintText(vbNullString, 0, , vbLeftJustify, True)  ' space - - - - - -
        Next nI
        Call oPrn.EndDoc(oVB_PicLogo, False)
    End If

    Call SP_STAxSplashByAct("Check Table Closing")
    Call SP_CHKxTblTmpClosing("TPSTClose")
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น

    'move file ej offline to path online
    '*KT 52-05-15 PH 1.8 ML-SF-060
    'EJ ไม่ส่งมายัง Server ขั้นตอน  1. เปิด POS แบบ Online (Open-> Close-->)  2. Disconnect LAN 3. Open POS in Offline 4. Sale --> Close --> Exit 5. Open POS in Online (Open -->Sale-->Close) 6. Print Closing Report --> Exit        EJ ที่ขายแบบ Offline จะไม่ส่งมายัง Server
     'ป้องกัน 1 เปิดโปรแกรม 2 ตอนที่เปิดโปรแกรมแล้วเข้าหน้าจอ Sale ทันที 3 ตอน Exit Terminal 4 ตอนจบบิลแล้วมีการสลับมาที่ On-line
    Call SP_STAxSplashByAct("Move file EJ")
    Call SP_MOVxEJToOnline
'    DoEvents '*KT 53-02-03
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "SP_MOVxEJToOnline[2]")
    
'    Call SP_STAxSplashByAct("Restart service MSSQLSERVER")
''    Call SP_PRCxStopServiceSQL
''    SleepEx 5000, 0
'     Call SP_PRCxStartServiceSQL 'ทำการ restart service ของ SQL EXPRESS 2005
''    SleepEx 5000, 0
''    bPass = False
''    Do While Not bPass
''        SleepEx 1000, 0
''        'bPass = GetServiceStatus()
''        bPass = SP_SQLxConnectService(tVB_CNOffDbSrc, tVB_CNOffDbSQLAuthUsrC, tVB_CNOffDbSQLAuthUsrP)
''    Loop
'    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    DoEvents

    '*KT 54-07-26 V.4.0.1 Phase4.0(Member Point) Promotion Coupon ทำการเก็บ Logo ไว้ใน array
    Call SP_STAxSplashByAct("Get  Logo Coupon")
    Call SP_GEToLogoCpn
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    Call SaveSetting("AdaPosMall", "POSFront", "nLogoCpn", "")
    If bVB_SetImageLogoCpn Then
        Call SP_STAxSplashByAct("Set Logo Coupon")
        Call oPrn.SetBitmapCpn(aVB_PicLogoCpn(0).oPic, aVB_PicLogoCpn(0).tID)
        If bVB_ImageLogoCpn Then
            Call SaveSetting("AdaPosMall", "POSFront", "nLogoCpn", CStr(aVB_PicLogoCpn(0).tPicName))
        End If
        Set oVB_PicLogoCpn = aVB_PicLogoCpn(0).oPic
        bVB_CNSplash = False
    End If
    Call SP_CRTxLogPrg("Get  Logo Coupon End.")

'*KT 54-09-16 V.3.5.1 เอาออกไม่ได้ใช้งานแล้ว
'----------------------------------------------------------------------------------------------------------
''    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "W_PRNbTicketPicture[1]")
'    oTransPrn.nXCPic = 5 / 2
'    oPrn.bPrnByOPOS = bVB_CNPrnOPOS
'    oPrn.tTmnNum = tVB_CNTerminalNum
'    oPrn.LineOffset = -0.12
'    If Not bVB_CNPrnOPOS Then oTransPrn.W_PRNbTicketPicture
''    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "W_PRNbTicketPicture[2]")
'----------------------------------------------------------------------------------------------------------
    
'    Set oRe.ActiveConnection = ocnAll
    If tVB_CNPthReScrip = "" Then tVB_CNPthReScrip = App.Path
    If Not SP_FLDbManFolder(tVB_CNPthReScrip, nEN_Fld1Exist) Then Call SP_FLDbManFolder(tVB_CNPthReScrip, nEN_Fld3Create)
'    oRe.Create tVB_CNPthReScrip & "\ScriptTrans.sql"
'    oRe.FilterTable = "TPSTDepositHD,TPSTDepositDT,TPSTDepositRC,TPSTDepositCD,TPSTVoucherHD,TPSTVoucherRC,TPSTHoldHD,TPSTHoldDT,TPSTHoldRC,TPSTHoldCD," & tVB_CNTblHD & "," & tVB_CNTblDT & "," & tVB_CNTblRC & "," & tVB_CNTblCD & "," & tVB_CNTblDTTmp & "," & tVB_CNTblCDTmp & ",TPSTLogDT"
    Set oFSO = Nothing
    
'    DoEvents 'ห้ามเอาออกไม่นั้น Error '*KT 53-02-03
    bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    Screen.MousePointer = vbDefault
    bUT_SignOut = False
    nVB_HisSelTrn = 0
    bVB_CNPrgFsh = True
    bVB_CNCreateTblTmp = False '*KT 51-10-08 V 3.0.131/V 3.0.200
    bVB_CNCreateTblTmpFrmCloseRpt = False '*KT 51-10-10 V 3.0.132/V 3.0.200
    
    bVB_CNTrnDiscEx = True '*56-05-29 Pae CommSheet TK-ISS3000-191-ST 00 00 01.xls
    '*Eaw 56-10-24
    tStrDate = Date & " " & Time
    'Call SleepEx(5000, 0) 'ให้ Relay ประมาณ 5 วินาทีเพื่อให้ ServiceOnOff ทำงานเสร็จก่อน
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        Call SP_SYSxChangeSystemDate    '*KT 51-08-08  Communication Sheet ML-SF-434 Date On Terminal.xls เปลี่ยน System date ของเครื่อง POS ให้ตรงกันกับ Server
        bVB_CNSysDateLocal = True
    Else
        bVB_CNSysDateLocal = False
    End If
    dVB_CurrentDate = SP_GETdCurrentDate
    Call SP_CRTxWriteLogTerminated("SystemDateBeforeChange : " & tStrDate) 'Write Log เวลาเครื่อง Local ก่อน Get ค่าจาก Server
    tStrDate = dVB_CurrentDate
    Call SP_CRTxWriteLogTerminated("SystemDateAfterChange : " & tStrDate) 'Write Log เวลาเครื่อง Local หลัง Get ค่าจาก Server
    '---------------------------------------------------
    bVB_ChkEmpPoint = True '*Eaw 56-11-21
    bMainFunc = True                                   '*Ao 2014/09/05   V 4.2.143
    tVB_OverSh = "0"
    
    '*Tao 61-01-18 เช็คว่าก่อนปิดโปรแกรม ทำรายการอะไรค้างอยู่
    If Not bVB_TrainingMode Then '*Tao 61-06-26 ต้องไม่ใช่ TrainingMode
    Call SP_CRTxLogPrg("SP_ChkxClosingMissing Start.")
    Select Case SP_ChkxClosingMissing
'    Case "37"
'        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Found Over/Short Report Missing Close Report" & " Form:mDB" & " Control:General" & " Events:Main")
'        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Re Processing Close Report" & " Form:mDB" & " Control:General" & " Events:Main")
'        wCNBegin.ogdFunc.Select 17, 0
'        wCNBegin.ocmOK_Click
'        If SP_CHKxIsFormOpen("wCNBegin") Then Unload wCNBegin
'        bVB_CNCloseTmnSta = True
'        bVB_CNClosingRptSta = False
'        bVB_CNClosingRptAgain = True
'        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Re Processing Closing Report" & " Form:mDB" & " Control:General" & " Events:Main")
'        wCNCloseMnu.ogdFunc.Select 7, 0
'        wCNCloseMnu.ocmOK_Click
'        If SP_CHKxIsFormOpen("wCNCloseMnu") Then Unload wCNCloseMnu
'        tUT_SaleDate = UT_tStringDate(SP_GETdSaleDate)
'        tVB_EJDate = tUT_SaleDate
'        oTerminal.SaleDate = CDate(tUT_SaleDate)
'        bVB_CNCloseTmnSta = False
'        bVB_CNClosingRptSta = False
'        bVB_CNClosingRptAgain = False
'        bUT_SignOut = False
        '*Tao 61-12-27 ยังไม่เสร็จ
'    Case "41"
'        tVB_CNMnuCurrent = "172005"
'        bUT_SignOut = True
'        bVB_CNCloseTmnSta = True
'        bVB_CNOverShortSta = True
'        tVB_Transaction = aVB_TransMode(41).ID
'        tVB_TransactionSub = tVB_Transaction
'        tVB_TransName = aVB_TransMode(41).Name
'        tVB_TransNameABB = aVB_NameAbb(tEN_AbbClosingReportTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbClosingReportTaxInc - 1).UsrLangSec
'        tVB_CNTypePrinter = aVB_TransMode(41).PrintType
'        Call wFunctionKB.W_PrintSignOut(False, False, True)
'        tUT_SaleDate = UT_tStringDate(SP_GETdSaleDate)
'        tVB_EJDate = tUT_SaleDate
'        oTerminal.SaleDate = CDate(tUT_SaleDate)
'        bVB_CNCloseTmnSta = False
'        bVB_CNClosingRptSta = False
'        bVB_CNClosingRptAgain = False
'        bUT_SignOut = False
'        tVB_Transaction = "97"
'        tVB_TransactionSub = "97"
'        bVB_FormPayAct = True
'    Case "46"
'        tVB_CNMnuCurrent = "172105"
'        bUT_SignOut = True
'        bVB_CNCloseTmnSta = True
'        bVB_CNOverShortSta = True
'        tVB_Transaction = aVB_TransMode(46).ID
'        tVB_TransactionSub = tVB_Transaction
'        tVB_TransName = aVB_TransMode(46).Name
'        tVB_TransNameABB = aVB_NameAbb(tEN_AbbClosingReportManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbClosingReportManualTax - 1).UsrLangSec
'        tVB_CNTypePrinter = aVB_TransMode(46).PrintType
'        Call wFunctionKB.W_PrintSignOut(False, False, True)
'        tUT_SaleDate = UT_tStringDate(SP_GETdSaleDate)
'        tVB_EJDate = tUT_SaleDate
'        oTerminal.SaleDate = CDate(tUT_SaleDate)
'        bVB_CNCloseTmnSta = False
'        bVB_CNClosingRptSta = False
'        bVB_CNClosingRptAgain = False
'        bUT_SignOut = False
'        tVB_Transaction = "97"
'        tVB_TransactionSub = "97"
'        bVB_FormPayAct = True
    Case "08"
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Found Close Report Missing Closing Report" & " Form:mDB" & " Control:General" & " Events:Main")
        bVB_CNCloseTmnSta = True
        bVB_CNClosingRptSta = False
        bVB_CNClosingRptAgain = True
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Re Processing Closing Report" & " Form:mDB" & " Control:General" & " Events:Main")
        wCNCloseMnu.ogdFunc.Select 7, 0
        If SP_CHKxIsFormOpen("wCNCloseMnu") Then Unload wCNCloseMnu
        tUT_SaleDate = UT_tStringDate(SP_GETdSaleDate)
        tVB_EJDate = tUT_SaleDate
        oTerminal.SaleDate = CDate(tUT_SaleDate)
        bVB_CNCloseTmnSta = False
        bVB_CNClosingRptSta = False
        bVB_CNClosingRptAgain = False
        bUT_SignOut = False
    End Select
    Call SP_CRTxLogPrg("SP_ChkxClosingMissing End.")
    End If
    '===========================================
    
'    Call SP_CRTxLogPrg("SP_SETxTokenObjects Start.")
'    Call SP_SETxTokenObjects '*Bmp 61-03-06 Token
'    Call SP_CRTxLogPrg("SP_SETxTokenObjects End.")
    
'    '*Tao 61-10-05 Send Echo at Start
'    Call SP_STAxSplashByAct("Send Echo to WPA")
'    Call SP_PRCbSendEcho
'    bVB_CNSplash = False
    
    bVB_FirstLoad = False '*Tao 60-01-20
    Call wBackGrand.W_PRVxFormActivate             '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "wBackGrand.Form_Activate[1]")
'     wBackGrand.Form_Activate
'    Call SP_CRTxWriteLogTerminated("ErrorNumber:0,Form:mDB," & "wBackGrand.Form_Activate[2]")
    Exit Sub
ErrHandle:
    On Error Resume Next
    Call oPrn.KillDoc
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:mDB" & " Control:General" & " Events:Main")
    Call SP_APPxEndProgram
End Sub

Public Sub W_CLRxClearDataHD_DT()
'-----------------------------------------------------------
'   Call:
'   Cmt:  ทำการ Clear ค่า Table ที่เกี่ยวกับ Transaction ใน DB ฝั่ง OnLine/OffLine
'----------------------------------------------------------
    Dim tSql$
    
    On Error Resume Next            '*AO51-03-19
    '*Eaw 56-02-06 For Recovery Process
    'ให้มีการตรวจสอบใน Table Temp ก่อนว่ามีข้อมูลที่ Trigger ยังไม่ส่งสู่ Table จริงหรือไม่ ถ้ามีจะไม่ Drop และ Clone ก่อน
    tVB_CNTblHDR = "TPSHD" & tVB_CNTerminalNum & "R"
    tVB_CNTblDTR = "TPSDT" & tVB_CNTerminalNum & "R"
'    tVB_CNTblRCR = "TPSRC" & tVB_CNTerminalNum & "R"  '*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
    tVB_CNTblCDR = "TPSCD" & tVB_CNTerminalNum & "R"
    tVB_CNTblDTRTmp = "TPSDT" & tVB_CNTerminalNum & "RTmp"
    tVB_CNTblCDRTmp = "TPSCD" & tVB_CNTerminalNum & "RTmp"
    
    'TPSHDXXXR
    If SP_TBLbChkExist(tVB_CNTblHD, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & tVB_CNTblHD & "',@DestinationTable ='" & tVB_CNTblHDR & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblHDR
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "INSERT INTO  " & tVB_CNTblHDR & " SELECT * FROM " & tVB_CNTblHD
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    'TPSDTXXXR
    If SP_TBLbChkExist(tVB_CNTblDT, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & tVB_CNTblDT & "',@DestinationTable ='" & tVB_CNTblDTR & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblDTR
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "INSERT INTO  " & tVB_CNTblDTR & " SELECT * FROM " & tVB_CNTblDT
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If

    '*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
    'TPSRCXXXR
'    If SP_TBLbChkExist(tVB_CNTblRC, oDB.ocnOffLine) Then
'        tSql = "STP_CNxCloneTable @SourceTable ='" & tVB_CNTblRC & "',@DestinationTable ='" & tVB_CNTblRCR & "'"
'        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'        tSql = "TRUNCATE TABLE  " & tVB_CNTblRCR
'        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'        tSql = "INSERT INTO  " & tVB_CNTblRCR & " SELECT * FROM " & tVB_CNTblRC
'        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'    End If
    
    '    TPSCDXXXR
    If SP_TBLbChkExist(tVB_CNTblCD, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & tVB_CNTblCD & "',@DestinationTable ='" & tVB_CNTblCDR & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblCDR
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "INSERT INTO  " & tVB_CNTblCDR & " SELECT * FROM " & tVB_CNTblCD
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    'TPSCDXXXRTmp
    If SP_TBLbChkExist(tVB_CNTblCDTmp, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & tVB_CNTblCDTmp & "',@DestinationTable ='" & tVB_CNTblCDRTmp & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblCDRTmp
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "INSERT INTO  " & tVB_CNTblCDRTmp & " SELECT * FROM " & tVB_CNTblCDTmp
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
'*Eaw 56-07-02 Test DTTmp
        'TPSDTXXXRTmp
    If SP_TBLbChkExist(tVB_CNTblDTTmp, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & tVB_CNTblDTTmp & "',@DestinationTable ='" & tVB_CNTblDTRTmp & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "TRUNCATE TABLE  " & tVB_CNTblDTRTmp
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        tSql = "INSERT INTO  " & tVB_CNTblDTRTmp & " SELECT * FROM " & tVB_CNTblDTTmp
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    '-----------------------------------------------------------
   
    'On Line
'     If SP_TBLbChkExist(tVB_CNTblDT) Then
'        tSql = "DELETE " & tVB_CNTblDT
'        tSql = tSql & "  FROM " & tVB_CNTblHD & " INNER JOIN "
'        tSql = tSql & "  " & tVB_CNTblDT & " ON " & tVB_CNTblHD & ".FTTmnNum = " & tVB_CNTblDT & ".FTTmnNum AND"
'        tSql = tSql & "  " & tVB_CNTblHD & ".FTShdTransNo = " & tVB_CNTblDT & ".FTShdTransNo AND " & tVB_CNTblHD & ".FTShdTransType = " & tVB_CNTblDT & ".FTShdTransType AND"
'        tSql = tSql & "  " & tVB_CNTblHD & ".FDShdTransDate = " & tVB_CNTblDT & ".FDShdTransDate"
'        tSql = tSql & "  WHERE " & tVB_CNTblHD & ".FTTmnNum='" & tVB_CNTerminalNum & "'"
'        tSql = tSql & "  AND " & tVB_CNTblHD & ".FDShdTransDate ='" & tUT_SaleDate & "'"
'        tSql = tSql & "  AND " & tVB_CNTblHD & ".FTShdStaDoc = '2'"
'        Call SP_SQLbExecute(tSql)
'    End If
'     If SP_TBLbChkExist(tVB_CNTblRC) Then
'        tSql = "DELETE " & tVB_CNTblRC
'        tSql = tSql & "  FROM " & tVB_CNTblHD & " INNER JOIN "
'        tSql = tSql & "  " & tVB_CNTblRC & " ON " & tVB_CNTblHD & ".FTTmnNum = " & tVB_CNTblRC & ".FTTmnNum AND"
'        tSql = tSql & "  " & tVB_CNTblHD & ".FTShdTransNo = " & tVB_CNTblRC & ".FTShdTransNo AND " & tVB_CNTblHD & ".FTShdTransType = " & tVB_CNTblRC & ".FTShdTransType AND"
'        tSql = tSql & "  " & tVB_CNTblHD & ".FDShdTransDate = " & tVB_CNTblRC & ".FDShdTransDate"
'        tSql = tSql & "  WHERE " & tVB_CNTblHD & ".FTTmnNum='" & tVB_CNTerminalNum & "'"
'        tSql = tSql & "  AND " & tVB_CNTblHD & ".FDShdTransDate ='" & tUT_SaleDate & "'"
'        tSql = tSql & "  AND " & tVB_CNTblHD & ".FTShdStaDoc = '2'"
'        Call SP_SQLbExecute(tSql)
'    End If
'    If SP_TBLbChkExist(tVB_CNTblCD) Then
'        tSql = "DELETE " & tVB_CNTblCD
'        tSql = tSql & "  FROM " & tVB_CNTblHD & " INNER JOIN "
'        tSql = tSql & "  " & tVB_CNTblCD & " ON " & tVB_CNTblHD & ".FTTmnNum = " & tVB_CNTblCD & ".FTTmnNum AND"
'        tSql = tSql & "  " & tVB_CNTblHD & ".FTShdTransNo = " & tVB_CNTblCD & ".FTShdTransNo AND " & tVB_CNTblHD & ".FTShdTransType = " & tVB_CNTblCD & ".FTShdTransType AND"
'        tSql = tSql & "  " & tVB_CNTblHD & ".FDShdTransDate = " & tVB_CNTblCD & ".FDShdTransDate"
'        tSql = tSql & "  WHERE " & tVB_CNTblHD & ".FTTmnNum='" & tVB_CNTerminalNum & "'"
'        tSql = tSql & "  AND " & tVB_CNTblHD & ".FDShdTransDate ='" & tUT_SaleDate & "'"
'        tSql = tSql & "  AND " & tVB_CNTblHD & ".FTShdStaDoc = '2'"
'        Call SP_SQLbExecute(tSql)
'    End If
'    If SP_TBLbChkExist(tVB_CNTblHD) Then
'        tSql = "DELETE " & tVB_CNTblHD
'        tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
'        tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
'        tSql = tSql & "  AND FTShdStaDoc = '2'"
'        Call SP_SQLbExecute(tSql)
'    End If
    
    'Off Line
    If SP_TBLbChkExist(tVB_CNTblDT, oDB.ocnOffLine) Then
        tSql = "DELETE " & tVB_CNTblDT
        tSql = tSql & "  FROM " & tVB_CNTblHD & " INNER JOIN "
        tSql = tSql & "  " & tVB_CNTblDT & " ON " & tVB_CNTblHD & ".FTTmnNum = " & tVB_CNTblDT & ".FTTmnNum AND"
        tSql = tSql & "  " & tVB_CNTblHD & ".FTShdTransNo = " & tVB_CNTblDT & ".FTShdTransNo AND " & tVB_CNTblHD & ".FTShdTransType = " & tVB_CNTblDT & ".FTShdTransType AND"
        tSql = tSql & "  " & tVB_CNTblHD & ".FDShdTransDate = " & tVB_CNTblDT & ".FDShdTransDate"
        tSql = tSql & "  WHERE " & tVB_CNTblHD & ".FTTmnNum='" & tVB_CNTerminalNum & "'"
        tSql = tSql & "  AND " & tVB_CNTblHD & ".FDShdTransDate ='" & tUT_SaleDate & "'"
        tSql = tSql & "  AND " & tVB_CNTblHD & ".FTShdStaDoc = '2'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
     If SP_TBLbChkExist(tVB_CNTblRC, oDB.ocnOffLine) Then
        tSql = "DELETE " & tVB_CNTblRC
        tSql = tSql & "  FROM " & tVB_CNTblHD & " INNER JOIN "
        tSql = tSql & "  " & tVB_CNTblRC & " ON " & tVB_CNTblHD & ".FTTmnNum = " & tVB_CNTblRC & ".FTTmnNum AND"
        tSql = tSql & "  " & tVB_CNTblHD & ".FTShdTransNo = " & tVB_CNTblRC & ".FTShdTransNo AND " & tVB_CNTblHD & ".FTShdTransType = " & tVB_CNTblRC & ".FTShdTransType AND"
        tSql = tSql & "  " & tVB_CNTblHD & ".FDShdTransDate = " & tVB_CNTblRC & ".FDShdTransDate"
        tSql = tSql & "  WHERE " & tVB_CNTblHD & ".FTTmnNum='" & tVB_CNTerminalNum & "'"
        tSql = tSql & "  AND " & tVB_CNTblHD & ".FDShdTransDate ='" & tUT_SaleDate & "'"
        tSql = tSql & "  AND " & tVB_CNTblHD & ".FTShdStaDoc = '2'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    If SP_TBLbChkExist(tVB_CNTblCD, oDB.ocnOffLine) Then
        tSql = "DELETE " & tVB_CNTblCD
        tSql = tSql & "  FROM " & tVB_CNTblHD & " INNER JOIN "
        tSql = tSql & "  " & tVB_CNTblCD & " ON " & tVB_CNTblHD & ".FTTmnNum = " & tVB_CNTblCD & ".FTTmnNum AND"
        tSql = tSql & "  " & tVB_CNTblHD & ".FTShdTransNo = " & tVB_CNTblCD & ".FTShdTransNo AND " & tVB_CNTblHD & ".FTShdTransType = " & tVB_CNTblCD & ".FTShdTransType AND"
        tSql = tSql & "  " & tVB_CNTblHD & ".FDShdTransDate = " & tVB_CNTblCD & ".FDShdTransDate"
        tSql = tSql & "  WHERE " & tVB_CNTblHD & ".FTTmnNum='" & tVB_CNTerminalNum & "'"
        tSql = tSql & "  AND " & tVB_CNTblHD & ".FDShdTransDate ='" & tUT_SaleDate & "'"
        tSql = tSql & "  AND " & tVB_CNTblHD & ".FTShdStaDoc = '2'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    If SP_TBLbChkExist(tVB_CNTblHD, oDB.ocnOffLine) Then
        tSql = "DELETE " & tVB_CNTblHD
        tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
        tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
        tSql = tSql & "  AND FTShdStaDoc = '2'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    If SP_TBLbChkExist("TSysLogTableUpdate", oDB.ocnOffLine) Then               '*Ao 2013/01/10  TK
        tSql = "DELETE TSysLogTableUpdate"
        tSql = tSql & " WHERE FTSltGroup = '1'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    '*Tao 61-05-30 ควรลบข้อมูล TPSTSalePoint ที่ไม่สมบูรณ์ด้วย (ถ้าจบบิล Remark จะเท่ากับ 0 หรือ 1 เท่านั้น)
    If bVB_FirstLoad Then 'ทำเฉพาะเปิดโปรแกรมครั้งแรกเท่านั้น
        If SP_TBLbChkExist("TPSTSalePoint", oDB.ocnOffLine) Then
            tSql = "DELETE TPSTSalePoint"
            tSql = tSql & vbCrLf & "WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate='" & tUT_SaleDate & "'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTRemark,'')=''"
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
    End If
    '=============================================================================
End Sub

Private Sub W_PRCxGetUsrHistoryReEntry(ByVal ptTransType As String)
'----------------------------------------------------
'Call:
'Cmt:
'*KT 51-11-03
'---------------------------------------------------
    Dim tSql$, tWhr$
    Dim orsTemp As ADODB.Recordset
    
    tSql = "SELECT FTEmpCode,FTEmpFName ,FTEmpLName ,FTLelCode,FTEmpPW "
    tSql = tSql & vbCrLf & " FROM TCNMEmpMtn"
    tSql = tSql & vbCrLf & " WHERE FTEmpCode = '" & tVB_LEmpID & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            tVB_CNUserCode = tVB_LEmpID
            tVB_CNUserName = Trim(SP_FEDtChkString(orsTemp, "FTEmpFName") & " " & SP_FEDtChkString(orsTemp, "FTEmpLName"))
            tUT_UsrLevel = SP_FEDtChkString(orsTemp, "FTLelCode")
            tUT_UsrPwd = SP_FEDtChkString(orsTemp, "FTEmpPW")
        Else
            tVB_CNUserCode = ""
            tVB_CNUserName = ""
            tUT_UsrLevel = ""
            tUT_UsrPwd = ""
        End If
    Else
        tVB_CNUserCode = ""
        tVB_CNUserName = ""
        tUT_UsrLevel = ""
        tUT_UsrPwd = ""
    End If
    Call SP_DATxRsClose(orsTemp)
    Select Case ptTransType
        Case "38"
            tVB_Transaction = aVB_TransMode(37).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(37).Name
            tVB_CNTypePrinter = aVB_TransMode(37).PrintType
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(37).PrintType
            bVB_ReEntryMode = True
            tVB_CNReEntryIDMain = "1720"
            tVB_CNMnuCurrent = "172001"
        Case "43"
            tVB_Transaction = aVB_TransMode(42).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(42).Name
            tVB_CNTypePrinter = aVB_TransMode(42).PrintType
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(42).PrintType
            bVB_ReEntryMode = True
            tVB_CNReEntryIDMain = "1721"
            tVB_CNMnuCurrent = "172101"
            tVB_CNTerminalNum = GetSetting("AdaPosMall", "POSFront", "TmnNum")
            tUT_SaleDate = GetSetting("AdaPosMall", "POSFront", "SaleDate")
            
            tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
            tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
            tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
            tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
            tVB_CNTblDTTmp = "TPSDT" & tVB_CNTerminalNum & "Tmp"
            tVB_CNTblCDTmp = "TPSCD" & tVB_CNTerminalNum & "Tmp"
    End Select
    tWhr = "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tWhr = tWhr & "  AND FTShdStaDoc = '1'"
    tUT_LoginCode = SP_SQLtFunction("MAX", "FTLogCode", "TPSTSalHD", tWhr, oDB.ocnOffLine)
    'PH 2.0 ML-SF-050 Closing Report ไม่เท่ากับ Flag Sale
    If tUT_LoginCode = "" Then
        If nVB_CNLastLogin > 0 Then
                tUT_LoginCode = SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & Format(nVB_CNLastLogin + 1, "00000")
        Else
                tUT_LoginCode = SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & "00001"
                nVB_CNLastLogin = 1
        End If
    End If
    '*KT 52-03-10
    tSql = "SELECT FTCfgOff "
    tSql = tSql & " FROM TConfigTerminal"
    tSql = tSql & " WHERE FTCfgCode='LRunning'"
    Set orsTemp = ocnCfg.Execute(tSql)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        tVB_TransactionMode = SP_FEDtChkString(orsTemp, "FTCfgOff")
    Else
        tVB_TransactionMode = "0"
    End If
    tVB_CNAccTypeReEntryMode = GetSetting("AdaPosMall", "POSFront", "AccType")
    Call SP_DATxRsNothing(orsTemp)
End Sub

Private Sub W_PRCxGetUsrHistory()
'----------------------------------------------------
'Call:
'Cmt:
'*KT 51-11-03
'---------------------------------------------------
    Dim tSql$, tWhr$
    Dim orsTemp As ADODB.Recordset
    
    tSql = "SELECT FTEmpCode,FTEmpFName ,FTEmpLName ,FTLelCode,FTEmpPW "
    tSql = tSql & vbCrLf & " FROM TCNMEmpMtn"
    tSql = tSql & vbCrLf & " WHERE FTEmpCode = '" & tVB_LEmpID & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            tVB_CNUserCode = tVB_LEmpID
            tVB_CNUserName = Trim(SP_FEDtChkString(orsTemp, "FTEmpFName") & " " & SP_FEDtChkString(orsTemp, "FTEmpLName"))
            tUT_UsrLevel = SP_FEDtChkString(orsTemp, "FTLelCode")
            tUT_UsrPwd = SP_FEDtChkString(orsTemp, "FTEmpPW")
        Else
            tVB_CNUserCode = ""
            tVB_CNUserName = ""
            tUT_UsrLevel = ""
            tUT_UsrPwd = ""
        End If
    Else
        tVB_CNUserCode = ""
        tVB_CNUserName = ""
        tUT_UsrLevel = ""
        tUT_UsrPwd = ""
    End If
    Call SP_DATxRsClose(orsTemp)
    
    tVB_Transaction = aVB_TransMode(2).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(2).Name
    tVB_CNTypePrinter = aVB_TransMode(2).PrintType
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSale - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(2).PrintType
    tWhr = "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tWhr = tWhr & "  AND FTShdStaDoc = '1'"
    tUT_LoginCode = SP_SQLtFunction("MAX", "FTLogCode", "TPSTSalHD", tWhr, oDB.ocnOffLine)
    'PH 2.0 ML-SF-050 Closing Report ไม่เท่ากับ Flag Sale
    If tUT_LoginCode = "" Then
        If nVB_CNLastLogin > 0 Then
                tUT_LoginCode = SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & Format(nVB_CNLastLogin + 1, "00000")
        Else
                tUT_LoginCode = SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & "00001"
                nVB_CNLastLogin = 1
        End If
    End If
    '*KT 52-03-10
    tSql = "SELECT FTCfgOff "
    tSql = tSql & " FROM TConfigTerminal"
    tSql = tSql & " WHERE FTCfgCode='LRunning'"
    Set orsTemp = ocnCfg.Execute(tSql)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        tVB_TransactionMode = SP_FEDtChkString(orsTemp, "FTCfgOff")
    Else
        tVB_TransactionMode = "0"
    End If
    tVB_CNAccTypeReEntryMode = GetSetting("AdaPosMall", "POSFront", "AccType")
    Call SP_DATxRsNothing(orsTemp)
End Sub

Public Sub SP_STAxSplashByAct(ByVal ptRead As String, Optional ByVal pnTag As Long = 5, Optional ByVal pnTimer As Long = 20)
'----------------------------------------------------------
'   Call:
'   Cmt:    show from splash
'----------------------------------------------------------
    bVB_CNSplash = True '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    wCNSplash.tW_Caption = ptRead
    wCNSplash.nW_Tag = pnTag
    wCNSplash.nW_Time = pnTimer
    wCNSplash.Show vbModal
End Sub

'*TW 51-07-15 Service
Public Function SVCbOpenService() As Boolean
    Dim tTemp As String
    Dim oFSO As New FileSystemObject
On Error GoTo ErrHandle
    If Not SP_SVCbCheckIsNT() Then
        MsgBox "This program requires Windows NT/2000/XP"
        Exit Function
    End If
    
    tNT_Account = "LocalSystem"
    tNT_Password = ""
    tNT_SERVICE_NAME = "ServiceAutoClear"
    tNT_Service_Display_Name = "ServiceAutoClear"
    tNT_Service_Description = "Call Service Auto Clear"
    tNT_Service_File_Name = "ServiceAutoClear.exe"
    
    
    tTemp = App.Path & "\ServiceAutoClear.exe"
    If SP_FLEbManFile(tTemp, nEN_Fle1Exist) Then
        tTemp = App.Path & "\"
    Else
        tTemp = oFSO.GetParentFolderName(App.Path) & "\ServiceAutoClear.exe"
        If SP_FLEbManFile(tTemp, nEN_Fle1Exist) Then
            tTemp = oFSO.GetParentFolderName(App.Path) & "\"
        Else
            SVCbOpenService = False
            Exit Function
        End If
    End If
    tNT_AppPath = tTemp
    SVCbOpenService = SP_SVCbCheckService(1)
    Exit Function
ErrHandle:
    SVCbOpenService = False
End Function

Private Function W_GETnLineCharsOPOS() As Integer
'---------------------------------------------------------------------------------
'Cmt: '*KT 55-03-16 การอ่านค่า Line_Chars  จาก Register Config OPOS นั้น ถ้ามี OPOS มากกว่า 1 ประเภท อาจจะไม่ได้ใช้ Key "Line_Chars" เลยเปลี่ยนให้เป็น Function เพื่อ Error จะได้ Return มาเป็น 42
'---------------------------------------------------------------------------------
Dim tDefaultPOSPrinter As String

On Error GoTo ErrHandle
'    W_GETnLineCharsOPOS = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\OLEforRetail\ServiceOPOS\POSPrinter\FP410PRT\Receipt\", "Line_Chars")
    '*Ao 2013/01/29
    tDefaultPOSPrinter = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\OLEforRetail\ServiceOPOS\POSPrinter\", "DefaultPOSPrinter")
    If tDefaultPOSPrinter <> "" Then
        W_GETnLineCharsOPOS = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\OLEforRetail\ServiceOPOS\POSPrinter\" & tDefaultPOSPrinter & "\Receipt\", "Line_Chars")
    Else
        W_GETnLineCharsOPOS = 48
    End If
    Exit Function
ErrHandle:
    W_GETnLineCharsOPOS = 48
End Function

Public Sub W_DisableUSB(ByVal pnStatus As Integer) '*TON 58-12-01
        'pnStatus 3 enable
        'pnStatus 4 disable
        On Error Resume Next
        Dim objReg As Object
        Dim tAlowUSB As String
        tAlowUSB = GetSetting("AdaPosMall", "POSFront", "Auto-Disable-USB") '*TON 58-12-01
        Set objReg = CreateObject("Wscript.shell")
        If pnStatus <> 3 Then pnStatus = 4
'=========================
'        Dim tRunFile As String
'        Dim tdev As String
'        Dim nRet As Long
'        If pnStatus = 3 Then
'            tRunFile = "devcon.exe enable  USBSTOR*"
'            nRet = Shell(tRunFile, vbMinimizedFocus)
'        Else
'            tRunFile = "devcon.exe disable USBSTOR*"
'            nRet = Shell(tRunFile, vbMinimizedFocus)
'        End If
'==========================

        objReg.RegWrite "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\AdaPosMall\POSFront\Auto-Disable-USB", pnStatus - 3, "REG_SZ"
'        If tAlowUSB = "0" Then
'            objReg.Regwrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\USBSTOR\Start", 3, "REG_DWORD"
'            Call SP_CRTxLogPrg(Format(Now, "yyyy-MM-dd hh:mm:ss") & "  USB port is Enabled")
'        Else
'            objReg.Regwrite "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\USBSTOR\Start", pnStatus, "REG_DWORD"
            If pnStatus = 4 Then
                    Call SP_CRTxLogPrg(Format(Now, "yyyy-MM-dd hh:mm:ss") & "  USB port is Disabled")
            Else
                    Call SP_CRTxLogPrg(Format(Now, "yyyy-MM-dd hh:mm:ss") & "  USB port is Enabled")
            End If
'        End If
End Sub

Private Sub DB_AlterTable()
Dim tSql As String
On Error Resume Next
    '*Tao 60-01-20 เพิ่มเมนู EJ Viewer ให้เท่ากัน ทั้ง Online และ Offline
    tSql = "IF NOT EXISTS ( SELECT FTSmnID FROM TSysMenu WHERE FTSmnID  = '1607')"
    tSql = tSql & vbCrLf & "BEGIN"
    tSql = tSql & vbCrLf & "INSERT INTO [dbo].[TSysMenu] ([FDDateUpd], [FTTimeUpd], [FTWhoUpd], [FDDateIns], [FTTimeIns], [FTWhoIns], [FTRemark], [FTSmnRoot], [FTSmnGrp], [FNSmnSeq], [FTSmnID], [FTSmnOrder], [FTSmnName], [FTDefLangPri], [FTDefLangSec], [FTUsrLangPri], [FTUsrLangSec], [FTSmnPthImg], [FTSmnActBtnGrp1], [FTSmnActBtnGrp2], [FTSmnActBtnGrp3], [FTSmnActBtnGrp4], [FTSmnActBtnGrp5], [FTSmnUsage], [FTSmnApp], [FNSmnVisible], [FNSmnAutho])"
    tSql = tSql & vbCrLf & "VALUES (CONVERT (varchar(10), GETDATE(),120),CONVERT (varchar(8), GETDATE(),108),'Manager', CONVERT (varchar(10), GETDATE(),120),CONVERT (varchar(8), GETDATE(),108),'Manager', N'', N'1600', N'Main Menu', 6, N'1607', N'G08', N'omnFntTrnEJViewer', N'EJ Viewer', N'EJ Viewer', N'EJ Viewer', N'EJ Viewer', N'\MenuPNG\J05.PNG', N'000000000000', N'0000', N'0000', N'00', N'000000', N'1', N'AdaStoreBack', 1, 1)"
    tSql = tSql & vbCrLf & "End"
    Call SP_SQLvExecute(tSql, , oDB.ocnOnLine, False)
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine, False)
    tSql = "IF NOT EXISTS ( SELECT FTSpmID FROM TSysPosMenu WHERE FTSpmID  = '1607')"
    tSql = tSql & vbCrLf & "BEGIN"
    tSql = tSql & vbCrLf & "INSERT INTO [dbo].[TSysPosMenu] ([FDDateUpd], [FTTimeUpd], [FTWhoUpd], [FDDateIns], [FTTimeIns], [FTWhoIns], [FTRemark], [FTSpmGrp], [FNSpmSeq], [FTSpmID], [FTSpmOrder], [FTSpmName], [FTSpmNameUser], [FTDefLangPri], [FTDefLangSec], [FTUsrLangPri], [FTUsrLangSec], [FTSpmRef])"
    tSql = tSql & vbCrLf & "VALUES (CONVERT (varchar(10), GETDATE(),120),CONVERT (varchar(8), GETDATE(),108),'Manager', CONVERT (varchar(10), GETDATE(),120),CONVERT (varchar(8), GETDATE(),108),'Manager', N'', N'Main Menu', 19, N'1607', N'G19', N'EJ Viewer', N'EJ Viewer', N'EJ Viewer', N'', N'', N'', N'W_PRCxTrnMnuEJViewer')"
    tSql = tSql & vbCrLf & "End"
    Call SP_SQLvExecute(tSql, , oDB.ocnOnLine, False)
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine, False)
End Sub
