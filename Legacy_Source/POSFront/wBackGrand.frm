VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#11.1#0"; "Codejock.SkinFramework.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form wBackGrand 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Mall"
   ClientHeight    =   5145
   ClientLeft      =   2415
   ClientTop       =   2460
   ClientWidth     =   9660
   Icon            =   "wBackGrand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin SCANNERLib.Scanner oScanner 
      Left            =   360
      Top             =   960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
   End
   Begin MSComctlLib.StatusBar ostForm 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4770
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1323
            MinWidth        =   1323
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
            Text            =   "Ready..."
            TextSave        =   "Ready..."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4764
            MinWidth        =   4764
            Text            =   "TheMall/AdaDB"
            TextSave        =   "TheMall/AdaDB"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "User : 009/Manager"
            TextSave        =   "User : 009/Manager"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "20/04/2018"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1376
            MinWidth        =   1376
            TextSave        =   "15:47"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1147
            MinWidth        =   1147
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin XtremeSkinFramework.SkinFramework osfForm 
      Left            =   0
      Top             =   0
      _Version        =   720897
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "wBackGrand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Sub W_PRVxFormActivate()
On Error GoTo ErrHandle
     '*KT 52-12-17 POSFront 2.0.30
     Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRVxFormActivate[1]")
     Call SP_CRTxLogMallCardExtend("Date:" & Date & " Time:" & Time$ & " tVB_Transaction :" & tVB_Transaction)   '*Bump 57-11-04 TK_Issue Monitor tVB_Transaction เปลี่ยน ทำให้ sale เปลี่ยนเป็น Total Report
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) And oTerminal.SaleMode = nEN_OnlineNormal Then
         ostForm.Panels(1).Text = "SRV"
     Else
         ostForm.Panels(1).Text = "LOCAL"
    End If
    If bVB_CNPrgFsh Then
               
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnOpen:
                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRVxFormActivate[2]")
                Call W_PRCxOpen
'*KT 53-02-03 ป้องกันการ Terminated
'                ostForm.Panels(4).Text = "User : " & tVB_CNUserCode & "/" & tVB_CNUserName '*KT 52-12-17 POSFront 3.0.10
'                Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Open Transaction" & " Form:" & Me.Name & " Control:Form" & " Events:Form_Activate")
            Case tEN_TrnStartBank:
                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRVxFormActivate[3]")
                Call W_PRCxStartBank
            Case tEN_TrnSale, tEN_TrnSaleTaxInc, tEN_TrnSaleManualTax:
                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRVxFormActivate[4]")
'                ostForm.Panels(4).Text = "User : " & tVB_CNUserCode & "/" & tVB_CNUserName '*KT 52-12-17 POSFront 3.0.10
                ostForm.Panels(4).Text = "User : " & tVB_CNUserName '*Eaw 56-07-15 CommSheet TK-ISS3000-278 TR FAPL-058.xls
                Call W_PRCxSale
            Case tEN_TrnReturn:
            Case tEN_TrnIPV:
            Case tEN_TrnDeposit:
            Case tEN_TrnRedeem:
            Case tEN_TrnClose:
'*KT 52-12-17 POSFront 2.0.30
'                Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Closing Report" & " Form:" & Me.Name & " Control:Form" & " Events:Form_Activate")
'                Call W_PRCxClose
            Case tEN_TrnPrintClose:
            Case tEN_TrnVIPSale:
            Case tEN_TrnTaxExceptSale:
            Case tEN_TrnPickUp:
            Case tEN_TrnDutyFreeSale:
            Case tEN_TrnIPVReturn:
            Case tEN_TrnVoucher:
            Case tEN_TrnDepositCancel:
            Case tEN_TrnIPVRedeem:
            Case tEN_TrnNoSale:
            Case tEN_TrnAddDeposit:
            Case tEN_TrnAddDepositCancel:
            Case tEN_TrnCancelVoucher:
            Case tEN_TrnOtherReceive:
            Case tEN_TrnOtherDisbursement:
            Case tEN_TrnDepositMoney:
            Case tEN_TrnTraining:
            Case tEN_TrnIPVVIPService:
            Case tEN_TrnIPVTaxExceptSale:
            Case tEN_TrnIPVDutyFreeSale:
            Case 97:
                Call SP_CRTxWriteLogTerminated("No_564 Check variable bUT_SignOut Form:" & Me.Name & " Events:W_PRVxFormActivate Case 97_bUT_SignOut:" & bUT_SignOut) '*Bump 57-11-04 เพิ่ม Log CommSheet TK-ISS3000-564.xlsx พิมพ์ Over Short แต่ไม่พิมพ์ Closing ของ Re-entrt Tanant Sales
                If bUT_SignOut Then
                    'wCNCloseMnu.ogdFunc.Select 2, 0 '*Eaw 55-12-28  Comment
                    Call SP_CRTxWriteLogTerminated("No_564 Check Closing focus Menu Form:" & Me.Name & " Events:W_PRVxFormActivate Case 97 before Select (7,0)") '*Bump 57-11-04 เพิ่ม Log CommSheet TK-ISS3000-564.xlsx พิมพ์ Over Short แต่ไม่พิมพ์ Closing ของ Re-entrt Tanant Sales
                    bVB_CNClosingRptSta = False
                    '*Tao 60-01-17 ใช้ Load ก่อน( กัน error แล้วเด้งออก)
                    If Not SP_CHKxIsFormOpen("wCNCloseMnu") Then
                        Call SP_GETxMainMnu
                        Load wCNCloseMnu
                    Else
                        If wCNCloseMnu.ogdFunc.Aggregate(flexSTMax, 0, 0, wCNCloseMnu.ogdFunc.rows - 1, 0) < 8 Then
                            Unload wCNCloseMnu
                            Call SP_GETxMainMnu
                            Load wCNCloseMnu
                        End If
                    End If
                    '======================================
                    wCNCloseMnu.ogdFunc.Select 7, 0  '*Eaw 55-12-28 -เป็นการ Auto Print Closing ใน MainMenu (wCNCloseMnu.frm)'*TON 58-12-01 TK-ISS3000-581
                    Call SP_CRTxWriteLogTerminated("No_564 Check Closing focus Menu Form:" & Me.Name & " Events:W_PRVxFormActivate Case 97 after Select (7,0)") '*Bump 57-11-04 เพิ่ม Log CommSheet TK-ISS3000-564.xlsx พิมพ์ Over Short แต่ไม่พิมพ์ Closing ของ Re-entrt Tanant Sales
                    wCNCloseMnu.ocmOk_Click
                    '*KT 58-05-27 TK-ISS3000-587 ป้องกันการไม่พิมพ์ใบ Closing Report
                    '********************************
                    If Not bVB_CNClosingRptSta Then
                         Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRVxFormActivate[Closing Report not success]")
                         If Not bVB_CNTrnComplete Then
                            wCNCloseMnu.ogdFunc.Select 7, 0 '*TON 58-12-01 TK-ISS3000-581
                            wCNCloseMnu.ocmOk_Click
                         End If
                         If Not bVB_CNClosingRptSta Then
                            bVB_CNClosingRptAgain = True
                         End If
                    End If
                     '********************************
                    If bVB_ReEntryMode Then
                        Call SP_GETbTmnData
                        tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                        tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                        tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                        tVB_CNTblDTTmp = "TPSDT" & tVB_CNTerminalNum & "Tmp"
                        tVB_CNTblCDTmp = "TPSCD" & tVB_CNTerminalNum & "Tmp"
                        tUT_SaleDate = UT_tStringDate(SP_GETdSaleDate)
                    End If
                    If Not bVB_ReEntryMode Then
                        wCNCloseMnu.ogdFunc.Select 0, 0
                    End If
                    
                 End If
                 If Not bVB_ReEntryMode Then
                    wCNCloseMnu.Show vbModal
                Else
                    wCNReEntryMnu.Show vbModal '*Tao 60-01-20
                End If
            Case 98:
'                ostForm.Panels(4).Text = "User : " & tVB_CNUserCode & "/" & tVB_CNUserName '*KT 52-12-17 POSFront 3.0.10
                ostForm.Panels(4).Text = "User : " & tVB_CNUserName '*Eaw 56-07-15 CommSheet TK-ISS3000-278 TR FAPL-058.xls
                Call W_PRCxReEntry
            Case 99: '*KT 52-12-17 POSFront 2.0.30
                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRVxFormActivate[5]")
'                If Screen.ActiveForm.Name <> "wCNCloseMnu" Then
'                    wCNCloseMnu.Show vbModal
'                End If
                
'                If bUT_SignOut Then
'                    wCNCloseMnu.ogdFunc.Select 2, 0
'                    wCNCloseMnu.ocmOK_Click
'                    wCNCloseMnu.ogdFunc.Select 0, 0
'                 End If
                 wCNCloseMnu.Show vbModal
            Case 100: '*KT 52-12-17 POSFront 2.0.30
                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRVxFormActivate[6]")
                Call W_PRCxClose
            Case Else
                'none
                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRVxFormActivate[Case Else]")
        End Select
    Else
        Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRVxFormActivate[bVB_CNPrgFsh = false]")
    End If
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:W_PRVxFormActivate[7]")
    Resume Next '*Tao 60-01-19 ให้กลับไปทำที่เหลือต่อ
'    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Private Sub Form_Activate()
'    MsgBox "TEST Form_Activate()"
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
End Sub

Private Sub Form_Load()
   'Set ค่าเบื้องต้น
   Dim tWhe$
   Dim tPW$
   Dim tUsr$
On Error GoTo ErrHandle
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[1]")
    tUsr = GetSetting("AdaPosMall", "POSFront", "User")
    tPW = GetSetting("AdaPosMall", "POSFront", "PW")
    Me.Caption = tVB_CNAppName & " " & "V." & App.Major & "." & App.Minor & "." & App.Revision
    '*Ao 52-02-23
    If bVb_SetResorution Then
        nVB_ResolutionX = GetSystemMetrics(0)
        nVB_ResolutionY = GetSystemMetrics(1)
        If (nVB_ResolutionX = 640 And nVB_ResolutionY = 480) Then
            bVB_Res640X480 = True
            nVB_ResolutionX = 640
            nVB_ResolutionY = 480
        Else
            bVB_Res640X480 = False
            nVB_ResolutionX = 800
            nVB_ResolutionY = 600
        End If
        Call SP_PRCxScreen(nVB_ResolutionX, nVB_ResolutionY, Me)
    End If
    '__________________________________________________________
    Call SP_FRMxLoadThemeStyle(Me, osfForm)
    If bVB_Res640X480 Then
            ostForm.Panels(1).Width = 750.04
            ostForm.Panels(2).Width = 1800
            ostForm.Panels(3).Width = 2000.12
            ostForm.Panels(4).Width = 2000.12
            ostForm.Panels(5).Width = 1000.06
            ostForm.Panels(6).Width = 780.09
            ostForm.Panels(7).Width = 650.26
            ostForm.Panels(8).Width = 720
    End If
    
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) And oTerminal.SaleMode = nEN_OnlineNormal Then
         ostForm.Panels(1).Text = "SRV"
     Else
         ostForm.Panels(1).Text = "LOCAL"
    End If
    ostForm.Panels(3).Text = oDB.OffServerName & "." & oDB.OffDbName  'DB Offline

    Me.Show
'    ostForm.Panels(4).Text = "User : " & tUsr & "/" & tPW   '*Ao 51-04-05
    ostForm.Panels(4).Text = "User : " & tPW '*Eaw 56-07-15 CommSheet TK-ISS3000-278 TR FAPL-058.xls
    If Not bVB_CNPrgFsh Then
       Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[2]")
       '*KT 54-04-08 V.3.0.56 ป้องกันเครื่องค้างและโปรแกรมหลุด
        'Call SP_STAxSplash
        'bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
        Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[3]")
    End If
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[4]")
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:Form_Load[5]")
End Sub

Private Sub W_PRCxOpen()
'PH 2.0.0
'---------------------------------------------------------------------------------------------
'Call:
'Cmt: เป็นการทำงานครั้งแรกที่มีการเปิดโปรแกรมใหม่ทุกครั้งและพบว่าเป็นการเปิดโปรแกรมที่พร้อมจะทำการ Open Terminal
'--------------------------------------------------------------------------------------------

'*KT 53-02-03 ทำการย้ายจาก wBackGrand/W_PRCxOpen ไปที่ wCNLogIn/W_PRCxOpen เนื่องจากเหตุผล 2 ประการ
'1. เครื่อง T.2000 ถ้ามีการทำงานใน Form_Activate ใช้เวลานาน จะทำให้ wCNLogIn จะซ่อนอยู่ด้านหลัง wBackGrand
'2. กรณีที่กดเลือกเมนู Open จากหน้าจอ Main Menu ระบบจะทำการลบข้อมูลใน Local ที่เป็นตาราง Temp ทั้งที่ ทั้งที่ยังไม่มีการใส่ Code Cashier
'    Dim orsTemp As ADODB.Recordset
'    Dim tSql$
'    Dim tErr$
    
'    Call SP_STAxMouse(Me, True)
    '*KT 51-10-08 V 3.0.131/V 3.0.200
'    '-----------------------------------------------------
'    If Not bVB_CNCreateTblTmp Or Not bVB_CNCreateTblTmpFrmCloseRpt Then
'        Call UT_CreateSaleTableTMP
'        bVB_CNCreateTblTmp = True
'    End If
'    '-----------------------------------------------------
'
''PH 2.0.0 ตรวจสอบข้อมูลที่ Local ตลอด
'    tSql = "select  name from Sys.sysobjects where name = '" & "TRG_Tmp2Sale" & tVB_CNTerminalNum & "'"
'    Call SP_SQLvExecute(tSql, orsTemp)
'    If orsTemp.EOF And orsTemp.EOF Then
'        '*KT 52-01-14 กรณีตรวจสอบเจอว่าไม่มี Trigger แล้วยังสร้างไม่สำเร็จทั้งฐานข้อมูล server และ local
'        If Not SP_DATbCrtTrigger(False, True) Then
'              Call SP_MSGnShowing("tms_CN005168", nCS_Warning)
'              Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Tricger" & " Form:" & Me.Name & " Step: W_PRCxOpen" & " Events:SP_DATbCrtTrigger")
'        End If
'    End If
    
'    Call SP_DATxRsNothing(orsTemp)
    '------------------------------------------------------
On Error GoTo ErrHandle
    tVB_Transaction = aVB_TransMode(0).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(0).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbOpen - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbOpen - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(0).PrintType
    '*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตังต้น 3.0.19
    '*KT 53-05-18 ป้องกันการขึ้น Dot send
     If Not bVB_CNCreateTblTmp Then
        tVB_Running = SP_AUTtGenCode("TPSTSalHD") '*KT 53-02-2010  PH3-ML-SF-088 POSFront V3.0.18 Open error(-2147217865) เนื่องจากมีการย้าย Code การสร้าง Table Temp ไปไว้ที่หน้าจอ LogIn แทน ในวันที่ *KT 53-02-03
    Else
         tVB_Running = SP_AUTtGenCode(tVB_CNTblHD)
    End If
'    Call SP_STAxMouse(Me, False)
    wCNLogIn.Show vbModal
    Exit Sub
ErrHandle:
    '*KT 54-03-26
    'wCNCloseMnu.Show vbModal
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:W_PRCxOpen") '*KT 54-04-12 V.3.0.57 เก็บ Log
End Sub

Private Sub W_PRCxStartBank()
On Error GoTo ErrHandle
    tVB_Transaction = aVB_TransMode(1).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(1).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbStartBank - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbStartBank - 1).UsrLangSec
    tVB_Running = SP_AUTtGenCode(tVB_CNTblHD)
    tVB_CNTypePrinter = aVB_TransMode(1).PrintType
    wTenderMain.Show vbModal
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:W_PRCxStartBank") '*KT 54-04-12 V.3.0.57 เก็บ Log
End Sub

Private Sub W_PRCxSale()
'PH 2.0.0
    Dim orsTemp As ADODB.Recordset
    Dim tSql$
    Dim tErr$
On Error GoTo ErrHanle

    'Off Line
    If Not SP_TBLbChkExist(tVB_CNTblHD, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalHD" & "',@DestinationTable ='" & tVB_CNTblHD & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If

    'for TPSTSalDTXXX
    If Not SP_TBLbChkExist(tVB_CNTblDT, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalDT" & "',@DestinationTable ='" & tVB_CNTblDT & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    'for TPSTSalDTXXXTmp
    If Not SP_TBLbChkExist(tVB_CNTblDTTmp, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalDT" & "',@DestinationTable ='" & tVB_CNTblDTTmp & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If

    'for TPSTSalRCXXX
    If Not SP_TBLbChkExist(tVB_CNTblRC, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalRC" & "',@DestinationTable ='" & tVB_CNTblRC & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    'for TPSTSalCDXXX
    If Not SP_TBLbChkExist(tVB_CNTblCD, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalCD" & "',@DestinationTable ='" & tVB_CNTblCD & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    'for TPSTSalCDXXXTmp
    If Not SP_TBLbChkExist(tVB_CNTblCDTmp, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalCD" & "',@DestinationTable ='" & tVB_CNTblCDTmp & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    bVB_CNCreateTblTmp = True '*KT 51-10-08 V 3.0.131/V 3.0.200
    bVB_CNCreateTblTmpFrmCloseRpt = False '*KT 51-10-13 V 3.0.132/V 3.0.200
    
'PH 2.0.0 ตรวจสอบข้อมูลที่ Local ตลอด
    tSql = "select  name  from sysobjects where name = '" & "TRG_Tmp2Sale" & tVB_CNTerminalNum & "'"
    Call SP_SQLvExecute(tSql, orsTemp)
    If orsTemp.EOF And orsTemp.EOF Then
         '*KT 52-01-14 กรณีตรวจสอบเจอว่าไม่มี Trigger แล้วยังสร้างไม่สำเร็จทั้งฐานข้อมูล server และ local
        If Not SP_DATbCrtTrigger(False, True) Then
             Call SP_MSGnShowing("tms_CN005168", nCS_Warning)
             Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Tricger" & " Form:" & Me.Name & " Step: W_PRCxSale" & " Events:SP_DATbCrtTrigger")
        End If
    End If

    Call SP_DATxRsNothing(orsTemp)
    '------------------------------------------------------
        '*Eaw 56-08-27
        Select Case tVB_CNMnuCurrent
            Case "172001"  'Re-Entry Tax Include
                tVB_Transaction = aVB_TransMode(47).ID
                tVB_TransactionSub = tVB_Transaction
                tVB_TransName = aVB_TransMode(47).Name
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbOpenTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbOpenTaxInc - 1).UsrLangSec
                tVB_CNTypePrinter = aVB_TransMode(47).PrintType
                Call SP_INSxTransHD(tVB_CNTblHD)
                Call wFunctionKB.W_SetToComplete
                Call SP_UPDxSetCompleteTerminalMtn
            Case "172101" 'Re-Entry Tax Manual(Tenant Sales)
                tVB_Transaction = aVB_TransMode(48).ID
                tVB_TransactionSub = tVB_Transaction
                tVB_TransName = aVB_TransMode(48).Name
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbOpenManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbOpenManualTax - 1).UsrLangSec
                tVB_CNTypePrinter = aVB_TransMode(48).PrintType
                Call SP_INSxTransHD(tVB_CNTblHD)
                Call wFunctionKB.W_SetToComplete
                Call SP_UPDxSetCompleteTerminalMtn
                Call SP_TmnxUpdateOpenVirtualPos '*Eaw 56-09-30  Update status ให้ Virtual Pos ที่ใช้งานอยู่
                
'                If Not SP_TMNxUpdateOpen(oTerminal.SaleMode) Then '*BG 56-10-31 Mark 1 Line ไม่ให้อัพเดทค่า TmnStatus 2 Row
'                    Call SP_STAxMouse(Me, False)
'                    Call SP_MSGnShowing("tms_CN005211", tCS_CNMsgErr)
'                    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:SP_TMNxUpdateOpen =false" & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCxOpen")
'                    Exit Sub
'                End If
        End Select
        Select Case tVB_CNMnuCurrent
            Case "172001"
                tVB_Transaction = aVB_TransMode(37).ID
                tVB_TransactionSub = tVB_Transaction
                tVB_TransName = aVB_TransMode(37).Name
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangSec
                tVB_CNTypePrinter = aVB_TransMode(37).PrintType
            Case "172101"
                tVB_Transaction = aVB_TransMode(42).ID
                tVB_TransactionSub = tVB_Transaction
                tVB_TransName = aVB_TransMode(42).Name
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangSec
                tVB_CNTypePrinter = aVB_TransMode(42).PrintType
            Case Else
                '---------------------------------------------------------------------------------*BG 56-11-01
                If bVB_ReEntryMode Then
                    Select Case tVB_CNReEntryIDMain
                        Case "1720"
                            tVB_Transaction = aVB_TransMode(37).ID
                            tVB_TransactionSub = tVB_Transaction
                            tVB_TransName = aVB_TransMode(37).Name
                            tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangSec
                            tVB_CNTypePrinter = aVB_TransMode(37).PrintType
                        Case "1721"
                            tVB_Transaction = aVB_TransMode(42).ID
                            tVB_TransactionSub = tVB_Transaction
                            tVB_TransName = aVB_TransMode(42).Name
                            tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangSec
                            tVB_CNTypePrinter = aVB_TransMode(42).PrintType
                    End Select
                Else
                    tVB_Transaction = aVB_TransMode(2).ID
                    tVB_TransactionSub = tVB_Transaction
                    tVB_TransName = aVB_TransMode(2).Name
                    tVB_TransNameABB = aVB_NameAbb(tEN_AbbSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSale - 1).UsrLangSec
                    tVB_CNTypePrinter = aVB_TransMode(2).PrintType
                End If
'                tVB_Transaction = aVB_TransMode(2).ID
'                tVB_TransactionSub = tVB_Transaction
'                tVB_TransName = aVB_TransMode(2).Name
'                tVB_TransNameABB = aVB_NameAbb(tEN_AbbSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSale - 1).UsrLangSec
'                tVB_CNTypePrinter = aVB_TransMode(2).PrintType
                '--------------------------------------------------------------------------------------
    End Select
    tVB_Running = SP_AUTtGenCode(tVB_CNTblHD)
    
    Call SP_INSxTransHD(tVB_CNTblHD)
    Call SP_TMNxUpdateOpenOnline   'PH 2.0.0
    wMain.Show vbModal  '*KT 52-02-23  640 หรือ 800
'    If wMain.Visible = False Then
'        wMain.Show vbModal  '*KT 52-02-23  640 หรือ 800
'    End If
'    Debug.Print "wMain.Visible = True"
'    wMain.Show vbModal  '*KT 52-02-23  640 หรือ 800
    Exit Sub
ErrHanle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxSale[ErrHanle]")
    Exit Sub
End Sub

Private Sub W_PRCxReEntry()
'PH 2.0.0
    Dim orsTemp As ADODB.Recordset
    Dim tSql$
    Dim tErr$
On Error GoTo ErrHanle

    'Off Line
    If Not SP_TBLbChkExist(tVB_CNTblHD, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalHD" & "',@DestinationTable ='" & tVB_CNTblHD & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If

    'for TPSTSalDTXXX
    If Not SP_TBLbChkExist(tVB_CNTblDT, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalDT" & "',@DestinationTable ='" & tVB_CNTblDT & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    'for TPSTSalDTXXXTmp
    If Not SP_TBLbChkExist(tVB_CNTblDTTmp, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalDT" & "',@DestinationTable ='" & tVB_CNTblDTTmp & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If

    'for TPSTSalRCXXX
    If Not SP_TBLbChkExist(tVB_CNTblRC, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalRC" & "',@DestinationTable ='" & tVB_CNTblRC & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    'for TPSTSalCDXXX
    If Not SP_TBLbChkExist(tVB_CNTblCD, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalCD" & "',@DestinationTable ='" & tVB_CNTblCD & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    'for TPSTSalCDXXXTmp
    If Not SP_TBLbChkExist(tVB_CNTblCDTmp, oDB.ocnOffLine) Then
        tSql = "STP_CNxCloneTable @SourceTable ='" & "TPSTSalCD" & "',@DestinationTable ='" & tVB_CNTblCDTmp & "'"
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    
    bVB_CNCreateTblTmp = True '*KT 51-10-08 V 3.0.131/V 3.0.200
    bVB_CNCreateTblTmpFrmCloseRpt = False '*KT 51-10-13 V 3.0.132/V 3.0.200
    
'PH 2.0.0 ตรวจสอบข้อมูลที่ Local ตลอด
    tSql = "select  name  from sysobjects where name = '" & "TRG_Tmp2Sale" & tVB_CNTerminalNum & "'"
    Call SP_SQLvExecute(tSql, orsTemp)
    If orsTemp.EOF And orsTemp.EOF Then
         '*KT 52-01-14 กรณีตรวจสอบเจอว่าไม่มี Trigger แล้วยังสร้างไม่สำเร็จทั้งฐานข้อมูล server และ local
        If Not SP_DATbCrtTrigger(False, True) Then
             Call SP_MSGnShowing("tms_CN005168", nCS_Warning)
             Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Tricger" & " Form:" & Me.Name & " Step: W_PRCxSale" & " Events:SP_DATbCrtTrigger")
        End If
    End If

    Call SP_DATxRsNothing(orsTemp)
    '------------------------------------------------------
    
'    tVB_Transaction = aVB_TransMode(2).ID
'    tVB_TransactionSub = tVB_Transaction
'    tVB_TransName = aVB_TransMode(2).Name
'    tVB_TransNameABB = aVB_NameAbb(tEN_AbbSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSale - 1).UsrLangSec
'    tVB_CNTypePrinter = aVB_TransMode(2).PrintType
'    tVB_Running = SP_AUTtGenCode(tVB_CNTblHD)
'    Call SP_INSxTransHD(tVB_CNTblHD)
'    Call SP_TMNxUpdateOpenOnline   'PH 2.0.0

'    wMain.Show vbModal  '*KT 52-02-23  640 หรือ 800
'    tVB_CNMnuCurrent = "1606"
'    If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
'            wCNCloseMnu.Show vbModal
'            Exit Sub
'    End If
    wCNReEntryMnu.otmForm.Enabled = True            '*Ao 2014/02/03
    wCNReEntryMnu.Show vbModal
    If Not bVB_ReEntryMode Then
        wCNCloseMnu.Show vbModal
    End If
    Exit Sub
ErrHanle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxSale[ErrHanle]")
    Exit Sub
End Sub

Private Sub W_PRCxClose()
On Error GoTo ErrHandle
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxClose[1]")
    wCNLogIn.Show vbModal
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:W_PRCxClose[2]") '*KT 54-04-12 V.3.0.57 เก็บ Log
End Sub

'*KT 53-02-03 ป้องกันการ Terminated ไม่มีความจำเป็นต้องใช้งาน
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = 1
'End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    If Not oTKToken Is Nothing Then Set oTKToken = Nothing '*Bmp 61-02-15
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:Form_Unload") '*KT 54-04-12 V.3.0.57 เก็บ Log
End Sub
