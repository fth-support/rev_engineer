VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form wCNLogIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login;klg_Title"
   ClientHeight    =   3735
   ClientLeft      =   5070
   ClientTop       =   4830
   ClientWidth     =   6405
   ControlBox      =   0   'False
   Icon            =   "wCNLogIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6405
   Tag             =   "2;"
   Begin POSFront.uAdaOnOff uAdaOnOff2 
      Height          =   255
      Left            =   25
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOnLine       =   "wCNLogIn.frx":6852
      PicOffLine      =   "wCNLogIn.frx":6F4C
   End
   Begin MSComctlLib.StatusBar ostForm 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   3315
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
            MinWidth        =   529
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "TheMall_Srv\PosData"
            TextSave        =   "TheMall_Srv\PosData"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "TheMall_Pc\PosData"
            TextSave        =   "TheMall_Pc\PosData"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl otcForm3 
      Height          =   1335
      Left            =   15
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2265
      Width           =   1815
      _Version        =   720897
      _ExtentX        =   3201
      _ExtentY        =   2355
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin POSFront.uAdaOnOff uAdaOnOff1 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm2 
      Height          =   2535
      Left            =   15
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   15
      Width           =   1815
      _Version        =   720897
      _ExtentX        =   3201
      _ExtentY        =   4471
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Image oimForm 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   90
         Picture         =   "wCNLogIn.frx":7646
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1605
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm4 
      Height          =   1335
      Left            =   1815
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2265
      Width           =   4575
      _Version        =   720897
      _ExtentX        =   8070
      _ExtentY        =   2355
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Timer otmForm 
         Interval        =   5000
         Left            =   120
         Top             =   360
      End
      Begin VB.CommandButton ocmCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel;klg_Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton ocmOk 
         Caption         =   "&Ok;klg_Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2535
      Left            =   1815
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   15
      Width           =   4575
      _Version        =   720897
      _ExtentX        =   8070
      _ExtentY        =   4471
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.TextBox otbUser 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   10
         TabIndex        =   0
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox otbPW 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cashier ID;klg_Cashier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   120
         Width           =   2835
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password;klg_Pw"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   1080
         Width           =   2145
      End
   End
   Begin VB.Image oimForm2 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   11040
      Picture         =   "wCNLogIn.frx":CC29
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   1455
   End
End
Attribute VB_Name = "wCNLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private orsForm As ADODB.Recordset
Private tW_Version As String
Private bW_Unload As Boolean

Private Function W_PRCbVerifyTmnNum() As Boolean
'--------------------------------------------------------
'Call:
'Cmt: ทำการตรวจสอบว่า Tetminal มีการใช้งานกับเครื่องอื่นหรือไม่ '*Eaw 56-09-05
'-------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    W_PRCbVerifyTmnNum = False
    tSql = "SELECT FTTmnCode "
    tSql = tSql & " FROM TCNMTerminalMtn"
    tSql = tSql & " WHERE  FTTmnName = '" & tVB_CNComName & "'"
    tSql = tSql & " AND FTTmnStatus = 'A'"
    If oDB.SQLbAvailableServer(oDB.OnServerName, oDB.OnComIP) Then
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    Else
        W_PRCbVerifyTmnNum = False
    End If
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        W_PRCbVerifyTmnNum = True
    Else
        W_PRCbVerifyTmnNum = False
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_PRCbVerifyTmnNum = False
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then SendKeys "{Tab}"
     If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
End Sub
    
Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[1]")
    tW_Version = "V." & App.Major & "." & App.Minor & "." & App.Revision
    uAdaOnOff1.ShowPic = False
    uAdaOnOff1.OnLineCaption = tVB_CNOnCaption
    uAdaOnOff1.OffLineCaption = tVB_CNOffCaption
    uAdaOnOff1.Online = IIf(oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP), True, False)  'oDB.IsCurOnline
    uAdaOnOff2.ShowPic = True
    uAdaOnOff2.OnLineCaption = tVB_CNOnCaption
    uAdaOnOff2.OffLineCaption = tVB_CNOffCaption
    uAdaOnOff1.Online = IIf(oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP), True, False)  'oDB.IsCurOnline
    
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Me.Caption = Me.Caption & "                                                                 " & tW_Version
    '*Tee 51-02-19
    ostForm.Panels(2).Text = oDB.OnServerName & "." & oDB.OnDbName   'DB Online
    ostForm.Panels(2).ToolTipText = oDB.OnServerName & "." & oDB.OnDbName   'DB Online
    ostForm.Panels(3).Text = oDB.OffServerName & "." & oDB.OffDbName  'DB Offline
    ostForm.Panels(3).ToolTipText = oDB.OffServerName & "." & oDB.OffDbName  'DB Offline

'    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) And oTerminal.SaleMode = nEN_OnlineNormal Then
'       wBackGrand.ostForm.Panels(1).Text = "SRV"
'    Else
'       wBackGrand.ostForm.Panels(1).Text = "LOCAL"
'   End If

    Call SaveSetting("AdaPosMall", "POSFront", "RepStop", "1") '*KT 52-07-27 เริ่มให้ ServiceAutoReplicate.EXE หยุดทำงาน
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[2]")
'    DoEvents '*KT 53-01-14 DoEvents เอาออกก่อน

    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:Form_Load[3]")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
'    otmForm.Enabled = False '*KT 53-01-14 DoEvents
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    otmForm.Enabled = False '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
    bW_Unload = True
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOpen
'            Unload Me
            tVB_Transaction = "99"
            tVB_TransactionSub = "99"
'            Exit Sub  '        'แทนการออกจากโปรแกรม กรณีเปิดโปรแกรมครั้งแรก
        Case tEN_TrnStartBank:
        Case tEN_TrnSale:
        Case tEN_TrnReturn:
        Case tEN_TrnIPV:
        Case tEN_TrnDeposit:
        Case tEN_TrnRedeem:
        Case tEN_TrnClose:
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
        Case tEN_TrnCreateVIP:
        Case tEN_TrnSaleTaxInc, tEN_TrnSaleManualTax:
            tVB_CNUserAlwC = "" '*Eaw 56-08-31 เคลียร์ User กรณีกด Esc
            Unload Me
            Exit Sub
        Case 99:
        Case 100
'            Unload Me
            tVB_Transaction = "99"
            tVB_TransactionSub = "99"
'            Exit Sub   'wBackGrand.Show Exit Sub
    End Select
    Unload Me
    Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
End Sub


Private Sub ocmCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call ocmCancel_Click
    End Select
End Sub

Private Sub ocmOK_Click()
On Error GoTo ErrHandle
'    Select Case SP_VALnText2Double(tVB_TransactionSub)
'        Case tEN_TrnOpen: Call W_PRCxOpen
'        Case 100:    Call W_PRCxClose '*KT 52-12-17 POSFront 2.0.30
'    End Select
'    If bW_Unload Then
'        Unload Me
'        Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
'    Else
'        Call SP_OBJxSetFocus(otbUser)
'    End If
'    Exit Sub
    otmForm.Enabled = False '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOpen, tEN_TrnSaleTaxInc, tEN_TrnSaleManualTax
            tVB_CNAccTypeReEntryMode = ""
            Call W_PRCxOpen
            If bW_Unload Then
                Unload Me
                If tVB_CNMnuCurrent <> "1720" And tVB_CNMnuCurrent <> "1721" Then
                    Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
                End If
            Else
                 If otmForm.Enabled = False Then otmForm.Enabled = True '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
                Call SP_OBJxSetFocus(otbUser)
            End If
        Case 100
            Call W_PRCxClose '*KT 52-12-17 POSFront 2.0.30
            If bW_Unload Then
                Unload Me
            Else
                 If otmForm.Enabled = False Then otmForm.Enabled = True '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
                Call SP_OBJxSetFocus(otbUser)
            End If
        Case Else
             If otmForm.Enabled = False Then otmForm.Enabled = True '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
    End Select
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:ocmOK_Click")
End Sub


Private Sub ocmOk_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call ocmCancel_Click
    End Select
End Sub
    
Private Sub W_PRCxLogin()
'---------------------------------------------------
'Call:
'Cmt: สำหรับการ Transaction Open
'--------------------------------------------------
    Dim tSql$, tPass$
    Dim tUc$, tUp$
    Dim bErr As Boolean
    Dim bNoMsg As Boolean
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    bW_Unload = False

    tUc = Trim(otbUser.Text)
    tUp = Trim(otbPW.Text)

    
    If (tUc = "") Or (tUp = "") Then
        Call SP_MSGnShowing("tms_CN999017", nCS_Warning)
        otbUser.Text = ""
        otbPW.Text = ""
        Call SP_OBJxSetFocus(otbUser)
        Exit Sub
    End If
    
    'Call SP_USRxSupportLogin(tUc)        'Auto Gen Support Login
    'search in database
    tSql = "SELECT * FROM TCNMEmpMtn"
    tSql = tSql & vbCrLf & " WHERE (FTEmpCode = '" & tUc & "')"
    If SP_SQLvExecute(tSql, orsForm, oDB.ocnOnLine, , , , , True, True) <> 0 Then
        Call SP_MSGnShowing("tms_CN005054", nCS_Warning)
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub    'error, exit
    End If
    If Not (orsForm.BOF And orsForm.EOF) Then                       'empty database
        tSql = "SELECT FTEmpCode,FTTmnNum FROM TCNMTerminalMtn"
        tSql = tSql & vbCrLf & " WHERE FTTmnStatus = 'A'"
        tSql = tSql & vbCrLf & " AND FTEmpCode = '" & tUc & "'"
        tSql = tSql & vbCrLf & " AND FTTmnNum <> '" & tVB_CNTerminalNum & "'"
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
        If (orsTemp.BOF And orsTemp.EOF) Then                        'empty database
            tPass = SP_FEDtChkString(orsForm, "FTEmpPW")
            'ไม่มีการ Encript
            If UCase(tUp) = UCase(tPass) Then
                Call W_SETxUserValue                                                     'set user parameter
                bW_Unload = True
            Else                            'not correct password
                bErr = True             'no any user
            End If
        Else
'            If bVB_CNAlwDLogInPolicy And SP_FEDtChkString(orsForm, "FTLelCode") = tVB_CNLevelDLogInPolicy Then
            If bVB_CNAlwDLogInPolicy Then   '*Eaw 56-05-23  CommSheet TK-ISS3000-203-ST 00 00 01.xls ตรวจสอบว่า FTStpStaAlw = Y/N เท่านั้น
                tPass = SP_FEDtChkString(orsForm, "FTEmpPW")
                'ไม่มีการ Encript
                If UCase(tUp) = UCase(tPass) Then
                    Call W_SETxUserValue                                                     'set user parameter
                    bW_Unload = True
                Else                            'not correct password
                    bErr = True             'no any user
                End If
            Else
                Call SP_MSGnShowing("tms_CN005132", nCS_Warning)
                otbUser.Text = ""
                otbPW.Text = ""
                Call SP_OBJxSetFocus(otbUser)
                Exit Sub
            End If
        End If
    Else
        bErr = True             'no any user
    End If
    If bErr Then
        If Not bNoMsg Then
            Call SP_MSGnShowing("tms_CN005054", nCS_Warning)
        End If
        otbUser.Text = ""
        otbPW.Text = ""
        Call SP_OBJxSetFocus(otbUser)
    End If
    Exit Sub
ErrHandle:
    bW_Unload = False
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCxLogin")
End Sub

Private Sub W_PRCxLoginClose()
'---------------------------------------------------
'Call:
'Cmt: สำหรับการ Transaction Close
'--------------------------------------------------

    Dim tSql$, tPass$
    Dim tUc$, tUp$
    Dim bErr As Boolean
    Dim bNoMsg As Boolean
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    bW_Unload = False

    tUc = Trim(otbUser.Text)
    tUp = Trim(otbPW.Text)

    If (tUc = "") Or (tUp = "") Then
        Call SP_MSGnShowing("tms_CN999017", nCS_Warning)
        otbUser.Text = ""
        otbPW.Text = ""
        Call SP_OBJxSetFocus(otbUser)
        Exit Sub
    End If
    
    'Call SP_USRxSupportLogin(tUc)        'Auto Gen Support Login
    'search in database
    tSql = "SELECT * FROM TCNMEmpMtn"
    tSql = tSql & " WHERE (FTEmpCode = '" & tUc & "')"
    If SP_SQLvExecute(tSql, orsForm, oDB.ocnOnLine, , , , , True, True) <> 0 Then
        Call SP_MSGnShowing("tms_CN005054", nCS_Warning)
        Call SP_DATxRsClose(orsTemp)
        Exit Sub      'error, exit
    End If
    If Not (orsForm.BOF And orsForm.EOF) Then                       'empty database
        If bUT_SignOut Then
            tPass = SP_FEDtChkString(orsForm, "FTEmpPW")
             If UCase(tUp) = UCase(tPass) Then
                If SP_USRtMnuGetAutho(nEN_Usr2AthRead, tVB_CNMnuCurrent, tUc) = "0" Then      '0:ไม่ผ่าน, 1:ผ่าน  '***KT 51-12-04 V 1.0.1 แก้ไขเรื่องการทำ Exit Terminal ไม่ได้เมื่อ User/Password เหมือนกันเพราะส่งค่า User ผิด ดันส่ง Password ไปแทน
                    Call SP_MSGnShowing("tms_CN006008", tCS_CNMsgWarn)
                Else
                    bW_Unload = True
                End If
            Else                            'not correct password
                bErr = True             'no any user
            End If
        Else
            tPass = SP_FEDtChkString(orsForm, "FTEmpPW")
            'ไม่มีการ Encript
            If UCase(tUp) = UCase(tPass) Then
                If SP_USRtMnuGetAutho(nEN_Usr2AthRead, tVB_CNMnuCurrent, tUc) = "0" Then      '0:ไม่ผ่าน, 1:ผ่าน  '***KT 51-12-04 V 1.0.1 แก้ไขเรื่องการทำ Exit Terminal ไม่ได้เมื่อ User/Password เหมือนกันเพราะส่งค่า User ผิด ดันส่ง Password ไปแทน
                    Call SP_MSGnShowing("tms_CN006008", tCS_CNMsgWarn)
                Else
                    bW_Unload = True
                End If
            Else                            'not correct password
                bErr = True             'no any user
            End If
        End If
    Else
        bErr = True             'no any user
    End If
    If bErr Then
        If Not bNoMsg Then
            Call SP_MSGnShowing("tms_CN005054", nCS_Warning)
        End If
        otbUser.Text = ""
        otbPW.Text = ""
        Call SP_OBJxSetFocus(otbUser)
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub W_SETxUserValue()
'----------------------------------------------------------
'   Call:
'   Ret:    occur when login correctly
'   Modify By :
'----------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    
    bUT_FuncLockKB = False      'set default lock KB
    
    tVB_CNUserCode = SP_FEDtChkString(orsForm, "FTEmpCode")
    tVB_CNUserName = Trim(SP_FEDtChkString(orsForm, "FTEmpFName") & " " & SP_FEDtChkString(orsForm, "FTEmpLName"))
    tUT_UsrLevel = SP_FEDtChkString(orsForm, "FTLelCode")
    tUT_UsrPwd = SP_FEDtChkString(orsForm, "FTEmpPW")
    
    tSql = "SELECT FTCfgCode,FTTmnName "
    tSql = tSql & " FROM TConfigTerminal"
    tSql = tSql & " WHERE FTCfgCode='LEmpID'"
    Set orsTemp = ocnCfg.Execute(tSql)
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        If orsTemp.RecordCount > 1 Then
            tSql = "DELETE FROM"
            tSql = tSql & "  TConfigTerminal"
            tSql = tSql & " WHERE FTCfgCode='LEmpID'"
            Call ocnCfg.Execute(tSql)
            tSql = "INSERT INTO TConfigTerminal"
            tSql = tSql & "(FTTmnCode,FTTmnName,FTCfgCode,FTCfgOn,FTCfgOff,FTCfgValue,FTCfgDescTha,FTCfgDescEng)"
            tSql = tSql & " VALUES ("
            tSql = tSql & "'" & tVB_CNTerminalNum & "'"
            tSql = tSql & ",'" & tVB_CNComName & "'"
            tSql = tSql & ",'" & "LEmpID" & "'"
            tSql = tSql & ",'" & "" & "'"
            tSql = tSql & ",'" & "" & "'"
            tSql = tSql & ",'" & tVB_CNUserCode & "'"
            tSql = tSql & ",'" & "รหัสพนักงานล่าสุด" & "'"
            tSql = tSql & ",'" & "Last EmployeeID" & "'"
            tSql = tSql & ")"
        Else
            tSql = "UPDATE"
            tSql = tSql & "  TConfigTerminal"
            tSql = tSql & " SET FTCfgValue = '" & tVB_CNUserCode & "'"
            tSql = tSql & " ,FTTmnName = '" & tVB_CNComName & "'"
            tSql = tSql & " ,FTTmnCode = '" & tVB_CNTerminalNum & "'"
            tSql = tSql & " WHERE FTCfgCode='LEmpID'"
        End If
    Else
        tSql = "INSERT INTO TConfigTerminal"
        tSql = tSql & "(FTTmnCode,FTTmnName,FTCfgCode,FTCfgOn,FTCfgOff,FTCfgValue,FTCfgDescTha,FTCfgDescEng)"
        tSql = tSql & " VALUES ("
        tSql = tSql & "'" & tVB_CNTerminalNum & "'"
        tSql = tSql & ",'" & tVB_CNComName & "'"
        tSql = tSql & ",'" & "LEmpID" & "'"
        tSql = tSql & ",'" & "" & "'"
        tSql = tSql & ",'" & "" & "'"
        tSql = tSql & ",'" & tVB_CNUserCode & "'"
        tSql = tSql & ",'" & "รหัสพนักงานล่าสุด" & "'"
        tSql = tSql & ",'" & "Last EmployeeID" & "'"
        tSql = tSql & ")"
    End If
    Call ocnCfg.Execute(tSql)
    
    tSql = "SELECT FTCfgCode,FTTmnName "
    tSql = tSql & " FROM TConfigTerminal"
    tSql = tSql & " WHERE FTCfgCode='LEmpPW'"
    Set orsTemp = ocnCfg.Execute(tSql)
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        If orsTemp.RecordCount > 1 Then
            tSql = "DELETE FROM"
            tSql = tSql & "  TConfigTerminal"
            tSql = tSql & " WHERE FTCfgCode='LEmpPW'"
            Call ocnCfg.Execute(tSql)
            tSql = "INSERT INTO TConfigTerminal"
            tSql = tSql & "(FTTmnCode,FTTmnName,FTCfgCode,FTCfgOn,FTCfgOff,FTCfgValue,FTCfgDescTha,FTCfgDescEng)"
            tSql = tSql & " VALUES ("
            tSql = tSql & "'" & tVB_CNTerminalNum & "'"
            tSql = tSql & ",'" & tVB_CNComName & "'"
            tSql = tSql & ",'" & "LEmpPW" & "'"
            tSql = tSql & ",'" & "" & "'"
            tSql = tSql & ",'" & "" & "'"
            tSql = tSql & ",'" & tUT_UsrPwd & "'"
            tSql = tSql & ",'" & "รหัสผ่านพนักงานล่าสุด" & "'"
            tSql = tSql & ",'" & "Last Employee Password" & "'"
            tSql = tSql & ")"
        Else
            tSql = "UPDATE"
            tSql = tSql & "  TConfigTerminal"
            tSql = tSql & " SET FTCfgValue = '" & tUT_UsrPwd & "'"
            tSql = tSql & " ,FTTmnName = '" & tVB_CNComName & "'"
            tSql = tSql & " ,FTTmnCode = '" & tVB_CNTerminalNum & "'"
            tSql = tSql & " WHERE FTCfgCode='LEmpPW'"
        End If
    Else
        tSql = "INSERT INTO TConfigTerminal"
        tSql = tSql & "(FTTmnCode,FTTmnName,FTCfgCode,FTCfgOn,FTCfgOff,FTCfgValue,FTCfgDescTha,FTCfgDescEng)"
        tSql = tSql & " VALUES ("
        tSql = tSql & "'" & tVB_CNTerminalNum & "'"
        tSql = tSql & ",'" & tVB_CNComName & "'"
        tSql = tSql & ",'" & "LEmpPW" & "'"
        tSql = tSql & ",'" & "" & "'"
        tSql = tSql & ",'" & "" & "'"
        tSql = tSql & ",'" & tUT_UsrPwd & "'"
        tSql = tSql & ",'" & "รหัสผ่านพนักงานล่าสุด" & "'"
        tSql = tSql & ",'" & "Last Employee Password" & "'"
        tSql = tSql & ")"
    End If
    Call ocnCfg.Execute(tSql)
    Call SP_DATxRsNothing(orsTemp)
      
    nVB_CNCutLang = 1       'first default
    cUT_Print.UT_SETxPosLang = nVB_CNCutLang - 1
    nVB_CNCutMsg = IIf(nVB_CNCutLang = 1, 1, 2)
    nVB_CNCutTag = IIf(nVB_CNCutLang = 1, 2, 3)
    nVB_CNCutConst = IIf(nVB_CNCutLang = 1, 1, 2)
End Sub

Private Sub W_PRCxOpen()
'-----------------------------------------------
'Call:
'Cmt: Mode การ Open
'----------------------------------------------
    Dim nMaxCode As Long
    Dim nMaxLogIn As Long
    Dim tMaxCode$
    Dim orsTemp As ADODB.Recordset
    Dim tSql$
    Dim tErr$
On Error GoTo ErrHandle
    
    Call W_PRCxLogin
    If bW_Unload Then
        If Not W_PRCbVerifyEmpCode Then
            Call SP_MSGnShowing("tms_CN005075", nCS_Warning)
            Call SP_OBJxSetFocus(otbUser)
            bW_Unload = False
            Exit Sub
        End If
        If aVB_MainMnu(nVB_CurentMainMnu).Autho Then
             If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                 Exit Sub
            End If
        End If
        '*Eaw 56-09-05
'        If W_PRCbVerifyTmnNum Then
'            Call SP_MSGnShowing("tms_CN005077", nCS_Warning)
'            Call SP_OBJxSetFocus(otbUser)
'            bW_Unload = False
'            Exit Sub
'        End If
        Call SP_STAxMouse(Me, True)
        '*KT 53-02-03 ทำการย้ายจาก wBackGrand/W_PRCxOpen เนื่องจากเหตุผล 2 ประการ
        '1. เครื่อง T.2000 ถ้ามีการทำงานใน Form_Activate ใช้เวลานาน จะทำให้ wCNLogIn จะซ่อนอยู่ด้านหลัง wBackGrand
        '2. กรณีที่กดเลือกเมนู Open จากหน้าจอ Main Menu ระบบจะทำการลบข้อมูลใน Local ที่เป็นตาราง Temp ทั้งที่ ทั้งที่ยังไม่มีการใส่ Code Cashier
        '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If Not bVB_CNCreateTblTmp Or Not bVB_CNCreateTblTmpFrmCloseRpt Then
            Call UT_CreateSaleTableTMP
            bVB_CNCreateTblTmp = True
        End If
       'PH 2.0.0 ตรวจสอบข้อมูลที่ Local ตลอด
        tSql = "select  name from Sys.sysobjects where name = '" & "TRG_Tmp2Sale" & tVB_CNTerminalNum & "'"
        Call SP_SQLvExecute(tSql, orsTemp)
        If orsTemp.EOF And orsTemp.EOF Then
            '*KT 52-01-14 กรณีตรวจสอบเจอว่าไม่มี Trigger แล้วยังสร้างไม่สำเร็จทั้งฐานข้อมูล server และ local
            If Not SP_DATbCrtTrigger(False, True) Then
                  Call SP_MSGnShowing("tms_CN005168", nCS_Warning)
                  Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Tricger" & " Form:" & Me.Name & " Step: W_PRCxOpen" & " Events:SP_DATbCrtTrigger")
            End If
        End If
        '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If tVB_CNMnuCurrent = "1720" Or tVB_CNMnuCurrent = "1721" Or tVB_CNMnuCurrent = "1722" Then '*Eaw 56-08-27 เคส 1720/1721 ไม่ Running Log
            'none
        Else
            nMaxCode = CLng(Val(Format(SP_SQLtFunction("MAX", "RIGHT(FTLogCode,5)", "TPSTSalHD", " WHERE Left(FTLogCode,9)='" & SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & "'", oDB.ocnOnLine, True))))
            'เปรียบเทียบระหว่าง Master กับ transaction
            If nMaxCode > nVB_CNLastLogin Then
                 nMaxLogIn = nMaxCode
            Else
                nMaxLogIn = nVB_CNLastLogin
            End If
            tMaxCode = Format(nMaxLogIn + 1, "00000")
            nVB_CNLastLogin = nMaxLogIn + 1 '*KT 52-01-21 เก็บค่าไว้ในตัวแปร
            tUT_LoginCode = SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & tMaxCode
        
'        If tVB_CNMnuCurrent = "1720" Or tVB_CNMnuCurrent = "1721" Or tVB_CNMnuCurrent = "1722" Then '*Eaw 56-08-27 เคส 1720/1721 ไม่ Running Log จึงย้ายขึ้นไปทำก่อนหน้า
'            'none
'        Else
            Call SP_INSxTransHD(tVB_CNTblHD)
            Call wFunctionKB.W_SetToComplete
            'Comsheet ML-SF-289
            Call SP_UPDxSetCompleteTerminalMtn
            Call SaveSetting("AdaPosMall", "POSFront", "CountBill", 1)
            Call SaveSetting("AdaPosMall", "POSFront", "TmnNum", tVB_CNTerminalNum)
            Call SaveSetting("AdaPosMall", "POSFront", "Mode", "1")
            Call SaveSetting("AdaPosMall", "POSFront", "AccType", tVB_CNAccTypeReEntryMode)
            'PH 2.0 ML-SF-050 Closing Report ไม่เท่ากับ Flash Sale ป้องกันฐานข้อมูล Local ไม่ Update Status Open ของ Terminal
            If oTerminal.SaleMode = nEN_OnlineNormal Then
                If Not SP_TMNxUpdateOpen(oTerminal.SaleMode) Then
                    Call SP_STAxMouse(Me, False)
                    Call SP_MSGnShowing("tms_CN005211", tCS_CNMsgErr)
                    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:SP_TMNxUpdateOpen =false" & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCxOpen")
                    Exit Sub
                End If
            Else
                If Not SP_TMNxUpdateOpenEx(oTerminal.SaleMode) Then
                    Call SP_STAxMouse(Me, False)
                    Call SP_MSGnShowing("tms_CN005211", tCS_CNMsgErr)
                    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:SP_TMNxUpdateOpenEx =false" & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCxOpen")
                    Exit Sub
                 End If
            End If
            tVB_Transaction = aVB_TransMode(1).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(1).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbStartBank - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbStartBank - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(1).PrintType
    '        Set wMain = wMainRet '*KT 52-02-23 640 or 800
            bUT_SignOut = False
            bVB_CNUnLock = False
            Call SaveSetting("AdaPosMall", "POSFront", "User", tVB_CNUserCode)
            Call SaveSetting("AdaPosMall", "POSFront", "PW", tVB_CNUserName)
            Call SP_STAxMouse(Me, False)
        End If
    Else
        Call SP_OBJxSetFocus(otbUser)
        bW_Unload = False
    End If
    Exit Sub
ErrHandle:
    Call SP_STAxMouse(Me, False)
    bW_Unload = False
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCxOpen")
End Sub

Private Sub W_PRCxClose()
'-----------------------------------------------
'Call:
'Cmt: Mode การ Open
'----------------------------------------------
    Call SP_STAxMouse(Me, True)
    Call W_PRCxLoginClose
    If bW_Unload Then
        'Call SP_PRCxDropTblTmp(True, False)  'ปัจจุบันไม่ได้ทำงานแล้ว เพราะส่ง parameter ตัวที่ 2 เป็น False
        'move file ej offline to path online
        '*KT 52-05-15 PH 1.8 ML-SF-060
        'EJ ไม่ส่งมายัง Server ขั้นตอน  1. เปิด POS แบบ Online (Open-> Close-->)  2. Disconnect LAN 3. Open POS in Offline 4. Sale --> Close --> Exit 5. Open POS in Online (Open -->Sale-->Close) 6. Print Closing Report --> Exit        EJ ที่ขายแบบ Offline จะไม่ส่งมายัง Server
        'ป้องกัน 1 เปิดโปรแกรม 2 ตอนที่เปิดโปรแกรมแล้วเข้าหน้าจอ Sale ทันที 3 ตอน Exit Terminal 4 ตอนจบบิลแล้วมีการสลับมาที่ On-line
        Call SP_MOVxEJToOnline
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Exit Terminal" & " Form:" & Me.Name & " Control:General" & " Events:W_PRCxClose")
        Call SP_APPxEndProgram
    Else
        Call SP_OBJxSetFocus(otbUser)
        bW_Unload = False
    End If
    Call SP_STAxMouse(Me, False)
End Sub



'Private Sub W_UPDxTransHD()
''------------------------------------------------------
''Call:
''Cmt:   *KT 50-12-24 ปรับปรุงข้อมูล Transaction HD
''-----------------------------------------------------
'    Dim tSql$
'    tSql = "UPDATE " & tVB_CNTblHD
'    tSql = tSql & " SET FTEmpCode ='" & tVB_CNUserCode & "'"
'    tSql = tSql & ",FTShdStaDoc = '1'"
'    tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
'    tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
'    tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
'    tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
'    tSql = tSql & "  AND FTShdStaDoc = '2'"
'    Call SP_SQLbExecute(tSql)
'End Sub

Private Sub otbPW_GotFocus()
    otbPW.SelStart = 0
    otbPW.SelLength = Len(otbPW.Text)
End Sub

Private Sub otbPW_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbPW.Text = ""
    End Select
End Sub

Private Sub otbPW_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call ocmOK_Click
End Sub

Private Sub otbUser_GotFocus()
    otbUser.SelStart = 0
    otbUser.SelLength = Len(otbUser.Text)
End Sub

Private Sub otbUser_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbUser.Text = ""
        Case vbKeyReturn
            SendKeys "{Tab}"
        Case 13
            If Trim(otbUser.Text) <> "" Then
                SendKeys "{Tab}"
            End If
    End Select
'    DoEvents '*KT 53-01-14 DoEvents
End Sub

Private Sub otbUser_LostFocus()
    If UCase(otbUser.Text) = UCase(tCS_CNPwdNoUsr) Then
        otbPW.MaxLength = 6
    Else
        otbPW.MaxLength = 4
    End If
End Sub

Private Sub otmForm_Timer()
Dim tSrvComName$
    If oDB.SQLbAvailableServer(oDB.OnServerName, oDB.OnComIP) Then
        uAdaOnOff1.Online = True
        uAdaOnOff2.Online = True
        'wBackGrand.ostForm.Panels(3).Text = oDB.OnServerName & "." & oDB.OnDbName   'DB Online 'PH 2.0.0 OLD
'        DoEvents '*KT 53-01-14 DoEvents
    Else
        uAdaOnOff1.Online = False
        uAdaOnOff2.Online = False
        'wBackGrand.ostForm.Panels(3).Text = oDB.OffServerName & "." & oDB.OffDbName  'DB Offline 'PH 2.0.0 OLD
'        DoEvents        '*KT 53-01-14
    End If
   
'    If oDB.SQLbAvailableServer(oDB.OnServerName, oDB.OnComIP) And oTerminal.SaleMode = nEN_OnlineNormal Then
'       wBackGrand.ostForm.Panels(1).Text = "SRV"
'    Else
'       wBackGrand.ostForm.Panels(1).Text = "LOCAL"
'   End If

End Sub

Private Function W_PRCbVerifyEmpCode() As Boolean
'--------------------------------------------------------
'Call:
'Cmt: ทำการตรวจสอบว่า Cashier มีการใช้งานกับเครื่องอื่นหรือไม่
'-------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT FTEmpCode "
    tSql = tSql & " FROM TCNMTerminalMtn"
    tSql = tSql & " WHERE FTEmpCode = '" & tVB_CNUserCode & "'"
    tSql = tSql & " AND  FTTmnName <> '" & tVB_CNComName & "'"
    tSql = tSql & " AND ISNULL(FTTmnStaClossing,'N') <> 'Y'"
    tSql = tSql & " AND FTTmnStatus = 'A'"
'    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
    If oDB.SQLbAvailableServer(oDB.OnServerName, oDB.OnComIP) Then   '*Eaw 56-05-23  CommSheet TK-ISS3000-203-ST 00 00 01.xls
'        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, True, , , , True, True)
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    Else
        Call SP_SQLvExecute(tSql, orsTemp)
    End If
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        W_PRCbVerifyEmpCode = True
    Else
'        If tUT_UsrLevel = "00001" Then
'            W_PRCbVerifyEmpCode = True
'        Else
'            W_PRCbVerifyEmpCode = False
'        End If
        W_PRCbVerifyEmpCode = True   '*Eaw 56-05-23  CommSheet TK-ISS3000-203-ST 00 00 01.xls  ตรวจสอบว่า FTStpStaAlw = Y/N เท่านั้น
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_PRCbVerifyEmpCode = False
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
End Function

