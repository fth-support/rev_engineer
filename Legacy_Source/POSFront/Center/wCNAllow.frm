VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form wCNAllow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "อนุมัติโดย;klg_Title"
   ClientHeight    =   3780
   ClientLeft      =   1605
   ClientTop       =   6885
   ClientWidth     =   6405
   ControlBox      =   0   'False
   Icon            =   "wCNAllow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6405
   Tag             =   "2;"
   Begin POSFront.uAdaOnOff uAdaOnOff2 
      Height          =   255
      Left            =   25
      TabIndex        =   12
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
      PicOnLine       =   "wCNAllow.frx":6852
      PicOffLine      =   "wCNAllow.frx":6F4C
   End
   Begin MSComctlLib.StatusBar ostForm 
      Height          =   420
      Left            =   0
      TabIndex        =   11
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
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl otcForm3 
      Height          =   1335
      Left            =   15
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2250
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
         TabIndex        =   5
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
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
         Picture         =   "wCNAllow.frx":7646
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1605
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm4 
      Height          =   1335
      Left            =   1815
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2250
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
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   3
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
         TabIndex        =   2
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2535
      Left            =   1815
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   4575
      _Version        =   720897
      _ExtentX        =   8070
      _ExtentY        =   4471
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
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
         TabIndex        =   10
         Tag             =   "2;"
         Top             =   1080
         Width           =   2145
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
         TabIndex        =   9
         Tag             =   "2;"
         Top             =   120
         Width           =   2835
      End
   End
   Begin VB.Image oimForm2 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   11040
      Picture         =   "wCNAllow.frx":CC29
      Stretch         =   -1  'True
      Top             =   8985
      Width           =   1455
   End
End
Attribute VB_Name = "wCNAllow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tW_ID As String

Private orsForm As ADODB.Recordset
Private tW_Version As String
Private bW_Unload As Boolean
Private tW_Usr As String
Private tW_Name As String
Private tW_Pw  As String

Private Function W_CHKbDoubleLoginUser(ByVal ptUc As String) As Boolean
'----------------------------------------------------------
'Call:
'Cmt : ตรวจสอบว่า User นี้มีการใช้งานที่เครื่อง POS  อื่นหรือไม่ ถ้าไม่มีให้ผ่านและทำการ Overwrite EmpCode ใหม่แทน EmpCode เดิมในตาราง TCNMTerminalMtn*Eaw 56-26-09  CommSheet TK-ISS3000-361-FAPL-127.xls
'----------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT FTEmpCode "
    tSql = tSql & vbCrLf & " FROM TCNMTerminalMtn "
    tSql = tSql & vbCrLf & " WHERE FTEmpCode  = '" & ptUc & "'"
'    tSql = tSql & vbCrLf & " AND FTTmnNum <> '" & tVB_CNTerminalNum & "'"
    '*Eaw 56-10-14 แยกเคส Double Ligin ระหว่าง Re-Entry กับ Normal Mode
    If bVB_ReEntryMode Then
        tSql = tSql & vbCrLf & " AND ( FTTmnNum <> '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND FTTmnNum <> '" & tVB_CNVirTerminalNum & "')"
    Else
        tSql = tSql & vbCrLf & " AND FTTmnNum <> '" & tVB_CNTerminalNum & "'"
    End If
    tSql = tSql & vbCrLf & "AND FTTmnStatus = 'A'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        W_CHKbDoubleLoginUser = True
    Else
        W_CHKbDoubleLoginUser = False
        '*Eaw 56-10-02 Overwrite new EmpCode
'        tSql = "UPDATE TCNMTerminalMtn"                                    '*Ao 2013/10/15   V4.0.82 ย้ายไปที่  W_PRCxTrn
'        tSql = tSql & vbCrLf & " SET FTEmpCode = '" & ptUc & "'"
'        tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'        Call SP_SQLvExecute(tSql)
'        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then Call SP_SQLvExecute(tSql, , oDB.ocnOnLine)
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGShowError
End Function

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then SendKeys "{Tab}"
'End Sub

Private Sub Form_Load()
Dim tSql  As String

On Error GoTo ErrHandle
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[1]")
    Call SP_STAxMouse(Me, True)
    tW_Version = "V." & App.Major & "." & App.Minor & "." & App.Revision
    uAdaOnOff1.ShowPic = False
    uAdaOnOff1.OnLineCaption = tVB_CNOnCaption
    uAdaOnOff1.OffLineCaption = tVB_CNOffCaption
    uAdaOnOff1.Online = IIf(oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP), True, False) ' oDB.IsCurOnline
    uAdaOnOff2.ShowPic = True
    uAdaOnOff2.OnLineCaption = tVB_CNOnCaption
    uAdaOnOff2.OffLineCaption = tVB_CNOffCaption
    uAdaOnOff2.Online = IIf(oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP), True, False) ' oDB.IsCurOnline
    
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Me.Caption = Me.Caption & "                                                                 " & tW_Version

    ostForm.Panels(2).Text = oDB.OnServerName & "." & oDB.OnDbName   'DB Online
    ostForm.Panels(3).Text = oDB.OffServerName & "." & oDB.OffDbName  'DB Offline
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnPrintClose
            ocmCancel.Enabled = False
        Case Else
            ocmCancel.Enabled = True
    End Select
    '------------------------------------------------------------------'*Ao 2013/10/15   V4.0.82
    If bVB_ChkMode Then
        If bVB_CheckerMode Then
            tSql = "UPDATE TCNMTerminalMtn"                                 '*Ao 2013/10/15   V4.0.82
            tSql = tSql & vbCrLf & "SET FTEmpCode = ''"
            If bVB_ReEntryMode Then
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNVirTerminalNum & "'"
            Else
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
            End If
            Call SP_SQLvExecute(tSql)
            If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then Call SP_SQLvExecute(tSql, , oDB.ocnOnLine)
            wMain.ostForm.Panels(4).Text = ""
        End If
    End If
    '------------------------------------------------------------------
    'DoEvents '*KT 53-01-14 DoEvents
    Call SP_STAxMouse(Me, False)
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[2]")
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[3]")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
     'otmForm.Enabled = False '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:
    
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Set wCNAllow = Nothing
'End Sub

Private Sub ocmCancel_Click()
    otmForm.Enabled = False '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
    tVB_CNUserAlwC = ""
    tVB_CNUserAlwN = ""
    bVB_CNPwdAllow = False
'    Unload Me
    '*Eaw 56-05-31 CommSheet TK-ISS3000-215-ST 00 00 01.xls
'    If Not bVB_CheckerMode Then
'        Unload Me
'    End If
    If bVB_ChkMode = False Then
        Unload Me               '*Ao 2013/10/15   V4.0.80
    End If
    
End Sub

Private Sub ocmCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call ocmCancel_Click
    End Select
End Sub

Private Sub ocmOK_Click()
    Call W_PRCxTrn
End Sub

Private Sub ocmOk_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call ocmCancel_Click
    End Select
End Sub
    
Private Sub W_PRCxLogin()
    Dim tSql$, tPass$
    Dim tUc$, tUp$
    Dim bErr As Boolean
    Dim bNoMsg As Boolean

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
            
    tSql = "SELECT * FROM TCNMEmpMtn"
    tSql = tSql & " WHERE FTEmpCode = '" & tUc & "'"

    If SP_SQLvExecute(tSql, orsForm, oDB.ocnOnLine, , , , , True, True) <> 0 Then
        Call SP_MSGnShowing("tms_CN005054", nCS_Warning)
        Exit Sub
    End If
    If Not (orsForm.BOF And orsForm.EOF) Then                       'empty database
'        If Not bVB_CNAlwDLogInPolicy Then   '*Eaw 56-26-09  CommSheet TK-ISS3000-361-FAPL-127.xls ถ้า FTStpStaAlw = N ไม่ให้ผ่าน
'            If W_CHKbDoubleLoginUser(tUc) Then
'                Call SP_MSGnShowing("tms_CN005132", nCS_Warning)
'                otbUser.Text = ""
'                otbPW.Text = ""
'                Call SP_OBJxSetFocus(otbUser)
'                Exit Sub
'            End If
'        End If

        tPass = SP_FEDtChkString(orsForm, "FTEmpPW")
        'ไม่มีการ Encript
        If UCase(tUp) = UCase(tPass) Then
            tW_Usr = SP_FEDtChkString(orsForm, "FTEmpCode")
            tW_Pw = SP_FEDtChkString(orsForm, "FTEmpPW")
            tW_Name = Trim(SP_FEDtChkString(orsForm, "FTEmpFName") & " " & SP_FEDtChkString(orsForm, "FTEmpLName"))
'            If SP_USRtMnuGetAutho(nEN_Usr2AthRead, tW_ID, tW_Usr) = "0" Then       '0:ไม่ผ่าน, 1:ผ่าน
'                Call SP_MSGnShowing("tms_CN006008", tCS_CNMsgWarn)             '*Ao 2013/10/15   V4.0.82
'                tW_Usr = ""
'                tW_Pw = ""
'                tW_Name = ""
'                Exit Sub
'            End If
            If bVB_ChkMode = False Then
                If SP_USRtMnuGetAutho(nEN_Usr2AthRead, tW_ID, tW_Usr) = "0" Then       '0:ไม่ผ่าน, 1:ผ่าน
                    Call SP_MSGnShowing("tms_CN006008", tCS_CNMsgWarn)          '*Ao 2013/10/15   V4.0.82
                    tW_Usr = ""
                    tW_Pw = ""
                    tW_Name = ""
                    Exit Sub
                End If
            End If
'            bW_Unload = True
            '---------------------------------------------------------------------------------'*Ao 2013/10/17   ย้ายลงมาจากข้างบน
            If Not bVB_CNAlwDLogInPolicy Then   '*Eaw 56-26-09  CommSheet TK-ISS3000-361-FAPL-127.xls ถ้า FTStpStaAlw = N ไม่ให้ผ่าน
                If W_CHKbDoubleLoginUser(tUc) Then
                    Call SP_MSGnShowing("tms_CN005132", nCS_Warning)
                    otbUser.Text = ""
                    otbPW.Text = ""
                    Call SP_OBJxSetFocus(otbUser)
                    Exit Sub
                Else
                    bW_Unload = True
                End If
            Else
                bW_Unload = True
            End If
            '---------------------------------------------------------------------------------
            Unload Me
        Else                            'not correct password
            bErr = True
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
End Sub

Private Sub W_PRCxTrn()
'-----------------------------------------------
'Call:
'Cmt: Mode การ Transaction ทั่วไป
'----------------------------------------------
Dim tSql  As String

    otmForm.Enabled = False '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
    Call W_PRCxLogin
    If bW_Unload Then
'        tVB_CNUserAlwC = tW_Usr           'allowed user code
'        tVB_CNUserAlwN = tW_Name           'allowed user name
'        bVB_CNPwdAllow = True
'        If bVB_CheckerMode = True Then          '*Aoo 2012/07/31 PRJ10001TK-External Specification Function Menu-00.02.00.doc
'            tVB_CNUserCode = tW_Usr
'            tVB_CNUserName = tW_Name
''            wMain.ostForm.Panels(4).Text = "User : " & tVB_CNUserCode & "/" & tVB_CNUserName   '*Eaw 56-06-04 CommSheet TK-ISS3000-215-ST 00 00 01.xls
'            wMain.ostForm.Panels(4).Text = "User : " & tVB_CNUserName '*Eaw 56-07-15 CommSheet TK-ISS3000-278 TR FAPL-058.xls
'        End If
        '-----------------------------------------------------------
        tVB_CNUserAlwC = tW_Usr           'allowed user code
        tVB_CNUserAlwN = tW_Name           'allowed user name
        bVB_CNPwdAllow = True
        If bVB_ChkMode Then
'            tVB_CNUserAlwC = tW_Usr           'allowed user code
'            tVB_CNUserAlwN = tW_Name           'allowed user name
'            bVB_CNPwdAllow = True
            If bVB_CheckerMode = True Then          '*Aoo 2012/07/31 PRJ10001TK-External Specification Function Menu-00.02.00.doc
                tVB_CNUserCode = tW_Usr
                tVB_CNUserName = tW_Name
    '            wMain.ostForm.Panels(4).Text = "User : " & tVB_CNUserCode & "/" & tVB_CNUserName   '*Eaw 56-06-04 CommSheet TK-ISS3000-215-ST 00 00 01.xls
                wMain.ostForm.Panels(4).Text = "User : " & tVB_CNUserName '*Eaw 56-07-15 CommSheet TK-ISS3000-278 TR FAPL-058.xls
                tSql = "UPDATE TCNMTerminalMtn"
                tSql = tSql & vbCrLf & "SET FTEmpCode = '" & tW_Usr & "'"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
                Call SP_SQLvExecute(tSql)
                If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then Call SP_SQLvExecute(tSql, , oDB.ocnOnLine)
            End If
        End If
        '-----------------------------------------------------------
        Unload Me
    Else
        If otmForm.Enabled = False Then otmForm.Enabled = True '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุดจากการที่ Timer ทำงานพร้อม Unload
        Call SP_OBJxSetFocus(otbUser)
    End If
End Sub

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
    If KeyAscii = 13 Then
        Call ocmOK_Click
    End If
End Sub

Private Sub otbUser_GotFocus()
    otbUser.SelStart = 0
    otbUser.SelLength = Len(otbUser.Text)
End Sub

Private Sub otbUser_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbUser.Text = ""
        Case 13
            If Trim(otbUser.Text) <> "" Then
                SendKeys "{Tab}"
            End If
    End Select
End Sub

Private Sub otbUser_LostFocus()
    If UCase(otbUser.Text) = UCase(tCS_CNPwdNoUsr) Then
        otbPW.MaxLength = 6
    Else
        otbPW.MaxLength = 4
    End If
End Sub

Private Sub otmForm_Timer()
    If oDB.SQLbAvailableServer(oDB.OnServerName, oDB.OnComIP) Then
        uAdaOnOff1.Online = True
        uAdaOnOff2.Online = True
    Else
        uAdaOnOff1.Online = False
        uAdaOnOff2.Online = False
    End If
End Sub





