VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNForeignPay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Foreign Currency;klg_Title"
   ClientHeight    =   2595
   ClientLeft      =   3285
   ClientTop       =   5430
   ClientWidth     =   5100
   ControlBox      =   0   'False
   Icon            =   "wCNForeignPay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5100
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1935
      Index           =   2
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   5085
      _Version        =   720897
      _ExtentX        =   8969
      _ExtentY        =   3413
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbNum 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   0
         Text            =   "0"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :;klg_Amt"
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
         Left            =   240
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label olaCurLocal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   1680
         TabIndex        =   5
         Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
         Top             =   480
         Width           =   2970
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Needed :;klg_Cur"
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
         Left            =   240
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   480
         Width           =   1260
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1395
      Width           =   5085
      _Version        =   720897
      _ExtentX        =   8969
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Left            =   3360
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
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
         Left            =   1800
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNForeignPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub W_SETxDefault()
    Dim tPrint As String '*Eaw 56-08-08
    Dim tForCurrency As String '*Eaw 56-08-08
    With otbNum
        .Text = 0
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    '*Eaw 56-08-08 CommSheet TK-ISS3000-271-TR.xls
    If bVB_OPOSLineDisplayEnable Then
        tPrint = "NEEDED:" & Space(13 - ((Len(tVB_ForeignCurSign)) + Len(SP_FMTcCurPoint(cVB_NeedAmt)))) & tVB_ForeignCurSign & SP_FMTcCurPoint(cVB_NeedAmt) & ";"
        tPrint = tPrint & "BALANCE:" & Space(12 - Len(SP_FMTcCurPoint(cVB_Balance))) & SP_FMTcCurPoint(cVB_Balance)
        Call SP_PrnPollDisplay(tPrint)
    End If
End Sub

Private Sub Form_Activate()
    Call SP_OTBxSelText(otbNum)
End Sub

'*KT 53-01-19 ไม่ใช้งานแล้ว ทำให้เกิดการ setfocus มายัง TabIndex 1
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'      If KeyCode = 13 Then SendKeys "{Tab}"
'End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_SETxDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    cVB_Pay = 0
     '*Bump 56-12-04   เพิ่มการ Clear 3 Line V.4.0.109
    tVB_FTRteCode = ""
    cVB_FCSrcRteAmt = 0
    cVB_FCSrcRteFac = 0
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim cPay As Double
    Dim cChange As Double '*Eaw 56-08-08
'    If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbNum.Text))) Then Exit Sub       '*AO51-03-26   Comm.Sheet  ML-SF-100
    
    cVB_FCSrcRteAmt = SP_VALcText2Double(otbNum.Text)
    cPay = SP_VALcText2Double(otbNum.Text) * cVB_FCSrcRteFac
'    cPay = cVB_FCSrcRteAmt * SP_FMTcCurPoint(cVB_FCSrcRteFac, True, True)  '*Eaw 56-03-19 แก้ไขการ Rounding สำหรับ TK Project
    cPay = SP_FMTcCurPoint(cPay, True, True)
    If cPay = 0 Then 'check amount zero
        Call SP_MSGnShowing("tms_CN005064", nCS_Error)
        Exit Sub
    End If
    '*Eaw 56-09-17 CommSheet TK-ISS3000-358-FAPL-124_Eaw.xls
    If cVB_MinDeno <> 0 Then
        If (cVB_FCSrcRteAmt < cVB_MinDeno) Then
            Call SP_MSGnShowing("tms_CN005064", nCS_Error)
            Call SP_OTBxSelText(otbNum)
            Exit Sub
        End If
        '*Eaw 56-09-25 CommSheet TK-ISS3000-358-FAPL-124_Eaw.xls
        If (cVB_FCSrcRteAmt <> ((cVB_FCSrcRteAmt \ cVB_MinDeno) * cVB_MinDeno)) Then
            Call SP_MSGnShowing("tms_CN005064", nCS_Error)
            Call SP_OTBxSelText(otbNum)
            Exit Sub
        End If
    End If
    '*Eaw 56-08-08 CommSheet TK-ISS3000-271-TR.xls
    If cVB_MaxChange <> 0 Then
        cChange = cPay - cVB_Balance
        If cChange > cVB_MaxChange Then
            Call SP_MSGnShowing("tms_CN005066", nCS_Error)
            Call SP_OTBxSelText(otbNum)
            Exit Sub
        End If
    End If
    cVB_Pay = cPay
    Unload Me
End Sub

Private Sub otbNum_Change()         '*AO51-03-12    Comm.Sheet SF- 070
    otbNum.Text = SP_DEFtChangeNum(Trim(otbNum.Text))
    otbNum.SelStart = Len(otbNum.Text)
End Sub

Private Sub otbNum_GotFocus()
    otbNum.SelStart = 0
    otbNum.SelLength = Len(otbNum.Text)
End Sub

Private Sub otbNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbNum.Text = ""
    End Select
End Sub

Private Sub otbNum_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case 13
            Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

