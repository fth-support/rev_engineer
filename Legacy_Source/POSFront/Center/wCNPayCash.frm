VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNPayCash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash;klg_Title"
   ClientHeight    =   2130
   ClientLeft      =   8580
   ClientTop       =   4335
   ClientWidth     =   4725
   ControlBox      =   0   'False
   Icon            =   "wCNPayCash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4725
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1455
      Index           =   2
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   4725
      _Version        =   720897
      _ExtentX        =   8334
      _ExtentY        =   2566
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbAmount 
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
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :;klg_Amount"
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
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   600
         Width           =   2760
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   930
      Width           =   4725
      _Version        =   720897
      _ExtentX        =   8334
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
         Left            =   3000
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
         Left            =   1440
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNPayCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Form_Activate()
'    Call SP_OTBxSelText(otbAmount)
'End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    wCNPayCash.Caption = tVB_TenderName
    otbAmount.MaxLength = 13  '*Eaw 56-03-02  TK-ISS3000-043 กรณี Auto fill balance
    otbAmount.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)    '*Eaw 56-02-02
    cVB_Pay = 0 '*Tao 58-12-14 เคลียร์ค่าการชำระเงิน
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    cVB_Pay = 0
    If tVB_TenderCode <> "T019" Then cVB_Balance = cVB_Balance - cVB_RoundRcv    '*Eaw 56-07-04
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim cPay As Double
On Error GoTo ErrHandle
    'If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbAmount.Text))) Then Exit Sub       '*AO51-03-26   Comm.Sheet  ML-SF-100  '*Eaw 56-05-27 CommSheet TK-ISS3000-201-ST 00 00 01.xls
    If (tVB_TenderCode = "T001") Then If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbAmount.Text))) Then Exit Sub '*Tao 60-08-28 เช็คยอดเงิน ให้สามารถ ทอนเงินได้
    cPay = SP_VALcText2Double(otbAmount.Text)
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
'             If CDbl(Val(Format(otbAmount.Text))) > cVB_Balance Then
             If SP_VALcText2Double(CDbl(Val(Format(otbAmount.Text)))) > SP_VALcText2Double(cVB_Balance) Then   '*Eaw 56-03-13
                Call SP_MSGnShowing("tms_CN005064", nCS_Warning)
                Call SP_OBJxSetFocus(otbAmount)   ' otbAmount.SetFocus
                Exit Sub
            End If
            cVB_Pay = cPay
        '*KT 52-09-19 PH2-ML-SF-144 Cancel Voucher ต้องตรวจสอบ Tender ที่จ่ายจะต้องเท่ากับ Voucher Amount ปัจจุบันสามารถจ่ายได้มากกว่า Voucher Amount ได้และมีการทอนเงิน ซึ่งผิด และจะทำให้ข้อมูล Upload file ผิด(ต้องตรวจสอบเช่นเดียวกับ Return )
        Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnCancelVoucher
'           If CDbl(Val(Format(otbAmount.Text))) > cVB_Balance Then
            If SP_VALcText2Double(CDbl(Val(Format(otbAmount.Text)))) > SP_VALcText2Double(cVB_Balance) Then   '*Eaw 56-03-13
                Call SP_MSGnShowing("tms_CN005064", nCS_Warning)
                Call SP_OBJxSetFocus(otbAmount)   ' otbAmount.SetFocus
                Exit Sub
            End If
            cVB_Pay = cPay
        '*KT 52-12-09 PH3.0-ML-SF-42
        Case tEN_TrnOtherDisbursement
'            If CDbl(Val(Format(otbAmount.Text))) > cVB_Balance Then
            If SP_VALcText2Double(CDbl(Val(Format(otbAmount.Text)))) > SP_VALcText2Double(cVB_Balance) Then   '*Eaw 56-03-13
                Call SP_MSGnShowing("tms_CN005064", nCS_Warning)
                Call SP_OBJxSetFocus(otbAmount)   ' otbAmount.SetFocus
                Exit Sub
            End If
            cVB_Pay = cPay
        Case Else
            If cPay = 0 And cVB_Balance <> 0 Then
                'MsgBox "Bad Amount"
                Call SP_MSGnShowing("tms_CN005064", nCS_Warning)
                Call SP_OBJxSetFocus(otbAmount)   ' otbAmount.SetFocus
                Exit Sub
            End If
            cVB_Pay = cPay
            Call SP_CRTxLogCashBalEvent(Me.Name, "ocmOK_Click[1]", "Cash Tender", cVB_Balance)
            Call SP_CRTxLogCashPayEvent(Me.Name, "ocmOK_Click[1]", "Cash Tender", cVB_Pay)
    End Select
    Unload Me
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_OBJxSetFocus(otbAmount)   ' otbAmount.SetFocus
End Sub

Private Sub otbAmount_Change()      '*AO51-03-12    Comm.Sheet SF- 070
    otbAmount.Text = SP_DEFtChangeNum(Trim(otbAmount.Text))
    otbAmount.SelStart = Len(otbAmount.Text)
End Sub

'Private Sub otbAmount_Change()             '*AO51-03-12   Comm.Sheet SF-070
'    If otbAmount.Text <> "" Then
'        If InStr(1, otbAmount.Text, ".") > 0 Then
'            otbAmount.Text = Format(SP_VALcText2Double(otbAmount.Text), "#,##0." & String(nVB_CNDecAmtForShw, "#"))
'            otbAmount.SelStart = Len(otbAmount.Text)
'        Else
'            otbAmount.Text = Format(SP_VALcText2Double(otbAmount.Text), "#,##0")
'            otbAmount.SelStart = Len(otbAmount.Text)
'        End If
'    End If
'End Sub
Private Sub otbAmount_GotFocus()
    With otbAmount
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub otbAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbAmount.Text = ""
    End Select
End Sub

Private Sub otbAmount_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0                    '*AO51-03-12   Comm.Sheet SF-070
'            If InStr(Trim(otbAmount.Text), ".") > 0 Then
'                KeyAscii = 0
'            End If
        Case 13
            Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            Else
                otbAmount.MaxLength = 11  '*Eaw 56-03-02 TK-ISS3000-063 ถ้าคีย์มูลค่าเองให้คีย์ได้ 8+2
            End If
    End Select
End Sub
