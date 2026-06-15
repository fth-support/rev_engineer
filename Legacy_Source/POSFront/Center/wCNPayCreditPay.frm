VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNPayCreditPay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Card;klg_Title"
   ClientHeight    =   2205
   ClientLeft      =   11595
   ClientTop       =   5940
   ClientWidth     =   4725
   ControlBox      =   0   'False
   Icon            =   "wCNPayCreditPay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4725
   Begin XtremeSuiteControls.TabControl otcForm 
      Height          =   1560
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   4725
      _Version        =   720897
      _ExtentX        =   8334
      _ExtentY        =   2752
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
         TabIndex        =   4
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
         Left            =   240
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   675
         Width           =   2760
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1035
      Width           =   4725
      _Version        =   720897
      _ExtentX        =   8334
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
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
         Left            =   3000
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNPayCreditPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub W_FRMxSetInitial()
'    cVB_Pay = 0                              '*Eaw 56-02-02
'    cVB_PayCredit = 0
'    cVB_Pay = cVB_Balance                  '*Ao 2013/02/07
    cVB_PayCredit = cVB_Balance
    otbNum.MaxLength = 13    '*Eaw 56-03-02  TK-ISS3000-043 กรณี Auto fill balance
    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
End Sub

'******************************************************************************
Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
    cVB_Pay = 0 '*Tao 58-12-14 เคลียร์ค่าการชำระเงิน
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    cVB_Pay = 0
    cVB_PayCredit = 0
    Unload Me
End Sub

Private Sub ocmOK_Click()
    If SP_VALcText2Double(otbNum.Text) <= 0 Then
        Exit Sub
    End If
    '*Eaw 56-01-29
'    If SP_VALcText2Double(otbNum.Text) > cVB_Balance Then
'        Exit Sub
'    End If
    If SP_VALcText2Double(otbNum.Text) > SP_VALcText2Double(cVB_Balance) Then
        Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------------------------------
    cVB_PayCredit = SP_VALcText2Double(otbNum.Text)
    Unload Me
'    wCNPayCredit.Show vbModal  '*Eaw 56-04-29 NETS Port
    Select Case UCase(tVB_TenderCode)    '*Eaw 56-04-29 NETS Port
        Case "T002", "T003"
            wCNPayCredit.Show vbModal
        Case "T013", "T015", "T016"
            wCNPayNETS.Show vbModal
        Case Else
            'none
    End Select
End Sub

Private Sub otbNum_Change()
    otbNum.Text = SP_DEFtChangeNum(Trim(otbNum.Text))
    otbNum.SelStart = Len(otbNum.Text)
End Sub

Private Sub otbNum_GotFocus()
    With otbNum
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
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
            Else
                otbNum.MaxLength = 11 '*Eaw 56-03-02  TK-ISS3000-043 กรณีให้คีย์มูลค่าเองมีขนาด 8+2
            End If
    End Select
End Sub
