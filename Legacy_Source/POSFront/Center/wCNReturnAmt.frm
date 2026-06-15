VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNReturnAmt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amount"
   ClientHeight    =   2520
   ClientLeft      =   8265
   ClientTop       =   9105
   ClientWidth     =   4740
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4740
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1815
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   4725
      _Version        =   720897
      _ExtentX        =   8334
      _ExtentY        =   3201
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
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   0
         Text            =   "0"
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :"
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
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1170
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4725
      _Version        =   720897
      _ExtentX        =   8334
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmOk 
         Caption         =   "&Ok"
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
      Begin VB.CommandButton ocmCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
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
   End
End
Attribute VB_Name = "wCNReturnAmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bW_IsFreeItem As Boolean

Private Sub Form_Load()
'*KT 53-01-26 Communication Sheet PH3-ML-SB-073 Deal & Free  in Return Transaction.xls
    Call SP_FRMxChildCenter(Me)
    If bW_IsFreeItem Then
        'Free Item
        otbAmount.Text = SP_FMTcCurPoint(cVB_CNRetSaleAmt, True, True)
        otbAmount.Locked = True
        otbAmount.BackColor = nVB_CNClrBLock
    Else
        'other
        otbAmount.Text = ""
        otbAmount.Locked = False
        otbAmount.BackColor = nVB_CNClrBUnLock
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    bVB_CNReturnOk = False
    Unload Me
End Sub

Private Sub ocmOk_Click()
    cVB_CNDiscAmt = SP_VALcText2Double(otbAmount.Text)
    bVB_CNReturnOk = True
    Unload Me
End Sub

Private Sub otbAmount_Change()
    otbAmount.Text = SP_DEFtChangeNum(Trim(otbAmount.Text))
    otbAmount.SelStart = Len(otbAmount.Text)
End Sub

Private Sub otbAmount_GotFocus()
    If Not bW_IsFreeItem Then
        otbAmount.SelStart = 0
        otbAmount.SelLength = Len(otbAmount.Text)
    End If
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
        Case 46             '·È¹ÔÂÁ
            KeyAscii = 0
        Case 13
            Call ocmOk_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub


