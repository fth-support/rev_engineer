VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNGSTReEntry 
   Caption         =   "GST Amount:;klg_Title"
   ClientHeight    =   2850
   ClientLeft      =   4500
   ClientTop       =   4830
   ClientWidth     =   6150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6150
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   ">60"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   0
      TabIndex        =   10
      Top             =   8400
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   "45-60"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   0
      TabIndex        =   9
      Top             =   7920
      Width           =   1455
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   "31-45"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   7440
      Width           =   1455
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   "21-30"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   6960
      Width           =   1455
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   "16-20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.OptionButton optAge 
      BackColor       =   &H00FFDED5&
      Caption         =   "<15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   6000
      Width           =   1455
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2280
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -250
      Width           =   6165
      _Version        =   720897
      _ExtentX        =   10874
      _ExtentY        =   4022
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GST Amount:;klg_Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   585
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   960
         Width           =   4725
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Width           =   6165
      _Version        =   720897
      _ExtentX        =   10874
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
         Left            =   2880
         TabIndex        =   1
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
         Left            =   4440
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNGSTReEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxInitail
End Sub

Private Sub W_FRMxInitail()
    otbAmount.Text = "0.00"
End Sub

Private Sub ocmCancel_Click()
    bVB_CNInputGSTReEntry = False
    bVB_CNTrnDisc = False '*Eaw 56-08-27
    '-------------------------------------------
    tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum           '*Ao 2014/09/05   V 4.2.143
    tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
    bMainFunc = True
    '-------------------------------------------
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim tSql As String
    cUT_VatAmt = Val(Format(otbAmount.Text))
    cVB_GSTVatAmt = Val(Format(otbAmount.Text)) '*Eaw 56-07-31CommSheet TK-ISS3000-283 TR FAPL-063.xls
    tSql = "UPDATE " & tVB_CNTblHD & " SET FCShdVat = " & SP_FMTcCurPoint(cUT_VatAmt, True, False)
    tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
    If SP_SQLvExecute(tSql) = 0 Then
        bVB_CNInputGSTReEntry = True
        bMainFunc = False                                   '*Ao 2014/09/05   V 4.2.143
        Unload Me
    End If
End Sub

Private Sub otbAmount_Change()
    otbAmount.Text = SP_DEFtChangeNum(Trim(otbAmount.Text))
    otbAmount.SelStart = Len(otbAmount.Text)
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
'            If InStr(Trim(otbForm.Text), ".") > 0 Then     '*AO51-03-12    Comm.Sheet SF- 070
'                KeyAscii = 0
'            End If
        Case 13
            Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbAmount_GotFocus()
    With otbAmount
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub



