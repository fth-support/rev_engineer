VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNWalletReturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wallet Refund"
   ClientHeight    =   3660
   ClientLeft      =   8265
   ClientTop       =   9105
   ClientWidth     =   5325
   ControlBox      =   0   'False
   Icon            =   "wCNWalletReturn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5325
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2955
      Index           =   2
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   -240
      Width           =   5325
      _Version        =   720897
      _ExtentX        =   9393
      _ExtentY        =   5212
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbPartnerID 
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
         Left            =   2220
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1380
         Width           =   2655
      End
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
         Left            =   2205
         MaxLength       =   11
         TabIndex        =   2
         Text            =   "0"
         Top             =   2100
         Width           =   2655
      End
      Begin XtremeSuiteControls.ComboBox ocbWalletType 
         Height          =   495
         Left            =   2235
         TabIndex        =   0
         Top             =   630
         Width           =   2655
         _Version        =   720897
         _ExtentX        =   4683
         _ExtentY        =   873
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wallet Type :"
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
         TabIndex        =   9
         Top             =   690
         Width           =   1665
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WPS Trans ID :"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1470
         Width           =   1950
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
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1170
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2460
      Width           =   5325
      _Version        =   720897
      _ExtentX        =   9393
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
         Left            =   2040
         TabIndex        =   3
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
         Left            =   3600
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNWalletReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*Tao 60-09-01 Wallet Return

Private Sub Form_Activate()
Call SP_OBJxSetFocus(ocbWalletType)
End Sub

Private Sub Form_Load()
    Dim nIndex As Integer
    Call SP_FRMxChildCenter(Me)
    For nIndex = 0 To UBound(aVB_WalletType) - 1
        ocbWalletType.AddItem (aVB_WalletType(nIndex).TenderType)
    Next nIndex
    Select Case tVB_TenderCode
    Case "T019"
        ocbWalletType.ListIndex = 0
    Case "T020"
        ocbWalletType.ListIndex = 1
    Case "T021"
        ocbWalletType.ListIndex = 2
    Case "T022"
        ocbWalletType.ListIndex = 3
    End Select
    otbPartnerID.Text = aVB_WalletLog.PartnerTransID
    otbAmount.Text = SP_FMTcCurPoint(aVB_WalletLog.ReqAmt, True, True)
    If bVB_Reference Then '*Tao 60-09-04 Return refer Sale ไม่ควรให้แก้ไข
        ocbWalletType.Enabled = False
        If (otbPartnerID.Text = "") Then
            otbPartnerID.Enabled = True
        Else
            otbPartnerID.Enabled = False
        End If
    Else
        ocbWalletType.Enabled = True
        otbPartnerID.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocbWalletType_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
    Call SP_OBJxSetFocus(otbPartnerID)
Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown 'ให้กด ลูกศร ได้
    '
Case Else
    KeyCode = 0
End Select
End Sub

Private Sub ocmCancel_Click()
    bVB_CNReturnOk = False
    Unload Me
End Sub

Private Sub ocmOk_Click()
    '*Tao 61-08-02 ไม่ใส่ก็ได้ (ตาม Requirement)
'    If Trim(otbPartnerID.Text) = "" Then '่ไม่ใส่
'        Call SP_MSGnShowing("tms_CN999042", nCS_Warning)
'        Call SP_OBJxSetFocus(otbPartnerID)
'        Exit Sub
'    End If
    If CDbl(Val(Format(otbAmount.Text))) > cVB_Balance Then
        Call SP_MSGnShowing("tms_CN005064", nCS_Warning)
        otbAmount.Text = SP_FMTcCurPoint(aVB_WalletLog.ReqAmt, True, True)
        otbAmount.SelStart = 0
        otbAmount.SelLength = Len(otbAmount.Text)
        Call SP_OBJxSetFocus(otbAmount)   ' otbAmount.SetFocus
        Exit Sub
    End If
    If bVB_Return And bVB_Reference Then '*Tao 60-09-04 Return refer Sale ยอด Return ต้องไม่มากกว่ายอด Sale
        If CDbl(Val(Format(otbAmount.Text))) > aVB_WalletLog.ReqAmt Then
            Call SP_MSGnShowing("tms_CN005064", nCS_Warning)
            otbAmount.Text = SP_FMTcCurPoint(aVB_WalletLog.ReqAmt, True, True)
            otbAmount.SelStart = 0
            otbAmount.SelLength = Len(otbAmount.Text)
            Call SP_OBJxSetFocus(otbAmount)   ' otbAmount.SetFocus
            Exit Sub
        End If
    End If
    Select Case ocbWalletType.ListIndex
    Case 0
        If Not aVB_WalletType(0).StaReturn Then
            Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(0).TenderType) 'ไม่อนุญาตคืนสินค้าด้วย Alipay
            Call SP_OBJxSetFocus(ocbWalletType)
            Exit Sub
        End If
    Case 1
        If Not aVB_WalletType(1).StaReturn Then
            Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(1).TenderType) 'ไม่อนุญาตคืนสินค้าด้วย WeChatpay
            Call SP_OBJxSetFocus(ocbWalletType)
            Exit Sub
        End If
    Case 2
        If Not aVB_WalletType(2).StaReturn Then
            Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(2).TenderType) 'ไม่อนุญาตคืนสินค้าด้วย Bluepay
            Call SP_OBJxSetFocus(ocbWalletType)
            Exit Sub
        End If
    Case 3
        If Not aVB_WalletType(3).StaReturn Then
            Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(3).TenderType) 'ไม่อนุญาตคืนสินค้าด้วย Others
            Call SP_OBJxSetFocus(ocbWalletType)
            Exit Sub
        End If
    Case Else
        Call SP_MSGnShowing("tms_CN999049", nCS_Warning) 'กรุณาเลือก Wallet Type
        Call SP_OBJxSetFocus(ocbWalletType)
        Exit Sub
    End Select
    wFunctionKB.tW_TransID = Trim(otbPartnerID.Text)
    wFunctionKB.tW_WalletName = Trim(ocbWalletType.Text)
    cVB_CNDiscAmt = SP_VALcText2Double(otbAmount.Text)
    bVB_CNReturnOk = True
    Unload Me
End Sub

Private Sub otbAmount_Change()
    otbAmount.Text = SP_DEFtChangeNum(Trim(otbAmount.Text))
    otbAmount.SelStart = Len(otbAmount.Text)
End Sub

Private Sub otbAmount_GotFocus()
    otbAmount.SelStart = 0
    otbAmount.SelLength = Len(otbAmount.Text)
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
            KeyAscii = 0
        Case 13
            Call ocmOk_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbPartnerID_Change()
    otbPartnerID.SelStart = Len(otbPartnerID.Text)
End Sub

Private Sub otbPartnerID_GotFocus()
    otbPartnerID.SelStart = 0
    otbPartnerID.SelLength = Len(otbPartnerID.Text)
End Sub

Private Sub otbPartnerID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbPartnerID.Text = ""
    End Select
End Sub

Private Sub otbPartnerID_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case 13
            Call SP_OBJxSetFocus(otbAmount)
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub
