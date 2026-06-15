VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNDiscLineItems 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amount;klg_Title"
   ClientHeight    =   2505
   ClientLeft      =   7215
   ClientTop       =   2775
   ClientWidth     =   4710
   ControlBox      =   0   'False
   Icon            =   "wCNDiscLineItems.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4710
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1815
      Index           =   2
      Left            =   0
      TabIndex        =   3
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
         Left            =   1920
         MaxLength       =   11
         TabIndex        =   0
         Text            =   "0"
         Top             =   600
         Width           =   2535
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
         Left            =   240
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   600
         Width           =   1170
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1300
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
Attribute VB_Name = "wCNDiscLineItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub W_FRMxSetInitial()
'    otbAmount.Text = SP_FMTtNum2Shw(0)
    With otbAmount
        .Text = 0
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    cVB_DiscLineItem = 0
    Unload Me
End Sub

Private Sub ocmOK_Click()
'*Eaw 56-07-04 CommSheet TK-ISS3000-201-ST.xls ไม่ต้องเช็คเหรียญ
'    If Not (nVB_TypeDiscLineItem = 1) Then          '*AO51-03-12  ลดแบบเปอร์เซ็นไม่ต้องเช็ค
'        If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbAmount.Text))) Then Exit Sub       '*AO51-03-26   Comm.Sheet  ML-SF-100
'    End If
    cVB_DiscLineItem = SP_VALcText2Double(otbAmount.Text)
    If bVB_LineDiscItems Then
        oDisc.bRecalByItem = False
        oDisc.DISxClearAll
        oDisc.dTranDate = tUT_SaleDate
        oDisc.tTerminalNo = tVB_CNTerminalNum
        oDisc.tTransNo = tVB_Running
        oDisc.cDiscValue = cVB_DiscLineItem
        oDisc.nDTSeqNo = xVB_SKULnDisc.nSeqNo
        oDisc.nENLnDisType = nVB_TypeDiscLineItem
        '*KT 53-10-30 Preview Day
        '---------------------------------------------------------------------------------------------------------------------------
        oDisc.bIsPreviewDay = IIf(tVB_CNSKUFlagPrv = "Y", True, False)
        If bVB_CNTrStaRegPreviewDay Then
            If tVB_CNSKUFlagPrv = "Y" Then
                Call SP_INIxSwitchModePreviewDay(True)
            Else
                Call SP_INIxSwitchModePreviewDay(False)
            End If
        End If
        '---------------------------------------------------------------------------------------------------------------------------
        Call oDisc.DIScDiscount(nEN_Disc05LnDisc, eEN_VerifyDiscType.nEN_Type1ByItem) '*KT 52-09-23 Analyse code
    Else
        oDisc.bRecalByItem = False
        oDisc.DISxClearAll
        oDisc.dTranDate = tUT_SaleDate
        oDisc.tTerminalNo = tVB_CNTerminalNum
        oDisc.tTransNo = tVB_Running
        oDisc.cDiscValue = cVB_DiscLineItem
        oDisc.nDTSeqNo = xVB_SKULnDisc.nSeqNo
        oDisc.nENLnDisType = nVB_TypeDiscLineItem
        '*KT 53-10-30 Preview Day
        '---------------------------------------------------------------------------------------------------------------------------
        oDisc.bIsPreviewDay = IIf(tVB_CNSKUFlagPrv = "Y", True, False)
        If bVB_CNTrStaRegPreviewDay Then
            If tVB_CNSKUFlagPrv = "Y" Then
                Call SP_INIxSwitchModePreviewDay(True)
            Else
                Call SP_INIxSwitchModePreviewDay(False)
            End If
        End If
        '---------------------------------------------------------------------------------------------------------------------------
        Call oDisc.DIScDiscount(nEN_Disc06LnMDown, eEN_VerifyDiscType.nEN_Type1ByItem) '*KT 52-09-23 Analyse code
    End If
    cVB_DiscLineItem = oDisc.cLastDiscount
    Unload Me
End Sub
Private Sub otbAmount_Change()         '*AO51-03-12    Comm.Sheet SF- 070
'    otbAmount.Text = SP_DEFtChangeNum(Trim(otbAmount.Text))
'*Eaw 56-08-06 CommSheet TK-ISS3000-306-UAT.xls เคส % off ไม่ปรับเป็นทศนิยมและคีย์ได้ไม่เกิน 3 ตัว
    Select Case nVB_TypeDiscLineItem
        Case nEN_LnDis1Per
            otbAmount.MaxLength = 3
        Case Else
            otbAmount.MaxLength = 11
            otbAmount.Text = SP_DEFtChangeNum(Trim(otbAmount.Text))
    End Select
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
'            If InStr(Trim(otbAmount.Text), ".") > 0 Then       '*AO51-03-12    Comm.Sheet SF- 070
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
