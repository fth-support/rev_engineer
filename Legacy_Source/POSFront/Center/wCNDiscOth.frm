VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNDiscOth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amount"
   ClientHeight    =   2505
   ClientLeft      =   10830
   ClientTop       =   8010
   ClientWidth     =   4755
   ControlBox      =   0   'False
   Icon            =   "wCNDiscOth.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4755
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1815
      Index           =   2
      Left            =   15
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
      Height          =   1230
      Index           =   1
      Left            =   15
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1305
      Width           =   4725
      _Version        =   720897
      _ExtentX        =   8334
      _ExtentY        =   2170
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel(Esc);klg_Cancel"
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
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNDiscOth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    bVB_DiscAccept = False
    
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    'Call SP_FRMxSetCapForm(Me)
    ocmOk.Caption = cUT.UT_STRtToken(ocmOk.Tag, ";", nVB_CNCutLang + 1)
    ocmCancel.Caption = cUT.UT_STRtToken(ocmCancel.Tag, ";", nVB_CNCutLang + 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    bVB_DiscAccept = False
    Unload Me
End Sub

Private Sub ocmOK_Click()

'*Eaw 56-07-04 CommSheet TK-ISS3000-201-ST.xls ไม่ต้องเช็คเหรียญ
'    If Not (nVB_TypeDiscLineItem = 1) Then                  '*AO51-03-12  ลดแบบเปอร์เซ็นไม่ต้องเช็ค
'        If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbAmount.Text))) Then Exit Sub       '*AO51-03-26   Comm.Sheet  ML-SF-100
'    End If
    
    cVB_DiscTransAmt = SP_VALcText2Double(otbAmount.Text)
    Select Case tVB_CNTrnDiscType
        Case "1814" 'Transaction Discount
            oDisc.bRecalByItem = False
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.cDiscValue = cVB_DiscTransAmt
            oDisc.nENLnDisType = nVB_TypeDiscLineItem
            Call oDisc.DIScDiscount(nEN_Disc13TranDisc, eEN_VerifyDiscType.nEN_Type2ByDcn)
            bVB_DiscAccept = True
        Case "1815" 'Transaction Markdown
            oDisc.bRecalByItem = False
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.cDiscValue = cVB_DiscTransAmt
            oDisc.nENLnDisType = nVB_TypeDiscLineItem
            Call oDisc.DIScDiscount(nEN_Disc14TranMDown, eEN_VerifyDiscType.nEN_Type2ByDcn)
            bVB_DiscAccept = True
        Case "1893" 'Trnsaction Markdown1
            oDisc.bRecalByItem = False
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.cDiscValue = cVB_DiscTransAmt
            oDisc.nENLnDisType = nVB_TypeDiscLineItem
            oDisc.tCardID = tVB_CNTranMDownNo
            Call oDisc.DIScDiscount(nEN_Disc23TranMDown1, eEN_VerifyDiscType.nEN_Type2ByDcn)
            bVB_DiscAccept = True
        Case "1892" 'Trnsaction Markdown2
            oDisc.bRecalByItem = False
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.cDiscValue = cVB_DiscTransAmt
            oDisc.nENLnDisType = nVB_TypeDiscLineItem
            oDisc.tCardID = tVB_CNTranMDownNo
            Call oDisc.DIScDiscount(nEN_Disc24TranMDown2, eEN_VerifyDiscType.nEN_Type2ByDcn)
            bVB_DiscAccept = True
        Case "1891" 'Trnsaction Markdown3
            oDisc.bRecalByItem = False
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.cDiscValue = cVB_DiscTransAmt
            oDisc.nENLnDisType = nVB_TypeDiscLineItem
            oDisc.tCardID = tVB_CNTranMDownNo
            Call oDisc.DIScDiscount(nEN_Disc25TranMDown3, eEN_VerifyDiscType.nEN_Type2ByDcn)
            bVB_DiscAccept = True
        Case "1890" 'Trnsaction Markdown4
            oDisc.bRecalByItem = False
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.cDiscValue = cVB_DiscTransAmt
            oDisc.nENLnDisType = nVB_TypeDiscLineItem
            oDisc.tCardID = tVB_CNTranMDownNo
            Call oDisc.DIScDiscount(nEN_Disc26TranMDown4, eEN_VerifyDiscType.nEN_Type2ByDcn)
            bVB_DiscAccept = True
        Case "1889" 'Trnsaction Markdown5
            oDisc.bRecalByItem = False
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.cDiscValue = cVB_DiscTransAmt
            oDisc.nENLnDisType = nVB_TypeDiscLineItem
            oDisc.tCardID = tVB_CNTranMDownNo
            Call oDisc.DIScDiscount(nEN_Disc27TranMDown5, eEN_VerifyDiscType.nEN_Type2ByDcn)
            bVB_DiscAccept = True
    End Select
    Unload Me
End Sub

Private Sub otbAmount_Change()      '*AO51-03-12    Comm.Sheet SF- 070
    otbAmount.Text = SP_DEFtChangeNum(Trim(otbAmount.Text))
    otbAmount.SelStart = Len(otbAmount.Text)
End Sub

'Private Sub otbAmount_Change()           '*AO51-03-12    Comm.Sheet SF- 070
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
'            If InStr(Trim(otbAmount.Text), ".") > 0 Then         '*AO51-03-12    Comm.Sheet SF- 070
'                KeyAscii = 0
'            End If
        Case 13
            KeyAscii = 0
            Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select

End Sub

