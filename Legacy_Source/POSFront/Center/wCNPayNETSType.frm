VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNPayNETSType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NET Type;klg_Title"
   ClientHeight    =   3900
   ClientLeft      =   2310
   ClientTop       =   3225
   ClientWidth     =   4665
   ControlBox      =   0   'False
   Icon            =   "wCNPayNETSType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4665
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1080
      Index           =   2
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3090
      Width           =   4680
      _Version        =   720897
      _ExtentX        =   8255
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
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
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   120
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
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   120
         Width           =   1455
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdFunc 
      Height          =   2510
      Left            =   0
      TabIndex        =   1
      Top             =   555
      Width           =   4680
      _cx             =   8255
      _cy             =   4427
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   11889258
      ForeColorFixed  =   0
      BackColorSel    =   11889258
      ForeColorSel    =   16777215
      BackColorBkg    =   16768725
      BackColorAlternate=   16768725
      GridColor       =   16777152
      GridColorFixed  =   16768725
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16768725
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNPayNETSType.frx":058A
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   840
      Index           =   0
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   4680
      _Version        =   720897
      _ExtentX        =   8255
      _ExtentY        =   1482
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NET Type;klg_Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   0
         Width           =   4215
      End
   End
End
Attribute VB_Name = "wCNPayNETSType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey As String
Private nMaxRow As Long

'NET                                    (Function Code 59)
'NET BCA                            (Function Code 65)
'NET Gift Card/Prepaid   (Function Code 70)
'NET Pre-paid                       (Function Code 58)

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
    Dim bFound As Boolean
    Dim nI As Long
On Error GoTo ErrHandle
    With ogdFunc
        bFound = False
         For nI = .FixedRows To .rows - 1
                If .TextMatrix(nI, 0) <> "" Then
                    If Val(Format(.TextMatrix(nI, 0))) = pnFuc Then
                        .Row = nI
                        bFound = True
                        Exit For
                    End If
                End If
         Next nI
         If Not bFound Then Exit Sub
         .Select .Row, 0, .Row, .Cols - 1
         .SetFocus
        .TopRow = .Row
    End With
    Exit Sub
ErrHandle:
    Exit Sub
End Sub

Private Sub W_FRMxSetInitial()
Dim tTmp  As String
    
    cVB_Pay = 0
    tVB_NETType = ""
    tTmp = SP_INItGetValue(tVB_CNPthFileINI, Me.Name, "klg_NET1")
    If tTmp <> "" Then
        With ogdFunc
            '*Tao 62-08-29 Seq
            .TextMatrix(0, 0) = "1"
            .TextMatrix(1, 0) = "2"
            .TextMatrix(2, 0) = "3"
            .TextMatrix(3, 0) = "4"
            
            '*Tao 62-08-29 Name
            .TextMatrix(0, 1) = "1." & cUT.UT_STRtToken("2;" & SP_INItGetValue(tVB_CNPthFileINI, Me.Name, "klg_NET1"), ";", nVB_CNCutLang + 1)
            .TextMatrix(1, 1) = "2." & cUT.UT_STRtToken("2;" & SP_INItGetValue(tVB_CNPthFileINI, Me.Name, "klg_NET2"), ";", nVB_CNCutLang + 1)
            .TextMatrix(2, 1) = "3." & cUT.UT_STRtToken("2;" & SP_INItGetValue(tVB_CNPthFileINI, Me.Name, "klg_NET3"), ";", nVB_CNCutLang + 1)
'            .TextMatrix(3, 1) = "4." & cUT.UT_STRtToken("2;" & SP_INItGetValue(tVB_CNPthFileINI, Me.Name, "klg_NET4"), ";", nVB_CNCutLang + 1)  '*Eaw 56-05-16 CommSheet TK-ISS3000-180-ST 00 00 01.xls
            .TextMatrix(3, 1) = "4." & cUT.UT_STRtToken("2;" & SP_INItGetValue(tVB_CNPthFileINI, Me.Name, "klg_NET4"), ";", nVB_CNCutLang + 1) '*Tao 62-07-19 NET QR
        End With
    End If

End Sub

'******************************************************************************

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
    Exit Sub
ErrHandle:

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call ocmOK_Click
    Else
        Select Case KeyAscii
            Case 48 To 58
                With ogdFunc
                    tW_HotKey = tW_HotKey & Chr(KeyAscii)
                    If tW_HotKey <> "" Then
                         W_PRCxCallFunc (Val(tW_HotKey))
                    End If
                    If IsNumeric(tW_HotKey) Then
                        If CLng(tW_HotKey) >= nMaxRow Then tW_HotKey = "" ': otmForm.Enabled = False
                    End If
                End With
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    cVB_Pay = 0
    tVB_NETType = ""
    Unload Me
End Sub

Private Sub ocmOK_Click()
    With ogdFunc
        If .TextMatrix(.Row, 1) = "" Then Exit Sub
'        tVB_NETType = .TextMatrix(.Row, 1)
        tVB_CrdType = .TextMatrix(.Row, 1)                                        '*Ao 2013/07/10
        nVB_CNNETSType = .Row  '*Eaw 56-04-29 NETS Port  เก็บไปแยกประเภทย่อยของ NETS
    End With
    Call SP_CRTxLogPrg("Payment = " & tVB_CrdType)
    Unload Me
'    wCNPayNETS.Show vbModal
    If SP_NETbCheckPort Then '*Eaw 56-06-17
'        wCNPayCreditPay.Show vbModal
        '*Eaw 56-06-11
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc, tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnCancelVoucher, tEN_TrnOtherDisbursement  '*Eaw 56-11-26 เพิ่มเคสจากโหมดปกติเข้าไป CommSheet TK-ISS3000-363-FAPL-129.xls
               wCNPayNETS.Show vbModal
            Case Else
                wCNPayCreditPay.Show vbModal
        End Select
    Else
        wCNPayNETS.Show vbModal
    End If
End Sub

Private Sub ogdFunc_DblClick()
    Call ocmOK_Click
End Sub
