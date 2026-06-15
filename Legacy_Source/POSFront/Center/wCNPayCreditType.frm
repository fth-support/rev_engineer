VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNPayCreditType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Card;klg_Title"
   ClientHeight    =   2430
   ClientLeft      =   4245
   ClientTop       =   4200
   ClientWidth     =   5865
   ControlBox      =   0   'False
   Icon            =   "wCNPayCreditType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5865
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1515
      Width           =   5865
      _Version        =   720897
      _ExtentX        =   10345
      _ExtentY        =   2117
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
         Left            =   4200
         TabIndex        =   2
         TabStop         =   0   'False
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
         Left            =   2640
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   240
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1815
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   5865
      _Version        =   720897
      _ExtentX        =   10345
      _ExtentY        =   3201
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VSFlex7Ctl.VSFlexGrid ogdForm 
         Height          =   1095
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   5535
         _cx             =   9763
         _cy             =   1931
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   11889258
         ForeColorSel    =   16777215
         BackColorBkg    =   16768725
         BackColorAlternate=   16768725
         GridColor       =   16768725
         GridColorFixed  =   16768725
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16768725
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"wCNPayCreditType.frx":058A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
         WallPaperAlignment=   9
      End
   End
End
Attribute VB_Name = "wCNPayCreditType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim nMaxRow As Long
    nMaxRow = ogdForm.Aggregate(flexSTMax, 0, 0, ogdForm.rows - 1, 0)
    If KeyAscii = 13 Then
        Call ocmOK_Click
    Else
        Select Case KeyAscii
            Case 48 To 58
                With ogdForm
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

Private Sub Form_Load()
    tW_HotKey = ""
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_SETxGridFunc
End Sub

Private Sub W_SETxGridFunc()
'----------------------------------------------------
'Cmt: ทำการ set label ตาม System Maintenance/Tab Tender Type
          '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
'----------------------------------------------------
    With ogdForm
        .TextMatrix(0, 1) = "1." & aVB_TenderMnu(1).Menu
        .TextMatrix(1, 1) = "2." & aVB_TenderMnu(2).Menu
    End With
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = 1
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    cVB_Pay = 0
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim nI As Long
    Select Case ogdForm.Row
        Case 0
            For nI = 0 To UBound(aVB_TenderMnu)
                If UCase(aVB_TenderMnu(nI).ID) = "T002" Then
                    nVB_CurTenderSeqNo = nI
                    Exit For
                End If
            Next nI
            tVB_CNMnuCurrent = "1817"
            If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
            End If
            tVB_TenderCode = "T002"
            tVB_TenderType = aVB_TenderMnu(nVB_CurTenderSeqNo).Name '"Credit Card" '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
            wCNPayCredit.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu '"Credit Card" '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
            tVB_TenderName = aVB_TenderMnu(nVB_CurTenderSeqNo).DescSale
            wCNPayCreditPay.Show vbModal        '*Ao 2013/01/03
'            wCNPayCredit.Show vbModal
        Case 1
            For nI = 0 To UBound(aVB_TenderMnu)
                If UCase(aVB_TenderMnu(nI).ID) = "T003" Then
                    nVB_CurTenderSeqNo = nI
                    Exit For
                End If
            Next nI
            tVB_CNMnuCurrent = "1818"
            If aVB_TenderMnu(nVB_CurTenderSeqNo).Autho Then
                If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
            End If
            tVB_TenderCode = "T003"
            tVB_TenderType = aVB_TenderMnu(nVB_CurTenderSeqNo).Name '"Installment Credit" '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
            wCNPayCredit.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu  '"Installment Credit" '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
            tVB_TenderName = aVB_TenderMnu(nVB_CurTenderSeqNo).DescSale
            wCNPayCreditPay.Show vbModal        '*Ao 2013/01/03
'            wCNPayCredit.Show vbModal
    End Select
    Unload Me
End Sub

Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_HotKey = ""
End Sub

Private Sub ogdForm_Click()
    Call ocmOK_Click
End Sub


Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
    Dim bFound As Boolean
    Dim nI As Long
On Error GoTo ErrHandle
    With ogdForm
        bFound = False
         For nI = .FixedRows To .rows - 1
                If .TextMatrix(nI, 0) <> "" Then
                    If CLng(Val(Format(.TextMatrix(nI, 0)))) = pnFuc Then
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

