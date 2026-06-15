VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrative Menu;klg_Title"
   ClientHeight    =   3795
   ClientLeft      =   5505
   ClientTop       =   2760
   ClientWidth     =   5865
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "wAdmin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5865
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   0
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2850
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
      Height          =   3135
      Index           =   2
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   5865
      _Version        =   720897
      _ExtentX        =   10345
      _ExtentY        =   5530
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VSFlex7Ctl.VSFlexGrid ogdForm 
         Height          =   2535
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   5535
         _cx             =   9763
         _cy             =   4471
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
         Rows            =   5
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"wAdmin.frx":058A
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
Attribute VB_Name = "wAdmin"
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
    Call W_FNCxAddFunc2Grid
    Call SP_STAxMouse(Me, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Select Case ogdForm.Row
        Case 0
            nVB_RTD = 1
            wCNAllow.tW_ID = tVB_CNMnuCurrent
            wCNAllow.Show vbModal
'            DoEvents '*KT 53-01-14 DoEvents
            If tVB_CNUserAlwC <> "" Then
                wCNMoveFile.Caption = ogdForm.TextMatrix(ogdForm.Row, 1)
                '*Eaw 56-11-14
                wCNMoveFile.otbForm(0).Enabled = False
                wCNMoveFile.ocmBrw(0).Enabled = False
                wCNMoveFile.otbForm(1).Enabled = True
                wCNMoveFile.ocmBrw(1).Enabled = True
                wCNMoveFile.Show vbModal
            End If
        Case 1
            nVB_RTD = 2
            wCNAllow.tW_ID = tVB_CNMnuCurrent
            wCNAllow.Show vbModal
'            DoEvents '*KT 53-01-14 DoEvents
            If tVB_CNUserAlwC <> "" Then
                wCNMoveFile.Caption = ogdForm.TextMatrix(ogdForm.Row, 1)
                '*Eaw 56-11-14
                wCNMoveFile.otbForm(0).Enabled = True
                wCNMoveFile.ocmBrw(0).Enabled = True
                wCNMoveFile.otbForm(1).Enabled = False
                wCNMoveFile.ocmBrw(1).Enabled = False
                wCNMoveFile.Show vbModal
            End If
        Case 2
            nVB_RTD = 3
            wCNAllow.tW_ID = tVB_CNMnuCurrent
            wCNAllow.Show vbModal
'            DoEvents '*KT 53-01-14 DoEvents
            If tVB_CNUserAlwC <> "" Then
                wCNMoveFile.Caption = ogdForm.TextMatrix(ogdForm.Row, 1)
                wCNMoveFile.otbForm(1).Enabled = False '*Eaw 56-10-18
                wCNMoveFile.ocmBrw(1).Enabled = False '*Eaw 56-10-18
                wCNMoveFile.Show vbModal
            End If
        Case 3 'Change Operation Date
            'Hold
        Case 4 'Change Transaction Number
            wCNAllow.tW_ID = tVB_CNMnuCurrent
            wCNAllow.Show vbModal
'            DoEvents '*KT 53-01-14 DoEvents
            If tVB_CNUserAlwC <> "" Then
                wCNChgTrnNo.Show vbModal
            End If
    End Select
    Unload Me
End Sub

Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_HotKey = ""
End Sub

Private Sub ogdForm_DblClick()
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

Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call:
'Cmt: ทำการแสดงรายการของ Admin Menu
'-------------------------------------------------------
Dim nRow%
Dim tSmnID$
With ogdForm
    .rows = 5
    .TextMatrix(0, 1) = Trim(SP_GETtVocabularyTicketFromArray("160401", aVB_MainMnuAll, tSmnID)) '"Copy Transaction To Device"
    .TextMatrix(1, 1) = Trim(SP_GETtVocabularyTicketFromArray("160402", aVB_MainMnuAll, tSmnID)) '"Copy Transaction From Device"
    .TextMatrix(2, 1) = Trim(SP_GETtVocabularyTicketFromArray("160403", aVB_MainMnuAll, tSmnID)) '"Update Master"
    .TextMatrix(3, 1) = Trim(SP_GETtVocabularyTicketFromArray("160404", aVB_MainMnuAll, tSmnID)) '"Change Operation Date"
    .TextMatrix(4, 1) = Trim(SP_GETtVocabularyTicketFromArray("160405", aVB_MainMnuAll, tSmnID)) '"Change Terminal Number"
    For nRow = 0 To 4
            .TextMatrix(nRow, 1) = " " & nRow + 1 & "." & .TextMatrix(nRow, 1)
            .TextMatrix(nRow, 0) = nRow + 1
    Next nRow
End With
End Sub
