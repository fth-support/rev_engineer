VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNTDisc2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Discount;klg_Title"
   ClientHeight    =   2505
   ClientLeft      =   6315
   ClientTop       =   8010
   ClientWidth     =   4395
   ControlBox      =   0   'False
   Icon            =   "wCNTDisc2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4395
   Tag             =   "2;"
   Begin VSFlex7Ctl.VSFlexGrid ogdForm 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _cx             =   7858
      _cy             =   2778
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
      Rows            =   3
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"wCNTDisc2.frx":058A
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
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4400
      _Version        =   720897
      _ExtentX        =   7761
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
         Left            =   2760
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
         Left            =   1200
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNTDisc2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey As String
Private nMaxRow As Long
'Private bFirstLoad As Boolean

Private Sub W_DEFxChangeTitle()      '*AO51-03-13 Comm.Sheet  SF- 074
'_________________________
'   Call:
'   Cmt:    เปลื่ยน Title
'*KT 54-12-02 V.4.0.9 ML-PH4-R3-014 สามารถเลือก Transaction Markdown 1 -5 สลับกันได้  และการให้ส่วนลดเป็นไปตาม Double Discount
'_________________________
    Select Case tVB_CNTrnDiscType
        Case "1815", "1889", "1890", "1891", "1892", "1893"
            Me.Tag = "2;" & SP_INItGetValue(tVB_CNPthFileINI, Me.Name, "klg_Title2")
            Me.Caption = cUT.UT_STRtToken(Me.Tag, ";", nVB_CNCutLang + 1)
    End Select
    
'    If tVB_CNTrnDiscType = "1815" Then
'        Me.Tag = "2;" & SP_INItGetValue(tVB_CNPthFileINI, Me.Name, "klg_Title2")
'        Me.Caption = cUT.UT_STRtToken(Me.Tag, ";", nVB_CNCutLang + 1)
'    End If
End Sub
Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
'_____________________________________
'   Call:
'   Cmt:
'_____________________________________
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call ocmOk_Click
        KeyAscii = 0
    Else
        Select Case KeyAscii
            Case 48 To 58
                With ogdForm
                    tW_HotKey = tW_HotKey & Chr(KeyAscii)
                    If tW_HotKey <> "" Then
                         Call W_PRCxCallFunc(Val(tW_HotKey))
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
    Call W_DEFxChangeTitle
    Call W_FRMxSetInitial  '*KT 54-12-02 V.4.0.9 ML-PH4-R3-014 สามารถเลือก Transaction Markdown 1 -5 สลับกันได้  และการให้ส่วนลดเป็นไปตาม Double Discount
    nMaxRow = ogdForm.Aggregate(flexSTMax, 0, 0, ogdForm.rows - 1, 0)
    'ogdForm.SetFocus
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'     If UnloadMode = 0 Then Cancel = 1
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    bVB_DiscAccept = False
    Unload Me
End Sub

Private Sub ocmOk_Click()
    With ogdForm
        Select Case .Row
            Case 0
                wCNDiscOth.olaForm(0).Caption = "% Off"
                wCNDiscOth.Caption = "% Off"
                nVB_TypeDiscLineItem = nEN_LnDis1Per
            Case 1
                wCNDiscOth.olaForm(0).Caption = "Amount Off"
                 wCNDiscOth.Caption = "Amount Off"
                 nVB_TypeDiscLineItem = nEN_LnDis2Amt
            Case 2
                wCNDiscOth.olaForm(0).Caption = "New Price"
                wCNDiscOth.Caption = "New Price"
                nVB_TypeDiscLineItem = nEN_LnDis3NewPrice
        End Select
    End With
    wCNDiscOth.Show vbModal
    Unload Me
End Sub

Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_HotKey = ""
End Sub

Private Sub ogdForm_DblClick()
    Call ocmOk_Click
End Sub

Private Sub W_FRMxSetInitial()
'*KT 54-12-02 V.4.0.9 ML-PH4-R3-014 สามารถเลือก Transaction Markdown 1 -5 สลับกันได้  และการให้ส่วนลดเป็นไปตาม Double Discount
     Select Case tVB_CNTrnDiscType
        Case "1893"
            Me.Caption = Me.Caption & "1"
        Case "1892"
            Me.Caption = Me.Caption & "2"
        Case "1891"
            Me.Caption = Me.Caption & "3"
        Case "1890"
            Me.Caption = Me.Caption & "4"
        Case "1889"
            Me.Caption = Me.Caption & "5"
    End Select
End Sub
