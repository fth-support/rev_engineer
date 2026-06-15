VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNReturnMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return Menu;klg_Title"
   ClientHeight    =   2175
   ClientLeft      =   3450
   ClientTop       =   4650
   ClientWidth     =   4665
   ControlBox      =   0   'False
   Icon            =   "wCNReturnMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4665
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1080
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1360
      Width           =   4665
      _Version        =   720897
      _ExtentX        =   8229
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
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
         Left            =   3120
         TabIndex        =   2
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
         Left            =   1560
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   120
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1640
      Index           =   1
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   15
      Width           =   4665
      _Version        =   720897
      _ExtentX        =   8229
      _ExtentY        =   2893
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VSFlex7Ctl.VSFlexGrid ogdForm 
         Height          =   1095
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   4455
         _cx             =   7858
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
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"wCNReturnMenu.frx":058A
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
Attribute VB_Name = "wCNReturnMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'การทำรายการคืน (Return Transaction) การบันทึกเมื่อลูกค้า ต้องการคืนสินค้า กรณีที่ได้ทำการขายจบแล้ว(จบบิลแล้ว) โดยพนักงานเก็บเงินจะต้องสามารถทำคืนรายการขายสินค้านั้นๆ
'แบ่งเป็น 2 กรณี คือ
' 1) การคืนแบบอ้างอิงรายละเอียดการขาย
'เป็นการคืนบางรายการ กรณียังไม่ปิดการขายประจำวัน โดยลูกค้าต้องนำบิลของรายการขายมาด้วยและพนักงานเก็บเงินต้องบันทึกหมายเลขเครื่องและหมายเลขลำดับการขายของรายการขายนั้นๆ
'ระบบต้องดึงข้อมูลของรายการขายขึ้นมาเพื่อให้พนักงานเก็บเงินเลือกรายการสินค้าที่ต้องการคืนหลังจากทำรายการจบ โดยการคืนให้ลูกค้า จะต้องพิมพ์บิลที่แสดงว่าเป็นรายการคืนพร้อมรายละเอียด
'ส่วนลดซึ่งการคืนประเภทนี้ ไม่สามารถทำรายการได้ ในกรณีรายการขายมีการให้โปรโมชั่นแบบเงื่อนไขพิเศษ
' 2) การคืนแบบไม่อ้างอิงรายละเอียดการขาย
'เป็นการคืนบางรายการ กรณีที่ข้ามวันหรือกรณีรายการขายมีการให้โปรโมชั่นแบบเงื่อนไขพิเศษโดยลูกค้าต้องนำบิลของรายการขายมาด้วยและพนักงานเก็บเงินต้องบันทึกรายการที่ต้องการคืนและบันทึกส่วนลดเอง
'หลังจากทำรายการจบโดยการคืนเงินให้ลูกค้าจะต้องพิมพ์บิลที่แสดงว่ารายการคืนพร้อมรายละเอียดการคืนและส่วนลด
Option Explicit
Private tW_HotKey$
Private nMaxRow As Long

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

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    bVB_Return = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            Call ogdForm_DblClick
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call ocmOK_Click
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
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FNCxAddFunc2Grid
    nMaxRow = ogdForm.Aggregate(flexSTMax, 0, 0, ogdForm.rows - 1, 0)
     '*KT 52-01-23 Communication Sheet PH2-ML-SF-069 Return refer sale can't scroll down on item selection screen.xls
    Call SP_OBJxSetFocus(ogdForm)
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = 1
'End Sub

Private Sub ocmOK_Click()
    With ogdForm
        tVB_CNMnuCurrent = .TextMatrix(.Row, 2)
        If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
        If bVB_CNUserOveride Then Call SP_UPDxOverideEmp
        
    Select Case tVB_CNMnuCurrent
        Case "172002"
            If .Row = 0 Then
                bVB_Reference = False
                bVB_Return = True
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbReturnTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbReturnTaxInc - 1).UsrLangSec
                dVB_TransDateRet = Date '*KT 67-06-07 TTK-ISS3000-873
            Else
                bVB_Reference = True
                bVB_Return = True
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbReturnTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbReturnTaxInc - 1).UsrLangSec
            End If
        Case "172102"
            If .Row = 0 Then
                bVB_Reference = False
                bVB_Return = True
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbReturnManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbReturnManualTax - 1).UsrLangSec
                dVB_TransDateRet = Date  '*KT 67-06-07 TTK-ISS3000-873
            Else
                bVB_Reference = True
                bVB_Return = True
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbReturnManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbReturnManualTax - 1).UsrLangSec
            End If
        Case Else
            If .Row = 0 Then
                bVB_Reference = False
                bVB_Return = True
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbReturn - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbReturn - 1).UsrLangSec
                dVB_TransDateRet = Date  '*KT 67-06-07 TTK-ISS3000-873
            Else
                bVB_Reference = True
                bVB_Return = True
                tVB_TransNameABB = aVB_NameAbb(tEN_AbbRefReturn - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbRefReturn - 1).UsrLangSec
            End If
    End Select
'    Call wMainRet.W_SETxIndicator(0)       '*Ao 2014/06/25  ไม่ใช้
    End With
    Unload Me
End Sub

Private Sub ogdForm_DblClick()
    Call ocmOK_Click
End Sub

Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call:
'Cmt: ทำการแสดงรายการของ Return or Refer Sale Return
'-------------------------------------------------------
Dim nRow%
Dim tSmnID$
With ogdForm
    .rows = 2
    Select Case tVB_CNMnuCurrent
        Case "172002"
            .TextMatrix(0, 1) = Trim(SP_GETtVocabularyTicketFromArray("170701", aVB_TransMnuAll, tSmnID)) '"Return")
            .TextMatrix(0, 2) = "172002"
            .TextMatrix(1, 1) = Trim(SP_GETtVocabularyTicketFromArray("170702", aVB_TransMnuAll, tSmnID)) '"Refer Sale Return"
            .TextMatrix(1, 2) = "172002"
        Case "172102"
            .TextMatrix(0, 1) = Trim(SP_GETtVocabularyTicketFromArray("170701", aVB_TransMnuAll, tSmnID)) '"Return")
            .TextMatrix(0, 2) = "172102"
            .TextMatrix(1, 1) = Trim(SP_GETtVocabularyTicketFromArray("170702", aVB_TransMnuAll, tSmnID)) '"Refer Sale Return"
            .TextMatrix(1, 2) = "172102"
        Case Else
            .TextMatrix(0, 1) = Trim(SP_GETtVocabularyTicketFromArray("170701", aVB_TransMnuAll, tSmnID)) '"Return")
            .TextMatrix(0, 2) = tSmnID
            .TextMatrix(1, 1) = Trim(SP_GETtVocabularyTicketFromArray("170702", aVB_TransMnuAll, tSmnID)) '"Refer Sale Return"
            .TextMatrix(1, 2) = tSmnID
    End Select

    For nRow = 0 To 1
            .TextMatrix(nRow, 1) = " " & nRow + 1 & "." & .TextMatrix(nRow, 1)
            .TextMatrix(nRow, 0) = nRow + 1
    Next nRow
End With
End Sub
