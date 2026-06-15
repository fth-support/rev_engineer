VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wShowKB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ฟังก์ชัน;klg_Title"
   ClientHeight    =   6315
   ClientLeft      =   4020
   ClientTop       =   1740
   ClientWidth     =   5640
   ControlBox      =   0   'False
   Icon            =   "wShowKB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   5640
   Tag             =   "2;ฟังก์ชัน;Function KB"
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1000
      Index           =   1
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5595
      Width           =   5640
      _Version        =   720897
      _ExtentX        =   9948
      _ExtentY        =   1764
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Timer otmForm 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   360
         Top             =   240
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
         Left            =   4080
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   100
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
         Left            =   2520
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   100
         Width           =   1455
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdForm 
      Height          =   5000
      Left            =   0
      TabIndex        =   0
      Tag             =   "2;ลำดับ;No.;ฟังก์ชัน;Function;คำอธบาย;Description"
      Top             =   600
      Width           =   5640
      _cx             =   9948
      _cy             =   8819
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   11889258
      ForeColorFixed  =   -2147483634
      BackColorSel    =   11889258
      ForeColorSel    =   16777215
      BackColorBkg    =   16768725
      BackColorAlternate=   16768725
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wShowKB.frx":0442
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
      BackColorFrozen =   12632256
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   840
      Index           =   0
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   5640
      _Version        =   720897
      _ExtentX        =   9948
      _ExtentY        =   1482
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.Label olaTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Function;klg_Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   320
         Width           =   3855
      End
   End
End
Attribute VB_Name = "wShowKB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey As String
Private bW_KeyEnter As Boolean '***KT 51-11-20 ป้องกันการ Enter ซ้ำและมีการบันทึก  มากกว่า 1 ครั้งใน
Private nMaxRow As Long

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
    Dim bFound As Boolean
    Dim nI As Long
On Error GoTo ErrHandle
    With ogdForm
        bFound = False
         For nI = .FixedRows To .rows - 1
                If .TextMatrix(nI, 0) <> "" Then
                    If SP_VALnText2Double(.TextMatrix(nI, 0)) = pnFuc Then
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

Private Sub W_FRMxShowKB()
'---------------------------------------------------
'Call:
'Cmt: *KT 51-01-02 แสดงรายการบน Grid
'---------------------------------------------------
    Dim tSql$
    Dim nCol As Long
    Dim nI As Long
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle             '*KT 54-03-25 ป้องกันโปรแกรมหลุด
    tSql = "SELECT   TSysKeyboard.FNSkbSeqGrp AS FNSkbSeqGrp,TSysPosMenu.FTUsrLangPri AS FTUsrLangPri,TSysPosMenu.FTUsrLangSec AS FTUsrLangSec"
    tSql = tSql & ",TSysPosMenu.FTSpmNameUser AS FTSpmNameUser,TSysKeyboard.FTSkbCode AS FTSkbCode"
    tSql = tSql & ",TSysKeyboard.FTSkbRef AS FTSkbRef,TSysKeyboard.FNSkbKey AS FNSkbKey,TSysKeyboard.FNSkbShift AS FNSkbShift"
    tSql = tSql & ",TSysMenu.FTSmnUsage AS FTSmnUsage,TSysMenu.FNSmnVisible AS FNSmnVisible,TSysMenu.FNSmnAutho AS FNSmnAutho"
    tSql = tSql & ",TSysKeyboard.FTSkbName"
    tSql = tSql & " FROM  TSysMenu INNER JOIN"
    tSql = tSql & "   TSysPosMenu ON TSysMenu.FTSmnID = TSysPosMenu.FTSpmID INNER JOIN"
    tSql = tSql & "    TSysKeyboard ON TSysPosMenu.FTSpmID = TSysKeyboard.FTSmnID"
    tSql = tSql & " WHERE TSysKeyboard.FTSkbGrp='KB'"
    tSql = tSql & " ORDER BY TSysKeyboard.FNSkbSeqGrp"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    Set ogdForm.DataSource = orsTemp
    ogdForm.Refresh
    With ogdForm
         For nCol = 0 To orsTemp.Fields.Count - 1
             Select Case Mid(orsTemp.Fields(nCol).Name, 2, 1)
                 Case "T"
                        .ColDataType(nCol) = flexDTString
                        .ColAlignment(nCol) = flexAlignLeftCenter
                 Case "N"
                        .ColDataType(nCol) = flexDTLong
                        .ColAlignment(nCol) = flexAlignRightCenter
                        .ColFormat(nCol) = "########0"
                 Case "C"
                        .ColDataType(nCol) = flexDTCurrency
                        .ColAlignment(nCol) = flexAlignRightCenter
                        .ColFormat(nCol) = SP_FMTtNum2Set
                 Case "D"
                        .ColDataType(nCol) = flexDTDate
                        .ColAlignment(nCol) = flexAlignCenterCenter
                        .ColFormat(nCol) = "dd/mm/yyyy"
                 Case "B"
                        .ColDataType(nCol) = flexDTBoolean
                        .ColAlignment(nCol) = flexAlignCenterCenter
                 Case Else
             End Select
            .ColKey(nCol) = orsTemp.Fields(nCol).Name
            .ColHidden(nCol) = True
        Next nCol
        .ColHidden(.ColIndex("FTSpmNameUser")) = False
        .ColWidth(.ColIndex("FTSpmNameUser")) = 2800
        .ColWidth(.ColIndex("FTSkbName")) = 1200
         For nI = .FixedRows To .rows - 1
            If .ValueMatrix(nI, .ColIndex("FNSmnVisible")) = 0 Then .RowHidden(nI) = True
            If .TextMatrix(nI, .ColIndex("FNSkbSeqGrp")) <> "" Then
                .TextMatrix(nI, .ColIndex("FTSpmNameUser")) = .TextMatrix(nI, .ColIndex("FNSkbSeqGrp")) & "." & .TextMatrix(nI, .ColIndex("FTSpmNameUser"))
            End If
         Next nI
    End With
    Set orsTemp = Nothing 'ไม่ได้คำสั่ง Close กรณี set datasource ใส่ grid
    Exit Sub
ErrHandle:
    Set orsTemp = Nothing 'ไม่ได้คำสั่ง Close กรณี set datasource ใส่ grid
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Call ocmOK_Click
        If bW_KeyEnter Then
            KeyAscii = 0
        Else
            bW_KeyEnter = True
            Call ocmOK_Click
        End If
    Else
        Select Case KeyAscii
            Case 48 To 58
                otmForm.Enabled = True
                With ogdForm
                    tW_HotKey = tW_HotKey & Chr(KeyAscii)
                    If tW_HotKey <> "" Then
                         Call W_PRCxCallFunc(Val(tW_HotKey))
                    End If
                    If IsNumeric(tW_HotKey) Then
                        If CLng(tW_HotKey) >= nMaxRow Then tW_HotKey = "": otmForm.Enabled = False
                    End If
                End With
        End Select
    End If
End Sub

Private Sub Form_Load()
     tW_HotKey = ""
     bW_KeyEnter = False
     bVB_FKBCancel = False
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxShowKB
    nMaxRow = ogdForm.Aggregate(flexSTMax, 0, 0, ogdForm.rows - 1, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    otmForm.Enabled = False
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:
    
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'
'End Sub

Private Sub ocmCancel_Click()
    bVB_FKBCancel = True
    bVB_FormPayAct = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
    With ogdForm
        bVB_SetFocus = False
        nVB_FKBKeyCode = SP_FMTtNum2Shw(.TextMatrix(.Row, .ColIndex("FNSkbKey")), True)
        nVB_FKBShift = SP_FMTtNum2Shw(.TextMatrix(.Row, .ColIndex("FNSkbShift")), True)
        If .TextMatrix(.Row, .ColIndex("FTSkbRef")) = "W_KBQtyChange" Then
            bVB_FactQtyChange = True
        Else
            bVB_FactQtyChange = False
        End If
        bVB_FormPayAct = True
        Unload Me
    End With
End Sub

'Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    tW_HotKey = ""
'End Sub

Private Sub ogdForm_DblClick()
    Call ocmOK_Click
End Sub

Private Sub otmForm_Timer()
'    Dim nMaxRow As Long
'    nMaxRow = ogdForm.Aggregate(flexSTMax, 0, 0, ogdForm.rows - 1, 0)
    If IsNumeric(tW_HotKey) Then tW_HotKey = "": otmForm.Enabled = False
End Sub

