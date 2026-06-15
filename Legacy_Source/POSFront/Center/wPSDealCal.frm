VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form wPSDealCal 
   Caption         =   "Calculate Formula"
   ClientHeight    =   3600
   ClientLeft      =   465
   ClientTop       =   2685
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   14385
   Begin VB.CommandButton ocmSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   3075
      Width           =   1215
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdDeal 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   14175
      _cx             =   25003
      _cy             =   5106
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      SubtotalPosition=   0
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
      Editable        =   2
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
   Begin VB.CommandButton ocmDel 
      Caption         =   "RemoveRow"
      Height          =   375
      Left            =   4875
      TabIndex        =   5
      Top             =   3075
      Width           =   1215
   End
   Begin VB.CommandButton ocmAdd 
      Caption         =   "AddRow"
      Height          =   375
      Left            =   3915
      TabIndex        =   4
      Top             =   3075
      Width           =   855
   End
   Begin VB.ComboBox ocbDeal 
      Height          =   315
      ItemData        =   "wPSDealCal.frx":0000
      Left            =   555
      List            =   "wPSDealCal.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3105
      Width           =   2055
   End
   Begin VB.CommandButton ocmUpdate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   2715
      TabIndex        =   0
      Top             =   3075
      Width           =   1095
   End
   Begin VB.Label olaForm 
      AutoSize        =   -1  'True
      Caption         =   "DealType:ÊèÇ¹Å´¾ÔàÈÉ 1:Free, 2:New Price, 3:Disc Amt, 4:Disc %"
      Height          =   195
      Index           =   1
      Left            =   8520
      TabIndex        =   3
      Top             =   3120
      Width           =   4665
   End
   Begin VB.Label olaForm 
      Caption         =   "Deal"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
End
Attribute VB_Name = "wPSDealCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub W_PMTxCalDeal()
'    Dim cTotalAmtGrp As Double
'    Dim nQtyDiv%
'    Select Case UCase(ocbDeal.Text)
'        Case "DEAL7"
'            Me.Caption = PT_CALcDeal07 '(nQtyDiv, cTotalAmtGrp)
'        Case "DEAL8"
'            Me.Caption = PT_CALcDeal08 '(nQtyDiv, cTotalAmtGrp)
'        Case "DEAL9"
'            Me.Caption = PT_CALcDeal09 '(nQtyDiv, cTotalAmtGrp)
'        Case "DEAL10"
'            Me.Caption = PT_CALcDeal10 '(nQtyDiv, cTotalAmtGrp)
'        Case Else
'            Me.Caption = PT_CALcDeal01_06 '(nQtyDiv, cTotalAmtGrp)
'    End Select
'
'End Sub
'
'Private Sub ocmAdd_Click()
'ogdDeal.rows = ogdDeal.rows + 1
'End Sub
'
'Private Sub ocmDel_Click()
'   If ogdDeal.rows = ogdDeal.FixedRows Then Exit Sub
'   If ogdDeal.Row > ogdDeal.FixedRows Then
'       ogdDeal.RemoveItem (ogdDeal.Row)
'   Else
'   End If
'        ogdDeal.rows = ogdDeal.rows - 1
'End Sub
'
'Private Sub ocmSave_Click()
'Dim tListIndex$
'
'tListIndex = Format(ocbDeal.ListIndex + 1, "000")
' Call PT_CALxSaveResult(tListIndex)
'End Sub
'
'Private Sub ocmUpdate_Click()
'Call W_PMTxCalDeal
'End Sub
'
''OPEN WHEN NEED TO USED CAL AMOUNTKEY
''Private Sub ogdDeal_CellChanged(ByVal Row As Long, ByVal Col As Long)
''Dim nGetQty%, nPmtQty%, nDivQty%
''Dim nDivQtyAll%
''With ogdDeal
''    Select Case Col
''        Case .ColIndex("FCDacQtyDiv")
''            nGetQty = .ValueMatrix(Row, .ColIndex("FNDacQtyGet"))
''            nPmtQty = .ValueMatrix(Row, .ColIndex("FNDacQtyBuy"))
''            nDivQty = .ValueMatrix(Row, .ColIndex("FCDacQtyDiv"))
''            nDivQtyAll = nGetQty * nPmtQty * nDivQty
''            If nDivQtyAll > .ValueMatrix(Row, .ColIndex("FNDacQtyKey")) Then nDivQtyAll = .ValueMatrix(Row, .ColIndex("FNDacQtyKey"))
''             .TextMatrix(Row, .ColIndex("FCDacQtyDivAll")) = nDivQtyAll
''    End Select
''End With
''End Sub
'
'Private Sub ogdDeal_DblClick()
'    Call ogdDeal.SaveGrid(App.Path & "\SaveGrid.txt", flexFileCommaText, True)
'End Sub
