VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wTender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tender Amount;klg_Title"
   ClientHeight    =   8040
   ClientLeft      =   1500
   ClientTop       =   1920
   ClientWidth     =   11970
   ControlBox      =   0   'False
   Icon            =   "wTender.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11970
   Tag             =   "2;"
   Begin VB.CommandButton ocmDetail 
      Caption         =   "Detail(F4);klg_Detail"
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
      Left            =   6900
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   8070
      Visible         =   0   'False
      Width           =   1455
   End
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
      Height          =   585
      Left            =   3900
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   7110
      Width           =   1515
   End
   Begin VB.CommandButton ocmOK 
      Caption         =   "&OK;klg_Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1155
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   7110
      Width           =   1515
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdOtherCash 
      Height          =   570
      Left            =   120
      TabIndex        =   32
      Top             =   5520
      Width           =   6255
      _cx             =   11033
      _cy             =   1005
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
      BackColor       =   16769505
      ForeColor       =   0
      BackColorFixed  =   11889258
      ForeColorFixed  =   0
      BackColorSel    =   11889258
      ForeColorSel    =   16777215
      BackColorBkg    =   16768725
      BackColorAlternate=   16769505
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
      SelectionMode   =   0
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wTender.frx":6852
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
      WallPaperAlignment=   4
   End
   Begin XtremeSuiteControls.TabControl otaForm 
      Height          =   8955
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "2;&บันทึกข้อมูล;&Data Entry;&ตัวเลือก;&Optional"
      Top             =   -360
      Width           =   12015
      _Version        =   720897
      _ExtentX        =   21193
      _ExtentY        =   15796
      _StockProps     =   64
      Appearance      =   10
      Color           =   32
      PaintManager.Layout=   2
      ItemCount       =   4
      SelectedItem    =   3
      Item(0).Caption =   "Coin"
      Item(0).ControlCount=   0
      Item(1).Caption =   "Currentcy"
      Item(1).ControlCount=   0
      Item(2).Caption =   "GiftCard"
      Item(2).ControlCount=   0
      Item(3).Caption =   "Item"
      Item(3).ControlCount=   20
      Item(3).Control(0)=   "olaTotal"
      Item(3).Control(1)=   "otbSumTender"
      Item(3).Control(2)=   "otcForm1(2)"
      Item(3).Control(3)=   "ogdFunc"
      Item(3).Control(4)=   "otcForm1(0)"
      Item(3).Control(5)=   "otcForm1(1)"
      Item(3).Control(6)=   "ogdFunc1"
      Item(3).Control(7)=   "otbSumForeign"
      Item(3).Control(8)=   "otbSumCash"
      Item(3).Control(9)=   "otbSumTotalGrid1"
      Item(3).Control(10)=   "ogdGiftVoucher"
      Item(3).Control(11)=   "ogdSumCash"
      Item(3).Control(12)=   "ogdSumForeign"
      Item(3).Control(13)=   "otcForm1(3)"
      Item(3).Control(14)=   "ogdGVFunc"
      Item(3).Control(15)=   "ogdOtherGV"
      Item(3).Control(16)=   "ogdTotalCash"
      Item(3).Control(17)=   "ogdTotalGV"
      Item(3).Control(18)=   "ogdVendorVchr"
      Item(3).Control(19)=   "otcForm1(4)"
      Begin VSFlex7Ctl.VSFlexGrid ogdVendorVchr 
         Height          =   570
         Left            =   6480
         TabIndex        =   35
         Top             =   7760
         Width           =   5415
         _cx             =   9551
         _cy             =   1005
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
         BackColor       =   16769505
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wTender.frx":691B
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
         WallPaperAlignment=   4
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   1350
         Index           =   2
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   7260
         Width           =   6255
         _Version        =   720897
         _ExtentX        =   11033
         _ExtentY        =   2381
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         PaintManager.FixedTabWidth=   50
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   780
         Index           =   4
         Left            =   6480
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   7260
         Visible         =   0   'False
         Width           =   5415
         _Version        =   720897
         _ExtentX        =   9551
         _ExtentY        =   1376
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   3360
            TabIndex        =   40
            Top             =   30
            Width           =   1515
         End
         Begin VB.Label olaFrom 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Foreign Currency"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   5
            Left            =   6120
            TabIndex        =   39
            Top             =   3840
            Width           =   4155
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Vch"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   -120
            TabIndex        =   38
            Top             =   30
            Width           =   2355
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   2400
            TabIndex        =   37
            Top             =   30
            Visible         =   0   'False
            Width           =   795
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid ogdSumCash 
         Height          =   570
         Left            =   120
         TabIndex        =   22
         Top             =   7260
         Visible         =   0   'False
         Width           =   6255
         _cx             =   11033
         _cy             =   1005
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483647
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wTender.frx":69ED
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
      Begin VSFlex7Ctl.VSFlexGrid ogdGVFunc 
         Height          =   1050
         Left            =   6480
         TabIndex        =   30
         Top             =   4680
         Width           =   5415
         _cx             =   9551
         _cy             =   1852
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
         BackColor       =   16769505
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
         SelectionMode   =   0
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   15
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wTender.frx":6A51
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
         WallPaperAlignment=   4
      End
      Begin VB.TextBox otbSumCash 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         MaxLength       =   10
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   8160
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.TextBox otbSumForeign 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7650
         MaxLength       =   10
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   8040
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.TextBox otbSumTender 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   5160
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VSFlex7Ctl.VSFlexGrid ogdFunc 
         Height          =   4770
         Left            =   120
         TabIndex        =   0
         Top             =   960
         Width           =   6255
         _cx             =   11033
         _cy             =   8414
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
         BackColor       =   16769505
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
         SelectionMode   =   0
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   15
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wTender.frx":6B23
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
         WallPaperAlignment=   4
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   780
         Index           =   0
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Width           =   6255
         _Version        =   720897
         _ExtentX        =   11033
         _ExtentY        =   1376
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin VB.Label olaFrom 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   3
            Left            =   1960
            TabIndex        =   21
            Top             =   30
            Width           =   2355
         End
         Begin VB.Label olaForm 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Collections"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   -500
            TabIndex        =   13
            Top             =   30
            Width           =   3555
         End
         Begin VB.Label olaFrom 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   0
            Left            =   3760
            TabIndex        =   12
            Top             =   30
            Width           =   2355
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid ogdFunc1 
         Height          =   3090
         Left            =   6480
         TabIndex        =   1
         Top             =   960
         Width           =   5415
         _cx             =   9551
         _cy             =   5450
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
         BackColor       =   16769505
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
         SelectionMode   =   0
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wTender.frx":6BF5
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
         WallPaperAlignment=   4
         Begin VB.TextBox otbSumGiftVoucher 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   960
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   5760
            Visible         =   0   'False
            Width           =   4215
         End
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   780
         Index           =   1
         Left            =   6480
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   5415
         _Version        =   720897
         _ExtentX        =   9551
         _ExtentY        =   1376
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin VB.Label olaFrom 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   2
            Left            =   3360
            TabIndex        =   16
            Top             =   30
            Width           =   1515
         End
         Begin VB.Label olaFrom 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Foreign Currency"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   1
            Left            =   -600
            TabIndex        =   15
            Top             =   30
            Width           =   4155
         End
      End
      Begin VB.TextBox otbSumTotalGrid1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   7320
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VSFlex7Ctl.VSFlexGrid ogdSumForeign 
         Height          =   570
         Left            =   6480
         TabIndex        =   23
         Top             =   7280
         Visible         =   0   'False
         Width           =   5415
         _cx             =   9551
         _cy             =   1005
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483647
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wTender.frx":6CC7
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
      Begin VSFlex7Ctl.VSFlexGrid ogdGiftVoucher 
         Height          =   570
         Left            =   120
         TabIndex        =   24
         Top             =   6000
         Visible         =   0   'False
         Width           =   6255
         _cx             =   11033
         _cy             =   1005
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
         BackColor       =   16769505
         ForeColor       =   0
         BackColorFixed  =   11889258
         ForeColorFixed  =   0
         BackColorSel    =   11889258
         ForeColorSel    =   16777215
         BackColorBkg    =   16768725
         BackColorAlternate=   16769505
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
         SelectionMode   =   0
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wTender.frx":6D2F
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
         WallPaperAlignment=   4
      End
      Begin VSFlex7Ctl.VSFlexGrid ogdOtherGV 
         Height          =   570
         Left            =   6480
         TabIndex        =   25
         Top             =   5880
         Width           =   5415
         _cx             =   9551
         _cy             =   1005
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
         BackColor       =   16769505
         ForeColor       =   0
         BackColorFixed  =   11889258
         ForeColorFixed  =   0
         BackColorSel    =   11889258
         ForeColorSel    =   16777215
         BackColorBkg    =   16768725
         BackColorAlternate=   16769505
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
         SelectionMode   =   0
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wTender.frx":6DF8
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
         WallPaperAlignment=   4
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   780
         Index           =   3
         Left            =   6480
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4200
         Width           =   5415
         _Version        =   720897
         _ExtentX        =   9551
         _ExtentY        =   1376
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin VB.Label olaQty 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   2400
            TabIndex        =   31
            Top             =   30
            Width           =   795
         End
         Begin VB.Label olaGVName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Gift Voucher"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   -120
            TabIndex        =   29
            Top             =   30
            Width           =   2355
         End
         Begin VB.Label olaFrom 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Foreign Currency"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Index           =   4
            Left            =   6120
            TabIndex        =   28
            Top             =   3840
            Width           =   4155
         End
         Begin VB.Label olaGVAmt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   3600
            TabIndex        =   27
            Top             =   30
            Width           =   1515
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid ogdTotalCash 
         Height          =   570
         Left            =   120
         TabIndex        =   33
         Top             =   6600
         Width           =   6255
         _cx             =   11033
         _cy             =   1005
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483647
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wTender.frx":6EC1
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
      Begin VSFlex7Ctl.VSFlexGrid ogdTotalGV 
         Height          =   570
         Left            =   6480
         TabIndex        =   34
         Top             =   6600
         Width           =   5415
         _cx             =   9551
         _cy             =   1005
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483647
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wTender.frx":6F2A
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
      Begin VB.Label olaTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   7560
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin VB.CommandButton ocmNext 
      Caption         =   "Next(F6)"
      Height          =   495
      Left            =   2190
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1455
   End
   Begin VB.CommandButton ocmPrevious 
      Caption         =   "Previous(F5)"
      Height          =   495
      Left            =   630
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1455
   End
End
Attribute VB_Name = "wTender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'หมายเหตุ : หน้าจอนี้สามารถระบุจำนวนเงินได้ทั้งหมด 10 degit เป็น ทศนิยม 2 degit  จำนวนเงินเต็ม 7 degit เครื่องจุด 1 degit
Option Explicit
Const tW_StdHD = "Shd"
Const tW_StdDT = "Sdt"
'Dim bW_FirstLoad As Boolean

Private nW_DiscType As Integer
Private tW_KeyPress As String

Private Type TypeCurrency
    FTSkuCode As String
    FCSdtStkFac As Double
    FCSdtRegPrice As Double
End Type

Private Sub W_UPDxVirtualPosRC()
'---------------------------------------------------------------
'   Call:
'   Cmt: '*Eaw 56-10-01 เก็บค่า VirtualPos ให้ตาราง SaleRC ใน Re-Entry Mode
'---------------------------------------------------------------
    Dim tSql As String
    Dim tTerminal As String
On Error GoTo HandleError
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
            tTerminal = ""
            If bVB_ReEntryMode Then
                tTerminal = tVB_CNVirTerminalNum
            End If
            tSql = "UPDATE " & tVB_CNTblRC
            tSql = tSql & vbCrLf & "SET FTRemark = '" & tTerminal & "'"
            tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
            tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
            Call SP_SQLvExecute(tSql)
    End Select
HandleError:

End Sub
Private Sub W_UPDxVirtualPosDT()
'---------------------------------------------------------------
'   Call:
'   Cmt: '*Eaw 56-10-01 เก็บค่า VirtualPos ให้ตาราง SaleDT ใน Re-Entry Mode
'---------------------------------------------------------------
    Dim tSql As String
    Dim tTerminal As String
On Error GoTo HandleError

    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
            tTerminal = ""
            If bVB_ReEntryMode Then
                tTerminal = tVB_CNVirTerminalNum
            Else
                tTerminal = tVB_CNTerminalNum
            End If
            tSql = "UPDATE " & tVB_CNTblDT
            tSql = tSql & vbCrLf & "SET FTWahCode = '" & tTerminal & "'"
            tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
            tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
            Call SP_SQLvExecute(tSql)
    End Select
    Exit Sub
HandleError:

End Sub

Private Function W_CHKbOverShortReport(ByVal pcSumEnter As Double) As Boolean
'---------------------------------------------------------------
'   Call:
'   Cmt: '*Ao 2013/07/31
'---------------------------------------------------------------
Dim tSql As String, tTenderType As String, tTmp As String, tTransWhere As String
Dim orsTemp As ADODB.Recordset
Dim cTenderALL As Double, cTender As Double, cChg As Double, cGVch As Double, cCurrency As Double, cACTUAL As Double, cTmp As Double, cVVch As Double '*Tao 62-09-05 เพิ่ม cVVch
Dim bChk As Boolean
Dim cCashTransfer As Double, CashTransfer As Double, cCashTrans As Double, cGVchTrans As Double, cCurrencyTrans As Double, cVVchTrans As Double '*Tao 62-09-05 เพิ่ม cVVchTrans
Dim cCHK  As Double, cCurrencyTmp As Double, cCurrencyTransTmp As Double
Dim aCurrency() As TypeCurrency
Dim nI  As Integer, nLoop As Integer
Dim tSKUCode As String

On Error GoTo HandleError
    
    W_CHKbOverShortReport = False
    
    tSql = wFunctionKB.W_GETtCloseTenderSQLEx(tVB_CNTblHD, tVB_CNTblRC, "FTShdTransType", True)
    Call SP_SQLvExecute(tSql, orsTemp)
    
    cTender = 0
    If (orsTemp.BOF And orsTemp.EOF) Then
        cTender = 0
    Else
        orsTemp.Filter = "FTTdmCode = 'T001'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                cTender = cTender + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))
                orsTemp.MoveNext
            Loop
        End If
    End If
    
    orsTemp.Filter = adFilterNone
    
    cChg = 0
    If (orsTemp.BOF And orsTemp.EOF) Then
        cChg = 0
    Else
        orsTemp.MoveFirst
        Do While Not (orsTemp.EOF)
            tTenderType = SP_FEDtChkString(orsTemp, "FTTdmCode")
            Select Case UCase(Trim(tTenderType))
                Case "T001", "T004", "T005", "T006", "T007", "T008", "T009", "T010", "T011", "T012", "T014"
                    cChg = cChg + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSrcChg"))
            End Select
            orsTemp.MoveNext
        Loop
    End If
    cTenderALL = cTender - cChg
    
    orsTemp.Filter = adFilterNone
    
    cGVch = 0
    If (orsTemp.BOF And orsTemp.EOF) Then
        cGVch = 0
    Else
        orsTemp.Filter = "FTTdmCode = 'T005'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                If SP_FEDtChkString(orsTemp, "FTTdmCode") = "T018" And (SP_FEDtChkString(orsTemp, "FTShdTransType") = "04" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "39" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "44" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "14") Then
                    cGVch = cGVch + 0
                Else
                    cGVch = cGVch + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))
                End If
                orsTemp.MoveNext
            Loop
        End If
    End If
    
    orsTemp.Filter = adFilterNone
    
    '*Tao 62-09-05 เพิ่ม cVVch
    cVVch = 0
    If (orsTemp.BOF And orsTemp.EOF) Then
        cVVch = 0
    Else
        orsTemp.Filter = "FTTdmCode = 'T006'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                cVVch = cVVch + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))
                orsTemp.MoveNext
            Loop
        End If
    End If
    
    orsTemp.Filter = adFilterNone
    
    Call SP_DATxRsNothing(orsTemp)
    
    tSql = wFunctionKB.W_GETtCloseTenderCurSQL(tVB_CNTblHD, tVB_CNTblRC, "FTShdTransType", True)
    Call SP_SQLvExecute(tSql, orsTemp)
    
    ReDim aCurrency(0)        '*Ao 2013/10/11   V4.0.78
    cCurrency = 0
    If (orsTemp.BOF And orsTemp.EOF) Then
        cCurrency = 0
    Else
        orsTemp.MoveFirst
        Do While Not orsTemp.EOF
'            cCurrency = cCurrency + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))   '*Ao 2013/10/09   การปัดเศษทำให้มีเศษเหลือ
'            cCurrencyTmp = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSrcRteFac") * SP_FEDvChkNumeric(orsTemp, "FCSrcRteAmt"), True, False)
'            cCurrency = cCurrency + cCurrencyTmp
            '--------------------------------------------------------------------------------------'*Ao 2013/10/11
            ReDim Preserve aCurrency(UBound(aCurrency) + 1)
            aCurrency(UBound(aCurrency)).FTSkuCode = Trim(SP_FEDtChkString(orsTemp, "FTRteCode"))
            aCurrency(UBound(aCurrency)).FCSdtStkFac = SP_FEDvChkNumeric(orsTemp, "FCSrcRteFac")
            aCurrency(UBound(aCurrency)).FCSdtRegPrice = SP_FEDvChkNumeric(orsTemp, "FCSrcRteAmt")
            '--------------------------------------------------------------------------------------
            orsTemp.MoveNext
        Loop
    End If
    '----------------------------------------------------------------------------
    '*Ao 2013/09/23     CommSheet TK-ISS3000-305-UAT_00.00.03 FSBT NG
    
    tTransWhere = tVB_CNTblHD & ".FDShdTransDate = '" & tUT_SaleDate & "' AND " & tVB_CNTblHD & ".FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tTransWhere = tTransWhere & vbCrLf & "AND " & tVB_CNTblHD & ".FTShdTransType = '12'"
    tTransWhere = tTransWhere & vbCrLf & "GROUP BY FTSkuCode,FTSkuAbbName,FTSkuAbbNameSnd,FTSdtStaPdt,FCSdtStkFac"
    
    tSql = "SELECT FTSkuCode,FTSkuAbbName,FTSkuAbbNameSnd,FCSdtStkFac,SUM(FCSdtSaleAmt) AS FCSum,SUM(FCSdtRegPrice) AS FCSum2,FTSdtStaPdt"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblHD & " INNER JOIN"
    tSql = tSql & vbCrLf & tVB_CNTblDT & " ON " & tVB_CNTblHD & ".FTTmnNum = " & tVB_CNTblDT & ".FTTmnNum AND " & tVB_CNTblHD & ".FTShdTransNo = " & tVB_CNTblDT & ".FTShdTransNo AND"
    tSql = tSql & vbCrLf & tVB_CNTblHD & ".FTShdTransType = " & tVB_CNTblDT & ".FTShdTransType And " & tVB_CNTblHD & ".FDShdTransDate = " & tVB_CNTblDT & ".FDShdTransDate"
    tSql = tSql & vbCrLf & "WHERE " & tTransWhere
    Call SP_SQLvExecute(tSql, orsTemp)
    
    cCashTransfer = 0
    cCashTrans = 0
    cGVchTrans = 0
    cCurrencyTrans = 0
    cVVchTrans = 0 '*Tao 62-09-05 เพิ่ม cVVchTrans
    
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        orsTemp.Filter = "FTSdtStaPdt = '0' OR FTSdtStaPdt = '1'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                cCashTrans = cCashTrans + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSum"))
                orsTemp.MoveNext
            Loop
        Else
            cCashTrans = 0
        End If
        
        orsTemp.Filter = adFilterNone
        orsTemp.Filter = "FTSdtStaPdt = '2'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                cGVchTrans = cGVchTrans + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSum"))
                orsTemp.MoveNext
            Loop
        Else
            cGVchTrans = 0
        End If
        
        '*Tao 62-09-05 เพิ่ม cVVchTrans
        orsTemp.Filter = adFilterNone
        orsTemp.Filter = "FTSdtStaPdt = '6'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
                cVVchTrans = cVVchTrans + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSum"))
                orsTemp.MoveNext
            Loop
        Else
            cVVchTrans = 0
        End If
        
        orsTemp.Filter = adFilterNone
        orsTemp.Filter = "FTSdtStaPdt = '3'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            orsTemp.MoveFirst
            Do While Not orsTemp.EOF
'                cCurrencyTrans = cCurrencyTrans + CDbl(SP_FEDvChkNumeric(orsTemp, "FCSum"))     '*Ao 2013/10/09   การปัดเศษทำให้มีเศษเหลือ
'                cCurrencyTransTmp = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSdtStkFac") * SP_FEDvChkNumeric(orsTemp, "FCSum2"), True, False)
'                cCurrencyTrans = cCurrencyTrans + cCurrencyTransTmp
'                orsTemp.MoveNext
                '----------------------------------------------------------------------------------------------'*Ao 2013/10/11
                tSKUCode = Trim(SP_FEDtChkString(orsTemp, "FTSkuCode"))
                nI = -1
                For nLoop = 0 To UBound(aCurrency)
                    If tSKUCode = aCurrency(nLoop).FTSkuCode Then
                        nI = nLoop
                        Exit For
                    Else
                        nI = -1
                    End If
                Next nLoop
                If nI = -1 Then
                    cCurrencyTransTmp = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCSdtStkFac") * SP_FEDvChkNumeric(orsTemp, "FCSum2"), True, False)
                    cCurrencyTrans = cCurrencyTrans + cCurrencyTransTmp
                Else
                    cCurrencyTmp = SP_FMTcCurPoint((aCurrency(nLoop).FCSdtRegPrice - SP_FEDvChkNumeric(orsTemp, "FCSum2")) * SP_FEDvChkNumeric(orsTemp, "FCSdtStkFac"), True, False)
                    cCurrency = cCurrency + cCurrencyTmp
                    aCurrency(nLoop).FCSdtRegPrice = 0
                    aCurrency(nLoop).FCSdtStkFac = 0
                End If
                orsTemp.MoveNext
                '----------------------------------------------------------------------------------------------
            Loop
        Else
            cCurrencyTrans = 0
        End If
'        '----------------------------------------------------------------------------------------------'*Ao 2013/10/11
'        For nLoop = 0 To UBound(aCurrency)
'            cCurrencyTmp = SP_FMTcCurPoint(aCurrency(nLoop).FCSdtRegPrice * aCurrency(nLoop).FCSdtStkFac, True, False)
'            cCurrency = cCurrency + cCurrencyTmp
'        Next nLoop
'        '----------------------------------------------------------------------------------------------
    End If
    '----------------------------------------------------------------------------------------------'*Ao 2013/10/11
    For nLoop = 0 To UBound(aCurrency)
        cCurrencyTmp = SP_FMTcCurPoint(aCurrency(nLoop).FCSdtRegPrice * aCurrency(nLoop).FCSdtStkFac, True, False)
        cCurrency = cCurrency + cCurrencyTmp
    Next nLoop
    '----------------------------------------------------------------------------------------------
    cCashTransfer = cCashTrans + cGVchTrans + cCurrencyTrans + cVVchTrans '*Tao 62-09-05 เพิ่ม cVVchTrans
    '----------------------------------------------------------------------------
    cACTUAL = cTenderALL + cGVch + cCurrency + cVVch '*Tao 62-09-05 เพิ่ม cVVch
    
    cCHK = cACTUAL - cCashTransfer
    
    If SP_FMTcCurPoint(pcSumEnter, True, False) = SP_FMTcCurPoint(cCHK, True, False) Then
        bChk = False
    Else
        If pcSumEnter > cCHK Then
            tTmp = "Amount Input is Excess."
        Else
            tTmp = "Amount Input is Short."
        End If
        cTmp = SP_FMTcCurPoint(CDbl(pcSumEnter) - CDbl(cCHK), False, True)
        wOver.tW_Msg = tTmp
        wOver.cW_Amount = cTmp
        wOver.Show vbModal
        If bVB_OverShort Then
            bChk = False
        Else
            bChk = True
        End If
    End If
    W_CHKbOverShortReport = bChk
    
    Call SP_DATxRsNothing(orsTemp)
Exit Function

HandleError:
    Call SP_DATxRsNothing(orsTemp)
    W_CHKbOverShortReport = False
End Function

Private Sub W_FNCxAddFunc2GridGV()
'-----------------------------------------------------------
'   Parameter :
'   Cmt :
'   File Refer :
'   Developer : PATCHARIN
'   Date Create : 26/08/2013 :  16:22:07
'-----------------------------------------------------------
Dim nI As Integer, nR As Integer
Dim orsTemp As ADODB.Recordset
Dim tSql As String
On Error GoTo ErrHandle
    With ogdGVFunc

        tSql = "SELECT FTGvcDesc,FCTGvcAmount FROM TCNMGiftVoucher"
        tSql = tSql & vbCrLf & " ORDER BY FCTGvcAmount"
        
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read TCNMBanknote
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            nR = 0
            .rows = orsTemp.RecordCount
            Do While Not orsTemp.EOF
                .TextMatrix(nR, 0) = nR + 1
                .TextMatrix(nR, 1) = SP_FMTtNum2Shw(SP_FEDtChkString(orsTemp, "FCTGvcAmount")) & " " & SP_FEDtChkString(orsTemp, "FTGvcDesc") 'SP_FMTtNum2Shw(SP_FEDvChkNumeric(orsTemp, "FCBnkAmt")) & " " & SP_FEDtChkString(orsTemp, "FTBnkNameEn")
                .TextMatrix(nR, 2) = 0
                .TextMatrix(nR, 3) = 0
                .TextMatrix(nR, 4) = SP_FMTcCurPoint(SP_FEDtChkString(orsTemp, "FCTGvcAmount"), , False) 'SP_FEDtChkString(orsTemp, "FCTGvcAmount")
'                .TextMatrix(nR, 4) = SP_FEDvChkNumeric(orsTemp, "FCBnkAmt")
'                .TextMatrix(nR, 5) = SP_FEDtChkString(orsTemp, "FTBnkCode")
'                .TextMatrix(nR, 6) = SP_FEDtChkString(orsTemp, "FTBnkNameTh")
'                .TextMatrix(nR, 7) = SP_FEDtChkString(orsTemp, "FTBnkStaTpye")
                .Cell(flexcpBackColor, nR, 2, nR, 2) = &H80000005
                nR = nR + 1
                orsTemp.MoveNext
            Loop
            .ColAlignment(2) = flexAlignRightCenter
        End If
    End With
    
    With ogdOtherGV
            Select Case SP_VALnText2Double(tVB_TransactionSub)
                 Case tEN_TrnPickUp, tEN_TrnDepositMoney, tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                       ' .rows = .rows + 1
                        nR = 0
                        .TextMatrix(nR, 0) = nR + 1
'                        .TextMatrix(nR, 1) = "GIFT VOUCHER"
                        .TextMatrix(nR, 1) = "OTHER GVch"
                        .TextMatrix(nR, 2) = 0
                        .TextMatrix(nR, 3) = SP_FMTcCurPoint(0, False, True)
                        .TextMatrix(nR, 4) = 0
'                        .TextMatrix(nR, 5) = "T018"
                        .TextMatrix(nR, 5) = "T005"                             '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
'                        .TextMatrix(nR, 5) = "T002"                             '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
'                        .TextMatrix(nR, 6) = "GIFT VOUCHER"
                        .TextMatrix(nR, 6) = "OTHER GVch"
                        .TextMatrix(nR, 7) = "2"
                        .Cell(flexcpBackColor, nR, 3, nR, 3) = &H80000005
             End Select
            .ColAlignment(3) = flexAlignRightCenter
    End With

    Exit Sub
ErrHandle:
    Set orsTemp = Nothing
    Call SP_MSGShowError
End Sub

Private Sub W_FNCxAddFunc2GridVendorVchr()
'*Tao 62-07-05 เพิ่ม Vendor Voucher
Dim nI As Integer, nR As Integer
Dim tVVchrName As String
On Error GoTo ErrHandle
    For nI = LBound(aVB_TenderMnu) To UBound(aVB_TenderMnu) - 1
        If aVB_TenderMnu(nI).ID = "T006" Then
            tVVchrName = "VVch" 'aVB_TenderMnu(nI).Name
            Exit For
        End If
    Next nI
    With ogdVendorVchr
        Select Case SP_VALnText2Double(tVB_TransactionSub)
             Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                    nR = 0
                    .TextMatrix(nR, 0) = nR + 1
                    .TextMatrix(nR, 1) = tVVchrName
                    .TextMatrix(nR, 2) = 0
                    .TextMatrix(nR, 3) = SP_FMTcCurPoint(0, False, True)
                    .TextMatrix(nR, 4) = 0
                    .TextMatrix(nR, 5) = "T006"
                    .TextMatrix(nR, 6) = tVVchrName
                    .TextMatrix(nR, 7) = "5"
                    '*Tao 62-08-27 ไม่ให้คีย์จำนวน
                    .ColHidden(2) = True
                    .ColWidth(1) = 2800
                    '======================
                    .Cell(flexcpBackColor, nR, 2, nR, 3) = &H80000005
                    .ColAlignment(3) = flexAlignRightCenter
                    .Visible = True
                    otcForm1(4).Visible = True
                Case Else
                    .Visible = False
                    otcForm1(4).Visible = False
         End Select
    End With
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Sub W_FRMxSetInitial()
    otbSumTender.Text = 0
    otaForm.Item(0).Visible = False
    otaForm.Item(1).Visible = False
    otaForm.Item(2).Visible = False
    ocmPrevious.Visible = False
'    Call SP_OBJxSetFocus(otbSumTender)
End Sub

Private Function W_PRCbStartBank() As Boolean
On Error GoTo ErrHandle '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
    W_PRCbStartBank = False
    '*KT 51-06-23 Comsheet ML-SF-361
'    If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbSumTender.Text))) Then Exit Function      '*AO51-03-26   Comm.Sheet  ML-SF-100
    Call SP_INSxTransHD(tVB_CNTblHD)
    Call W_INSxTranDT

    Call W_INSxTranRC
    Call W_UPDxTransHD
    Call SP_UPDxSetCompleteTerminalMtn
    'process ตรวจการเปิดลิ้นชัก
    tVB_Transaction = aVB_TransMode(2).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(2).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSale - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(2).PrintType
    W_PRCbStartBank = True  '*AO51-03-26   Comm.Sheet  ML-SF-100 '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
    Exit Function
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCbStartBank") '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
End Function

Private Sub W_PRCxPickUp()
On Error GoTo ErrHandle '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
'    otbSumTender.Text = SP_CHKtCurPoint(otbSumTender.Text)          '*AO51-03-12
    If SP_VALcText2Double(otbSumTender.Text) > 0 Then
        Call SP_INSxTransHD(tVB_CNTblHD)
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnDepositMoney, tEN_TrnPickUp
                wMain.olaDocNo.Caption = tVB_Running '*KT 52-02-23 640 or 800
        End Select
        Call W_INSxTranDT
        Call W_INSxTranRC
        Call W_UPDxTransHD
        Call SP_UPDxSetCompleteTerminalMtn
        bVB_CNStartBankDT = True
        Unload Me
        'wCNBegin.Show vbModal
    Else
        Call SP_MSGnShowing("tms_CN005242", nCS_Warning)
        'Call ocmCancel_Click
    End If
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCxPickUp") '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
End Sub

Private Sub W_PRCxDepositMny()
On Error GoTo ErrHandle '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
'    otbSumTender.Text = SP_CHKtCurPoint(otbSumTender.Text)          '*AO51-03-12
    If SP_VALcText2Double(otbSumTender.Text) > 0 Then
        Call SP_INSxTransHD(tVB_CNTblHD)
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnDepositMoney, tEN_TrnPickUp
                wMain.olaDocNo.Caption = tVB_Running  '*KT 52-02-23 640 or 800
        End Select
        Call W_INSxTranDT
        Call W_INSxTranRC
        Call W_UPDxTransHD
        Call SP_UPDxSetCompleteTerminalMtn
        bVB_CNStartBankDT = True
        Unload Me
        'wCNBegin.Show vbModal
    Else
        'Call ocmCancel_Click
        Call SP_MSGnShowing("tms_CN005242", nCS_Warning)
    End If
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:W_PRCxDepositMny") '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
End Sub

Private Sub W_PRCxOverShortReport()
On Error GoTo ErrHandle
    If SP_VALcText2Double(otbSumTender.Text) >= 0 Then
        If W_CHKbOverShortReport(otbSumTender.Text) Then
            Exit Sub
        End If
        Call SP_INSxTransHD(tVB_CNTblHD)
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                wMain.olaDocNo.Caption = tVB_Running  '*KT 52-02-23 640 or 800
        End Select
        Call W_INSxTranDT
        Call W_INSxTranRC
        Call W_UPDxTransHD
        Call SP_UPDxSetCompleteTerminalMtn
        bVB_CNStartBankDT = True
        Unload Me
    Else
        Call ocmCancel_Click
    End If
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_UPDxTransHD()
'------------------------------------------------------
'Call:
'Cmt:   *KT 50-12-24 ปรับปรุงข้อมูล Transaction HD
'-----------------------------------------------------
    Dim tSql$
    tSql = "UPDATE " & tVB_CNTblHD
    tSql = tSql & vbCrLf & " SET FCShdTotal =" & SP_FMTcCurPoint(SP_VALcText2Double(otbSumTender.Text), True, False)
    tSql = tSql & vbCrLf & ",FCShdGrand=" & SP_FMTcCurPoint(SP_VALcText2Double(otbSumTender.Text), True, False)
    tSql = tSql & vbCrLf & ",FTShdStaDoc  = '1'"
    tSql = tSql & vbCrLf & ",FTShdStaPaid='3'"               '3:จ่ายครบ
    tSql = tSql & vbCrLf & ",FTShdStaPrcDoc='1'"
    '*KT 52-06-25 เก็บ Flag ว่ามีการทำงานในสถานะ Online หรือ Offline (สถานะ Online/Offline   ว่าง:Online  1:Offline) ส่วนการ Connect Server
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        tSql = tSql & ",FTShdStaType = '1'"         'Online
    Else
        tSql = tSql & ",FTShdStaType = '0'"         'Offline
    End If
    '*KT 51-09-22 เก็บ Flag ว่ามีการทำงานในสถานะ Online หรือ Offline (สถานะ Online/Offline   ว่าง:Online  1:Offline) ส่วนการใช้ DB
    tSql = tSql & ",FTShdStaPrcGL='1'"         'Offline
    tSql = tSql & ",FTCcyCode ='1'"  '1: แบบ Detail  0: แบบก้อน
    tSql = tSql & ",FTXihDCStatusMALL = '" & tVB_OverSh & "' "      '0 = Menu 17 ,1 = Menu 18     '*Ao 2014/07/31     CommSheet TK-ISS3000-481.0.0.1   V 4.1.129
    
    tSql = tSql & vbCrLf & " WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND ISNULL(FTShdStaDoc,'0')  <> '1'"
    'Call DB_ExecuteSQL(tSql)
    Call SP_SQLvExecute(tSql)
    tVB_OverSh = "0"
    cVB_CoinsStartBank = 0
    cVB_CurrencyStartBank = 0
    cVB_TotalStartBank = SP_VALcText2Double(otbSumTender.Text)
End Sub

Private Sub W_PRCxSumAmount(ByVal pgdSource As VSFlexGrid, ByVal pogdTarget As VSFlexGrid)
    Dim nRow As Long
    Dim sum1 As Double
    Dim sum2 As Double
    sum1 = 0: sum2 = 0
    For nRow = 1 To pgdSource.rows - 1
        sum1 = sum1 + SP_VALcText2Double(pgdSource.TextMatrix(nRow, 1))
        sum2 = sum2 + SP_VALcText2Double(pgdSource.TextMatrix(nRow, 2))
    Next nRow
    pogdTarget.TextMatrix(0, 1) = Format(sum1, "#,##0.00")
    pogdTarget.TextMatrix(0, 2) = Format(sum2, "#,##0.00")
End Sub

'Private Sub Form_Activate()
'On Error GoTo ErrHandle
'    If bW_FirstLoad Then
'        Call SP_OBJxSetFocus(otbSumTender)
'    End If
'    bW_FirstLoad = False
'    Exit Sub
'ErrHandle:
'
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    Select Case KeyCode
        Case vbKeyEscape
            Call ocmCancel_Click
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnStartBank
            ocmCancel.Visible = False
            ocmOk.Left = ocmCancel.Left
    End Select
    Call W_FRMxSetInitial
    If Not oOpos Is Nothing Then
        If bVB_OPOSCashDrawerEnable Then Call oOpos.KickDrawer
    Else
         Call cUT_Print.UT_KickDrawer(Me)
    End If
    
    Call W_FNCxAddFunc2Grid
    Call W_FNCxAddFunc2GridForeign
    Call W_FNCxAddFunc2GridGV  '*BG 56-08-26 PRJ10001TK-DCR_5607007(285TR Deno Screen)-00.01.00.xls
    Call W_FNCxAddFunc2GridVendorVchr '*Tao 62-07-05 เพิ่ม Vendor Voucher
    'ogdFunc.SetFocus
    ogdFunc.SelectionMode = flexSelectionByRow
    ogdFunc.Select 0, 2
    'ogdFunc.EditCell
   
    
'   Me.ZOrder 0 '*KTs 53-02-03 ป้องกันการ Terminated หรือ  Form wBackGrand ขึ้นมาบังหน้าจอที่ต้องทำงาน
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Load")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    With oOpos
        '*KT 54-04-12 V.3.0.57 ป้องกันโปรแกรมหลุด
        'If .DrawerOpen And bVB_OPOSCashDrawerEnable Then wMsgDrawer.Show vbModal
        If bVB_OPOSCashDrawerEnable Then
            If .DrawerOpen Then
                wMsgDrawer.Show vbModal
            End If
        End If
    End With
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Unload")
End Sub

Private Sub ocmCancel_Click()
    Dim tSql$
On Error GoTo ErrHandle
    bVB_CNStartBankDT = False
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnStartBank:
            'process ตรวจการเปิดลิ้นชั  ก
            '*KT 51-06-23 Comsheet ML-SF-361
            Call ocmOK_Click
        Case tEN_TrnPickUp:
            Unload Me
            'wCNBegin.Show vbModal
        Case tEN_TrnDepositMoney
            Unload Me
            'wCNBegin.Show vbModal
        Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
            Unload Me
            'wCNBegin.Show vbModal
    End Select
    Exit Sub
ErrHandle:
    
End Sub

Private Sub ocmOK_Click()
On Error GoTo ErrHandle
'    Select Case SP_VALnText2Double(tVB_TransactionSub)
'        Case tEN_TrnOpen:
'        Case tEN_TrnStartBank
'            If Not (W_PRCbStartBank) Then Exit Sub
'        Case tEN_TrnPickUp: Call W_PRCxPickUp: Exit Sub
'        Case tEN_TrnDepositMoney: Call W_PRCxDepositMny: Exit Sub
'    End Select
'    Unload Me
'    Exit Sub
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOpen:
            Unload Me
            Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
        Case tEN_TrnStartBank
            If W_PRCbStartBank Then
                Unload Me
                Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
            Else
                Exit Sub
            End If
        Case tEN_TrnPickUp:
            Call W_PRCxPickUp
            Exit Sub
        Case tEN_TrnDepositMoney:
            Call W_PRCxDepositMny
            Exit Sub
        Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
            Call W_PRCxOverShortReport
            bVB_CNOverShortSta = True
            Exit Sub
    End Select
    Exit Sub
ErrHandle:
    
End Sub

'Private Sub ocmPrevious_Click()
'    Select Case otaForm.SelectedItem
'        Case 1
'            otaForm.Item(0).Visible = True
'            otaForm.Item(0).Selected = True
'            otaForm.Item(1).Visible = False
'            otaForm.Item(2).Visible = False
'            ocmPrevious.Visible = False
'            ogdCoin.Row = 1
'            ogdCoin.Col = 1
'            ogdCoin.EditCell
'        Case 2
'            otaForm.Item(0).Visible = False
'            otaForm.Item(1).Visible = True
'            otaForm.Item(1).Selected = True
'            otaForm.Item(2).Visible = False
'            ocmNext.ZOrder 0
'            ogdCurrentcy.Row = 1
'            ogdCurrentcy.Col = 1
'            ogdCurrentcy.EditCell
'    End Select
'End Sub



Private Sub ogdFunc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_KeyPress = ""
'    ogdFunc.HighLight = flexHighlightAlways
'    ogdFunc.Select NewRow, 2        'move highlight to first row
'    ogdFunc.Col = 2
'    DoEvents 'ห้ามเอาออกไม่นั้นจะทำให้ Grid ไม่มีการ Select Row '*KT 51-10-09 V 3.0.131/V 3.0.200
'    ogdFunc.EditCell
End Sub

Private Sub ogdFunc_ChangeEdit()
    With ogdFunc
        Select Case .Col
            Case 2
                'MsgBox tW_KeyPress
                'MsgBox SP_DEFtChangeNum(tW_KeyPress)
                '.EditText = SP_DEFtChangeNum(tW_KeyPress)
                .EditText = tW_KeyPress
        End Select
    End With
End Sub

'NEUNG XXX
'Private Sub ogdFunc_DblClick()
'    Call W_FNCxAddAmt
'End Sub

Private Sub ogdFunc_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With ogdFunc
        Select Case Col
            Case 2
                If .TextMatrix(Row, Col) = "" Then .TextMatrix(Row, Col) = 0
                .TextMatrix(Row, 3) = SP_VALcText2Double(.TextMatrix(Row, 2)) * SP_VALcText2Double(.TextMatrix(Row, 4))
                Call W_DEFxSumStartBank
        End Select
    End With
End Sub

Private Sub ogdFunc_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With ogdFunc
        Select Case Col
            Case 2
                '.FocusRect = flexFocusHeavy
                '.HighLight = flexHighlightAlways
                .EditMaxLength = 5 '*Eaw 56-04-03
            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub ogdFunc_GotFocus()
    With ogdFunc
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
         tW_KeyPress = "" 'Replace(.TextMatrix(.Row, .Col), ".", "")
    End With
End Sub

'Private Sub ogdFunc_EnterCell()
'    With ogdFunc
'        If .Col = 2 Then
'            If .Row = .rows - 1 Then
'                ogdFunc1.SetFocus
'                ogdFunc1.Select 0, 2
'            Else
'                ogdFunc.SetFocus
'                ogdFunc.Select ogdFunc.Row + 1, 2
'            End If
'        End If
'    End With
'End Sub

Private Sub ogdFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    With ogdFunc
        Select Case KeyCode
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                    '*BG 56-08-28 Comsheet-285
'                        ogdGiftVoucher.SetFocus
'                        ogdGiftVoucher.Select 0, 3
                        ogdOtherCash.SetFocus
                        ogdOtherCash.Select 0, 3
                    '***
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc.SetFocus
                        ogdFunc.Select .Row + 1, 2
                    End If
                End If
            Case vbKeyDown
                If .Row = .rows - 1 Then
                '*BG 56-08-28 comsheet-285
'                    ogdGiftVoucher.SetFocus
'                    ogdGiftVoucher.Select 0, 3
                    ogdOtherCash.SetFocus
                    ogdOtherCash.Select 0, 3
                '****
                End If
            Case vbKeyUp
                If .Row = 0 Then
                '*BG 56-09-02
'                    ogdFunc1.SetFocus
'                    ogdFunc1.Select ogdFunc1.rows - 1, 3
                    ogdOtherGV.SetFocus
                    ogdOtherGV.Select ogdOtherGV.rows - 1, 3
                '****
                End If
        End Select
    End With
End Sub

Private Sub ogdFunc_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With ogdFunc
        Select Case KeyAscii
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    'MsgBox .EditText
                    If .EditText <> "" Then
                        'Call ogdFunc_KeyDown(vbKeyDown, 0)
                        '*BG 56-08-28 comsheet-285
'                        ogdGiftVoucher.SetFocus
'                        ogdGiftVoucher.Select 0, 3
                        '****
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc.SetFocus
                        ogdFunc.Select Row + 1, 2
                    End If
                End If
            Case Else
                Select Case KeyAscii
                    Case vbKeyDelete
                        tW_KeyPress = ""
                    Case vbKeyBack
                    '*Eaw 56-04-22  CommSheet TK-ISS3000-154-IT2-Eaw.xlsx
'                        If Len(Trim(tW_KeyPress)) > 0 Then
'                            tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
'                        End If
                        If Len(Trim(tW_KeyPress)) > 0 Then
                            If Len(Trim(tW_KeyPress)) > 5 Then
                                tW_KeyPress = Left(tW_KeyPress, 4)
                            Else
                                tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
                            End If
                        End If
                        With ogdFunc
                            Select Case .Col
                                Case 2
                                    .Text = tW_KeyPress
                            End Select
                        End With
                    '------------------------------------------------------------------------------------------
                    Case Else
                        Call cUT.UT_KBDKeyCurrency(KeyAscii, tW_KeyPress)    'check key, currency
                        If Chr(KeyAscii) = "." Then KeyAscii = 0
                        If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii)) 'ถ้า  KeyAscii = Enter หมายถึงจบการแก้ไขราคาต่อหน่วย
                End Select
        End Select
    End With
End Sub

Private Sub ogdFunc_LostFocus()
    With ogdFunc
        .SelectionMode = flexSelectionFree
        .Editable = flexEDNone
        '.Cell(flexcpBackColor, 0, 2, .rows - 1, 2) = &H80000013
        .HighLight = flexHighlightNever
    End With
End Sub

Private Sub ogdGiftVoucher_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
     tW_KeyPress = ""
End Sub

Private Sub ogdGiftVoucher_ChangeEdit()
    With ogdGiftVoucher
        Select Case .Col
            Case 3
                'MsgBox tW_KeyPress
                .EditText = SP_DEFtChangeNum(tW_KeyPress)
        End Select
    End With
End Sub

Private Sub ogdGiftVoucher_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With ogdGiftVoucher
        Select Case Col
            Case 3
                If .TextMatrix(Row, Col) = "" Then .TextMatrix(Row, Col) = Format(0, "#0.00000")
               ' .TextMatrix(Row, 2) = SP_FMTtNum2ShwCurrency(SP_VALcText2Double(.TextMatrix(Row, 2)) * SP_VALcText2Double(.TextMatrix(Row, 6)))
                Call W_DEFxSumStartBank
        End Select
    End With
End Sub

Private Sub ogdGiftVoucher_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With ogdGiftVoucher
        Select Case Col
            Case 3
                '.FocusRect = flexFocusHeavy
                .EditMaxLength = 11 '*Eaw 56-04-03
            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub ogdGiftVoucher_GotFocus()
    With ogdGiftVoucher
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
         tW_KeyPress = "" 'Replace(.TextMatrix(.Row, .Col), ".", "")
    End With
End Sub

Private Sub ogdGiftVoucher_KeyDown(KeyCode As Integer, Shift As Integer)
    With ogdGiftVoucher
        Select Case KeyCode
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select 0, 3
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc.SetFocus
                        ogdFunc.Select .Row + 1, 2
                    End If
                End If
            Case vbKeyDown
               If .Row = .rows - 1 Then
                    ogdFunc1.SetFocus
                    ogdFunc1.Select 0, 3
                End If
            Case vbKeyUp
                If .Row = 0 Then
                    ogdFunc.SetFocus
                    ogdFunc.Select ogdFunc.rows - 1, 2
                End If
        End Select
    End With
End Sub

Private Sub ogdGiftVoucher_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With ogdGiftVoucher
        Select Case KeyAscii
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select 0, 3
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select Row + 1, 3
                    End If
                End If
            Case Else
                Select Case KeyAscii
                    Case vbKeyDelete
                        tW_KeyPress = ""
                    Case vbKeyBack
                    '*Eaw 56-04-22  CommSheet TK-ISS3000-154-IT2-Eaw.xlsx
'                        If Len(Trim(tW_KeyPress)) > 0 Then
'                            tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
'                        End If
                        If Len(Trim(tW_KeyPress)) > 0 Then
                            If Len(Trim(tW_KeyPress)) > 11 Then
                                tW_KeyPress = Left(tW_KeyPress, 9)
                            Else
                                tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
                            End If
                        End If
                        With ogdGiftVoucher
                            Select Case .Col
                                Case 2
                                    .Text = tW_KeyPress
                            End Select
                        End With
                    '------------------------------------------------------------------------------------------
                    Case Else
                        Call cUT.UT_KBDKeyCurrency(KeyAscii, tW_KeyPress)    'check key, currency
                        If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii)) 'ถ้า  KeyAscii = Enter หมายถึงจบการแก้ไขราคาต่อหน่วย
                End Select
        End Select
    End With
End Sub

Private Sub ogdGiftVoucher_LostFocus()
    With ogdGiftVoucher
        .SelectionMode = flexSelectionFree
        .Editable = flexEDNone
        '.Cell(flexcpBackColor, 0, 3, .rows - 1, 3) = &H80000013
        .HighLight = flexHighlightNever
    End With
End Sub

Private Sub ogdFunc1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
     tW_KeyPress = ""
'    ogdFunc1.HighLight = flexHighlightAlways
'    ogdFunc1.Select NewRow, 2        'move highlight to first row
'    ogdFunc1.Col = 2
'    DoEvents 'ห้ามเอาออกไม่นั้นจะทำให้ Grid ไม่มีการ Select Row '*KT 51-10-09 V 3.0.131/V 3.0.200
'    ogdFunc1.EditCell
End Sub

Private Sub ogdFunc1_ChangeEdit()
    With ogdFunc1
        Select Case .Col
            Case 3
                'MsgBox tW_KeyPress
                .EditText = SP_DEFtChangeNum(tW_KeyPress)
                
        End Select
    End With
End Sub

'Private Sub ogdFunc1_DblClick()
'    Call W_FNCxAddForCurrencyAmt
'End Sub

Private Sub ogdFunc1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With ogdFunc1
        Select Case Col
            Case 3
                If .TextMatrix(Row, Col) = "" Then .TextMatrix(Row, Col) = Format(0, "#0.00000")
               ' .TextMatrix(Row,2) = SP_FMTtNum2ShwCurrency(SP_VALcText2Double(.TextMatrix(Row, 2)) * SP_VALcText2Double(.TextMatrix(Row, 6)))
                Call W_DEFxSumStartBank
        End Select
    End With
End Sub

Private Sub ogdFunc1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With ogdFunc1
        Select Case Col
            Case 3
                '.FocusRect = flexFocusHeavy
                .EditMaxLength = 11  '*Eaw 56-04-03
            Case Else
                Cancel = True
        End Select
    End With
End Sub

'Private Sub ogdFunc1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
'        ogdFunc1.Cell(flexcpBackColor, NewRow, 2, NewRow, 2) = &H80000013
'End Sub

'Private Sub ogdFunc1_EnterCell()
'    With ogdFunc1
'        If .Col = 2 Then
'            If .Row = .rows - 1 Then
'                   Call SP_OBJxSetFocus(ocmOK)
'            End If
'        End If
'    End With
'End Sub

Private Sub ogdFunc1_GotFocus()
    With ogdFunc1
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
         tW_KeyPress = "" 'Replace(.TextMatrix(.Row, .Col), ".", "")
    End With
End Sub

Private Sub ogdFunc1_KeyDown(KeyCode As Integer, Shift As Integer)
    With ogdFunc1
        Select Case KeyCode
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc.SetFocus
                        ogdFunc.Select 0, 2
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select .Row + 1, 3
                    End If
                End If
            Case vbKeyDown
               If .Row = .rows - 1 Then
               '*BG 56-09-02
'                    ogdFunc.SetFocus
'                    ogdFunc.Select 0, 3
                    ogdGVFunc.SetFocus
                    ogdGVFunc.Select 0, 2
                    '****
                End If
            Case vbKeyUp
                If .Row = 0 Then
                '*BG 56-09-02
'                    ogdGiftVoucher.SetFocus
'                    ogdGiftVoucher.Select ogdGiftVoucher.rows - 1, 3
                    ogdOtherCash.SetFocus
                    ogdOtherCash.Select ogdOtherCash.rows - 1, 2
                    '***
                End If
        End Select
    End With
End Sub

Private Sub ogdFunc1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With ogdFunc1
        Select Case KeyAscii
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc.SetFocus
                        ogdFunc.Select 0, 2
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select Row + 1, 3
                    End If
                End If
            Case Else
                Select Case KeyAscii
                    Case vbKeyDelete
                        tW_KeyPress = ""
                    Case vbKeyBack
                    '*Eaw 56-04-22  CommSheet TK-ISS3000-154-IT2-Eaw.xlsx
'                        If Len(Trim(tW_KeyPress)) > 0 Then
'                            tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
'                        End If
                        If Len(Trim(tW_KeyPress)) > 0 Then
                            If Len(Trim(tW_KeyPress)) > 11 Then
                                tW_KeyPress = Left(tW_KeyPress, 9)
                            Else
                                tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
                            End If
                        End If
                        With ogdFunc1
                            Select Case .Col
                                Case 2
                                    .Text = tW_KeyPress
                            End Select
                        End With
                    '------------------------------------------------------------------------------------------
                    Case Else
                        Call cUT.UT_KBDKeyCurrency(KeyAscii, tW_KeyPress)    'check key, currency
                        If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii)) 'ถ้า  KeyAscii = Enter หมายถึงจบการแก้ไขราคาต่อหน่วย
                End Select
        End Select
    End With
End Sub

Private Sub ogdFunc1_LostFocus()
    With ogdFunc1
        .SelectionMode = flexSelectionFree
        .Editable = flexEDNone
        '.Cell(flexcpBackColor, 0, 3, .rows - 1, 3) = &H80000013
        .HighLight = flexHighlightNever
    End With
End Sub

Private Sub ogdGVFunc_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With ogdGVFunc
        Select Case Col
            Case 2
                If .TextMatrix(Row, Col) = "" Then .TextMatrix(Row, Col) = 0
                .TextMatrix(Row, 3) = SP_VALcText2Double(.TextMatrix(Row, 2)) * SP_VALcText2Double(.TextMatrix(Row, 4))
                Call W_DEFxSumStartBank
        End Select
    End With
End Sub

Private Sub ogdGVFunc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_KeyPress = ""
End Sub

Private Sub ogdGVFunc_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With ogdGVFunc
        Select Case Col
            Case 2
                '.FocusRect = flexFocusHeavy
                '.HighLight = flexHighlightAlways
                .EditMaxLength = 5 '*Eaw 56-04-03
            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub ogdGVFunc_ChangeEdit()
    With ogdGVFunc
        Select Case .Col
            Case 2
                'MsgBox tW_KeyPress
                'MsgBox SP_DEFtChangeNum(tW_KeyPress)
                '.EditText = SP_DEFtChangeNum(tW_KeyPress)
                .EditText = tW_KeyPress
        End Select
    End With

End Sub

Private Sub ogdGVFunc_GotFocus()
    With ogdGVFunc
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
         tW_KeyPress = "" 'Replace(.TextMatrix(.Row, .Col), ".", "")
    End With
End Sub

Private Sub ogdGVFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    With ogdGVFunc
        Select Case KeyCode
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdOtherGV.SetFocus
                        ogdOtherGV.Select 0, 3
                    End If
                Else
                    If .EditText <> "" Then
                        ogdGVFunc.SetFocus
                        ogdGVFunc.Select .Row + 1, 2
                    End If
                End If
            Case vbKeyDown
                If .Row = .rows - 1 Then
                    ogdOtherGV.SetFocus
                    ogdOtherGV.Select 0, 3
                End If
            Case vbKeyUp
                If .Row = 0 Then
                    ogdFunc1.SetFocus
                    ogdFunc1.Select ogdFunc1.rows - 1, 3
                End If
        End Select
    End With
End Sub

Private Sub ogdGVFunc_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With ogdGVFunc
        Select Case KeyAscii
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    'MsgBox .EditText
                    If .EditText <> "" Then
                        'Call ogdFunc_KeyDown(vbKeyDown, 0)
                        ogdOtherGV.SetFocus
                        ogdOtherGV.Select 0, 3
                    End If
                Else
                    If .EditText <> "" Then
                        ogdGVFunc.SetFocus
                        ogdGVFunc.Select Row + 1, 2
                    End If
                End If
            Case Else
                Select Case KeyAscii
                    Case vbKeyDelete
                        tW_KeyPress = ""
                    Case vbKeyBack
                    '*Eaw 56-04-22  CommSheet TK-ISS3000-154-IT2-Eaw.xlsx
'                        If Len(Trim(tW_KeyPress)) > 0 Then
'                            tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
'                        End If
                        If Len(Trim(tW_KeyPress)) > 0 Then
                            If Len(Trim(tW_KeyPress)) > 5 Then
                                tW_KeyPress = Left(tW_KeyPress, 4)
                            Else
                                tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
                            End If
                        End If
                        With ogdGVFunc
                            Select Case .Col
                                Case 2
                                    .Text = tW_KeyPress
                            End Select
                        End With
                    '------------------------------------------------------------------------------------------
                    Case Else
                        Call cUT.UT_KBDKeyCurrency(KeyAscii, tW_KeyPress)    'check key, currency
                        If Chr(KeyAscii) = "." Then KeyAscii = 0
                        If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii)) 'ถ้า  KeyAscii = Enter หมายถึงจบการแก้ไขราคาต่อหน่วย
                End Select
        End Select
    End With

End Sub

Private Sub ogdGVFunc_LostFocus()
    With ogdGVFunc
        .SelectionMode = flexSelectionFree
        .Editable = flexEDNone
        '.Cell(flexcpBackColor, 0, 2, .rows - 1, 2) = &H80000013
        .HighLight = flexHighlightNever
    End With
End Sub

Private Sub ogdOtherGV_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_KeyPress = ""
End Sub

Private Sub ogdOtherGV_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With ogdOtherGV
        Select Case Col
            Case 3
                '.FocusRect = flexFocusHeavy
                .EditMaxLength = 11 '*Eaw 56-04-03
            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub ogdOtherGV_ChangeEdit()
    With ogdOtherGV
        Select Case .Col
            Case 3
                'MsgBox tW_KeyPress
                .EditText = SP_DEFtChangeNum(tW_KeyPress)
        End Select
    End With
End Sub

Private Sub ogdOtherGV_GotFocus()
    With ogdOtherGV
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
         tW_KeyPress = "" 'Replace(.TextMatrix(.Row, .Col), ".", "")
    End With
End Sub

Private Sub ogdOtherGV_KeyDown(KeyCode As Integer, Shift As Integer)
    With ogdOtherGV
        Select Case KeyCode
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select 0, 3
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc.SetFocus
                        ogdFunc.Select .Row + 1, 2
                    End If
                End If
            Case vbKeyDown
               If .Row = .rows - 1 Then
               '*BG 56-09-02
'                    ogdFunc1.SetFocus
'                    ogdFunc1.Select 0, 3
                    ogdFunc.SetFocus
                    ogdFunc.Select 0, 3
                '****
                End If
            Case vbKeyUp
                If .Row = 0 Then
                '*BG 56-09-02
'                    ogdFunc.SetFocus
'                    ogdFunc.Select ogdFunc.rows - 1, 2
                    ogdGVFunc.SetFocus
                    ogdGVFunc.Select ogdGVFunc.rows - 1, 2
                '****
                End If
        End Select
    End With

End Sub

Private Sub ogdOtherGV_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With ogdOtherGV
        Select Case KeyAscii
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select 0, 3
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select Row + 1, 3
                    End If
                End If
            Case Else
                Select Case KeyAscii
                    Case vbKeyDelete
                        tW_KeyPress = ""
                    Case vbKeyBack
                    '*Eaw 56-04-22  CommSheet TK-ISS3000-154-IT2-Eaw.xlsx
'                        If Len(Trim(tW_KeyPress)) > 0 Then
'                            tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
'                        End If
                        If Len(Trim(tW_KeyPress)) > 0 Then
                            If Len(Trim(tW_KeyPress)) > 11 Then
                                tW_KeyPress = Left(tW_KeyPress, 9)
                            Else
                                tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
                            End If
                        End If
                        With ogdOtherGV
                            Select Case .Col
                                Case 2
                                    .Text = tW_KeyPress
                            End Select
                        End With
                    '------------------------------------------------------------------------------------------
                    Case Else
                        Call cUT.UT_KBDKeyCurrency(KeyAscii, tW_KeyPress)    'check key, currency
                        If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii)) 'ถ้า  KeyAscii = Enter หมายถึงจบการแก้ไขราคาต่อหน่วย
                End Select
        End Select
    End With
End Sub

Private Sub ogdOtherGV_LostFocus()
    With ogdOtherGV
        .SelectionMode = flexSelectionFree
        .Editable = flexEDNone
        '.Cell(flexcpBackColor, 0, 3, .rows - 1, 3) = &H80000013
        .HighLight = flexHighlightNever
    End With
End Sub

Private Sub ogdOtherCash_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With ogdOtherCash
        Select Case Col
            Case 3
'                If .TextMatrix(Row, Col) = "" Then .TextMatrix(Row, Col) = 0
'                .TextMatrix(Row, 3) = SP_VALcText2Double(.TextMatrix(Row, 2)) * SP_VALcText2Double(.TextMatrix(Row, 4))
                If .TextMatrix(Row, Col) = "" Then .TextMatrix(Row, Col) = Format(0, "#0.00000")
                Call W_DEFxSumStartBank
'                ogdTotalGV.TextMatrix(0, 2) = ogdTotalGV.ValueMatrix(0, 2) + ogdOtherGV.ValueMatrix(0, 3)
        End Select
    End With
End Sub

Private Sub ogdOtherCash_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_KeyPress = ""
End Sub

Private Sub ogdOtherCash_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With ogdOtherCash
        Select Case Col
            Case 3
                '.FocusRect = flexFocusHeavy
                .EditMaxLength = 11 '*Eaw 56-04-03
            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub ogdOtherCash_ChangeEdit()
    With ogdOtherCash
        Select Case .Col
            Case 3
                'MsgBox tW_KeyPress
                .EditText = SP_DEFtChangeNum(tW_KeyPress)
        End Select
    End With
End Sub

Private Sub ogdOtherCash_GotFocus()
    With ogdOtherCash
        .HighLight = flexHighlightAlways
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
         tW_KeyPress = "" 'Replace(.TextMatrix(.Row, .Col), ".", "")
    End With
End Sub

Private Sub ogdOtherCash_KeyDown(KeyCode As Integer, Shift As Integer)
    With ogdOtherCash
        Select Case KeyCode
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select 0, 3
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc.SetFocus
                        ogdFunc.Select .Row + 1, 2
                    End If
                End If
            Case vbKeyDown
               If .Row = .rows - 1 Then
                    ogdFunc1.SetFocus
                    ogdFunc1.Select 0, 3
                End If
            Case vbKeyUp
                If .Row = 0 Then
                    ogdFunc.SetFocus
                    ogdFunc.Select ogdFunc.rows - 1, 2
                End If
        End Select
    End With
End Sub

Private Sub ogdOtherCash_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With ogdOtherCash
        Select Case KeyAscii
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select 0, 3
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select Row + 1, 3
                    End If
                End If
            Case Else
                Select Case KeyAscii
                    Case vbKeyDelete
                        tW_KeyPress = ""
                    Case vbKeyBack
                    '*Eaw 56-04-22  CommSheet TK-ISS3000-154-IT2-Eaw.xlsx
'                        If Len(Trim(tW_KeyPress)) > 0 Then
'                            tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
'                        End If
                        If Len(Trim(tW_KeyPress)) > 0 Then
                            If Len(Trim(tW_KeyPress)) > 11 Then
                                tW_KeyPress = Left(tW_KeyPress, 9)
                            Else
                                tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
                            End If
                        End If
                        With ogdOtherCash
                            Select Case .Col
                                Case 2
                                    .Text = tW_KeyPress
                            End Select
                        End With
                    '------------------------------------------------------------------------------------------
                    Case Else
                        Call cUT.UT_KBDKeyCurrency(KeyAscii, tW_KeyPress)    'check key, currency
                        If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii)) 'ถ้า  KeyAscii = Enter หมายถึงจบการแก้ไขราคาต่อหน่วย
                End Select
        End Select
    End With
End Sub

Private Sub ogdOtherCash_LostFocus()
    With ogdOtherCash
        .SelectionMode = flexSelectionFree
        .Editable = flexEDNone
        '.Cell(flexcpBackColor, 0, 3, .rows - 1, 3) = &H80000013
        .HighLight = flexHighlightNever
    End With
End Sub

Private Sub ogdOtherGV_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With ogdOtherGV
        Select Case Col
            Case 3
'                If .TextMatrix(Row, Col) = "" Then .TextMatrix(Row, Col) = 0
'                .TextMatrix(Row, 3) = SP_VALcText2Double(.TextMatrix(Row, 2)) * SP_VALcText2Double(.TextMatrix(Row, 4))
                Call W_DEFxSumStartBank
'                ogdTotalGV.TextMatrix(0, 2) = ogdTotalGV.ValueMatrix(0, 2) + ogdOtherGV.ValueMatrix(0, 3)
        End Select
    End With
End Sub

Private Sub ogdVendorVchr_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With ogdVendorVchr
        Select Case Col
            Case 3
                Call W_DEFxSumStartBank
        End Select
    End With
End Sub

Private Sub otbSumTender_Change()       '*AO51-03-12
    otbSumTender.Text = SP_DEFtChangeNum(Trim(otbSumTender.Text))
    otbSumTender.SelStart = Len(otbSumTender.Text)
End Sub

Private Sub otbSumTender_GotFocus()
        otbSumTender.SelStart = 0
        otbSumTender.SelLength = Len(otbSumTender.Text)
End Sub

Private Sub otbSumTender_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbSumTender.Text = ""
    End Select
End Sub

Private Sub otbSumTender_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbSumTender.Text), ".") > 0 Then
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

Public Sub W_INSxTranRC()
'------------------------------------------------------------------
'Call:
'Cmt: *KT 51-08-18 ทำการบันทึกข้อมูล RC ของ Start Bank
'------------------------------------------------------------------
    Dim tSql$
    Dim cFAmt As Double
    Dim cVB_Pay As Double
    Dim cNet As Double
    Dim cVB_Change As Double
    Dim nI As Long
    Dim nSeqNo As Long
On Error GoTo ErrHandle

    cVB_Change = 0
    cFAmt = CDbl(Val(Format(otbSumCash.Text)))
    cVB_Pay = cFAmt
    cNet = cFAmt
    tVB_CrdCardNo = ""
    cVB_FCSrcRteAmt = 0
    cVB_FCSrcRteFac = 0
    cVB_RoundChg = 0
    tVB_CrdApprCode = ""
    nUT_SeqRC = 1
    tVB_FTRteCode = ""
'    tVB_CrdExpM = "": tVB_CrdExpY = ""
'    tVB_TenderCode = aVB_TenderMnu(0).ID
'    tVB_TenderType = aVB_TenderMnu(0).Name '*KT 53-10-15
'    tVB_TenderName = aVB_TenderMnu(0).DescSale
    For nI = 0 To UBound(aVB_TenderMnu)                                                             '*Ao 2013/05/31
        If UCase(aVB_TenderMnu(nI).ID) = UCase("T001") Then
            nSeqNo = nI
            Exit For
        End If
    Next nI
    tVB_TenderCode = aVB_TenderMnu(nSeqNo).ID
    tVB_TenderType = aVB_TenderMnu(nSeqNo).Name
    tVB_TenderName = aVB_TenderMnu(nSeqNo).DescSale
    
    tSql = "DELETE " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
    Call SP_SQLvExecute(tSql)
    
    'save RC Cash
    tSql = "INSERT INTO " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "(FTTmnNum,FTShdTransNo,FNSrcSeqNo,"
    tSql = tSql & vbCrLf & "FTShdTransType,FDShdTransDate,FTTdmCode,"
    tSql = tSql & vbCrLf & "FTTdmType,FTSrcCardNo,FTApprCode,"
    tSql = tSql & vbCrLf & "FCSrcCardChg,FCSrcFAmt,FCSrcAmt,"
    tSql = tSql & vbCrLf & "FCSrcNet,FCSrcChg,FTSrcStaPrc,"
    tSql = tSql & vbCrLf & "FTSrcRmk,FTSrcRetDocRef,FCSrcRetAmt,"
    tSql = tSql & vbCrLf & "FNShdStaSend,"
    tSql = tSql & vbCrLf & "FTRteCode,"
    tSql = tSql & vbCrLf & "FCSrcRteAmt ,"
    tSql = tSql & vbCrLf & "FCSrcRteFac,"
    tSql = tSql & vbCrLf & "FCSrcRndMnyChg,"
    tSql = tSql & vbCrLf & "FTSrcAbb,"
    tSql = tSql & vbCrLf & "FTSrcCardExp,"
    tSql = tSql & vbCrLf & "FCSrcOntopAmt,"
    tSql = tSql & vbCrLf & "FNSrcCrdPoint,"
    tSql = tSql & vbCrLf & "FTREFNO,"  '*Eaw 56-11-05
    tSql = tSql & vbCrLf & tCS_CNLastFedIns
    tSql = tSql & vbCrLf & " VALUES('" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nUT_SeqRC & ","
    tSql = tSql & vbCrLf & "'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & tVB_TenderCode & "',"
    tSql = tSql & vbCrLf & "'" & tVB_TenderType & "','" & tVB_CrdCardNo & "','" & tVB_CrdApprCode & "',"
    tSql = tSql & vbCrLf & 0 & "," & SP_FMTcCurPoint(cFAmt, True, False) & "," & SP_FMTcCurPoint(cVB_Pay, True, False) & ","
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cNet, True, False) & "," & SP_FMTcCurPoint(cVB_Change, True, False) & ",'',"
    tSql = tSql & vbCrLf & "'',''," & 0 & ","
    tSql = tSql & vbCrLf & "0,"
    tSql = tSql & vbCrLf & "'" & tVB_FTRteCode & "',"
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteAmt, True, False) & ","
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteFac, True, False) & ","
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_RoundChg, True, False) & ","
    'tSql = tSql & vbCrLf & "'" & tVB_TenderName & "',"
    tSql = tSql & vbCrLf & "'" & "Cash" & "',"
    If tVB_CrdExpM = "" And tVB_CrdExpY = "" Then
        tSql = tSql & "'" & "" & "',"
    Else
        tSql = tSql & "'" & tVB_CrdExpM & "/" & tVB_CrdExpY & "',"
    End If
    tSql = tSql & 0 & ","
    tSql = tSql & 0 & ","
    tSql = tSql & vbCrLf & "'" & tVB_TenderType & "',"  '*Eaw 56-11-05
    tSql = tSql & UT_tLastInsSql & ")"
    Call SP_SQLvExecute(tSql)
    
    'Gift Voucher
    cFAmt = CDbl(Val(Format(otbSumGiftVoucher.Text)))
    cVB_Pay = cFAmt
    cNet = cFAmt
    tVB_CrdCardNo = ""
    cVB_FCSrcRteAmt = 0
    cVB_FCSrcRteFac = 0
    cVB_RoundChg = 0
    tVB_CrdApprCode = ""
    nUT_SeqRC = 2
    tVB_FTRteCode = ""
'    tVB_TenderCode = aVB_TenderMnu(8).ID
'    tVB_TenderType = aVB_TenderMnu(8).Name '*KT 53-10-15
'    tVB_TenderName = aVB_TenderMnu(8).DescSale
    For nI = 0 To UBound(aVB_TenderMnu)
        If UCase(aVB_TenderMnu(nI).ID) = UCase("T005") Then
            nSeqNo = nI
            Exit For
        End If
    Next nI
    tVB_TenderCode = aVB_TenderMnu(nSeqNo).ID
    tVB_TenderType = aVB_TenderMnu(nSeqNo).Name
    tVB_TenderName = aVB_TenderMnu(nSeqNo).DescSale
            
    tSql = "INSERT INTO " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "(FTTmnNum,FTShdTransNo,FNSrcSeqNo,"
    tSql = tSql & vbCrLf & "FTShdTransType,FDShdTransDate,FTTdmCode,"
    tSql = tSql & vbCrLf & "FTTdmType,FTSrcCardNo,FTApprCode,"
    tSql = tSql & vbCrLf & "FCSrcCardChg,FCSrcFAmt,FCSrcAmt,"
    tSql = tSql & vbCrLf & "FCSrcNet,FCSrcChg,FTSrcStaPrc,"
    tSql = tSql & vbCrLf & "FTSrcRmk,FTSrcRetDocRef,FCSrcRetAmt,"
    tSql = tSql & vbCrLf & "FNShdStaSend,"
    tSql = tSql & vbCrLf & "FTRteCode,"
    tSql = tSql & vbCrLf & "FCSrcRteAmt ,"
    tSql = tSql & vbCrLf & "FCSrcRteFac,"
    tSql = tSql & vbCrLf & "FCSrcRndMnyChg,"
    tSql = tSql & vbCrLf & "FTSrcAbb,"
    tSql = tSql & vbCrLf & "FTSrcCardExp,"
    tSql = tSql & vbCrLf & "FCSrcOntopAmt,"
    tSql = tSql & vbCrLf & "FNSrcCrdPoint,"
    tSql = tSql & vbCrLf & "FTREFNO,"  '*Eaw 56-11-05
    tSql = tSql & vbCrLf & tCS_CNLastFedIns
    tSql = tSql & vbCrLf & " VALUES('" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nUT_SeqRC & ","
    tSql = tSql & vbCrLf & "'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & tVB_TenderCode & "',"
    tSql = tSql & vbCrLf & "'" & tVB_TenderType & "','" & tVB_CrdCardNo & "','" & tVB_CrdApprCode & "',"
    tSql = tSql & vbCrLf & 0 & "," & SP_FMTcCurPoint(cFAmt, True, False) & "," & SP_FMTcCurPoint(cVB_Pay, True, False) & ","
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cNet, True, False) & "," & SP_FMTcCurPoint(cVB_Change, True, False) & ",'',"
    tSql = tSql & vbCrLf & "'',''," & 0 & ","
    tSql = tSql & vbCrLf & "0,"
    tSql = tSql & vbCrLf & "'" & tVB_FTRteCode & "',"
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteAmt, True, False) & ","
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteFac, True, False) & ","
    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_RoundChg, True, False) & ","
    'tSql = tSql & vbCrLf & "'" & tVB_TenderName & "',"
    tSql = tSql & vbCrLf & "'" & "GVchr" & "',"
    If tVB_CrdExpM = "" And tVB_CrdExpY = "" Then
        tSql = tSql & "'" & "" & "',"
    Else
        tSql = tSql & "'" & tVB_CrdExpM & "/" & tVB_CrdExpY & "',"
    End If
    tSql = tSql & 0 & ","
    tSql = tSql & 0 & ","
    tSql = tSql & vbCrLf & "'" & tVB_TenderType & "',"  '*Eaw 56-11-05
    tSql = tSql & UT_tLastInsSql & ")"
    Call SP_SQLvExecute(tSql)
    
     'save RC Foreign Currency
    With ogdFunc1
'            tVB_TenderCode = aVB_TenderMnu(6).ID
'            tVB_TenderType = aVB_TenderMnu(6).Name '*KT 53-10-15
'            tVB_TenderName = aVB_TenderMnu(6).DescSale
            For nI = 0 To UBound(aVB_TenderMnu)
                If UCase(aVB_TenderMnu(nI).ID) = UCase("T012") Then
                    nSeqNo = nI
                    Exit For
                End If
            Next nI
            tVB_TenderCode = aVB_TenderMnu(nSeqNo).ID
            tVB_TenderType = aVB_TenderMnu(nSeqNo).Name
            tVB_TenderName = aVB_TenderMnu(nSeqNo).DescSale
            
             For nI = .FixedRows To .rows - 1
                    nUT_SeqRC = nUT_SeqRC + 1
'                    cFAmt = SP_FMTcCurPoint(CDbl(Val(Format(.TextMatrix(nI, 7)))) * CDbl(Val(Format(.TextMatrix(nI, 3)))), True, False)        '*Ao 2013/09/18
                    cFAmt = SP_FMTcCurPoint(CDbl(Val(Format(.TextMatrix(nI, 7)))) * CDbl(Val(Format(.TextMatrix(nI, 3)))), True, False)
                    If cFAmt = 0 Then
                        If CDbl(Val(Format(.ValueMatrix(nI, 3)))) > 0 Then
                            cFAmt = CDbl(Val(Format(.ValueMatrix(nI, 7)))) * CDbl(Val(Format(.ValueMatrix(nI, 3))))
                        End If
                    End If
                    cVB_Pay = cFAmt
                    cNet = cFAmt
                    tVB_CrdCardNo = ""
                    cVB_FCSrcRteAmt = SP_FMTcCurPoint(CDbl(Val(Format(.TextMatrix(nI, 3)))), True, False)
'                    cVB_FCSrcRteFac = SP_FMTcCurPoint(CDbl(Val(Format(.TextMatrix(nI, 7)))), True, False)
                    cVB_FCSrcRteFac = Format(CDbl(Val(Format(.TextMatrix(nI, 7)))), "###0.00000")               '*Ao 2013/12/04  CommSheet TK-ISS3000-440
                    cVB_RoundChg = 0
                    tVB_CrdApprCode = ""
                    tVB_FTRteCode = .TextMatrix(nI, 5)
                    
                    tSql = "INSERT INTO " & tVB_CNTblRC
                    tSql = tSql & vbCrLf & "(FTTmnNum,FTShdTransNo,FNSrcSeqNo,"
                    tSql = tSql & vbCrLf & "FTShdTransType,FDShdTransDate,FTTdmCode,"
                    tSql = tSql & vbCrLf & "FTTdmType,FTSrcCardNo,FTApprCode,"
                    tSql = tSql & vbCrLf & "FCSrcCardChg,FCSrcFAmt,FCSrcAmt,"
                    tSql = tSql & vbCrLf & "FCSrcNet,FCSrcChg,FTSrcStaPrc,"
                    tSql = tSql & vbCrLf & "FTSrcRmk,FTSrcRetDocRef,FCSrcRetAmt,"
                    tSql = tSql & vbCrLf & "FNShdStaSend,"
                    tSql = tSql & vbCrLf & "FTRteCode,"
                    tSql = tSql & vbCrLf & "FCSrcRteAmt ,"
                    tSql = tSql & vbCrLf & "FCSrcRteFac,"
                    tSql = tSql & vbCrLf & "FCSrcRndMnyChg,"
                    tSql = tSql & vbCrLf & "FTSrcAbb,"
                    tSql = tSql & vbCrLf & "FTSrcCardExp,"
                    tSql = tSql & vbCrLf & "FCSrcOntopAmt,"
                    tSql = tSql & vbCrLf & "FNSrcCrdPoint,"
                    tSql = tSql & vbCrLf & "FTREFNO,"  '*Eaw 56-11-05
                    tSql = tSql & vbCrLf & tCS_CNLastFedIns
                    tSql = tSql & vbCrLf & " VALUES('" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nUT_SeqRC & ","
                    tSql = tSql & vbCrLf & "'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & tVB_TenderCode & "',"
                    tSql = tSql & vbCrLf & "'" & tVB_TenderType & "','" & tVB_CrdCardNo & "','" & tVB_CrdApprCode & "',"
'                    tSql = tSql & vbCrLf & 0 & "," & SP_FMTcCurPoint(cFAmt, True, False) & "," & SP_FMTcCurPoint(cVB_Pay, True, False) & ","               '*Ao 2013/09/18
'                    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cNet, True, False) & "," & SP_FMTcCurPoint(cVB_Change, True, False) & ",'',"
                    tSql = tSql & vbCrLf & 0 & "," & cFAmt & "," & cVB_Pay & ","
                    tSql = tSql & vbCrLf & cNet & "," & SP_FMTcCurPoint(cVB_Change, True, False) & ",'',"
                    tSql = tSql & vbCrLf & "'',''," & 0 & ","
                    tSql = tSql & vbCrLf & "0,"
                    tSql = tSql & vbCrLf & "'" & tVB_FTRteCode & "',"
                    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteAmt, True, False) & ","
'                    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_FCSrcRteFac, True, False) & ","
                    tSql = tSql & vbCrLf & Format(cVB_FCSrcRteFac, "###0.00000") & ","            '*Ao 2013/12/04  CommSheet TK-ISS3000-440
                    tSql = tSql & vbCrLf & SP_FMTcCurPoint(cVB_RoundChg, True, False) & ","
                   ' tSql = tSql & vbCrLf & "'" & tVB_TenderName & "',"
                    tSql = tSql & vbCrLf & "'" & tVB_FTRteCode & "',"
                    If tVB_CrdExpM = "" And tVB_CrdExpY = "" Then
                        tSql = tSql & "'" & "" & "',"
                    Else
                        tSql = tSql & "'" & tVB_CrdExpM & "/" & tVB_CrdExpY & "',"
                    End If
                    tSql = tSql & 0 & ","
                    tSql = tSql & 0 & ","
                    tSql = tSql & vbCrLf & "'" & tVB_FTRteCode & "',"  '*Eaw 56-11-05
                    tSql = tSql & UT_tLastInsSql & ")"
                    Call SP_SQLvExecute(tSql)
             Next nI
              tVB_FTRteCode = "" '*Ao 2013/11/05
              cVB_FCSrcRteAmt = 0
              cVB_FCSrcRteFac = 0
              cVB_RoundChg = 0
              tVB_CrdExpM = ""
              tVB_CrdExpY = ""
    End With
    Call W_UPDxVirtualPosRC '*Eaw 56-10-01
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'NEUNG XXX
Private Sub W_FNCxAddFunc2Grid(Optional ByVal ptChkBnkStaAlw As String = "Y")
'--------------------------------------------------------
'Call:
'Cmt: แสดง Banknote
'-------------------------------------------------------
Dim nI As Integer, nR As Integer
Dim orsTemp As ADODB.Recordset
Dim tSql As String
On Error GoTo ErrHandle
    With ogdFunc

        tSql = "SELECT * FROM TCNMBanknote"
        If ptChkBnkStaAlw = "Y" Then
            tSql = tSql & vbCrLf & " WHERE FTBnkStaAlw='Y'"
        End If
        tSql = tSql & vbCrLf & " ORDER BY FCBnkAmt,FTBnkStaTpye,FTBnkCode"
        
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read TCNMBanknote
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            nR = 0
            .rows = orsTemp.RecordCount
            Do While Not orsTemp.EOF
                .TextMatrix(nR, 0) = nR + 1
                .TextMatrix(nR, 1) = SP_FMTtNum2Shw(SP_FEDvChkNumeric(orsTemp, "FCBnkAmt")) & " " & SP_FEDtChkString(orsTemp, "FTBnkNameEn")
                .TextMatrix(nR, 2) = 0
                .TextMatrix(nR, 3) = SP_FMTcCurPoint(0, False, True)
                .TextMatrix(nR, 4) = SP_FEDvChkNumeric(orsTemp, "FCBnkAmt")
                .TextMatrix(nR, 5) = SP_FEDtChkString(orsTemp, "FTBnkCode")
                .TextMatrix(nR, 6) = SP_FEDtChkString(orsTemp, "FTBnkNameTh")
                .TextMatrix(nR, 7) = SP_FEDtChkString(orsTemp, "FTBnkStaTpye")
                .Cell(flexcpBackColor, nR, 2, nR, 2) = &H80000005
                nR = nR + 1
                orsTemp.MoveNext
            Loop
            .ColAlignment(2) = flexAlignRightCenter
        End If
    End With
    
    With ogdOtherCash
            Select Case SP_VALnText2Double(tVB_TransactionSub)
                 Case tEN_TrnPickUp, tEN_TrnDepositMoney, tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                       ' .rows = .rows + 1
                        nR = 0
                        .TextMatrix(nR, 0) = nR + 1
'                        .TextMatrix(nR, 1) = "GIFT VOUCHER"
                        .TextMatrix(nR, 1) = "OTHER CASH"
                        .TextMatrix(nR, 2) = 1
                        .TextMatrix(nR, 3) = SP_FMTcCurPoint(0, False, True)
                        .TextMatrix(nR, 4) = 0
'                        .TextMatrix(nR, 5) = "T018"
'                        .TextMatrix(nR, 5) = "T005"                             '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        .TextMatrix(nR, 5) = "T002"                             '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
'                        .TextMatrix(nR, 6) = "GIFT VOUCHER"
                        .TextMatrix(nR, 6) = "OTHER CASH"
                        .TextMatrix(nR, 7) = "2"
                        .Cell(flexcpBackColor, nR, 3, nR, 3) = &H80000005
             End Select
            .ColAlignment(3) = flexAlignRightCenter
    End With
          
    Exit Sub
ErrHandle:
    Set orsTemp = Nothing
    Call SP_MSGShowError
End Sub

'NEUNG XXX
Private Sub W_FNCxAddFunc2Grid560826(Optional ByVal ptChkBnkStaAlw As String = "Y")
'--------------------------------------------------------
'Call:
'Cmt: แสดง Banknote
'-------------------------------------------------------
Dim nI As Integer, nR As Integer
Dim orsTemp As ADODB.Recordset
Dim tSql As String
On Error GoTo ErrHandle
    With ogdFunc

        tSql = "SELECT * FROM TCNMBanknote"
        If ptChkBnkStaAlw = "Y" Then
            tSql = tSql & vbCrLf & " WHERE FTBnkStaAlw='Y'"
        End If
        tSql = tSql & vbCrLf & " ORDER BY FCBnkAmt,FTBnkStaTpye,FTBnkCode"
        
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read TCNMBanknote
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            nR = 0
            .rows = orsTemp.RecordCount
            Do While Not orsTemp.EOF
                .TextMatrix(nR, 0) = nR + 1
                .TextMatrix(nR, 1) = SP_FMTtNum2Shw(SP_FEDvChkNumeric(orsTemp, "FCBnkAmt")) & " " & SP_FEDtChkString(orsTemp, "FTBnkNameEn")
                .TextMatrix(nR, 2) = 0
                .TextMatrix(nR, 3) = SP_FMTcCurPoint(0, False, True)
                .TextMatrix(nR, 4) = SP_FEDvChkNumeric(orsTemp, "FCBnkAmt")
                .TextMatrix(nR, 5) = SP_FEDtChkString(orsTemp, "FTBnkCode")
                .TextMatrix(nR, 6) = SP_FEDtChkString(orsTemp, "FTBnkNameTh")
                .TextMatrix(nR, 7) = SP_FEDtChkString(orsTemp, "FTBnkStaTpye")
                .Cell(flexcpBackColor, nR, 2, nR, 2) = &H80000005
                nR = nR + 1
                orsTemp.MoveNext
            Loop
            .ColAlignment(2) = flexAlignRightCenter
        End If
    End With
    
    With ogdGiftVoucher
            Select Case SP_VALnText2Double(tVB_TransactionSub)
                 Case tEN_TrnPickUp, tEN_TrnDepositMoney, tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                       ' .rows = .rows + 1
                        nR = 0
                        .TextMatrix(nR, 0) = nR + 1
                        .TextMatrix(nR, 1) = "GIFT VOUCHER"
                        .TextMatrix(nR, 2) = 1
                        .TextMatrix(nR, 3) = SP_FMTcCurPoint(0, False, True)
                        .TextMatrix(nR, 4) = 0
'                        .TextMatrix(nR, 5) = "T018"
                        .TextMatrix(nR, 5) = "T005"                             '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        .TextMatrix(nR, 6) = "GIFT VOUCHER"
                        .TextMatrix(nR, 7) = "2"
                        .Cell(flexcpBackColor, nR, 3, nR, 3) = &H80000005
             End Select
            .ColAlignment(3) = flexAlignRightCenter
    End With
          
    Exit Sub
ErrHandle:
    Set orsTemp = Nothing
    Call SP_MSGShowError
End Sub

'NEUNG XXX
Private Sub W_FNCxAddFunc2GridForeign(Optional ByVal ptChkFcrStaAlw As String = "Y")
'--------------------------------------------------------
'Call:
'Cmt: แสดง TCNMForCurrency
'-------------------------------------------------------
Dim nI As Integer, nR As Integer
Dim orsTemp As ADODB.Recordset
Dim tSql As String
On Error GoTo ErrHandle
    With ogdFunc1

        tSql = "SELECT * FROM TCNMForCurrency"
        tSql = tSql & vbCrLf & " ORDER BY FTFcrCode"
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read TCNMForCurrency
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            nR = 0
            .rows = orsTemp.RecordCount
            Do While Not orsTemp.EOF
                .TextMatrix(nR, 0) = nR + 1
                
                .TextMatrix(nR, 1) = Trim(SP_FEDtChkString(orsTemp, "FTFcrType"))   '*Ao 2013/03/14  TK-ISS3000-094
                .TextMatrix(nR, 2) = 1
                .TextMatrix(nR, 3) = SP_FMTcCurPointForeign(0, False, True)
                .TextMatrix(nR, 4) = SP_FEDtChkString(orsTemp, "FTTdmCode")
                .TextMatrix(nR, 5) = SP_FEDtChkString(orsTemp, "FTFcrCode")
                .TextMatrix(nR, 6) = SP_FEDtChkString(orsTemp, "FTFcrSign")
                .TextMatrix(nR, 7) = SP_FMTtNum2ShwCurrency(SP_FEDvChkNumeric(orsTemp, "FCFcrExchang"))
                
                .Cell(flexcpBackColor, nR, 3, nR, 3) = &H80000005
                
                nR = nR + 1
                orsTemp.MoveNext
            Loop
            .ColAlignment(2) = flexAlignRightCenter
            
        End If
    End With
    Exit Sub
ErrHandle:
    Set orsTemp = Nothing
    Call SP_MSGShowError
End Sub

'NEUNG XXX
'Private Sub W_FNCxAddAmt()
'    Dim nRow As Integer
'
'    With ogdFunc
'        bVB_CNStartBank = True
'        nRow = .TextMatrix(.Row, 0)
'
'               cVB_CNDiscAmt = CDbl(Val(Format(.TextMatrix(.Row, 2))))
'                wCNReturnAmt.bW_IsFreeItem = False
'                wCNReturnAmt.Show vbModal
'
'        Call W_DEFxStartBankAmt(nRow)
'
'         bVB_CNStartBank = False
'
'         Call W_DEFxSumStartBank
'    End With
'
'End Sub

Private Sub W_DEFxStartBankAmt(ByVal pnDiscType As Integer)
Dim nStartBankAmt As Integer
    nStartBankAmt = 0
    nStartBankAmt = cVB_CNDiscAmt

    If SP_FMTcCurPoint(nStartBankAmt, False, True) <> 0 Then
        With ogdFunc
                    .TextMatrix(pnDiscType - 1, 2) = SP_FMTcCurPoint(nStartBankAmt * .TextMatrix(pnDiscType - 1, 3), False, True)
        End With
    End If
    
End Sub
  
  Private Sub W_DEFxStartBankForCurrencyAmt(ByVal pnDiscType As Integer)
  Dim nForCurrencyAmt As Integer
  
  nForCurrencyAmt = 0
  nForCurrencyAmt = cVB_CNDiscAmt
  
    If SP_FMTcCurPoint(nForCurrencyAmt, False, True) <> 0 Then
        With ogdFunc1
                    .TextMatrix(pnDiscType - 1, 2) = SP_FMTtNum2ShwCurrency(nForCurrencyAmt * .TextMatrix(pnDiscType - 1, 6))
        End With
    End If
    
End Sub
  
Private Sub W_DEFxSumStartBank()
Dim nLoop As Integer
Dim cSumCash As Double
Dim cSumCur As Double
Dim cSumGiftVoucher As Double
Dim cSumGV As Double '*BG 56-08-26
Dim cSumVVch As Double '*Tao 62-09-05 ยอดรวม Vendor Voucher
    cSumCash = 0
    cSumGiftVoucher = 0
    
    With ogdFunc
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnPickUp, tEN_TrnDepositMoney, tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                For nLoop = .FixedRows To .rows - 1
                    If .TextMatrix(nLoop, 3) <> "" Then
                         cSumCash = cSumCash + SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False)
                    End If
                Next nLoop
            Case Else
                For nLoop = .FixedRows To .rows - 1
                    If .TextMatrix(nLoop, 3) <> "" Then
                         cSumCash = cSumCash + SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False)
                    End If
                Next nLoop
        End Select
    End With
'    cSumCash = cSumCash + SP_FMTcCurPoint(ogdOtherCash.ValueMatrix(0, 3), True, False)
    
    With ogdGiftVoucher
        cSumGiftVoucher = SP_FMTcCurPoint(ogdGiftVoucher.ValueMatrix(0, 3), True, False)
    End With
    
    '*Tao 62-09-05 ยอดรวม Vendor Voucher
    Select Case SP_VALnText2Double(tVB_TransactionSub)
    Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
        cSumVVch = SP_FMTcCurPoint(ogdVendorVchr.ValueMatrix(0, 3), True, False)
    Case Else
        cSumVVch = 0
    End Select

'    otbSumCash.Text = SP_FMTcCurPoint(cSumCash, True, True)
'    otbSumGiftVoucher.Text = SP_FMTcCurPoint(cSumGiftVoucher, True, True)
'    otbSumTotalGrid1.Text = SP_FMTcCurPoint(cSumGiftVoucher + cSumCash, True, True)
    ogdSumCash.TextMatrix(0, 2) = SP_FMTcCurPoint(cSumGiftVoucher + cSumCash, True, True)
    ogdTotalCash.TextMatrix(0, 2) = SP_FMTcCurPoint(ogdOtherCash.ValueMatrix(0, 3) + cSumCash, True, True) ' *BG 56-08-27  comsheet-285
    
    cSumCur = 0
    With ogdFunc1
        For nLoop = .FixedRows To .rows - 1
            If .TextMatrix(nLoop, 3) <> "" Then
                 cSumCur = cSumCur + SP_FMTcCurPoint(Val(Format(.ValueMatrix(nLoop, 3))) * Val(Format(.ValueMatrix(nLoop, 7))), True, False)
            End If
        Next nLoop
    End With
    otbSumForeign.Text = SP_FMTcCurPoint(cSumCur, True, True)
    ogdSumForeign.TextMatrix(0, 2) = SP_FMTcCurPoint(cSumCur, True, True)
'    otbSumTender.Text = SP_FMTcCurPoint(cSumGiftVoucher + cSumCash + cSumCur, True, True)
    '*BG 56-08-26
        With ogdGVFunc
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnPickUp, tEN_TrnDepositMoney, tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                For nLoop = .FixedRows To .rows - 1
                    If .TextMatrix(nLoop, 3) <> "" Then
                         cSumGV = cSumGV + CDbl(SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False)) '* SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False))
                    End If
                Next nLoop
            Case Else
                For nLoop = .FixedRows To .rows - 1
                    If .TextMatrix(nLoop, 3) <> "" Then
                         cSumGV = cSumGV + CDbl(SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False)) ' * SP_FMTcCurPoint(.ValueMatrix(nLoop, 3), True, False))
                    End If
                Next nLoop
        End Select
    End With
    cSumGV = cSumGV + SP_FMTcCurPoint(ogdOtherGV.ValueMatrix(0, 3), True, False)
    ogdTotalGV.TextMatrix(0, 2) = SP_FMTcCurPoint(cSumGV, True, True)
'    otbSumCash.Text = SP_FMTcCurPoint(cSumGV + cSumCash, True, True)
    otbSumCash.Text = SP_FMTcCurPoint(ogdTotalCash.ValueMatrix(0, 2), True, True)
'    otbSumTotalGrid1.Text = SP_FMTcCurPoint(cSumGV + cSumCash, True, True)
    otbSumGiftVoucher.Text = SP_FMTcCurPoint(cSumGV, True, True)
    otbSumTender.Text = (cSumGV * 100) + (cSumCur * 100) + SP_FMTcCurPoint(ogdTotalCash.ValueMatrix(0, 2) * 100, True, False) + (cSumVVch * 100)
    
    '****
End Sub

'Private Sub W_FNCxAddForCurrencyAmt()
'    Dim nRow As Integer
'
'    With ogdFunc1
'        bVB_CNForCurrency = True
'        nRow = .TextMatrix(.Row, 0)
'        cVB_CNDiscAmt = CDbl(Val(Format(.TextMatrix(.Row, 2))))
'        wCNReturnAmt.bW_IsFreeItem = False
'        wCNReturnAmt.Show vbModal
'        Call W_DEFxStartBankForCurrencyAmt(nRow)
'
'         bVB_CNForCurrency = False
'
'         Call W_DEFxSumStartBankForCurrency
'         Call W_DEFxSumStartBank
'    End With
'
'End Sub

'Private Sub W_DEFxSumStartBankForCurrency()
'Dim nLoop As Integer
'Dim pcSumCash As Double
'    pcSumDisc = 0
'    With ogdFunc1
'        For nLoop = .FixedRows To .rows - 1
'        If .TextMatrix(nLoop, 2) <> "" Then
'             pcSumCash = pcSumCash + SP_FMTcCurPoint(.TextMatrix(nLoop, 2), True, True)
'        End If
'        Next nLoop
'    End With
'
'    otbSumForeign.Text = SP_FMTcCurPoint(pcSumCash, True, True)
'End Sub

Private Function W_INSxTranDT() As Boolean
'------------------------------------------------
'Call:
'Cmt: *Insert Data To DT
'-----------------------------------------------
    Dim nSeq%, cNet As Double
    Dim tSql$, tSkuBarC$, tSKUName$
    Dim tSKUName2$, tSKUNameSrvDoc$, tSKUNameSrvDoc2$
    Dim cVat As Double, cVatable As Double
    Dim cCostIn As Double, cQtyAll As Double
    Dim cQty As Double, cRegPrice As Double, cSalePrice As Double, cB4DisChg As Double, cSaleAmt As Double, cSaleUnit As Double, cSdtStkFac As Double
    Dim tWhr$
    Dim nEff As Long
    Dim nI As Integer, nR As Integer
    Dim tSdtStaPdt As Integer
    Dim nRow As Long
    Dim nGV As Integer
On Error GoTo ErrHandle

        W_INSxTranDT = False
        
        With ogdFunc
            Select Case SP_VALnText2Double(tVB_TransactionSub)
                Case tEN_TrnPickUp, tEN_TrnDepositMoney, tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                      nRow = .rows - 1
                Case Else
                     nRow = .rows - 1
            End Select
             nR = 1
             For nI = .FixedRows To nRow
                    If .TextMatrix(nI, 0) <> "" Then
                    
                            tUT_SKUCode = .TextMatrix(nI, 5)
                            tSKUName = Trim(.TextMatrix(nI, 1))     'ชื่อสินค้า(Default ENG)1
                            tSKUName2 = Trim(.TextMatrix(nI, 6))    'ชื่อสินค้า(Local)2
                            tSKUNameSrvDoc = Trim(.TextMatrix(nI, 1))    'ชื่อสินค้า Service Doc(Default ENG) 1
                            tSKUNameSrvDoc2 = Trim(.TextMatrix(nI, 6))     'ชื่อสินค้า Service Doc (Local)2
                           
                            cQty = SP_FMTcCurPoint(.ValueMatrix(nI, 2), True, False)
                            cQtyAll = SP_FMTcCurPoint(.ValueMatrix(nI, 2), True, False)
                            cRegPrice = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cSalePrice = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cSaleUnit = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cB4DisChg = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cSdtStkFac = 1
                            cSaleAmt = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                           If cSaleAmt = 0 Then
                                cQty = 0
                                cQtyAll = 0
                           End If
                        
                            If Trim(.TextMatrix(nI, 7)) = "Y" Then
                                tSdtStaPdt = "1"
                            Else
                                tSdtStaPdt = "0"
                            End If
                           
                        tSql = "INSERT INTO " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                        tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                        tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                        tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                        tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                        tSql = tSql & vbCrLf & ",FTSdtBarCode"      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        
                        tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                        tSql = tSql & vbCrLf & " VALUES ("
                        tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                        tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                        tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                        tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                        tSql = tSql & vbCrLf & "," & cRegPrice
                        tSql = tSql & vbCrLf & "," & cSaleAmt
                        tSql = tSql & vbCrLf & "," & cSalePrice
                        tSql = tSql & vbCrLf & "," & cSaleUnit
                        tSql = tSql & vbCrLf & "," & cB4DisChg
                        tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                        tSql = tSql & vbCrLf & "," & cSdtStkFac
                        tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"         '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        
                        tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                        If SP_SQLbExecute(tSql, , , nEff) Then
                           If nEff = 1 Then
                               W_INSxTranDT = True
                               nR = nR + 1
                           End If
                        End If
                    End If
                    
             Next nI
             
        End With
        
        '*BG 56-08-29 comsheet-285
        With ogdOtherCash '
            If SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnPickUp Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnDepositMoney Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReport Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReportTaxInc Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReportManualTax Then
'                   tUT_SKUCode = .TextMatrix(0, 5)
                   tUT_SKUCode = "001"
                   tSKUName = Trim(.TextMatrix(0, 1))     'ชื่อสินค้า(Default ENG)1
                   tSKUName2 = Trim(.TextMatrix(0, 6))    'ชื่อสินค้า(Local)2
                   tSKUNameSrvDoc = Trim(.TextMatrix(0, 1))    'ชื่อสินค้า Service Doc(Default ENG) 1
                   tSKUNameSrvDoc2 = Trim(.TextMatrix(0, 6))     'ชื่อสินค้า Service Doc (Local)2

                   cQty = SP_FMTcCurPoint(.ValueMatrix(0, 2), True, False)
                   cQtyAll = SP_FMTcCurPoint(.ValueMatrix(0, 2), True, False)
                   cRegPrice = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
                   cSalePrice = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
                   cSaleUnit = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
                   cB4DisChg = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
                   cSdtStkFac = 1
                   cSaleAmt = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
                    If cSaleAmt = 0 Then
                        cQty = 0
                        cQtyAll = 0
                   End If
                   tSdtStaPdt = "1"

                    tSql = "INSERT INTO " & tVB_CNTblDT
                    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                    tSql = tSql & vbCrLf & ",FTSdtBarCode"      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns

                    tSql = tSql & vbCrLf & " VALUES ("
                    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                    tSql = tSql & vbCrLf & "," & cRegPrice
                    tSql = tSql & vbCrLf & "," & cSaleAmt
                    tSql = tSql & vbCrLf & "," & cSalePrice
                    tSql = tSql & vbCrLf & "," & cSaleUnit
                    tSql = tSql & vbCrLf & "," & cB4DisChg
                    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                    tSql = tSql & vbCrLf & "," & cSdtStkFac
                    tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"     '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls

                    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                    If SP_SQLbExecute(tSql, , , nEff) Then
                        If nEff = 1 Then
                            W_INSxTranDT = True
                            nR = nR + 1
                        End If
                    End If
             End If
        End With

'        With ogdGiftVoucher                 '*Ao 2013/05/04   เพิ่ม SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnDepositMoney
        With ogdGVFunc  '*BG 56-08-29 comsheet - 285 วนรอบเพื่อเก็บค่า GV แต่ละประเภทตาม Back
            nRow = .rows - 1
            For nGV = .FixedRows To nRow
                If .TextMatrix(nGV, 0) <> "" Then
                    If SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnPickUp Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnDepositMoney Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReport Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReportTaxInc Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReportManualTax Then
'                          tUT_SKUCode = .TextMatrix(nGV, 5)
                          tUT_SKUCode = "T005"
                          tSKUName = Trim(.TextMatrix(nGV, 1))     'ชื่อสินค้า(Default ENG)1
'                          tSKUName2 = Trim(.TextMatrix(nGV, 6))    'ชื่อสินค้า(Local)2
                          tSKUName2 = Trim(.TextMatrix(nGV, 1))    'ชื่อสินค้า(Local)2
                          tSKUNameSrvDoc = Trim(.TextMatrix(nGV, 1))    'ชื่อสินค้า Service Doc(Default ENG) 1
'                          tSKUNameSrvDoc2 = Trim(.TextMatrix(nGV, 6))     'ชื่อสินค้า Service Doc (Local)2
                          tSKUNameSrvDoc2 = Trim(.TextMatrix(nGV, 1))     'ชื่อสินค้า Service Doc (Local)2
                          cQty = SP_FMTcCurPoint(.ValueMatrix(nGV, 2), True, False)
                          cQtyAll = SP_FMTcCurPoint(.ValueMatrix(nGV, 2), True, False)
                          cRegPrice = SP_FMTcCurPoint(.ValueMatrix(nGV, 3), True, False)
                          cSalePrice = SP_FMTcCurPoint(.ValueMatrix(nGV, 3), True, False)
                          cSaleUnit = SP_FMTcCurPoint(.ValueMatrix(nGV, 3), True, False)
                          cB4DisChg = SP_FMTcCurPoint(.ValueMatrix(nGV, 3), True, False)
                          cSdtStkFac = 1
                          cSaleAmt = SP_FMTcCurPoint(.ValueMatrix(nGV, 3), True, False)
                           If cSaleAmt = 0 Then
                               cQty = 0
                               cQtyAll = 0
                          End If
                          tSdtStaPdt = "2"
                          
                           tSql = "INSERT INTO " & tVB_CNTblDT
                           tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                           tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                           tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                           tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                           tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                           tSql = tSql & vbCrLf & ",FTSdtBarCode"      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                           tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                           
                           tSql = tSql & vbCrLf & " VALUES ("
                           tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                           tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                           tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                           tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                           tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                           tSql = tSql & vbCrLf & "," & cRegPrice
                           tSql = tSql & vbCrLf & "," & cSaleAmt
                           tSql = tSql & vbCrLf & "," & cSalePrice
                           tSql = tSql & vbCrLf & "," & cSaleUnit
                           tSql = tSql & vbCrLf & "," & cB4DisChg
                           tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                           tSql = tSql & vbCrLf & "," & cSdtStkFac
                           tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"     '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
'                           tSql = tSql & vbCrLf & ",'" & "T005" & "'"    '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                           tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                           If SP_SQLbExecute(tSql, , , nEff) Then
                               If nEff = 1 Then
                                   W_INSxTranDT = True
                                   nR = nR + 1
                               End If
                           End If
                    End If
            End If
        Next nGV
    End With
    
    With ogdOtherGV '
        If SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnPickUp Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnDepositMoney Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReport Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReportTaxInc Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReportManualTax Then
'                   tUT_SKUCode = .TextMatrix(0, 5)
            tUT_SKUCode = "T005"
            tSKUName = Trim(.TextMatrix(0, 1))     'ชื่อสินค้า(Default ENG)1
            tSKUName2 = Trim(.TextMatrix(0, 6))    'ชื่อสินค้า(Local)2
            tSKUNameSrvDoc = Trim(.TextMatrix(0, 1))    'ชื่อสินค้า Service Doc(Default ENG) 1
            tSKUNameSrvDoc2 = Trim(.TextMatrix(0, 6))     'ชื่อสินค้า Service Doc (Local)2

            cQty = SP_FMTcCurPoint(.ValueMatrix(0, 2), True, False)
            cQtyAll = SP_FMTcCurPoint(.ValueMatrix(0, 2), True, False)
            cRegPrice = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
            cSalePrice = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
            cSaleUnit = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
            cB4DisChg = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
            cSdtStkFac = 1
            cSaleAmt = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
             If cSaleAmt = 0 Then
                 cQty = 0
                 cQtyAll = 0
            Else
                cQty = 1
            End If
            tSdtStaPdt = "2"

             tSql = "INSERT INTO " & tVB_CNTblDT
             tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
             tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
             tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
             tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
             tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
             tSql = tSql & vbCrLf & ",FTSdtBarCode"      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
             tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns

             tSql = tSql & vbCrLf & " VALUES ("
             tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
             tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
             tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
             tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
             tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
             tSql = tSql & vbCrLf & "," & cRegPrice
             tSql = tSql & vbCrLf & "," & cSaleAmt
             tSql = tSql & vbCrLf & "," & cSalePrice
             tSql = tSql & vbCrLf & "," & cSaleUnit
             tSql = tSql & vbCrLf & "," & cB4DisChg
             tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
             tSql = tSql & vbCrLf & "," & cSdtStkFac
             tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"     '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls

             tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
             If SP_SQLbExecute(tSql, , , nEff) Then
                 If nEff = 1 Then
                     W_INSxTranDT = True
                     nR = nR + 1
                 End If
             End If
         End If
    End With
        
'            If SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnPickUp Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnDepositMoney Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReport Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReportTaxInc Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReportManualTax Then
'                   tUT_SKUCode = .TextMatrix(0, 5)
'                   tSKUName = Trim(.TextMatrix(0, 1))     'ชื่อสินค้า(Default ENG)1
'                   tSKUName2 = Trim(.TextMatrix(0, 6))    'ชื่อสินค้า(Local)2
'                   tSKUNameSrvDoc = Trim(.TextMatrix(0, 1))    'ชื่อสินค้า Service Doc(Default ENG) 1
'                   tSKUNameSrvDoc2 = Trim(.TextMatrix(0, 6))     'ชื่อสินค้า Service Doc (Local)2
'
'                   cQty = SP_FMTcCurPoint(.ValueMatrix(0, 2), True, False)
'                   cQtyAll = SP_FMTcCurPoint(.ValueMatrix(0, 2), True, False)
'                   cRegPrice = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
'                   cSalePrice = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
'                   cSaleUnit = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
'                   cB4DisChg = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
'                   cSdtStkFac = 1
'                   cSaleAmt = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
'                    If cSaleAmt = 0 Then
'                        cQty = 0
'                        cQtyAll = 0
'                   End If
'                   tSdtStaPdt = "2"
'
'                    tSql = "INSERT INTO " & tVB_CNTblDT
'                    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
'                    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
'                    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
'                    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
'                    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
'                    tSql = tSql & vbCrLf & ",FTSdtBarCode"      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
'                    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
'
'                    tSql = tSql & vbCrLf & " VALUES ("
'                    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
'                    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
'                    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
'                    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
'                    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
'                    tSql = tSql & vbCrLf & "," & cRegPrice
'                    tSql = tSql & vbCrLf & "," & cSaleAmt
'                    tSql = tSql & vbCrLf & "," & cSalePrice
'                    tSql = tSql & vbCrLf & "," & cSaleUnit
'                    tSql = tSql & vbCrLf & "," & cB4DisChg
'                    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
'                    tSql = tSql & vbCrLf & "," & cSdtStkFac
'                    tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"     '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
'
'                    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
'                    If SP_SQLbExecute(tSql, , , nEff) Then
'                        If nEff = 1 Then
'                            W_INSxTranDT = True
'                            nR = nR + 1
'                        End If
'                    End If
'             End If
'      End With
        
        '*Tao 62-08-05 Enter Vendor Voucher
        cQty = ogdVendorVchr.ValueMatrix(0, 2) 'จำนวน Vendor Voucher ที่ป้อนเข้าไป
        cNet = ogdVendorVchr.ValueMatrix(0, 3) 'ยอดรวม Vendor Voucher ที่ป้อนเข้าไป
        tUT_SKUCode = "T006"
        tSKUName = "VVch" 'ชื่อสินค้า(Default ENG)1
        tSKUName2 = "VVch" 'ชื่อสินค้า(Local)2
        tSKUNameSrvDoc = tSKUName 'ชื่อสินค้า Service Doc(Default ENG) 1
        tSKUNameSrvDoc2 = tSKUName2   'ชื่อสินค้า Service Doc (Local)2
        If cNet = 0 Then
          cQty = 0
          cQtyAll = 0
        Else
          If (cQty = 0) Then cQty = 1
          cQtyAll = cQty
        End If
        cRegPrice = cNet
        cSalePrice = cNet
        cSaleUnit = cNet
        cB4DisChg = cNet
        cSdtStkFac = cNet
        cSaleAmt = cNet
        cSaleAmt = cNet
        tSdtStaPdt = "6"
    
        tSql = "INSERT INTO " & tVB_CNTblDT
        tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
        tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
        tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
        tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
        tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
        tSql = tSql & vbCrLf & ",FTSdtBarCode"
        tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
        tSql = tSql & vbCrLf & " VALUES ("
        tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
        tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
        tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
        tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
        tSql = tSql & vbCrLf & "," & cSalePrice
        tSql = tSql & vbCrLf & "," & cSaleAmt
        tSql = tSql & vbCrLf & "," & cSalePrice
        tSql = tSql & vbCrLf & "," & cSaleUnit
        tSql = tSql & vbCrLf & "," & cB4DisChg
        tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
        tSql = tSql & vbCrLf & "," & cSdtStkFac
        tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"
        tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
        If SP_SQLbExecute(tSql, , , nEff) Then
            If nEff = 1 Then
                W_INSxTranDT = True
                nR = nR + 1
            End If
        End If
        '=======================================
        
        With ogdFunc1
             For nI = .FixedRows To .rows - 1
                    If .TextMatrix(nI, 0) <> "" Then
                        tUT_SKUCode = .TextMatrix(nI, 5) 'เก็บเป็นรหัสสกุลเงิน เพื่อเชื่อมกับ tSdtStaPdt = "4" และ tSdtStaPdt = "5"
                        tSKUName = Trim(.TextMatrix(nI, 1))     'ชื่อสินค้า(Default ENG)1
                        tSKUName2 = Trim(.TextMatrix(nI, 1))    'ชื่อสินค้า(Local)2
                        tSKUNameSrvDoc = Trim(.TextMatrix(nI, 1))    'ชื่อสินค้า Service Doc(Default ENG) 1
                        tSKUNameSrvDoc2 = Trim(.TextMatrix(nI, 1))     'ชื่อสินค้า Service Doc (Local)2
                        
                        cQty = 1
                        cQtyAll = 1
                        cRegPrice = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
'                        cSalePrice = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                        cSalePrice = cRegPrice
'                        cSaleUnit = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                        cSaleUnit = cRegPrice
'                        cB4DisChg = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                        cB4DisChg = cRegPrice
                        cSdtStkFac = CDbl(Val(Format(.ValueMatrix(nI, 7)))) 'ทศนิยม 5 ตำแหน่ง
'                        cSaleAmt = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                        cSaleAmt = SP_FMTcCurPoint(cRegPrice * cSdtStkFac, True, False)
                        If cSaleAmt = 0 Then                                        '*Ao 2013/09/18      แก้ไข การคูณแล้วปัดเศษ เช่น       0.01 * 0.00123 = 0.0000123     ปัดแล้วได้ 0.00
                            If cRegPrice > 0 Then
                                cSaleAmt = cRegPrice * cSdtStkFac
                            Else
                                cQty = 0
                                cQtyAll = 0
                            End If
'                            cQty = 0
'                            cQtyAll = 0
                        End If
                        tSdtStaPdt = "3"
                
                           tSql = "INSERT INTO " & tVB_CNTblDT
                           tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                           tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                           tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                           tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                           tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                           tSql = tSql & vbCrLf & ",FTSdtBarCode"      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                           tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                           
                           tSql = tSql & vbCrLf & " VALUES ("
                           tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                           tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                           tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                           tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                           tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                           tSql = tSql & vbCrLf & "," & cSalePrice
                           tSql = tSql & vbCrLf & "," & cSaleAmt
                           tSql = tSql & vbCrLf & "," & cSalePrice
                           tSql = tSql & vbCrLf & "," & cSaleUnit
                           tSql = tSql & vbCrLf & "," & cB4DisChg
                           tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                           tSql = tSql & vbCrLf & "," & cSdtStkFac
                           'tSql = tSql & vbCrLf & ",'T012'"     '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                           tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"     '*BG 56-07-08 CommSheet TK-ISS3000-208-ST 00 00 01.xls
                           tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                           If SP_SQLbExecute(tSql, , , nEff) Then
                               If nEff = 1 Then
                                   W_INSxTranDT = True
                                   nR = nR + 1
                               End If
                           End If
                    End If
                    
             Next nI
        End With
                
    Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                  Call W_PRNbOverShort(nR)
                  Call W_UPDxVirtualPosDT '*Eaw 56-10-01
    End Select
    
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function W_INSxTranDT560829() As Boolean
'------------------------------------------------
'Call:
'Cmt: *Insert Data To DT
'-----------------------------------------------
    Dim nSeq%, cNet As Double
    Dim tSql$, tSkuBarC$, tSKUName$
    Dim tSKUName2$, tSKUNameSrvDoc$, tSKUNameSrvDoc2$
    Dim cVat As Double, cVatable As Double
    Dim cCostIn As Double, cQtyAll As Double
    Dim cQty As Double, cRegPrice As Double, cSalePrice As Double, cB4DisChg As Double, cSaleAmt As Double, cSaleUnit As Double, cSdtStkFac As Double
    Dim tWhr$
    Dim nEff As Long
    Dim nI As Integer, nR As Integer
    Dim tSdtStaPdt As Integer
    Dim nRow As Long
On Error GoTo ErrHandle

        W_INSxTranDT560829 = False
        
        With ogdFunc
            Select Case SP_VALnText2Double(tVB_TransactionSub)
                Case tEN_TrnPickUp, tEN_TrnDepositMoney, tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                      nRow = .rows - 1
                Case Else
                     nRow = .rows - 1
            End Select
             nR = 1
             For nI = .FixedRows To nRow
                    If .TextMatrix(nI, 0) <> "" Then
                    
                            tUT_SKUCode = .TextMatrix(nI, 5)
                            tSKUName = Trim(.TextMatrix(nI, 1))     'ชื่อสินค้า(Default ENG)1
                            tSKUName2 = Trim(.TextMatrix(nI, 6))    'ชื่อสินค้า(Local)2
                            tSKUNameSrvDoc = Trim(.TextMatrix(nI, 1))    'ชื่อสินค้า Service Doc(Default ENG) 1
                            tSKUNameSrvDoc2 = Trim(.TextMatrix(nI, 6))     'ชื่อสินค้า Service Doc (Local)2
                           
                            cQty = SP_FMTcCurPoint(.ValueMatrix(nI, 2), True, False)
                            cQtyAll = SP_FMTcCurPoint(.ValueMatrix(nI, 2), True, False)
                            cRegPrice = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cSalePrice = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cSaleUnit = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cB4DisChg = SP_FMTcCurPoint(.ValueMatrix(nI, 4), True, False)
                            cSdtStkFac = 1
                            cSaleAmt = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                           If cSaleAmt = 0 Then
                                cQty = 0
                                cQtyAll = 0
                           End If
                        
                            If Trim(.TextMatrix(nI, 7)) = "Y" Then
                                tSdtStaPdt = "1"
                            Else
                                tSdtStaPdt = "0"
                            End If
                           
                        tSql = "INSERT INTO " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                        tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                        tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                        tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                        tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                        tSql = tSql & vbCrLf & ",FTSdtBarCode"      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        
                        tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                        tSql = tSql & vbCrLf & " VALUES ("
                        tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                        tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                        tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                        tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                        tSql = tSql & vbCrLf & "," & cRegPrice
                        tSql = tSql & vbCrLf & "," & cSaleAmt
                        tSql = tSql & vbCrLf & "," & cSalePrice
                        tSql = tSql & vbCrLf & "," & cSaleUnit
                        tSql = tSql & vbCrLf & "," & cB4DisChg
                        tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                        tSql = tSql & vbCrLf & "," & cSdtStkFac
                        tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"         '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                        
                        tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                        If SP_SQLbExecute(tSql, , , nEff) Then
                           If nEff = 1 Then
                               W_INSxTranDT560829 = True
                               nR = nR + 1
                           End If
                        End If
                    End If
                    
             Next nI
             
        End With
        
        With ogdGiftVoucher                 '*Ao 2013/05/04   เพิ่ม SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnDepositMoney
            If SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnPickUp Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnDepositMoney Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReport Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReportTaxInc Or SP_VALnText2Double(tVB_TransactionSub) = tEN_TrnOverShortReportManualTax Then
                   tUT_SKUCode = .TextMatrix(0, 5)
                   tSKUName = Trim(.TextMatrix(0, 1))     'ชื่อสินค้า(Default ENG)1
                   tSKUName2 = Trim(.TextMatrix(0, 6))    'ชื่อสินค้า(Local)2
                   tSKUNameSrvDoc = Trim(.TextMatrix(0, 1))    'ชื่อสินค้า Service Doc(Default ENG) 1
                   tSKUNameSrvDoc2 = Trim(.TextMatrix(0, 6))     'ชื่อสินค้า Service Doc (Local)2
                
                   cQty = SP_FMTcCurPoint(.ValueMatrix(0, 2), True, False)
                   cQtyAll = SP_FMTcCurPoint(.ValueMatrix(0, 2), True, False)
                   cRegPrice = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
                   cSalePrice = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
                   cSaleUnit = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
                   cB4DisChg = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
                   cSdtStkFac = 1
                   cSaleAmt = SP_FMTcCurPoint(.ValueMatrix(0, 3), True, False)
                    If cSaleAmt = 0 Then
                        cQty = 0
                        cQtyAll = 0
                   End If
                   tSdtStaPdt = "2"
                   
                    tSql = "INSERT INTO " & tVB_CNTblDT
                    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                    tSql = tSql & vbCrLf & ",FTSdtBarCode"      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                    
                    tSql = tSql & vbCrLf & " VALUES ("
                    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                    tSql = tSql & vbCrLf & "," & cRegPrice
                    tSql = tSql & vbCrLf & "," & cSaleAmt
                    tSql = tSql & vbCrLf & "," & cSalePrice
                    tSql = tSql & vbCrLf & "," & cSaleUnit
                    tSql = tSql & vbCrLf & "," & cB4DisChg
                    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                    tSql = tSql & vbCrLf & "," & cSdtStkFac
                    tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"     '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    
                    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                    If SP_SQLbExecute(tSql, , , nEff) Then
                        If nEff = 1 Then
                            W_INSxTranDT560829 = True
                            nR = nR + 1
                        End If
                    End If
             End If
      End With
        
        With ogdFunc1
             For nI = .FixedRows To .rows - 1
                    If .TextMatrix(nI, 0) <> "" Then
                           tUT_SKUCode = .TextMatrix(nI, 5) 'เก็บเป็นรหัสสกุลเงิน เพื่อเชื่อมกับ tSdtStaPdt = "4" และ tSdtStaPdt = "5"
                           tSKUName = Trim(.TextMatrix(nI, 1))     'ชื่อสินค้า(Default ENG)1
                           tSKUName2 = Trim(.TextMatrix(nI, 1))    'ชื่อสินค้า(Local)2
                           tSKUNameSrvDoc = Trim(.TextMatrix(nI, 1))    'ชื่อสินค้า Service Doc(Default ENG) 1
                           tSKUNameSrvDoc2 = Trim(.TextMatrix(nI, 1))     'ชื่อสินค้า Service Doc (Local)2
                           
                           cQty = 1
                           cQtyAll = 1
                           cRegPrice = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                           cSalePrice = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                           cSaleUnit = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                           cB4DisChg = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                           cSdtStkFac = CDbl(Val(Format(.ValueMatrix(nI, 7)))) 'ทศนิยม 5 ตำแหน่ง
                           cSaleAmt = SP_FMTcCurPoint(.ValueMatrix(nI, 3), True, False)
                           cSaleAmt = SP_FMTcCurPoint(cSaleAmt * cSdtStkFac, True, False)
                           If cSaleAmt = 0 Then
                                cQty = 0
                                cQtyAll = 0
                           End If
                           tSdtStaPdt = "3"
                
                           tSql = "INSERT INTO " & tVB_CNTblDT
                           tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                           tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                           tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                           tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                           tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                           tSql = tSql & vbCrLf & ",FTSdtBarCode"      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                           tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                           
                           tSql = tSql & vbCrLf & " VALUES ("
                           tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                           tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                           tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                           tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                           tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                           tSql = tSql & vbCrLf & "," & cSalePrice
                           tSql = tSql & vbCrLf & "," & cSaleAmt
                           tSql = tSql & vbCrLf & "," & cSalePrice
                           tSql = tSql & vbCrLf & "," & cSaleUnit
                           tSql = tSql & vbCrLf & "," & cB4DisChg
                           tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                           tSql = tSql & vbCrLf & "," & cSdtStkFac
                           'tSql = tSql & vbCrLf & ",'T012'"     '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                           tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"     '*BG 56-07-08 CommSheet TK-ISS3000-208-ST 00 00 01.xls
                           tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                           If SP_SQLbExecute(tSql, , , nEff) Then
                               If nEff = 1 Then
                                   W_INSxTranDT560829 = True
                                   nR = nR + 1
                               End If
                           End If
                    End If
                    
             Next nI
        End With
        
        Select Case SP_VALnText2Double(tVB_TransactionSub)
                Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                      Call W_PRNbOverShort(nR)
        End Select
        
        Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub W_PRNbOverShort(ByVal pnSeqNo As Long)

    Dim cACTUAL As Double
    Dim cQtyACTUAL As Double
    
    Dim nSeq%, cNet As Double
    Dim tSql$, tSkuBarC$, tSKUName$
    Dim tSKUName2$, tSKUNameSrvDoc$, tSKUNameSrvDoc2$
    Dim cVat As Double, cVatable As Double
    Dim cCostIn As Double, cQtyAll As Double
    Dim cQty As Double, cRegPrice As Double, cSalePrice As Double, cB4DisChg As Double, cSaleAmt As Double, cSaleUnit As Double, cSdtStkFac As Double
    Dim tWhr$
    Dim nEff As Long
    Dim nI As Integer, nR As Integer
    Dim tSdtStaPdt As Integer
    Dim nRow As Long
    
    Dim orsPickup As ADODB.Recordset
    Dim orsTemp As ADODB.Recordset
    Dim tTransWhere As String
    Dim nError As Long
    Dim cReg As Double
    Dim cRate As Double
    Dim cACTUALTmp As Double
    Dim cChg As Double
    Dim tTenderType As String '*Eaw 56-06-26 CommSheet TK-ISS3000-FAPL-029.xls
    Dim tTransType As String  '*Eaw 56-06-28
    
On Error GoTo ErrHandler
    nR = pnSeqNo
    tSql = wFunctionKB.W_GETtCloseTenderSQLEx(tVB_CNTblHD, tVB_CNTblRC, "FTShdTransType", True)
    Call SP_SQLvExecute(tSql, orsTemp)
    
   'cash
    cACTUAL = 0
    cQtyACTUAL = 0
    cChg = 0   '*Eaw 56-06-26 CommSheet TK-ISS3000-FAPL-029.xls
    If Not (orsTemp.BOF And orsTemp.EOF) Then
    orsTemp.Filter = "FTTdmCode = 'T001'"
    If Not (orsTemp.BOF And orsTemp.EOF) Then
'       cACTUAL = SP_FEDvChkNumeric(orsTemp, "FCSrcNet")
'       cQtyACTUAL = SP_FEDvChkNumeric(orsTemp, "FCQty")
'       cChg = SP_FEDvChkNumeric(orsTemp, "FCSrcChg")                       '*Ao 2013/05/31 CommSheet TK-ISS3000-217-ST 00 00 01-Ao.xls
        orsTemp.MoveFirst
        Do While Not orsTemp.EOF
             tTransType = SP_FEDtChkString(orsTemp, "FTShdTransType")
            cACTUAL = cACTUAL + SP_FEDvChkNumeric(orsTemp, "FCSrcNet")
'            Select Case UCase(Trim(tTransType))
''                Case "03", "06", "07", "10", "11", "13", "19", "22"
'                Case "03", "06", "07", "10", "11", "13", "19", "22", "38", "43", "04", "16", "17", "20" '*BG 56-09-04 Comsheet-285
'                    cQtyACTUAL = cQtyACTUAL + SP_FEDvChkNumeric(orsTemp, "FCQty")
'            End Select
            cQtyACTUAL = cQtyACTUAL + SP_FEDvChkNumeric(orsTemp, "FCQty")           '*Ao 2013/09/23 CommSheet TK-ISS3000-350-UAT-FAPL-118_00.00.03 FSBT NG
            orsTemp.MoveNext
        Loop
    End If
    orsTemp.Filter = adFilterNone
    End If
    If Not (orsTemp.BOF And orsTemp.EOF) Then  '*Eaw 56-06-26 CommSheet TK-ISS3000-FAPL-029.xls
        orsTemp.MoveFirst
        Do While Not (orsTemp.EOF)
            tTenderType = SP_FEDtChkString(orsTemp, "FTTdmCode")
            Select Case UCase(Trim(tTenderType))
'                Case "T004", "T005", "T006", "T007", "T008", "T009", "T010", "T011", "T012", "T014"
                Case "T001", "T004", "T005", "T006", "T007", "T008", "T009", "T010", "T011", "T012", "T014"
                    cChg = cChg + SP_FEDvChkNumeric(orsTemp, "FCSrcChg")
            End Select
            orsTemp.MoveNext
        Loop
   End If
    cACTUAL = cACTUAL - cChg                                '*Ao 2013/05/31 CommSheet TK-ISS3000-217-ST 00 00 01-Ao.xls
    tUT_SKUCode = "Cash"
    tSKUName = "Cash"  'ชื่อสินค้า(Default ENG)1
    tSKUName2 = "Cash"   'ชื่อสินค้า(Local)2
    tSKUNameSrvDoc = "Cash"  'ชื่อสินค้า Service Doc(Default ENG) 1
    tSKUNameSrvDoc2 = "Cash"    'ชื่อสินค้า Service Doc (Local)2
    
    cQty = cQtyACTUAL
    cQtyAll = cQtyACTUAL
    
    cRegPrice = cACTUAL
    cSalePrice = cACTUAL
    cSaleUnit = cACTUAL
    cB4DisChg = cACTUAL
    cSdtStkFac = cACTUAL
    cSaleAmt = cACTUAL
    cSaleAmt = cACTUAL
    tSdtStaPdt = "4"

    tSql = "INSERT INTO " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
    tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
    tSql = tSql & vbCrLf & " VALUES ("
    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleAmt
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleUnit
    tSql = tSql & vbCrLf & "," & cB4DisChg
    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
    tSql = tSql & vbCrLf & "," & cSdtStkFac
    tSql = tSql & vbCrLf & ",'T001'"      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
    If SP_SQLbExecute(tSql, , , nEff) Then
        If nEff = 1 Then
            'none
        End If
    End If

   'GVch
   nR = nR + 1
   cACTUAL = 0
   cQtyACTUAL = 0
   If Not (orsTemp.BOF And orsTemp.EOF) Then
'        orsTemp.Filter = "FTTdmCode = 'T004' OR FTTdmCode = 'T005'  OR FTTdmCode = 'T006' OR FTTdmCode = 'T007' OR FTTdmCode = 'T008' OR FTTdmCode = 'T009'  OR FTTdmCode = 'T010' OR FTTdmCode = 'T011' OR FTTdmCode = 'T018'"
'        orsTemp.Filter = "FTTdmCode = 'T004' OR FTTdmCode = 'T005'  OR FTTdmCode = 'T006' OR FTTdmCode = 'T007' OR FTTdmCode = 'T008' OR FTTdmCode = 'T009'  OR FTTdmCode = 'T010' OR FTTdmCode = 'T011'"  '*Eaw 56-06-26 ไม่นับ Credit Voucher เป็น GVch
        orsTemp.Filter = "FTTdmCode = 'T005'"   '*BG 56-08-07  คิดรวมแค่ GV อย่างเดียว" CommSheet TK-ISS3000-303-UAT_Question.xls
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            'cACTUAL = SP_FEDvChkNumeric(orsTemp, "FCSrcNet")
            'cQtyACTUAL = SP_FEDvChkNumeric(orsTemp, "FCQty")
            Do While Not orsTemp.EOF
                If SP_FEDtChkString(orsTemp, "FTTdmCode") = "T018" And (SP_FEDtChkString(orsTemp, "FTShdTransType") = "04" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "39" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "44" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "14") Then
                    cACTUAL = cACTUAL + 0
                    cQtyACTUAL = cQtyACTUAL + 0
                Else
                    cACTUAL = cACTUAL + SP_FEDvChkNumeric(orsTemp, "FCSrcNet")
                    cQtyACTUAL = cQtyACTUAL + SP_FEDvChkNumeric(orsTemp, "FCQty")
                End If
                orsTemp.MoveNext
            Loop
        End If
         orsTemp.Filter = adFilterNone
   End If
   
    tUT_SKUCode = "GVch"
    tSKUName = "GVch"  'ชื่อสินค้า(Default ENG)1
    tSKUName2 = "GVch"   'ชื่อสินค้า(Local)2
    tSKUNameSrvDoc = "GVch"  'ชื่อสินค้า Service Doc(Default ENG) 1
    tSKUNameSrvDoc2 = "GVch"    'ชื่อสินค้า Service Doc (Local)2
    
    cQty = cQtyACTUAL
    cQtyAll = cQtyACTUAL
    
    cRegPrice = cACTUAL
    cSalePrice = cACTUAL
    cSaleUnit = cACTUAL
    cB4DisChg = cACTUAL
    cSdtStkFac = cACTUAL
    cSaleAmt = cACTUAL
    cSaleAmt = cACTUAL
    tSdtStaPdt = "4"

    tSql = "INSERT INTO " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
    tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
    tSql = tSql & vbCrLf & " VALUES ("
    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleAmt
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleUnit
    tSql = tSql & vbCrLf & "," & cB4DisChg
    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
    tSql = tSql & vbCrLf & "," & cSdtStkFac
    tSql = tSql & vbCrLf & ",'T005'"        '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
    If SP_SQLbExecute(tSql, , , nEff) Then
        If nEff = 1 Then
            'none
        End If
    End If
   
   '*Tao 62-08-06 Sum Vendor Voucher
   nR = nR + 1
   cACTUAL = 0
   cQtyACTUAL = 0
   If Not (orsTemp.BOF And orsTemp.EOF) Then
        orsTemp.Filter = "FTTdmCode = 'T006'"
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            Do While Not orsTemp.EOF
'                If SP_FEDtChkString(orsTemp, "FTTdmCode") = "T018" And (SP_FEDtChkString(orsTemp, "FTShdTransType") = "04" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "39" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "44" Or SP_FEDtChkString(orsTemp, "FTShdTransType") = "14") Then
'                    cACTUAL = cACTUAL + 0
'                    cQtyACTUAL = cQtyACTUAL + 0
'                Else
                    cACTUAL = cACTUAL + SP_FEDvChkNumeric(orsTemp, "FCSrcNet")
                    cQtyACTUAL = cQtyACTUAL + SP_FEDvChkNumeric(orsTemp, "FCQty")
'                End If
                orsTemp.MoveNext
            Loop
        End If
         orsTemp.Filter = adFilterNone
   End If
   
    tUT_SKUCode = "T006"
    tSKUName = "VVch"  'ชื่อสินค้า(Default ENG)1
    tSKUName2 = "VVch"   'ชื่อสินค้า(Local)2
    tSKUNameSrvDoc = tSKUName  'ชื่อสินค้า Service Doc(Default ENG) 1
    tSKUNameSrvDoc2 = tSKUName2    'ชื่อสินค้า Service Doc (Local)2
    
    cQty = cQtyACTUAL
    cQtyAll = cQtyACTUAL
    
    cRegPrice = cACTUAL
    cSalePrice = cACTUAL
    cSaleUnit = cACTUAL
    cB4DisChg = cACTUAL
    cSdtStkFac = cACTUAL
    cSaleAmt = cACTUAL
    cSaleAmt = cACTUAL
    tSdtStaPdt = "4"

    tSql = "INSERT INTO " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
    tSql = tSql & vbCrLf & ",FTSdtBarCode"
    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
    tSql = tSql & vbCrLf & " VALUES ("
    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleAmt
    tSql = tSql & vbCrLf & "," & cSalePrice
    tSql = tSql & vbCrLf & "," & cSaleUnit
    tSql = tSql & vbCrLf & "," & cB4DisChg
    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
    tSql = tSql & vbCrLf & "," & cSdtStkFac
    tSql = tSql & vbCrLf & ",'T006'"
    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
    If SP_SQLbExecute(tSql, , , nEff) Then
        If nEff = 1 Then
            'none
        End If
    End If
    
    'orsTemp.Close
    Call SP_DATxRsNothing(orsTemp)
    
    tSql = wFunctionKB.W_GETtCloseTenderCurSQL(tVB_CNTblHD, tVB_CNTblRC, "FTShdTransType", True)
    Call SP_SQLvExecute(tSql, orsTemp)
    With ogdFunc1
            For nI = .FixedRows To .rows - 1
                    nR = nR + 1
                    cACTUAL = 0
                    cQtyACTUAL = 0
                    cACTUALTmp = 0
                     If Not (orsTemp.BOF And orsTemp.EOF) Then
                        orsTemp.Filter = "FTRteCode = '" & .TextMatrix(nI, 5) & "'"
                        If Not (orsTemp.BOF And orsTemp.EOF) Then
                            'cACTUAL = SP_FEDvChkNumeric(orsTemp, "FCSrcNet") 'TK-ISS3000-164 Over/Short Report - Cash & Foreign Currency problem
'                            cACTUAL = SP_FEDvChkNumeric(orsTemp, "FCSrcRteAmt")
                            cACTUAL = SP_FEDvChkNumeric(orsTemp, "FCSrcNet")               '*Ao แก้ไขทำให้การคำนวณ OVER/SHORT(BASE)
                            cACTUALTmp = SP_FEDvChkNumeric(orsTemp, "FCSrcRteAmt")
                            cQtyACTUAL = SP_FEDvChkNumeric(orsTemp, "FCQty")
                        End If
                        orsTemp.Filter = adFilterNone
                    End If
                    tUT_SKUCode = .TextMatrix(nI, 5) 'เก็บเป็นรหัสสกุลเงิน เพื่อเชื่อมกับ tSdtStaPdt = "4" และ tSdtStaPdt = "5"
                    tSKUName = Trim(.TextMatrix(nI, 1))  'ชื่อสินค้า(Default ENG)1
                    tSKUName2 = Trim(.TextMatrix(nI, 5))   'ชื่อสินค้า(Local)2
                    tSKUNameSrvDoc = tSKUName  'ชื่อสินค้า Service Doc(Default ENG) 1
                    tSKUNameSrvDoc2 = tSKUName2   'ชื่อสินค้า Service Doc (Local)2
                    
                    cQty = cQtyACTUAL
                    cQtyAll = cQtyACTUAL
                    
                    cRegPrice = cACTUAL
'                    cSalePrice = cACTUAL
                    cSalePrice = cACTUALTmp     '*Ao แก้ไขทำให้การคำนวณ OVER/SHORT(BASE)
                    cSaleUnit = cACTUAL
                    cB4DisChg = cACTUAL
                    cSdtStkFac = cACTUAL
                    cSaleAmt = cACTUAL
                    cSaleAmt = cACTUAL
                    tSdtStaPdt = "4"
                
                    tSql = "INSERT INTO " & tVB_CNTblDT
                    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                    tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                    tSql = tSql & vbCrLf & " VALUES ("
                    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                    tSql = tSql & vbCrLf & "," & cSalePrice
                    tSql = tSql & vbCrLf & "," & cSaleAmt
                    tSql = tSql & vbCrLf & "," & cSalePrice
                    tSql = tSql & vbCrLf & "," & cSaleUnit
                    tSql = tSql & vbCrLf & "," & cB4DisChg
                    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                    tSql = tSql & vbCrLf & "," & cSdtStkFac
                    'tSql = tSql & vbCrLf & ",'T012'"          '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'" '*BG 56-07-10   CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                    If SP_SQLbExecute(tSql, , , nEff) Then
                        If nEff = 1 Then
                            'none
                        End If
                    End If
        Next nI
  End With

   '***CASH TRANSFER REPORT***
   'summary pickup money
   '57-05-08   20130504-ReEntry.xls Re-Entry ไม่ต้องแสดง ASH TRANSFER REPORT
   Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOverShortReport
            tTransWhere = tVB_CNTblHD & ".FDShdTransDate = '" & tUT_SaleDate & "' AND " & tVB_CNTblHD & ".FTTmnNum ='" & tVB_CNTerminalNum & "'"
            tTransWhere = tTransWhere & vbCrLf & " AND " & tVB_CNTblHD & ".FTShdTransType = '12'"
'           tTransWhere = tTransWhere & vbCrLf & " AND " & tVB_CNTblHD & ".FTLogCode  = '" & tUT_LoginCode & "'"  '*BG 56-09-09   ไม่วงไม่แวแล้ว พี่โอ๋บอก
            tTransWhere = tTransWhere & vbCrLf & " GROUP BY FTSkuCode,FTSkuAbbName,FTSkuAbbNameSnd,FTSdtStaPdt,FCSdtStkFac"
            tSql = "SELECT FTSkuCode,FTSkuAbbName,FTSkuAbbNameSnd,FCSdtStkFac,SUM(FCSdtSaleAmt) AS FCSum,SUM(FCSdtRegPrice) AS FCSum2,FTSdtStaPdt"
            tSql = tSql & vbCrLf & "FROM   " & tVB_CNTblHD & " INNER JOIN"
            tSql = tSql & vbCrLf & "               " & tVB_CNTblDT & " ON " & tVB_CNTblHD & ".FTTmnNum = " & tVB_CNTblDT & ".FTTmnNum AND " & tVB_CNTblHD & ".FTShdTransNo = " & tVB_CNTblDT & ".FTShdTransNo AND "
            tSql = tSql & vbCrLf & "               " & tVB_CNTblHD & ".FTShdTransType = " & tVB_CNTblDT & ".FTShdTransType And " & tVB_CNTblHD & ".FDShdTransDate = " & tVB_CNTblDT & ".FDShdTransDate"
            tSql = tSql & vbCrLf & " WHERE " & tTransWhere
            nError = SP_SQLvExecute(tSql, orsPickup, oDB.ocnOffLine, , , , , False)
            If nError = 0 Then
                  cNet = 0
                  nR = nR + 1
                 If Not (orsPickup.BOF And orsPickup.EOF) Then
                        orsPickup.Filter = "FTSdtStaPdt = '0' OR FTSdtStaPdt = '1'"
                        If Not (orsPickup.BOF And orsPickup.EOF) Then
                            Do While Not orsPickup.EOF
                                cNet = cNet + SP_FEDvChkNumeric(orsPickup, "FCSum") 'TK-ISS3000-164 Over/Short Report - Cash & Foreign Currency problem
                                orsPickup.MoveNext
                            Loop
                        End If
                        orsPickup.Filter = adFilterNone
                 End If
'                tUT_SKUCode = "Cash"
                tUT_SKUCode = "T001"                                                     '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                tSKUName = "Cash" 'ชื่อสินค้า(Default ENG)1
                tSKUName2 = "Cash" 'ชื่อสินค้า(Local)2
                tSKUNameSrvDoc = tSKUName 'ชื่อสินค้า Service Doc(Default ENG) 1
                tSKUNameSrvDoc2 = tSKUName2   'ชื่อสินค้า Service Doc (Local)2
                If cNet = 0 Then
                  cQty = 0
                  cQtyAll = 0
                Else
                  cQty = 1
                  cQtyAll = 1
                End If
                cRegPrice = cNet
                cSalePrice = cNet
                cSaleUnit = cNet
                cB4DisChg = cNet
                cSdtStkFac = cNet
                cSaleAmt = cNet
                cSaleAmt = cNet
                tSdtStaPdt = "5"
            
                tSql = "INSERT INTO " & tVB_CNTblDT
                tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                tSql = tSql & vbCrLf & " VALUES ("
                tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                tSql = tSql & vbCrLf & "," & cSalePrice
                tSql = tSql & vbCrLf & "," & cSaleAmt
                tSql = tSql & vbCrLf & "," & cSalePrice
                tSql = tSql & vbCrLf & "," & cSaleUnit
                tSql = tSql & vbCrLf & "," & cB4DisChg
                tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                tSql = tSql & vbCrLf & "," & cSdtStkFac
                tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"           '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                If SP_SQLbExecute(tSql, , , nEff) Then
                    If nEff = 1 Then
                       'none
                    End If
                End If
                
                'CV
                nR = nR + 1
                cNet = 0
                 If Not (orsPickup.BOF And orsPickup.EOF) Then
                        orsPickup.Filter = "FTSdtStaPdt = '2'"
                        If Not (orsPickup.BOF And orsPickup.EOF) Then
                            Do While Not orsPickup.EOF
                                cNet = cNet + SP_FEDvChkNumeric(orsPickup, "FCSum")
                                orsPickup.MoveNext
                            Loop
                        End If
                        orsPickup.Filter = adFilterNone
                 End If
'                tUT_SKUCode = "GVch"                        '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                tUT_SKUCode = "T005"
                tSKUName = "GVch" 'ชื่อสินค้า(Default ENG)1
                tSKUName2 = "GVch" 'ชื่อสินค้า(Local)2
                tSKUNameSrvDoc = tSKUName 'ชื่อสินค้า Service Doc(Default ENG) 1
                tSKUNameSrvDoc2 = tSKUName2   'ชื่อสินค้า Service Doc (Local)2
                If cNet = 0 Then
                  cQty = 0
                  cQtyAll = 0
                Else
                  cQty = 1
                  cQtyAll = 1
                End If
                cRegPrice = cNet
                cSalePrice = cNet
                cSaleUnit = cNet
                cB4DisChg = cNet
                cSdtStkFac = cNet
                cSaleAmt = cNet
                cSaleAmt = cNet
                tSdtStaPdt = "5"
            
                tSql = "INSERT INTO " & tVB_CNTblDT
                tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                tSql = tSql & vbCrLf & " VALUES ("
                tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                tSql = tSql & vbCrLf & "," & cSalePrice
                tSql = tSql & vbCrLf & "," & cSaleAmt
                tSql = tSql & vbCrLf & "," & cSalePrice
                tSql = tSql & vbCrLf & "," & cSaleUnit
                tSql = tSql & vbCrLf & "," & cB4DisChg
                tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                tSql = tSql & vbCrLf & "," & cSdtStkFac
                tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"           '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                If SP_SQLbExecute(tSql, , , nEff) Then
                    If nEff = 1 Then
                       'none
                    End If
                End If
                                
                'Foreign Currency
                With ogdFunc1
                       For nI = .FixedRows To .rows - 1
                            nR = nR + 1
                            cNet = 0
                            cReg = 0
                            cRate = 0
                             If Not (orsPickup.BOF And orsPickup.EOF) Then
                                orsPickup.Filter = "FTSdtStaPdt = '3' AND FTSkuCode = '" & .TextMatrix(nI, 5) & "'"
                                If Not (orsPickup.BOF And orsPickup.EOF) Then
                                    Do While Not orsPickup.EOF
                                        cNet = cNet + SP_FEDvChkNumeric(orsPickup, "FCSum")
                                        cReg = cReg + SP_FEDvChkNumeric(orsPickup, "FCSum2")
                                        cRate = SP_FEDvChkNumeric(orsPickup, "FCSdtStkFac")
                                        orsPickup.MoveNext
                                    Loop
                                End If
                                orsPickup.Filter = adFilterNone
                            End If
                           
                            'tUT_SKUCode = .TextMatrix(ni, 4)
                            tUT_SKUCode = .TextMatrix(nI, 5)    '*BG 56-07-10   CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                            tSKUName = .TextMatrix(nI, 1) 'ชื่อสินค้า(Default ENG)1
                            tSKUName2 = .TextMatrix(nI, 5)   'ชื่อสินค้า(Local)2
                            tSKUNameSrvDoc = tSKUName 'ชื่อสินค้า Service Doc(Default ENG) 1
                            tSKUNameSrvDoc2 = tSKUName2   'ชื่อสินค้า Service Doc (Local)2
                            
                            If cNet = 0 Then
                              cQty = 0
                              cQtyAll = 0
                            Else
                              cQty = 1
                              cQtyAll = 1
                            End If
                            cRegPrice = cReg
                            cSalePrice = cReg
                            cSaleUnit = cReg
                            cB4DisChg = cReg
'                            cSdtStkFac = cRate             '*Ao 2013/10/14   V 4.0.79
                            cSdtStkFac = CDbl(Val(Format(.ValueMatrix(nI, 7)))) 'ทศนิยม 5 ตำแหน่ง
                            cSaleAmt = cNet
                            tSdtStaPdt = "5"
                        
                            tSql = "INSERT INTO " & tVB_CNTblDT
                            tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                            tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                            tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                            tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                            tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                            tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                            tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                            tSql = tSql & vbCrLf & " VALUES ("
                            tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                            tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                            tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                            tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                            tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                            tSql = tSql & vbCrLf & "," & cSalePrice
                            tSql = tSql & vbCrLf & "," & cSaleAmt
                            tSql = tSql & vbCrLf & "," & cSalePrice
                            tSql = tSql & vbCrLf & "," & cSaleUnit
                            tSql = tSql & vbCrLf & "," & cB4DisChg
                            tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                            tSql = tSql & vbCrLf & "," & cSdtStkFac
                            tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"           '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                            tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                            If SP_SQLbExecute(tSql, , , nEff) Then
                                If nEff = 1 Then
                                   'none
                                End If
                            End If
                       Next nI
                End With
                
                '----------------------------------------------------------------------------------------------------
                '*Ao 2013/09/23  CommSheet TK-ISS3000-305-UAT_00.00.03 FSBT NG
                orsPickup.Filter = adFilterNone
                orsPickup.Filter = "FTSdtStaPdt = '3' AND FTSkuCode = 'T012'"
                If Not (orsPickup.BOF And orsPickup.EOF) Then
                    nR = nR + 1
                    cNet = 0
                    cReg = 0
'                    cRate = 0
                    orsPickup.MoveFirst
                    Do While Not orsPickup.EOF
                        cNet = cNet + SP_FEDvChkNumeric(orsPickup, "FCSum")
                        cReg = cReg + SP_FEDvChkNumeric(orsPickup, "FCSum2")
'                        cRate = SP_FEDvChkNumeric(orsPickup, "FCSdtStkFac")
                        orsPickup.MoveNext
                    Loop
                Else
                    nR = nR + 1
                    cNet = 0
                    cReg = 0
'                    cRate = 0
                End If
                    
                If True Then                        '*Ao 2013/10/09    เก็บตลอด
                    If cNet = 0 Then
                      cQty = 0
                      cQtyAll = 0
                    Else
                      cQty = 1
                      cQtyAll = 1
                    End If
                    
                    tUT_SKUCode = "T012"
                    tSKUName = "Foreign(Base)"
                    tSKUName2 = "Foreign(Base)"
                    tSKUNameSrvDoc = "Foreign(Base)"
                    tSKUNameSrvDoc2 = "Foreign(Base)"
                            
                    cRegPrice = cReg
                    cSalePrice = cReg
                    cSaleUnit = cReg
                    cB4DisChg = cReg
'                    cSdtStkFac = cRate                 '*Ao 2013/10/14
                    cSdtStkFac = 1
                    cSaleAmt = cNet
                    tSdtStaPdt = "5"
                    
                    tSql = "INSERT INTO " & tVB_CNTblDT
                    tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode"
                    tSql = tSql & vbCrLf & ",FTSkuAbbName,FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc"
                    tSql = tSql & vbCrLf & ",FTShdTransType,FDShdTransDate,FTShdTransTime,FCSdtQty,FCSdtQtyAll"
                    tSql = tSql & vbCrLf & ",FCSdtSalePrice,FCSdtSaleAmt,FCSdtRegPrice,FCSdtSaleUnit,FCSdtB4DisChg"
                    tSql = tSql & vbCrLf & ",FTSdtStaPdt,FCSdtStkFac"
                    tSql = tSql & vbCrLf & ",FTSdtBarCode"                      '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & "," & tCS_CNLastFedIns
                    tSql = tSql & vbCrLf & " VALUES ("
                    tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nR & ",'" & tUT_SKUCode & "'"
                    tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "'"
                    tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
                    tSql = tSql & vbCrLf & ",'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & Time$ & "'"
                    tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cQty, True, False) & "," & SP_FMTcCurPoint(cQtyAll, True, False)
                    tSql = tSql & vbCrLf & "," & cSalePrice
                    tSql = tSql & vbCrLf & "," & cSaleAmt
                    tSql = tSql & vbCrLf & "," & cSalePrice
                    tSql = tSql & vbCrLf & "," & cSaleUnit
                    tSql = tSql & vbCrLf & "," & cB4DisChg
                    tSql = tSql & vbCrLf & ",'" & tSdtStaPdt & "'"
                    tSql = tSql & vbCrLf & "," & cSdtStkFac
                    tSql = tSql & vbCrLf & ",'" & tUT_SKUCode & "'"           '*Ao 2013/05/30  CommSheet TK-ISS3000-208-ST 00 00 01-Ao.xls
                    tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
                    If SP_SQLbExecute(tSql, , , nEff) Then
                        If nEff = 1 Then
                           'none
                        End If
                    End If
                       
                End If
                '----------------------------------------------------------------------------------------------------
            End If
    End Select
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsPickup)
   Exit Sub
ErrHandler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsPickup)
End Sub

'*Tao 62-07-05 เพิ่ม Vendor Voucher
Private Sub ogdVendorVchr_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_KeyPress = ""
End Sub

Private Sub ogdVendorVchr_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With ogdVendorVchr
        Select Case Col
            Case 2
                .EditMaxLength = 5
            Case 3
                .EditMaxLength = 11
            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub ogdVendorVchr_ChangeEdit()
    With ogdVendorVchr
        Select Case .Col
            Case 2
                .EditText = tW_KeyPress
            Case 3
                .EditText = SP_DEFtChangeNum(tW_KeyPress)
        End Select
    End With
End Sub

Private Sub ogdVendorVchr_GotFocus()
    With ogdVendorVchr
        .HighLight = flexHighlightAlways
        .SelectionMode = flexHighlightWithFocus
        .Editable = flexEDKbdMouse
         tW_KeyPress = ""
    End With
End Sub

Private Sub ogdVendorVchr_KeyDown(KeyCode As Integer, Shift As Integer)
    With ogdVendorVchr
        Select Case KeyCode
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    If .EditText <> "" Then
                        ogdFunc1.SetFocus
                        ogdFunc1.Select 0, 3
                    End If
                Else
                    If .EditText <> "" Then
                        ogdFunc.SetFocus
                        ogdFunc.Select .Row + 1, 2
                    End If
                End If
            Case vbKeyDown
               If .Row = .rows - 1 Then
                    ogdFunc.SetFocus
                    ogdFunc.Select 0, 3
                End If
            Case vbKeyUp
                If .Row = 0 Then
                    ogdGVFunc.SetFocus
                    ogdGVFunc.Select ogdGVFunc.rows - 1, 3
                End If
        End Select
    End With
End Sub

Private Sub ogdVendorVchr_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With ogdVendorVchr
        Select Case KeyAscii
            Case 13
                tW_KeyPress = ""
                If .Row = .rows - 1 Then
                    Select Case Col
                    Case 2
                        ogdVendorVchr.Select Row, Col + 1, Row, Col + 1
                    Case 3
                        'ogdVendorVchr.Select Row, Col - 1, Row, Col - 1
                    End Select
                Else
                    If .EditText <> "" Then
                        ogdVendorVchr.SetFocus
                        ogdVendorVchr.Select Row + 1, 2
                    End If
                End If
            Case Else
                Select Case KeyAscii
                    Case vbKeyDelete
                        tW_KeyPress = ""
                    Case vbKeyBack
                        If Len(Trim(tW_KeyPress)) > 0 Then
                            If Len(Trim(tW_KeyPress)) > 11 Then
                                tW_KeyPress = Left(tW_KeyPress, 9)
                            Else
                                tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
                            End If
                        End If
                        With ogdVendorVchr
                            Select Case .Col
                                Case 2
                                    .Text = tW_KeyPress
                            End Select
                        End With
                    Case Else
                        Call cUT.UT_KBDKeyCurrency(KeyAscii, tW_KeyPress)
                        If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii))
                End Select
        End Select
    End With
End Sub

Private Sub ogdVendorVchr_LostFocus()
    With ogdVendorVchr
        .SelectionMode = flexSelectionFree
        .Editable = flexEDNone
        .HighLight = flexHighlightNever
    End With
End Sub

