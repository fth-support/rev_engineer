VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form wMainRet640 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "wMainRet640"
   ClientHeight    =   6780
   ClientLeft      =   2145
   ClientTop       =   2025
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox opbForm2 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   9585
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   9585
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   735
         Index           =   4
         Left            =   4875
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   15
         Width           =   1815
         _Version        =   720897
         _ExtentX        =   3201
         _ExtentY        =   1296
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin VB.Label olaOnOff 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "On-line"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   300
            Left            =   105
            TabIndex        =   24
            Top             =   45
            Width           =   1275
         End
         Begin VB.Image oimOnOff 
            Height          =   360
            Left            =   1380
            Picture         =   "wMainRet640.frx":0000
            Top             =   40
            Width           =   360
         End
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   735
         Index           =   5
         Left            =   6675
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   15
         Width           =   1455
         _Version        =   720897
         _ExtentX        =   2566
         _ExtentY        =   1296
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin VB.Image oimToolsBar 
            Height          =   240
            Index           =   0
            Left            =   75
            Picture         =   "wMainRet640.frx":06EA
            Stretch         =   -1  'True
            Tag             =   "2;เครื่องจุดขาย;POS Machine"
            Top             =   110
            Width           =   240
         End
         Begin VB.Label olaPos 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00001"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   345
            Left            =   450
            TabIndex        =   26
            Top             =   45
            Width           =   915
         End
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   735
         Index           =   8
         Left            =   8115
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   15
         Width           =   1465
         _Version        =   720897
         _ExtentX        =   2584
         _ExtentY        =   1296
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin VB.Image oimToolsBar 
            Height          =   240
            Index           =   1
            Left            =   60
            Picture         =   "wMainRet640.frx":6F3C
            Stretch         =   -1  'True
            Tag             =   "2;เวลาขาย;Sale Time"
            Top             =   110
            Width           =   240
         End
         Begin VB.Label olaDocNo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   345
            Left            =   480
            TabIndex        =   28
            Top             =   45
            Width           =   720
         End
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   735
         Index           =   3
         Left            =   0
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   15
         Width           =   4890
         _Version        =   720897
         _ExtentX        =   8625
         _ExtentY        =   1296
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin VB.Label olaIndiCator 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Sale"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   300
            Left            =   4200
            TabIndex        =   50
            Top             =   75
            Width           =   525
         End
         Begin VB.Label olaTransName 
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Sale"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   300
            Left            =   120
            TabIndex        =   22
            Top             =   75
            Width           =   4365
         End
         Begin VB.Label olaTrainingMode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Training Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   300
            Left            =   2280
            TabIndex        =   21
            Top             =   75
            Visible         =   0   'False
            Width           =   2475
         End
      End
   End
   Begin VB.PictureBox opbForm 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   0
      ScaleHeight     =   1080
      ScaleWidth      =   9585
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   9585
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   1335
         Index           =   0
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   -30
         Width           =   9580
         _Version        =   720897
         _ExtentX        =   16898
         _ExtentY        =   2355
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin MSComctlLib.ImageList oilOnOff 
            Left            =   1800
            Top             =   -360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   24
            ImageHeight     =   24
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   7
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "wMainRet640.frx":735E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "wMainRet640.frx":7A58
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "wMainRet640.frx":8152
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "wMainRet640.frx":884C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "wMainRet640.frx":F0AE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "wMainRet640.frx":15910
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "wMainRet640.frx":1C172
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Image oimLogo 
            Height          =   870
            Left            =   120
            Picture         =   "wMainRet640.frx":229D4
            Stretch         =   -1  'True
            Top             =   90
            Width           =   1305
         End
         Begin VB.Label olaBarcode 
            AutoSize        =   -1  'True
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            Caption         =   "8851907130541"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   1755
            TabIndex        =   18
            Top             =   60
            Width           =   3510
         End
         Begin VB.Label olaPrice 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            Caption         =   "132.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   960
            Left            =   5040
            TabIndex        =   17
            Top             =   165
            Width           =   4380
         End
         Begin VB.Label olaSKUName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Name."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Left            =   1755
            TabIndex        =   16
            Top             =   615
            Width           =   2085
         End
         Begin VB.Label olaChange 
            AutoSize        =   -1  'True
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            Caption         =   "เงินทอน"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   555
            Left            =   2640
            TabIndex        =   15
            Tag             =   "2;เงินทอน;Change"
            Top             =   320
            Visible         =   0   'False
            Width           =   1590
         End
      End
   End
   Begin VB.Timer otmLock 
      Interval        =   1500
      Left            =   1560
      Top             =   5640
   End
   Begin MSComctlLib.StatusBar ostForm 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6435
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "Ready..."
            TextSave        =   "Ready..."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3405
            MinWidth        =   3405
            Text            =   "TheMall/AdaDB"
            TextSave        =   "TheMall/AdaDB"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3405
            MinWidth        =   3405
            Text            =   "User : 009/Manager"
            TextSave        =   "User : 009/Manager"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "25/12/2024"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "11:03"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   970
            MinWidth        =   970
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   970
            MinWidth        =   970
         EndProperty
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdForm 
      Height          =   3350
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   9580
      _cx             =   16898
      _cy             =   5909
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      ForeColorFixed  =   16777215
      BackColorSel    =   11889258
      ForeColorSel    =   16777215
      BackColorBkg    =   16768725
      BackColorAlternate=   16768725
      GridColor       =   33023
      GridColorFixed  =   16768725
      TreeColor       =   16768725
      FloodColor      =   192
      SheetBorder     =   16768725
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   4
      GridLineWidth   =   1
      Rows            =   11
      Cols            =   33
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wMainRet640.frx":265E6
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
      WallPaperAlignment=   4
      Begin VB.PictureBox opbWeight 
         BorderStyle     =   0  'None
         Height          =   2400
         Left            =   2520
         ScaleHeight     =   2400
         ScaleWidth      =   5100
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   280
         Visible         =   0   'False
         Width           =   5100
         Begin XtremeSuiteControls.TabControl otcForm1 
            Height          =   2655
            Index           =   21
            Left            =   0
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   -255
            Width           =   5085
            _Version        =   720897
            _ExtentX        =   8969
            _ExtentY        =   4683
            _StockProps     =   64
            Appearance      =   10
            Color           =   -1484272203
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "59.00"
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
               Index           =   8
               Left            =   2160
               TabIndex        =   12
               Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
               Top             =   2160
               Width           =   690
            End
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0.434"
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
               Index           =   7
               Left            =   1800
               TabIndex        =   11
               Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
               Top             =   1560
               Width           =   690
            End
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ลูกตาลอ่อน"
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
               Index           =   6
               Left            =   2520
               TabIndex        =   10
               Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
               Top             =   960
               Width           =   1710
            End
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "202133002550004342"
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
               Index           =   5
               Left            =   1920
               TabIndex        =   9
               Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
               Top             =   360
               Width           =   2700
            End
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sub Total :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   222
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Index           =   4
               Left            =   480
               TabIndex        =   8
               Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
               Top             =   2160
               Width           =   1560
            End
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Weight :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   222
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Index           =   3
               Left            =   480
               TabIndex        =   7
               Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
               Top             =   1560
               Width           =   1230
            End
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Items Name :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   222
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Index           =   2
               Left            =   480
               TabIndex        =   6
               Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
               Top             =   960
               Width           =   1950
            End
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Barcode :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   222
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Index           =   0
               Left            =   480
               TabIndex        =   5
               Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
               Top             =   360
               Width           =   1350
            End
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   4
               Tag             =   "2;ชื่อสินค้า;Name"
               Top             =   1440
               Width           =   45
            End
         End
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2055
      Index           =   24
      Left            =   -120
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4900
      Width           =   9700
      _Version        =   720897
      _ExtentX        =   17110
      _ExtentY        =   3625
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.TextBox otbPrompt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         MaxLength       =   20
         TabIndex        =   30
         Top             =   60
         Width           =   2880
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   1335
         Index           =   2
         Left            =   135
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   525
         Width           =   6300
         _Version        =   720897
         _ExtentX        =   11112
         _ExtentY        =   2355
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin XtremeSuiteControls.TabControl otcForm1 
            Height          =   1310
            Index           =   19
            Left            =   1245
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   0
            Width           =   5055
            _Version        =   720897
            _ExtentX        =   8916
            _ExtentY        =   2311
            _StockProps     =   64
            Appearance      =   10
            Color           =   -1484272203
            PaintManager.Position=   2
            Begin VB.Label olaDisplay 
               BackStyle       =   0  'Transparent
               Caption         =   "^ 12-25 JULY 2007  HAPPY ENDING SALE  ฉลองเปิดเดอะมอลล์ ท่าพระ ^  โฉมใหม่ "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   240
               Index           =   1
               Left            =   60
               TabIndex        =   35
               Top             =   420
               Width           =   4950
            End
            Begin VB.Label olaDisplay 
               BackStyle       =   0  'Transparent
               Caption         =   "^ 12-25 JULY 2007  HAPPY ENDING SALE  ฉลองเปิดเดอะมอลล์"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   240
               Index           =   2
               Left            =   60
               TabIndex        =   34
               Top             =   720
               Width           =   4950
            End
            Begin VB.Label olaDisplay 
               BackStyle       =   0  'Transparent
               Caption         =   "^ 12-25 JULY 2007  HAPPY ENDING SALE  ฉลองเปิดเดอะมอลล์ ท่าพระ ^  โฉมใหม่ "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   222
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   240
               Index           =   0
               Left            =   60
               TabIndex        =   33
               Top             =   120
               Width           =   4950
            End
         End
         Begin XtremeSuiteControls.TabControl otcForm1 
            Height          =   1310
            Index           =   20
            Left            =   0
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   0
            Width           =   1260
            _Version        =   720897
            _ExtentX        =   2222
            _ExtentY        =   2311
            _StockProps     =   64
            Appearance      =   10
            Color           =   -1484272203
            PaintManager.Position=   2
            Begin SCANNERLib.Scanner oScanner 
               Left            =   120
               Top             =   480
               _Version        =   65536
               _ExtentX        =   741
               _ExtentY        =   741
               _StockProps     =   0
            End
            Begin VB.Timer otmDisplay 
               Interval        =   33000
               Left            =   120
               Top             =   0
            End
            Begin VB.Label olaFn 
               BackStyle       =   0  'Transparent
               Caption         =   "Function"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   222
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   38
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label olaFn 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Fn"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   495
               Index           =   1
               Left            =   120
               TabIndex        =   37
               Top             =   480
               Width           =   1215
            End
         End
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   1775
         Index           =   22
         Left            =   6405
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   60
         Width           =   3300
         _Version        =   720897
         _ExtentX        =   5821
         _ExtentY        =   3131
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin VB.Label olaTSubReg 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   2505
            TabIndex        =   45
            Top             =   45
            Width           =   630
         End
         Begin VB.Label olbSubReg 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Reg."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Left            =   90
            TabIndex        =   44
            Tag             =   "SubTotal;ยอดคงเหลือ"
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label olbSubPro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Pro."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Left            =   90
            TabIndex        =   43
            Tag             =   "SubTotal;ยอดคงเหลือ"
            Top             =   480
            Width           =   1080
         End
         Begin VB.Label olaTSubPro 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   2505
            TabIndex        =   42
            Top             =   480
            Width           =   630
         End
         Begin VB.Label olbSubTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   315
            Left            =   90
            TabIndex        =   41
            Tag             =   "SubTotal;ยอดคงเหลือ"
            Top             =   945
            Width           =   1290
         End
         Begin VB.Label olaTSubTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   2505
            TabIndex        =   40
            Top             =   945
            Width           =   630
         End
      End
      Begin XtremeSuiteControls.TabControl otcForm1 
         Height          =   735
         Index           =   23
         Left            =   4440
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   60
         Width           =   1980
         _Version        =   720897
         _ExtentX        =   3492
         _ExtentY        =   1296
         _StockProps     =   64
         Appearance      =   10
         Color           =   -1484272203
         PaintManager.Position=   2
         Begin VB.Label olbQtySumItem 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty.(Item)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   48
            Tag             =   "SubTotal;ยอดคงเหลือ"
            Top             =   45
            Width           =   1035
         End
         Begin VB.Label olaSumQty 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1635
            TabIndex        =   47
            Top             =   0
            Width           =   150
         End
      End
      Begin VB.Label olaRef 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   49
         Top             =   60
         Width           =   1215
      End
   End
   Begin VB.Menu omnLogo 
      Caption         =   "Logo"
      Visible         =   0   'False
      Begin VB.Menu omnLogo1 
         Caption         =   "Change"
      End
      Begin VB.Menu omnLogo2 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "wMainRet640"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*Ao 52-02-25
' ถ้ามีการแก้ Code ใน From wMainRet ให้แก้ใน From wMainRet640 ด้วย
'wMainRet ใช้ Resorution 800*600
'wMainRet640 ใช้ Resorution 640*480
'________________________________________________________
Private orsSKU As ADODB.Recordset
Public tW_KeyPress As String 'เก็บ ราคาที่ Key  ในการกดแป้นพิมพ์ที่ grid
Private tW_KeyCode As String 'เก็บ KeyCode ในการกดแป้นพิมพ์ที่ grid
Const tW_Fed = "FNSdtSeqNo,FTPmt, FTSymbol,FTSdtBarCode,FTSkuAbbName,FNSdtQty,FCSdtRegPrice,FCSdtSaleUnit,FCSdtSaleAmt,FCSdtQty" & _
                                ",FCSdtQtyAll,FTSdtUnitName,FCSdtDis,FCSdtChg,FTSkuCode" & _
                                ",FTTaxCode,FCSdtTax,FTSdtWeight ,FCSdtVat,FCSdtVatable" & _
                                ",FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc" & _
                                ",FTTaxSign,FTTaxStaSign,FCSdtSalePrice,FTSdtPriceOverideFlag,FTSdtCancelFlag,FTSdtStaPdt,FTSdtStaRfd" & _
                                ",FTSdtConsignFlag,FTSdtStaPreview,FTSdtStaNotAlwPrv"
Private nW_QtyItem As Long
Private cW_SaleItem As Double
Private dW_TimeDisplay As Date '*KT 52-12-19 POSFront 2.0.30 แก้ไขให้กาทำงานของ Timer ของการ Display ถูกต้องมากขึ้น เก็บเวลาเริ่มต้นจับเวลาการ Display

Public bW_FirstLoad As Boolean 'ตัวแปรบอกว่ามีการเปิดหน้าจอครั้งแรก True : เปิดครั้งแรก False : ไม่ใช่ครั้งแรก
Public nW_GridEnter As Long
Public bW_ChkDecimals As Boolean '*KT 52-11-26 PH 3.0-ML-SF-034  กด Price Override เป็นเศษสตางค์  เครื่องมีข้อความเตื่อน แต่เมื่อกด Enter แล้วใส่จำนวนเงินใหม่ เมื่อพิมพ์ Receipt จะพิมพ์ค่าที่ป้อนผิด
Public bW_GridEnter As Boolean '*KT 52-12-01 หลังเคาะ Enter ห้ามให้มีการ key อะไรจนกว่าจะรอให้ทำงานเสร็จก่อนแล้วสามารถ Key ได้ถ้ามีการแจ้งเตือน
Public nW_CntSentSale As Long '*Tao 61-09-24 นับเวลาที่จะ SentSale

'Public Sub W_SETxIndicator(ByVal ptType As Integer, Optional ByVal ptTransType As String)
''-----------------------------------------------------------
''   Parameter :
''   Cmt :
''   File Refer :
''   Developer : BUNG
''   Date Create : 30/04/2014 :  17:53:32
''-----------------------------------------------------------
'    Dim tTrnType As String
'    Dim tTypePrint As String
'    Dim tSql As String
'    Dim orsTemp As ADODB.Recordset
'    Dim tWhere As String
'
'    On Error GoTo ErrHandle
'
'        If ptTransType = "" Then
''            tTrnType = tVB_Transaction                                  '*Ao 2014/06/20
'            tTrnType = Format(tVB_Transaction, "00")
'        Else
'            tTrnType = Format(ptTransType, "00")
'        End If
'
'        Select Case ptType
'            Case 0 'SET
''                'Get Data
'                tSql = "SELECT FTSttTranCode,FTSttDesc,FTSttABB,FTSttPrintType FROM TSysTransType"
'                tSql = tSql & vbCrLf & " WHERE FTSttTranCode = '" & tTrnType & "'"
'                Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
'                If Not (orsTemp.EOF And orsTemp.BOF) Then
'                    tTypePrint = SP_FEDtChkString(orsTemp, "FTSttPrintType")
'                End If
'
'                'Show Data
'                    '*BG 57-05-02
'                Select Case tTypePrint
'                    Case "0"
'                        olaIndiCator.Caption = "S"
'                    Case "1"
'                        olaIndiCator.Caption = "T"
'                    Case "2"
'                        olaIndiCator.Caption = "A"
'                End Select
'
'            Case 1 'Change Shift + F7
'                 If UCase(tVB_CNTypePrinter) = UCase("2") Then
'                        Select Case tVB_CNTypePrinter
'                            Case "0"  'S
'                                tTrnType = "A"
'                                tVB_CNTypePrinter = "2"
'                            Case "2"  'A
'                                If Trim(olaIndiCator.Caption) = "A" Then
'                                    tTrnType = "S"
'                                Else
'                                    tTrnType = "A"
'                                End If
'                            Case Else
'                                tTrnType = "T"
'                        End Select
'                        olaIndiCator.Caption = tTrnType
'                Else
'                    Call SP_MSGxManualShow(tMS_CN522, vbExclamation + vbOKOnly)
'                End If
'            Case 2
'                olaIndiCator.Caption = vbNullString
'                Select Case tVB_CNTypePrinter
'                   Case "0"
'                       olaIndiCator.Caption = "S"
'                   Case "1"
'                       olaIndiCator.Caption = "T"
'                   Case "2"
'                       olaIndiCator.Caption = "A"
'                End Select
'        End Select
'
'    Exit Sub
'ErrHandle:
'
'End Sub
Public Function w_GetxSumSdtSaleAmt() As Double
    Dim nMRow As Long
    Dim i As Long
    Dim cSumZ As Double
    
    cSumZ = 0
    With ogdForm
        nMRow = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FNSdtSeqNo"), .rows - 1, .ColIndex("FNSdtSeqNo"))
        For i = 1 To nMRow
            cSumZ = cSumZ + 1
'            If .TextMatrix(i, .ColIndex("FTSdtCancelFlag")) <> "1" Then
'                cSumZ = cSumZ + .ValueMatrix(i, .ColIndex("FCSdtSaleAmt"))
'            End If
        Next i
    End With
    w_GetxSumSdtSaleAmt = cSumZ
End Function
Public Sub W_SETxIndicator(ByVal ptTransType As String, Optional ByVal pbChk As Boolean = False)
'-----------------------------------------------------------
'   Call :
'   Cmt :   '*Ao 2014/06/25
'-----------------------------------------------------------
Dim tTrnType As String, tIndiCator As String, tSql As String, tTypePrint As String
Dim orsTemp As ADODB.Recordset

On Error GoTo ErrHandle
        '0 = S   =  Slip
        '1 = T = Printer
        '2 = A = Auto
        
    If ptTransType = "" Then
        tTrnType = Format(tVB_TransactionSub, "00")
    Else
        tTrnType = Format(ptTransType, "00")
    End If
        
        If pbChk = True Then            'HOT KEY
            tIndiCator = Trim(olaIndiCator.Caption)
            If UCase(tIndiCator) = UCase("T") Then
                Call SP_MSGxManualShow(tMS_CN522, vbExclamation + vbOKOnly, tVB_CNAppTitle)
                Exit Sub
            End If
            tSql = "SELECT FTSttTranCode,FTSttDesc,FTSttABB,FTSttPrintType FROM TSysTransType"
            tSql = tSql & vbCrLf & "WHERE FTSttTranCode = '" & tTrnType & "'"
            Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
            If Not (orsTemp.EOF And orsTemp.BOF) Then
                tTypePrint = SP_FEDtChkString(orsTemp, "FTSttPrintType")
            End If
            If tTypePrint = "" Then
                olaIndiCator.Caption = "T"
                tVB_CNTypePrinter = "1"
            Else
                 If UCase(tTypePrint) = UCase("0") Or UCase(tTypePrint) = UCase("1") Then
                    Call SP_MSGxManualShow(tMS_CN522, vbExclamation + vbOKOnly, tVB_CNAppTitle)
                    Exit Sub
                End If
                If UCase(tIndiCator) = UCase("A") Then
                    olaIndiCator.Caption = "S"
                    tVB_CNTypePrinter = "0"
                Else
                    olaIndiCator.Caption = "A"
                    tVB_CNTypePrinter = "2"
                End If
            End If
        Else
            olaIndiCator.Caption = vbNullString
            Select Case tVB_CNTypePrinter
               Case "0"
                   olaIndiCator.Caption = "S"
               Case "1"
                   olaIndiCator.Caption = "T"
               Case "2"
                   olaIndiCator.Caption = "A"
            End Select
        End If
    Exit Sub
ErrHandle:

End Sub

Private Sub W_GETxSlipPrintType(ByVal ptTransType As String)
'________________________________________
'   Call:
'   Cmt: '*Eaw 56-09-04  ปรับ Print Type ให้ตรงกับที่ตั้งค่าไว้ '*Eaw 56-09-04 CommSheet TK-ISS3000-338 UAT FAPL-074.xlsx
'________________________________________
Dim tSql As String
Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT FTSttTranCode,FTSttDesc,FTSttABB,FTSttPrintType FROM TSysTransType"
    tSql = tSql & vbCrLf & " WHERE FTSttTranCode = '" & ptTransType & "'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
    If Not (orsTemp.EOF And orsTemp.BOF) Then
        tVB_CNTypePrinter = SP_FEDtChkString(orsTemp, "FTSttPrintType")
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGShowError
End Sub

Public Sub W_CHKxHHDSpace()
'-------------------------------------------------------------------------------
'   Call:
'   Cmt: '*Bump 56-03-08 TK-ISS3000-033 ตรวจสอบพื้นที่ HDD
'-------------------------------------------------------------------------------
    Dim orsTmp As ADODB.Recordset
    Dim tSql As String
    Dim tMsg1, tMsg2 As String
    tMsg1 = "" '"Hard disk is low space!"
    tMsg2 = "" '"Please contact Administrator"
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        'Transaction ที่ไม่ต้องเช็ค พื้นที่ HHD
        Case tEN_TrnOpen, tEN_TrnStartBank, tEN_TrnClose, tEN_TrnNoSale '*Nomal Mode
        Case tEN_TrnCopyMaster, tEN_TrnCopyFromDev, tEN_TrnCopyTrans '* Admin Menu
        Case tEN_TrnClosingReportManualTax, tEN_TrnClosingReportTaxInc '*Re entry Mode
        Case Else
            If Not SP_CHKbHDDFreeSpace() Then
                ''tms_CN014042,'tms_CN014043  ทาง FSBT อาจจะให้เปลี่ยนข้อความที่จะแสดงให้เปลี่ยนที่ TsysLang แทน
                tSql = " SELECT FTDefLangPri FROM TSysLang  WHERE FTSlnKey  IN('tms_CN014042' ,'tms_CN014043') ORDER BY  FTSlnKey "
                If SP_SQLvExecute(tSql, orsTmp, oDB.ocnOffLine) = 0 Then
                'If SP_SQLvExecute(tSQL, orsTemp) <> 0 Then
                    If Not (orsTmp.BOF And orsTmp.EOF) Then
                        Do While Not orsTmp.EOF
                            If tMsg1 = "" Then
                                tMsg1 = SP_FEDtChkString(orsTmp, "FTDefLangPri")
                            Else
                                 tMsg2 = SP_FEDtChkString(orsTmp, "FTDefLangPri")
                            End If
                        orsTmp.MoveNext
                        Loop
                    End If
                    Call SP_DATxRsNothing(orsTmp)
                Else
                    tMsg1 = "Hard disk is low space!"
                    tMsg2 = "Please contact Administrator"
                End If
                
                If Not bVB_IsFirstMsgDiskSpace Then '*Bump เช็คป้องกันการแสดงข้อความซ้ำ 2 ครั้งเมื่อเปิดครั้งแรก
                    bVB_IsFirstMsgDiskSpace = True
                Else
                    MsgBox (tMsg1 & vbCrLf & tMsg2), vbExclamation
                End If
            End If
    End Select
    '01  Open Terminal  Transaction
    '2   Start Bank
    '3   Sale Transaction
    '04  Return Transaction
    '5   IPV Transaction
    '06  Layaway Create Transaction
    '07  Layaway Pickup Transaction
    '08  Close Terminal Transaction
    '9   Print Closing; Report
    '10  VIP Service Sale
    '11  Tax Except Sale
    '12  Cash Transfer
    '13  Duty Free Sale
    '14  IPV Return Transaction
    '15  Voucher
    '16  Layaway Cancel Transaction
    '17  IPV Layaway Pickup Transaction
    '18  nO Sale
    '19  Add Layaway Transaction
    '20  Cancel Add Layaway
    '21  Cancel Voucher
    '22  Other Receive
    '23  Other Disbursement
    '24  Deposit Money
    '25  Training Mode
    '26  IPV VIP Service Sale
    '27  IPV Tax Except Sale
    '28  IPV Duty Free Sale
    '29  Suspend Transaction
    '30  Create VIP
    '31  Total Report
    '32  IPV Other Receive
    '33  IPV Other Disbursement
    '34  Copy Transaction to Device
    '35  Copy Transaction From Device
    '36  Update Master
    '37  Over/Short Report
    '38  Sale
    '39  Return
    '40  Total Report
    '41  Over/Short Report
    '42  Closing Report
    '43  Sale
    '44  Return
    '45  Total Report
    '46  Over/Short Report
    '47  Closing Report
    'FTSttTranCode FTSttDesc
End Sub

Public Sub W_PRCxUpdateLineOver(ByVal pcDiscNewOrg As Double, ByVal pcSourceAmt As Double, ByVal pnDTSeqNo As String)
Dim tSql As String, tTaxCode As String
Dim cQty As Double, cTaxRate As Double
Dim cVat As Double, cVatable As Double                  '*Aoo 2012/10/05
Dim cNet As Double, cSalePrice As Double
On Error GoTo Err_Handler
    
    With ogdForm
        .TextMatrix(pnDTSeqNo, .ColIndex("FCSdtSaleUnit")) = SP_FMTcCurPoint(pcDiscNewOrg, False, True)
        .TextMatrix(pnDTSeqNo, .ColIndex("FCSdtSalePrice")) = .TextMatrix(pnDTSeqNo, .ColIndex("FCSdtSaleUnit"))
        .TextMatrix(pnDTSeqNo, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(.ValueMatrix(pnDTSeqNo, .ColIndex("FCSdtQty")) * pcDiscNewOrg, False, True)
        cQty = SP_VALcText2Double(.ValueMatrix(pnDTSeqNo, .ColIndex("FCSdtQty")))
        cSalePrice = SP_VALcText2Double(.ValueMatrix(pnDTSeqNo, .ColIndex("FCSdtSalePrice")))  'Sale Pric
        cNet = SP_FMTcCurPoint(cQty * cSalePrice, True, False)                         'รวมเงิน
        cTaxRate = SP_VALcText2Double(.ValueMatrix(pnDTSeqNo, .ColIndex("FCSdtTax")))
        tTaxCode = .TextMatrix(pnDTSeqNo, .ColIndex("FTTaxCode"))
'        cVat = IIf(tTaxCode <> "00", SP_FMTcCurPoint(SP_CALcVatValue(IIf(tVB_CNRetVATInOrEx = "1", True, False), cNet, cTaxRate), True, False), 0)                       '1:vat, 2:novat (0)
        cVat = SP_FMTcCurPoint(SP_VALcText2Double(cNet * (cTaxRate / (100 + cTaxRate))), True, False)
        cVatable = IIf(tTaxCode <> "00", cNet - cVat, 0)
        tSql = "UPDATE " & tVB_CNTblDT
'        tSQL = tSQL & vbCrLf & "SET FCLineOver = " & Abs(pcDiscNewOrg - pcSourceAmt)         '*Eaw 56-03-05
        tSql = tSql & vbCrLf & "SET FCLineOver = " & SP_FMTcCurPoint(pcDiscNewOrg, True, False)    '*Eaw 56-03-05
        tSql = tSql & vbCrLf & ",FCSdtSalePrice = " & SP_FMTcCurPoint(cSalePrice, True, False)
        tSql = tSql & vbCrLf & ",FCSdtSaleUnit = " & SP_FMTcCurPoint(cSalePrice, True, False)
'        tSQL = tSQL & vbCrLf & ",FCSdtB4DisChg = " & SP_FMTcCurPoint(cSalePrice, True, False)              '*Ao 2013/04/05   CommSheet TK-ISS3000-116 00.00.03 FSBT.xlsx
        tSql = tSql & vbCrLf & ",FCSdtSaleAmt = " & cNet
        tSql = tSql & vbCrLf & ",FCSdtVat = " & SP_FMTcCurPoint(cVat, True, False)
        tSql = tSql & vbCrLf & ",FCSdtVatable = " & SP_FMTcCurPoint(cVatable, True, False)
        tSql = tSql & vbCrLf & ",FCSdtSalePricePmt = " & SP_FMTcCurPoint(cSalePrice, True, False)
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnDeposit
                tSql = tSql & vbCrLf & ",FTSdtPriceOverideFlag = 'Y'"           '*Ao 2013/07/26  CommSheet 235-ST att addon.pdf
'            Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnIPVRedeem
        End Select
'        tSql = tSql & vbCrLf & ",FTSdtPriceOverideFlag = 'Y'"
        If .TextMatrix(pnDTSeqNo, .ColIndex("FTPmt")) = "*" Then
            tSql = tSql & " ,FCSdtBSubReg = " & 0
            tSql = tSql & " ,FCSdtBSubPro = " & cSalePrice
            tSql = tSql & " ,FCSdtBluePrice = " & 0
            tSql = tSql & " ,FCSdtYellowPrice = " & cSalePrice
        Else
            tSql = tSql & " ,FCSdtBSubReg = " & cSalePrice
            tSql = tSql & " ,FCSdtBSubPro = " & 0
            tSql = tSql & " ,FCSdtBluePrice = " & cSalePrice
            tSql = tSql & " ,FCSdtYellowPrice = " & 0
        End If
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & "AND FNSdtSeqNo = " & pnDTSeqNo
        Call SP_SQLbExecute(tSql)
    End With
Exit Sub
Err_Handler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Public Sub W_FRMxActivate()
'-----------------------------------------------------------
'Call:
'Cmt: ให้การทำงานในแต่ละ Transactiom ในหน้าจอหลัก
'-----------------------------------------------------------
On Error GoTo ErrHandle
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnOpen:
        Case tEN_TrnStartBank:
        Case tEN_TrnSale: Call W_PRCxSale
        Case tEN_TrnReturn: Call W_PRCxReturn
        Case tEN_TrnIPV: Call W_PRCxIPV
        Case tEN_TrnDeposit: Call W_PRCxDeposit
        Case tEN_TrnRedeem: Call W_PRCxRedeem
        Case tEN_TrnClose:
        Case tEN_TrnPrintClose:
        Case tEN_TrnVIPSale: Call W_PRCxVIPSale
        Case tEN_TrnTaxExceptSale: Call W_PRCxTaxExceptSale
        Case tEN_TrnPickUp: Call W_PRCxPickUp
        Case tEN_TrnDutyFreeSale: Call W_PRCxDutyFree
        Case tEN_TrnIPVReturn:
        Case tEN_TrnVoucher: Call W_PRCxVoucher
        Case tEN_TrnDepositCancel: Call W_PRCxDepositCancel
        Case tEN_TrnIPVRedeem:
        Case tEN_TrnNoSale: Call W_PRCxNoSale
        Case tEN_TrnAddDeposit: Call W_PRCxAddDeposit
        Case tEN_TrnAddDepositCancel: Call W_PRCxAddDepositCancel
        Case tEN_TrnCancelVoucher: Call W_PRCxCancelVoucher
        Case tEN_TrnOtherReceive: Call W_PRCxOtherReceive
        Case tEN_TrnOtherDisbursement: Call W_PRCxOtherDisb
        Case tEN_TrnDepositMoney: Call W_PRCxDepositMny
        Case tEN_TrnTraining: Call W_PRCxTrainningMode
        Case tEN_TrnIPVVIPService:
        Case tEN_TrnIPVTaxExceptSale:
        Case tEN_TrnIPVDutyFreeSale:
        Case tEN_TrnCreateVIP: Call W_PRCxCreateVIPSale
        Case tEN_TrnTotalReport: Call W_PRCxTotalReport
        Case tEN_TrnSaleTaxInc: Call W_PRCxSale
        Case tEN_TrnSaleManualTax: Call W_PRCxSale
        Case tEN_TrnReturnTaxInc: Call W_PRCxReturn
        Case tEN_TrnReturnManualTax: Call W_PRCxReturn
        Case tEN_TrnTotalReportTaxInc:    Call W_PRCxTotalReport
        Case tEN_TrnTotalReportManualTax: Call W_PRCxTotalReport
        Case tEN_TrnOpenTaxInc: '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
        Case tEN_TrnOpenManualTax: '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
    End Select
        
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub W_PRCxTotalReport()
'    wCNAllow.tW_ID = tVB_CNMnuCurrent
'    wCNAllow.Show vbModal
'    DoEvents
'    If tVB_CNUserAlwC <> "" Then
On Error GoTo ErrHandle
        bVB_CNSplash = True  '*KT 54-04-08 V.3.0.56 ป้องกันเครื่องค้างและโปรแกรมหลุด
        'Call SP_STAxSplash '*KT 54-04-08 V.3.0.56 ป้องกันเครื่องค้างและโปรแกรมหลุด
        Call wFunctionKB.W_PrintSignOut
        bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
        If bVB_ReEntryMode Then
             wCNReEntryMnu2.Show vbModal: Exit Sub
        Else
            wCNBegin.Show vbModal
        End If
    Exit Sub
ErrHandle:
    
End Sub

Public Sub W_PRCxReEntry()
    '--------------------------------------------------------
'Call:
'Cmt: Transaction Sale
'Record Saleman ID                    Allow = Y
'Record Sex and Nationality     Allow = Y
'Record Customer's age            Allow = Y
'--------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
   Call W_FRMxInitailTrans
    'Sex/Nation/Age/Saleman
    'Comsheet ML-SF-080
    '-----------------------------------------------------
    Call SP_ChkxAlwInputALL
    If bVB_CNAlwInputSpn Then
        'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
        With wMain.oScanner
             If bVB_OPOSScannerEnable Then
                 If .DeviceEnabled Then
                     .DeviceEnabled = False
                     .Release
                     .Close
                 End If
             End If
         End With
        wCNSaleManID.Show vbModal
        If bVB_OPOSScannerEnable Then
            With oScanner
                     If .DeviceEnabled = False Then
                         .Open tVB_Scanner
                        .Claim 0 'TeamPos2000
                         .DeviceEnabled = True
                         .DataEventEnabled = True
                         .AutoDisable = True
                         .ClearInput
                    End If
            End With
        End If
    End If
    '-----------------------------------------------------
   bVB_SetFocus = True
   
   Exit Sub
ErrHandle:
    
End Sub

Public Sub W_PRCxSale()
    '--------------------------------------------------------
'Call:
'Cmt: Transaction Sale
'Record Saleman ID                    Allow = Y
'Record Sex and Nationality     Allow = Y
'Record Customer's age            Allow = Y
'--------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    
    Call W_FRMxInitailTrans
    'Sex/Nation/Age/Saleman
    'Comsheet ML-SF-080
    '-----------------------------------------------------
    Call SP_ChkxAlwInputALL
'    If bVB_CNAlwInputSex Then      '*Aoo 2012/07/31  ทำเป็น Function
'         wCNDetailCstSex.Show vbModal
'    End If
'
'    If bVB_CNAlwInputNat Then
'        wCNDetailCst.Show vbModal
'    End If
'
'    If bVB_CNAlwInputAge Then
'        wCNAge.Show vbModal
'    End If

    
    If bVB_CNAlwInputSpn Then
        'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
        With wMain.oScanner
             If bVB_OPOSScannerEnable Then
                 If .DeviceEnabled Then
                     .DeviceEnabled = False
                     .Release
                     .Close
                 End If
             End If
         End With
        wCNSaleManID.Show vbModal
        If bVB_OPOSScannerEnable Then
            With oScanner
                     If .DeviceEnabled = False Then
                         .Open tVB_Scanner
                        .Claim 0 'TeamPos2000
                         'Enable the device.
                         .DeviceEnabled = True
                         'Enable the event.
                         .DataEventEnabled = True
                         .AutoDisable = True
                         .ClearInput
                    End If
            End With
        End If
    End If
    '-----------------------------------------------------
   bVB_SetFocus = True
    '*Eaw 56-02-06 For Recovery Process
    If W_PRCbRecovery Then
        Call W_PRCxGetData
'        Call W_CHKxRCTmp  '*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
        Call W_CHKxCDTmp
        Call W_RECxProduct2Grid
        nVB_CNScollItem = ogdForm.Aggregate(flexSTMax, ogdForm.FixedRows, ogdForm.ColIndex("FNSdtSeqNo"), ogdForm.rows - 1, ogdForm.ColIndex("FNSdtSeqNo")) '*Tao 60-05-03 ถ้าขายค้างอยู่ แล้วถูกปิดโปรแกรม จะไม่สามารถ แก้ไขจำนวนและราคาได้
        Call SP_OBJxSetFocus(otbPrompt)
        If bVB_CNDiscount Then
            SleepEx 1000, 0
            Call otbPrompt_KeyDown(36, 0)           '*Ao 2014/05/07 PRJ10001TK-PH2-External Spec-Incomplete Transaction-00.02.04 FSBT
        End If
    End If
    '-----------------------------------------------------------------------
   Exit Sub
ErrHandle:
    
End Sub
    
Private Sub W_PRCxReturn()
'--------------------------------------------------------
'Call:
'Cmt: การคืนสินค้า มี 2 แบบคือ การคืนโดยไม่ต้องอ้างอิงใบ Ticket กับการคืนแบบการอ้างอิงใบ Ticket
'--------------------------------------------------------
On Error GoTo ErrHandle
    
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    If bVB_Reference Then
        wFunctionKB.W_KBRefund
        If Not bVB_Return Then
            If bVB_ReEntryMode Then
                 wCNReEntryMnu2.Show vbModal: Exit Sub
                 Exit Sub
            Else
                wCNBegin.Show vbModal
                Exit Sub
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_PRCxIPV()
'--------------------------------------------------------
'Call:
'Cmt: การ Void ทั้งบิลใบ Ticket ที่ทำรายการขายและคืน
'--------------------------------------------------------
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    wCNIPV.Show vbModal
    If Not bVB_IPV Then
        If bVB_ReEntryMode Then
             wCNReEntryMnu2.Show vbModal: Exit Sub
             Exit Sub
        Else
            wCNBegin.Show vbModal
            Exit Sub
        End If
    End If
    If bVB_IPV Then wFunctionKB.W_PRCxIPVTicket
    Call SP_OBJxSetFocus(otbPrompt)
'    If Not (GetInputState = 0) Then DoEvents '*KT 53-01-14 DoEvents
End Sub

Private Sub W_PRCxDeposit()
'--------------------------------------------------------
'Call:
'Cmt: Deposit Transaction (การทำเอกสารการมัดจำ)
'          *KT 52-11-16 PH3-ML-SF-013 Saleman ID on Deposit Tran
'--------------------------------------------------------
On Error GoTo ErrHandle
    Call W_FRMxInitailTrans
    
    Call SP_ChkxAlwInputALL
    
'    If bVB_CNAlwInputSex Then      '*Aoo 2012/07/31  ทำเป็น Function
'         wCNDetailCstSex.Show vbModal
'    End If
'
'    If bVB_CNAlwInputNat Then
'        wCNDetailCst.Show vbModal
'    End If
'
'    If bVB_CNAlwInputAge Then
'        wCNAge.Show vbModal
'    End If

    If bVB_CNAlwInputSpn Then
        'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
        With wMain.oScanner
             If bVB_OPOSScannerEnable Then
                 If .DeviceEnabled Then
                     .DeviceEnabled = False
                     .Release
                     .Close
                 End If
             End If
         End With
        wCNSaleManID.Show vbModal
        If bVB_OPOSScannerEnable Then
            With oScanner
                     If .DeviceEnabled = False Then
                         .Open tVB_Scanner
                        .Claim 0 'TeamPos2000
                         'Enable the device.
                         .DeviceEnabled = True
                         'Enable the event.
                         .DataEventEnabled = True
                         .AutoDisable = True
                         .ClearInput
                    End If
            End With
        End If
    End If
    '-----------------------------------------------------
   bVB_SetFocus = True
   Exit Sub
ErrHandle:
    
End Sub

Private Sub W_PRCxRedeem()
'--------------------------------------------------------
'Call:
'Cmt: Redeem Transaction (การชำระเงินมัดจำ)
'--------------------------------------------------------
On Error GoTo ErrHandle
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
    With wMain.oScanner
         If bVB_OPOSScannerEnable Then
             If .DeviceEnabled Then
                 .DeviceEnabled = False
                 .Release
                 .Close
             End If
         End If
     End With
    wCNDepositRedeem.Show vbModal
    If Not bVB_TrnRedeem Then wCNBegin.Show vbModal: Exit Sub
    wCNRedeemAmt.Show vbModal
    If Not bVB_TrnRedeem Then wCNBegin.Show vbModal: Exit Sub
    Call wFunctionKB.W_KBXHomeSUB
    Exit Sub
ErrHandle:
    
End Sub


Public Sub W_PRCxVoucher()
'--------------------------------------------------------
'Call:
'Cmt: Transaction Voucher Sale
'--------------------------------------------------------
On Error GoTo ErrHandle
    If bVB_TrnVoucher = False Then wCNBegin.Show vbModal: Exit Sub
    Call W_FRMxInitailTrans
    
    Call SP_ChkxAlwInputALL
    
'    If bVB_CNAlwInputNat Then                   '*Aoo 2012/07/31  ทำเป็น Function
'        wCNDetailCst.Show vbModal
'    End If
'    If bVB_CNAlwInputAge Then
'        wCNAge.Show vbModal
'    End If
    
    If bVB_CNAlwInputSpn Then
        'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
        With wMain.oScanner
             If bVB_OPOSScannerEnable Then
                 If .DeviceEnabled Then
                     .DeviceEnabled = False
                     .Release
                     .Close
                 End If
             End If
         End With
        wCNSaleManID.Show vbModal
    End If
    With wMain.oScanner
        If bVB_OPOSScannerEnable Then
            If .DeviceEnabled Then
                .DeviceEnabled = False
                .Release
                .Close
            End If
        End If
    End With
    Call wFunctionKB.W_KBXHomeSUB
    If bVB_OPOSScannerEnable Then
        With oScanner
                 If .DeviceEnabled = False Then
                     .Open tVB_Scanner
                    .Claim 0 'TeamPos2000
                     'Enable the device.
                     .DeviceEnabled = True
                     'Enable the event.
                     .DataEventEnabled = True
                     .AutoDisable = True
                     .ClearInput
                End If
        End With
    End If
     If Not bVB_AddVoucher Then wCNBegin.Show vbModal
     Exit Sub
ErrHandle:
    
End Sub

Private Sub W_PRCxNoSale()
'--------------------------------------------------------
'Call:
'Cmt: Transaction No Sale
'--------------------------------------------------------
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    wFunctionKB.W_KBXEndDrawer
    wCNBegin.Show vbModal: Exit Sub
End Sub

Private Sub W_PRCxDepositCancel()
'--------------------------------------------------------
'Call:
'Cmt: Deposit Cancel Transaction (การยกเลิกการมัดจำ)
'--------------------------------------------------------
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
    With wMain.oScanner
         If bVB_OPOSScannerEnable Then
             If .DeviceEnabled Then
                 .DeviceEnabled = False
                 .Release
                 .Close
             End If
         End If
     End With
     wCNDepositRedeem.Show vbModal
    If Not bVB_TrnDepositCancel Then wCNBegin.Show vbModal: Exit Sub
    wCNRedeemAmt.Show vbModal
    If Not bVB_TrnDepositCancel Then wCNBegin.Show vbModal: Exit Sub
    Call wFunctionKB.W_KBXHomeSUB
End Sub

Private Sub W_PRCxAddDepositCancel()
'--------------------------------------------------------
'Call:
'Cmt: Deposit Cancel Transaction (การยกเลิกการมัดจำ)
'--------------------------------------------------------
On Error GoTo ErrHandle
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
    With wMain.oScanner
         If bVB_OPOSScannerEnable Then
             If .DeviceEnabled Then
                 .DeviceEnabled = False
                 .Release
                 .Close
             End If
         End If
     End With
    wCNDepositRedeem.Show vbModal
    If Not bVB_TrnDepositCancel Then wCNBegin.Show vbModal: Exit Sub
    wCNRedeemAmt.Show vbModal
    If Not bVB_TrnDepositCancel Then wCNBegin.Show vbModal: Exit Sub
    Call wFunctionKB.W_KBXHomeSUB
    Exit Sub
ErrHandle:
    
End Sub

Public Sub W_PRCxCancelVoucher()
'--------------------------------------------------------
'Call:
'Cmt: Transaction Cancel Voucher
'--------------------------------------------------------
    Dim tDisplay$
On Error GoTo ErrHandle
    If bVB_TrnVoucher = False Then
        wCNBegin.Show vbModal
        Exit Sub
    End If
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    Call wFunctionKB.W_KBXHomeSUB
     If Not bVB_CancelVoucher Then
        If bVB_OPOSLineDisplayEnable Then
             If aVB_DSlipMag(0) = "" Or aVB_DSlipMag(1) = "" Then
                 tDisplay = "Welcome ....;"
            Else
                tDisplay = aVB_DSlipMag(0) & ";" & Trim(Mid(aVB_DSlipMag(1), 1, 20)) & Space(20 - Len(Trim(Mid(aVB_DSlipMag(1), 1, 20))))
            End If
            Call SP_PrnPollDisplay(tDisplay)
        End If
        wCNBegin.Show vbModal
        Exit Sub
    End If
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_PRCxAddDeposit()
'--------------------------------------------------------
'Call:
'Cmt: Deposit Cancel Transaction (การยกเลิกการมัดจำ)
'--------------------------------------------------------
On Error GoTo ErrHandle
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
    With wMain.oScanner
         If bVB_OPOSScannerEnable Then
             If .DeviceEnabled Then
                 .DeviceEnabled = False
                 .Release
                 .Close
             End If
         End If
     End With
    wCNDepositRedeem.Show vbModal
    If Not bVB_TrnAddDeposit Then wCNBegin.Show vbModal: Exit Sub
    wCNRedeemAmt.Show vbModal
    If Not bVB_TrnAddDeposit Then wCNBegin.Show vbModal: Exit Sub
    Call wFunctionKB.W_KBXHomeSUB
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_PRCxDepositMny()
'--------------------------------------------------------
'Call:
'Cmt:
'--------------------------------------------------------
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    wTenderMain.Show vbModal
End Sub

Public Sub W_PRCxOtherDisb()
'--------------------------------------------------------
'Call:
'Cmt: Transaction Other Disb
'--------------------------------------------------------
'    Call W_FRMxInitailTrans
'    wCNMainOther.Show vbModal
'    If Not bVB_TrnOthDisb Then wCNBegin.Show vbModal: Exit Sub
'    Call wFunctionKB.W_KBXHomeSUB
On Error GoTo ErrHandle
    Call W_FRMxInitailTrans
    '*KT 52-02-25 PH2-ML-SF-106 ในการทำ Other receive ไม่ขึ้นให้ใส่ Customer detail เช่น Age , Sex , Nationality etc…
    '-----------------------------------------------------
    Call SP_ChkxAlwInputALL
   
'    If bVB_CNAlwInputSex Then               '*Aoo 2012/07/31  ทำเป็น Function
'         wCNDetailCstSex.Show vbModal
'    End If
'
'    If bVB_CNAlwInputNat Then
'        wCNDetailCst.Show vbModal
'    End If
'
'    If bVB_CNAlwInputAge Then
'        wCNAge.Show vbModal
'    End If

    If bVB_CNAlwInputSpn Then
        'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
        With wMain.oScanner
             If bVB_OPOSScannerEnable Then
                 If .DeviceEnabled Then
                     .DeviceEnabled = False
                     .Release
                     .Close
                 End If
             End If
         End With
        wCNSaleManID.Show vbModal
        If bVB_OPOSScannerEnable Then
            With oScanner
                     If .DeviceEnabled = False Then
                         .Open tVB_Scanner
                        .Claim 0 'TeamPos2000
                         'Enable the device.
                         .DeviceEnabled = True
                         'Enable the event.
                         .DataEventEnabled = True
                         .AutoDisable = True
                         .ClearInput
                    End If
            End With
        End If
    End If
    '-----------------------------------------------------
    Exit Sub
ErrHandle:
    
End Sub

Public Sub W_PRCxOtherReceive()
'--------------------------------------------------------
'Call:
'Cmt: other Receive Transaction
'--------------------------------------------------------
On Error GoTo ErrHandle
'    Call W_FRMxInitailTrans
'    wCNMainOther.Show vbModal
'    If Not bVB_TrnOthReceive Then wCNBegin.Show vbModal: Exit Sub
'    Call wFunctionKB.W_KBXHomeSUB

    Call W_FRMxInitailTrans
    '*KT 52-02-25 PH2-ML-SF-106 ในการทำ Other receive ไม่ขึ้นให้ใส่ Customer detail เช่น Age , Sex , Nationality etc…
    '-----------------------------------------------------
    Call SP_ChkxAlwInputALL
    
'    If bVB_CNAlwInputSex Then                  '*Aoo 2012/07/31  ทำเป็น Function
'         wCNDetailCstSex.Show vbModal
'    End If
'
'    If bVB_CNAlwInputNat Then
'        wCNDetailCst.Show vbModal
'    End If
'
'    If bVB_CNAlwInputAge Then
'        wCNAge.Show vbModal
'    End If
    
    If bVB_CNAlwInputSpn Then
        'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
        With wMain.oScanner
             If bVB_OPOSScannerEnable Then
                 If .DeviceEnabled Then
                     .DeviceEnabled = False
                     .Release
                     .Close
                 End If
             End If
         End With
        wCNSaleManID.Show vbModal
        If bVB_OPOSScannerEnable Then
            With oScanner
                     If .DeviceEnabled = False Then
                         .Open tVB_Scanner
                        .Claim 0 'TeamPos2000
                         'Enable the device.
                         .DeviceEnabled = True
                         'Enable the event.
                         .DataEventEnabled = True
                         .AutoDisable = True
                         .ClearInput
                    End If
            End With
        End If
    End If
    '-----------------------------------------------------
    Exit Sub
ErrHandle:
    
End Sub

Public Sub W_PRCxVIPSale()
'--------------------------------------------------------
'Call:
'Cmt: VIP Sale Transaction
'--------------------------------------------------------
On Error GoTo ErrHandle
    If bVB_TrnVIP = False Then wCNBegin.Show vbModal: Exit Sub
   
    Call SP_ChkxAlwInputALL
    
    'Sex/Nation/Age/Saleman                     '*Aoo 2012/07/31  ทำเป็น Function
'    If bVB_CNAlwInputSex Then
'        wCNDetailCstSex.Show vbModal
'    End If
'
'    If bVB_CNAlwInputNat Then
'        wCNDetailCst.Show vbModal
'    End If
'
'    If bVB_CNAlwInputAge Then
'        wCNAge.Show vbModal
'    End If
    
    If bVB_CNAlwInputSpn Then
        'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
        With wMain.oScanner
             If bVB_OPOSScannerEnable Then
                 If .DeviceEnabled Then
                     .DeviceEnabled = False
                     .Release
                     .Close
                 End If
             End If
         End With
        wCNSaleManID.Show vbModal
        If bVB_OPOSScannerEnable Then
            With oScanner
                     If .DeviceEnabled = False Then
                         .Open tVB_Scanner
                        .Claim 0 'TeamPos2000
                         'Enable the device.
                         .DeviceEnabled = True
                         'Enable the event.
                         .DataEventEnabled = True
                         .AutoDisable = True
                         .ClearInput
                    End If
            End With
        End If
    End If
    wCNCstSpc.Show vbModal
    If bVB_TrnCreateVIPSale Then
        Call W_PRCxMergeForVIPSale
        wShwAmtVIP.Show vbModal
        Call W_FRMxInitailTrans(bVB_CNTrStaRegPreviewDay)
        If bVB_CNTrStaRegPreviewDay Then
            olaTransName.Caption = "Preview/" & tVB_TransName
        End If
        '*KT 53-11-04 Preview Day
        wFunctionKB.W_PRCxRetBillVIPSale
'        If Not (GetInputState = 0) Then DoEvents '*KT 53-01-14 DoEvents
    Else
        wCNBegin.Show vbModal
    End If
    Exit Sub
ErrHandle:
    
End Sub

Public Sub W_PRCxTaxExceptSale()
'--------------------------------------------------------
'Call:
'Cmt: Transaction Tax Except Sale
'--------------------------------------------------------
On Error GoTo ErrHandle
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_PRCxPickUp()
'--------------------------------------------------------
'Call:
'Cmt: Transaction PickUp
'--------------------------------------------------------
On Error GoTo ErrHandle
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
    wTenderMain.Show vbModal
    Exit Sub
ErrHandle:
    
End Sub

Public Sub W_PRCxDutyFree()
'--------------------------------------------------------
'Call:
'Cmt: Transaction Duty Free Sale
'--------------------------------------------------------
    'sex/Nation/Age/Saleman
    Call SP_ChkxAlwInputALL
    
'    If bVB_CNAlwInputSex Then                      '*Aoo 2012/07/31  ทำเป็น Function
'        wCNDetailCstSex.Show vbModal
'    End If
'
'    If bVB_CNAlwInputNat Then
'        wCNDetailCst.Show vbModal
'    End If
'
'    If bVB_CNAlwInputAge Then
'        wCNAge.Show vbModal
'    End If

    If bVB_CNAlwInputSpn Then
        'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
        With wMain.oScanner
             If bVB_OPOSScannerEnable Then
                 If .DeviceEnabled Then
                     .DeviceEnabled = False
                     .Release
                     .Close
                 End If
             End If
         End With
        wCNSaleManID.Show vbModal
        If bVB_OPOSScannerEnable Then
            With oScanner
                     If .DeviceEnabled = False Then
                         .Open tVB_Scanner
                        .Claim 0 'TeamPos2000
                         'Enable the device.
                         .DeviceEnabled = True
                         'Enable the event.
                         .DataEventEnabled = True
                         .AutoDisable = True
                         .ClearInput
                    End If
            End With
        End If
    End If
    Call W_FRMxInitailTrans
    bVB_SetFocus = True
End Sub


Public Sub W_PRCxTrainningMode()
'--------------------------------------------------------
'Call:
'Cmt: Trainning Mode ON/OFF
'--------------------------------------------------------
On Error GoTo ErrHandle
    Call W_FRMxInitailTrans
    wTrainMode.Show vbModal
    wCNBegin.Show vbModal
    Exit Sub
ErrHandle:
    
End Sub

Public Sub W_PRCxCreateVIPSale()
'--------------------------------------------------------
'Call:
'Cmt: Create VIP Sale Transaction
'--------------------------------------------------------
On Error GoTo ErrHandle
    If bVB_TrnVIP = False Then wCNBegin.Show vbModal: Exit Sub
    Call W_FRMxInitailTrans
    Call SP_ChkxAlwInputALL
'    If bVB_CNAlwInputSpn Then
'        With wMain.oScanner
'            If bVB_OPOSScannerEnable Then
'                If .DeviceEnabled Then
'                    .DeviceEnabled = False
'                    .Release
'                    .Close
'                End If
'            End If
'        End With
'        wCNSaleManID.Show vbModal
'    End If
    'PH 2.0 ML-SF-019 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
    With wMain.oScanner
        If bVB_OPOSScannerEnable Then
            If .DeviceEnabled Then
                .DeviceEnabled = False
                .Release
                .Close
            End If
        End If
    End With
    wCNCstSpc.Show vbModal
    If bVB_CNAlwInputSpn Then wCNSaleManID.Show vbModal  '*BG 56-11-08 CommSheet TK-ISS3000-417-FAPL-148_BG.xlsx
    If bVB_OPOSScannerEnable Then
        With oScanner
                 If .DeviceEnabled = False Then
                     .Open tVB_Scanner
                    .Claim 0 'TeamPos2000
                     'Enable the device.
                     .DeviceEnabled = True
                     'Enable the event.
                     .DataEventEnabled = True
                     .AutoDisable = True
                     .ClearInput
                End If
        End With
    End If
    If bVB_TrnCreateVIPSale Then
        wShwAmtVIP.Show vbModal
        otbPrompt.SetFocus
    Else
        wCNBegin.Show vbModal
    End If
    Exit Sub
ErrHandle:
    
End Sub

'Private Sub Command1_Click()
'    Dim tSql$
'    Dim orsTemp As ADodb.Recordset
'    tSql = "SELECT * FROM TCNMSku WHERE FTfSkuCode = '1'"
'    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True)
'    If Not (orsTemp.BOF And orsTemp.EOF) Then
'        MsgBox SP_FEDtChkString(orsTemp, "FTSkuAbbEng")
'    End If
'    Set orsTemp = Nothing
'End Sub

'Public Sub Form_Activate()
''    Debug.Print "1"
'    Dim nStep As Long
'On Error GoTo ErrHandle
'    nStep = 0
'    If bW_FirstLoad Then '*KT 51-06-18 Communication Sheet ML-SF-129(5)
'        nStep = 1
'        tVB_CNMnuCurrent = "1700"
'        If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
'            bW_FirstLoad = False
'            Call otbPrompt_KeyDown(vbKeyEscape, 0)
'            Exit Sub
'        End If
''         Debug.Print "1.1"
'         bW_FirstLoad = False
'        Call W_FRMxActivate
'    End If
'    nStep = 2
'    If Not bVB_FormPayAct Then
'        nStep = 3
'        tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
'        tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
'        If bVB_CNReSumAmount Then
'            '*KT 53-01-08 กรณรก่อนหน้ามีกด Total แล้วเคยได้ Deal ต้องการให้มีการแสดงข้อมูลใน Table ใหม่
''            If bVB_CNbHistoryHaveDael Then
''               wFunctionKB.W_AddProduct2GridAgain
''            End If
'            nStep = 4
'            Call W_SumAmount
'            bVB_CNReSumAmount = False
'        End If
'        nStep = 5
'        If bVB_OPOSScannerEnable Then
'            nStep = 6
'            With oScanner
'                     If .DeviceEnabled = False Then
'                        nStep = 7
'                         .Open tVB_Scanner
'                        .Claim 0 'TeamPos2000
'                         'Enable the device.
'                         .DeviceEnabled = True
'                         'Enable the event.
'                         .DataEventEnabled = True
'                         .AutoDisable = True
'                         .ClearInput
'                    End If
'            End With
'        End If
'    End If
'    nStep = 8
''    If Not (GetInputState = 0) Then DoEvents 'ห้ามเอาออกเพื่อให้เคอร์เซอร์ไปที่ otbPrompt '*Ao 53-01-14   DoEvents
'    If bVB_SetFocus Then Call SP_OBJxSetFocus(otbPrompt)
'    bW_FirstLoad = False
''    DoEvents
'    nStep = 9
'    Exit Sub
'ErrHandle:
'    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Activate" & "[" & nStep & "]")
'End Sub

'Private Sub Form_Initialize()
'    'Debug.Print "2"
'On Error GoTo ErrHandle
'    Call W_FRMxActivate
'    Exit Sub
'ErrHandle:
'    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Initialize")
'End Sub

'*KT 54-04-07 ไม่ใช้งานซ้ำซ้อนกับ TextPromt
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
'    If opbWeight.Visible = True Then opbWeight.Visible = False
'    If KeyCode = vbKeyEscape Then Call otbPrompt_KeyDown(vbKeyEscape, 0)
'    bVB_CheckVendor = True                  '*Ao 53-10-13   CheckVendor
'End Sub

Private Sub Form_Load()
    '*KT 51-10-31
    'default grid
    Dim tDisplay$
    Dim tSentComplate As String
    Dim tSta As String
On Error GoTo ErrHandle
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[1]")
    dW_TimeDisplay = Now() '*KT 52-12-19 POSFront 2.0.30 แก้ไขให้กาทำงานของ Timer ของการ Display ถูกต้องมากขึ้น
    'Debug.Print "3"
    bVB_CNStartBill = True '*KT 51-09-17 แก้ไขเรื่องการยิงสินค้าขึ้น Zero
    bVB_CNTenderAct = False
    bVB_CNReSumAmount = False
    bVB_TextPromptAct = False '*KT 51-09-17 แก้ไขเรื่องการยิงสินค้าขึ้น Zero
    
    Call SP_FRMxChildCenter(Me) 'เพื่อให้หน้าจออยู่ตรงกลางพอดี
    
'    oimLogo.AutoSize = True
'    If tVB_CNMainLogo <> "" Then
'        If (SP_FLEbManFile(Trim(tVB_CNMainLogo), nEN_Fle1Exist)) Then
'            oimLogo.LoadImage_FromFile (tVB_CNMainLogo)
'            oimLogo.StretchQuality = True
'        End If
'    End If
    
    If bVB_OPOSScannerEnable Then
        With oScanner
'                 If .DeviceEnabled = False Then
                     .Open tVB_Scanner
                    .Claim 0 'TeamPos2000
                     'Enable the device.
                     .DeviceEnabled = True
                     'Enable the event.
                     .DataEventEnabled = True
                     .AutoDisable = True
                     .ClearInput
                     '*Bump 56-02-28 Touch Scanner เพิ่ม code 3 บรรทัด
                    If .DeviceName = tSNR_TOUCH_SCANNER Then
                        .DecodeData = True
                    End If
'                End If
        End With
    End If
    
    Me.Caption = tVB_CNAppName & " " & "V." & App.Major & "." & App.Minor & "." & App.Revision
    Set oVB_TextClear = otbPrompt
    bW_FirstLoad = True
    
'    uAdaOnOff1.ShowPic = True
'    Set uAdaOnOff1.PicOnLine = oilOnOff.ListImages(1).Picture
'    Set uAdaOnOff1.PicOffLine = oilOnOff.ListImages(2).Picture
'    uAdaOnOff1.OnLineCaption = tVB_CNOnCaption
'    uAdaOnOff1.OffLineCaption = tVB_CNOffCaption
'    uAdaOnOff1.Online = IIf(oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP), True, False) ' oDB.IsCurOnline
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        oimOnOff.Picture = oilOnOff.ListImages(1).Picture
        olaOnOff.Caption = "On-line"
        'move file ej offline to path online
        Call SP_MOVxEJToOnline
        Call SP_CHKbOnlineEnable 'PH 2.0.0
    Else
        oimOnOff.Picture = oilOnOff.ListImages(2).Picture
        olaOnOff.Caption = "Off-line"
    End If

    
    Call W_SetGridHeading
    bVB_SetFocus = True
    With ogdForm
        .rows = .FixedRows
        .rows = nUT_GridMaxRow
    End With
'    Call W_FRMxInitailTrans  '*Eaw 56-05-13 CommSheet TK-ISS3000-187-ST 00 00 01.xls
    
'    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) And oTerminal.SaleMode = nEN_OnlineNormal Then
'         ostForm.Panels(1).Text = "SRV"
'     Else
'         ostForm.Panels(1).Text = "LOCAL"
'    End If

    'ostForm.Panels(3).Text = oDB.OffServerName & "." & oDB.OffDbName  'DB Offline '*
    ostForm.Panels(3).Text = tVB_CNOffDbSrc & "." & oDB.OffDbName  'DB Offline '*Tao 59-12-26
    
    'Set Timer Display on Main
    nVB_TimeDisplayMsg = SP_GETnTimeDisplayMsg
    Call W_FRMxShwMsg
    'Call W_FRMxActivate
'    ostForm.Panels(4).Text = "User : " & tVB_CNUserCode & "/" & tVB_CNUserName '*
    ostForm.Panels(4).Text = "User : " & tVB_CNUserName '*Eaw 56-07-15 CommSheet TK-ISS3000-278 TR FAPL-058.xls
    If bVB_OPOSLineDisplayEnable Then
        If aVB_DSlipMag(0) = "" Or aVB_DSlipMag(1) = "" Then
            tDisplay = "Welcome ....;"
        Else
            tDisplay = aVB_DSlipMag(0) & ";" & Trim(Mid(aVB_DSlipMag(1), 1, 20)) & Space(20 - Len(Trim(Mid(aVB_DSlipMag(1), 1, 20))))
        End If
        Call SP_PrnPollDisplay(tDisplay)
    End If
    
'    Me.WindowState = vbNormal 'เพื่อให้เมื่อ Double Click ที่หัวฟอร์มแล้วฟอร์มไม่ย่อลง '*KT 53-01-14 DoEvents ทำการ set properties ของ Form  WindowState = vbNormal  แทน

    'กรณีที่ตอน Close Terminal เป็น Offline แล้วตอน Open Terminal อีกครั้งเป็น Online ServiceTranfer ไม่ทำงานเพราะ SentSale = '0' อยู่
    tSentComplate = GetSetting("AdaPosMall", "POSFront", "SentSale")
    If tSentComplate = "0" Then
            Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")
    End If
    
    'If oTerminal.SaleMode = nEN_OnlineNormal Then 'PH 2.0.0 OLD
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        'move file ej offline to path online
        '*KT 52-05-15 PH 1.8 ML-SF-060
        'EJ ไม่ส่งมายัง Server ขั้นตอน  1. เปิด POS แบบ Online (Open-> Close-->)  2. Disconnect LAN 3. Open POS in Offline 4. Sale --> Close --> Exit 5. Open POS in Online (Open -->Sale-->Close) 6. Print Closing Report --> Exit        EJ ที่ขายแบบ Offline จะไม่ส่งมายัง Server
        'ป้องกัน 1 เปิดโปรแกรม 2 ตอนที่เปิดโปรแกรมแล้วเข้าหน้าจอ Sale ทันที 3 ตอน Exit Terminal 4 ตอนจบบิลแล้วมีการสลับมาที่ On-line
        Call SP_MOVxEJToOnline
    End If
    '*BG 57-04-30 ---------------------------------------
    Select Case tVB_CNTypePrinter
        Case "0"
            olaIndiCator.Caption = "S"
        Case "1"
            olaIndiCator.Caption = "T"
        Case "2"
            olaIndiCator.Caption = "A"
    End Select
    '------------------------------------------------------------
    
    Call SaveSetting("AdaPosMall", "POSFront", "RepStop", "0") '*KT 52-07-27 เริ่มให้ ServiceAutoReplicate.EXE ทำงาน
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[2]")
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:Form_Load[3]")
End Sub


'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = True
'End Sub
    
'Private Sub Form_Resize()
'    wMainRet.WindowState = vbMaximized
'End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Unload[1]")
    otmDisplay.Enabled = False
    otmLock.Enabled = False
   If Not orsSKU Is Nothing Then Set orsSKU = Nothing
    With oScanner
        If bVB_OPOSScannerEnable Then
            If .DeviceEnabled Then
                If bVB_ChkRetGST7 = True Then bVB_ChkRetGST7 = False '*Amorn 65-07-11 ปิดการทำ Return 7%
                .DeviceEnabled = False
                .Release
                .Close
            End If
        End If
    End With
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Unload[2]")
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:Form_Unload[3]")
End Sub

Private Sub ogdForm_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim tSql$, tWhr$
    Dim cSubTotal  As Double
    Dim nMRow As Long
    Dim i As Long
    Dim orsTemp As ADODB.Recordset
    Dim cQtyDiff As Double
    Dim cMemPriceMaxQty As Double
    Dim bCHKPromo As Boolean
    Dim cQty As Double
    Dim cPrice As Double
    Dim tBug$
    Dim cMaxQty  As Double
    Dim tPrint As String  '*Eaw 56-07-01CommSheet TK-ISS3000-252-TR.xls
    
On Error GoTo ErrHandle
    '*KT 51-09-2008 แก้ไขยิงสินค้าอฃแล้วขึ้น Zero
    If Row = 0 Or nW_GridEnter > 1 Then
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
    End If
    nW_GridEnter = nW_GridEnter + 1 'ป้องกันการทำงานซ้ำเนื่องจาก ogdForm_AfterEdit ถูกสั่งให้ทำงานที่ ogdForm_KeyPressEdit ซึ่งในการ Run EXE จะเกิด Enter 2 ครั้งทำให้
    'Program ทำงานซ้ำซ้อน
    '------------------------------------------------------------------------------
    With ogdForm
        Select Case UCase(.ColKey(Col))
            Case UCase("FCSdtSaleUnit")
                '*KT 52-11-26 PH 3.0-ML-SF-034  กด Price Override เป็นเศษสตางค์  เครื่องมีข้อความเตื่อน แต่เมื่อกด Enter แล้วใส่จำนวนเงินใหม่ เมื่อพิมพ์ Receipt จะพิมพ์ค่าที่ป้อนผิด
                If tW_KeyPress = "" Then
                    Call SP_DATxRsNothing(orsTemp)
                    nW_GridEnter = 0
                    bW_GridEnter = False
                    Exit Sub
                Else
                    If Not bW_ChkDecimals Then
'*Eaw 56-07-04 CommSheet TK-ISS3000-201-ST.xls ไม่ต้องเช็คเหรียญ
'                        If Not (W_CHKbBanknoteAlw(SP_VALcText2Double(SP_DEFtChangeNum(tW_KeyPress)))) Then
'                            tW_KeyPress = ""
'                            bVB_SetFocus = False
'                            nW_GridEnter = 0
'                            bW_GridEnter = False
'                            bW_ChkDecimals = False
'                            Call SP_MSGnShowing("tms_CN008028", nCS_Warning)   'คุณป้อนเศษสตางค์ไม่ถูกต้อง
'                            .Enabled = True
'                            .HighLight = flexHighlightAlways
'                            .TextMatrix(.Row, .ColIndex("FCSdtSaleUnit")) = SP_FMTcCurPoint(cW_SaleItem, True, True)
'                            .EditText = cW_SaleItem
'                            .Select Row, 1        'move highlight to first row
'                            .Col = .ColIndex("FCSdtSaleUnit")
'                            .EditCell
'                             tW_KeyPress = "" 'Replace(.TextMatrix(.Row, .Col), ".", "")
'                             bVB_PriceOveride = False
'                             Call SP_DATxRsNothing(orsTemp)
'                             Exit Sub
'                        Else
'                             bW_ChkDecimals = True
'                        End If
                    End If
                End If

                If .EditText = "" Then
                    cPrice = cW_SaleItem
                Else
                    cPrice = CDbl(Val(Format(.EditText)))
                End If
                
                If .TextMatrix(Row, Col) = "" Then .TextMatrix(Row, Col) = 0#
                If SP_VALcText2Double(.ValueMatrix(Row, .ColIndex("FNSdtQty"))) * cPrice > cCS_CNMaxSale Then
                        Call SP_MSGnShowing("tms_CN005170", nCS_Warning)
                        .TextMatrix(Row, .ColIndex("FCSdtSaleUnit")) = cW_SaleItem
                        cPrice = cW_SaleItem
                End If
                .TextMatrix(Row, .ColIndex("FCSdtSalePrice")) = SP_FMTcCurPoint(cPrice, False, True) '.TextMatrix(Row, .ColIndex("FCSdtSaleUnit"))
                Select Case .TextMatrix(Row, .ColIndex("FTSdtWeight"))
                    Case ""  'อื่นๆ ราคารวม = จำนวน * ราคาต่อหน่วย)
                        .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(.ValueMatrix(Row, .ColIndex("FCSdtQty")) * cPrice, False, True)
                    Case Else  'เครื่องชั่ง, น้ำหนัก สามารทำ Price Override
                        .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(.ValueMatrix(Row, .ColIndex("FCSdtQty")) * cPrice, False, True)
                End Select
                
                nMRow = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FNSdtSeqNo"), .rows - 1, .ColIndex("FNSdtSeqNo"))
                For i = 1 To nMRow
                    If .TextMatrix(i, .ColIndex("FTSdtCancelFlag")) <> "1" Then
                        cSubTotal = cSubTotal + CDbl(Val(Format(.TextMatrix(Row, .ColIndex("FCSdtSaleAmt")))))
                    End If
                Next i
                
                If cSubTotal > cCS_CNMaxSale Then
                        Call SP_MSGnShowing("tms_CN005170", nCS_Warning)
                        .TextMatrix(Row, .ColIndex("FCSdtSaleUnit")) = cW_SaleItem
                        .TextMatrix(Row, .ColIndex("FCSdtSalePrice")) = .TextMatrix(Row, .ColIndex("FCSdtSaleUnit"))
                         cPrice = cW_SaleItem
                        Select Case .TextMatrix(Row, .ColIndex("FTSdtWeight"))
                            Case ""  'อื่นๆ ราคารวม = จำนวน * ราคาต่อหน่วย)
                                .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(.ValueMatrix(Row, .ColIndex("FCSdtQty")) * cW_SaleItem, False, True)
                            Case Else  'เครื่องชั่ง, น้ำหนัก สามารทำ Price Override
                               .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(.ValueMatrix(Row, .ColIndex("FCSdtQty")) * cW_SaleItem, False, True)
                        End Select
                End If
                
                '*KT 52-05-19 PH 1.8 ML-SF-068 สินค้า Promotion เมื่อทำการคูณ (เปลี่ยนจำนวนสินค้า) จะช้ามาก
                'เดิมเจอคำสั่ง Stop
'                olaPrice.Caption = .TextMatrix(Row, .ColIndex("FCSdtSaleAmt"))
                olaPrice.Caption = SP_FMTcCurPoint(.TextMatrix(Row, .ColIndex("FCSdtSaleAmt")), False, True) 'แสดงราคารวมที่หน้าจอ
'                If Not (GetInputState = 0) Then DoEvents '*KT 53-01-14 DoEvents
                If UCase(olaPrice.Caption) = UCase("subtotal") Then
                    Call SP_MSGnShowing(tMS_CN534, tCS_CNMsgWarn)
                End If
                Call W_UpdateTSaleDT(Row, Col)
                
                'ReCal
                '-----------------------------------------------------------------              '*Ao 2010/12/16         Price Override
'                tSql = "SELECT FCScdAmt "
'                tSql = tSql & vbCrLf & " FROM " & tVB_CNTblCD
'                tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
'                tSql = tSql & vbCrLf & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
'                tSql = tSql & vbCrLf & " AND  FNSdtSeqNo=" & .TextMatrix(Row, .ColIndex("FNSdtSeqNo"))  'seq เดิมที่ถูกยกเลิก
''                tSql = tSql & vbCrLf & " AND NOT(FNDctNo IN ('1','2','3'))" '*KT 52-06-12 PH 2.0 ML-SF-027 Price Override ไม่สามารถใช้กับสินค้า Promotion ได้ (ให้สามารถทำ Price Overide กรณีที่ได้ส่วนลด Deal , PromotionTime,Promotion)
'                Call SP_SQLvExecute(tSql, orsTemp)
'                If Not (orsTemp.BOF Or orsTemp.EOF) Then
'                    oDisc.nENDiscType = nEN_Disc01Deal 'เริ่ม Recal ส่วนลดจาก Deal
'                    oDisc.dTranDate = tUT_SaleDate
'                    oDisc.tTerminalNo = tVB_CNTerminalNum
'                    oDisc.tTransNo = tVB_Running
'                    oDisc.nDTSeqNo = Val(Format(.TextMatrix(Row, .ColIndex("FNSdtSeqNo"))))
'                    oDisc.DISbClearDiscDTByItem 'oDisc.DISbClearDiscDT       '*KT 51-09-19 Clear DT ก่อนทำ '*KT 52-05-20 ทำการ Recal ส่วนลดเป็นแบบรายตัว
'                   Call oDisc.DISxPrcRecal(, , , Val(Format(.TextMatrix(Row, .ColIndex("FNSdtSeqNo"))))) '*KT 52-05-20 ทำการ Recal ส่วนลดเป็นแบบรายตัว
'                End If
                '------------------------------------------------------------
                .HighLight = flexHighlightNever
                .Editable = flexEDNone
                .Enabled = False
                 '*KT 52-05-21 Comment Code ไว้เพราะมีการทำการ Update ซ้ำซ้อน
'                tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
'                tWhr = tWhr & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
'                If .TextMatrix(Row, .ColIndex("FTSdtWeight")) = "" Then  'สินค้าปกติ
'                     tWhr = tWhr & " AND  FTSdtBarCode=''" & .TextMatrix(Row, .ColIndex("FTSdtBarCode")) & "'' AND FNSdtSeqNo=" & .TextMatrix(Row, .ColIndex("FNSdtSeqNo"))
'                Else  'สินค้าน้ำหนัก
'                     tWhr = tWhr & " AND  FTSdtInstoreBar=''" & .TextMatrix(Row, .ColIndex("FTSdtBarCode")) & "'' AND FNSdtSeqNo=" & .TextMatrix(Row, .ColIndex("FNSdtSeqNo"))
'                End If
'                tSql = "STP_CNxSQLFunction @ptFunc = '*' ,@ptTable = '" & tVB_CNTblDT & "',@ptWhere = '" & tWhr & "'"
'                Call SP_SQLvExecute(tSql, orsSKU)
'                Call W_UPDxPriceOverride(Row, Col, orsSKU) 'มีการบันทึกข้อมูลการแก้ไขราคา
                .TextMatrix(Row, .ColIndex("FCSdtDis")) = 0            '*Ao 2011/03/29  CommSheet ML-PRV-047-SF-Deposit Tran Subtotal incorrect when Price Override.xls
                .TextMatrix(Row, .ColIndex("FTSdtPriceOverideFlag")) = "Y"
                .ColDataType(Col) = flexDTDouble
                .ColAlignment(Col) = flexAlignRightCenter
                .ColFormat(Col) = SP_FMTtNum2Set
                 Call W_SumAmount
                 ostForm.Panels(2).Text = "Ready..." '*
                 bVB_PriceOveride = False
                 otbPrompt.Locked = False
                 bVB_SetFocus = True
                 Call W_ClearValAfterInsTSDT
                 Call SP_OBJxSetFocus(otbPrompt)
                 Call SP_DATxRsClose(orsTemp)
            Case UCase("FNSdtQty")
                tBug = "1:" & Time
                'If .TextMatrix(Row, Col) = "" Or CDbl(Val(Format(.TextMatrix(Row, Col)))) = 0 Then .TextMatrix(Row, Col) = 1
                If .TextMatrix(Row, Col) = "" Or CDbl(Val(Format(.EditText))) = 0 Then
                    .TextMatrix(Row, Col) = 1
                    .EditText = 1
                End If
                cQty = CDbl(Val(Format(.EditText)))
                
                '*Ao 2011/05/26 CommSheet ST-FL-027 Check input from scanner when change QTY.xls
                '*Ao 2011/05/18 Kampon Preecheun Email 18 พฤษภาคม 2011 12:13
'                If .TextMatrix(Row, .ColIndex("FNSdtQty")) > cCS_CNMaxQty Then
'                    If SP_MSGnShowing("tms_CN008036", tCS_CNMsgYesNoDefNo, , , cQty) = vbNo Then
'                        .TextMatrix(Row, .ColIndex("FNSdtQty")) = nW_QtyItem
'                        cQty = nW_QtyItem
'                    Else
'                        'cCS_CNMaxQty  ไม่สามารถใช้  Len(cCS_CNMaxQty)
'                        cMaxQty = cCS_CNMaxQty
'                        .TextMatrix(Row, .ColIndex("FNSdtQty")) = Left(.TextMatrix(Row, .ColIndex("FNSdtQty")), Len(cMaxQty))
'                        cQty = .TextMatrix(Row, .ColIndex("FNSdtQty"))
'                    End If
'                End If
                '*BG 57-04-18
                SleepEx 1000, 0
                 If cQty <= cCS_CNMaxQty Then
                 
                 Else
'                If (CDbl(Val(.TextMatrix(Row, .ColIndex("FNSdtQty")))) > CDbl(Val(cCS_CNMaxQty))) And (Len(Trim(.TextMatrix(Row, .ColIndex("FNSdtQty")))) > Len(Trim(cCS_CNMaxQty))) Then
                    Call SP_MSGxManualShow(tMS_CN521, vbExclamation)
                    .TextMatrix(Row, .ColIndex("FNSdtQty")) = nW_QtyItem
                    Exit Sub
              
                End If
                '************************
                
                
                If CDbl(Val(Format(cQty * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit"))))) > cCS_CNMaxSale Then
'                        Call SP_MSGnShowing("tms_CN005170", nCS_Warning)
                        '*AO 2011/05/26  CommSheet ST-FL-027 Check input from scanner when change QTY.xls
                        Call SP_MSGnShowingFrm("tms_CN005170", nCS_Warning)
                        .TextMatrix(Row, .ColIndex("FNSdtQty")) = nW_QtyItem
                        cQty = nW_QtyItem
                End If
                
                
                .TextMatrix(Row, .ColIndex("FCSdtQty")) = SP_FMTcCurPoint(cQty, False, True)
                Select Case .TextMatrix(Row, .ColIndex("FTSdtWeight"))
                    Case ""  'อื่นๆ ราคารวม = จำนวน * ราคาต่อหน่วย)
                        .TextMatrix(Row, .ColIndex("FCSdtQty")) = cQty
                        .TextMatrix(Row, .ColIndex("FCSdtQtyAll")) = cQty * cUT_StkFac * cUT_SKUUnitF
                        .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(cQty * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit")), False, True)
                    Case Else  'เครื่องชั่ง, น้ำหนัก สามารถทำ Qty Change
                        .TextMatrix(Row, .ColIndex("FCSdtQty")) = cQty
                        .TextMatrix(Row, .ColIndex("FCSdtQtyAll")) = cQty * cUT_StkFac * cUT_SKUUnitF
                        .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(cQty * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit")), False, True)
                End Select
                
                nMRow = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FNSdtSeqNo"), .rows - 1, .ColIndex("FNSdtSeqNo"))
                For i = 1 To nMRow
                    If .TextMatrix(i, .ColIndex("FTSdtCancelFlag")) <> "1" Then
                        cSubTotal = cSubTotal + SP_VALcText2Double(.TextMatrix(Row, .ColIndex("FCSdtSaleAmt")))
                    End If
                Next i
                If cSubTotal > cCS_CNMaxSale Then
                        Call SP_MSGnShowing("tms_CN005170", nCS_Warning)
                        .TextMatrix(Row, .ColIndex("FNSdtQty")) = nW_QtyItem
                        Select Case .TextMatrix(Row, .ColIndex("FTSdtWeight"))
                            Case ""  'อื่นๆ ราคารวม = จำนวน * ราคาต่อหน่วย)
                                .TextMatrix(Row, .ColIndex("FCSdtQty")) = .TextMatrix(Row, .ColIndex("FNSdtQty"))
                                .TextMatrix(Row, .ColIndex("FCSdtQtyAll")) = nW_QtyItem * cUT_StkFac * cUT_SKUUnitF
                                .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(nW_QtyItem * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit")), False, True)
                            Case Else  'เครื่องชั่ง, น้ำหนัก สามารถทำ Qty Change
                                .TextMatrix(Row, .ColIndex("FCSdtQty")) = .TextMatrix(Row, .ColIndex("FNSdtQty"))
                                .TextMatrix(Row, .ColIndex("FCSdtQtyAll")) = nW_QtyItem * cUT_StkFac * cUT_SKUUnitF
                                .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(nW_QtyItem * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit")), False, True)
                        End Select
                End If
                '*KT 53-10-30 Preview Day
                '---------------------------------------------------------------------------------------------------------------------------------
                If bVB_CNTrStaRegPreviewDay Then
                    If .TextMatrix(Row, .ColIndex("FTSdtStaPreview")) = "Y" And .TextMatrix(Row, .ColIndex("FTSdtStaNotAlwPrv")) = "N" Then
                        Call SP_INIxSwitchModePreviewDay(True)
                    Else
                        Call SP_INIxSwitchModePreviewDay(False)
                    End If
                End If
                '---------------------------------------------------------------------------------------------------------------------------------
                'ตรวจสอบ Limit จำนวนสินค้าที่ซื้อจาก Max Qty ของ Member Price
                'กรณีสินค้านั้นเป็นสินค้า Member Price แล้วเกินจำนวน Max Qty ของทั้งบิลจะไม่ยอมให้ขาย
                cQtyDiff = CDbl(Val(Format(.EditText)) - nW_QtyItem)
                'cQtyDiff = CDbl(CLng(Val(Format(.EditText))) - nW_QtyItem) '*KT 52-04-17 ให้เปลียนจาก .ValueMatrix มาใช้  .EditText กรรีต้องการค่าใน Cell เมื่อมีการการทำงานใน AfterEdit

                bCHKPromo = False
                If bVB_CNMemberPrice Then
                    If xVB_CNMemberP.ID <> "" And xVB_CNMemberP.MallCardMemberPrice = "Y" And .TextMatrix(Row, .ColIndex("FTSdtStaRfd")) = "Y" Then
                        tSql = "SELECT FNPmpMemMaxQty"
                        tSql = tSql & vbCrLf & "FROM " & tVB_CNTblMBPrice
                        tSql = tSql & vbCrLf & "WHERE FTSkuCode = '" & .TextMatrix(Row, .ColIndex("FTSKUCode")) & "'"
                        tSql = tSql & vbCrLf & "AND FDPmpMemEndDate >= '" & SP_DTEtStrSTD(SP_GETdChkDateDisc) & "'" '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
                        tSql = tSql & vbCrLf & "AND ISNULL(FTPmpStaDiscType,'0') <> '0'"
                        If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                            If Not (orsTemp.BOF And orsTemp.EOF) Then
                                cMemPriceMaxQty = SP_FEDvChkNumeric(orsTemp, "FNPmpMemMaxQty")
                            Else
                                 cMemPriceMaxQty = 0
                            End If
                       Else
                            cMemPriceMaxQty = 0
                       End If
                       If cMemPriceMaxQty > 0 Then
                            If Not wFunctionKB.W_PRCbMaxQtyMemberPriceBySKU(.TextMatrix(Row, .ColIndex("FTSKUCode")), cQtyDiff, cMemPriceMaxQty, 0) Then
            '                    Call SP_MSGnShowing("tms_CN005071", nCS_Error)
                                 MsgBox "Max Qty Member Price"
                                .TextMatrix(Row, .ColIndex("FNSdtQty")) = nW_QtyItem
                                Select Case .TextMatrix(Row, .ColIndex("FTSdtWeight"))
                                    Case ""  'อื่นๆ ราคารวม = จำนวน * ราคาต่อหน่วย)
                                        .TextMatrix(Row, .ColIndex("FCSdtQty")) = .TextMatrix(Row, .ColIndex("FNSdtQty"))
                                        .TextMatrix(Row, .ColIndex("FCSdtQtyAll")) = nW_QtyItem * cUT_StkFac * cUT_SKUUnitF
                                        .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(nW_QtyItem * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit")), False, True)
                                    Case Else  'เครื่องชั่ง, น้ำหนัก  สามารถทำ Qty change
                                        .TextMatrix(Row, .ColIndex("FCSdtQty")) = .TextMatrix(Row, .ColIndex("FNSdtQty"))
                                        .TextMatrix(Row, .ColIndex("FCSdtQtyAll")) = nW_QtyItem * cUT_StkFac * cUT_SKUUnitF
                                        .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(nW_QtyItem * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit")), False, True)
                                End Select
                            End If
                        Else
                            bCHKPromo = True
                        End If
                    Else
                        bCHKPromo = True
                    End If
                Else
                    bCHKPromo = True
                End If
                'ตรวจสอบ Limit จำนวนสินค้าที่ซื้อจาก Max Qty ของ Promotion
                'กรณีสินค้านั้นเป็นสินค้า Promotion แล้วเกินจำนวน Max Qty ของทั้งบิลจะไม่ยอมให้ขาย
                If bCHKPromo Then
                    If Not wFunctionKB.W_PRCbMaxQtyPmtTimeSaleBySKU(.TextMatrix(Row, .ColIndex("FTSKUCode")), cQtyDiff) Then
                        Call SP_MSGnShowing("tms_CN005071", nCS_Error)
                        .TextMatrix(Row, .ColIndex("FNSdtQty")) = nW_QtyItem
                        Select Case .TextMatrix(Row, .ColIndex("FTSdtWeight"))
                            Case ""  'อื่นๆ ราคารวม = จำนวน * ราคาต่อหน่วย)
                                .TextMatrix(Row, .ColIndex("FCSdtQty")) = .TextMatrix(Row, .ColIndex("FNSdtQty"))
                                .TextMatrix(Row, .ColIndex("FCSdtQtyAll")) = nW_QtyItem * cUT_StkFac * cUT_SKUUnitF
                                .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(nW_QtyItem * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit")), False, True)
                            Case Else  'เครื่องชั่ง, น้ำหนัก สามารถทำ Qty change
                                .TextMatrix(Row, .ColIndex("FCSdtQty")) = .TextMatrix(Row, .ColIndex("FNSdtQty"))
                                .TextMatrix(Row, .ColIndex("FCSdtQtyAll")) = nW_QtyItem * cUT_StkFac * cUT_SKUUnitF
                                .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(nW_QtyItem * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit")), False, True)
                        End Select
                    End If
                End If
                'ตรวจสอบ Limit จำนวนสินค้าที่ซื้อจาก Max Qty ของ Promotion
                'กรณีสินค้านั้นเป็นสินค้า Promotion แล้วเกินจำนวน Max Qty ของทั้งบิลจะไม่ยอมให้ขาย
                If bCHKPromo Then
                    If Not wFunctionKB.W_PRCbMaxQtyPromotionBySKU(.TextMatrix(Row, .ColIndex("FTSKUCode")), cQtyDiff) Then
                            Call SP_MSGnShowing("tms_CN005070", nCS_Error)
                            .TextMatrix(Row, .ColIndex("FNSdtQty")) = nW_QtyItem
                            Select Case .TextMatrix(Row, .ColIndex("FTSdtWeight"))
                                Case ""  'อื่นๆ ราคารวม = จำนวน * ราคาต่อหน่วย)
                                    .TextMatrix(Row, .ColIndex("FCSdtQty")) = .TextMatrix(Row, .ColIndex("FNSdtQty"))
                                    .TextMatrix(Row, .ColIndex("FCSdtQtyAll")) = nW_QtyItem * cUT_StkFac * cUT_SKUUnitF
                                    .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(nW_QtyItem * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit")), False, True)
                                Case Else  'เครื่องชั่ง, น้ำหนัก สามารถทำ Qty Change
                                    .TextMatrix(Row, .ColIndex("FCSdtQty")) = .TextMatrix(Row, .ColIndex("FNSdtQty"))
                                    .TextMatrix(Row, .ColIndex("FCSdtQtyAll")) = nW_QtyItem * cUT_StkFac * cUT_SKUUnitF
                                    .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(nW_QtyItem * .ValueMatrix(Row, .ColIndex("FCSdtSaleUnit")), False, True)
                            End Select
                    End If
                End If
                '*KT 52-05-19 PH 1.8 ML-SF-068 สินค้า Promotion เมื่อทำการคูณ (เปลี่ยนจำนวนสินค้า) จะช้ามาก
                'เดิมเจอคำสั่ง Stop
'                olaPrice.Caption = .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) 'แสดงราคารวมที่หน้าจอ           '*Ao 2014/02/26  CommSheet TK-ISS3000-462-Tender Cash 0.00_Aoo-570228
'                If Not (GetInputState = 0) Then DoEvents '*KT 53-01-14 DoEvents
                If UCase(olaPrice.Caption) = UCase("subtotal") Then
                    Call SP_MSGnShowing(tMS_CN534, tCS_CNMsgWarn)
                End If
                 Call W_UpdateTSaleDT(Row, Col)
'                 tBug = tBug & vbCrLf & "2:" & Time                '*Ao 2014/05/28
                 olaPrice.Caption = .TextMatrix(Row, .ColIndex("FCSdtSaleAmt")) 'แสดงราคารวมที่หน้าจอ
                'ReCal
                '-----------------------------------------------------------------
                tSql = "SELECT FCScdAmt "
                tSql = tSql & " FROM " & tVB_CNTblCD
                tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                tSql = tSql & " AND  FNSdtSeqNo=" & .TextMatrix(Row, .ColIndex("FNSdtSeqNo"))  'seq เดิมที่ถูกยกเลิก
                tSql = tSql & " AND  (FTSdtStkCode <> '1')"                         '*Ao 2010/12/16 Price Override
                
                Call SP_SQLvExecute(tSql, orsTemp)
                If Not (orsTemp.BOF Or orsTemp.EOF) Then
                    oDisc.nENDiscType = nEN_Disc01Deal 'เริ่ม Recal ส่วนลดจาก Deal
                    oDisc.dTranDate = tUT_SaleDate
                    oDisc.tTerminalNo = tVB_CNTerminalNum
                    oDisc.tTransNo = tVB_Running
                    oDisc.nDTSeqNo = Val(Format(.TextMatrix(Row, .ColIndex("FNSdtSeqNo"))))
                    oDisc.DISbClearDiscDTByItem 'oDisc.DISbClearDiscDT       '*KT 51-09-19 Clear DT ก่อนทำ '*KT 52-05-20 ทำการ Recal ส่วนลดเป็นแบบรายตัว
                   Call oDisc.DISxPrcRecal(, , , Val(Format(.TextMatrix(Row, .ColIndex("FNSdtSeqNo"))))) '*KT 52-05-20 ทำการ Recal ส่วนลดเป็นแบบรายตัว
                    'Debug.Print "DISxPrcRecal"
                End If
                '------------------------------------------------------------
                .HighLight = flexHighlightNever
                .Editable = flexEDNone
                .Enabled = False
                
                 tBug = tBug & vbCrLf & "3:" & Time
                'Insert Table LogDT
                tSql = "SELECT *  "
                tSql = tSql & " FROM " & tVB_CNTblDT
                tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                If .TextMatrix(Row, .ColIndex("FTSdtWeight")) = "" Then  'สินค้าปกติ
                    tSql = tSql & " AND  FTSdtBarCode='" & .TextMatrix(Row, .ColIndex("FTSdtBarCode")) & "' AND FNSdtSeqNo=" & .TextMatrix(Row, .ColIndex("FNSdtSeqNo"))  'seq เดิมที่ถูกยกเลิก
                Else  'สินค้าน้ำหนัก
                     tSql = tSql & " AND  FTSdtInstoreBar='" & .TextMatrix(Row, .ColIndex("FTSdtBarCode")) & "' AND FNSdtSeqNo=" & .TextMatrix(Row, .ColIndex("FNSdtSeqNo"))  'seq เดิมที่ถูกยกเลิก
                End If
                Call SP_SQLvExecute(tSql, orsTemp)
                Call wFunctionKB.W_PRCxLogDT(orsTemp, eQtyChg, SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll"))
                Call SP_DATxRsNothing(orsTemp)
                Call W_SumAmount
                 otbPrompt.Locked = False
                 bVB_SetFocus = True
                 Call W_ClearValAfterInsTSDT
                 Call SP_OBJxSetFocus(otbPrompt)
                  tBug = tBug & vbCrLf & "4:" & Time
                  'MsgBox tBug
        End Select
        '*Eaw 56-07-01CommSheet TK-ISS3000-252-TR.xls
        'Display
        If bVB_OPOSLineDisplayEnable Then
            tPrint = wMain.ogdForm.TextMatrix(Row, wMain.ogdForm.ColIndex("FTSkuAbbName"))
            tPrint = tPrint & ";" & wMain.ogdForm.TextMatrix(Row, wMain.ogdForm.ColIndex("FCSdtSaleAmt"))
            Call SP_PrnPollDisplay(tPrint)
        End If
    End With
    Exit Sub
ErrHandle:
    Call W_ClearValAfterInsTSDT               '*Ao 2011/03/31  CommSheet ML-PRV-054-SF Wrong qty on last item.xls
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Step:Qty Change/Price Override" & " Events:ogdForm_AfterEdit")
End Sub

'*KT 51-05-26 Comsheet ML-SF-274
Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With ogdForm
         If .TextMatrix(NewRow, .ColIndex("FNSdtSeqNo")) <> "" Then
            .TopRow = NewRow
            nVB_CNScollItem = NewRow
        End If
    End With
End Sub

Private Sub ogdForm_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With ogdForm
        nW_QtyItem = SP_VALcText2Double(.TextMatrix(NewRow, .ColIndex("FNSdtQty")))
        cW_SaleItem = CDbl(Val(Format(.TextMatrix(NewRow, .ColIndex("FCSdtSaleUnit")))))
    End With
End Sub

Private Sub ogdForm_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    'Dim nLastRow As Long
    Dim tPrint As String
    With ogdForm
        'nLastRow = .Aggregate(flexSTCount, .FixedRows, .ColIndex("FNSdtSeqNo"), .FixedRows, .ColIndex("FNSdtSeqNo"))
        If .TextMatrix(NewRowSel, .ColIndex("FNSdtSeqNo")) = "" Then
            Cancel = True
        Else
            '*Eaw 56-09-06 CommSheet TK-ISS3000-252
            'Display
            If bVB_OPOSLineDisplayEnable Then
                tPrint = wMain.ogdForm.TextMatrix(NewRowSel, wMain.ogdForm.ColIndex("FTSkuAbbName"))
                tPrint = tPrint & ";" & SP_FMTcCurPoint(.TextMatrix(NewRowSel, wMain.ogdForm.ColIndex("FCSdtSaleAmt")), False, True)
                Call SP_PrnPollDisplay(tPrint)
            End If
            olaPrice.Caption = SP_FMTcCurPoint(.TextMatrix(NewRowSel, wMain.ogdForm.ColIndex("FCSdtSaleAmt")), False, True)
            olaBarcode.Caption = .TextMatrix(NewRowSel, wMain.ogdForm.ColIndex("FTSdtBarCode"))
            olaSKUName.Caption = .TextMatrix(NewRowSel, wMain.ogdForm.ColIndex("FTSkuAbbName"))
        End If
        
    End With
End Sub

Private Sub ogdForm_GotFocus()
    With ogdForm
        Select Case UCase(.ColKey(.Col))
            Case UCase("FCSdtSaleUnit")
                tW_KeyPress = Replace(.TextMatrix(.Row, .Col), ".", "")
            Case Else
                tW_KeyPress = ""
        End Select
    End With
End Sub

Public Sub ogdForm_KeyDown(KeyCode As Integer, Shift As Integer)
'-----------------------------------------------------------------------------------------------------------------------
'cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
'กรุณาตรวจสอบความถูกต้องของ Communication Sheet ML-SF-381,ML-SF-382 ด้วยนะครับ
'-----------------------------------------------------------------------------------------------------------------------
    Dim tKeyCode$, tShift$
    Dim tProc$
    Dim orsFunc As ADODB.Recordset
    Dim tSql As String
On Error GoTo ErrHandle
    With ogdForm
            bVB_CheckVendor = True                  '*Ao 53-10-13   CheckVendor
            If Shift = 4 And KeyCode <> 18 Then KeyCode = 0      'do not accept Alt
            Select Case KeyCode
                '*KT 51-07-23 Communication Sheet ML-SF-412 Can't Deal When scroll Item.xls
                '------------------------------------------------------------------------------------------------------------------------------------------
                Case vbKeyHome
                    .HighLight = flexHighlightNever
                    .Editable = flexEDNone
                    .Enabled = False
                     Call W_SumAmount
                     otbPrompt.Locked = False
                     bVB_SetFocus = True
                    Call otbPrompt_KeyDown(vbKeyHome, 0)
                '------------------------------------------------------------------------------------------------------------------------------------------
                
                Case vbKeyDown
                    If .Row >= .rows - 1 Then KeyCode = 0: Call SP_DATxRsNothing(orsFunc): Exit Sub
                    If .TextMatrix(.Row + 1, .ColIndex("FNSdtSeqNo")) = "" Then KeyCode = 0: Call SP_DATxRsNothing(orsFunc): Exit Sub
                Case vbKeyEscape
                    .HighLight = flexHighlightNever
                    .Editable = flexEDNone
                    .Enabled = False
                     Call W_SumAmount
                     otbPrompt.Locked = False
                     bVB_SetFocus = True
                     Call SP_OBJxSetFocus(otbPrompt)
                 Case 18, 27, 33, 34, 35, 36, 45, 112 To 124, 128, 129
                    Call SP_FRMxSetTimeCount(False)       'หยุดเวลา Lock
                    tSql = "SELECT FTSkbRef FROM TSysKeyboard   WHERE FNSkbKey='" & CStr(KeyCode) & "'"
                    tSql = tSql & " AND FNSkbShift = '" & CStr(Shift) & "'"
                    tSql = tSql & " AND FTSkbwMain =  '1'"
                    tSql = tSql & " AND FTSkbUsage ='Y'" '*KT 54-04-08 V.3.0.56
                    Call SP_SQLvExecute(tSql, orsFunc, oDB.ocnOnLine, , , , , True, True) 'read master
                    If orsFunc.EOF Then
                        Call SP_MSGnShowing("tms_CN005107", nCS_Information)    'tCS_MSG22 ฟังก์ชันนี้ยังไม่ได้กำหนด หรือไม่ได้รับอนุญาตให้ใช้งาน;This function is undefined or not allowed to use. '*KT 54-04-08 V.3.0.56 เปลี่ยน message
                        KeyCode = 0
                        Call SP_DATxRsNothing(orsFunc)
                        If Not bUT_FuncLockKB Then Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock '*KT 52-12-15 กรณีกด Keyboard ไม่มี function
                        Exit Sub
                    End If
'                    nUT_UserLevelFunc = SP_VALcText2Double(tUT_UsrLevel)
                    tProc = SP_FEDtChkString(orsFunc, "FTSkbRef")
                    
                    If (tProc = "W_KBRcvVenCpnCancel") Or (tProc = "W_KBRcvVenCpnChange") Or (tProc = "") Then                      '*Ao 2011/02/24  CommSheet ML-PRV-041-Runtime Error '*KT 54-04-08 V.3.0.56 ป้องกัน Error จาก CallByName
                        Call SP_MSGnShowing("tms_CN005107", nCS_Information)    'tCS_MSG22 ฟังก์ชันนี้ยังไม่ได้กำหนด หรือไม่ได้รับอนุญาตให้ใช้งาน;This function is undefined or not allowed to use. '*KT 54-04-08 V.3.0.56 เปลี่ยน message
                        KeyCode = 0
                        Call SP_DATxRsNothing(orsFunc)
                        If Not bUT_FuncLockKB Then Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock '*KT 52-12-15 กรณีกด Keyboard ไม่มี function
                        Exit Sub
                    End If
                    Call SP_DATxRsNothing(orsFunc)
                    
                    otmDisplay.Enabled = False  'ไม่ต้องการให้ทำงานหลังจากมีการมี key text promt '*KT 54-04-08 V.3.0.56
                    
                    '*KT 54-04-09 V.3.0.56 close scanner
                     If bVB_OPOSScannerEnable Then
                        With oScanner
                            If .DeviceEnabled Then
                                .DeviceEnabled = False
                                .Release
                                .Close
                            End If
                        End With
                    End If

                    tUT_FuncKBName = tProc  'strore name of function kb
                    bMainFunc = False 'เรียกคำสั่ง CallByName จากหน้าจอ main
                    Call CallByName(wFunctionKB, tProc, VbMethod)
                    
                    '*KT 54-04-09 V.3.0.56 open scanner
                    If bVB_OPOSScannerEnable Then
                        With oScanner
                                If .DeviceEnabled = False Then
                                     .Open tVB_Scanner
                                    .Claim 0 'TeamPos2000
                                     'Enable the device.
                                     .DeviceEnabled = True
                                     'Enable the event.
                                     .DataEventEnabled = True
                                     .AutoDisable = True
                                     .ClearInput
                                End If
                        End With
                    End If
                    KeyCode = 0
                    bMainFunc = True
                    If Not bUT_FuncLockKB Then Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock '*KT 52-12-15 กรณีกด Keyboard ไม่มี function '*KT 54-04-08 V.3.0.56
                     If Not otmDisplay.Enabled Then otmDisplay.Enabled = True '*KT 54-04-08 V.3.0.56
                Case 65 To 90, 8, 46, 48 To 57, 110, 37 To 40
                    'none
                Case Else
                    'none
                     KeyCode = 0
            End Select
    End With
    Call SP_DATxRsNothing(orsFunc)
    Exit Sub
ErrHandle: '*KT 54-04-08 ป้องกันโปรแกรมหลุด V.3.0.56
    Call SP_DATxRsNothing(orsFunc)
    Call SP_MSGxShowError(Err.Number, Err.Description)

End Sub

Private Sub ogdForm_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'-----------------------------------------------------------------------------------------------------------------------
'cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
'กรุณาตรวจสอบความถูกต้องของ Communication Sheet ML-SF-381,ML-SF-382 ด้วยนะครับ
'-----------------------------------------------------------------------------------------------------------------------
    With ogdForm
        tW_KeyCode = KeyCode
        Select Case UCase(.ColKey(Col))
            Case UCase("FNSdtQty")
                Select Case KeyCode
                    Case vbKeyDelete
                        .EditText = ""
                End Select
            Case UCase("FCSdtSaleUnit")
                Select Case KeyCode
                    Case vbKeyDelete
                        .EditText = "0."
                End Select
                '*Tao 60-05-03 เมื่อตัวเลขมากกว่า 10 หลัก จะทำให้ Price Override ไม่ได้
                If Len(.TextMatrix(Row, .ColIndex("FCSdtSaleUnit"))) > 10 Then
                    .TextMatrix(Row, .ColIndex("FCSdtSaleUnit")) = Chr(Asc(KeyCode))
                     .EditText = Chr(Asc(KeyCode))
                End If
                '====================================================
        End Select
    End With
End Sub

Private Sub ogdForm_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'-----------------------------------------------------------------------------------------------------------------------
'cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
'กรุณาตรวจสอบความถูกต้องของ Communication Sheet ML-SF-381,ML-SF-382 ด้วยนะครับ
'-----------------------------------------------------------------------------------------------------------------------
    With ogdForm
        Call cUT.UT_KBDChkKeyAscii(KeyAscii)
        Select Case UCase(.ColKey(Col))
            Case UCase("FNSdtQty")
                If KeyAscii = 46 Then KeyAscii = 0 'Not allow dot '*KT 52-11-26 PH 3.0-ML-SF-034  กด Price Override เป็นเศษสตางค์  เครื่องมีข้อความเตื่อน แต่เมื่อกด Enter แล้วใส่จำนวนเงินใหม่ เมื่อพิมพ์ Receipt จะพิมพ์ค่าที่ป้อนผิด
                Select Case KeyAscii
                    Case 13
                        bW_GridEnter = True
                        nW_GridEnter = nW_GridEnter + 1
                        Call cUT.UT_KBDKeyNumeric(KeyAscii)
                        If IsNumeric(Chr(KeyAscii)) And Not bVB_FKeyQtyChange Then
                            If Len(.EditText) >= 6 Then KeyAscii = 0
                        End If
                        KeyAscii = 0
                        bVB_FKeyQtyChange = False
                        If nW_GridEnter = 1 Then ogdForm_AfterEdit Row, Col
                        Exit Sub
                    Case Else
                        If bW_GridEnter Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                        bW_GridEnter = False
                        Select Case KeyAscii
                            Case vbKeyDelete, vbKeyBack
                                'none
                            Case Else
                                Call cUT.UT_KBDKeyNumeric(KeyAscii)
                        End Select
                        If IsNumeric(Chr(KeyAscii)) And Not bVB_FKeyQtyChange Then
                            If Len(.EditText) >= 6 Then KeyAscii = 0
                        End If
                        bVB_FKeyQtyChange = False
                End Select
                 If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii)) 'ถ้า  KeyAscii = Enter หมายถึงจบการแก้ไขราคาต่อหน่วย
            Case UCase("FCSdtSaleUnit")
                If KeyAscii <> 13 Then
                    If (KeyAscii = vbKeyBack) Then      'backspace
                        If Len(Trim(tW_KeyPress)) > 0 Then
                            tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
                        End If
                        Exit Sub
                    End If
                    Call cUT.UT_KBDKeyCurrency(KeyAscii, tW_KeyPress)    'check key, currency
                    If Len(.EditText) >= 11 Then KeyAscii = 0
                Else
                    '*KT 51-07-02 Communication Sheet ML-SF-381
                    '======================================================
                    'จำเป้นต้องใช้ tW_KeyPress แทน .TextMatrix(Row, Col) เพราะว่าค่าของ .TextMatrix(Row, Col) บางครั้งยังไม่รวมตัวเลขที่ Key ล่าสุดจนผ่าน ChangEdit ก่อน
                    'ใครที่ต้องการแก้ กรุณาตรวจสอบให้รอบครอบด้วยนะ เพราะ Grid ต้อง Control ด้วย Keybord ได้
                    'เกิดปัญหาว่าตอน Run ผ่าน Code ทำงานได้แต่เวลา Complier ไม่ผ่าน
                    
                    '*KT 52-11-26 PH 3.0-ML-SF-034  กด Price Override เป็นเศษสตางค์  เครื่องมีข้อความเตื่อน แต่เมื่อกด Enter แล้วใส่จำนวนเงินใหม่ เมื่อพิมพ์ Receipt จะพิมพ์ค่าที่ป้อนผิด
'*Eaw 56-07-04 CommSheet TK-ISS3000-201-ST.xls ไม่ต้องเช็คเหรียญ
'                    If Not (W_CHKbBanknoteAlw(SP_VALcText2Double(SP_DEFtChangeNum(tW_KeyPress)))) Then
'                        KeyAscii = 0
'                        tW_KeyPress = ""
'                        bVB_SetFocus = False
'                        nW_GridEnter = 0
'                        bW_ChkDecimals = False
'                        Call SP_MSGnShowing("tms_CN008028", nCS_Warning)   'คุณป้อนเศษสตางค์ไม่ถูกต้อง
'                        .Enabled = True
'                        .HighLight = flexHighlightAlways
'                        .TextMatrix(.Row, .ColIndex("FCSdtSaleUnit")) = SP_FMTcCurPoint(cW_SaleItem, True, True)
'                        .EditText = cW_SaleItem
'                        .Select Row, 1        'move highlight to first row
'                        .Col = .ColIndex("FCSdtSaleUnit")
'                        .EditCell
'                         tW_KeyPress = "" 'Replace(.TextMatrix(.Row, .Col), ".", "")
'                         bVB_PriceOveride = False
'                         Exit Sub     '*AO51-03-26   Comm.Sheet  ML-SF-100
'                    Else
'                         bW_ChkDecimals = True
'                    End If
                    '======================================================
                    nW_GridEnter = nW_GridEnter + 1
                    If nW_GridEnter = 1 Then ogdForm_AfterEdit Row, Col
                    Exit Sub
                End If
                If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii)) 'ถ้า  KeyAscii = Enter หมายถึงจบการแก้ไขราคาต่อหน่วย
        End Select
    End With
End Sub

Private Sub ogdForm_LostFocus()
        With ogdForm
            bUT_VoidBrows = False
            .Select .Row, 1
            .HighLight = flexHighlightNever
            .Editable = flexEDNone
            .Enabled = False
             otbPrompt.Locked = False
             bVB_SetFocus = True
        End With
End Sub

Private Sub oimLogo_DblClick()
    
On Error GoTo ErrHandle
    If bVb_SetResorution = False Then
        With wMain.oScanner
            If bVB_OPOSScannerEnable Then
                If .DeviceEnabled Then
                    .DeviceEnabled = False
                    .Release
                    .Close
                End If
            End If
        End With
'        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "ออกชั่วคราว" & " Form:" & Me.Name & " Control:oimLogo" & " Events:DblClick")
        Call SP_APPxEndProgram
    End If
    Exit Sub
ErrHandle:
    
End Sub

'Private Sub oimLogo_DblClick(ByVal Button As Integer)
'    'If bVb_SetResorution = False Then
'On Error Resume Next
'        With wMain.oScanner
'            If bVB_OPOSScannerEnable Then
'                If .DeviceEnabled Then
'                    .DeviceEnabled = False
'                    .Release
'                    .Close
'                End If
'            End If
'        End With
'        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "ออกชั่วคราว" & " Form:" & Me.Name & " Control:oimLogo" & " Events:DblClick")
'        Call SP_APPxEndProgram
'    'End If
'End Sub

'Private Sub oimLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 And Shift = 0 Then
'         PopupMenu omnLogo, vbPopupMenuRightButton   '+ vbPopupMenuLeftAlign
'    End If
'End Sub

Private Sub olaFn_Click(Index As Integer)
'*KT 51-06-18 Comsheet  XXX  ให้ทำงานที่ wFunctionKB/W_KBXALT ที่เดี่ยวทั้งการคลิ้กจากรูปหรือใช้ Hot Key กด Alt
'     wShowKB.Show vbModal
'     bMainFunc = bVB_FactQtyChange
'     If Not bVB_FKBCancel Then Call wMain.otbPrompt_KeyDown(nVB_FKBKeyCode, nVB_FKBShift)
    If Not bUT_FuncLockKB Then Call CallByName(wFunctionKB, "W_KBXALT", VbMethod)
 '------------------------------------------------------------------
End Sub



'Private Sub omnLogo1_Click()
'    Dim orsTemp As ADODB.Recordset
'    Dim tFilter$, tName$
'    Dim tSql$
'
'    tFilter = tCS_CNFilterBMP & "|" & tCS_CNFilterGIF & "|" & tCS_CNFilterJPG & "|" & tCS_CNFilterICO & "|" & tCS_CNFilterPNG
'    ocdform.Flags = &H4 ' hide check box read only
'    tName = SP_SHWtDialogue(ocdform, nEN_Ocd1Open, tFilter)
'    'If Not SP_OBJbCheckFilePic(oimLogo, tName) Then Exit Sub
'
'
'    tVB_CNMainLogo = Dir(tName)
'    If (SP_FLEbManFile(Trim(tName), nEN_Fle1Exist)) Then
''        oimLogo.AutoSize = True
''        oimLogo.LoadImage_FromFile (tVB_CNMainLogo)
''        oimLogo.StretchQuality = True
'        tSql = "SELECT FTScfUsrValue  FROM  TSysConfig"
'        tSql = tSql & vbCrLf & "WHERE    FTScfCode = 'FLogo' AND FTScfSeq = '001'"
'        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
'        If Not (orsTemp.BOF Or orsTemp.EOF) Then
'             tSql = "UPDATE TSysConfig"
'             tSql = tSql & vbCrLf & "SET FTScfUsrValue = '" & tVB_CNMainLogo & "'"
'             tSql = tSql & vbCrLf & "WHERE    FTScfCode = 'FLogo' AND FTScfSeq = '001'"
'        Else
'             tSql = "INSERT INTO  TSysConfig"
'             tSql = tSql & vbCrLf & "("
'             tSql = tSql & vbCrLf & "FTScfCode,FTScfSeq,FTScfApp,FTScfAppUse,FTScfStaGrp"
'             tSql = tSql & vbCrLf & ",FTScfStaAll,FTScfNameLang1,FTScfNameLang2,FTScfNameLang3,FTScfNameLang4"
'             tSql = tSql & vbCrLf & ",FTScfDesLang1,FTScfDesLang2,FTScfDesLang3,FTScfDesLang4"
'             tSql = tSql & vbCrLf & ",FTScfEdit,FTScfEditType,FTScfEditLen,FTScfStaUse"
'             tSql = tSql & vbCrLf & ",FTScfStaShw,FTScfDefValue,FTScfDefRef,FTScfUsrValue"
'             tSql = tSql & vbCrLf & ",FTScfUsrRef,"
'             tSql = tSql & vbCrLf & tCS_CNLastFedIns
'             tSql = tSql & vbCrLf & "VALUES ("
'             tSql = tSql & vbCrLf & "'FLogo','001','All','SF','LO'"
'             tSql = tSql & vbCrLf & ",'1','','','',''"
'             tSql = tSql & vbCrLf & ",'','','',''"
'             tSql = tSql & vbCrLf & ",'','','',''"
'             tSql = tSql & vbCrLf & ",'','','','" & tName & "'"
'             tSql = tSql & vbCrLf & ",''"
'             tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
'        End If
'        Call SP_SQLvExecute(tSql, , oDB.ocnOnLine, , , , , True, True) 'Server
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'End Sub

'Private Sub omnLogo2_Click()
'    Dim tSql As String
'On Error GoTo ErrHandle
'    tSql = "UPDATE TSysConfig"
'    tSql = tSql & vbCrLf & "SET FTScfUsrValue = ''"
'    tSql = tSql & vbCrLf & "WHERE    FTScfCode = 'FLogo' AND FTScfSeq = '001'"
'    Call SP_SQLvExecute(tSql, , oDB.ocnOnLine, , , , , True, True) 'read master
'    oimLogo.ClearImage
'    Exit Sub
'ErrHandle:
'
'End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
    If opbWeight.Visible = True Then opbWeight.Visible = False
    '*KT 51-11-18 กรณีที่หน้าจอทิ้งไว้นานระบบจะ Lock TextPrompt กรณีที่ Scan จะ Scan ไม่ได้
    If Not bUT_FuncLockKB And Not bVB_TextPromptAct And bVB_CNStartBill And wMain.Name = Screen.ActiveForm.Name And Not bW_FirstLoad Then
        '*KT        'Scan สินค้า คูณ จำนวน ทำให้เครื่องแฮ๊ง
        'Terminal No. 4304, เวลาประมาณ 20.05 น. เกิดจาก (1) Scan SKU รายการที่ 16, (2) แก้ไขจำนวน
        '(3) ในระหว่างที่ Cashier คนที่ 1 กำลังแก้ไขจำนวน Cashier คนที่ 2 ทำการ Scan SKU จึงทำให้เครื่องแฮ๊ง
        'ต้อง End Task แล้ว Restart Terminal ใหม่
        '** ดูรูปภาพประกอบจาก Standby_MALL8_04032009\4304-M8
        '** คำแนะนำจาก FSBT ถ้าเป็นเหตุการณ์ลักษณะนี้  โปรแกรมควรมี ข้อความแจ้ง หรือ ป้องกัน ไม่ให้มี Process ซ้อนกัน 52-03-24 PH1.5-ML-SF-110
        If Not ogdForm.Enabled Then
            With oScanner

                'otbPrompt.Text = .ScanData '*Bump 56-02-19 Comment บรรทัดนี้
              '*Bump 56-02-19 เพิ่ม Code 8 บรรทัด
              '***************************************
                If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
                    .DecodeData = True
                    otbPrompt.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel) '*BG 56-11-02 Scan ครั้งไหนๆ ก็จะได้ PdtCode
                End If
                If .DeviceName = tSNR_TOUCH_SCANNER Then
'                    MsgBox "ScanDataType = " & .ScanDataType
'                    MsgBox "ScanData = " & .ScanData
'                    MsgBox "ScanDataLabel = " & .ScanDataLabel
                    otbPrompt.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel) '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
                Else
                    otbPrompt.Text = .ScanData '*เป็นอันเดิม
                End If
                '***************************************
                .ClearInput
                .DeviceEnabled = True
                .DataEventEnabled = True
'                If Not (GetInputState = 0) Then DoEvents '*Ao 53-01-14   DoEvents
                If otbPrompt.Text <> "" Then
                    '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
                    If Asc(Right(otbPrompt.Text, 1)) = 13 Then
                        otbPrompt.Text = Mid(otbPrompt.Text, 1, Len(otbPrompt.Text) - 1)
                    End If
                    Call otbPrompt_KeyDown(13, 0)
                End If
                'ทำการ Open Scanner อีกครั้ง ไม่นั้นครั้งต่อไปจะ Scan สินค้าไม่ได้ครับ
                '*KT 52-01-08 ไม่ต้อง Open Scanner อีกครั้ง ยกเว้นกรณี Form Active
    '            .Open tVB_Scanner
    '            .Claim 0 'TeamPos2000
    '            'Enable the device.
    '            .DeviceEnabled = True
    '            'Enable the event.
    '            .DataEventEnabled = True
    '            .AutoDisable = True
'                .ClearInput
            End With
        Else
             With oScanner
                If .DeviceEnabled = False Then
                    .Open tVB_Scanner
                    .Claim 0 'TeamPos2000
                    'Enable the device.
                    .DeviceEnabled = True
                    'Enable the event.
                    .DataEventEnabled = True
                    .AutoDisable = True
                    .ClearInput
                End If
            End With
        End If
    Else
        With oScanner
            If .DeviceEnabled = False Then
                 .Open tVB_Scanner
                .Claim 0 'TeamPos2000
                 'Enable the device.
                 .DeviceEnabled = True
                 'Enable the event.
                 .DataEventEnabled = True
                 .AutoDisable = True
                 .ClearInput
            End If
         End With
    End If
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Step:Scan item" & " Events:oScanner_DataEvent")
End Sub

Private Sub otbPrompt_GotFocus()
    '*KT 53-01-14- DoEvents
   If bW_FirstLoad Then '*KT 51-06-18 Communication Sheet ML-SF-129(5)
        tVB_CNMnuCurrent = "1700"
        If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
            Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:otbPrompt_GotFocus[1]")
            bW_FirstLoad = False
            Call otbPrompt_KeyDown(vbKeyEscape, 0)
            Exit Sub
        End If
         bW_FirstLoad = False
         Call W_FRMxActivate
    End If
    bVB_SetFocus = True
End Sub

Public Sub otbPrompt_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim tKeyCode$, tShift$
    Dim tProc$
    Dim orsFunc As ADODB.Recordset
    Dim tSql As String

On Error GoTo ErrHandle
    nW_CntSentSale = 0 '*Tao 61-09-24 เริ่มนับใหม่

    bVB_CheckVendor = True                  '*Ao 53-10-13   CheckVendor
    '*KT 51-09-17 แก้ไขเรื่องการยิงสินค้าขึ้น Zero
    '----------------------------------------------------------------------------------------------------------------
    If Not bVB_CNStartBill Then
        KeyCode = 0
        Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
        Exit Sub
    End If
    If bVB_TextPromptAct Then
        KeyCode = 0
        Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
        Exit Sub
    End If
    '---------------------------------------------------------------------------------------------------------------
    If opbWeight.Visible = True And (tVB_2LineCode <> "" Or cVB_2LinePrice > 0) Then   '*Eaw 56-02-05 2Linebarcode
        opbWeight.Visible = True
    Else
        opbWeight.Visible = False
    End If
'    If opbWeight.Visible = True Then
'        opbWeight.Visible = False
'        '*KT 52-05-06 กรณีการแสดงสินค้าน้ำหนักแล้วให้สามารถทำการ key สินค้าหรือ scan สินค้าต่อไปไดเลย
'        'Exit Sub
'    End If
    If olaChange.Visible = True And bVB_OPOSCashDrawerEnable Then Exit Sub
    bW_LockKeyAscii = False
    
    If bUT_FuncLockKB Then
        If KeyCode = 112 And Shift = 2 Then
            If otbPrompt.Visible And otbPrompt.Enabled Then Call SP_OBJxSetFocus(otbPrompt)   'otbPrompt.SetFocus
            If bW_Ok Then
                bUT_FuncLockKB = False
                otbPrompt.Text = ""         'no message for unlock
                KeyCode = 0
                bW_LockKeyAscii = True
                Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
                Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock
                Exit Sub
            End If
        Else
            KeyCode = 0
            bW_LockKeyAscii = True
            Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
            If otbPrompt.Visible And otbPrompt.Enabled Then Call SP_OBJxSetFocus(otbPrompt)  'otbPrompt.SetFocus
            Exit Sub
        End If
    Else
        If otmLock.Enabled = False Then
            Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock
        End If
    End If
    'If Shift = 4 And KeyCode <> 18 Then KeyCode = 0      'do not accept Alt
    Select Case KeyCode
        Case 13, 18, 27, 33, 34, 35, 36, 45, 112 To 124, 128, 129, 46
        
            '*KT 54-09-20 V.4.0.1 กรณีที่ไม่เป็น OPOS ในการแสดงเงินทอน
            If Not bVB_OPOSCashDrawerEnable And bVB_CNShwChange Then
                If nUT_SumQty = 0 Then
                     If wMain.olaChange.Visible Then
                        wMain.olaChange.Visible = False
                        olaPrice.Caption = ""
                     End If
                     bVB_CNShwChange = False
                End If
            End If

            'Enter, Esc, PgUp, PgDn, End, Home, Ins, Del, F1 to F12,Del
            If bVB_IPV Then
                Select Case KeyCode
                    Case vbKeyHome
                        'non
                    Case Else
                        otbPrompt.Text = ""
                        KeyCode = 0
                        Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
                        Exit Sub
                End Select
            End If
            Call SP_FRMxSetTimeCount(False)       'จับเวลา Lock
            tSql = "SELECT FTSkbRef FROM TSysKeyboard   WHERE FNSkbKey='" & CStr(KeyCode) & "'"
            tSql = tSql & " AND FNSkbShift = '" & CStr(Shift) & "'"
            tSql = tSql & " AND FTSkbwMain =  '1'"
            tSql = tSql & " AND FTSkbUsage ='Y'"
            Call SP_SQLvExecute(tSql, orsFunc, oDB.ocnOnLine, , , , , True, True) 'read master
            If orsFunc.EOF Then
                Call SP_MSGnShowing("tms_CN005107", nCS_Information)
                KeyCode = 0
'                orsFunc.Close
                Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
                If Not bUT_FuncLockKB Then Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock '*KT 52-12-15 กรณีกด Keyboard ไม่มี function
                Exit Sub
            End If
'            nUT_UserLevelFunc = SP_VALcText2Double(tUT_UsrLevel)
            tProc = SP_FEDtChkString(orsFunc, "FTSkbRef")
            orsFunc.Close
            tUT_FuncKBName = tProc  'strore name of function kb
            
            '*Eaw 56-02-06  ถ้ามีการจ่ายเงินบางส่วนแล้ว ไม่สามารถทำรายการใดๆได้อีก ยกเว้นการจ่ายให้ครบจำนวน
'            If bVB_CNPaid Or bVB_CNDiscount Then
            If bVB_CNDiscount Then
                If tProc <> "W_KBXHomeSUB" Then
                    Call SP_MSGnShowing("tms_CN005107", nCS_Information)
                    otbPrompt.Text = ""
                    Call SP_OTBxSelText(otbPrompt)
                    Exit Sub
                End If
            End If
            '-------------------------------------------------------------------------------------------------
            If (tProc = "W_KBRcvVenCpnCancel") Or (tProc = "W_KBRcvVenCpnChange") Or (tProc = "") Then                          '*Ao 2011/02/24  CommSheet ML-PRV-041-Runtime Error '*KT 54-04-08 V.3.0.56 ให้ทำการ Exit Sub ก่อนจะไปทำการ Close Scanner
                Call SP_MSGnShowing("tms_CN005107", nCS_Information)
                KeyCode = 0
'                orsFunc.Close
                Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
                If Not bUT_FuncLockKB Then Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock '*KT 52-12-15 กรณีกด Keyboard ไม่มี function
                Exit Sub
            End If
            If tVB_2LineCode <> "" Or cVB_2LinePrice > 0 Then  '*Eaw 56-02-05 2Linebarcode
                If tProc = "W_KBXEsc" Then   '*Eaw 56-03-02     เพิ่ม if นี้มา
                    opbWeight.Visible = False
                    tVB_2LineCode = ""     '*Eaw 56-03-05
                    cVB_2LinePrice = 0   '*Eaw 56-03-05
                    Exit Sub
                ElseIf tProc <> "W_KBXEnter" Then
                    tVB_2LineCode = ""
                    cVB_2LinePrice = 0
                    opbWeight.Visible = False
                Else
                    opbWeight.Visible = True
                End If
'                If tProc <> "W_KBXEnter" Then
'                    tVB_2LineCode = ""
'                    cVB_2LinePrice = 0
'                    opbWeight.Visible = False
'                Else
'                    opbWeight.Visible = True
'                End If

            End If
            If (tProc = "W_KBDiscount91") Then                           '*Eaw 56-01-31 แก้ไขเรื่อง Hot Key ของ Store Wide Discount
                Call SP_MSGnShowing("tms_CN005245", nCS_Error)
                KeyCode = 0
'                orsFunc.Close
                Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
                If Not bUT_FuncLockKB Then Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock '*KT 52-12-15 กรณีกด Keyboard ไม่มี function
                Exit Sub
            End If
            'bMainFunc = True 'เรียกคำสั่ง CallByName จากหน้าจอ main
            If KeyCode = vbKeyHome Then
                If SP_VALcText2Double(olaTSubTotal.Caption) > cCS_CNMaxSale Then
                    Call SP_MSGnShowing("tms_CN005170", nCS_Warning)
                    KeyCode = 0
                    Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
                     If Not bUT_FuncLockKB Then Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock '*KT 52-12-15 กรณีกด Keyboard ไม่มี function
                    Exit Sub
                End If
                '------------------------------------------------------------------------
                Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax           '*Ao 2014/09/05   V 4.2.143
                    If bMainFunc Then
                        bVB_CNTrnDiscEx = True
                    End If
                End Select
                '------------------------------------------------------------------------
                Call SP_PRCxMve2DTCDTmp
                Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnSale, tEN_TrnDeposit, tEN_TrnVIPSale, tEN_TrnTaxExceptSale, tEN_TrnDutyFreeSale
                        Call DE_PRCxDeal
                End Select
                bVB_CNCalAutoTran = False '*KT 51-05-29 ML-SF-275
            End If
            If tProc <> "" And tProc <> "W_KBXEnter" Then
                If bVB_OPOSScannerEnable Then
                    With oScanner
                        If .DeviceEnabled Then
                            .DeviceEnabled = False
                            .Release
                            .Close
                        End If
                    End With
                End If
            End If
            If tProc <> "W_KBLockKB" Then otmDisplay.Enabled = False  'ไม่ต้องการให้ทำงานหลังจากมีการมี key text promt
            bVB_TextPromptAct = True
            
            Call CallByName(wFunctionKB, tProc, VbMethod)
'            If Not (GetInputState = 0) Then DoEvents '*KT 53-01-14 DoEvents

            If bVB_OPOSScannerEnable Then
                'If tProc = "W_KBXEsc" And tVB_Transaction = "99" Then '*KT 52-11-20 เก็บ log สำหรับกรณีที่ Close Terminal
                '    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:W_KBXEsc" & " Form:" & Me.Name & " Step:Close Terminal[1]" & " Events:otbPrompt_KeyDown")
                'End If
                With oScanner
                    If .DeviceEnabled = False Then
                        .Open tVB_Scanner
                        .Claim 0 'TeamPos2000
                         'Enable the device.
                        .DeviceEnabled = True
                         'Enable the event.
                        .DataEventEnabled = True
                        .AutoDisable = True
                        .ClearInput
                        'If tProc = "W_KBXEsc" And tVB_Transaction = "99" Then '*KT 52-11-20 เก็บ log สำหรับกรณีที่ Close Terminal
                        '    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:W_KBXEsc" & " Form:" & Me.Name & " Step:Close Terminal[2]" & " Events:otbPrompt_KeyDown")
                        'End If
                    End If
                End With
                'If tProc = "W_KBXEsc" And tVB_Transaction = "99" Then '*KT 52-11-20 เก็บ log สำหรับกรณีที่ Close Terminal
                '    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:W_KBXEsc" & " Form:" & Me.Name & " Step:Close Terminal[3]" & " Events:otbPrompt_KeyDown")
                'End If
            End If
            bVB_TextPromptAct = False
            KeyCode = 0
            
            If bUT_SignOut Then
                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:otbPrompt_KeyDown")
                Unload Me
                Call wBackGrand.W_PRVxFormActivate             '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
                Exit Sub     'signout delete  maxsaledoc from transaction table ,update (TPSMPos,TPSTLogIn)
            End If
            If Not bUT_FuncLockKB Then Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock
            If Not otmDisplay.Enabled Then otmDisplay.Enabled = True
            bW_LockKeyAscii = True
            If Not wMain.otbPrompt.Locked Then
                'otbPrompt.SetFocus
                 Call SP_OBJxSetFocus(otbPrompt)
            End If
        Case 65 To 90, 8, 46, 48 To 57, 110, 37 To 40
            'A to Z, Del, BackSpace, 0 to 9, decimal point(.),<--  ^  V  --> (ลูกศร)
            'do nothing, just keep keycode
    
        Case Else   'other than accepted should be set as nothing
            KeyCode = 0
    End Select
    Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsFunc) '*KT 52-12-15 ต้องทำลาย object recordset
    bVB_TextPromptAct = False
    If Err.Number = 438 Then            'not define function name in the database
        'Call SP_MSGnShowing("tms_CN999022", nCS_Information)    'tCS_MSG22 ฟังก์ชันนี้ยังไม่ได้กำหนด หรือไม่ได้รับอนุญาตให้ใช้งาน;This function is undefined or not allowed to use.
        Call SP_MSGxShowError(Err.Number, Err.Description) '*KT 54-04-09 V.3.0.56 เปลี่ยน message
    ElseIf Err.Number = 364 Then
       If bUT_SignOut Then
            Call SP_CRTxWriteLogTerminated("ErrorNumber:364" & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:otbPrompt_KeyDown") '*KT 54-04-09 V.3.0.56 เก็บ Log
            Unload Me
            Call wBackGrand.W_PRVxFormActivate             '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls '*KT 54-04-09 ให้ทำการเรียก wBackGrand.W_PRVxFormActivate   ด้วย
            Exit Sub
        End If
    ElseIf Err.Number = 440 Then 'Automation error
        Call SP_MSGxShowError(Err.Number, Err.Description)
    Else
        Call SP_MSGxShowError(Err.Number, Err.Description)
    End If
End Sub


Private Sub W_PRCxCancelPdtEx(ByVal ptPdtCode As String)
    Dim nMax As Long
    Dim nRow As Long
    Dim nLoop As Long
    Dim bFound As Boolean
On Error GoTo ErrHandle
    With ogdForm
            nMax = .Aggregate(flexSTMax, .FixedRows, 0, .rows - 1, 0)
            If nMax < 1 Then Exit Sub
            bFound = False
            For nRow = .FixedRows To nMax
                If .TextMatrix(nRow, 8) = ptPdtCode And .TextMatrix(nRow, 9) = "1" Then
                    bFound = True
                    Exit For
                End If
            Next nRow
            If bFound Then
                .Enabled = True
                .HighLight = flexHighlightWithFocus
                .Select nRow, 1                  'move highlight to first row
                .SetFocus
                Call ogdForm_KeyDown(vbKeyF3, 0)
            Else
                MsgBox "No Found", vbCritical
            End If
    End With
    Exit Sub
ErrHandle:
    
End Sub


Public Function W_InsertTSaleDT(ByVal pbRfn As Boolean) As Boolean
'------------------------------------------------
'Call:
'Cmt: *KT 50-12-25 Insert Data To DT
'-----------------------------------------------

    Dim nSeq%, cNet As Double
    Dim tSql$, tSkuBarC$, tSKUName$
    Dim tSKUName2$, tSKUNameSrvDoc$, tSKUNameSrvDoc2$
    Dim cVat As Double, cVatable As Double
    Dim cCostIn As Double, cQtyAll As Double
    Dim cQty As Double, cRegPrice As Double, cSalePrice As Double, cB4DisChg As Double
    Dim cNetDuty As Double, cQtyDuty As Double
    Dim tPromptSale$
    Dim tPromotion As String
    Dim nEff As Long
    Dim tSkuIndex As String
    Dim cTotalBFVat As Double
    
On Error GoTo ErrHandle
    W_InsertTSaleDT = False
    If tVB_SdtTaxInv = "" Then tVB_SdtTaxInv = "N" '*Bump 58-01-09 เก็บ Flag 'Y' ถ้า สินค้ามี Line Discount สำหรับการทำ Return Refer Sale
    With ogdForm
        nSeq = .Aggregate(flexSTCount, .FixedRows, .ColIndex("FNSdtSeqNo"), .rows - 1, .ColIndex("FNSdtSeqNo"))
        tSkuBarC = Trim(.TextMatrix(nSeq, .ColIndex("FTSdtBarCode")))  'รหัสบาร์โค้ด
        tSKUName = Trim(.TextMatrix(nSeq, .ColIndex("FTSkuAbbName")))     'ชื่อสินค้า(Default ENG)1
        tSKUName2 = Trim(.TextMatrix(nSeq, .ColIndex("FTSkuAbbNameSnd")))     'ชื่อสินค้า(Local)2
        tSKUNameSrvDoc = Trim(.TextMatrix(nSeq, .ColIndex("FTSkuAbbNameSrvDoc")))     'ชื่อสินค้า Service Doc(Default ENG) 1
        tSKUNameSrvDoc2 = Trim(.TextMatrix(nSeq, .ColIndex("FTSkuAbbNameSndSrvDoc")))     'ชื่อสินค้า Service Doc(Local) 2
        cQty = .ValueMatrix(nSeq, .ColIndex("FCSdtQty"))                              'จำนวนไม่ใช้จาก ตัวแปร cUT_Qty(เดิม), ใช้จาก grid จะครอบคลุมจำนวนแบบ เครื่องชั่ง และน้ำหนักด้วย (ใหม่)
        cRegPrice = .ValueMatrix(nSeq, .ColIndex("FCSdtRegPrice")) 'Regular Price
        cSalePrice = .ValueMatrix(nSeq, .ColIndex("FCSdtSalePrice")) 'Sale Pric
        cNet = .ValueMatrix(nSeq, .ColIndex("FCSdtSaleAmt"))                              'รวมเงิน
        '*KT 53-02-11 Communication Sheet PH3-ML-SF-084 When scroll cancel item which is discounted from Mall card.xls
        'กรณีที่มีการ Recal ของ Return แบบ Refer
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                If bVB_Reference Then
                    cB4DisChg = cRegPrice
                    tSkuIndex = tVB_CNSkuIndex              '*Ao 2011/02/24  CommSheet ML-PRV-042-EXPfile
                Else
                    cB4DisChg = cSalePrice
                    tSkuIndex = nSeq                                     '*Neung 2011/10/24 CommSheet ML-PRV-071 V.3.0.68
                End If
'                tSkuIndex = tVB_CNSkuIndex              '*Ao 2011/02/24  CommSheet ML-PRV-042-EXPfile
            Case Else
                cB4DisChg = cSalePrice
                tSkuIndex = nSeq                                    '*Ao 2011/02/24  CommSheet ML-PRV-042-EXPfile
        End Select

        'cVat = IIf(tVB_SKUTaxCode <> "00", SP_FMTcCurPoint(SP_CALcVatValue(IIf(tVB_CNRetVATInOrEx = "1", True, False), cNet, cVB_SKUTaxRate), True, False), 0)                            '1:vat, 2:novat (0)
         '*KT 66-08-09  case tax code "00" non tax   or  tax code "02" zero vat  set vat valuse=0
        'cVat = IIf(tVB_SKUTaxCode <> "00", SP_FMTcCurPoint(SP_CALcVatValue(IIf(tVB_CNRetVATInOrEx = "2", True, False), cNet, cVB_SKUTaxRate), True, False), 0)   ''tVB_CNRetVATInOrEx 1:รวมใน, 2: แยกนอก แก้ไขให้ถูกต้องตามการคำนวณใน SP_CALcVatValue '*Bump 66-03-23
        cVat = IIf((tVB_SKUTaxCode = "00") Or (tVB_SKUTaxCode = "02" And cVB_SKUTaxRate = 0), 0, SP_FMTcCurPoint(SP_CALcVatValue(IIf(tVB_CNRetVATInOrEx = "2", True, False), cNet, cVB_SKUTaxRate), True, False))         '*tVB_CNRetVATInOrEx 1:รวมใน, 2: แยกนอก แก้ไขให้ถูกต้องตามการคำนวณใน SP_CALcVatValue '*Bump 66-03-23
        cVatable = IIf(tVB_SKUTaxCode <> "00", cNet - cVat, 0)
        
        cQtyAll = cQty * cUT_StkFac * cUT_SKUUnitF
        
        tPromptSale = .TextMatrix(nSeq, .ColIndex("FTSdtPriceOverideFlag"))
        
        cNetDuty = IIf(tVB_Transaction = Format(tEN_TrnDutyFreeSale, "00"), cNet, 0)
        cQtyDuty = IIf(tVB_Transaction = Format(tEN_TrnDutyFreeSale, "00"), cQtyAll, 0)
                                    
        .TextMatrix(nSeq, .ColIndex("FCSdtVat")) = cVat
        .TextMatrix(nSeq, .ColIndex("FCSdtVatable")) = cVatable
        
        'จะเช็คเกี่ยวกับ FTSkuStaMallEmpCard='N' จะมี '*' อยู่หน้า BarCode ยังไม่เกี่ยวกับ Promotion
        tPromotion = IIf(Trim(.TextMatrix(nSeq, .ColIndex("FTPmt"))) = "", "Y", "N")       '*AO 51-04-02  Comm.Sheet ML-SF-179
        '*Ao 2013/01/31  TK Req
        cTotalBFVat = SP_FMTcCurPoint((cRegPrice * cQtyAll) - (((cRegPrice * cQtyAll) * cVB_SKUTaxRate) / (100 + cVB_SKUTaxRate)), True, False)
        
        Select Case Val(Format(tVB_TransactionSub)) '*Eaw 56-08-26 CommSheet TK-ISS3000-283 TR FAPL-63 00.00.04.xls
            Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax
                cVatable = cNet
                cTotalBFVat = cNet
        End Select
        tSql = "INSERT INTO " & tVB_CNTblDT
        tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode,"
        tSql = tSql & vbCrLf & "FTSkuAbbName,FTShdTransType,FDShdTransDate,FTSdtBarCode,FTSdtStkCode,"
        tSql = tSql & vbCrLf & "FCSdtStkFac,FTTaxCode,FTDpcCode,FTSkuDepCode,FTSkuBrdCode,"
        tSql = tSql & vbCrLf & "FTPunCode,FTSdtUnitName,FCSdtRegPrice,FTSdtEventPrice,FCSdtQty,"
        tSql = tSql & vbCrLf & "FCSdtSaleTaxAmt,FCSdtReturnTaxAmt,"
        tSql = tSql & vbCrLf & "FTSdtInstoreBar,FCSdtB4DisChg,FTSdtDisChgTxt,FCSdtDis,"
        tSql = tSql & vbCrLf & "FCSdtChg,FCSdtSaleAmt,FCSdtVat,FCSdtVatable,FCSdtQtyAll,"
        tSql = tSql & vbCrLf & "FTSdtStaPdt,"
        tSql = tSql & vbCrLf & "FTSdtWeight,FTSdtConsignFlag,"
        tSql = tSql & vbCrLf & "FTSdtCancelFlag,FCSdtTax,FCSdtSalePrice,FTShdTransTime,"
        tSql = tSql & vbCrLf & "FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc,"
        tSql = tSql & vbCrLf & "FTTaxSign,FTTaxStaSign,"
        tSql = tSql & vbCrLf & "FCSdtSaleUnit,"
        tSql = tSql & vbCrLf & "FCSdtQtyAllDisAvi,"
        tSql = tSql & vbCrLf & "FCSdtAmtDisAvi,"
        tSql = tSql & vbCrLf & "FCSdtDutyFreeAmt,"
        tSql = tSql & vbCrLf & "FCSdtDutyFreeQty,"
        tSql = tSql & vbCrLf & "FCSdtDutyDeposit,"
        tSql = tSql & vbCrLf & "FCSdtWeight,"
        tSql = tSql & vbCrLf & "FTSdtPriceOverideFlag,"
        tSql = tSql & vbCrLf & "FCSdtSalePricePmt,"
        tSql = tSql & vbCrLf & "FTSkuStaMallEmpCard , "   '*AO 51-04-02  เพิ่ม Field FTSkuStaMallEmpCard"
        tSql = tSql & vbCrLf & "FNSdtStaRef,"
        tSql = tSql & vbCrLf & "FCSdtBSubPro,"
        tSql = tSql & vbCrLf & "FCSdtBSubReg,"
        tSql = tSql & vbCrLf & "FCSdtDiscSubPro,"
        tSql = tSql & vbCrLf & "FCSdtDiscSubReg,"
        tSql = tSql & vbCrLf & "FCSdtYellowPrice,"
        tSql = tSql & vbCrLf & "FCSdtBluePrice,"
        tSql = tSql & vbCrLf & "FTSdtStaRfd," '*KT 52-02-13 Amount of member price
        tSql = tSql & vbCrLf & "FCSdtMemberPrice," '*KT 52-02-13 Amount of member price
        tSql = tSql & vbCrLf & "FTSdtEventMemPrice," '*KT 52-02-18 Event ID of member price
        tSql = tSql & vbCrLf & "FCSdtB4DisChgOrg," '*KT 52-02-18 ทำการเก็บค่าเดิม FCSdtB4DisChg ก่อนได้ member card
        tSql = tSql & vbCrLf & "FCSdtPriceOveride," '*KT 52-02-13 price overide
        tSql = tSql & vbCrLf & "FTSdtMemberPointAlw," '*KT 52-06-15 Member Point By SKU Y: Allow N: Not Allow
        tSql = tSql & vbCrLf & "FCSdtMemberPoint,"  '*KT 52-06-15 Amount of member price
        tSql = tSql & vbCrLf & "FTSdtEventMemPoint," '*KT 52-06-15 Event ID Member Price
        tSql = tSql & vbCrLf & "FTSdtMemPointCooAlw," '*KT 52-06-15 Y: Allow Member Point(Cooperate[General]) N: Not Allow Member Point(Cooperate[General])
        tSql = tSql & vbCrLf & "FTSkuIndex,"
        tSql = tSql & vbCrLf & "FTSdtStaPreview," '*KT 53-10-30 Flag แสดงสถานะสินค้า Preview Day
        tSql = tSql & vbCrLf & "FTSdtStaNotAlwPrv," '*KT 53-10-30 Flag แสดงสถานะสินค้า Preview Day
        tSql = tSql & vbCrLf & "FCAFGroupPM,FCOverrideAmt,"             '*Ao 2010/12/16 Price Override
        tSql = tSql & vbCrLf & "FTSdtSpnCode,"             '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
        tSql = tSql & vbCrLf & "FTStdStaFixPointFlag," '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
        tSql = tSql & vbCrLf & "FNSdtSKUPointTime," '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
        tSql = tSql & vbCrLf & "FCSdtTotalBFVat,FTPdtNoDis," '*Ao 2013/01/31 TK Req
        tSql = tSql & vbCrLf & "FTSdtTaxInv," '*Eaw 56-08-26 CommSheet TK-ISS3000-317-UAT-FAPL-93.xls
        tSql = tSql & vbCrLf & tCS_CNLastFedIns
        tSql = tSql & vbCrLf & " VALUES ("
        tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "','" & nSeq & "','" & tUT_SKUCode & "',"
        tSql = tSql & vbCrLf & "'" & SP_STRtSiggleCode(tSKUName) & "','" & tVB_Transaction & "','" & tUT_SaleDate & "','" & IIf(oSKU.GETInStoreBar, tUT_TextPrompt, tSkuBarC) & "','',"
        tSql = tSql & vbCrLf & SP_FMTcCurPoint(cUT_StkFac, True, False) & ",'" & tVB_SKUTaxCode & "','" & tVB_CNDpcCode & "','" & tVB_CNSkuDepCode & "','" & tVB_CNSkuBrdCode & "',"
        tSql = tSql & vbCrLf & "'" & tUT_SKUUnitC & "','" & tUT_SKUUnitN & "'," & SP_FMTcCurPoint(cRegPrice, True, False) & ",'" & tVB_CNEventPrice & "'," & SP_FMTcCurPoint(cQty, True, False) & ","
        tSql = tSql & vbCrLf & "0,0,"
        tSql = tSql & vbCrLf & "'" & IIf(oSKU.GETInStoreBar, tVB_InStoreBar, "") & "'," & cB4DisChg & ",''," & cUT_DisAmt & "," '*KT 52-03-19 เก็บค่า InStoreBar ใหม่ต้องเป็นหมายเลขรหัสที่ยิงบาร์โค้ดสินค้าประเภทน้ำหนัก
        tSql = tSql & vbCrLf & cUT_ChgAmt & "," & cNet & "," & cVat & "," & cVatable & "," & cQtyAll & ","
        tSql = tSql & vbCrLf & "'" & Format(tVB_CNSkuStaAlw, "0000000000000000000000") & "'," '*KT 54-06-20 V.4.0.0 Ph.4.0 Req.13 การให้ Special Point(SKU Point)และสิทธิส่วนลด, '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow  '*KT 55-08-09 PRJ10001TK-Internal Specification-SKU-00.02.00.doc '*KT 67-12-12 TTK-ISS3000-885,TTK-ISS3000-886 เก็บสถานะ  FTSkuItemType  BP Point  -22
        tSql = tSql & vbCrLf & "'" & tVB_Weight & "','" & tVB_SKUConsignFlag & "'," '*KT 52-08-06
        tSql = tSql & vbCrLf & "'0'," & cVB_SKUTaxRate & "," & cSalePrice & ",'" & Time$ & "',"
        tSql = tSql & vbCrLf & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "',"
        tSql = tSql & vbCrLf & "'" & tVB_SKUTaxSign & "','" & tVB_SKUTaxStaSign & "',"
        tSql = tSql & vbCrLf & cSalePrice & ","
        tSql = tSql & vbCrLf & cQtyAll & ","
        tSql = tSql & vbCrLf & cNet & ","
        tSql = tSql & vbCrLf & cNetDuty & ","
        tSql = tSql & vbCrLf & cQtyDuty & ","
        tSql = tSql & vbCrLf & SP_FMTcCurPoint(cQtyAll * cVB_DutyFreeDeposit, True, False) & ","
        tSql = tSql & vbCrLf & cVB_Weight & ","
        tSql = tSql & vbCrLf & "'" & tPromptSale & "',"
        tSql = tSql & vbCrLf & cB4DisChg & "," 'ราคาสำหรับคำนวณ Deal 5,6
        tSql = tSql & vbCrLf & "'" & tPromotion & "',"     '*AO 51-04-02
        tSql = tSql & vbCrLf & nSeq & ","
        tSql = tSql & vbCrLf & IIf(tPromotion = "N", cNet, 0) & ","
        tSql = tSql & vbCrLf & IIf(tPromotion = "N", 0, cNet) & ","
'        tSql = tSql & vbCrLf & IIf(tPromotion = "N", cNet, 0) & "," '*KT 55-06-21 ปรับข้อมูล FCSdtDiscSubPro,FCSdtDiscSubReg,FCSdtBluePrice,FCSdtYellowPrice
'        tSql = tSql & vbCrLf & IIf(tPromotion = "N", 0, cNet) & ","
        tSql = tSql & vbCrLf & 0 & ","
        tSql = tSql & vbCrLf & 0 & ","
        tSql = tSql & vbCrLf & IIf(tPromotion = "N", cNet, 0) & ","
        tSql = tSql & vbCrLf & IIf(tPromotion = "N", 0, cNet) & ","
        tSql = tSql & vbCrLf & "'" & tVB_StaMemPriceAlw & "',"     '*KT 52-02-13 Amount of member price
        tSql = tSql & vbCrLf & SP_FMTcCurPoint(cUT_MemberPrice, True, False) & ","  '*KT 52-02-13 Amount of member price
        tSql = tSql & vbCrLf & "'" & tVB_CNEventMemPrice & "',"
        tSql = tSql & vbCrLf & cB4DisChg & ","
        tSql = tSql & vbCrLf & IIf(tPromptSale = "N", 0, cRegPrice) & ","
        tSql = tSql & vbCrLf & "'" & tVB_CNSkuPointAlw & "',"
        tSql = tSql & vbCrLf & cVB_CNSkuPointAmt & ","
        tSql = tSql & vbCrLf & "'" & tVB_CNSkuEventPoint & "',"
        tSql = tSql & vbCrLf & "'" & tVB_CNSkuPointCooAlw & "',"
        tSql = tSql & vbCrLf & "'" & tSkuIndex & "',"
        tSql = tSql & vbCrLf & "'" & tVB_CNSKUFlagPrv & "'," '*KT 53-10-30 Flag แสดงสถานะสินค้า Preview Day
        tSql = tSql & vbCrLf & "'" & tVB_CNSKUFlagNotAlwPrvUse & "'," '*KT 53-10-30 Flag แสดงสถานะสินค้า Preview Day
        tSql = tSql & vbCrLf & cVB_AFGroupPM & "," & cVB_OverrideAmt & ","    '*Ao 2010/12/16 Price Override
        tSql = tSql & vbCrLf & "'" & tVB_SalemanID & "',"             '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
        tSql = tSql & vbCrLf & "'" & tVB_CNSKUStaFixPointFlag & "',"             '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
        '*KT 54-11-21 V.4.0.5 DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                If bVB_Reference Then
                    tSql = tSql & vbCrLf & IIf(tVB_CNSKUStaFixPointFlag = "Y", 0, cVB_CNSkuPointTime) & ","           '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
                Else
                    tSql = tSql & vbCrLf & IIf(tVB_CNSKUStaFixPointFlag = "Y", 0, cVB_CNSkuPointAmt) & ","           '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
                End If
            Case Else
                tSql = tSql & vbCrLf & IIf(tVB_CNSKUStaFixPointFlag = "Y", 0, cVB_CNSkuPointAmt) & ","           '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
        End Select
        'tSql = tSql & vbCrLf & IIf(tVB_CNSKUStaFixPointFlag = "Y", 0, cVB_CNSkuPointAmt) & ","           '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
        tSql = tSql & vbCrLf & cTotalBFVat & ",'" & tVB_PdtNoDis & "',"            '*Ao 2013/01/31
        'tSql = tSql & vbCrLf & "'N'," '*Eaw 56-08-26 CommSheet TK-ISS3000-317-UAT-FAPL-93.xls ตอน Insert ให้ SKU ทุกตัวมีค่า Coop flag='N'
         tSql = tSql & vbCrLf & "RTRIM(LTRIM('" & Trim(tVB_SdtTaxInv) & "'))," '*Bump 58-01-09 ใช้เก็บ Flag 'Y' ถ้า สินค้ามี Line Discount สำหรับการทำ Return Refer Sale'*TON 58-10-16 เพิ่มการตัดช่องว่าง
        tSql = tSql & vbCrLf & UT_tLastInsSql & ")"
    End With
    If SP_SQLvExecute(tSql, , , , , nEff) = 0 Then
        If nEff > 0 Then
            W_InsertTSaleDT = True
        Else
            W_InsertTSaleDT = False
        End If
    Else
        W_InsertTSaleDT = False
    End If
    Exit Function
ErrHandle:
    W_InsertTSaleDT = False
    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "W_InsertTSaleDT" & " Form:" & Me.Name & " Control:Form" & " Events:W_InsertTSaleDT")
End Function

Public Sub W_UpdateTSaleDT(ByVal pnRow As Long, ByVal pnCol As Long)
'-----------------------------------------------------------
'   Call:
'   Cmt:    update data into DT
'----------------------------------------------------------
        Dim tSql$
        Dim tSqlF$
        Dim tSqlV$
        Dim nI As Long
        Dim cQty As Double
        Dim cQtyAll As Double
        Dim cVat As Double
        Dim cVatable As Double
        Dim cNet As Double
        Dim cSalePrice As Double, cRegPrice As Double, cTotalBFVat As Double
        Dim nEff As Long

On Error GoTo ErrHandle
        
        With ogdForm
            tSql = "UPDATE " & tVB_CNTblDT & " SET "
            Select Case .ColKey(pnCol)
                Case "FCSdtSaleUnit"
                    cQty = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtQty")))                               'จำนวนไม่ใช้จาก ตัวแปร cUT_Qty(เดิม), ใช้จาก grid จะครอบคลุมจำนวนแบบ เครื่องชั่ง และน้ำหนักด้วย (ใหม่)
                    cSalePrice = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtSalePrice")))  'Sale Pric
                    'MsgBox cSalePrice
                    cNet = SP_FMTcCurPoint(cQty * cSalePrice, True, False)                         'รวมเงิน
                    cVB_SKUTaxRate = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtTax")))
                    tVB_SKUTaxCode = .TextMatrix(pnRow, .ColIndex("FTTaxCode"))
                    'cVat = IIf(tVB_SKUTaxCode <> "00", SP_FMTcCurPoint(SP_CALcVatValue(IIf(tVB_CNRetVATInOrEx = "1", True, False), cNet, cVB_SKUTaxRate), True, False), 0)                            '1:vat, 2:novat (0)
                    cVat = IIf(tVB_SKUTaxCode <> "00", SP_FMTcCurPoint(SP_CALcVatValue(IIf(tVB_CNRetVATInOrEx = "2", True, False), cNet, cVB_SKUTaxRate), True, False), 0)  'tVB_CNRetVATInOrEx 1:รวมใน, 2: แยกนอก แก้ไขให้ถูกต้องตามการคำนวณใน SP_CALcVatValue '*Bump 66-03-23
                    cVatable = IIf(tVB_SKUTaxCode <> "00", cNet - cVat, 0)
                    tSql = tSql & " FCSdtSalePrice=" & SP_FMTcCurPoint(cSalePrice, True, False)
                    tSql = tSql & ",FCSdtSaleUnit=" & SP_FMTcCurPoint(cSalePrice, True, False)
                    tSql = tSql & ",FCSdtB4DisChg=" & SP_FMTcCurPoint(cSalePrice, True, False)
                    tSql = tSql & ",FCSdtSaleAmt=" & cNet
                    tSql = tSql & ",FCSdtVat=" & SP_FMTcCurPoint(cVat, True, False)
                    tSql = tSql & ",FCSdtVatable=" & SP_FMTcCurPoint(cVatable, True, False)
                    tSql = tSql & ",FTSdtPriceOverideFlag='Y'"
                    tSql = tSql & ",FTEmpCodeOveride='" & tVB_CNUserAlwC & "'"
                    tSql = tSql & ",FCSdtDis = 0" '*Ao 2011/03/29 CommSheet ML-PRV-047-SF-Deposit Tran Subtotal incorrect when Price Override.xls
                    If .TextMatrix(pnRow, .ColIndex("FTPmt")) = "*" Then
                        tSql = tSql & ",FCSdtBSubPro=" & cNet
                        tSql = tSql & ",FCSdtDiscSubPro=0"
                        tSql = tSql & ",FCSdtBSubReg=0"
                        tSql = tSql & ",FCSdtDiscSubReg=0"
                        tSql = tSql & ",FCSdtYellowPrice= " & cNet
                        tSql = tSql & ",FCSdtBluePrice= 0"
                    Else
                        tSql = tSql & ",FCSdtBSubPro=0"
                        tSql = tSql & ",FCSdtDiscSubPro=0"
                        tSql = tSql & ",FCSdtBSubReg=" & cNet
                        tSql = tSql & ",FCSdtDiscSubReg=0"
                        tSql = tSql & ",FCSdtYellowPrice= 0"
                        tSql = tSql & ",FCSdtBluePrice= " & cNet
                    End If
                    tSql = tSql & ",FCSdtAmtDisAvi=" & cNet
                    tSql = tSql & ",FCSdtSalePricePmt =" & cSalePrice '*KT 52-12-28 แก้ไข DCR PRJ09010TheMall-R3-DCR_12002(Deal)-00.02.00.xls
                Case "FNSdtQty"
                    cQty = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtQty")))                              'จำนวนไม่ใช้จาก ตัวแปร cUT_Qty(เดิม), ใช้จาก grid จะครอบคลุมจำนวนแบบ เครื่องชั่ง และน้ำหนักด้วย (ใหม่)
                    cQtyAll = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtQtyAll")))
                    cSalePrice = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtSaleUnit")))  'Sale Pric
                    If (cQty < 1) Or (cQtyAll < 1) Then
                        MsgBox "Please input change qty again"                                     '*Ao 2014/02/26  CommSheet TK-ISS3000-462-Tender Cash 0.00_Aoo-570228
                        .TextMatrix(pnRow, .ColIndex("FNSdtQty")) = nW_QtyItem
                        .TextMatrix(pnRow, .ColIndex("FCSdtQty")) = nW_QtyItem
                        .TextMatrix(pnRow, .ColIndex("FCSdtQtyAll")) = nW_QtyItem
                        cQty = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtQty")))                              'จำนวนไม่ใช้จาก ตัวแปร cUT_Qty(เดิม), ใช้จาก grid จะครอบคลุมจำนวนแบบ เครื่องชั่ง และน้ำหนักด้วย (ใหม่)
                        cQtyAll = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtQtyAll")))
                        cSalePrice = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtSaleUnit")))  'Sale Pric
                        .TextMatrix(pnRow, .ColIndex("FCSdtSaleAmt")) = SP_FMTcCurPoint(nW_QtyItem * .ValueMatrix(pnRow, .ColIndex("FCSdtSaleUnit")), False, True)
                    End If
                    cNet = SP_FMTcCurPoint(cQtyAll * cSalePrice, True, False)                              'รวมเงิน
                    cVB_SKUTaxRate = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtTax")))
                    tVB_SKUTaxCode = .TextMatrix(pnRow, .ColIndex("FTTaxCode"))
                    'cVat = IIf(tVB_SKUTaxCode <> "00", SP_FMTcCurPoint(SP_CALcVatValue(IIf(tVB_CNRetVATInOrEx = "1", True, False), cNet, cVB_SKUTaxRate), True, False), 0)                            '1:vat, 2:novat (0)
                    cVat = IIf(tVB_SKUTaxCode <> "00", SP_FMTcCurPoint(SP_CALcVatValue(IIf(tVB_CNRetVATInOrEx = "2", True, False), cNet, cVB_SKUTaxRate), True, False), 0)  '*tVB_CNRetVATInOrEx 1:รวมใน, 2: แยกนอก แก้ไขให้ถูกต้องตามการคำนวณใน SP_CALcVatValue '*Bump 66-03-23
                    cVatable = IIf(tVB_SKUTaxCode <> "00", cNet - cVat, 0)
                    '*Ao 2013/01/31  TK Req
                    cRegPrice = SP_VALcText2Double(.TextMatrix(pnRow, .ColIndex("FCSdtRegPrice")))
                    cTotalBFVat = (cRegPrice * SP_FMTcCurPoint(cQty, True, False)) - (((cRegPrice * SP_FMTcCurPoint(cQty, True, False)) * cVB_SKUTaxRate) / (100 + cVB_SKUTaxRate))
                    Select Case Val(Format(tVB_TransactionSub)) '*Eaw 56-08-26 CommSheet TK-ISS3000-283 TR FAPL-63 00.00.04.xls
                        Case tEN_TrnSaleManualTax, tEN_TrnReturnManualTax
                            cVatable = cNet
                            cTotalBFVat = cNet
                    End Select
                    tSql = tSql & " FCSdtQty=" & SP_FMTcCurPoint(cQty, True, False)
                    tSql = tSql & ",FCSdtQtyAll=" & SP_FMTcCurPoint(cQtyAll, True, False)
                    tSql = tSql & ",FCSdtSaleAmt=" & cNet
                    tSql = tSql & ",FCSdtVat=" & SP_FMTcCurPoint(cVat, True, False)
                    tSql = tSql & ",FCSdtVatable=" & SP_FMTcCurPoint(cVatable, True, False)
                    tSql = tSql & ",FCSdtTotalBFVat = " & SP_FMTcCurPoint(cTotalBFVat, True, False)         '*Ao 2013/01/31  TK Req
                    If .TextMatrix(pnRow, .ColIndex("FTPmt")) = "*" Then
                        tSql = tSql & ",FCSdtBSubPro= " & cNet
                        tSql = tSql & ",FCSdtDiscSubPro=0"
                        tSql = tSql & ",FCSdtBSubReg=0"
                        tSql = tSql & ",FCSdtDiscSubReg=0"
                        tSql = tSql & ",FCSdtYellowPrice= " & cNet
                        tSql = tSql & ",FCSdtBluePrice= 0"
                    Else
                        tSql = tSql & ",FCSdtBSubPro=0"
                        tSql = tSql & ",FCSdtDiscSubPro=0"
                        tSql = tSql & ",FCSdtBSubReg= " & cNet
                        tSql = tSql & ",FCSdtDiscSubReg=0"
                        tSql = tSql & ",FCSdtYellowPrice= 0"
                        tSql = tSql & ",FCSdtBluePrice= " & cNet
                    End If
                    tSql = tSql & ",FCSdtAmtDisAvi= " & cNet
            End Select
             tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
             tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
             If .TextMatrix(pnRow, .ColIndex("FTSdtWeight")) = "" Then
                 tSql = tSql & " AND  FTSdtBarCode='" & .TextMatrix(pnRow, .ColIndex("FTSdtBarCode")) & "' AND FNSdtSeqNo=" & .TextMatrix(pnRow, .ColIndex("FNSdtSeqNo"))
            Else
                tSql = tSql & " AND  FTSdtInstoreBar='" & .TextMatrix(pnRow, .ColIndex("FTSdtBarCode")) & "' AND FNSdtSeqNo=" & .TextMatrix(pnRow, .ColIndex("FNSdtSeqNo"))
            End If
        Call SP_SQLvExecute(tSql, , , , , nEff)
        If nEff > 0 Then
            '*Ao 2010/12/16 Price Override
            Select Case .ColKey(pnCol)
                Case "FCSdtSaleUnit"
                     tSql = "UPDATE " & tVB_CNTblDT & " SET "
                     tSql = tSql & "FCOverrideAmt = ROUND(FCSdtSaleAmt - (FCAFGroupPM * FCSdtQtyAll),2)"
                     tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                     tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                     If .TextMatrix(pnRow, .ColIndex("FTSdtWeight")) = "" Then
                         tSql = tSql & " AND  FTSdtBarCode='" & .TextMatrix(pnRow, .ColIndex("FTSdtBarCode")) & "' AND FNSdtSeqNo=" & .TextMatrix(pnRow, .ColIndex("FNSdtSeqNo"))
                    Else
                        tSql = tSql & " AND  FTSdtInstoreBar='" & .TextMatrix(pnRow, .ColIndex("FTSdtBarCode")) & "' AND FNSdtSeqNo=" & .TextMatrix(pnRow, .ColIndex("FNSdtSeqNo"))
                    End If
                    Call SP_SQLvExecute(tSql, , , , , nEff)
                    If nEff > 0 Then
                        tSql = "UPDATE " & tVB_CNTblCD & " SET "            'เก็บส่วนลดเดิมที่เคยได้ไว้เป็นประวัติ
                        tSql = tSql & "FTSdtStkCode = '1'"
                        tSql = tSql & " WHERE FNSdtSeqNo =" & .TextMatrix(pnRow, .ColIndex("FNSdtSeqNo"))
                        tSql = tSql & " AND FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tSql = tSql & " AND FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        Call SP_SQLvExecute(tSql, , , , , nEff)
                        If nEff = 0 Then
'                            Call SP_MSGxShowError(Err.Number, Err.Description)
                        End If
                    End If
                Case "FNSdtQty"
                    tSql = "UPDATE " & tVB_CNTblDT & " SET "
                    tSql = tSql & "FCOverrideAmt = ROUND(FCSdtSaleAmt - (FCAFGroupPM * FCSdtQtyAll),2)"
                    tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                    tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                     If .TextMatrix(pnRow, .ColIndex("FTSdtWeight")) = "" Then
                         tSql = tSql & " AND  FTSdtBarCode='" & .TextMatrix(pnRow, .ColIndex("FTSdtBarCode")) & "' AND FNSdtSeqNo=" & .TextMatrix(pnRow, .ColIndex("FNSdtSeqNo"))
                    Else
                        tSql = tSql & " AND  FTSdtInstoreBar='" & .TextMatrix(pnRow, .ColIndex("FTSdtBarCode")) & "' AND FNSdtSeqNo=" & .TextMatrix(pnRow, .ColIndex("FNSdtSeqNo"))
                    End If
                    tSql = tSql & " AND (FCOverrideAmt <> 0) "
                    Call SP_SQLvExecute(tSql, , , , , nEff)
            End Select
            '_____________________________________________________________
        End If
        
        End With
        Exit Sub
ErrHandle:
        Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Public Sub W_ClearValAfterInsTSDT()
'----------------------------------------------------------
'   Call:
'   Cmt:    clear values after insert data into DT
'----------------------------------------------------------
    '*NEW
    tUT_CmpTaxCode = ""
    tVB_CNDpcCode = ""
    tVB_CNSkuDepCode = ""
    tVB_CNSkuBrdCode = ""
    tUT_SKUUnitC = ""
    tUT_SKUUnitN = ""
    tUT_SKUCode = vbNullString
    cUT_SalePrice = 0
    cUT_SKUCost = 0
    cUT_ChgAmt = 0
    cUT_DisAmt = 0
    cUT_StkFac = 1
    bVB_SKUisPromotion = False
    '*KT 52-02-13 Member price
    tVB_StaMemPriceAlw = "N"
    cUT_MemberPrice = 0
    tVB_CNEventMemPrice = ""
    nVB_CNMemPriceMaxQty = 0
    tVB_CNMemPriceType = "" '*KT 52-09-07 MemberPrice + Consignment
    cVB_CNMemPriceAmt = 0 '*KT 52-09-07 MemberPrice + Consignment
    
     tVB_SKUConsignFlag = "N" 'Flag Consignment ของ SKU '*KT 52-08-06
     
     cUT_Qty = 1  '*KT 52-08-13 PH2-ML-SF- 110 VIP Create เมื่อ ถ้า Cancel รายการสินค้า เมื่อทำ VIP Sale แล้วขายสินค้าเพิ่ม จำนวนสินค้า และ จำนวนเงินจะติดลบ
     
     cVB_CNRetSaleAmt = 0  'ตัวแปรเก็บราคารวมของสินค้าก่อนลด *KT 53-01-26 Communication Sheet PH3-ML-SB-073 Deal & Free  in Return Transaction.xls
     
    If bVB_CNTrStaRegPreviewDay Then
        Call SP_INIxSwitchModePreviewDay(True)
    Else
        Call SP_INIxSwitchModePreviewDay(False)
    End If
    tVB_CNSKUFlagNotAlwPrvUse = "N"
    tVB_CNSKUFlagAlwPrvUsePromo = "N"
    cVB_AFGroupPM = 0               '*Ao 2010/12/16 Price Override
    cVB_OverrideAmt = 0
    cUT_SKUUnitF = 1                   '*Ao 2010/12/24    เปิดโปรแกรมมาที่หน้า Item เลย แล้วทำ Return จะทำให้ค่า cUT_SKUUnitF ยังไม่ได้ Set = 1
    
    '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
    tVB_CNSkuPointCooAlw = "Y"
    tVB_CNSkuPointAlw = "N"
    tVB_CNSkuEventPoint = ""
    cVB_CNSkuPointAmt = 0
    cVB_CNSkuPointTime = 0
    tVB_CNSkuPointAlwDiscountMall = "N"
    tVB_CNSKUStaFixPointFlag = "N"
    tVB_PdtNoDis = ""                                       '*Ao 2013/03/27  TK
    tVB_SdtTaxInv = "N" '*Bump 58-01-09 ใช้เก็บ Flag 'Y' ถ้า สินค้ามี Line Discount สำหรับการทำ Return Refer Sale
End Sub

Public Sub W_SumAmount(Optional ByVal pbCal As Boolean = True, Optional ByVal pbUpd As Boolean = True)
'----------------------------------------------------------
'   Call:
'   Cmt:   sum total of sale
'  ****************************************************************************
'   คำเตือน เนื่องจาก W_SumAmountEx กับ W_SumAmount มีการทำงานเหมือนกันมาก แต่การใช้งานของ W_SumAmountEx ใช้สำหรับกรณีที่ผ่านการกด Home จากหน้า Item Entry
'   แต่ W_SumAmount จะใช้สำหรับก่อนกด Home จากหน้า Item Entry
' *****************************************************************************
'----------------------------------------------------------
    Dim i As Long
    Dim tSql$, tWhr$, tValue$
    Dim nMRow As Long
    Dim cRegPrice As Double
    Dim cSaleAmt As Double
    Dim cQty As Double
    Dim cDisc As Double
    Dim cVat As Double
    Dim cVatable  As Double
'    Dim nQty As Long
    Dim tTaxCode As String
    Dim cTaxRate As Double
    Dim tOveride As String '*KT 51-10-09 V 3.0.131/V 3.0.200
    Dim cSalePrice As Double '*KT 51-10-09 V 3.0.131/V 3.0.200
    Dim cB4Disc As Double
    Dim tStaRfd As String
On Error GoTo ErrHandle
    cUT_Total = 0
    cUT_SubTotal = 0
    cUT_SubReg = 0
    cUT_SubPro = 0
    nUT_SumQty = 0
    cUT_SumNoVat = 0
    cUT_SumVat = 0
    cUT_VatAmt = 0
    cUT_VatableAmt = 0
    cVB_TotalDisc = 0
    cB4Disc = 0
    
    With ogdForm
        nMRow = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FNSdtSeqNo"), .rows - 1, .ColIndex("FNSdtSeqNo"))
        For i = 1 To nMRow
            If .TextMatrix(i, .ColIndex("FTSdtCancelFlag")) <> "1" Then
                tOveride = .TextMatrix(i, .ColIndex("FTSdtPriceOverideFlag")) '*KT 51-10-09 V 3.0.131/V 3.0.200
                cRegPrice = .ValueMatrix(i, .ColIndex("FCSdtRegPrice"))
                cSalePrice = .ValueMatrix(i, .ColIndex("FCSdtSaleUnit")) '*KT 51-10-09 V 3.0.131/V 3.0.200
                cSaleAmt = .ValueMatrix(i, .ColIndex("FCSdtSaleAmt"))
                cQty = .ValueMatrix(i, .ColIndex("FCSdtQty"))
                cDisc = .ValueMatrix(i, .ColIndex("FCSdtDis"))
'                nQty = .ValueMatrix(i, .ColIndex("FNSdtQty"))
                cVat = .ValueMatrix(i, .ColIndex("FCSdtVat"))
                cVatable = .ValueMatrix(i, .ColIndex("FCSdtVatable"))
                tTaxCode = .TextMatrix(i, .ColIndex("FTTaxCode"))
                cTaxRate = SP_VALcText2Double(.TextMatrix(i, .ColIndex("FCSdtTax")))
                tStaRfd = .TextMatrix(i, .ColIndex("FTSdtStaRfd"))
                '*KT 51-10-09 V 3.0.131/V 3.0.200
                '--------------------------------------------------------------------------------------
                If UCase(tOveride) = "Y" Then
                    Select Case Val(Format(tVB_TransactionSub))
                        Case tEN_TrnTaxExceptSale
                            cRegPrice = (cVat + cVatable + SP_FMTcCurPoint(SP_VALcText2Double((cDisc / cQty)), True, False) / cQty) '***KT 51-11-24 V 1.0.1
                        Case Else
                            cRegPrice = cSalePrice + SP_FMTcCurPoint(SP_VALcText2Double((cDisc / cQty)), True, False) '***KT 51-11-24 V 1.0.1
                    End Select
                Else
                    '*KT 52-04-08 Member Price
                    If UCase(tStaRfd) = "Y" Then
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnTaxExceptSale
                                cRegPrice = (cVat + cVatable + SP_FMTcCurPoint(SP_VALcText2Double((cDisc / cQty)), True, False) / cQty) '***KT 51-11-24 V 1.0.1
                            Case Else
                                cRegPrice = cSalePrice + SP_FMTcCurPoint(SP_VALcText2Double((cDisc / cQty)), True, False) '***KT 51-11-24 V 1.0.1
                        End Select
                    Else
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnTaxExceptSale
                                cRegPrice = (cVat + cVatable + SP_FMTcCurPoint(SP_VALcText2Double((cDisc / cQty)), True, False) / cQty) '***KT 51-11-24 V 1.0.1
                        End Select
                    End If
                End If
                '-------------------------------------------------------------------------------------
                
                cUT_Total = cUT_Total + (cQty * cRegPrice) 'ก่อนลด
                cUT_SubTotal = cUT_SubTotal + cSaleAmt  'หลักลด
                
                If Trim(.TextMatrix(i, .ColIndex("FTPmt"))) <> "*" Then
                    cUT_SubReg = cUT_SubReg + cSaleAmt
                Else
                    cUT_SubPro = cUT_SubPro + cSaleAmt
                End If

                nUT_SumQty = nUT_SumQty + cQty
                cVB_TotalDisc = cVB_TotalDisc + cDisc
                '*KT 51-05-26 Comsheet ML-SF-291
                 Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnTaxExceptSale
                        If Trim(tTaxCode) = "00" Then
                             cUT_SumNoVat = cUT_SumNoVat + SP_FMTcCurPoint(((cRegPrice * cQty) - ((cRegPrice * cQty) * cTaxRate / (100 + cTaxRate))), True, False)  '*KT 52-02-28  cRegPrice ไม่ต้องคูณ cQty กรณี TaxExceptSale
                        Else
                            cUT_SumVat = cUT_SumVat + SP_FMTcCurPoint(((cRegPrice * cQty) - ((cRegPrice * cQty) * cTaxRate / (100 + cTaxRate))), True, False) '*KT 52-02-28  cRegPrice ไม่ต้องคูณ cQty กรณี TaxExceptSale
                        End If
                    Case Else
                         If Trim(tTaxCode) = "00" Then
                            cUT_SumNoVat = cUT_SumNoVat + (cRegPrice * cQty)
                        Else
                            cUT_SumVat = cUT_SumVat + (cRegPrice * cQty)
                        End If
                  End Select
                  cUT_VatAmt = cUT_VatAmt + cVat
                  cUT_VatableAmt = cUT_VatableAmt + cVatable
            End If
        Next i
    End With
    
    '*TW 51-08-08
    If pbCal Then Call oDisc.DISxGetSubRegPro(cUT_SubReg, cUT_SubPro)
    cUT_SubTotal = cUT_SubReg + cUT_SubPro
    olaTSubReg.Caption = SP_FMTcCurPoint(cUT_SubReg, False, True)
    olaTSubPro.Caption = SP_FMTcCurPoint(cUT_SubPro, False, True)
    olaTSubTotal.Caption = SP_FMTcCurPoint(cUT_SubTotal, False, True)
    
    '*AO51-03-20        Comm.Sheet  SF-111
    olaSumQty.Caption = nUT_SumQty
'    If Not (GetInputState = 0) Then DoEvents '*KT 53-01-14 DoEvents
    
    'โยนค่าใส่ตัวแปรต่างๆๆ ในการชำระเงิน Tender
    cVB_TotalDisc = cUT_Total - cUT_SubTotal
    
    'ทำการปรับแปรุงยอดของ Sub Reg และ Sub Pro

    tValue = " SET FCShdSubReg = " & SP_FMTcCurPoint(cUT_SubReg, True, False)
    tValue = tValue & vbCrLf & " ,FCShdSubPro = " & SP_FMTcCurPoint(cUT_SubPro, True, False)
    tValue = tValue & vbCrLf & " ,FCShdNonVat = " & SP_FMTcCurPoint(cUT_SumNoVat, True, False)
    tValue = tValue & vbCrLf & " ,FCShdUnNontVat = " & SP_FMTcCurPoint(cUT_SumVat, True, False)
    tValue = tValue & vbCrLf & " ,FCShdVat = " & SP_FMTcCurPoint(cUT_VatAmt, True, False)
    tValue = tValue & vbCrLf & " ,FCShdVatable = " & SP_FMTcCurPoint(cUT_VatableAmt, True, False)
    tValue = tValue & vbCrLf & " ,FCShdB4DisChg = " & SP_FMTcCurPoint(cUT_Total, True, False)
    tValue = tValue & vbCrLf & " ,FCShdTotal= " & SP_FMTcCurPoint(cUT_SubTotal, True, False)
    
    tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
    tWhr = tWhr & vbCrLf & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
    tSql = "STP_CNxSQLFunctionUpd @ptValue='" & tValue & "',@ptTable='" & tVB_CNTblHD & "',@ptWhere='" & tWhr & "'"
    If pbUpd Then
        Call SP_SQLbExecute(tSql)
    End If
    Exit Sub
ErrHandle:
    
End Sub

Public Sub W_SumAmountEx(Optional ByVal pbCal As Boolean = True)
'----------------------------------------------------------
'   Call:
'   Cmt:   sum total of sale
'  ****************************************************************************
'   คำเตือน เนื่องจาก W_SumAmountEx กับ W_SumAmount มีการทำงานเหมือนกันมาก แต่การใช้งานของ W_SumAmountEx ใช้สำหรับกรณีที่ผ่านการกด Home จากหน้า Item Entry
'   แต่ W_SumAmount จะใช้สำหรับก่อนกด Home จากหน้า Item Entry
' *****************************************************************************
'----------------------------------------------------------
    Dim i As Long
    Dim tSql$, tWhr$, tValue$
    Dim orsTemp As ADODB.Recordset
    Dim cRegPrice As Double
    Dim cQty As Double
    Dim cSaleAmt As Double
    Dim tCancelFlag As String
    Dim cDis As Double
    Dim cVat As Double
    Dim cVatable  As Double
    Dim tTax As String
    Dim tPmt As String
    Dim cTaxRate As Double
    Dim tOveride As String '*KT 51-10-09 V 3.0.131/V 3.0.200
    Dim cSalePrice As Double '*KT 51-10-09 V 3.0.131/V 3.0.200
    Dim cB4Disc As Double
    Dim tStaRfd As String
On Error GoTo ErrHandle
    cUT_Total = 0
    cUT_SubTotal = 0
    cUT_SubReg = 0
    cUT_SubPro = 0
    nUT_SumQty = 0
    cUT_SumNoVat = 0
    cUT_SumVat = 0
    cUT_VatAmt = 0
    cUT_VatableAmt = 0
    cVB_TotalDisc = 0

    tSql = "SELECT FCSdtTax,FTSdtCancelFlag,FCSdtRegPrice,FCSdtSaleAmt,FCSdtQty,FCSdtDis"
    tSql = tSql & vbCrLf & ",FTTaxCode,FCSdtVat,FCSdtVatable,FTSkuStaMallEmpCard"
    tSql = tSql & vbCrLf & ",FCSdtBSubReg,FCSdtBSubPro,FTSdtPriceOverideFlag,FCSdtSaleUnit,FTSdtStaRfd" '*KT 51-10-09 V 3.0.131/V 3.0.200
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "ORDER BY FNSdtSeqNo"
    Call SP_SQLvExecute(tSql, orsTemp)
    With ogdForm
        Do While Not orsTemp.EOF
             tOveride = SP_FEDtChkString(orsTemp, "FTSdtPriceOverideFlag") '*KT 51-10-09 V 3.0.131/V 3.0.200
            cRegPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")
            cSalePrice = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleUnit") '*KT 51-10-09 V 3.0.131/V 3.0.200
            cSaleAmt = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt")
            cQty = SP_FEDvChkNumeric(orsTemp, "FCSdtQty")
            tCancelFlag = SP_FEDtChkString(orsTemp, "FTSdtCancelFlag")
            cDis = SP_FEDvChkNumeric(orsTemp, "FCSdtDis") '***KT 51-11-24
            cVat = SP_FEDvChkNumeric(orsTemp, "FCSdtVat")
            cVatable = SP_FEDvChkNumeric(orsTemp, "FCSdtVatable")
            tTax = SP_FEDtChkString(orsTemp, "FTTaxCode")
            tPmt = IIf(SP_FEDtChkString(orsTemp, "FTSkuStaMallEmpCard") = "N", "*", "")
            cTaxRate = SP_FEDvChkNumeric(orsTemp, "FCSdtTax")
             tStaRfd = SP_FEDtChkString(orsTemp, "FTSdtStaRfd")
            '*KT 51-10-09 V 3.0.131/V 3.0.200
            '--------------------------------------------------------------------------------------
            If UCase(tOveride) = "Y" Then
                    Select Case Val(Format(tVB_TransactionSub))
                        Case tEN_TrnTaxExceptSale
                            cRegPrice = SP_FMTcCurPoint(SP_VALcText2Double((cVat + cVatable + (cDis / cQty)) / cQty), True, False) '***KT 51-11-24 V 1.0.1 การ + Discount ต้องเป็น Discount ที่เป็น Discount ต่อหน่วย
                        Case Else
                            cRegPrice = SP_FMTcCurPoint(SP_VALcText2Double(cSalePrice + (cDis / cQty)), True, False) '***KT 51-11-24 V 1.0.1 การ + Discount ต้องเป็น Discount ที่เป็น Discount ต่อหน่วย
                    End Select
            Else
                    '*KT 52-04-08 Member Price
                    If UCase(tStaRfd) = "Y" Then
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnTaxExceptSale
                                cRegPrice = SP_FMTcCurPoint(SP_VALcText2Double((cVat + cVatable + (cDis / cQty)) / cQty), True, False) '***KT 51-11-24 V 1.0.1 การ + Discount ต้องเป็น Discount ที่เป็น Discount ต่อหน่วย
                            Case Else
                               cRegPrice = cSalePrice + SP_FMTcCurPoint(SP_VALcText2Double((cDis / cQty)), True, False) '***KT 51-11-24 V 1.0.1
                        End Select
                    Else
                        Select Case Val(Format(tVB_TransactionSub))
                            Case tEN_TrnTaxExceptSale
                               cRegPrice = SP_FMTcCurPoint(SP_VALcText2Double((cVat + cVatable + (cDis / cQty)) / cQty), True, False)
                        End Select
                    End If
            End If
            '-------------------------------------------------------------------------------------
            
            If tCancelFlag <> "1" Then
                cUT_Total = cUT_Total + (cQty * cRegPrice)
                cUT_SubPro = cUT_SubPro + SP_FEDvChkNumeric(orsTemp, "FCSdtBSubPro")
                cUT_SubReg = cUT_SubReg + SP_FEDvChkNumeric(orsTemp, "FCSdtBSubReg")
                
                nUT_SumQty = nUT_SumQty + cQty
                cVB_TotalDisc = cVB_TotalDisc + cDis
                Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnTaxExceptSale
                        If Trim(tTax) = "00" Then
                            cUT_SumNoVat = cUT_SumNoVat + (cRegPrice * cQty) 'cUT_SumNoVat + SP_FMTcCurPoint(((cRegPrice * cQty) - ((cRegPrice * cQty) * cTaxRate / (100 + cTaxRate))), True, False)
                        Else
                            cUT_SumVat = cUT_SumVat + (cRegPrice * cQty)  'cUT_SumVat + SP_FMTcCurPoint(((cRegPrice * cQty) - ((cRegPrice * cQty) * cTaxRate / (100 + cTaxRate))), True, False)
                        End If
                    Case Else
                         If Trim(tTax) = "00" Then
                            cUT_SumNoVat = cUT_SumNoVat + (cRegPrice * cQty)
                        Else
                            cUT_SumVat = cUT_SumVat + (cRegPrice * cQty)
                        End If
                End Select
                cUT_VatAmt = cUT_VatAmt + cVat
                cUT_VatableAmt = cUT_VatableAmt + cVatable
            End If
            orsTemp.MoveNext
        Loop
    End With
    
    '*TW 51-08-08
     If pbCal Then Call oDisc.DISxGetSubRegPro(cUT_SubReg, cUT_SubPro)
    cUT_SubTotal = cUT_SubReg + cUT_SubPro
    cVB_TotalDisc = cUT_Total - cUT_SubTotal
    olaTSubReg.Caption = SP_FMTcCurPoint(cUT_SubReg, False, True)
    olaTSubPro.Caption = SP_FMTcCurPoint(cUT_SubPro, False, True)
    olaTSubTotal.Caption = SP_FMTcCurPoint(cUT_SubTotal, False, True)
    
    'ทำการปรับแปรุงยอดของ Sub Reg และ Sub Pro
    tValue = " SET FCShdSubReg = " & SP_FMTcCurPoint(cUT_SubReg, True, False)
    tValue = tValue & vbCrLf & " ,FCShdSubPro = " & SP_FMTcCurPoint(cUT_SubPro, True, False)
    tValue = tValue & vbCrLf & " ,FCShdNonVat = " & SP_FMTcCurPoint(cUT_SumNoVat, True, False)
    tValue = tValue & vbCrLf & " ,FCShdUnNontVat = " & SP_FMTcCurPoint(cUT_SumVat, True, False)
    tValue = tValue & vbCrLf & " ,FCShdVat = " & SP_FMTcCurPoint(cUT_VatAmt, True, False)
    tValue = tValue & vbCrLf & " ,FCShdVatable = " & SP_FMTcCurPoint(cUT_VatableAmt, True, False)
    tValue = tValue & vbCrLf & " ,FCShdB4DisChg = " & SP_FMTcCurPoint(cUT_Total, True, False)
    tValue = tValue & vbCrLf & " ,FCShdTotal= " & SP_FMTcCurPoint(cUT_SubTotal, True, False)
    
    tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
    tWhr = tWhr & vbCrLf & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"

    tSql = "STP_CNxSQLFunctionUpd @ptValue='" & tValue & "',@ptTable='" & tVB_CNTblHD & "',@ptWhere='" & tWhr & "'"
    Call SP_SQLbExecute(tSql)
    Exit Sub
ErrHandle:
    
End Sub

Public Sub W_SetGridHeading()
'----------------------------------------------------------
'   Call:
'   Cmt:   set heading of selling grid
'----------------------------------------------------------
    Dim tColW$, iCol%
On Error GoTo ErrHandle
    'tColW = "705;195;3075;3405;705;1350;1350;1560"      '<= tFed, ส่วน tFed จะมีเท่าไหร่ก็ได้ แต่ห้ามเกินจำนวน cols ของ grid
    '*Ao 52-04-04
    If bVB_Res640X480 Then
        tColW = "540;195;195;1755;2700;600;1215;1215;1305"              'Resorution 640*480
    Else
        tColW = "605;195;195;2595;3305;805;1510;1510;1720"              'Resorution 800 * 600
    End If

    iCol = 9     'last col to show, if > iCol จะถูก hidden
    With ogdForm
        Call SP_GRDxSetProperty(ogdForm, tW_Fed, tColW, iCol, True, 0)
          '*KT 55-08-17 Takashimaya PRJ10001TK-External Specification SKU-00.01.00.doc
         .ColHidden(1) = True
        'Show
        .TextMatrix(.FixedRows - 1, .ColIndex("FNSdtSeqNo")) = UT_tToken(tCS_ColNo, ";", UT_nPosCutTag)                         'No
        .TextMatrix(.FixedRows - 1, .ColIndex("FTPmt")) = ""
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSymbol")) = ""
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSdtBarCode")) = UT_tToken(tCS_ColBarCode, ";", UT_nPosCutTag)      'SKU Name
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSkuAbbName")) = UT_tToken(tCS_ColProductName, ";", UT_nPosCutTag)      'Product Name
        .TextMatrix(.FixedRows - 1, .ColIndex("FNSdtQty")) = UT_tToken(tCS_ColQty, ";", UT_nPosCutTag)                                'Qty
        .TextMatrix(.FixedRows - 1, .ColIndex("FCSdtRegPrice")) = UT_tToken(tCS_ColRegPrice, ";", UT_nPosCutTag)                  'Reg. Price
        .TextMatrix(.FixedRows - 1, .ColIndex("FCSdtSaleUnit")) = UT_tToken(tCS_ColSalePrice, ";", UT_nPosCutTag)                        'Sale Price
        .TextMatrix(.FixedRows - 1, .ColIndex("FCSdtSaleAmt")) = UT_tToken(tCS_ColSubTotal, ";", UT_nPosCutTag)                 'Sub Total
        .ColAlignment(.ColIndex("FNSdtSeqNo")) = flexAlignCenterCenter        'special
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 11
    End With
    Exit Sub
ErrHandle:
    
End Sub

Public Sub W_UPDxTransHD()
'------------------------------------------------------
'Call:
'Cmt:   *KT 50-12-24 ปรับปรุงข้อมูล Transaction HD
'-----------------------------------------------------
    Dim tSql$, tValue$, tWhr$
On Error GoTo ErrHandle
    tValue = " SET FCShdTotal =" & SP_FMTcCurPoint(cUT_SubTotal, True, False)
    tValue = tValue & vbCrLf & " , FCShdGrand=" & SP_FMTcCurPoint(cUT_SubTotal, True, False)
    If xVB_CNMemberP.ID <> "" Then  '*Eaw 56-02-06  การ Recovery ข้อมูล
        tValue = tValue & vbCrLf & ",FTShdCpnCodeRef = ''" & xVB_CNMemberP.ID & "''"
    End If
    tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "''"
    tWhr = tWhr & vbCrLf & "  AND FTShdTransNo=''" & tVB_Running & "''"
    tWhr = tWhr & vbCrLf & "  AND FTShdTransType=''" & tVB_Transaction & "''"
    tWhr = tWhr & vbCrLf & "  AND FDShdTransDate =''" & tUT_SaleDate & "''"
    tWhr = tWhr & vbCrLf & "  AND FTShdStaDoc = ''2''"

    tSql = "STP_CNxSQLFunctionUpd @ptValue='" & tValue & "',@ptTable='" & tVB_CNTblHD & "',@ptWhere='" & tWhr & "'"
    Call SP_SQLbExecute(tSql)
    Exit Sub
ErrHandle:

End Sub

Public Sub W_FRMxInitailTrans(Optional ByVal pbPreview As Boolean = False)
''----------------------------------------------------
'Call:
'Cmt: ทำการ Clear หน้าจอเพื่อเริ่มต้นทำ Transaction ใหม่
'----------------------------------------------------
On Error GoTo ErrHandle
    'Bar
    olaDocNo.Caption = tVB_Running
    '*BG 56-09-24 Comsheet - 356 Print EJ ปรับค่าให้แสดง Virtual POS หน้าขาย
'    olaPos.Caption = tVB_CNTerminalNum
    If bVB_ReEntryMode Then
        olaPos.Caption = tVB_CNVirTerminalNum
    Else
        olaPos.Caption = tVB_CNTerminalNum
    End If
    '****
    olaTransName.Caption = IIf(bVB_TrainingMode, "Tr\" & tVB_TransName, IIf(bVB_ReEntryMode, "Re Entry\" & tVB_TransName, tVB_TransName))
    'Header
    olaBarcode.Caption = ""
    olaSKUName.Caption = ""
    '*KT 54-09-20 V.4.0.1 กรณีที่ไม่เป็น OPOS ในการแสดงเงินทอน
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case 1, 5, 8, 9, 14, 17, 25, 26, 27, 28, 29, 30, 31, 32, 34, 35, 36
                olaPrice.Caption = ""
        Case Else
                If bVB_OPOSCashDrawerEnable Or (Not olaChange.Visible) Then
                    olaPrice.Caption = ""
                Else  'Eaw 56-11-29 เพิ่มส่วนเคส else ทั้งหมด
'                    If SP_VALcText2Double(olaPrice.Caption) = 0 Then
'                        olaChange.Visible = False
'                        olaPrice.Caption = ""
'                    End If
                    olaChange.Visible = False
                    olaPrice.Caption = ""
                End If
    End Select
    '*BG 57-04-28 PRJ10001TK-PH2-Internal Spec-Goods Return on POS-00.01.00.docx
    '            4-Return
'            16-Layaway Cancel
'            20-Add Deposit Cancel
'            39-Return(Tax Include)
'            44-Return(Manual Tax)
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case 4, 16, 20, 39, 44
            olaTransName.BackStyle = 1
'            olaTransName.BackColor = &HFF00&  'สีเขียว
'            olaTransName.BackColor = &HC0E0FF      'สีส้มอ่อน
'            olaTransName.BackColor = &H80C0FF       'สีส้มกลาง
            olaTransName.BackColor = &H80FF&    'สีส้มเข้ม
        Case Else
            olaTransName.BackStyle = 0
    End Select
    '-----------------------------------------------------------------------------------------------------------------------------
    'Footer
    olaSumQty.Caption = ""
    olaTSubReg.Caption = ""
    olaTSubPro.Caption = ""
    olaTSubTotal.Caption = ""
    olaRef.Caption = ""
    otbPrompt.Text = ""
    'otbPrompt.Locked = False
    cUT_Qty = 1
    bVB_Recall = False
    bVB_PriceOveride = False
    '*Thirapong 2020-10-12 ST on reprint issues.xlsx
    tVB_MemPointID = ""
    tVB_MallCardIDRef = ""
    '********************************
    'Prevew Day
    If Not pbPreview Then
        bVB_CNTrStaRegPreviewDay = False
        Call SP_INIxSwitchModePreviewDay(False)
        Call SP_INIxClearVarPrv
    End If
    Call W_SETxIndicator(tVB_TransactionSub, False)       '*Ao 2014/06/21
    
    '*Bump 56-03-08 TK-ISS3000-033 ตรวจสอบพื้นที่ HHD
    Call W_CHKxHHDSpace
    Exit Sub
ErrHandle:
    
End Sub

Private Sub otbPrompt_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle
    Call UT_ChkKeyAscii(KeyAscii)
    If bW_LockKeyAscii Then KeyAscii = 0
    If KeyAscii = 22 Then                                               '*Ao 2011/03/22       Ctrl+V = KeyAscii = 22
        Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:otbPrompt_KeyPress[1]")
        If InStr(1, Clipboard.GetText, "'") > 0 Then
            otbPrompt.Text = Replace(Clipboard.GetText, "'", "")
            If otbPrompt.Text <> "" Then
                otbPrompt.SelStart = Len(otbPrompt.Text)
                otbPrompt.SelLength = Len(otbPrompt.Text)
            End If
            KeyAscii = 0
        End If
    End If
  Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:otbPrompt_KeyPress[2]")
End Sub

Private Sub otmDisplay_Timer()
'--------------------------------------------------------------
'Cmt: Timer สำหรับการแสดงข้อความจาก server
'           *KT 52-10-27 ปรับปรุง
'-------------------------------------------------------------
    Dim tFile$
    Dim oFSO As New FileSystemObject
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
On Error GoTo ErrHandle

    '*KT 52-12-19 POSFront 2.0.30 แก้ไขให้กาทำงานของ Timer ของการ Display ถูกต้องมากขึ้น
    If DateDiff("n", Format(dW_TimeDisplay, "yyyy/MM/dd HH:mm:ss"), Format(Now, "yyyy/MM/dd HH:mm:ss")) >= nVB_TimeDisplayMsg Then
    
        '*Tao 61-09-24 นับเวลาเริ่ม SentSale
        If (Screen.ActiveForm.Name = wMain.Name) And (oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP)) Then
            nW_CntSentSale = nW_CntSentSale + 1
            If bUT_FuncLockKB Then
                If nW_CntSentSale >= (CInt(60 / (otmDisplay.Interval / 1000)) * 2.5) Then
                    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1") 'ถ้าอยู่ในโหมด LockKB ให้ส่งทุกๆ 5 นาที
                    nW_CntSentSale = 0
                End If
            Else
                If nW_CntSentSale >= (CInt(60 / (otmDisplay.Interval / 1000)) * 5) Then
                    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1") 'ถ้าอยู่ในโหมดปรกต แต่ไม่ได้กดปุ่มใดๆ ให้ส่งทุกๆ 5 นาที
                    nW_CntSentSale = 0
                End If
            End If
        End If
        '===========================
        
        Call W_FRMxShwMsg
    End If
    
    tSql = "select  name from Sys.sysobjects where name = '" & "TRG_Tmp2Sale" & tVB_CNTerminalNum & "'"
    Call SP_SQLvExecute(tSql, orsTemp)
    If orsTemp.EOF And orsTemp.EOF Then
        '*KT 52-01-14 กรณีตรวจสอบเจอว่าไม่มี Trigger แล้วยังสร้างไม่สำเร็จทั้งฐานข้อมูล server และ local
        If Not SP_DATbCrtTrigger(False, True) Then
              Call SP_MSGnShowing("tms_CN005168", nCS_Warning)
              Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Tricger" & " Form:" & Me.Name & " Step: otmDisplay_Timer" & " Events:SP_DATbCrtTrigger")
        End If
    End If
    
    If bUT_FuncLockKB Then
        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
            oimOnOff.Picture = oilOnOff.ListImages(1).Picture
            olaOnOff.Caption = "On-line"
        Else
            oimOnOff.Picture = oilOnOff.ListImages(2).Picture
            olaOnOff.Caption = "Off-line"
        End If
    End If
    
    'PH 2.0.0
    'เปิด EXE ที่ทำหน้าที่เป็น Service ในการ Replicate
    
        'เปิด EXE ที่ทำหน้าที่เป็น Service ในการเก็บค่า Offline Online
    '-------------------------------------------------------------------------------------
    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceOnOff.exe"
    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
            tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
            tFile = tFile + "\POS\ServiceOnOff.exe"
            If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
                Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Can not run ServiceOnOff.exe" & " Form:" & Me.Name & " Control:otmDisplay" & " Events:Timer")
                Call SP_APPxEndProgram
            End If
    End If
    nVB_lServiceOnOffWnd = FindWindow(vbNullString, "ServiceOnOff")
    If nVB_lServiceOnOffWnd = 0 Then
          'Shell tFile, vbHide
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "ShellEx ServiceOnOff.exe" & " Form:" & Me.Name & " Control:otmDisplay" & " Events:Timer")  '*KT 52-12-17 POSFront 2.0.30
        Call ShellEx(nVB_lServiceOnOffWnd, "open", tFile, "", "", 10)
    End If
'    If Not (GetInputState = 0) Then DoEvents '*Ao 53-01-14   DoEvents
    '-----------------------------------------------------------------------------------
    
    '-------------------------------------------------------------------------------------
    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceAutoReplicate.exe"
    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
            tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
            tFile = tFile + "\POS\ServiceAutoReplicate.exe"
            If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
                Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Can not run ServiceAutoReplicate.exe" & " Form:" & Me.Name & " Control:otmDisplay" & " Events:Timer")
                Call SP_APPxEndProgram
            End If
    End If
    nVB_lServiceAutoReplicateWnd = FindWindow(vbNullString, "ServiceAutoReplicate")
    If nVB_lServiceAutoReplicateWnd = 0 Then
          'Shell tFile, vbHide
        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "ShellEx ServiceAutoReplicate.exe" & " Form:" & Me.Name & " Control:otmDisplay" & " Events:Timer") '*KT 52-12-17 POSFront 2.0.30
        Call ShellEx(nVB_lServiceAutoReplicateWnd, "open", tFile, "", "", 10)
    End If
'    If Not (GetInputState = 0) Then DoEvents '*Ao 53-01-14   DoEvents
    '-----------------------------------------------------------------------------------
    
   'เปิด EXE ที่ทำหน้าที่เป็น ServiceTranfer
    '-------------------------------------------------------------------------------------
    tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceTranfer.exe"
    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
            tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
            tFile = tFile + "\POS\ServiceTranfer.exe"
            If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
                Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Can not run ServiceTranfer.exe" & " Form:" & Me.Name & " Control:otmDisplay" & " Events:Timer")
                Call SP_APPxEndProgram
            End If
    End If
    nVB_lServiceTransWnd = FindWindow(vbNullString, "ServiceTranfer")
    If nVB_lServiceTransWnd = 0 Then
            'Shell tFile, vbHide
            Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "ShellEx ServiceTranfer.exe" & " Form:" & Me.Name & " Control:otmDisplay" & " Events:Timer")  '*KT 52-12-17 POSFront 2.0.30
            Call ShellEx(nVB_lServiceTransWnd, "open", tFile, "", "", 10)
    End If
'    If Not (GetInputState = 0) Then DoEvents '*Ao 53-01-14   DoEvents
    dW_TimeDisplay = Now() '*KT 52-12-19 POSFront 2.0.30 แก้ไขให้กาทำงานของ Timer ของการ Display ถูกต้องมากขึ้น
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    dW_TimeDisplay = Now() '*KT 52-12-19 POSFront 2.0.30 แก้ไขให้กาทำงานของ Timer ของการ Display ถูกต้องมากขึ้น
    Call SP_DATxRsNothing(orsTemp)
    If Err.Number = -2147467259 Or Err.Number = -2147217871 Or Err.Number = 91 Or Err.Number = 3704 Then    '-2147467259, -2147217871, 91, 3704 '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
        Call SP_MSGnShowing("tms_CN005199", nCS_Warning)
        Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:otmDisplay" & " Events:Timer")
        Call SP_APPxEndProgram
    ElseIf Err.Number = 440 Then
        Call SP_MSGxShowError(Err.Number, Err.Description)
    Else
        Call SP_MSGxShowError(Err.Number, Err.Description)
    End If
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " Form:" & Err.Description & Me.Name & ",otmDisplay_Timer[ERROR]")
End Sub

Private Sub otmLock_Timer()
    '*KT 51-10-31
    '*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตังต้น 3.0.19
    Dim bOnOff As Boolean
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
    Dim tErr$
    Dim bOnOffCur As Boolean
    Dim tSta As String
    Dim tStaOk As String
    Dim oForm As Form
    Dim tWhe As String
    Dim tTrnType As String
On Error GoTo ErrHandle '*KT 52-01-09 ใส่ on error เพื่อดูว่าเกิด Error Number 91 ซึ่งส่วนใหญ่เกิดจากการลง SQLEXPRESS 2005 ไม่สมบูรณ์เมื่อมาทำงานบน Offline
    
    nW_TimerLock = nW_TimerLock - 1
    If nW_TimerLock = 0 And Screen.ActiveForm.Name = wMain.Name Then
        Call otbPrompt_KeyDown(vbKeyF1, 2) 'Lock
        nW_CntSentSale = 0 '*Tao 61-09-24 เริ่มนับใหม่
        Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1") '*Tao 61-09-24 Start Service Transfer
        Call SP_FRMxSetTimeCount(False) 'Disable Timmer(otmLock)
    End If
    If nW_TimerLock < 0 Then nW_TimerLock = nVB_CNTimeLock + 1 '*KT 52-08-17 Lock Keyboard
    
    '*KT 51-02-27 ทำการ Update Ststusbar หน้าจอหลัก
    bOnOff = IIf(olaOnOff.Caption = "On-line", True, False)
    bOnOffCur = oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP)
    If bOnOffCur Then
        oimOnOff.Picture = oilOnOff.ListImages(1).Picture
        olaOnOff.Caption = "On-line"
        'move file ej offline to path online
        If Not bOnOff And bOnOffCur Then
            Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")
            Call SP_MOVxEJToOnline
            Call SP_CHKbOnlineEnable 'PH 2.0.0
            Call SP_TMNxUpdateOpenOnline
        End If
    Else
        oimOnOff.Picture = oilOnOff.ListImages(2).Picture
        olaOnOff.Caption = "Off-line"
    End If
     tSta = GetSetting("AdaPosMall", "POSFront", "RepSta")
     If tSta = "0" Then
        wMain.ostForm.Panels(9).Picture = Nothing
        wMain.ostForm.Panels(9).Text = "REP"
     Else
        wMain.ostForm.Panels(9).Text = ""
        tStaOk = GetSetting("AdaPosMall", "POSFront", "RepOk")
        If tStaOk = "0" Then
             wMain.ostForm.Panels(9).Picture = wMain.oilOnOff.ListImages(5).Picture
        Else
             wMain.ostForm.Panels(9).Picture = wMain.oilOnOff.ListImages(4).Picture
        End If
     End If
     
'    If bOnOffCur And oTerminal.SaleMode = nEN_OnlineNormal Then
'         ostForm.Panels(1).Text = "SRV"
'     Else
'         ostForm.Panels(1).Text = "LOCAL"
'    End If
    
    tSta = GetSetting("AdaPosMall", "POSFront", "SentSale")
    If tSta = "1" Or Not bOnOffCur Then
        If bUT_FuncLockKB And bOnOffCur Then
             wMain.ostForm.Panels(10).Picture = wMain.oilOnOff.ListImages(6).Picture
        Else
            wMain.ostForm.Panels(10).Picture = wMain.oilOnOff.ListImages(7).Picture
        End If
    Else
        wMain.ostForm.Panels(10).Picture = wMain.oilOnOff.ListImages(6).Picture
    End If
    
'    If Not (GetInputState = 0) Then DoEvents '*Ao 53-01-14   DoEvents
    
    If Not bVB_CNSysDateLocal And bOnOffCur Then
        dVB_CurrentDate = SP_GETdCurrentDate
        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
             Call SP_SYSxChangeSystemDate    '*KT 51-08-08  Communication Sheet ML-SF-434 Date On Terminal.xls เปลี่ยน System date ของเครื่อง POS ให้ตรงกันกับ Server
             bVB_CNSysDateLocal = True
         End If
    End If
    
    '*KT 51-07-24 Unlock
'    If bOnOffCur And bOnOff = False And oDB.IsReExecute = False Then   '*Ao 2013/10/18   V4.0.83
'        Call SP_TMNxUpdateEmpOnline
'        If SP_CHKbUnLock Then
'            Call SP_MSGnShowing("tms_CN005166", nCS_Error)
'            Call SP_PRCxCancelBillOnUnLock
'        End If
'    End If

    'display message to cashier  *KT 52-12-15 ทำการป้องกันการเข้าทำการ set timer  กรณีมีการแก้ไขเวลาที่ Active
    If (otmDisplay.Interval / 60000) <> nVB_TimeDisplayMsg Then
        If nVB_TimeDisplayMsg > 0 Then otmDisplay.Interval = IIf((nVB_TimeDisplayMsg * 60) * 1000 > 65535, 65535, (nVB_TimeDisplayMsg * 60) * 1000)
    End If

    'Verify Printer by OPOS
    If bVB_CNPrnOPOS Then Call SP_SHWxErrPrinter(False, tVB_CNTypePrinter)      '*Ao 2013/07/16  CommSheet TK-ISS3000-274 TR FAPL-053.xls
    Call SP_DATxRsNothing(orsTemp)
    
    '*KT 51-11-07 ป้องการการทำรายการเหล่านี้ ณ ที่ offline
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnIPV, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVRedeem, tEN_TrnIPVReturn, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVVIPService, tEN_TrnIPVOtherReceive, tEN_TrnIPVOtherDisbursement '
            If Not bOnOffCur Then
                MsgBox "Terminal is Offline", vbCritical 'tms_CN005222
                Select Case UCase(Screen.ActiveForm.Name)
                    Case UCase("wMainRet"), UCase("wMainRet640") 'wMain.Name '*KT 52-12-01 เปลี่ยนจาก  wMain.Name เป็น Fix ชื่อ wMainRet,wMainRet640
                        'เริ่ม Ticket ใหม่
                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                         'Clear Data DT,CD
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tSql = "DELETE " & tVB_CNTblHD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblCD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblRC
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
'                        Call UT_GenSaleNewDoc '*KT 53-02-26 ย้ายไปอยู่ใน wFunction/W_KBXEsc
                        bVB_IPV = False '*KT 53-02-26 จำเป็นต้องให้  bVB_IPV = False  ไม่อย่างนั้นจะทำงานใน otbPrompt_KeyDown vbKeyEscape, 0 ไม่ได้
                        otbPrompt_KeyDown vbKeyEscape, 0
'                    Case UCase("wCNSubPay"), UCase("wCNSubPay640")   'wSub.Name '*KT 52-12-01 เปลี่ยนจาก wSub.Name เป็น Fix ชื่อ wCNSubPay,wCNSubPay640
'                        'เริ่ม Ticket ใหม่
'                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
'                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
'                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
'                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
'                         'Clear Data DT,CD
'                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
'                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
'                        tSql = "DELETE " & tVB_CNTblHD
'                        tSql = tSql & vbCrLf & tWhe
'                        Call SP_SQLvExecute(tSql)
'                        tSql = "DELETE " & tVB_CNTblDT
'                        tSql = tSql & vbCrLf & tWhe
'                        Call SP_SQLvExecute(tSql)
'                        tSql = "DELETE " & tVB_CNTblCD
'                        tSql = tSql & vbCrLf & tWhe
'                        Call SP_SQLvExecute(tSql)
'                        tSql = "DELETE " & tVB_CNTblRC
'                        tSql = tSql & vbCrLf & tWhe
'                        Call SP_SQLvExecute(tSql)
'                        Call UT_GenSaleNewDoc
'                         bVB_IPV = False
'                         wSub.ocmCancel_Click
'                    Case UCase("wCNSubDisc2"), UCase("wCNSubDisc640")    'wDisc.Name '*KT 52-12-01 เปลี่ยนจาก wSub.Name เป็น Fix ชื่อ wCNSubPay,wCNSubPay640
'                        'เริ่ม Ticket ใหม่
'                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
'                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
'                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
'                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
'                         'Clear Data DT,CD
'                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
'                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
'                        tSql = "DELETE " & tVB_CNTblHD
'                        tSql = tSql & vbCrLf & tWhe
'                        Call SP_SQLvExecute(tSql)
'                        tSql = "DELETE " & tVB_CNTblDT
'                        tSql = tSql & vbCrLf & tWhe
'                        Call SP_SQLvExecute(tSql)
'                        tSql = "DELETE " & tVB_CNTblCD
'                        tSql = tSql & vbCrLf & tWhe
'                        Call SP_SQLvExecute(tSql)
'                        tSql = "DELETE " & tVB_CNTblRC
'                        tSql = tSql & vbCrLf & tWhe
'                        Call SP_SQLvExecute(tSql)
'                        Call UT_GenSaleNewDoc
'                         bVB_IPV = False
'                        wDisc.ocmCancel_Click
                    Case Else
                        bVB_TicketNotSuccess = True
                End Select
            End If
        Case tEN_TrnDeposit                 '*Ao 2013/09/25 LayawayCreate OFFLINE
        
        Case tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel        '*Ao 2013/09/25 LayawayCreate OFFLINE
            If Not bOnOffCur Then
                MsgBox "Terminal is Offline", vbCritical 'tms_CN005222
                Select Case Screen.ActiveForm.Name
                    Case UCase("wMainRet"), UCase("wMainRet640") 'wMain.Name '*KT 52-12-01 เปลี่ยนจาก  wMain.Name เป็น Fix ชื่อ wMainRet,wMainRet640
                        'เริ่ม Ticket ใหม่
                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        'Clear Data DT,CD
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tSql = "DELETE " & tVB_CNTblHD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblCD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblRC
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        Call UT_GenSaleNewDoc
                         bVB_TrnDepositRedeem = False
                         bVB_TrnRedeem = False
                         bVB_TrnDepositCancel = False
                         bVB_TrnAddDeposit = False
                        otbPrompt_KeyDown vbKeyEscape, 0
                    Case UCase("wCNSubPay"), UCase("wCNSubPay640")   'wSub.Name '*KT 52-12-01 เปลี่ยนจาก wSub.Name เป็น Fix ชื่อ wCNSubPay,wCNSubPay640
                        'เริ่ม Ticket ใหม่
                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        'Clear Data DT,CD
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tSql = "DELETE " & tVB_CNTblHD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblCD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblRC
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        Call UT_GenSaleNewDoc
                         bVB_TrnDepositRedeem = False
                         bVB_TrnRedeem = False
                         bVB_TrnDepositCancel = False
                         bVB_TrnAddDeposit = False
                        wSub.ocmCancel_Click
                    Case UCase("wCNSubDisc2"), UCase("wCNSubDisc640")    'wDisc.Name '*KT 52-12-01 เปลี่ยนจาก wSub.Name เป็น Fix ชื่อ wCNSubPay,wCNSubPay640
                         'เริ่ม Ticket ใหม่
                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        'Clear Data DT,CD
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tSql = "DELETE " & tVB_CNTblHD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblCD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblRC
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        Call UT_GenSaleNewDoc
                         bVB_TrnDepositRedeem = False
                         bVB_TrnRedeem = False
                         bVB_TrnDepositCancel = False
                         bVB_TrnAddDeposit = False
                        wDisc.ocmCancel_Click
                    Case Else
                        bVB_TicketNotSuccess = True
                End Select
            End If
        Case tEN_TrnVoucher, tEN_TrnCancelVoucher
             If Not bOnOffCur Then
                MsgBox "Terminal is Offline", vbCritical 'tms_CN005222
                Select Case Screen.ActiveForm.Name
                    Case UCase("wMainRet"), UCase("wMainRet640") 'wMain.Name '*KT 52-12-01 เปลี่ยนจาก  wMain.Name เป็น Fix ชื่อ wMainRet,wMainRet640
                        'เริ่ม Ticket ใหม่
                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        'Clear Data DT,CD
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tSql = "DELETE " & tVB_CNTblHD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblCD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblRC
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        Call UT_GenSaleNewDoc
                        bVB_TrnVoucher = False
                        bVB_AddVoucher = False
                        tVB_VoucherNo = ""
                        bVB_CancelVoucher = False
                        tVB_CancelVoucherNo = ""
                    Case UCase("wCNSubPay"), UCase("wCNSubPay640")   'wSub.Name '*KT 52-12-01 เปลี่ยนจาก wSub.Name เป็น Fix ชื่อ wCNSubPay,wCNSubPay640
                        'เริ่ม Ticket ใหม่
                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        'Clear Data DT,CD
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tSql = "DELETE " & tVB_CNTblHD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblCD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblRC
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        Call UT_GenSaleNewDoc
                        bVB_TrnVoucher = False
                        bVB_AddVoucher = False
                        tVB_VoucherNo = ""
                        bVB_CancelVoucher = False
                        tVB_CancelVoucherNo = ""
                        wSub.ocmCancel_Click
                    Case UCase("wCNSubDisc2"), UCase("wCNSubDisc640")    'wDisc.Name '*KT 52-12-01 เปลี่ยนจาก wSub.Name เป็น Fix ชื่อ wCNSubPay,wCNSubPay640
                        'เริ่ม Ticket ใหม่
                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        'Clear Data DT,CD
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tSql = "DELETE " & tVB_CNTblHD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblCD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblRC
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        Call UT_GenSaleNewDoc
                        bVB_TrnVoucher = False
                        bVB_AddVoucher = False
                        tVB_VoucherNo = ""
                        bVB_CancelVoucher = False
                        tVB_CancelVoucherNo = ""
                        wDisc.ocmCancel_Click
                    Case Else
                        bVB_TicketNotSuccess = True
                End Select
            End If
        Case tEN_TrnVIPSale, tEN_TrnCreateVIP
             If Not bOnOffCur Then
                MsgBox "Terminal is Offline", vbCritical 'tms_CN005222
                Select Case Screen.ActiveForm.Name
                    Case UCase("wMainRet"), UCase("wMainRet640") 'wMain.Name '*KT 52-12-01 เปลี่ยนจาก  wMain.Name เป็น Fix ชื่อ wMainRet,wMainRet640
                        'เริ่ม Ticket ใหม่
                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        'Clear Data DT,CD
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tSql = "DELETE " & tVB_CNTblHD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblCD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblRC
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        Call UT_GenSaleNewDoc
                        otbPrompt_KeyDown vbKeyEscape, 0
                    Case UCase("wCNSubPay"), UCase("wCNSubPay640")   'wSub.Name '*KT 52-12-01 เปลี่ยนจาก wSub.Name เป็น Fix ชื่อ wCNSubPay,wCNSubPay640
                        'เริ่ม Ticket ใหม่
                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        'Clear Data DT,CD
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tSql = "DELETE " & tVB_CNTblHD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblCD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblRC
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        Call UT_GenSaleNewDoc
                        wSub.ocmCancel_Click
                    Case UCase("wCNSubDisc2"), UCase("wCNSubDisc640")    'wDisc.Name '*KT 52-12-01 เปลี่ยนจาก wSub.Name เป็น Fix ชื่อ wCNSubPay,wCNSubPay640
                        'เริ่ม Ticket ใหม่
                         tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
                         tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
                         tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
                         tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
                        'Clear Data DT,CD
                        tWhe = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
                        tWhe = tWhe & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                        tSql = "DELETE " & tVB_CNTblHD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblDT
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblCD
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        tSql = "DELETE " & tVB_CNTblRC
                        tSql = tSql & vbCrLf & tWhe
                        Call SP_SQLvExecute(tSql)
                        Call UT_GenSaleNewDoc
                        wDisc.ocmCancel_Click
                    Case Else
                        bVB_TicketNotSuccess = True
                End Select
            End If
   End Select
   Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    If Err.Number = -2147467259 Or Err.Number = -2147217871 Or Err.Number = 91 Or Err.Number = 3704 Then   '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
        Call SP_MSGnShowing("tms_CN005199", nCS_Warning)
'        Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:otmLock" & " Events:Timer")
        Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:otmLock" & " Events:Timer")
        Call SP_APPxEndProgram
    ElseIf Err.Number = 440 Then
        Call SP_MSGxShowError(Err.Number, Err.Description)
    Else
        Call SP_MSGxShowError(Err.Number, Err.Description)
    End If
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " Form:" & Err.Description & Me.Name & ",otmLock_Timer[ERROR]")
End Sub


Private Sub W_FRMxShwMsg()
'----------------------------------------------
'Call :
'Cmt : แสดงข้อมูลแจ้ง Cashier ทุกเครื่อง Pos
'-----------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    tSql = "SELECT FTMsgLine1,FTMsgLine2,FTMsgLine3 "
    tSql = tSql & " FROM TCNMMessage"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            olaDisplay(0).Caption = SP_FEDtChkString(orsTemp, "FTMsgLine1")
            olaDisplay(1).Caption = SP_FEDtChkString(orsTemp, "FTMsgLine2")
            olaDisplay(2).Caption = SP_FEDtChkString(orsTemp, "FTMsgLine3")
        Else
            olaDisplay(0).Caption = ""
            olaDisplay(1).Caption = ""
            olaDisplay(2).Caption = ""
        End If
    Else
        olaDisplay(0).Caption = ""
        olaDisplay(1).Caption = ""
        olaDisplay(2).Caption = ""
    End If
'    If Not (GetInputState = 0) Then DoEvents '*KT 53-01-14 DoEvents
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub


Private Sub W_PRCxMergeForVIPSale()
'------------------------------------------------------------
'   Ret:
'   Cmt: ทำการรวมสินค้าที่มีการขาย VIP Service Sale ในแต่ละครั้งของลูกค้า VIP
'------------------------------------------------------------
    Dim tSql$, tTable$
    Dim orsHD As ADODB.Recordset
    Dim orsDT As ADODB.Recordset
    Dim orsCD As ADODB.Recordset
    Dim cGrand As Double
    Dim cTotal As Double
    Dim nSeqNo As Long
    Dim nSeqNoCD  As Long
    
    tSql = "STP_CNxSQLSelHoldHDSumByWhr @ptTransType ='" & Format(tEN_TrnCreateVIP, "00") & "',@ptTransDate = '" & tUT_SaleDate & "',@ptCstID ='" & tVB_CstID & "'"
    Call SP_SQLvExecute(tSql, orsHD, oDB.ocnOnLine) 'PH 2.0.0 connect database server
    If Not (orsHD.BOF Or orsHD.EOF) Then
        cTotal = SP_FEDvChkNumeric(orsHD, "SUMFCShdTotal")
        cGrand = SP_FEDvChkNumeric(orsHD, "SUMFCShdGrand")
        'update HD
        tSql = "UPDATE " & tVB_CNTblHD
        tSql = tSql & " SET FCShdTotal = " & cTotal
        tSql = tSql & " , FCShdGrand = " & cGrand
        tSql = tSql & "," & UT_tLastUpdSql
        tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
        tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
        tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
        tSql = tSql & "  AND FTShdStaDoc = '2'"
        tSql = tSql & " AND  FTCstVIPCode='" & tVB_CstID & "'"
        Call SP_SQLbExecute(tSql)
                        
        'insert DT
        tTable = tVB_CNTblDT
        tSql = "STP_CNxSQLSelHoldJoinHDDTByWhr @ptTransType ='" & Format(tEN_TrnCreateVIP, "00") & "',@ptTransDate = '" & tUT_SaleDate & "',@ptCstID ='" & tVB_CstID & "'"
        If SP_SQLvExecute(tSql, orsDT, oDB.ocnOnLine) <> 0 Then Exit Sub 'PH 2.0.0 connect database server
        '*KT 52-02-13  PH2-ML-SF-099 เมื่อมีการ Create ต่างเครื่องแล้วสินค้ามีส่วนลดแบบ Line item ( Line discount, Line mark down , CoDis )  แล้วเมื่อทำ VIP Service ส่วนลดจะหายไป แก้ไข STP_CNxSQLSelHoldJoinHDCDByWhr_4.0.sql
        tSql = "STP_CNxSQLSelHoldJoinHDCDByWhr @ptTransType ='" & Format(tEN_TrnCreateVIP, "00") & "',@ptTransDate = '" & tUT_SaleDate & "',@ptCstID ='" & tVB_CstID & "'"
        If SP_SQLvExecute(tSql, orsCD, oDB.ocnOnLine) <> 0 Then Exit Sub 'PH 2.0.0 connect database server
        
        nSeqNo = 1
        nSeqNoCD = 1
        Do While Not orsDT.EOF
            'CD
            tTable = tVB_CNTblCD
            orsCD.Filter = "FTShdTransNo ='" & SP_FEDtChkString(orsDT, "FTShdTransNo") & "' AND FTSkuCode ='" & SP_FEDtChkString(orsDT, "FTSkuCode") & "' AND FNSdtSeqNo =" & SP_FEDvChkNumeric(orsDT, "FNSdtSeqNo") & " AND FTTmnNum ='" & SP_FEDtChkString(orsDT, "FTTmnNum") & "'" '*KT 52-02-13 PH2-ML-SF-099 เมื่อมีการ Create ต่างเครื่องแล้วสินค้ามีส่วนลดแบบ Line item ( Line discount, Line mark down , CoDis )  แล้วเมื่อทำ VIP Service ส่วนลดจะหายไป
            If Not (orsCD.BOF Or orsCD.EOF) Then
                orsCD.MoveFirst
                Do While Not orsCD.EOF
                    Call wFunctionKB.W_DupTblBtwSleNHold(orsCD, tTable, nSeqNo, False, nSeqNo, nSeqNoCD, , , 2, nSeqNoCD)
                    nSeqNoCD = nSeqNoCD + 1
                    orsCD.MoveNext
                Loop
            End If
            orsCD.Filter = adFilterNone
            
            'DT
            tTable = tVB_CNTblDT
            Call wFunctionKB.W_DupTblBtwSleNHold(orsDT, tTable, nSeqNo, False, nSeqNo, , , True, 2)
            nSeqNo = nSeqNo + 1
            orsDT.MoveNext
        Loop
        
    End If
    
    
    Call SP_DATxRsNothing(orsHD)
    Call SP_DATxRsNothing(orsDT)
    Call SP_DATxRsNothing(orsCD)
    
End Sub

Private Sub ogdForm_ChangeEdit()            '*AO 51-03-31  Comm.Sheet ML-SF-160
'-----------------------------------------------------------------------------------------------------------------------
'cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
'กรุณาตรวจสอบความถูกต้องของ Communication Sheet ML-SF-381,ML-SF-382 ด้วยนะครับ
'-----------------------------------------------------------------------------------------------------------------------
    With ogdForm
        Select Case UCase(.ColKey(.Col))
            Case UCase("FCSdtSaleUnit")
                Select Case tW_KeyCode
                    Case 13 'ถ้าไม่ดัก KeyCode Enter จะหลุดไปที่ Case Else ทำให้ทำงานผิดพลาดนะครับ กรุณาระวังด้วย
                        Exit Sub
                    Case vbKeyDelete
                        .EditText = "0."
                         tW_KeyPress = "0"
                    Case Else
                        .EditText = SP_DEFtChangeNum(tW_KeyPress)
                        bW_ChkDecimals = False
                End Select
            Case UCase("FNSdtQty")
                Select Case tW_KeyCode
                    Case vbKeyDelete
                        .EditText = ""
                End Select
        End Select
    End With
End Sub

Public Function W_CHKbBanknoteAlw(ByVal pcAmt As Double) As Boolean           '*AO51-03-26   Comm.Sheet  ML-SF-100
'----------------------------------------------------------
'   Call:
'   Cmt:    Banknote
'----------------------------------------------------------
    Dim nDot As Integer, nIndex As Integer
    Dim cCHK As Double
    Dim tTmp As String
    Dim bLastFor As Boolean
    Dim cDiv As Double
    Dim nResultDiv As Long
On Error GoTo Err_Handle

    W_CHKbBanknoteAlw = False

    If bVB_ConfigBanknote Then     'เช็คสตางค์
        If Trim(CStr(aVB_CNStang(0))) = "" Then                     'ค่าว่าง ไม่มีเช็คสตางค์ปล่อยผ่าน
            W_CHKbBanknoteAlw = True
            Exit Function
        End If
        tTmp = Trim(CStr(pcAmt))
        nDot = InStr(1, tTmp, ".", vbBinaryCompare)
        If nDot = 0 Then                'ค่า .00 ไม่มีสตางค์
            W_CHKbBanknoteAlw = True
            Exit Function
        End If
        tTmp = Right(tTmp, (Len(tTmp) - nDot) + 1)
        cCHK = SP_VALcText2Double(tTmp)
        bLastFor = True

        Do While cCHK > 0 And bLastFor
            For nIndex = 0 To UBound(aVB_CNStang)
                If cCHK > 0 Then
                    '*KT 54-08-31 V.4.0.1 แก้ไขหน่วงเงินที่เล็กที่สุด
                     '----------------------------------------------------------------------------------
                    If SP_FMTcCurPoint(aVB_CNStang(nIndex)) <= SP_FMTcCurPoint(cCHK) Then
                        cCHK = SP_FMTcCurPoint(cCHK - aVB_CNStang(nIndex))
                        If nIndex = UBound(aVB_CNStang) Then
                            If cCHK > 0 Then
                                cDiv = SP_FMTcCurPoint(aVB_CNStang(nIndex))
                                nResultDiv = Int(SP_FMTcCurPoint(cCHK) / cDiv)
                                If (nResultDiv * cDiv) = cCHK Then
                                    cCHK = 0
                                End If
                            End If
                        End If
                    End If
                     '----------------------------------------------------------------------------------
                Else
                    Exit For
                End If
            Next nIndex
            bLastFor = False
        Loop
        If SP_FMTcCurPoint(cCHK) = 0 Then               'เช็คอีกครั้งว่าป้อนถูกไม่ ???
            W_CHKbBanknoteAlw = True
        Else
            W_CHKbBanknoteAlw = False
        End If
    Else
        W_CHKbBanknoteAlw = True           'ไม่ได้เช็คสตางค์ปล่อยผ่าน
    End If

    Exit Function
Err_Handle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    W_CHKbBanknoteAlw = True   'เกิด Error ปล่อยผ่าน
End Function

'Private Function W_DEFbDelectCD(ByVal pnSeqNo As Integer) As Boolean
''________________________________________
''   Call:
''   Cmt:
''________________________________________
'Dim tSql As String
'Dim nEff As Long
'On Error GoTo ErrHandle
'    W_DEFbDelectCD = False
'    If tVB_TransactionSub = tEN_TrnReturn Then       '* Ao 51-06-27 การคืนแบบไม่มีเงื่อนไข ไม่ต้องมีส่วนลดใดๆเกิดขึ้น
'        If bVB_Reference = False Then
'            tSql = "DELETE  "
'            tSql = tSql & " FROM " & tVB_CNTblCD
'            tSql = tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
'            tSql = tSql & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
'            tSql = tSql & " AND FNSdtSeqNo=" & pnSeqNo  'seq เดิมที่ถูกยกเลิก
'            Call SP_SQLvExecute(tSql, , , , , nEff)
'            If nEff > 0 Then
'                oDisc.dTranDate = tUT_SaleDate
'                oDisc.tTerminalNo = tVB_CNTerminalNum
'                oDisc.tTransNo = tVB_Running
'                oDisc.C_DATxSumDT
'            End If
'            W_DEFbDelectCD = True
'        End If
'    End If
'    Exit Function
'ErrHandle:
'    Call SP_MSGShowError
'End Function


Public Function W_PRCbRecovery() As Boolean
'________________________________________
'   Call:
'   Cmt: '*Eaw 56-01-07   กู้ข้อมูลบิลที่ยังไม่จบหากต้องการกู้ข้อมูลเดิมที่ทำไปแล้ว หรือไม่กู้ข้อมูล รายการ Transaction ที่ต้องกู้มีดังนี้
'Notmal Transaction
'sale = 30
'Return b= 04
'deposit = 06
'VIP Create = 30
'tax excep sale = 11
'duty free = 13
'Re entry Transaction
'SaleTaxInc = 38
'ReturnTaxInc = 39
'SaleManualTax = 44
'ReturnManualTax = 45
'________________________________________
    Dim tSql As String
    Dim tMemID$
    Dim orsTemp As ADODB.Recordset
     Dim orsTemp2 As ADODB.Recordset
    Dim nNum1 As Long
    Dim nNum2 As Long
    Dim nCountNoRe As Long
    Dim tWhere As String
On Error GoTo ErrHandle
    W_PRCbRecovery = False
    bVB_IsRecovery = False
    If SP_TBLbChkExist(tVB_CNTblHDR, oDB.ocnOffLine) Then
        tSql = "SELECT *  FROM " & tVB_CNTblHDR
        tSql = vbCrLf & tSql & " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "'"
        tSql = tSql & " AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & "  AND FDShdTransDate='" & tUT_SaleDate & "'"
        tSql = tSql & " AND FTShdStaDoc='" & 2 & "'"
        tSql = tSql & " AND FCShdTotal <> 0"
'        tSql = tSql & " AND FTShdTransType IN (03,04,06,10,11,13,38,39,44,45) "
        tSql = tSql & " AND FTShdTransType IN (03,04,06,11,13,30,38,39,44,45) "
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine)
        If Not (orsTemp.EOF And orsTemp.BOF) Then
'            If MsgBox("The Previous transaction is uncompleted ! " & vbCrLf & " Do you want to Recovery the transaction? ", vbQuestion + vbYesNo, "Recovery Transaction") = vbYes Then
            If True Then                            '*Ao 2014/05/28
                W_PRCbRecovery = True
                bVB_IsRecovery = True
                tVB_Transaction = SP_FEDtChkString(orsTemp, "FTShdTransType")
                tVB_TransactionSub = tVB_Transaction
                 Call W_FRMxInitailTrans '*KT 66-08-09 TTK-ISS3000-829 GST Rate for old Layaway transactions 2024
'                If Val(Format(tVB_TransactionSub)) = tEN_TrnDeposit Then    '*Eaw 56-07 การ Recovery ข้อมูล
'                    tVB_DepositNo = SP_FEDtChkString(orsTemp, "FTCdpDepositNo")
'                    cVB_DepositAmt = SP_FEDvChkNumeric(orsTemp, "FCShdDepositAmt")
'                    bVB_CNTrnDisc = True
'                End If
                Select Case Val(Format(tVB_TransactionSub))
                    Case tEN_TrnDeposit     '*Eaw 56-07 การ Recovery ข้อมูล
                        tVB_DepositNo = SP_FEDtChkString(orsTemp, "FTCdpDepositNo")
                        cVB_DepositAmt = SP_FEDvChkNumeric(orsTemp, "FCShdDepositAmt")
                        bVB_CNTrnDisc = True
                    Case tEN_TrnReturn    '*KT 66-08-18 TTK-ISS3000-839 support recovery return transaction
                        tVB_TerminalNoRet = SP_FEDtChkString(orsTemp, "FTShdDepRefTmnNum")
                        tVB_TransNoRet = SP_FEDtChkString(orsTemp, "FTShdDepReTransNo")
                        dVB_TransDateRet = SP_FEDdChkDate(orsTemp, "dVB_TransDateRet")
                        If tVB_TerminalNoRet <> "" Then
                            tSql = "SELECT FTShdTransType,FTCstCardCode,FTShdCshOrCrd,FTCstCardType FROM TPSTSalHD WITH(NOLOCK)"
                            tSql = tSql & vbCrLf & " WHERE FTTmnnUm='" & tVB_TerminalNoRet & "'"
                            tSql = tSql & vbCrLf & " AND FTTmnnUm='" & tVB_TransNoRet & "'"
                            tSql = tSql & vbCrLf & " AND FTTmnnUm='" & Format(dVB_TransDateRet, "yyyy/MMdd") & "'"
                            If SP_SQLvExecute(tSql, orsTemp2, , , , , , True) = 0 Then
                                If Not (orsTemp2.BOF And orsTemp2.EOF) Then
                                    tVB_TransTypeRet = SP_FEDtChkString(orsTemp2, "FTShdTransType")
                                    tVB_MemPointID = SP_FEDtChkString(orsTemp2, "FTCstCardCode")
                                    tVB_MallCardIDRef = SP_FEDtChkString(orsTemp2, "FTShdCshOrCrd") '*Tao 61-08-15 MallCardID จาก Manual ให้ใส่ *
                                    tVB_MemExp = SP_FEDtChkString(orsTemp2, "FTCstCardType") '*Tao 61-08-15 Expire Date
                                
                                    tWhere = "WHERE FTTmnNum = '" & tVB_TerminalNoRet & "'" & " AND FTShdTransNo = '" & tVB_TransNoRet & "'" & "  AND FDShdTransDate = '" & Format(dVB_TransDateRet, "yyyy/MM/dd") & "'" & " AND FTShdTransType = '" & tVB_TransTypeRet & "'"
                                    nNum1 = Val(Format(SP_SQLtFunction("COUNT", "FNSdtSeqNo", tVB_CNTblDT, tWhere, , True)))
                                    nNum2 = ogdForm.Aggregate(flexSTCount, ogdForm.FixedRows, ogdForm.ColIndex("FNSdtSeqNo"), ogdForm.rows - 1, ogdForm.ColIndex("FNSdtSeqNo"))
                                    tWhere = "WHERE ISNULL(FCSdtReturnQty,0)= ISNULL(FCSdtQty,0) AND FTTmnNum = '" & tVB_TerminalNoRet & "'" & " AND FTShdTransNo = '" & tVB_TransNoRet & "'" & "  AND FDShdTransDate = '" & Format(dVB_TransDateRet, "yyyy/MM/dd") & "'" & " AND FTShdTransType = '" & tVB_TransTypeRet & "'"
                                    bVB_ReturnAll = Val(Format(SP_SQLtFunction("COUNT", "FNSdtSeqNo", tVB_CNTblDT, tWhere, , True)))
                                End If
                            End If

                            If (nNum1 - nCountNoRe) = nNum2 Then
                                bVB_ReturnAll = True
                            Else
                                bVB_ReturnAll = False
                            End If
                            '*KT 62-09-06 เก็บสถานะว่ามีการคืนครั้งเดียวทั้งบิล
                            If nNum1 = nNum2 And nCountNoRe = 0 Then
                                bVB_ReturnAllBill = True
                            Else
                                bVB_ReturnAllBill = False
                            End If
                        Else
                            tVB_TerminalNoRet = ""
                            tVB_TransNoRet = ""
                            dVB_TransDateRet = Date
                            tVB_TransTypeRet = ""
                            tVB_MemPointID = ""
                            tVB_MallCardIDRef = ""
                            tVB_MemExp = ""
                            
                            bVB_ReturnAllBill = False
                            bVB_ReturnAll = False
                        End If
                    Case Else
                        'none
                End Select
                
                xVB_CNMemberP.ID = SP_FEDtChkString(orsTemp, "FTShdCpnCodeRef")
                If xVB_CNMemberP.ID <> "" Then
                    xVB_CNMemberP.MallCardMemberPrice = "Y"
                    bVB_CNMemberPrice = True
                    wMain.ostForm.Panels(2).Text = "Mem: " & xVB_CNMemberP.ID
                End If
                 '*KT 66-08-09 TTK-ISS3000-829 GST Rate for old Layaway transactions 2024
                '-----------------------------------------------------------------------------------------------------
                cVB_OldGstRate = SP_FEDvChkNumeric(orsTemp, "FCShdVATRate")
                If cVB_OldGstRate > 0 And tVB_TransactionSub = "04" Then
                    bVB_ChkRetGST7 = True
                End If
                 '-----------------------------------------------------------------------------------------------------
                
                Call W_GETxSlipPrintType(tVB_Transaction)  '*Eaw 56-09-04 CommSheet TK-ISS3000-338 UAT FAPL-074.xlsx
                '*Eaw 56-07-02
'                tSql = "UPDATE " & tVB_CNTblHD
'                tSql = tSql & vbCrLf & "SET FCShdWpTax = " & SP_FEDvChkNumeric(orsTemp, "FCShdWpTax")
'                tSql = tSql & vbCrLf & "WHERE FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
'                tSql = tSql & vbCrLf & "AND FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
                tSql = "UPDATE " & tVB_CNTblHD
                tSql = tSql & vbCrLf & "SET FCShdWpTax = " & SP_FEDvChkNumeric(orsTemp, "FCShdWpTax")
                tSql = tSql & vbCrLf & ",FNCstCardPoint = 0"
                tSql = tSql & vbCrLf & ",FTCstCardCode = '" & SP_FEDtChkString(orsTemp, "FTCstCardCode") & "'"          '*Ao 2014/05/07  PRJ10001TK-PH2-External Spec-Incomplete Transaction-00.02.04 FSBT
                tSql = tSql & vbCrLf & ",FTCstCardType = '" & SP_FEDtChkString(orsTemp, "FTCstCardType") & "'"
                tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "' AND FTShdTransNo = '" & tVB_Running & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "' AND FDShdTransDate = '" & tUT_SaleDate & "'"
                Call SP_SQLvExecute(tSql)
                Call SP_DATxRsNothing(orsTemp)
                Call SP_DATxRsNothing(orsTemp2)
                Exit Function
            End If
        End If
        tSql = "DROP TABLE  " & tVB_CNTblHDR
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        If SP_TBLbChkExist(tVB_CNTblDTR, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblDTR
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
        If SP_TBLbChkExist(tVB_CNTblCDR, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblCDR
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
        If SP_TBLbChkExist(tVB_CNTblRCR, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblRCR
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
        If SP_TBLbChkExist(tVB_CNTblDTRTmp, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblDTRTmp
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
        If SP_TBLbChkExist(tVB_CNTblCDRTmp, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblCDRTmp
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
        W_PRCbRecovery = False
        bVB_IsRecovery = False
    Else  '*Eaw 56-07-03
        If SP_TBLbChkExist(tVB_CNTblDTR, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblHDR
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
        If SP_TBLbChkExist(tVB_CNTblDTR, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblDTR
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
        If SP_TBLbChkExist(tVB_CNTblCDR, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblCDR
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
        If SP_TBLbChkExist(tVB_CNTblRCR, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblRCR
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
        If SP_TBLbChkExist(tVB_CNTblDTRTmp, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblDTRTmp
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
        If SP_TBLbChkExist(tVB_CNTblCDRTmp, oDB.ocnOffLine) Then
            tSql = "DROP TABLE  " & tVB_CNTblCDRTmp
            Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsTemp2)
    Exit Function
ErrHandle:
    W_PRCbRecovery = False
    bVB_IsRecovery = False
    Call SP_MSGShowError
    Call SP_DATxRsNothing(orsTemp)
    Call SP_DATxRsNothing(orsTemp2)
End Function

Public Function W_PRCxGetData()
'________________________________________
'   Call:
'   Cmt: '*Eaw 56-01-07   กู้ข้อมูลบิลที่ยังไม่จบโดยดึงข้อมูลจากตาราง Tmp ที่สร้างขึ้นมาเก็บข้อมูลชั่วคราว
'________________________________________
Dim tSql$, tWhr$, tUpd$
Dim orsTemp As ADODB.Recordset
Dim tName As String '*Eaw 56-11-01
On Error GoTo ErrHandle
    tName = SP_STRtSiggleCode(tVB_CNUserName)  '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    '*Eaw 56-04-10  CommSheet TK-ISS3000-129-IT2 00.00.01.xls
    tUpd = " SET FDDateUpd = '" & UT_tStringDate(Date) & "',"
    tUpd = tUpd & vbCrLf & " FTTimeUpd = '" & Time$ & "',"
    tUpd = tUpd & vbCrLf & " FTWhoUpd = '" & tName & "'," '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    tUpd = tUpd & vbCrLf & " FDDateIns = '" & UT_tStringDate(Date) & "',"
    tUpd = tUpd & vbCrLf & " FTTimeIns = '" & Time$ & "',"
    tUpd = tUpd & vbCrLf & " FTWhoIns = '" & tName & "'" '*Eaw 56-11-01ให้ชื่อมีเขาเดี่ยวได้
    '----------------------------------------------------------------------------------------------------------
    tWhr = "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & "  AND FTShdTransNo='" & tVB_Running & "'"
'    tWhr = tWhr & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = "DELETE FROM  " & tVB_CNTblHD & " WITH(ROWLOCK) "
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    
    tSql = "INSERT INTO  " & tVB_CNTblHD & " SELECT * FROM " & tVB_CNTblHDR
    tSql = tSql & tWhr
'    tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & " AND FTShdStaDoc='" & 2 & "'"
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    
    tSql = "DROP TABLE  " & tVB_CNTblHDR
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '*Eaw 56-04-10  CommSheet TK-ISS3000-129-IT2 00.00.01.xls
    tSql = "UPDATE " & tVB_CNTblHD & tUpd & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '-----------------------------------------------------
    tSql = "INSERT INTO  " & tVB_CNTblDT & " SELECT * FROM " & tVB_CNTblDTR
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '--------------------------------------------------------------
'*Eaw 56-07-02 Test DTTmp
    tSql = "DELETE FROM  " & tVB_CNTblDTTmp & " WITH(ROWLOCK) "
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "INSERT INTO  " & tVB_CNTblDTTmp & " SELECT * FROM " & tVB_CNTblDTRTmp
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '--------------------------------------------------------------
    tSql = "INSERT INTO  " & tVB_CNTblCD & " SELECT * FROM " & tVB_CNTblCDR
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '--------------------------------------------------------------
    tSql = "DELETE FROM  " & tVB_CNTblCDTmp & " WITH(ROWLOCK) "
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    tSql = "INSERT INTO  " & tVB_CNTblCDTmp & " SELECT * FROM " & tVB_CNTblCDRTmp
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
'    --------------------------------------------------------------
'*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
'    tSql = "INSERT INTO  " & tVB_CNTblRC & " SELECT * FROM " & tVB_CNTblRCR
'    tSql = tSql & tWhr
'    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    '--------------------------------------------------------------
    Exit Function
ErrHandle:
    Call SP_MSGShowError
End Function

Public Function W_RECxProduct2Grid()
'________________________________________
'   Call:
'   Cmt: '*Eaw 56-01-07   add ข้อมูลที่กู้มาจาก DB ลงใน Grid
'________________________________________
Dim tSql$, tWhr$
Dim orsTemp As ADODB.Recordset
Dim cPrice As Double
Dim tBarcode As String
Dim tSKUName As String
Dim nRow As Long
On Error GoTo ErrHandle
    tWhr = "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & "  AND FTShdTransNo='" & tVB_Running & "'"
    tWhr = tWhr & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    
    tSql = "SELECT *  FROM " & tVB_CNTblDT
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine)
    If Not (orsTemp.EOF And orsTemp.BOF) Then
    orsTemp.MoveFirst
        Do While Not orsTemp.EOF
                With ogdForm
                     nRow = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FNSdtSeqNo"), .rows - 1, .ColIndex("FNSdtSeqNo"))
                     If nRow = .rows - 1 Then .rows = .rows + 1
                     If .rows - 1 >= nUT_GridMaxRow Then
                        .TopRow = .rows - (nUT_GridMaxRow - 1) '.TopRow + 1
                    End If

                    tVB_SKUSymbol = ""
                     '*Eaw 55-12-06  เปลี่ยนเป็นใช้เครื่องหมาย *สำหรับ Not Alw Member Card ,# สำหรับ Not Alw Emp Card
                    tVB_CNSkuStaAlw = SP_FEDtChkString(orsTemp, "FTSdtStaPdt")
                    If Mid(tVB_CNSkuStaAlw, 20, 1) = "0" Then tVB_SKUSymbol = "#"       'ไม่ให้ส่วนลดEmp card
                    If Mid(tVB_CNSkuStaAlw, 9, 1) = "0" Then tVB_SKUSymbol = "*"       'ไม่ให้ส่วนลดMember card
                    If Mid(tVB_CNSkuStaAlw, 20, 1) = "0" And Mid(tVB_CNSkuStaAlw, 9, 1) = "0" Then tVB_SKUSymbol = "*" 'ไม่ให้ส่วนลด Mall+Emp
                    
                    .TextMatrix(nRow + 1, .ColIndex("FNSdtSeqNo")) = SP_FEDvChkNumeric(orsTemp, "FNSdtSeqNo")
                    .TextMatrix(nRow + 1, .ColIndex("FTPmt")) = IIf(bVB_SKUisPromotion, "*", "")
                    .TextMatrix(nRow + 1, .ColIndex("FTSymbol")) = tVB_SKUSymbol '*KT 55-08-17 Takashimaya
                    .TextMatrix(nRow + 1, .ColIndex("FTSdtBarCode")) = SP_FEDtChkString(orsTemp, "FTSdtBarCode")
                    tBarcode = SP_FEDtChkString(orsTemp, "FTSdtBarCode")
                    .TextMatrix(nRow + 1, .ColIndex("FTSkuCode")) = SP_FEDtChkString(orsTemp, "FTSkuCode")  'รหัสสินค้า
                    .TextMatrix(nRow + 1, .ColIndex("FCSdtQty")) = SP_FEDvChkNumeric(orsTemp, "FCSdtQty") 'จำนวนเป็นจำนวนเต็ม
                    .TextMatrix(nRow + 1, .ColIndex("FCSdtRegPrice")) = SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")       'ราคา ป.ป.
'                    cPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")  '*Eaw 56-07-03
                    .TextMatrix(nRow + 1, .ColIndex("FNSdtQty")) = SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll") 'จำนวนเป็นจำนวนเต็ม
                    .TextMatrix(nRow + 1, .ColIndex("FCSdtSalePrice")) = SP_FEDvChkNumeric(orsTemp, "FCSdtSalePrice")       'ราคา ป.ป.
                    cPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleUnit")  '*Eaw 56-07-03
                    .TextMatrix(nRow + 1, .ColIndex("FCSdtSaleUnit")) = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleUnit")       'ราคา ป.ป.
'                    cPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleUnit")
                    .TextMatrix(nRow + 1, .ColIndex("FTSkuAbbName")) = SP_FEDtChkString(orsTemp, "FTSkuAbbName")
                    tSKUName = SP_FEDtChkString(orsTemp, "FTSkuAbbName")
                    .TextMatrix(nRow + 1, .ColIndex("FTSkuAbbNameSnd")) = SP_FEDtChkString(orsTemp, "FTSkuAbbNameSnd")
                    .TextMatrix(nRow + 1, .ColIndex("FTSkuAbbNameSrvDoc")) = SP_FEDtChkString(orsTemp, "FTSkuAbbNameSrvDoc")
                    .TextMatrix(nRow + 1, .ColIndex("FTSkuAbbNameSndSrvDoc")) = SP_FEDtChkString(orsTemp, "FTSkuAbbNameSndSrvDoc")
                    .TextMatrix(nRow + 1, .ColIndex("FTSdtUnitName")) = SP_FEDtChkString(orsTemp, "FTSdtUnitName")
                    .TextMatrix(nRow + 1, .ColIndex("FCSdtDis")) = SP_FEDvChkNumeric(orsTemp, "FCSdtDis")
                    .TextMatrix(nRow + 1, .ColIndex("FCSdtChg")) = SP_FEDvChkNumeric(orsTemp, "FCSdtChg")
                    .TextMatrix(nRow + 1, .ColIndex("FTTaxCode")) = SP_FEDtChkString(orsTemp, "FTTaxCode")
                    .TextMatrix(nRow + 1, .ColIndex("FCSdtTax")) = SP_FEDvChkNumeric(orsTemp, "FCSdtTax")
                    .TextMatrix(nRow + 1, .ColIndex("FCSdtQtyAll")) = SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAll")
                    .TextMatrix(nRow + 1, .ColIndex("FTTaxSign")) = SP_FEDtChkString(orsTemp, "FTSkuAbbName")
                    .TextMatrix(nRow + 1, .ColIndex("FTTaxStaSign")) = SP_FEDtChkString(orsTemp, "FTSkuAbbName")
                    .TextMatrix(nRow + 1, .ColIndex("FCSdtSaleAmt")) = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt")
                    .TextMatrix(nRow + 1, .ColIndex("FTSdtPriceOverideFlag")) = SP_FEDtChkString(orsTemp, "FTSdtPriceOverideFlag")
                    .TextMatrix(nRow + 1, .ColIndex("FTSdtCancelFlag")) = SP_FEDtChkString(orsTemp, "FTSdtCancelFlag")
                    .TextMatrix(nRow + 1, .ColIndex("FTSdtStaPdt")) = SP_FEDtChkString(orsTemp, "FTSdtStaPdt")
                    .TextMatrix(nRow + 1, .ColIndex("FTSdtStaRfd")) = SP_FEDtChkString(orsTemp, "FTSdtStaRfd")
                    .TextMatrix(nRow + 1, .ColIndex("FTSdtWeight")) = SP_FEDtChkString(orsTemp, "FTSdtWeight")
                    .TextMatrix(nRow + 1, .ColIndex("FTSdtConsignFlag")) = SP_FEDtChkString(orsTemp, "FTSdtConsignFlag")
                    .TextMatrix(nRow + 1, .ColIndex("FTSdtStaPreview")) = SP_FEDtChkString(orsTemp, "FTSdtStaPreview")
                     .TextMatrix(nRow + 1, .ColIndex("FTSdtStaNotAlwPrv")) = SP_FEDtChkString(orsTemp, "FTSdtStaNotAlwPrv")
                    .Refresh
                orsTemp.MoveNext
            End With
            Loop
        End If
        Call W_SumAmount
        tVB_TransName = aVB_TransMode(Val(Format(tVB_TransactionSub)) - 1).Name
        '*KT 66-08-09 TTK-ISS3000-829 GST Rate for old Layaway transactions 2024
        If bVB_ChkRetGST7 Then
            tVB_TransName = cVB_OldGstRate & "% " & tVB_TransName
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbReturn - 1).UsrLangPri & "" & cVB_OldGstRate & "%;" & aVB_NameAbb(tEN_AbbReturn - 1).UsrLangSec & "" & cVB_OldGstRate & "%"
        End If
        
        olaTransName.Caption = IIf(bVB_TrainingMode, "Tr\" & tVB_TransName, IIf(bVB_ReEntryMode, "Re Entry\" & tVB_TransName, tVB_TransName))
        olaPrice.Caption = SP_FMTcCurPoint(cPrice)
        olaBarcode.Caption = tBarcode
        olaSKUName.Caption = tSKUName
        olaTSubReg.Caption = SP_FMTcCurPoint(cUT_SubReg)
        olaTSubPro.Caption = SP_FMTcCurPoint(cUT_SubPro)
        olaTSubTotal.Caption = SP_FMTcCurPoint(cUT_SubTotal)
        olaSumQty.Caption = nUT_SumQty

        
        tSql = "DROP TABLE  " & tVB_CNTblDTR
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)

        Call SP_DATxRsNothing(orsTemp)
        Exit Function
        
ErrHandle:
    Call SP_MSGShowError
End Function

'*Eaw 56-07-02 CommSheet TK-ISS3000-249-TR.xls ไม่ต้อง Recovery RC
'Public Function W_CHKxRCTmp()
''________________________________________
''   Call:
''   Cmt: '*Eaw 56-01-07   Recovery Transaction ตรวจสอบว่าบิลที่กู้คืนมีการจ่ายเงินบางส่วนแล้วหรือไม่
''________________________________________
'Dim tSql$, tWhr$
'Dim orsTemp As ADODB.Recordset
'On Error GoTo ErrHandle
'    bVB_CNPaid = False
'    tWhr = "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
'    tWhr = tWhr & "  AND FTShdTransNo='" & tVB_Running & "'"
'    tWhr = tWhr & "  AND FTShdTransType='" & tVB_Transaction & "'"
'    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
'    tSql = "SELECT *  FROM " & tVB_CNTblRCR
'    tSql = tSql & tWhr
'    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine)
'    If Not (orsTemp.EOF And orsTemp.BOF) Then
'        bVB_CNPaid = True
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'    Exit Function
'ErrHandle:
'    Call SP_MSGShowError
'End Function

Public Function W_CHKxCDTmp()
'________________________________________
'   Call:
'   Cmt: '*Eaw 56-01-07   Recovery Transaction ตรวจสอบว่าบิลที่กู้คืนมีการให้ส่วนลดท้ายบิลแล้วหรือไม่
'________________________________________
Dim tSql$, tWhr$
Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    bVB_CNDiscount = False
    tWhr = "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & "  AND FTShdTransNo='" & tVB_Running & "'"
    tWhr = tWhr & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tWhr = tWhr & "  AND FNDctNo IN (7,8,9,10,11,12,13,14,23,24,25,26,27)"
    tSql = "SELECT * FROM " & tVB_CNTblCDRTmp
    tSql = tSql & tWhr
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine)
    If Not (orsTemp.EOF And orsTemp.BOF) Then
        bVB_CNDiscount = True
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    Call SP_MSGShowError
End Function




