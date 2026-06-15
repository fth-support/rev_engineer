VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form wCNEJViewer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Electronic Journal Viewer;klg_Title"
   ClientHeight    =   7980
   ClientLeft      =   1185
   ClientTop       =   1155
   ClientWidth     =   11520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.DirListBox odiEJ 
      Height          =   990
      Left            =   0
      TabIndex        =   16
      Top             =   10440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.FileListBox oflEJ 
      Height          =   1455
      Left            =   2760
      TabIndex        =   15
      Top             =   10080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.DriveListBox odrEJ 
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   10080
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.PictureBox opbForm 
      BackColor       =   &H80000009&
      Height          =   5940
      Left            =   15600
      ScaleHeight     =   5880
      ScaleWidth      =   3195
      TabIndex        =   7
      Top             =   7440
      Width           =   3255
   End
   Begin MSComctlLib.ImageList oilForm 
      Left            =   6960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":0000
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":6862
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":D0C4
            Key             =   "Sep1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":D4F5
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":13D57
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":1A5B9
            Key             =   "Defresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":20E1B
            Key             =   "Sep2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":2124C
            Key             =   "MPreviors"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":27AAE
            Key             =   "MNext"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":2E310
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":34B72
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":3B3D4
            Key             =   "Sep3"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":3B805
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":40FF7
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":47859
            Key             =   "Pay"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":4E0BB
            Key             =   "Cancle"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":5491D
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":5B17F
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":619E1
            Key             =   "Post"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":68243
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":6EAA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":84ACF
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "wCNEJViewer.frx":89729
            Key             =   "THA"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox oucPanel 
      Height          =   8055
      Index           =   0
      Left            =   0
      ScaleHeight     =   7995
      ScaleWidth      =   11505
      TabIndex        =   8
      Top             =   0
      Width           =   11565
      Begin VB.PictureBox oucPanel 
         Appearance      =   0  'Flat
         Height          =   825
         Index           =   1
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   11130
         TabIndex        =   9
         Top             =   120
         Width           =   11160
         Begin VB.CheckBox ockDate 
            Caption         =   "Check1"
            Height          =   195
            Left            =   240
            TabIndex        =   1
            Top             =   405
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CommandButton ocmBrwCstTel 
            Caption         =   "Search;klg_Search"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   9720
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   315
            Width           =   1305
         End
         Begin VB.TextBox otbForm 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   7680
            MaxLength       =   5
            TabIndex        =   0
            Tag             =   "2;"
            Top             =   315
            Width           =   1935
         End
         Begin VB.TextBox otbForm 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   4200
            MaxLength       =   5
            TabIndex        =   3
            Tag             =   "2;"
            Top             =   315
            Width           =   1740
         End
         Begin MSComCtl2.DTPicker odtSearch 
            Height          =   375
            Left            =   1320
            TabIndex        =   2
            Top             =   315
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   661
            _Version        =   393216
            Format          =   145620993
            CurrentDate     =   39351
         End
         Begin VB.Label olaForm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search;klg_Search"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Tag             =   "2"
            Top             =   30
            Width           =   1620
         End
         Begin VB.Label olaForm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เลขที่ใบสำคัญ;klg_TranNo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   6000
            TabIndex        =   12
            Tag             =   "2;"
            Top             =   360
            Width           =   2235
         End
         Begin VB.Label olaForm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date:;klg_Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   525
            TabIndex        =   11
            Tag             =   "2;"
            Top             =   405
            Width           =   1380
         End
         Begin VB.Label olaForm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "หมายเลข;klg_POSNo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   3120
            TabIndex        =   10
            Tag             =   "2;"
            Top             =   405
            Width           =   1830
         End
      End
      Begin RichTextLib.RichTextBox orbForm 
         Height          =   6225
         Left            =   4005
         TabIndex        =   6
         Top             =   1200
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   10980
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"wCNEJViewer.frx":89CC3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier MonoThai"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VSFlex7Ctl.VSFlexGrid ogdFile 
         Height          =   6255
         Left            =   120
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   1200
         Width           =   3645
         _cx             =   6429
         _cy             =   11033
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   12648384
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16761024
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   18
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wCNEJViewer.frx":89D49
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
      Begin VB.Label olaPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   7560
         Width           =   405
      End
   End
End
Attribute VB_Name = "wCNEJViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*Eaw 56-11-25 DCR_5608002(293 EJ Viewer at POS)
Option Explicit

'ตัวแปรสำหรับผู้ใช้
Dim bW_FirstLoad As Boolean
Dim bW_FRMAc As Boolean
Dim tW_MnuCurrent As String
Dim tW_FileFind As String
Dim tW_EJPath As String
Dim tW_CDatePath As String
Dim nW_FindRow As Long
Dim dW_OriginalDate As Long

Public Sub W_DEFxSetGrdCaptionAll(ByVal ptGridName As String)  'Copy Code Function นี้จาก Form ZAPTemplate เข้าไปไว้ใน Form ของตัวเอง
'----------------------------------------------------------
'   Call:set Caption for all Grid
'
'----------------------------------------------------------
    Select Case UCase(ptGridName)
        Case "OGDFILE"
            Call SP_GETxGridCaptionAndKey(ogdFile, "ogd_File", Me.Name)
    End Select
End Sub

Private Sub W_PRNxPrintEJ()
'__________________________
'   Call:
'   Cmt:    พิมพ์ EJ
'__________________________
    On Error GoTo ErrHandle
    If Trim$(orbForm.Text) <> "" Then
        Printer.CurrentX = 1000
        orbForm.SelPrint (Printer.hDC)
        Printer.EndDoc
    End If
    Exit Sub
    
ErrHandle:
    Call SP_MSGShowError
End Sub

Private Function W_GETtTrimNull(ByVal ptValue As String) As String
'-----------------------------------------------------------
'   Cmt:    ตัดตัวอักษร Null
'-----------------------------------------------------------
    Dim nCh&, nL&
    nL = Len(ptValue)
    W_GETtTrimNull = Trim$(ptValue)
    
    For nCh = 1 To nL
        If Asc(Mid(ptValue, nCh, 1)) < 32 Then
            W_GETtTrimNull = Trim$(Left$(ptValue, nCh - 1))
            Exit For
        End If
    Next nCh
End Function

Private Function W_GETbDir(poBuff As WIN32_FIND_DATA) As Boolean
'-----------------------------------------------------------
'   Cmt:    ตรวจสอบว่าเป็น Directory หรือไม่
'-----------------------------------------------------------
    Dim tFName$
    
    tFName = W_GETtTrimNull(poBuff.cFileName)
    W_GETbDir = ((poBuff.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) > 0)
End Function

Private Function W_SETFocusRows(ByVal ptEJName As String) As Boolean
    '-----------------------------------------------------------
    '*Bump 57-07-17 เซทโฟกัส EJ ตามที่ค้นหาด้วย TransNo.
    '-----------------------------------------------------------
    Dim nLRow As Long
    Dim nI As Long
    Dim bFound As Boolean
    Dim tFindText As String
On Error GoTo ErrHandle
    bFound = False
    W_SETFocusRows = False
    
    tFindText = ptEJName '**"DC000650023601.EJ"
    With ogdFile
        .Enabled = True
        .HighLight = flexHighlightAlways
        'nLRow = .Aggregate(flexSTMax, .FixedRows, .ColIndex("Name"), .rows - 1, .ColIndex("Name"))
        nLRow = ogdFile.rows - 1
    '       orbForm.LoadFile (ogdFile.TextMatrix(ogdFile.Row, ogdFile.ColIndex("FileName")))
    'ogdFile.SetFocus
        nI = nLRow
        Do While nI >= 1 And nI <= nLRow
                If Trim(.TextMatrix(nI, .ColIndex("Name"))) = tFindText Then bFound = True: Exit Do
            nI = nI - 1
        Loop
        If bFound Then
          '  .Enabled = True
           ' .HighLight = flexHighlightAlways
            .Select nI, 1
            .SetFocus
            nW_FindRow = nI
            W_SETFocusRows = True
        Else
            '.HighLight = flexHighlightNever
            '.Editable = flexEDNone
            '.Enabled = False
        End If
    End With
ErrHandle:
End Function
Private Sub W_FINxPathDte(ByVal ptPath$)
'-----------------------------------------------------------
'   Cmt:    ค้นหา Dir ใน Path  Date
'-----------------------------------------------------------
    Dim tTmnPath$, tBuff$, tFilter$, tFileName$
    Dim nHandle&
    Dim oBuff As WIN32_FIND_DATA


    tFilter = "??" & Left$(Trim(otbForm(6).Text) & "?????", 5) & Left$(Trim(otbForm(0).Text) & "?????", 5) & "????.EJ"

    nHandle = FindFirstFile(ptPath & "\" & tFilter, oBuff)
    If nHandle <> INVALID_HANDLE_VALUE Then
        If Not W_GETbDir(oBuff) Then
            tBuff = W_GETtTrimNull(oBuff.cFileName)
            tFileName = ptPath & "\" & tBuff
             '*Bump 57-07-17 TK-ISS3000PH2-521
            If W_SETFocusRows(tBuff) Then
                Call FindClose(nHandle)
                Exit Sub
            End If
            '-------------------------------------------------------
            With ogdFile
                .rows = .rows + 1
                .TextMatrix(.rows - 1, .ColIndex("FileName")) = tFileName
                .TextMatrix(.rows - 1, .ColIndex("Name")) = tBuff
                .TextMatrix(.rows - 1, .ColIndex("Date")) = tW_CDatePath
            End With
        End If

        Do While FindNextFile(nHandle, oBuff)
            If Not W_GETbDir(oBuff) Then
                tBuff = W_GETtTrimNull(oBuff.cFileName)
                tFileName = ptPath & "\" & tBuff
                With ogdFile
                    .rows = .rows + 1
                    .TextMatrix(.rows - 1, .ColIndex("FileName")) = tFileName
                    .TextMatrix(.rows - 1, .ColIndex("Name")) = tBuff
                    .TextMatrix(.rows - 1, .ColIndex("Date")) = tW_CDatePath
                End With
            End If
        Loop

        Call FindClose(nHandle)
    End If

End Sub

Private Sub W_FINxPathEJ(Optional ByVal ptDate$ = "*.*")
'-----------------------------------------------------------
'   Cmt:    ค้นหา Dir ใน Path Image
'-----------------------------------------------------------
    Dim tDtePath$, tPath$, tBuff$
    Dim nHandle&
    Dim oBuff As WIN32_FIND_DATA
    
    tPath = tW_EJPath & "\" & ptDate
    nHandle = FindFirstFile(tPath, oBuff)
    If nHandle <> INVALID_HANDLE_VALUE Then
        If W_GETbDir(oBuff) Then 'Directory
            tBuff = W_GETtTrimNull(oBuff.cFileName)
            If (tBuff <> ".") And (tBuff <> "..") Then
                tW_CDatePath = tBuff 'Current Date
                tDtePath = tW_EJPath & "\" & tBuff
                Call W_FINxPathDte(tDtePath)
            End If
        End If
            
        Do While FindNextFile(nHandle, oBuff)
            If W_GETbDir(oBuff) Then 'Directory
                tBuff = W_GETtTrimNull(oBuff.cFileName)
                If (tBuff <> ".") And (tBuff <> "..") Then
                    tW_CDatePath = tBuff 'Current Date
                    tDtePath = tW_EJPath & "\" & tBuff
                    Call W_FINxPathDte(tDtePath)
                End If
            End If
        Loop
            
        Call FindClose(nHandle)
    End If
End Sub

Private Sub Form_Load()
    Dim tTmp As String
    Dim oFSO As New FileSystemObject
    tVB_CNMnuCurrent = "1005"
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    
    Call SP_FRMxChildCenter(Me)
    Call ockDate_Click
    
    tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\Backup" 'หา Path ของการสร้าง ej ที่ Loacal (folder = "POSEJ")
    If Not SP_FLDbManFolder(tTmp, nEN_Fld3Create) Then 'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
           tTmp = App.Path & "\POSEJ\Backup" 'หา Path ของการสร้าง ej ที่ Loacal ใหม่
           Call SP_FLDbManFolder(tTmp, nEN_Fld3Create) 'ตรวจสอบว่า folder มีหรือไม่ ถ้าไม่มีให้สร้าง folder ใหม่
    End If
    
    Call SP_SETxCreateEJAll(True)
    
    tW_EJPath = tTmp
    ogdFile.rows = ogdFile.FixedRows
    odtSearch.Value = Date
    dW_OriginalDate = Date
    
    oflEJ.Pattern = "*.EJ"    'หาแต่ File EJ
    olaPath.Caption = tW_EJPath
    otbForm(6).Text = tVB_CNTerminalNum
    otbForm(6).Enabled = False
    Call SP_OBJxSetFocus(otbForm(0))                    '*Ao 2014/06/04
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ockDate_Click()
    odtSearch.Enabled = (ockDate.Value = 1)
End Sub

Private Sub ocmBrwCstTel_Click()
    Dim tPath$
    If Trim(otbForm(0).Text) = "" Then
        ogdFile.Clear '*TON 58-12-11
        ogdFile.rows = ogdFile.FixedRows
        orbForm.Text = ""
        ogdFile.TextMatrix(0, 0) = "Date"
        ogdFile.TextMatrix(0, 1) = "Name"
    End If
    If Format$(dW_OriginalDate, "YYMMDD") <> Format$(odtSearch.Value, "YYMMDD") Then
        ogdFile.rows = ogdFile.FixedRows
        orbForm.Text = ""
        dW_OriginalDate = odtSearch.Value
    End If
    
    'ogdFile.rows = ogdFile.FixedRows   '*Bump 57-07-17 comment ให้ใช้ ด้านบนแทน
    'orbForm.Text = ""

    If ockDate.Value <> 1 Then  'ไม่แยกวันที่
        Call W_FINxPathEJ
    Else 'แยกวันที่
        tPath = Format$(odtSearch.Value, "YYMMDD")
        Call W_FINxPathEJ(tPath)
    End If
    
    With ogdFile
        .ColSort(.ColIndex("Date")) = flexSortGenericAscending
        .ColSort(.ColIndex("Name")) = flexSortGenericAscending
    End With
    If ogdFile.rows <= ogdFile.FixedRows Then
        Exit Sub
    Else
'        orbForm.LoadFile (ogdFile.TextMatrix(1, ogdFile.ColIndex("FileName")))
        If nW_FindRow = 0 Then
            ogdFile.Row = 1
            ogdFile.RowSel = 1
        Else
            ogdFile.Row = nW_FindRow
            ogdFile.RowSel = nW_FindRow
            nW_FindRow = 0  '* Reset
        End If
        'ogdFile.Row = 1
       ' ogdFile.RowSel = 1
    End If
    otbForm(0).SetFocus '*Bump 57-07-17
End Sub

Private Sub odiEJ_Change()
    oflEJ.Path = odiEJ.Path
End Sub

Private Sub odrEJ_Change()
    odiEJ.Path = odrEJ.Drive
End Sub

Private Sub oflEJ_Click()
    Dim tFileEj$, tTxt$
    
    tFileEj = oflEJ.Path & "\" & oflEJ.FileName
    If tFileEj = "" Then Exit Sub
    
    If Not (UCase(Right(tFileEj, 2)) = "EJ") Then Exit Sub
    
    
    Call SP_STAxMouse(Me, True)
    orbForm.Visible = True
    orbForm.Text = ""
    Open tFileEj For Input As #1
    Do While Not EOF(1)
        Line Input #1, tTxt
        If InStr(1, tTxt, "[Secret]") <> 0 Then orbForm.Text = "": Exit Do
        orbForm.Text = orbForm.Text & tTxt & vbCrLf
    Loop
    Close #1
    Call SP_STAxMouse(Me, False)
    Exit Sub

    Call SP_STAxMouse(Me, False)
End Sub

Private Sub ogdFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    orbForm.Text = ""
    If ogdFile.rows <= ogdFile.FixedRows Then Exit Sub
    
    orbForm.LoadFile (ogdFile.TextMatrix(ogdFile.Row, ogdFile.ColIndex("FileName")))
    ogdFile.SetFocus
End Sub

Private Sub otbForm_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call ocmBrwCstTel_Click
    End Select
End Sub

Private Sub otbForm_LostFocus(Index As Integer)
    Dim tTmp$
    
    Select Case Index
    Case 0
        If Trim(otbForm(Index).Text) = "" Then Exit Sub
        tTmp = Right$("00000" & Trim(otbForm(Index).Text), 5)
        otbForm(Index).Text = tTmp
    End Select
End Sub


