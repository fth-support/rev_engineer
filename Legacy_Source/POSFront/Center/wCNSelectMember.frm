VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNSelectMember 
   Caption         =   "Member Card;Klg_Title"
   ClientHeight    =   5385
   ClientLeft      =   3735
   ClientTop       =   3135
   ClientWidth     =   10470
   ControlBox      =   0   'False
   Icon            =   "wCNSelectMember.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   10470
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   4920
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   10500
      _Version        =   720897
      _ExtentX        =   18521
      _ExtentY        =   8678
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VSFlex7Ctl.VSFlexGrid ogdForm 
         Height          =   4530
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   10215
         _cx             =   18018
         _cy             =   7990
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   8
         GridLinesFixed  =   4
         GridLineWidth   =   1
         Rows            =   11
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wCNSelectMember.frx":058A
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
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   990
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4680
      Width           =   10500
      _Version        =   720897
      _ExtentX        =   18521
      _ExtentY        =   1746
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
         Height          =   510
         Left            =   8880
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
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7200
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   100
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNSelectMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public orsMember As ADODB.Recordset

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_GRDxShowData
    tVB_FindMember = ""
    Call SP_OBJxSetFocus(ogdForm)
   
End Sub

Private Sub W_GRDxShowData()

    Dim nI As Long
    Dim tSta As String
    Dim tSql As String
    
On Error GoTo ErrHandle
    With ogdForm
        .rows = 1
        .rows = 11
    End With
    If Not (orsMember.BOF And orsMember.EOF) Then
       nI = 1
       Do While Not orsMember.EOF
            With ogdForm
                If nI >= .rows Then .rows = .rows + 1
                .TextMatrix(nI, 0) = nI
                .TextMatrix(nI, 1) = SP_FEDtChkString(orsMember, "FTMcdName") & " " & SP_FEDtChkString(orsMember, "FTMcdSurName")
                .TextMatrix(nI, 2) = SP_FEDtChkString(orsMember, "FTMcdCode")
                .TextMatrix(nI, 3) = Format(SP_FEDdChkDate(orsMember, "FDMcdExpDate"), "dd/mm/yyyy")
                 tSta = UCase(SP_FEDtChkString(orsMember, "FTMcdSta"))
                 If tSta = "A" Then
                    .TextMatrix(nI, 4) = "Active"
                ElseIf tSta = "E" Then
                    .TextMatrix(nI, 4) = "Expire"
                ElseIf tSta = "L" Then
                    .TextMatrix(nI, 4) = "Lose"
                ElseIf tSta = "H" Then
                    .TextMatrix(nI, 4) = "Hold"
                End If
                .TextMatrix(nI, 5) = SP_FEDtChkString(orsMember, "FCMcdTotalPoint")
            End With
            nI = nI + 1
            orsMember.MoveNext
       Loop
    End If
    With ogdForm
         .Enabled = True
        .HighLight = flexHighlightAlways
        .Select 1, 1
    End With
    Call SP_DATxRsNothing(orsMember)
    Exit Sub
ErrHandle:
     Call SP_DATxRsNothing(orsMember)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)          '*AO51-03-17
    Select Case KeyCode
        Case 13: Call ocmOK_Click
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    tVB_FindMember = ""
    Unload Me
End Sub

Private Sub ocmOK_Click()
    tVB_FindMember = ogdForm.TextMatrix(ogdForm.Row, 2)
    Unload Me
End Sub

Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With ogdForm
        If .TextMatrix(NewRow, 0) = "" Then
            .Select OldRow, 1
        Else
            .TopRow = NewRow
        End If
    End With
End Sub

Private Sub ogdForm_DblClick()
    Call ocmOK_Click
End Sub
