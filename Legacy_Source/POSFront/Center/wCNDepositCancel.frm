VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNDepositCancel 
   BorderStyle     =   0  'None
   Caption         =   "Cancel Add Layaway"
   ClientHeight    =   4410
   ClientLeft      =   8220
   ClientTop       =   3840
   ClientWidth     =   6780
   ControlBox      =   0   'False
   Icon            =   "wCNDepositCancel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl otcForm 
      Height          =   1335
      Index           =   1
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   -255
      Width           =   6765
      _Version        =   720897
      _ExtentX        =   11933
      _ExtentY        =   2355
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbDepositNo 
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   4
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Layaway Number:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   555
         Width           =   2280
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm 
      Height          =   3615
      Index           =   0
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   810
      Width           =   6765
      _Version        =   720897
      _ExtentX        =   11933
      _ExtentY        =   6376
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
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
         Left            =   5040
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton ocmOk 
         Caption         =   "&Ok"
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
         Left            =   3360
         TabIndex        =   5
         Top             =   2880
         Width           =   1455
      End
      Begin VSFlex7Ctl.VSFlexGrid ogdForm 
         Height          =   2325
         Left            =   255
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   6255
         _cx             =   11033
         _cy             =   4101
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16768725
         BackColorAlternate=   16768725
         GridColor       =   33023
         GridColorFixed  =   16768725
         TreeColor       =   16768725
         FloodColor      =   192
         SheetBorder     =   12615808
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   8
         GridLinesFixed  =   4
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   380
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wCNDepositCancel.frx":058A
         ScrollTrack     =   -1  'True
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
         WallPaperAlignment=   4
      End
   End
End
Attribute VB_Name = "wCNDepositCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tW_Sql As String
Public tW_Number As String

Private orsTemp As ADODB.Recordset
Private tW_HotKey As String
Private nMaxRow As Long

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
    Dim bFound As Boolean
    Dim nI As Long
    
    On Error GoTo ErrHandle
    With ogdForm
        bFound = False
        For nI = .FixedRows To .rows - 1
            If .TextMatrix(nI, 3) <> "" Then
                If CLng(Val(Format(.TextMatrix(nI, 3)))) = pnFuc Then
                    .Row = nI
                    bFound = True
                    Exit For
                End If
            End If
        Next nI
        tW_HotKey = ""
        If Not bFound Then Exit Sub
        .Select .Row, 0, .Row, .Cols - 1
        .SetFocus
        .TopRow = .Row
    End With
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_FNCxData2Grid(ByVal porsTemp As ADODB.Recordset)
'--------------------------------------------------------
'Call:
'Cmt:
'-------------------------------------------------------
Dim nRow As Integer
    With ogdForm
        porsTemp.Sort = "FCShdDepositBal DESC"
        porsTemp.MoveFirst
        .rows = 1
        nRow = 0
        Do While Not porsTemp.EOF
            nRow = nRow + 1
            .rows = nRow + 1
            .TextMatrix(nRow, 0) = SP_FEDtChkString(porsTemp, "FTTmnNum")
            .TextMatrix(nRow, 1) = SP_FEDtChkString(porsTemp, "FTShdTransNo")
            .TextMatrix(nRow, 2) = SP_FEDtChkString(porsTemp, "FDShdSysDate")
            .TextMatrix(nRow, 3) = nRow
            .TextMatrix(nRow, 4) = SP_FMTcCurPoint(SP_FEDvChkNumeric(porsTemp, "FCShdDepositAmt"), False, True)
            .TextMatrix(nRow, 5) = SP_FMTcCurPoint(SP_FEDvChkNumeric(porsTemp, "FCShdDepositBal"), False, True)
            porsTemp.MoveNext
        Loop
        .Row = 1
        .Select .Row, 0, .Row, .Cols - 1
'        .SetFocus
        .TopRow = 1
    End With
    
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
                         W_PRCxCallFunc (Val(tW_HotKey))
                    End If
                    If IsNumeric(tW_HotKey) Then
                        If CLng(tW_HotKey) >= nMaxRow Then tW_HotKey = ""
                    End If
                End With
        End Select
    End If
End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    If tW_Sql = "" Then Exit Sub
    
    If SP_SQLvExecute(tW_Sql, orsTemp, oDB.ocnOnLine) = 0 Then  'PH 2.0.0 'ต้องอ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            Call W_FNCxData2Grid(orsTemp)
            nMaxRow = ogdForm.Aggregate(flexSTCount, 0, 0, ogdForm.rows - 1, 0)
        End If
        otbDepositNo.Text = tW_Number
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    tVB_FTTmnNum = ""
    tVB_FTShdTransNo = ""
    dVB_CNDateAll = Format("01/01/2000", "dd/MM/yyyy")
    Unload Me
End Sub

Private Sub ocmOK_Click()
'    With ogdForm
'        If .Row < 1 Then Exit Sub
'        tVB_FTTmnNum = .TextMatrix(.Row, 0)
'        tVB_FTShdTransNo = .TextMatrix(.Row, 1)
'        dVB_CNDateAll = CDate(Format(.TextMatrix(.Row, 2), "dd/MM/yyyy"))
'    End With
'    If tVB_FTTmnNum <> "" Then
'        Call Form_Unload(0)
'    End If
    Call W_DEFxSELECT
End Sub

Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_HotKey = ""
End Sub

Private Sub W_DEFxSELECT()
    With ogdForm
        tVB_FTTmnNum = ""
        If .Row < 1 Then Exit Sub
        tVB_FTTmnNum = .TextMatrix(.Row, 0)
        tVB_FTShdTransNo = .TextMatrix(.Row, 1)
        dVB_CNDateAll = CDate(Format(.TextMatrix(.Row, 2), "dd/MM/yyyy"))
    End With
    If tVB_FTTmnNum <> "" Then
        Unload Me
    End If
End Sub

Private Sub ogdForm_DblClick()
'    With ogdForm
'        If .Row < 1 Then Exit Sub
'        tVB_FTTmnNum = .TextMatrix(.Row, 0)
'        tVB_FTShdTransNo = .TextMatrix(.Row, 1)
'        dVB_CNDateAll = CDate(Format(.TextMatrix(.Row, 2), "dd/MM/yyyy"))
'    End With
'    If tVB_FTTmnNum <> "" Then
'        Call Form_Unload(0)
'    End If
    Call W_DEFxSELECT
End Sub
