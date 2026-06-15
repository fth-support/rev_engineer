VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNDepositMnu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposit Menu;klg_Deposit Menu"
   ClientHeight    =   6420
   ClientLeft      =   7950
   ClientTop       =   3885
   ClientWidth     =   7215
   ControlBox      =   0   'False
   Icon            =   "wCNDepositMnu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7215
   Begin VSFlex7Ctl.VSFlexGrid ogdFunc 
      Height          =   4995
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   7200
      _cx             =   12700
      _cy             =   8811
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   16768725
      ForeColorFixed  =   0
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNDepositMnu.frx":058A
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
      Left            =   5535
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   5745
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
      Left            =   3975
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   5745
      Width           =   1455
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1080
      Index           =   2
      Left            =   -5400
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5630
      Width           =   12600
      _Version        =   720897
      _ExtentX        =   22225
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   840
      Index           =   0
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   7200
      _Version        =   720897
      _ExtentX        =   12700
      _ExtentY        =   1482
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Menu;klg_Deposit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   0
         Width           =   6975
      End
   End
   Begin VB.Label olaForm 
      Alignment       =   2  'Center
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "wCNDepositMnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey As String
Private nMaxRow As Long

Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call:
'Cmt: ทำการแสดงรายการของ Deposit or Redeem
'-------------------------------------------------------
Dim nRow%
Dim tSmnID As String

With ogdFunc
    .rows = 11
    .TextMatrix(0, 1) = SP_GETtVocabularyTicketFromArray("170101", aVB_TransMnuAll, tSmnID) '"Deposit"
    .TextMatrix(0, 2) = tSmnID
    .TextMatrix(1, 1) = SP_GETtVocabularyTicketFromArray("170102", aVB_TransMnuAll, tSmnID)  '"Redeem"
    .TextMatrix(1, 2) = tSmnID
    .TextMatrix(2, 1) = SP_GETtVocabularyTicketFromArray("170103", aVB_TransMnuAll, tSmnID)  '"Deposit Cancel"
    .TextMatrix(2, 2) = tSmnID
    .TextMatrix(3, 1) = SP_GETtVocabularyTicketFromArray("170104", aVB_TransMnuAll, tSmnID)  '"Add Deposit"
    .TextMatrix(3, 2) = tSmnID
    .TextMatrix(4, 1) = SP_GETtVocabularyTicketFromArray("170105", aVB_TransMnuAll, tSmnID)  '"Add Deposit Cancel"
    .TextMatrix(4, 2) = tSmnID
    For nRow = 0 To 4
        If nRow + 1 < 10 Then
            .TextMatrix(nRow, 1) = " " & nRow + 1 & "." & .TextMatrix(nRow, 1)
            .TextMatrix(nRow, 0) = nRow + 1
        Else
            .TextMatrix(nRow, 1) = nRow + 1 & "." & .TextMatrix(nRow, 1)
            .TextMatrix(nRow, 0) = nRow + 1
        End If
    Next nRow
    .rows = 5
End With
End Sub

Private Sub Form_Activate()
On Error GoTo ErrHandle
    Call SP_OBJxSetFocus(ogdFunc)
Exit Sub
ErrHandle:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    Dim nMaxRow As Long
'    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
    If KeyAscii = 13 Then
        Call ocmOK_Click
    Else
        Select Case KeyAscii
            Case 48 To 58
                With ogdFunc
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
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FNCxAddFunc2Grid
    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)

End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = True
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    Unload Me
End Sub

Private Sub ocmOK_Click()
    '*KT 51-10-31
'PH 2.0.0 OLD
'-----------------------------------------------------------------------------------------
'    If ocnAll.ConnectionString = oDB.ocnOffLine.ConnectionString Then
'             If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
'                If Not SP_SETxInitialToOnline Then
'                        MsgBox "Terminal is Offline"
'                        Call ocmCancel_Click
'                        Exit Sub
'                End If
'            Else
'                MsgBox "Terminal is Offline"
'                Call ocmCancel_Click
'                Exit Sub
'            End If
'    End If
'-----------------------------------------------------------------------------------------
'    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
'        If Not SP_CHKbOnlineEnable Then
'                MsgBox "Terminal is Offline", vbCritical
'                Call ocmCancel_Click
'                Exit Sub
'        End If
'    Else
'        MsgBox "Terminal is Offline", vbCritical
'        Call ocmCancel_Click
'        Exit Sub
'    End If
    If ogdFunc.Row <> 0 Then
        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then         '*Ao 2013/09/25 LayawayCreate OFFLINE
            If Not SP_CHKbOnlineEnable Then
                MsgBox "Terminal is Offline", vbCritical
                Call ocmCancel_Click
                Exit Sub
            End If
        Else
            MsgBox "Terminal is Offline", vbCritical
            Call ocmCancel_Click
            Exit Sub
        End If
    End If
    
    Select Case ogdFunc.Row
        Case 0 'Deposit
            tVB_CNMnuCurrent = ogdFunc.TextMatrix(ogdFunc.Row, 2)
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
            bVB_TrnDepositRedeem = True
            tVB_Transaction = aVB_TransMode(5).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(5).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbDeposit - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbDeposit - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(5).PrintType
            Call SP_INSxTransHD(tVB_CNTblHD)
        Case 1 'Redeem
            tVB_CNMnuCurrent = ogdFunc.TextMatrix(ogdFunc.Row, 2)
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
            bVB_TrnDepositRedeem = True
            tVB_Transaction = aVB_TransMode(6).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(6).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbRedeem - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbRedeem - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(6).PrintType
            Call SP_INSxTransHD(tVB_CNTblHD)
        Case 2 'Deposit cancel
            tVB_CNMnuCurrent = ogdFunc.TextMatrix(ogdFunc.Row, 2)
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
            bVB_TrnDepositRedeem = True
            tVB_Transaction = aVB_TransMode(15).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(15).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbDepositCancel - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbDepositCancel - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(15).PrintType
            Call SP_INSxTransHD(tVB_CNTblHD)
        Case 3 'Add Deposit
            tVB_CNMnuCurrent = ogdFunc.TextMatrix(ogdFunc.Row, 2)
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
            bVB_TrnDepositRedeem = True
            tVB_Transaction = aVB_TransMode(18).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(18).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbAddDeposit - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbAddDeposit - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(18).PrintType
            Call SP_INSxTransHD(tVB_CNTblHD)
        Case 4 'Adaa Deposit Cancel
            tVB_CNMnuCurrent = ogdFunc.TextMatrix(ogdFunc.Row, 2)
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then Exit Sub
            bVB_TrnDepositRedeem = True
            tVB_Transaction = aVB_TransMode(19).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(19).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbAddDepositCancel - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbAddDepositCancel - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(19).PrintType
            Call SP_INSxTransHD(tVB_CNTblHD)
    End Select
'    Call wMain.W_SETxIndicator(tVB_TransactionSub, False)              '*Ao 2014/06/25
    Unload Me
End Sub

Private Sub ogdFunc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_HotKey = ""
End Sub

Private Sub ogdFunc_DblClick()
    Call ocmOK_Click
End Sub

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
On Error GoTo ErrHandle
    With ogdFunc
        If pnFuc > .rows Or pnFuc = 0 Then Exit Sub
         .Row = pnFuc - 1
         .Select .Row, 0, .Row, .Cols - 1
         .SetFocus
         .TopRow = .Row
    End With
    Exit Sub
ErrHandle:
    Exit Sub
End Sub


