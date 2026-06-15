VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNReprintSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprint Menu"
   ClientHeight    =   4845
   ClientLeft      =   4125
   ClientTop       =   3945
   ClientWidth     =   4650
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "wCNReprintSel.frx":0000
   ScaleHeight     =   4845
   ScaleWidth      =   4650
   Begin VB.CommandButton ocmCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel(Esc)"
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
      Left            =   3000
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   4200
      Width           =   1575
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
      Left            =   1560
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   4200
      Width           =   1335
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1080
      Index           =   2
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4095
      Width           =   4680
      _Version        =   720897
      _ExtentX        =   8255
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdFunc 
      Height          =   3540
      Left            =   0
      TabIndex        =   2
      Top             =   555
      Width           =   4680
      _cx             =   8255
      _cy             =   6244
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
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
      SelectionMode   =   1
      GridLines       =   8
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
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNReprintSel.frx":0342
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
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   840
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   4680
      _Version        =   720897
      _ExtentX        =   8255
      _ExtentY        =   1482
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reprint Menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   0
         Width           =   4215
      End
   End
End
Attribute VB_Name = "wCNReprintSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey As String
Private bW_KeyEnter As Boolean
Private nMaxRow As Long

Public tW_TransType As String
Public tW_TransNo As String
Public tW_PrintType As String

Private Function W_GETbCVIssue(ByVal ptTransType As String, ByVal ptTmnNo As String, ByVal ptTransNo As String) As Boolean
'-------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 2013/10/02
'-------------------------------------------------------------
Dim orsTemp As ADODB.Recordset
Dim tSql As String

On Error GoTo ErrHandle
    W_GETbCVIssue = False
    tSql = "SELECT FTCVNum,FCCVAmt,FCCVBalAmt,FDCVExpDate,FTCustContact"
    tSql = tSql & vbCrLf & "FROM TPSTCV"
    tSql = tSql & vbCrLf & "WHERE FTTransType = '" & ptTransType & "'"
    tSql = tSql & vbCrLf & "AND FTPOSNo = '" & ptTmnNo & "'"
    tSql = tSql & vbCrLf & "AND FTTransNo = '" & ptTransNo & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            W_GETbCVIssue = True
        End If
    End If
    Exit Function
ErrHandle:
    
End Function

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    'Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    'Call SP_FRMxSetCapForm(Me)
    Call W_FNCxAddFunc2Grid
End Sub

Private Sub W_FNCxAddFunc2Grid()
    Dim nI As Long
    With ogdFunc
        .TextMatrix(0, 0) = "1"
        .TextMatrix(0, 1) = " " & "1" & "." & "NORMAL"
        .TextMatrix(1, 0) = "2"
        .TextMatrix(1, 1) = " " & "2" & "." & "CREDIT VOUCHER"
        .TextMatrix(2, 0) = "3"
        .TextMatrix(2, 1) = " " & "3" & "." & "SALE ENDORSEMENT"
    End With
    nMaxRow = 3
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    Dim nMaxRow As Long
'    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
    If KeyAscii = 13 Then
        'Call ocmOk_Click
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

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
    Dim bFound As Boolean
    Dim nI As Long
On Error GoTo ErrHandle
    With ogdFunc
        bFound = False
         For nI = .FixedRows To .rows - 1
                If .TextMatrix(nI, 0) <> "" Then
                    If Val(Format(.TextMatrix(nI, 0))) = pnFuc Then
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

Private Sub ocmCancel_Click()
    Unload Me
End Sub

Private Sub ocmOK_Click()
Select Case ogdFunc.Row
    Case 0
        oTransPrn.TypePriner = tW_PrintType
        oTransPrn.PhysicalTmnNum = tVB_CNPhysicalTerminalNum
        Call oTransPrn.OpenTransaction(Val(tW_TransType), CDate(tUT_SaleDate), tVB_CNTerminalNum, tW_TransNo, False, , True) '*KT 52-03-26 ทำการแก้ไขให้การ Rerint สามารถทำการ Reprint ข้ามกะได้ แต่ข้ามเครื่องไม่ได้
        oTransPrn.RePrint = True
        oTransPrn.RePrintNormal = True
        oTransPrn.RePrintCV1 = False
        oTransPrn.RePrintCV2 = False
        oTransPrn.PrintEJ = False 'Comsheet ML- SF- 399
        oTransPrn.PrintOut
        oTransPrn.RePrint = False
        oTransPrn.RePrintNormal = False
        oTransPrn.PrintEJ = True 'Comsheet ML- SF- 399
    Case 1
        If W_GETbCVIssue(tW_TransType, tVB_CNTerminalNum, tW_TransNo) = False Then
            Exit Sub
        End If
        oTransPrn.TypePriner = tW_PrintType
        oTransPrn.PhysicalTmnNum = tVB_CNPhysicalTerminalNum
        Call oTransPrn.OpenTransaction(Val(tW_TransType), CDate(tUT_SaleDate), tVB_CNTerminalNum, tW_TransNo, False, , True) '*KT 52-03-26 ทำการแก้ไขให้การ Rerint สามารถทำการ Reprint ข้ามกะได้ แต่ข้ามเครื่องไม่ได้
        oTransPrn.RePrint = True
        oTransPrn.RePrintNormal = False
        oTransPrn.RePrintCV1 = True
        oTransPrn.RePrintCV2 = False
        oTransPrn.PrintEJ = False 'Comsheet ML- SF- 399
        oTransPrn.PrintOut
        oTransPrn.RePrint = False
        oTransPrn.RePrintCV1 = False
        oTransPrn.PrintEJ = True 'Comsheet ML- SF- 399
    Case 2
        Select Case SP_VALnText2Double(tW_TransType)            '*Ao 2013/10/01 CommSheet TK-ISS3000-367
            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax, tEN_TrnAddDepositCancel, tEN_TrnDepositCancel '*Eaw 56-10-14 เพิ่ม tEN_TrnAddDepositCancel, tEN_TrnDepositCancel
                Exit Sub
       End Select
        oTransPrn.TypePriner = tW_PrintType
        oTransPrn.PhysicalTmnNum = tVB_CNPhysicalTerminalNum
        Call oTransPrn.OpenTransaction(Val(tW_TransType), CDate(tUT_SaleDate), tVB_CNTerminalNum, tW_TransNo, False, , True) '*KT 52-03-26 ทำการแก้ไขให้การ Rerint สามารถทำการ Reprint ข้ามกะได้ แต่ข้ามเครื่องไม่ได้
        oTransPrn.RePrint = True
        oTransPrn.RePrintNormal = False
        oTransPrn.RePrintCV1 = False
        oTransPrn.RePrintCV2 = True
        oTransPrn.PrintEJ = False 'Comsheet ML- SF- 399
        oTransPrn.PrintOut
        oTransPrn.RePrint = False
        oTransPrn.RePrintCV2 = False
        oTransPrn.PrintEJ = True 'Comsheet ML- SF- 399
End Select
End Sub

Private Sub ogdFunc_Click()
    Call ocmOK_Click
End Sub

Private Sub ogdFunc_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call ocmOK_Click
    End Select
End Sub
