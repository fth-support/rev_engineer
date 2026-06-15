VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNViewSaleman 
   Caption         =   "Sale Man ID;klg_Title"
   ClientHeight    =   5400
   ClientLeft      =   4170
   ClientTop       =   5730
   ClientWidth     =   9570
   ControlBox      =   0   'False
   Icon            =   "wCNViewSaleman.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9570
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   4935
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   -240
      Width           =   9570
      _Version        =   720897
      _ExtentX        =   16880
      _ExtentY        =   8705
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VSFlex7Ctl.VSFlexGrid ogdForm 
         Height          =   4530
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   9345
         _cx             =   16484
         _cy             =   7990
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"wCNViewSaleman.frx":058A
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
         Editable        =   1
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
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   960
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   9570
      _Version        =   720897
      _ExtentX        =   16880
      _ExtentY        =   1693
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmChgAll 
         Caption         =   "&Change All;klg_ChgAll"
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
         Left            =   2400
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton ocmChgOne 
         Caption         =   "&Change;klg_ChgOne"
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
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   360
         Width           =   2055
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
         Left            =   7920
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNViewSaleman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nW_SeqNo As Long
Private tW_SalemanIDOld As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim orsFunc As ADODB.Recordset
        Dim tSql As String
        Dim tProc As String
        
On Error GoTo ErrHandle
        tSql = "SELECT FTSkbRef FROM TSysKeyboard   WHERE FNSkbKey='" & CStr(KeyCode) & "'"
        tSql = tSql & " AND FNSkbShift = '" & CStr(Shift) & "'"
        tSql = tSql & " AND FTSkbUsage ='Y'"
        tSql = tSql & " AND FTSkbGrp='F'"
        If SP_SQLvExecute(tSql, orsFunc, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
            If Not (orsFunc.BOF And orsFunc.EOF) Then
                tProc = SP_FEDtChkString(orsFunc, "FTSkbRef")
                If tProc <> "" Then
                    Call CallByName(wCNViewSaleman, tProc, VbMethod)
                End If
            End If
        End If
        Call SP_DATxRsNothing(orsFunc)
        Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsFunc)
End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxShwSalemanData
    tVB_SalemanIDOld = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tVB_SalemanIDChg = ""
End Sub

Private Sub ocmCancel_Click()
    Unload Me
End Sub

Private Sub ocmChgAll_Click()
    Dim tSql As String
    Dim tRow As String
    With ogdForm
         tRow = ogdForm.TextMatrix(ogdForm.Row, 0)
    End With
   If tRow <> "" Then
        tW_SalemanIDOld = ogdForm.TextMatrix(ogdForm.Row, 1)
        tVB_SalemanIDOld = tW_SalemanIDOld
        tVB_SalemanIDChg = ""
        wCNSaleManID.Show vbModal
        If bVB_SalemanIDChg Then
            tSql = "UPDATE " & tVB_CNTblDT
            tSql = tSql & " SET FTSdtSpnCode = '" & tVB_SalemanIDChg & "'"
            tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
            tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
            tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
            tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
            tSql = tSql & "  AND FTSdtSpnCode = '" & tW_SalemanIDOld & "'"
            If SP_SQLvExecute(tSql, , oDB.ocnOffLine) = 0 Then
                 Call W_FRMxShwSalemanData
            End If
        End If
    End If
End Sub

Private Sub ocmChgOne_Click()
    Dim tSql As String
    Dim tRow As String
    
    With ogdForm
         tRow = ogdForm.TextMatrix(ogdForm.Row, 0)
    End With
   If tRow <> "" Then
        tW_SalemanIDOld = ogdForm.TextMatrix(ogdForm.Row, 1)
        tVB_SalemanIDOld = tW_SalemanIDOld
        nW_SeqNo = ogdForm.TextMatrix(ogdForm.Row, 0)
        tVB_SalemanIDChg = ""
        wCNSaleManID.Show vbModal
        If bVB_SalemanIDChg Then
            tSql = "UPDATE " & tVB_CNTblDT
            tSql = tSql & " SET FTSdtSpnCode = '" & tVB_SalemanIDChg & "'"
            tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
            tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
            tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
            tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
            tSql = tSql & "  AND FNSdtSeqNo=" & nW_SeqNo
            If SP_SQLvExecute(tSql, , oDB.ocnOffLine) = 0 Then
                 Call W_FRMxShwSalemanData
            End If
        End If
    End If
End Sub

Private Sub W_FRMxShwSalemanData()
'----------------------------------------------------------
'   Call:
'   Cmt:    Show Salemain ID in grid
'*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
'----------------------------------------------------------

    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim nI As Long
On Error GoTo ErrHandle
    
    With ogdForm
        .FixedRows = 1
        .FixedCols = 0
        .rows = 1
        .Cols = 4
    End With
    
    tSql = "SELECT FNSdtSeqNo, FTSdtSpnCode,FTSdtBarCode,FTSkuAbbName"
    tSql = tSql & "  FROM " & tVB_CNTblDT
    tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"

    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine) Then Call SP_DATxRsNothing(orsTemp): Exit Sub
    Set ogdForm.DataSource = orsTemp
    ogdForm.Refresh
    For nI = 0 To orsTemp.Fields.Count - 1
        ogdForm.ColKey(nI) = orsTemp.Fields(nI).Name
    Next nI
    If ogdForm.rows < 11 Then ogdForm.rows = 11
    ogdForm.FixedCols = 1
    Call W_SetGridHeading
    ogdForm.Enabled = True
    ogdForm.HighLight = flexHighlightAlways
    ogdForm.Select 1, 1                        'move highlight to first row
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
End Sub

Public Sub W_SetGridHeading()
'----------------------------------------------------------
'   Call:
'   Cmt:   set heading of selling grid
'----------------------------------------------------------
    Const tFed = "FNSdtSeqNo, FTSdtSpnCode,FTSdtBarCode,FTSkuAbbName"
    Dim tColW$, iCol%

    tColW = "550;1995;3105;3105"
    iCol = 4
    
    With ogdForm
        Call SP_GRDxSetProperty(ogdForm, tFed, tColW, iCol, True, 0)
        .TextMatrix(.FixedRows - 1, .ColIndex("FNSdtSeqNo")) = UT_tToken(tCS_ColNo, ";", UT_nPosCutTag)                         'No
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSdtSpnCode")) = UT_tToken(tCS_ColSaleman, ";", UT_nPosCutTag)      'Select
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSdtBarCode")) = UT_tToken(tCS_ColBarCode, ";", UT_nPosCutTag)     'SKU Name
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSkuAbbName")) = UT_tToken(tCS_ColProductName, ";", UT_nPosCutTag)     'Product Name
    End With
    
End Sub

Public Sub W_KBSalemanChgOne()
'----------------------------------------------------------
'   Call:
'   Cmt:   Change One
'----------------------------------------------------------
    Call ocmChgOne_Click
End Sub

Public Sub W_KBSalemanChgAll()
'----------------------------------------------------------
'   Call:
'   Cmt:   Change All
'----------------------------------------------------------
    Call ocmChgAll_Click
End Sub

Private Sub ogdForm_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Dim nMaxR As Long
    With ogdForm
        nMaxR = .Aggregate(flexSTMax, .FixedRows, 0, .rows - 1, 0)
        If NewRowSel > nMaxR Then
            .Select .FixedRows, 0, .FixedRows, .Cols - 1
        End If
    End With
End Sub

'Private Sub ogdForm_KeyDown(KeyCode As Integer, Shift As Integer)
'        Dim orsFunc As ADODB.Recordset
'        Dim tSql As String
'        Dim tProc As String
'
'On Error GoTo ErrHandle
'        If KeyCode = 122 Then Stop
'        tSql = "SELECT FTSkbRef FROM TSysKeyboard   WHERE FNSkbKey='" & CStr(KeyCode) & "'"
'        tSql = tSql & " AND FNSkbShift = '" & CStr(Shift) & "'"
'        tSql = tSql & " AND FTSkbUsage ='Y'"
'        tSql = tSql & " AND FTSkbGrp='F'"
'        If SP_SQLvExecute(tSql, orsFunc, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
'            If Not orsFunc.EOF Then
'                tProc = SP_FEDtChkString(orsFunc, "FTSkbRef")
'                If tProc <> "" Then
'                    Call CallByName(wCNViewSaleman, tProc, VbMethod)
'                End If
'            End If
'        End If
'        Call SP_DATxRsNothing(orsFunc)
'        Exit Sub
'ErrHandle:
'    Call SP_DATxRsNothing(orsFunc)
'End Sub

Private Sub ogdForm_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub
