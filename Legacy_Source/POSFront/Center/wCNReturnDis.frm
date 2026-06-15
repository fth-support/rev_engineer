VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNReturnDis 
   BorderStyle     =   0  'None
   Caption         =   "Discount Type"
   ClientHeight    =   6645
   ClientLeft      =   7020
   ClientTop       =   4515
   ClientWidth     =   6270
   ControlBox      =   0   'False
   Icon            =   "wCNReturnDis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VSFlex7Ctl.VSFlexGrid ogdFunc 
      Height          =   5610
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6255
      _cx             =   11033
      _cy             =   9895
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   17.25
         Charset         =   0
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
      Rows            =   17
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNReturnDis.frx":058A
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
      Height          =   780
      Index           =   0
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   6255
      _Version        =   720897
      _ExtentX        =   11033
      _ExtentY        =   1376
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   17.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4560
         TabIndex        =   4
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   17.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4275
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   840
      Index           =   1
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5835
      Width           =   6255
      _Version        =   720897
      _ExtentX        =   11033
      _ExtentY        =   1482
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmDisCount 
         Caption         =   "&Enter"
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
         Left            =   2880
         TabIndex        =   2
         Top             =   280
         Width           =   1455
      End
      Begin VB.Timer otmForm 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   120
         Top             =   480
      End
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
         Left            =   4560
         TabIndex        =   3
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton ocmTotal 
         Caption         =   "&Total"
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
         Left            =   1200
         TabIndex        =   1
         Top             =   280
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNReturnDis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nW_SeqNo As Integer
Public tW_SKUCode As String
Public tW_ConsignmentFlag As String
Private tW_HotKey As String, tW_ScdDiscType As String
Private nW_DiscType As Integer
Const tC_Msg001 = "เกิดข้อผิดพลาดในการใส่ข้อมูลส่วนลด;You can not add discount exceed the price."
Const tC_Msg003 = "เกิดข้อผิดพลาดในการเพิ่มข้อมูลส่วนลด;There are some error occer on add discount data."

Private Sub W_DEFxFindDisc()
'__________________________
'   Call:
'   Cmt:
'*KT 53-01-26 Communication Sheet PH3-ML-SB-073 Deal & Free  in Return Transaction.xls
'__________________________
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim nDctNo As Integer
Dim cScdAmt As Double
On Error GoTo ErrHandler
    With ogdFunc
        tSql = "SELECT SUM(FCScdAmt) AS FCScdAmt,FNDctNo,FTScdDiscType FROM " & tVB_CNTblCD
        tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
        tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nW_SeqNo
        tSql = tSql & vbCrLf & " GROUP BY FNDctNo,FTScdDiscType"
        If SP_SQLbExecute(tSql, orsTemp) Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                Do While Not orsTemp.EOF
                nDctNo = SP_FEDvChkNumeric(orsTemp, "FNDctNo")
                cScdAmt = SP_FEDvChkNumeric(orsTemp, "FCScdAmt")
                Select Case nDctNo
                    Case nEN_Disc01Deal
                        If SP_FEDtChkString(orsTemp, "FTScdDiscType") = "4" Then
                            .TextMatrix(16, 2) = SP_FMTcCurPoint(cScdAmt, False, True)
                        Else
                            .TextMatrix(nDctNo - 1, 2) = SP_FMTcCurPoint(cScdAmt, False, True)
                        End If
                    Case nEN_Disc21MPrice
                        .TextMatrix(15, 2) = SP_FMTcCurPoint(cScdAmt, False, True)
                    Case nEN_Disc22MDisc
                        .TextMatrix(15, 2) = SP_FMTcCurPoint(cScdAmt, False, True)
                    Case Else
                        .TextMatrix(nDctNo - 1, 2) = SP_FMTcCurPoint(cScdAmt, False, True)
                End Select
                orsTemp.MoveNext
                Loop
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Set orsTemp = Nothing
    Call SP_MSGShowError
End Sub

Private Sub W_DEFcSumMallandCredit(ByVal pnDiscType As Integer)
'__________________________________
'   Call:
'   Cmt:
'*KT 51-11-01
'________________________
Dim nLoop As Integer
Dim cDisAmt As Double
    Select Case pnDiscType
        Case nEN_Disc11MallCrd
            For nLoop = 0 To UBound(aVB_CNMallCard) - 1
                cDisAmt = cDisAmt + aVB_CNMallCard(nLoop).Amt
            Next nLoop
            cVB_CNDiscAmt = cDisAmt
        Case nEN_Disc12EmpCrd
            For nLoop = 0 To UBound(aVB_CNEmpCard) - 1
                cDisAmt = cDisAmt + aVB_CNEmpCard(nLoop).Amt
            Next nLoop
            cVB_CNDiscAmt = cDisAmt
        Case nEN_Disc15Ontop
            nLoop = 0
            For nLoop = 0 To UBound(aVB_CNOnTop)
                cDisAmt = cDisAmt + aVB_CNOnTop(nLoop).Amt
            Next nLoop
            cVB_CNDiscAmt = cDisAmt
    End Select
End Sub


Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call:
'Cmt: แสดงรายการส่วนลดประเภทต่างๆ
'*KT 53-01-26 Communication Sheet PH3-ML-SB-073 Deal & Free  in Return Transaction.xls
'-------------------------------------------------------
Dim nI As Integer, nR As Integer
Dim orsTemp As ADODB.Recordset
Dim tSql As String
On Error GoTo ErrHandle
    With ogdFunc
         '*KT 52-09-19 Member Price + Consignment + Return
        tSql = "SELECT * FROM TCNMDiscountType"
        tSql = tSql & vbCrLf & "WHERE (FNDctNo <> 22 OR FNDctNo <> 16) "
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            nR = 0
            .rows = orsTemp.RecordCount + 1
            Do While Not orsTemp.EOF
                .TextMatrix(nR, 0) = nR + 1 'SP_FEDvChkNumeric(orsTemp, "FNDctNo")
                .TextMatrix(nR, 1) = .TextMatrix(nR, 0) & "." & SP_FEDtChkString(orsTemp, "FTDctDiscType")
                .TextMatrix(nR, 2) = SP_FMTcCurPoint(0, False, True)
                .TextMatrix(nR, 3) = SP_FEDvChkNumeric(orsTemp, "FNDctNo")
                nR = nR + 1
                orsTemp.MoveNext
            Loop
            
            'Specail List
            'แก้ไข wording Deal&Free เป็น Deal
            .TextMatrix(0, 1) = .TextMatrix(0, 0) & "." & "Deal"
            'เพิ่มแถว Free Item
            .TextMatrix(nR, 0) = nR + 1
            .TextMatrix(nR, 1) = .TextMatrix(nR, 0) & "." & "Free Item"
            .TextMatrix(nR, 2) = .TextMatrix(0, 2)
            .TextMatrix(nR, 3) = .TextMatrix(0, 3)
            .ColAlignment(2) = flexAlignRightCenter
            
        End If
    End With
    Exit Sub
ErrHandle:
    Set orsTemp = Nothing
    Call SP_MSGShowError
End Sub

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
Dim bFound As Boolean
Dim nI As Long
On Error GoTo ErrHandle
    With ogdFunc
        bFound = False
         For nI = .FixedRows To .rows - 1
                If .TextMatrix(nI, 0) <> "" Then
                    If CLng(Val(Format(.TextMatrix(nI, 0)))) = pnFuc Then
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyHome
                Call ocmTotal_Click
    End Select
End Sub

Private Sub Form_Load()
    tW_HotKey = ""
    Call SP_FRMxChildCenter(Me)
'    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
'    Call SP_FRMxSetCapForm(Me)
    Call W_FNCxAddFunc2Grid
    Call W_DEFxFindDisc
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim nMaxRow As Long
    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
    If KeyAscii = 13 Then
        Call ocmDisCount_Click
    Else
        Select Case KeyAscii
            Case 48 To 58
                otmForm.Enabled = True
                With ogdFunc
                    tW_HotKey = tW_HotKey & Chr(KeyAscii)
                    If tW_HotKey <> "" Then
                         W_PRCxCallFunc (Val(tW_HotKey))
                    End If
                    If IsNumeric(tW_HotKey) Then
                        If CLng(tW_HotKey) >= nMaxRow Then tW_HotKey = "": otmForm.Enabled = False
                    End If
                End With
            Case vbKeyHome
                Call ocmTotal_Click
            Case vbKeyEscape
                
        End Select
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    otmForm.Enabled = False
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    Unload Me
End Sub

Private Sub ocmDisCount_Click()
'*KT 53-01-26 Communication Sheet PH3-ML-SB-073 Deal & Free  in Return Transaction.xls
    Dim nRow As Integer
    With ogdFunc
        bVB_CNReturnOk = True
        nRow = .TextMatrix(.Row, 0)
        nW_DiscType = .TextMatrix(.Row, 3)
        '*KT 54-09-13 V.4.0.1 Req3 Phase4.0(Member Point) Member Point on Other Receive หน้าจอ Discount สำหรับ Transaction Other Disbursement  ให้ทำรายการคืน Point ได้อย่างเดียว
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnOtherDisbursement
                If nW_DiscType <> nEN_Disc11MallCrd Then
                    Call SP_MSGnShowing("tms_CN005107", nCS_Information)
                    Exit Sub
                End If
        End Select
    
        Select Case nW_DiscType
            Case nEN_Disc01Deal
                cVB_CNDiscAmt = CDbl(Val(Format(.TextMatrix(.Row, 2))))
                If nRow <> "1" Then
                    'Free Item
                    wCNReturnAmt.bW_IsFreeItem = True
                    wCNReturnAmt.Show vbModal
                Else
                    'Deal
                    wCNReturnAmt.bW_IsFreeItem = False
                    wCNReturnAmt.Show vbModal
                End If
            Case nEN_Disc11MallCrd
                wCNReturnMall.nW_SeqNo = nW_SeqNo
                wCNReturnMall.cW_ChkPrice = SP_FMTcCurPoint(.TextMatrix(.Row, 2), False, True)
                wCNReturnMall.Show vbModal
                If bVB_CNReturnOk Then
                    Call W_DEFcSumMallandCredit(nW_DiscType)
                End If
            Case nEN_Disc12EmpCrd
                wCNReturnEmp.nW_SeqNo = nW_SeqNo
                wCNReturnEmp.cW_ChkPrice = SP_FMTcCurPoint(.TextMatrix(.Row, 2), False, True)
                wCNReturnEmp.Show vbModal
                If bVB_CNReturnOk Then
                    Call W_DEFcSumMallandCredit(nW_DiscType)
                End If
            Case nEN_Disc15Ontop
                wCNReturnCredit.nW_SeqNo = nW_SeqNo
                wCNReturnCredit.cW_ChkPrice = SP_FMTcCurPoint(.TextMatrix(.Row, 2), False, True)
                wCNReturnCredit.Show vbModal
                If bVB_CNReturnOk Then
                    Call W_DEFcSumMallandCredit(nW_DiscType)
                End If
            Case nEN_Disc21MPrice, nEN_Disc22MDisc '*KT 52-09-19 Member Price + Consignment + Return
                cVB_CNDiscAmt = CDbl(Val(Format(.TextMatrix(.Row, 2))))
                wCNReturnMBPrice.nW_SeqNo = nW_SeqNo
                wCNReturnMBPrice.Show vbModal
            Case Else
                cVB_CNDiscAmt = CDbl(Val(Format(.TextMatrix(.Row, 2))))
                wCNReturnAmt.bW_IsFreeItem = False
                wCNReturnAmt.Show vbModal
        End Select
        Call W_DEFxDisType(nRow)
    End With
End Sub
Private Sub ocmTotal_Click()
    Call W_DATxInsertDiscCD
    'Unload Me
End Sub
Private Sub ogdFunc_DblClick()
    Call ocmDisCount_Click
End Sub

Private Sub ogdFunc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyHome
                Call ocmTotal_Click
    End Select
End Sub

Private Sub otmForm_Timer()
    Dim nMaxRow As Long
    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
    If IsNumeric(tW_HotKey) Then tW_HotKey = "": otmForm.Enabled = False
End Sub
Private Sub W_DEFxDisType(ByVal pnDiscType As Integer)
    With ogdFunc
                .TextMatrix(pnDiscType - 1, 2) = SP_FMTcCurPoint(cVB_CNDiscAmt, False, True)
    End With
End Sub

Private Sub W_DATxGetDataDT(ByRef porsTemp As ADODB.Recordset)
'--------------------------------------------------
 'Get data from TPSTSalDT and collect in parameter
 '--------------------------------------------------
    Dim tSql As String

    On Error GoTo ErrHandler
    tSql = "SELECT FCSdtQtyAllDisAvi,FCSdtSaleAmt,FCSdtQty,FCSdtRegPrice"
    tSql = tSql & vbCrLf & " FROM " & tVB_CNTblDT
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nW_SeqNo
    tSql = tSql & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
    If Not SP_SQLbExecute(tSql, porsTemp) Then
        Set porsTemp = Nothing
    End If
    Exit Sub
ErrHandler:
    Set porsTemp = Nothing
    Call SP_MSGShowError
End Sub
Private Sub W_DEFxDeleteCD()
'_____________________________
'   Call:
'   Cmt:
'_____________________________
Dim tSql As String
Dim nEff As Long
On Error GoTo ErrHandler

    tSql = "DELETE FROM " & tVB_CNTblCD
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & nW_SeqNo
    Call SP_SQLbExecute(tSql, , , nEff)
    Exit Sub
ErrHandler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub
Private Sub W_DEFcSumDisc(ByRef pcSumDisc As Double)
Dim nLoop As Integer
    pcSumDisc = 0
    With ogdFunc
        For nLoop = .FixedRows To .rows - 1
            pcSumDisc = pcSumDisc + SP_FMTcCurPoint(.TextMatrix(nLoop, 2), True, True)
        Next nLoop
    End With
End Sub

Private Sub W_DATxInsertDiscCD()
'--------------------------------------------------
 'Insert Discount data to TPST  SalCD
 '*KT 53-01-26 Communication Sheet PH3-ML-SB-073 Deal & Free  in Return Transaction.xls
 '--------------------------------------------------
 
    Dim nLoop As Integer, nCard As Integer, nDiscType As Integer
    Dim cDisAmt As Double, cSdtQtyAllDisAvi As Double, cDisQty As Double, cSdtSaleAmt As Double
    Dim orsTemp As ADODB.Recordset
    Dim tEventID As String, tCardType As String, tScdCardID As String
    Dim cSumDisc As Double
    Dim bMemberPrice As Boolean

On Error GoTo ErrHandler
    bMemberPrice = True
    Call W_DATxGetDataDT(orsTemp)
    If (orsTemp.BOF And orsTemp.EOF) Then Exit Sub
'    cSdtQtyAllDisAvi = SP_FEDvChkNumeric(orsTemp, "FCSdtQtyAllDisAvi")
    'DELETE CD ทิ้งก่อน
    
    cSdtSaleAmt = SP_FEDvChkNumeric(orsTemp, "FCSdtSaleAmt")
    Call W_DEFcSumDisc(cSumDisc)
    If cSumDisc > cSdtSaleAmt And cSdtSaleAmt > 0 Then
        Call SP_MSGnShowing(tC_Msg001, tCS_CNMsgWarn)
        Exit Sub
    End If
    Call W_DEFxDeleteCD
    cSdtQtyAllDisAvi = 0
    cDisQty = SP_FEDvChkNumeric(orsTemp, "FCSdtQty")
    tW_ScdDiscType = CStr(nEN_LnDis2Amt)
    
    With ogdFunc
        For nLoop = .FixedRows To .rows - 1
            If .TextMatrix(nLoop, 2) > 0 Then
                cDisAmt = SP_FMTcCurPoint(.TextMatrix(nLoop, 2), True, True)
                tEventID = ""
                nDiscType = Trim(.TextMatrix(nLoop, 3))
                Select Case nDiscType 'Trim(.TextMatrix(nLoop, 3))
                    Case nEN_Disc11MallCrd
                        If bMemberPrice Then
                            cDisAmt = SP_FMTcCurPoint(.TextMatrix(15, 2), True, True)
                            tEventID = ""
                            nDiscType = Trim(.TextMatrix(15, 3))
                            If UCase(tW_ConsignmentFlag) = "Y" Then
                                nDiscType = nEN_Disc22MDisc
                            Else
                                nDiscType = nEN_Disc21MPrice
                            End If
                            tScdCardID = xVB_CNMemberP.ID
                            tCardType = xVB_CNMemberP.MallCardType
                            tVB_StaMemPriceAlw = "Y"
                            tVB_StaMemPricePromoAlw = False
                            nVB_CNMemPriceMaxQty = 99999999
                            tVB_SKUConsignFlag = tW_ConsignmentFlag
                            tVB_CNMemPriceType = tW_ScdDiscType
                            cVB_CNMemPriceAmt = cDisAmt
                            cUT_Qty = cDisQty
                            tVB_CNEventMemPrice = tEventID
                            cVB_CNMemPriceAmtDis = cDisAmt
                            cUT_RegPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")
                            cUT_MemberPrice = SP_FMTcCurPoint(cUT_RegPrice - (cDisAmt / cUT_Qty), True, False)
                            If W_PRCbUpdMemberPrice(nW_SeqNo, tVB_CNEventMemPrice, cUT_MemberPrice) Then
                                 Call wFunctionKB.W_PRCxMemberPrice(nW_SeqNo, tW_SKUCode, 0)
                            End If
                        End If
                        bMemberPrice = False
                        cDisAmt = SP_FMTcCurPoint(.TextMatrix(nLoop, 2), True, True)
                        tEventID = ""
                        nDiscType = Trim(.TextMatrix(nLoop, 3))
                        For nCard = 0 To UBound(aVB_CNMallCard) - 1
                            If Not (Trim(aVB_CNMallCard(nCard).IDCard) = "") Then
                                tScdCardID = aVB_CNMallCard(nCard).IDCard
                                cDisAmt = aVB_CNMallCard(nCard).Amt
                                tCardType = CStr(nCard + 1)
                                Call W_DATxInsertCD(cDisAmt, nW_SeqNo, tEventID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCardType, nDiscType)
                            End If
                        Next nCard
                    Case nEN_Disc12EmpCrd
                        If bMemberPrice Then
                            cDisAmt = SP_FMTcCurPoint(.TextMatrix(15, 2), True, True)
                            tEventID = ""
                            nDiscType = Trim(.TextMatrix(15, 3))
                            If UCase(tW_ConsignmentFlag) = "Y" Then
                                nDiscType = nEN_Disc22MDisc
                            Else
                                nDiscType = nEN_Disc21MPrice
                            End If
                            tScdCardID = xVB_CNMemberP.ID
                            tCardType = xVB_CNMemberP.MallCardType
                            tVB_StaMemPriceAlw = "Y"
                            tVB_StaMemPricePromoAlw = False
                            nVB_CNMemPriceMaxQty = 99999999
                            tVB_SKUConsignFlag = tW_ConsignmentFlag
                            tVB_CNMemPriceType = tW_ScdDiscType
                            cVB_CNMemPriceAmt = cDisAmt
                            cUT_Qty = cDisQty
                            tVB_CNEventMemPrice = tEventID
                            cVB_CNMemPriceAmtDis = cDisAmt
                            cUT_RegPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")
                            cUT_MemberPrice = SP_FMTcCurPoint(cUT_RegPrice - (cDisAmt / cUT_Qty), True, False)
                            If W_PRCbUpdMemberPrice(nW_SeqNo, tVB_CNEventMemPrice, cUT_MemberPrice) Then
                                 Call wFunctionKB.W_PRCxMemberPrice(nW_SeqNo, tW_SKUCode, 0)
                            End If
                        End If
                        bMemberPrice = False
                        cDisAmt = SP_FMTcCurPoint(.TextMatrix(nLoop, 2), True, True)
                        tEventID = ""
                        nDiscType = Trim(.TextMatrix(nLoop, 3))
                        For nCard = 0 To UBound(aVB_CNEmpCard) - 1
                            If Not (Trim(aVB_CNEmpCard(nCard).IDCard) = "") Then
                                tScdCardID = aVB_CNEmpCard(nCard).IDCard
                                cDisAmt = aVB_CNEmpCard(nCard).Amt
                                tCardType = CStr(nCard + 1)
                                tVB_DiscEmpChkLimit = W_GETtEmpChkLimit(tScdCardID)
                                Call W_DATxInsertCD(cDisAmt, nW_SeqNo, tEventID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCardType, nDiscType)
                            End If
                        Next nCard
                    Case nEN_Disc15Ontop
                        If bMemberPrice Then
                            cDisAmt = SP_FMTcCurPoint(.TextMatrix(15, 2), True, True)
                            tEventID = ""
                            nDiscType = Trim(.TextMatrix(15, 3))
                            If UCase(tW_ConsignmentFlag) = "Y" Then
                                nDiscType = nEN_Disc22MDisc
                            Else
                                nDiscType = nEN_Disc21MPrice
                            End If
                            tScdCardID = xVB_CNMemberP.ID
                            tCardType = xVB_CNMemberP.MallCardType
                            tVB_StaMemPriceAlw = "Y"
                            tVB_StaMemPricePromoAlw = False
                            nVB_CNMemPriceMaxQty = 99999999
                            tVB_SKUConsignFlag = tW_ConsignmentFlag
                            tVB_CNMemPriceType = tW_ScdDiscType
                            cVB_CNMemPriceAmt = cDisAmt
                            cUT_Qty = cDisQty
                            tVB_CNEventMemPrice = tEventID
                            cVB_CNMemPriceAmtDis = cDisAmt
                            cUT_RegPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")
                            cUT_MemberPrice = SP_FMTcCurPoint(cUT_RegPrice - (cDisAmt / cUT_Qty), True, False)
                            If W_PRCbUpdMemberPrice(nW_SeqNo, tVB_CNEventMemPrice, cUT_MemberPrice) Then
                                 Call wFunctionKB.W_PRCxMemberPrice(nW_SeqNo, tW_SKUCode, 0)
                            End If
                        End If
                        bMemberPrice = False
                        cDisAmt = SP_FMTcCurPoint(.TextMatrix(nLoop, 2), True, True)
                        tEventID = ""
                        nDiscType = Trim(.TextMatrix(nLoop, 3))
                        For nCard = 0 To UBound(aVB_CNOnTop)
                            If Not (Trim(aVB_CNOnTop(nCard).IDCard) = "") Then
                                tScdCardID = aVB_CNOnTop(nCard).IDCard
                                cDisAmt = aVB_CNOnTop(nCard).Amt
                                tCardType = CStr(nCard + 1)
                                Call W_DATxInsertCD(cDisAmt, nW_SeqNo, tEventID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCardType, nDiscType)
                            End If
                        Next nCard
                    Case nEN_Disc21MPrice, nEN_Disc22MDisc '*KT 52-09-19 Member Price + Consignment + Return
                        If bMemberPrice Then
                            If UCase(tW_ConsignmentFlag) = "Y" Then
                                nDiscType = nEN_Disc22MDisc
                            Else
                                nDiscType = nEN_Disc21MPrice
                            End If
                            tScdCardID = xVB_CNMemberP.ID
                            tCardType = xVB_CNMemberP.MallCardType
                            tVB_StaMemPriceAlw = "Y"
                            tVB_StaMemPricePromoAlw = False
                            nVB_CNMemPriceMaxQty = 99999999
                            tVB_SKUConsignFlag = tW_ConsignmentFlag
                            tVB_CNMemPriceType = tW_ScdDiscType
                            cVB_CNMemPriceAmt = cDisAmt
                            cUT_Qty = cDisQty
                            tVB_CNEventMemPrice = tEventID
                            cVB_CNMemPriceAmtDis = cDisAmt
                            cUT_RegPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")
                            cUT_MemberPrice = SP_FMTcCurPoint(cUT_RegPrice - (cDisAmt / cUT_Qty), True, False)
                            If W_PRCbUpdMemberPrice(nW_SeqNo, tVB_CNEventMemPrice, cUT_MemberPrice) Then
                                 Call wFunctionKB.W_PRCxMemberPrice(nW_SeqNo, tW_SKUCode, 0)
                            End If
                        End If
                        bMemberPrice = False
                    Case Else
                         Select Case nDiscType
                            Case nEN_Disc01Deal
                                If nLoop <> 0 Then
                                     'Free Item
                                     tW_ScdDiscType = "4" ' เดิมเป็น 2 ตลอด
                                     Call W_DATxInsertCD(cDisAmt, nW_SeqNo, tEventID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCardType, nDiscType, "7")
                                Else
                                    'Deal
                                     Call W_DATxInsertCD(cDisAmt, nW_SeqNo, tEventID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCardType, nDiscType, "6")
                                End If
                            Case nEN_Disc02PmtTime, nEN_Disc03Pmt
                                Call W_DATxInsertCD(cDisAmt, nW_SeqNo, tEventID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCardType, nDiscType)
                            Case Else
                                If bMemberPrice Then
                                    cDisAmt = SP_FMTcCurPoint(.TextMatrix(15, 2), True, True)
                                    tEventID = ""
                                    nDiscType = Trim(.TextMatrix(15, 3))
                                    If UCase(tW_ConsignmentFlag) = "Y" Then
                                        nDiscType = nEN_Disc22MDisc
                                    Else
                                        nDiscType = nEN_Disc21MPrice
                                    End If
                                    tScdCardID = xVB_CNMemberP.ID
                                    tCardType = xVB_CNMemberP.MallCardType
                                    tVB_StaMemPriceAlw = "Y"
                                    tVB_StaMemPricePromoAlw = False
                                    nVB_CNMemPriceMaxQty = 99999999
                                    tVB_SKUConsignFlag = tW_ConsignmentFlag
                                    tVB_CNMemPriceType = tW_ScdDiscType
                                    cVB_CNMemPriceAmt = cDisAmt
                                    cUT_Qty = cDisQty
                                    tVB_CNEventMemPrice = tEventID
                                    cVB_CNMemPriceAmtDis = cDisAmt
                                    cUT_RegPrice = SP_FEDvChkNumeric(orsTemp, "FCSdtRegPrice")
                                    cUT_MemberPrice = SP_FMTcCurPoint(cUT_RegPrice - (cDisAmt / cUT_Qty), True, False)
                                    If W_PRCbUpdMemberPrice(nW_SeqNo, tVB_CNEventMemPrice, cUT_MemberPrice) Then
                                         Call wFunctionKB.W_PRCxMemberPrice(nW_SeqNo, tW_SKUCode, 0)
                                    End If
                                End If
                                tScdCardID = ""
                                tCardType = ""
                                bMemberPrice = False
                                cDisAmt = SP_FMTcCurPoint(.TextMatrix(nLoop, 2), True, True)
                                tEventID = ""
                                nDiscType = Trim(.TextMatrix(nLoop, 3))
                                Call W_DATxInsertCD(cDisAmt, nW_SeqNo, tEventID, cDisQty, cSdtSaleAmt, cSdtQtyAllDisAvi, tScdCardID, tCardType, nDiscType)
                        End Select
                End Select
            End If
        Next nLoop
    End With
    Unload Me
    Exit Sub
ErrHandler:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function W_PRCbUpdMemberPrice(ByVal pnSeqNo As Long, ByVal ptEventID As String, ByVal pcMemberPrice As Double)
'------------------------------------------------------------------------------------------------------------------------
'Ret:บันทึกข้อมูลสำเร็จ True บันทึกข้อมูลไม่สำเร็จ  False
'Cmt: ทำการ update ราคา Member Price ,Event ID ที่ DT FCSdtMemberPrice,FTSdtEventMemPrice ก่อน
'-----------------------------------------------------------------------------------------------------------------------
    Dim tSql As String
    Dim nEff As Long
On Error GoTo ErrHandle
    nEff = 0
    tSql = "UPDATE " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "SET FCSdtMemberPrice = " & pcMemberPrice
    tSql = tSql & vbCrLf & ",FTSdtEventMemPrice = '" & ptEventID & "'"
    tSql = tSql & vbCrLf & " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & " AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND FNSdtSeqNo = " & pnSeqNo
    tSql = tSql & vbCrLf & " AND ISNULL(FTSdtCancelFlag,'0') <> '1'"
    If SP_SQLvExecute(tSql, , , , , nEff) = 0 Then
        If nEff > 0 Then
            W_PRCbUpdMemberPrice = True
        Else
            W_PRCbUpdMemberPrice = False
        End If
    Else
        W_PRCbUpdMemberPrice = False
    End If
    Exit Function
ErrHandle:
    W_PRCbUpdMemberPrice = False
End Function

Private Sub W_DATxInsertCD(ByVal pcDisAmt#, ByVal pnFNSdtSeqNo%, ByVal ptEventID$, ByVal pcDisQty#, ByVal pcAlwDis# _
, ByVal pcAlwDisQty#, ByVal ptScdCardID$, ByVal ptCardType$, ByVal pnDiscType%, Optional ByVal ptStaChkDepart$ = "")
'--------------------------------------------------
 'Insert Discount data to TPST  SalCD
 '*KT 53-01-26 Communication Sheet PH3-ML-SB-073 Deal & Free  in Return Transaction.xls
 '--------------------------------------------------
Dim tSql As String
Dim tField As String
Dim nMaxSeqNo As Integer
Dim tWhe As String
Dim nEffect As Long
Dim TWheMax As String
Dim nLoop As Integer, nMall As Integer, nTop As Integer
Dim cDisPmt As Double
Dim nScdOrder As Integer, nActSeqNo As Integer

On Error GoTo ErrHandle
   
    nScdOrder = W_GETnGroup(pnDiscType)
    nActSeqNo = W_DISnGetNextAct()
    tField = "FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FNScdSeqNo"
    tField = tField & ",FTSkuCode,FTShdTransType,FDShdTransDate,FTSdtBarCode,FTSdtStkCode"
    tField = tField & ",FCSdtStkFac,FTSkuDepCode,FTSkuBrdCode,FTPunCode,FCScdQtyB4"
    tField = tField & ",FCScdQty,FCScdQtyAF,FCScdAmtB4,FCScdAmtB4Avi,FCScdAmt"
    tField = tField & ",FCScdAmtAF,FCScdAmtAFAvi,FCScdDiscPer,FCScdDiscAmt,FCScdDiscNew"
    tField = tField & ",FTScdDiscType,FNDctNo,FTDctName,FTDctRefCode,FTScdEventID"
    tField = tField & ",FTScdCardID,FDScdExpDate,FTScdCardType,FTScdStaTrans"
    tField = tField & ",FTScdCstID,FTScdCardKey,FTScdStaChkExpire,FTScdStaChkDepart,FNScdSptSeqNo,FCScdMaxAlwDisc"
    tField = tField & ",FNScdActSeqNo"
    tField = tField & ",FNScdOrder,FCScdAmtTime,FCScdAmtPmt,FNScdDealGroup"
    tField = tField & ",FCScdDiscPerOrg,FCScdDiscAmtOrg,FCScdDiscNewOrg"
    
    tWhe = " WHERE FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " And  FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " And FDShdTransDate = '" & tUT_SaleDate & "'"
    TWheMax = tWhe
    tWhe = tWhe & " And FNSdtSeqNo = " & pnFNSdtSeqNo
    
    nMaxSeqNo = SP_VALcText2Double(SP_SQLtFunction("MAX", "FNScdSeqNo", "" & tVB_CNTblCD & "", TWheMax))
    nMaxSeqNo = nMaxSeqNo + 1
                
    tSql = "INSERT INTO " & tVB_CNTblCD & "(" & tField & SP_GETtWhoInsertField
    tSql = tSql & vbCrLf & " SELECT "
    tSql = tSql & vbCrLf & "FTStmCode,FTTmnNum, FTShdTransNo, FNSdtSeqNo"
    tSql = tSql & vbCrLf & "," & nMaxSeqNo & " AS FNScdSeqNo "
    tSql = tSql & vbCrLf & ",FTSkuCode,FTShdTransType,FDShdTransDate,FTSdtBarCode,FTSdtStkCode"
    tSql = tSql & vbCrLf & ",FCSdtStkFac,FTSkuDepCode,FTSkuBrdCode,FTPunCode"
    tSql = tSql & vbCrLf & "," & "FCSdtQtyAllDisAvi" & " AS FCScdQtyB4 "
    tSql = tSql & vbCrLf & "," & "FCSdtQty" & " AS FCScdQty"
    tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtQtyAllDisAvi - " & pcAlwDisQty & " < 0 THEN 0 ELSE FCSdtQtyAllDisAvi - " & pcAlwDisQty & " END AS FCScdQtyAF"
    tSql = tSql & vbCrLf & "," & "FCSdtSaleAmt" & " AS FCScdAmtB4"
    tSql = tSql & vbCrLf & "," & "FCSdtAmtDisAvi" & " AS FCScdAmtB4Avi"
    tSql = tSql & vbCrLf & "," & pcDisAmt & " AS FCScdAmt"
    tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtSaleAmt - " & pcDisAmt & " < 0 THEN 0 ELSE FCSdtSaleAmt - " & pcDisAmt & " END AS FCScdAmtAF"
    tSql = tSql & vbCrLf & "," & "CASE WHEN FCSdtAmtDisAvi -" & pcAlwDis & " <0 THEN 0 ELSE FCSdtAmtDisAvi -" & pcAlwDis & " END AS FCScdAmtAFAvi"
    tSql = tSql & vbCrLf & "," & 0 & " AS FCScdDiscPer"
    tSql = tSql & vbCrLf & "," & pcDisAmt & " AS FCScdDiscAmt"
    tSql = tSql & vbCrLf & "," & 0 & " AS FCScdDiscNew"
    tSql = tSql & vbCrLf & "," & tW_ScdDiscType & " AS FTScdDiscType"
    tSql = tSql & vbCrLf & "," & pnDiscType & " AS FNDctNo"
    tSql = tSql & vbCrLf & ",'" & SP_SQLtFunction("", "FTDctDiscType", "TCNMDiscountType", "WHERE FNDctNo = " & pnDiscType, oDB.ocnOnLine, True, True) & "' AS FTDctName" 'read master
    tSql = tSql & vbCrLf & ",'' AS FTDctRefCode"
    tSql = tSql & vbCrLf & ",'" & ptEventID & "' AS FTScdEventID"
    tSql = tSql & vbCrLf & ",'" & ptScdCardID & "' AS FTScdCardID"
    tSql = tSql & vbCrLf & ",'" & tUT_SaleDate & "' AS FDScdExpDate"
    tSql = tSql & vbCrLf & ",'" & ptCardType & "' AS FTScdCardType"

    
'*KT 51-09-19 Communication Sheet ML-SF-558 Print  Remain Discount for Employee Card.xls   เก็บค่าไว้ใน FTScdStaTrans ว่า Emp Card ใดไม่ต้องแสดง Remain Discount
'    Select Case nW_DiscType
'        Case nEN_Disc01Deal, nEN_Disc02PmtTime, nEN_Disc03Pmt, nEN_Disc05LnDisc, nEN_Disc06LnMDown, nEN_Disc04CoopDisc
'            tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
'        Case Else
'            tSql = tSql & vbCrLf & ",'1'  AS FTScdStaTrans"
'    End Select

    Select Case nW_DiscType
        Case nEN_Disc12EmpCrd
            If UCase(tVB_DiscEmpChkLimit) = "Y" Then
                tSql = tSql & vbCrLf & ",'1'  AS FTScdStaTrans"
            Else
                tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
            End If
        Case Else
            tSql = tSql & vbCrLf & ",'0'  AS FTScdStaTrans"
    End Select
'------------------------------------------------------------------------------------------------------------------------------------

    tSql = tSql & vbCrLf & ",'' AS FTScdCstID"
    tSql = tSql & vbCrLf & ",'" & ptCardType & "' AS FTScdCardKey"
    tSql = tSql & vbCrLf & ",'' AS FTScdStaChkExpire"
    tSql = tSql & vbCrLf & ",'" & ptStaChkDepart & "' AS FTScdStaChkDepart"
    tSql = tSql & vbCrLf & ",'' AS FNScdSptSeqNo"
    tSql = tSql & vbCrLf & ",0 AS FCScdMaxAlwDisc"
    tSql = tSql & vbCrLf & "," & nActSeqNo & " AS FNScdActSeqNo"
    tSql = tSql & vbCrLf & "," & nScdOrder & " AS FNScdOrder"

    If nW_DiscType = nEN_Disc02PmtTime Then
        cDisPmt = pcDisAmt
    Else
        cDisPmt = 0
    End If
    tSql = tSql & vbCrLf & "," & cDisPmt & " AS FCScdAmtTime"
    If nW_DiscType = nEN_Disc03Pmt Then
        cDisPmt = pcDisAmt
    Else
        cDisPmt = 0
    End If
    tSql = tSql & vbCrLf & "," & cDisPmt & " AS FCScdAmtPmt"
    If pnDiscType = nEN_Disc01Deal Then
        tSql = tSql & vbCrLf & ",'" & nMaxSeqNo & "' AS FNScdDealGroup"
    Else
        tSql = tSql & vbCrLf & ",'' AS FNScdDealGroup"
    End If
    tSql = tSql & vbCrLf & ",0 AS FCScdDiscPerOrg"
    tSql = tSql & vbCrLf & "," & pcDisAmt & " AS FCScdDiscAmtOrg"
    tSql = tSql & vbCrLf & ",0 AS FCScdDiscNewOrg"
    
    tSql = tSql & vbCrLf & Left(SP_GETtWhoInsertValue, Len(SP_GETtWhoInsertValue) - 1)
    tSql = tSql & vbCrLf & " FROM " & tVB_CNTblDT & ""
    tSql = tSql & vbCrLf & tWhe
    If SP_SQLbExecute(tSql, , False, nEffect) Then
        If nEffect <= 0 Then
            Call SP_MSGnShowing(tC_Msg003, tCS_CNMsgWarn)
        End If
    Else
        Call SP_MSGnShowing(tC_Msg003, tCS_CNMsgWarn)
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub
Private Function W_DISnGetNextAct() As Integer
Dim tWhe As String
Dim nActSeqNo As Integer

    On Error GoTo ErrHandle
    tWhe = " WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " And FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " And FDShdTransDate = '" & tUT_SaleDate & "'"
    nActSeqNo = CInt(SP_VALcText2Double(SP_SQLtFunction("MAX", "FNScdActSeqNo", "" & tVB_CNTblCD & "", tWhe)))
    nActSeqNo = nActSeqNo + 1
    W_DISnGetNextAct = nActSeqNo
    Exit Function
ErrHandle:
    nActSeqNo = 1
    W_DISnGetNextAct = nActSeqNo
End Function
Private Function W_GETnGroup(ByVal pnDiscType As Integer) As Integer
    '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
    Select Case pnDiscType
        Case nEN_Disc01Deal
            W_GETnGroup = 1
        Case nEN_Disc02PmtTime, nEN_Disc03Pmt, nEN_Disc04CoopDisc
            W_GETnGroup = 2
        Case nEN_Disc05LnDisc, nEN_Disc06LnMDown
            W_GETnGroup = 3
        Case nEN_Disc09CoopCpn
            W_GETnGroup = 4
        Case nEN_Disc10CoopBrand
            W_GETnGroup = 5
        Case nEN_Disc07AutoTran
            W_GETnGroup = 6
        Case nEN_Disc08MallCpn
            W_GETnGroup = 7
        Case nEN_Disc11MallCrd
            W_GETnGroup = 8
        Case nEN_Disc12EmpCrd
            W_GETnGroup = 9
        Case nEN_Disc13TranDisc
            W_GETnGroup = 10
        Case nEN_Disc14TranMDown
            W_GETnGroup = 11
        Case nEN_Disc16FreeGift
            W_GETnGroup = 12
        Case nEN_Disc15Ontop
            W_GETnGroup = 13
        Case nEN_Disc23TranMDown1, nEN_Disc24TranMDown2, nEN_Disc25TranMDown3, nEN_Disc26TranMDown4, nEN_Disc27TranMDown5 '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow
            W_GETnGroup = 7
    End Select
End Function

Private Function W_GETtEmpChkLimit(ByVal ptEmpID As String) As String
    Dim orsTemp As ADODB.Recordset
    Dim tSql$
On Error GoTo ErrHandle
    tSql = "SELECT  FTEcdStaCreditLimit "
    tSql = tSql & vbCrLf & " FROM   TCNMEmpCardMtn"
    tSql = tSql & vbCrLf & " WHERE '" & ptEmpID & "' BETWEEN FTEcdFrom AND FTEcdTo"
    tSql = tSql & vbCrLf & "  AND  LEN(" & ptEmpID & ") <= LEN(FTEcdTo)"
    tSql = tSql & vbCrLf & " AND FTEcdStaRang = 'Y'"
    tSql = tSql & vbCrLf & " ORDER BY LEN(FTEcdFrom)"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
                W_GETtEmpChkLimit = SP_FEDtChkString(orsTemp, "FTEcdStaCreditLimit")
        Else
                W_GETtEmpChkLimit = "N"
        End If
    Else
        W_GETtEmpChkLimit = "N"
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_GETtEmpChkLimit = "N"
     Call SP_DATxRsNothing(orsTemp)
End Function
