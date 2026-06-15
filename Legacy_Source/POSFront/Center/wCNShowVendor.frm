VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNShowVendor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detail Vendor Coupon;klg_Title"
   ClientHeight    =   4605
   ClientLeft      =   2685
   ClientTop       =   4125
   ClientWidth     =   8865
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8865
   Begin VSFlex7Ctl.VSFlexGrid ogdForm 
      Height          =   3690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8865
      _cx             =   15637
      _cy             =   6509
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Rows            =   9
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNShowVendor.frx":0000
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
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1560
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   3360
      Width           =   8880
      _Version        =   720897
      _ExtentX        =   15663
      _ExtentY        =   2752
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmChange 
         Caption         =   "Change Cp.(F5);klg_Change"
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
         Left            =   930
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton ocmCancelCp 
         Caption         =   "Cancel Cp.(F3);klg_CancelCp"
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
         Left            =   3060
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton ocmCancel 
         Cancel          =   -1  'True
         Caption         =   "&Close;klg_Close"
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
         Left            =   7320
         TabIndex        =   2
         Tag             =   "2;"
         Top             =   720
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
         Left            =   6030
         TabIndex        =   1
         Tag             =   "2;"
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Vendor Coupon :;klg_TotalV"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   360
         Width           =   4245
      End
      Begin VB.Label olaForm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   6480
         TabIndex        =   7
         Top             =   360
         Width           =   825
      End
      Begin VB.Label olaForm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   6940
         TabIndex        =   6
         Top             =   360
         Width           =   1905
      End
   End
End
Attribute VB_Name = "wCNShowVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nW_Col As Integer
Private tW_CardNo As String
Private tW_KeyPress As String 'เก็บ ราคาที่ Key  ในการกดแป้นพิมพ์ที่ grid
Private tW_KeyCode As String 'เก็บ KeyCode ในการกดแป้นพิมพ์ที่ grid
Private cW_SrcAmtTmp As Double
Private nW_VendorCnpQty  As Integer
Private tW_VendorCnpType As String
Private tW_VendorType As String

'Private Sub W_FNDxCheckBarCode(ByVal ptCardNo As String, ByVal pnRow As Integer)
''___________________________________________________
''Call:
''Cmt:
''___________________________________________________
'Dim orsTemp As ADODB.Recordset
'Dim tSql As String
'Dim tNo As String, tCouponNo As String, tPrice As String, tPriceDec As String
'
'   On Error GoTo ErrHandle
'   With ogdForm
'        tSql = "SELECT * FROM TCNMSkuBarCon"
'        tSql = tSql & " WHERE FNBafItemSize = '" & Len(Trim(ptCardNo)) & "'"
'        tSql = tSql & " AND FNBafMatchStr = '" & Left(Trim(ptCardNo), 1) & "'"
'        tSql = tSql & " AND FTBafCodeType = '2'"
'        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
'        If Not (orsTemp.EOF And orsTemp.BOF) Then
'            Call SP_FNDxCheckCfgBarCode(orsTemp, ptCardNo, tCouponNo, tPrice, tPriceDec)
'            tW_VendorCnpType = 1
'            .TextMatrix(pnRow, .ColIndex("FTSrcCardNo")) = tCouponNo
'            .TextMatrix(pnRow, .ColIndex("FCSrcAmtTmp")) = SP_FMTcCurPoint(tPrice & "." & tPriceDec, False)
'            .TextMatrix(pnRow, .ColIndex("FTVendorCnpType")) = tW_VendorCnpType
'        Else
'            tSql = "SELECT * FROM TCNMVndCoupon"
'            tSql = tSql & " WHERE FTVdcCode = '" & Trim(ptCardNo) & "'"
'            Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
'            If Not (orsTemp.EOF And orsTemp.BOF) Then
''                .TextMatrix(pnRow, .ColIndex("FTSrcCardNo")) = tCouponNo
'                tW_VendorCnpType = 2
'                .TextMatrix(pnRow, .ColIndex("FCSrcAmtTmp")) = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCVdcCpnAmt"), False)
'                .TextMatrix(pnRow, .ColIndex("FTVendorCnpType")) = tW_VendorCnpType
'            Else
'                tW_VendorCnpType = 3
'                .TextMatrix(pnRow, .ColIndex("FTVendorCnpType")) = tW_VendorCnpType
'            End If
'        End If
'    End With
'    Call SP_DATxRsNothing(orsTemp)
'Exit Sub
'ErrHandle:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'    tW_VendorCnpType = ""
'End Sub

Private Function W_CHKbCheckItemVendor(ByVal ptIndex As String) As Boolean
'___________________________________________________
'Call:
'Cmt:
'___________________________________________________
Dim cSumItem  As Double
Dim nRow As Integer, nI As Integer

    On Error GoTo ErrHandle
    With ogdForm
        nRow = .Aggregate(flexSTCount, .FixedRows, .ColIndex("FNSrcSeqNo"), .rows - 1, .ColIndex("FNSrcSeqNo"))
        For nI = 1 To nRow
            If ptIndex = Trim(.TextMatrix(nI, .ColIndex("FTSkuIndex"))) Then
                cSumItem = cSumItem + .ValueMatrix(nI, .ColIndex("FCSrcAmt"))
            End If
        Next nI
        If SP_VALcText2Double(cSumItem) > SP_VALcText2Double(wMain.ogdForm.TextMatrix(ptIndex, wMain.ogdForm.ColIndex("FCSdtSaleAmt"))) Then
            W_CHKbCheckItemVendor = True
        Else
            W_CHKbCheckItemVendor = False
        End If
    End With
    
 Exit Function
ErrHandle:
    W_CHKbCheckItemVendor = False
End Function

Private Sub W_FRMxSetInitial()
'___________________________________________________
'Call:
'Cmt:
'___________________________________________________
    nW_Col = 1
    If bVB_CheckVendor = False Then
'        ocmCancelCp.Enabled = False
    End If
    Call SP_OBJxSetFocus(ogdForm)
End Sub

'Private Sub W_FRMxEditPrice(ByVal pnRow As Integer, Optional ByVal pbCheck As Boolean = True)
Private Sub W_FRMxEditPrice()
'_____________________________________________
'Call:
'Cmt:
'_____________________________________________
Dim tSql As String, tCardNo As String, tTdmType As String
Dim orsTemp As ADODB.Recordset
Dim nCount As Integer, nSeqNo As Integer, cQty As Integer
Dim cAmt As Double, cVendorCpn As Double
Dim nRow As Integer
Dim tWhr As String, tVendorType As String

    On Error GoTo ErrHandle
'    With ogdForm
'        If pnRow = 0 Then Exit Sub
'        If pbCheck = True Then
'            If Trim(.TextMatrix(pnRow, .ColIndex("FNSrcSeqNo"))) = "" Then Exit Sub
'            nSeqNo = .TextMatrix(pnRow, .ColIndex("FNSrcSeqNo"))
'            tCardNo = Trim(.TextMatrix(pnRow, .ColIndex("FTSrcCardNo")))
'            cQty = .TextMatrix(pnRow, .ColIndex("FNVendorCnpQty"))
'            cAmt = SP_FMTcCurPoint(.TextMatrix(pnRow, .ColIndex("FCSrcAmt")), True, False)
'            tVendorType = .TextMatrix(pnRow, .ColIndex("FTVendorCnpType"))
'
'            tSql = "UPDATE " & tVB_CNTblRC
'            tSql = tSql & vbCrLf & "SET FTSrcCardNo = '" & tCardNo & "'"
'            tSql = tSql & vbCrLf & ",FCSrcFAmt = " & cAmt
'            tSql = tSql & vbCrLf & ",FCSrcAmt = " & cAmt
'            tSql = tSql & vbCrLf & ",FCSrcNet = " & cAmt
'            tSql = tSql & vbCrLf & ",FCSrcChg = 0"
'            tSql = tSql & vbCrLf & ",FTVendorCnpQty = '" & cQty & "'"
'            tSql = tSql & vbCrLf & ",FTVendorCnpType = '" & tVendorType & "'"
'            tSql = tSql & vbCrLf & ",FTSrcRmk='Change Vendor'"
'
'            tSql = tSql & vbCrLf & "WHERE FNSrcSeqNo=" & nSeqNo
'            tSql = tSql & vbCrLf & "AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
'            tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
'            tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
'            tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
'            Call SP_SQLvExecute(tSql)
'        End If
'    End With
    
    cVendorCpn = 0
     If bVB_CheckVendor = False Then
        With wSub
'            Debug.Print .ogdTender.rows
            For nRow = 1 To .ogdTender.rows - 1
                .ogdTender.TextMatrix(nRow, .ogdTender.ColIndex("FTTdmType")) = ""
                .ogdTender.TextMatrix(nRow, .ogdTender.ColIndex("FCSrcNet")) = ""
            Next nRow
'            wSub.ogdTender.Refresh
            tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
            tWhr = tWhr & vbCrLf & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
            tWhr = tWhr & vbCrLf & " AND  FTTdmCode = ''T006''" & " AND FTVendorCnpFlag=''1''"
            tWhr = tWhr & vbCrLf & " ORDER BY  FNSrcSeqNo ASC"
            tSql = "STP_CNxSQLFunction @ptFunc = 'FNSrcSeqNo,FTTdmType,FCSrcNet' ,@ptTable = '" & tVB_CNTblRC & "',@ptWhere = '" & tWhr & "'"
            Call SP_SQLvExecute(tSql, orsTemp)
            nRow = 0
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                orsTemp.MoveFirst
                tTdmType = SP_FEDtChkString(orsTemp, "FTTdmType")
                Do While Not orsTemp.EOF
                    cVendorCpn = cVendorCpn + CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))))
                    orsTemp.MoveNext
                Loop
                nRow = 1
                If nRow > .ogdTender.rows - 1 Then .ogdTender.rows = .ogdTender.rows + 1
                .ogdTender.TextMatrix(nRow, .ogdTender.ColIndex("FNSrcSeqNo")) = 1
                .ogdTender.TextMatrix(nRow, .ogdTender.ColIndex("FTTdmType")) = tTdmType
                .ogdTender.TextMatrix(nRow, .ogdTender.ColIndex("FCSrcNet")) = cVendorCpn
            End If
            tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
            tWhr = tWhr & vbCrLf & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
            tWhr = tWhr & vbCrLf & " AND  FTTdmCode = ''T006''" & " AND FTVendorCnpFlag=''0''"
            tWhr = tWhr & vbCrLf & " ORDER BY  FNSrcSeqNo ASC"
            tSql = "STP_CNxSQLFunction @ptFunc = 'FNSrcSeqNo,FTTdmType,FCSrcNet' ,@ptTable = '" & tVB_CNTblRC & "',@ptWhere = '" & tWhr & "'"
            Call SP_SQLvExecute(tSql, orsTemp)
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                orsTemp.MoveFirst
                Do While Not orsTemp.EOF
'                    nRow = .ogdTender.Aggregate(flexSTCount, .ogdTender.FixedRows, .ogdTender.ColIndex("FTTdmType"), .ogdTender.rows - 1, .ogdTender.ColIndex("FTTdmType")) + 1
                    nRow = nRow + 1
                    If nRow = .ogdTender.rows Then .ogdTender.rows = .ogdTender.rows + 1
                    .ogdTender.TextMatrix(nRow, .ogdTender.ColIndex("FNSrcSeqNo")) = nRow
                    .ogdTender.TextMatrix(nRow, .ogdTender.ColIndex("FTTdmType")) = SP_FEDtChkString(orsTemp, "FTTdmType")
                    .ogdTender.TextMatrix(nRow, .ogdTender.ColIndex("FCSrcNet")) = CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))))
                    cVendorCpn = cVendorCpn + CDbl(Val(Format(SP_FEDvChkNumeric(orsTemp, "FCSrcNet"))))
                    orsTemp.MoveNext
                 Loop
            End If
            .ogdTender.TopRow = nRow
            '________________________________________________________________________________________________
            '*Ao 53-10-16   ถ้าแก้ไขส่วนนี้ให้ไปแก้ไขที่ wCNSubPay,wCNSubPay640
            Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnDeposit, tEN_TrnAddDeposit
                    'non ทำการหา cVB_RoundRcv ที่ PRCxDeposit/wCNDepositAmt แล้ว
                    cVB_NetSaleRound = cVB_NetSale + cVB_NetDutyFreeDep + cVB_RoundRcv
                    .olaTotal.Caption = SP_FMTcCurPoint(cVB_Total, False, True)  ''*KT 51-04-10 Comsheet ML-SF-201 แสดงยอดรวมก่อนหักส่วนลดที่เกิดจาก Form wCNSubDisc2
                    .olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True) 'SP_FMTcCurPoint(cVB_TotalDisc, False, True) '*KT 51-04-10 Comsheet ML-SF-201 แสดงส่วนลดเฉพาะที่เกิดจากหน้า wCNSubDisc2
                    .olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total + cVB_NetDutyFreeDep - cVB_TotalDiscAft, False, True)
                    .olaAmtTender.Caption = SP_FMTtNum2Shw(0 + cVendorCpn, False, True)
                    cVB_Balance = cVB_NetSaleRound - cVendorCpn
                    .olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
                    .lblDepositAmt.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_DepositAmt), False, True)
                Case tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnRedeem
                    cVB_NetSaleRound = cVB_NetSale + cVB_NetDutyFreeDep + cVB_RoundRcv
                    .olaTotal.Caption = SP_FMTcCurPoint(cVB_Total, False, True)  ''*KT 51-04-10 Comsheet ML-SF-201 แสดงยอดรวมก่อนหักส่วนลดที่เกิดจาก Form wCNSubDisc2
                    .olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True) 'SP_FMTcCurPoint(cVB_TotalDisc, False, True) '*KT 51-04-10 Comsheet ML-SF-201 แสดงส่วนลดเฉพาะที่เกิดจากหน้า wCNSubDisc2
                    .olaNetSale.Caption = SP_FMTcCurPoint(cVB_Total + cVB_NetDutyFreeDep - cVB_TotalDiscAft, False, True)
                    .olaAmtTender.Caption = SP_FMTtNum2Shw(0 + cVendorCpn, False, True)
                    cVB_Balance = cVB_NetSaleRound - cVendorCpn
                    .olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
                    .lblDepositAmt.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_DepositAmt), False, True)
                Case Else
                    'cVB_RoundDiffAmt = SP_FMTcCurPoint((cVB_NetSale + cVB_NetDutyFreeDep) - ((cVB_NetSale + cVB_NetDutyFreeDep) \ 1), True, False)
                    cVB_RoundRcv = IIf(bVB_CNAlwRnd, SP_FMTcCurPoint(UT_GetnRoundDiff(cVB_NetSale + cVB_NetDutyFreeDep), True, False), 0) '*KT 51-04-05 Comsheet ML-SF-172,ML-SF-173
                    cVB_NetSaleRound = cVB_NetSale + cVB_RoundRcv
                    .olaTotal.Caption = SP_FMTcCurPoint(cVB_Total, False, True)  ''*KT 51-04-10 Comsheet ML-SF-201 แสดงยอดรวมก่อนหักส่วนลดที่เกิดจาก Form wCNSubDisc2
                    .olaDiscTotal.Caption = SP_FMTcCurPoint(cVB_TotalDiscAft, False, True) 'SP_FMTcCurPoint(cVB_TotalDisc, False, True) '*KT 51-04-10 Comsheet ML-SF-201 แสดงส่วนลดเฉพาะที่เกิดจากหน้า wCNSubDisc2
                    .olaNetSale.Caption = SP_FMTcCurPoint(cVB_NetSale, False, True)
                    .olaAmtTender.Caption = SP_FMTtNum2Shw(0 + cVendorCpn, False, True)
                    cVB_Balance = cVB_NetSaleRound + cVB_NetDutyFreeDep - cVendorCpn '*KT 51-04-23 กรณีที่ขายแบบ Duty Free ต้องบวก Deposit Amt ด้วยครับ
'                    Call SP_CRTxLogCashBalEvent(Me.Name, "Form_Load[1]", "Cash Tender", cVB_Balance)
                    .olaBalance.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
                    .lblDepositAmt.Caption = SP_FMTcCurPoint(SP_VALcText2Double(cVB_NetDutyFreeDep), False, True)
            End Select
            nUT_SeqRC = 0
            '________________________________________________________________________________________________
        End With
     End If

    Call SP_DATxRsNothing(orsTemp)
 Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub W_FRMxDeleteRC(ByVal ptSeqNo As Integer)
'_____________________________________________
'Call:
'Cmt:
'_____________________________________________
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim nI As Integer, nSeqNo As Integer, nMax As Integer
Dim nEff As Long
    On Error GoTo ErrHandle
    
    tSql = "DELETE FROM " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "WHERE FNSrcSeqNo=" & ptSeqNo
    tSql = tSql & vbCrLf & "AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
    
    Call SP_SQLvExecute(tSql, , , , , nEff)
    With ogdForm
        nMax = .Aggregate(flexSTCount, .FixedRows, .ColIndex("FNSrcSeqNo"), .rows - 1, .ColIndex("FNSrcSeqNo"))
        If ptSeqNo = nMax Then
            If W_FRMbShowICoupon Then
                Call W_SetGridHeading
                Call SP_OBJxSetFocus(ocmOk)
            Else
                .Clear
                Call W_SetGridHeading
            End If
            Exit Sub
        End If
    End With
    tSql = "SELECT * FROM " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "ORDER BY FNSrcSeqNo"

    Call SP_SQLvExecute(tSql, orsTemp)
     If Not (orsTemp.EOF And orsTemp.BOF) Then
        For nI = 1 To orsTemp.RecordCount
            nSeqNo = SP_FEDvChkNumeric(orsTemp, "FNSrcSeqNo")
            tSql = "UPDATE " & tVB_CNTblRC
            tSql = tSql & vbCrLf & "SET FNSrcSeqNo=" & nI
            tSql = tSql & vbCrLf & "WHERE FNSrcSeqNo=" & nSeqNo
            tSql = tSql & vbCrLf & "AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
            tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
            tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
            Call SP_SQLvExecute(tSql)
            orsTemp.MoveNext
        Next nI
        If W_FRMbShowICoupon Then
            Call W_SetGridHeading
            Call SP_OBJxSetFocus(ocmOk)
        Else
            ogdForm.Clear
            Call W_SetGridHeading
        End If
    Else
    
    End If
    
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function W_FRMbShowICoupon() As Boolean
'_____________________________________________
'Call:
'Cmt:
'_____________________________________________
Dim tSql As String
Dim orsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    W_FRMbShowICoupon = False
    tSql = "SELECT DISTINCT ISNULL(" & tVB_CNTblDT & ".FTSdtBarCode,'') AS FTSdtBarCode"
    'tSql = tSql & vbCrLf & ",ISNULL(" & tVB_CNTblDT & ".FTSkuAbbName,'Vendor Coupon') AS FTSkuAbbName" '*Tao 62-07-05 ของเดิม
    tSql = tSql & vbCrLf & ",ISNULL(" & tVB_CNTblDT & ".FTSkuAbbName,'Vendor Voucher') AS FTSkuAbbName" '*Tao 62-07-05 เปลี่ยนเป็น Vendor Voucher
    tSql = tSql & vbCrLf & "," & tVB_CNTblRC & ".FTSrcCardNo"
    tSql = tSql & vbCrLf & "," & tVB_CNTblRC & ".FCSrcAmt/" & tVB_CNTblRC & ".FTVendorCnpQty AS FCSrcAmtTmp"
    tSql = tSql & vbCrLf & "," & tVB_CNTblRC & ".FTVendorCnpQty," & tVB_CNTblRC & ".FCSrcAmt"
    tSql = tSql & vbCrLf & "," & tVB_CNTblRC & ".FNSrcSeqNo"
    tSql = tSql & vbCrLf & ",ISNULL(" & tVB_CNTblRC & ".FTSkuIndex,'') AS FTSkuIndex"
    tSql = tSql & vbCrLf & ",ISNULL(" & tVB_CNTblRC & ".FTVendorCnpType,'') AS FTVendorCnpType"
    tSql = tSql & vbCrLf & "FROM " & tVB_CNTblRC & " LEFT JOIN " & tVB_CNTblDT
    tSql = tSql & vbCrLf & "ON " & tVB_CNTblRC & ".FTSkuIndex = " & tVB_CNTblDT & ".FTSkuIndex"
    tSql = tSql & vbCrLf & "AND " & tVB_CNTblRC & ".FTTmnNum = " & tVB_CNTblDT & ".FTTmnNum"
    tSql = tSql & vbCrLf & "AND " & tVB_CNTblRC & ".FTShdTransNo = " & tVB_CNTblDT & ".FTShdTransNo"
    tSql = tSql & vbCrLf & "AND " & tVB_CNTblRC & ".FTShdTransType = " & tVB_CNTblDT & ".FTShdTransType"
    tSql = tSql & vbCrLf & "AND " & tVB_CNTblRC & ".FDShdTransDate = " & tVB_CNTblDT & ".FDShdTransDate"
    tSql = tSql & vbCrLf & "WHERE " & tVB_CNTblRC & ".FTTdmCode  ='T006'"
    
    If bVB_CheckVendor Then
        tSql = tSql & vbCrLf & "AND " & tVB_CNTblRC & ".FTVendorCnpFlag = '1'"
    End If
    tSql = tSql & vbCrLf & "AND " & tVB_CNTblRC & ".FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND " & tVB_CNTblRC & ".FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND " & tVB_CNTblRC & ".FTShdTransType = '" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & "AND " & tVB_CNTblRC & ".FDShdTransDate = '" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & "AND ISNULL(" & tVB_CNTblRC & ".FTSrcStaNoPay,'') <> '1'"           '*Ao 2011/01/17   CommSheet ML-PRV-020-Change Tender is wrong concept.xls
    tSql = tSql & vbCrLf & "ORDER BY FNSrcSeqNo"
    Call SP_SQLvExecute(tSql, orsTemp)
    If (orsTemp.EOF And orsTemp.BOF) Then
        W_FRMbShowICoupon = False
    Else
        W_FRMbShowICoupon = True
        Set ogdForm.DataSource = orsTemp
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_FRMbShowICoupon = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Sub W_SetGridHeading()
'----------------------------------------------------------
'   Call:
'   Cmt:   set heading of selling grid
'----------------------------------------------------------
    Const tFed = "FTSdtBarCode,FTSkuAbbName,FTSrcCardNo,FCSrcAmtTmp,FNVendorCnpQty,FCSrcAmt,FNSrcSeqNo,FTSkuIndex,FTVendorCnpType"
    Dim tColW$, iCol%
    On Error GoTo ErrHandle
    tColW = "1700;2000;1800;1100;700;1800;0;0;0"
    iCol = 9
    With ogdForm
        Call SP_GRDxSetProperty(ogdForm, tFed, tColW, iCol, True, 0)
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSdtBarCode")) = UT_tToken(tCS_ColBarCode, ";", UT_nPosCutTag)     'SKU Name
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSkuAbbName")) = UT_tToken(tCS_ColProductName, ";", UT_nPosCutTag)     'Product Name
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSrcCardNo")) = UT_tToken(tCS_ColCoupon, ";", UT_nPosCutTag)                         'No
        .TextMatrix(.FixedRows - 1, .ColIndex("FCSrcAmtTmp")) = UT_tToken(tCS_ColAmount, ";", UT_nPosCutTag)      'Select
        .TextMatrix(.FixedRows - 1, .ColIndex("FNVendorCnpQty")) = UT_tToken(tCS_ColQty, ";", UT_nPosCutTag)                                'Qty
        .TextMatrix(.FixedRows - 1, .ColIndex("FCSrcAmt")) = UT_tToken(tCS_ColTAmount, ";", UT_nPosCutTag)                 'Sub Total
        .TextMatrix(.FixedRows - 1, .ColIndex("FNSrcSeqNo")) = "FNSrcSeqNo"
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSkuIndex")) = "FTSkuIndex"
        .TextMatrix(.FixedRows - 1, .ColIndex("FTVendorCnpType")) = "FTVendorCnpType"
        .ColHidden(6) = True
        .ColHidden(7) = True
        .ColHidden(8) = True
        If .Aggregate(flexSTCount, .FixedRows, .ColIndex("FNSrcSeqNo"), .rows - 1, .ColIndex("FNSrcSeqNo")) + 1 < 9 Then
            .rows = 9
        End If
    End With
    Exit Sub
ErrHandle:
    ogdForm.rows = 9
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'-----------------------------------------------------------------------------------------------------------------------
'cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
'กรุณาตรวจสอบความถูกต้องด้วย
'-----------------------------------------------------------------------------------------------------------------------
Dim nRow As Integer
   
    With ogdForm
         Select Case KeyCode
            Case 38, 40                                     'ลุกศรขึ้น ลง
'                nRow = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FNSrcSeqNo"), .rows - 1, .ColIndex("FNSrcSeqNo"))
'                If .TextMatrix(.Row, .ColIndex("FNSrcSeqNo")) = "" Then
'                    .HighLight = flexHighlightNever
'                    Exit Sub
'                End If
                nRow = .RowSel

                If nRow = 0 Then nRow = 1
'                .Enabled = True
                .HighLight = flexHighlightAlways
                .Select nRow, 1                        'move highlight to first row
                 DoEvents 'ห้ามเอาออกไม่นั้นจะทำให้ Grid ไม่มีการ Select Row
                .SetFocus
            Case 116                           '39, 41, 116, 9 'ลุกศรซ้าย ขวา,F5,Tab
'                If nW_Col > 3 Then nW_Col = 1
'                nW_Col = nW_Col + 1
'                If ocmChange.Enabled = True Then
'                    Call ocmChange_Click
'                End If
            Case 13
                Call SP_OBJxSetFocus(ocmOk)

            Case 114                                                'F3
'                If bVB_CheckVendor Then
'                    MsgBox "DELETE"
'                    nRow = .RowSel
'                    Call W_FRMxDeleteRC(nRow)
'                End If
                
        End Select
   End With
End Sub

Private Sub ocmCancelCp_Click()
Dim nRow As Integer
Dim tWhr As String, tSql As String
Dim orsTemp As ADODB.Recordset

    With ogdForm
        nRow = .RowSel
        If nRow = 0 Then nRow = 1
        If bVB_CheckVendor Then
            If .TextMatrix(nRow, .ColIndex("FTSdtBarCode")) = "" Then Exit Sub
            If SP_MSGnShowing("tms_CN005106", nCS_OkCancelDefaultButton1, , , vbCrLf & .TextMatrix(nRow, .ColIndex("FTSdtBarCode")) & ":" & .TextMatrix(nRow, .ColIndex("FTSrcCardNo"))) = vbOK Then
                Call W_FRMxDeleteRC(nRow)
                olaForm(0).Caption = SP_FMTtNum2Shw(.Aggregate(flexSTSum, .FixedRows, .ColIndex("FCSrcAmt"), .rows - 1, .ColIndex("FCSrcAmt")))
                olaForm(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FNVendorCnpQty"), .rows - 1, .ColIndex("FNVendorCnpQty"))
            End If
        Else
            If .TextMatrix(nRow, .ColIndex("FNSrcSeqNo")) = "" Then Exit Sub
            tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
            tWhr = tWhr & vbCrLf & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
            tWhr = tWhr & vbCrLf & " AND  NOT FTTdmCode = ''T006''"
            tWhr = tWhr & vbCrLf & " ORDER BY  FNSrcSeqNo ASC"
            tSql = "STP_CNxSQLFunction @ptFunc = 'FNSrcSeqNo,FTTdmType,FCSrcNet' ,@ptTable = '" & tVB_CNTblRC & "',@ptWhere = '" & tWhr & "'"
            Call SP_SQLvExecute(tSql, orsTemp)
            If Not (orsTemp.BOF And orsTemp.EOF) Then
'                Call SP_MSGnShowing("tms_CN005109", nCS_Warning)
                Call SP_MSGnShowing("tms_CN005107", nCS_Information)
                Exit Sub
            End If
            If .TextMatrix(nRow, .ColIndex("FTSdtBarCode")) = "" Then
                If SP_MSGnShowing("tms_CN005106", nCS_OkCancelDefaultButton1, , , vbCrLf & .TextMatrix(nRow, .ColIndex("FTSkuAbbName")) & ":" & .TextMatrix(nRow, .ColIndex("FTSrcCardNo"))) = vbOK Then
                    Call W_FRMxDeleteRC(nRow)
                    olaForm(0).Caption = SP_FMTtNum2Shw(.Aggregate(flexSTSum, .FixedRows, .ColIndex("FCSrcAmt"), .rows - 1, .ColIndex("FCSrcAmt")))
                    olaForm(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FNVendorCnpQty"), .rows - 1, .ColIndex("FNVendorCnpQty"))
'                    Call ogdForm_AfterEdit(nRow, .ColIndex("FNSrcSeqNo"))
'                    Call W_FRMxEditPrice(nRow, False)
                    Call W_FRMxEditPrice
                End If
            Else
                Call SP_MSGnShowing("tms_CN005107", nCS_Information)
            End If
        End If
        Call SP_OBJxSetFocus(ogdForm)
    End With
End Sub

Private Sub ogdForm_KeyDown(KeyCode As Integer, Shift As Integer)
'-----------------------------------------------------------------------------------------------------------------------
'cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
'กรุณาตรวจสอบความถูกต้องด้วย
'-----------------------------------------------------------------------------------------------------------------------
Dim nRow As Integer
Dim orsTmp As ADODB.Recordset
Dim tTmp As String, tSql As String

    On Error GoTo ErrHandle
    
    With ogdForm
         Select Case KeyCode
            Case 38, 40                                     'ลุกศรขึ้น ลง
'                nRow = .Aggregate(flexSTMax, .FixedRows, .ColIndex("FNSrcSeqNo"), .rows - 1, .ColIndex("FNSrcSeqNo"))
'                If .TextMatrix(.Row, .ColIndex("FNSrcSeqNo")) = "" Then
'                    .HighLight = flexHighlightNever
'                    Exit Sub
'                End If
                nRow = .RowSel

                If nRow = 0 Then nRow = 1
'                .Enabled = True
                .HighLight = flexHighlightAlways
                .Select nRow, 1                        'move highlight to first row
                 DoEvents 'ห้ามเอาออกไม่นั้นจะทำให้ Grid ไม่มีการ Select Row
                .SetFocus
'            Case 116                           '39, 41, 116, 9 'ลุกศรซ้าย ขวา,F5,Tab
'                If nW_Col > 3 Then nW_Col = 1
'                nW_Col = nW_Col + 1
'                If ocmChange.Enabled = True Then
'                    tSql = "SELECT FTSkbRef FROM TSysKeyboard   WHERE FNSkbKey='" & CStr(KeyCode) & "'"
'                    tSql = tSql & " AND FNSkbShift = '" & CStr(Shift) & "'"
'                    tSql = tSql & " AND FTSkbwMain =  '1'"
'                    tSql = tSql & " AND FTSkbUsage ='Y'"
'                    Call SP_SQLvExecute(tSql, orsTmp, oDB.ocnOnLine, , , , , True, True) 'read master
'                    If Not (orsTmp.BOF And orsTmp.EOF) Then
'                        Call ocmChange_Click
'                    End If
'                End If
            Case 13
                Call SP_OBJxSetFocus(ocmOk)
'            Case 114                                                'F3
'                If bVB_CheckVendor Then                     'Vendor Item
'                    Call ocmCancelCp_Click
'                End If
            Case Else
                    tSql = "SELECT FTSkbRef FROM TSysKeyboard   WHERE FNSkbKey='" & CStr(KeyCode) & "'"
                    tSql = tSql & " AND FNSkbShift = '" & CStr(Shift) & "'"
                   ' tSql = tSql & " AND FTSkbwMain =  '1'" '*KT 54-04-08 V.3.0.56
                    tSql = tSql & " AND FTSkbUsage ='Y'"
                    Call SP_SQLvExecute(tSql, orsTmp, oDB.ocnOnLine, , , , , True, True) 'read master
                    If Not (orsTmp.BOF And orsTmp.EOF) Then
                        tTmp = SP_FEDtChkString(orsTmp, "FTSkbRef")
                        Select Case UCase(tTmp)
                            Case UCase("W_KBRcvVenCpnCancel")
                                Call ocmCancelCp_Click
                            Case UCase("W_KBRcvVenCpnChange")
                                If ocmChange.Enabled = True Then
                                    Call ocmChange_Click
                                End If
                        End Select
                    End If
        End Select
   End With
   Exit Sub
ErrHandle:
   
End Sub

Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With ogdForm
         If .TextMatrix(NewRow, .ColIndex("FNSrcSeqNo")) <> "" Then
            .TopRow = NewRow
            nVB_CNScollItem = NewRow
        End If
    End With
End Sub

Private Sub ogdForm_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
'-----------------------------------------------------------------------------------------------------------------------
'cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
'กรุณาตรวจสอบความถูกต้องด้วย
'-----------------------------------------------------------------------------------------------------------------------
    With ogdForm
        If .TextMatrix(NewRowSel, .ColIndex("FNSrcSeqNo")) = "" Then
            Cancel = True
        End If
    End With
End Sub

Private Sub ogdForm_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'-----------------------------------------------------------------------------------------------------------------------
'cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
'กรุณาตรวจสอบความถูกต้องด้วย
'-----------------------------------------------------------------------------------------------------------------------
    tW_CardNo = ""
    cW_SrcAmtTmp = 0
    nW_VendorCnpQty = 0
    tW_VendorType = ""
    With ogdForm
        Select Case UCase(.ColKey(Col))
            Case UCase("FTSrcCardNo"), UCase("FCSrcAmtTmp"), UCase("FNVendorCnpQty")
                tW_CardNo = .TextMatrix(Row, .ColIndex("FTSrcCardNo"))
                cW_SrcAmtTmp = .ValueMatrix(Row, .ColIndex("FCSrcAmtTmp"))
                nW_VendorCnpQty = .ValueMatrix(Row, .ColIndex("FNVendorCnpQty"))
                tW_VendorType = .TextMatrix(Row, .ColIndex("FTVendorCnpType"))
        End Select
    End With
End Sub

'Private Sub ogdForm_AfterEdit(ByVal Row As Long, ByVal Col As Long)
''-----------------------------------------------------------------------------------------------------------------------
''cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
''กรุณาตรวจสอบความถูกต้องด้วย
''-----------------------------------------------------------------------------------------------------------------------
'Dim tTmp As String
'Dim cSrcAmt As Double, cSrcAmtTmp As Double, cBalance As Double
'
'    On Error GoTo ErrHandle
'    With ogdForm
'        Select Case UCase(.ColKey(Col))
'            Case UCase("FTSrcCardNo")
'                 If .TextMatrix(Row, .ColIndex("FTSrcCardNo")) = "" Then
'                    .TextMatrix(Row, .ColIndex("FTSrcCardNo")) = tW_CardNo
'                    Exit Sub
'                Else
'                    If .TextMatrix(Row, .ColIndex("FTSrcCardNo")) <> tW_CardNo Then
'                        Call W_FNDxCheckBarCode(.TextMatrix(Row, .ColIndex("FTSrcCardNo")), Row)
'                    Else
'                        Exit Sub
'                    End If
'                End If
'            Case UCase("FCSrcAmtTmp")
'                 tTmp = .ValueMatrix(Row, .ColIndex("FCSrcAmtTmp"))
'                If SP_VALcText2Double(tTmp) = SP_VALcText2Double(0) Then
'                    If .TextMatrix(Row, .ColIndex("FNVendorCnpQty")) = 0 Then
'                        .TextMatrix(Row, .ColIndex("FNVendorCnpQty")) = 1
'                    End If
'                    .TextMatrix(Row, .ColIndex("FCSrcAmtTmp")) = .ValueMatrix(Row, .ColIndex("FCSrcAmt")) / .ValueMatrix(Row, .ColIndex("FNVendorCnpQty"))
'                End If
'            Case UCase("FNVendorCnpQty")
'                tTmp = .EditText
'                If SP_VALnText2Double(tTmp) = SP_VALnText2Double(0) Then
'                    .TextMatrix(Row, .ColIndex("FNVendorCnpQty")) = 1
'                End If
'        End Select
'        cSrcAmt = .ValueMatrix(Row, .ColIndex("FCSrcAmtTmp")) * .ValueMatrix(Row, .ColIndex("FNVendorCnpQty"))
'        cSrcAmtTmp = .TextMatrix(Row, .ColIndex("FCSrcAmt"))
'        .TextMatrix(Row, .ColIndex("FCSrcAmt")) = cSrcAmt
'        .HighLight = flexHighlightNever
'
'        If bVB_CheckVendor Then
'            If W_CHKbCheckItemVendor(Trim(.TextMatrix(Row, .ColIndex("FTSkuIndex")))) Then
'                Call SP_MSGnShowing("tms_CN005066", nCS_Error)
'                .TextMatrix(Row, .ColIndex("FTSrcCardNo")) = tW_CardNo
'                .TextMatrix(Row, .ColIndex("FCSrcAmtTmp")) = cW_SrcAmtTmp
'                .TextMatrix(Row, .ColIndex("FNVendorCnpQty")) = nW_VendorCnpQty
'                .TextMatrix(Row, .ColIndex("FTVendorCnpType")) = tW_VendorType
'                .TextMatrix(Row, .ColIndex("FCSrcAmt")) = .ValueMatrix(Row, .ColIndex("FCSrcAmtTmp")) * .ValueMatrix(Row, .ColIndex("FNVendorCnpQty"))
'                olaForm(0).Caption = SP_FMTtNum2Shw(.Aggregate(flexSTSum, .FixedRows, .ColIndex("FCSrcAmt"), .rows - 1, .ColIndex("FCSrcAmt")))
'                olaForm(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FNVendorCnpQty"), .rows - 1, .ColIndex("FNVendorCnpQty"))
'                Exit Sub
'            End If
'        Else
'            cBalance = SP_VALcText2Double(wSub.olaNetSale.Caption)
'            olaForm(0).Caption = SP_FMTtNum2Shw(.Aggregate(flexSTSum, .FixedRows, .ColIndex("FCSrcAmt"), .rows - 1, .ColIndex("FCSrcAmt")))
'            olaForm(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FNVendorCnpQty"), .rows - 1, .ColIndex("FNVendorCnpQty"))
'            If SP_VALcText2Double(olaForm(0).Caption) >= cBalance Then
'                Call SP_MSGnShowing("tms_CN005066", nCS_Error)
'                .TextMatrix(Row, .ColIndex("FTSrcCardNo")) = tW_CardNo
'                .TextMatrix(Row, .ColIndex("FCSrcAmtTmp")) = cW_SrcAmtTmp
'                .TextMatrix(Row, .ColIndex("FNVendorCnpQty")) = nW_VendorCnpQty
'                .TextMatrix(Row, .ColIndex("FTVendorCnpType")) = tW_VendorType
'                .TextMatrix(Row, .ColIndex("FCSrcAmt")) = .ValueMatrix(Row, .ColIndex("FCSrcAmtTmp")) * .ValueMatrix(Row, .ColIndex("FNVendorCnpQty"))
'                olaForm(0).Caption = SP_FMTtNum2Shw(.Aggregate(flexSTSum, .FixedRows, .ColIndex("FCSrcAmt"), .rows - 1, .ColIndex("FCSrcAmt")))
'                olaForm(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FNVendorCnpQty"), .rows - 1, .ColIndex("FNVendorCnpQty"))
'                Exit Sub
'            End If
'        End If
'        olaForm(0).Caption = SP_FMTtNum2Shw(.Aggregate(flexSTSum, .FixedRows, .ColIndex("FCSrcAmt"), .rows - 1, .ColIndex("FCSrcAmt")))
'        olaForm(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FNVendorCnpQty"), .rows - 1, .ColIndex("FNVendorCnpQty"))
'    End With
'    Call W_FRMxEditPrice(Row)
'Exit Sub
'ErrHandle:
''    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub

Private Sub ocmChange_Click()
Dim nRow As Integer, nCol As Integer
Dim tSql As String, tWhr As String
Dim orsTemp As ADODB.Recordset
            
    'NOT T006
    tWhr = " WHERE  FTTmnNum=''" & tVB_CNTerminalNum & "'' AND FTShdTransNo=''" & tVB_Running & "''"
    tWhr = tWhr & vbCrLf & " AND  FTShdTransType=''" & tVB_Transaction & "'' AND FDShdTransDate=''" & tUT_SaleDate & "''"
    tWhr = tWhr & vbCrLf & " AND  NOT FTTdmCode = ''T006''"
    tWhr = tWhr & vbCrLf & " ORDER BY  FNSrcSeqNo ASC"

    tSql = "STP_CNxSQLFunction @ptFunc = 'FNSrcSeqNo,FTTdmType,FCSrcNet' ,@ptTable = '" & tVB_CNTblRC & "',@ptWhere = '" & tWhr & "'"
    Call SP_SQLvExecute(tSql, orsTemp)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        Call SP_MSGnShowing("tms_CN005107", nCS_Information)
        Exit Sub
    End If
    
    With ogdForm
        nRow = .RowSel
        If bVB_CheckVendor Then
            If Not (.TextMatrix(nRow, .ColIndex("FTSdtBarCode")) = "") Then
                tVB_CardNo = .TextMatrix(nRow, .ColIndex("FTSrcCardNo"))
                cVB_Amt = .ValueMatrix(nRow, .ColIndex("FCSrcAmtTmp"))
                nVB_CnpQty = .ValueMatrix(nRow, .ColIndex("FNVendorCnpQty"))
                nVB_SeqNo = .ValueMatrix(nRow, .ColIndex("FNSrcSeqNo"))
                tVB_VendorCnpType = .TextMatrix(nRow, .ColIndex("FTVendorCnpType"))
                cVB_SrcAmt = .TextMatrix(nRow, .ColIndex("FCSrcAmt"))
                tVB_SkuIndex = .TextMatrix(nRow, .ColIndex("FTSkuIndex"))
                bVB_VedFrom = False
                wCNPayNewVendor.Show vbModal
'                Call W_FRMxEditPrice
                If W_FRMbShowICoupon Then
                    Call W_SetGridHeading
                    olaForm(0).Caption = SP_FMTtNum2Shw(.Aggregate(flexSTSum, .FixedRows, .ColIndex("FCSrcAmt"), .rows - 1, .ColIndex("FCSrcAmt")))
                    olaForm(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FNVendorCnpQty"), .rows - 1, .ColIndex("FNVendorCnpQty"))
'                    Call SP_OBJxSetFocus(ocmOk)            '*Ao 2010-12-14 CommSheet ML-PRV-022-Can't change vendor cpn again.xls
                Else
                    .Clear
                    Call W_SetGridHeading
                    ocmChange.Enabled = False
            '        olaForm(0).Caption = "0.00"
            '        olaForm(1).Caption = "0"
                End If
            Else
                tVB_CardNo = ""
                cVB_Amt = 0
                nVB_CnpQty = 0
                nVB_SeqNo = 0
                tVB_VendorCnpType = ""
                cVB_SrcAmt = 0
            End If
        Else
            If .TextMatrix(nRow, .ColIndex("FTSdtBarCode")) = "" Then
                tVB_CardNo = .TextMatrix(nRow, .ColIndex("FTSrcCardNo"))
                cVB_Amt = .ValueMatrix(nRow, .ColIndex("FCSrcAmtTmp"))
                nVB_CnpQty = .ValueMatrix(nRow, .ColIndex("FNVendorCnpQty"))
                nVB_SeqNo = .ValueMatrix(nRow, .ColIndex("FNSrcSeqNo"))
                tVB_VendorCnpType = .TextMatrix(nRow, .ColIndex("FTVendorCnpType"))
                cVB_SrcAmt = .TextMatrix(nRow, .ColIndex("FCSrcAmt"))
                tVB_SkuIndex = .TextMatrix(nRow, .ColIndex("FTSkuIndex"))
                bVB_VedFrom = False
                wCNPayNewVendor.Show vbModal
                Call W_FRMxEditPrice
                If W_FRMbShowICoupon Then
                    Call W_SetGridHeading
                    olaForm(0).Caption = SP_FMTtNum2Shw(.Aggregate(flexSTSum, .FixedRows, .ColIndex("FCSrcAmt"), .rows - 1, .ColIndex("FCSrcAmt")))
                    olaForm(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FNVendorCnpQty"), .rows - 1, .ColIndex("FNVendorCnpQty"))
'                    Call SP_OBJxSetFocus(ocmOk)            '*Ao 2010-12-14 CommSheet ML-PRV-022-Can't change vendor cpn again.xls
                Else
                    .Clear
                    Call W_SetGridHeading
                    ocmChange.Enabled = False
            '        olaForm(0).Caption = "0.00"
            '        olaForm(1).Caption = "0"
                End If
            Else
                tVB_CardNo = ""
                cVB_Amt = 0
                nVB_CnpQty = 0
                nVB_SeqNo = 0
                tVB_VendorCnpType = ""
                cVB_SrcAmt = 0
            End If
        End If
    End With
'    With ogdForm
'        If nW_Col > 3 Then nW_Col = 1
'        nW_Col = nW_Col + 1
'        nRow = .RowSel
'        If nRow = 0 Then nRow = 1
'        If bVB_CheckVendor Then
'            If nW_Col = 3 Then
'                If (.TextMatrix(nRow, .ColIndex("FTVendorCnpType")) = "1") Or (.TextMatrix(nRow, .ColIndex("FTVendorCnpType")) = "2") Then
'                    nW_Col = 4
'                End If
'            End If
'            If Not (.TextMatrix(nRow, .ColIndex("FTSdtBarCode")) = "") Then
'                .Enabled = True
'                .HighLight = flexHighlightAlways
'        '        .HighLight = flexHighlightWithFocus
'                .Select nRow, 1        'move highlight to first row
'                .Col = nW_Col
'                 DoEvents 'ห้ามเอาออกไม่นั้นจะทำให้ Grid ไม่มีการ Select Row
'                .EditCell
'            End If
'        Else
'            If nW_Col = 3 Then
'                If (.TextMatrix(nRow, .ColIndex("FTVendorCnpType")) = "1") Or (.TextMatrix(nRow, .ColIndex("FTVendorCnpType")) = "2") Then
'                    nW_Col = 4
'                End If
'            End If
'            If .TextMatrix(nRow, .ColIndex("FTSdtBarCode")) = "" Then
'                .Enabled = True
'                .HighLight = flexHighlightAlways
'        '        .HighLight = flexHighlightWithFocus
'                .Select nRow, 1        'move highlight to first row
'                .Col = nW_Col
'                 DoEvents 'ห้ามเอาออกไม่นั้นจะทำให้ Grid ไม่มีการ Select Row
'                .EditCell
'            End If
'        End If
'
'    End With
End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    Me.Top = IIf(bVB_Res640X480, cCS_Top680x480Vendor, cCS_Top800x600Vendor)
'    Me.Left = IIf(bVB_Res640X480, cCS_Left680x480Vendor, cCS_Left800x600Vendor)

    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    With ogdForm
        .FixedRows = 1
        .FixedCols = 0
        .rows = 2
        .Cols = 9
    If W_FRMbShowICoupon Then
        Call W_SetGridHeading
        olaForm(0).Caption = SP_FMTtNum2Shw(.Aggregate(flexSTSum, .FixedRows, .ColIndex("FCSrcAmt"), .rows - 1, .ColIndex("FCSrcAmt")))
        olaForm(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("FNVendorCnpQty"), .rows - 1, .ColIndex("FNVendorCnpQty"))
        Call SP_OBJxSetFocus(ocmOk)
    Else
        .Clear
        Call W_SetGridHeading
        ocmChange.Enabled = False
'        olaForm(0).Caption = "0.00"
'        olaForm(1).Caption = "0"
    End If
    End With
    Call W_FRMxSetInitial
End Sub

Private Sub ocmCancel_Click()
    tVB_VendorCpnNo = ""
    Unload Me
End Sub

Private Sub ocmOK_Click()
    tVB_VendorCpnNo = ""
    Unload Me
End Sub

Private Sub ogdForm_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With ogdForm
        Call cUT.UT_KBDChkKeyAscii(KeyAscii)

        Select Case KeyAscii
            Case 8, 13          'backspace, Enter
                Exit Sub
            Case 6                  'DELETE
                .TextMatrix(Row, Col) = ""
            Case Else
                Select Case UCase(.ColKey(Col))
                    Case UCase("FTSrcCardNo")
                        Call cUT.UT_KBDKeyNumeric(KeyAscii)
                    Case UCase("FCSrcAmtTmp")
                        Call cUT.UT_KBDKeyNumeric(KeyAscii)
                    Case UCase("FNVendorCnpQty")
                        Call cUT.UT_KBDKeyNumeric(KeyAscii)
                End Select
        End Select
    End With
    
End Sub

'Private Sub ogdForm_GotFocus()
''-----------------------------------------------------------------------------------------------------------------------
''cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
''กรุณาตรวจสอบความถูกต้องด้วย
''-----------------------------------------------------------------------------------------------------------------------
'    With ogdForm
'        Select Case UCase(.ColKey(.Col))
'            Case UCase("FCSrcAmtTmp")
'                tW_KeyPress = Replace(.TextMatrix(.Row, .Col), ".", "")
'            Case Else
'                tW_KeyPress = ""
'        End Select
'    End With
'End Sub

'Private Sub ogdForm_ChangeEdit()            '*AO 51-03-31  Comm.Sheet ML-SF-160
''-----------------------------------------------------------------------------------------------------------------------
''cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
''กรุณาตรวจสอบความถูกต้องด้วย
''-----------------------------------------------------------------------------------------------------------------------
'    With ogdForm
'        Select Case UCase(.ColKey(.Col))
'            Case UCase("FCSrcAmtTmp")
'                Select Case tW_KeyCode
'                    Case 13 'ถ้าไม่ดัก KeyCode Enter จะหลุดไปที่ Case Else ทำให้ทำงานผิดพลาดนะครับ กรุณาระวังด้วย
'                        Exit Sub
'                    Case vbKeyDelete
'                        .EditText = "0.00"
'                         tW_KeyPress = "0"
'                    Case Else
'                        .EditText = SP_DEFtChangeNum(tW_KeyPress)
''                        bW_ChkDecimals = False
'                End Select
''            Case UCase("FNSdtQty")
''                Select Case tW_KeyCode
''                    Case vbKeyDelete
''                        .EditText = ""
''                End Select
'        End Select
'    End With
'End Sub

'Private Sub ogdForm_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
''-----------------------------------------------------------------------------------------------------------------------
''cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
''กรุณาตรวจสอบความถูกต้องของ Communication Sheet ML-SF-381,ML-SF-382 ด้วยนะครับ
''-----------------------------------------------------------------------------------------------------------------------
'    With ogdForm
'        Call cUT.UT_KBDChkKeyAscii(KeyAscii)
'        Select Case UCase(.ColKey(Col))
'            Case UCase("FCSrcAmtTmp")
'                If KeyAscii <> 13 Then
'                    If (KeyAscii = vbKeyBack) Then      'backspace
'                        If Len(Trim(tW_KeyPress)) > 0 Then
'                            tW_KeyPress = Left(tW_KeyPress, Len(tW_KeyPress) - 1)
'                            .EditText = SP_DEFtChangeNum(tW_KeyPress)
'                        End If
'                        Exit Sub
'                    End If
'                    Call cUT.UT_KBDKeyCurrency(KeyAscii, tW_KeyPress)    'check key, currency
'                    If Len(.EditText) >= 11 Then KeyAscii = 0
'                Else
'                    '*KT 51-07-02 Communication Sheet ML-SF-381
'                    '======================================================
'                    'จำเป้นต้องใช้ tW_KeyPress แทน .TextMatrix(Row, Col) เพราะว่าค่าของ .TextMatrix(Row, Col) บางครั้งยังไม่รวมตัวเลขที่ Key ล่าสุดจนผ่าน ChangEdit ก่อน
'                    'ใครที่ต้องการแก้ กรุณาตรวจสอบให้รอบครอบด้วยนะ เพราะ Grid ต้อง Control ด้วย Keybord ได้
'                    'เกิดปัญหาว่าตอน Run ผ่าน Code ทำงานได้แต่เวลา Complier ไม่ผ่าน
'                    If Not (W_CHKbBanknoteAlw(SP_VALcText2Double(SP_DEFtChangeNum(tW_KeyPress)))) Then
'                        KeyAscii = 0
'                        tW_KeyPress = ""
''                        bVB_SetFocus = False
''                        nW_GridEnter = 0
''                        bW_ChkDecimals = False
''                        Call SP_MSGnShowing("tms_CN008028", nCS_Warning)   'คุณป้อนเศษสตางค์ไม่ถูกต้อง
''                        .Enabled = True
''                        .HighLight = flexHighlightAlways
''                        .TextMatrix(.Row, .ColIndex("FCSdtSaleUnit")) = SP_FMTcCurPoint(cW_SaleItem, True, True)
''                        .EditText = cW_SaleItem
''                        .Select Row, 1        'move highlight to first row
''                        .Col = .ColIndex("FCSdtSaleUnit")
''                        .EditCell
''                         tW_KeyPress = "" 'Replace(.TextMatrix(.Row, .Col), ".", "")
''                         bVB_PriceOveride = False
'                         Exit Sub
''                    Else
''                         bW_ChkDecimals = True
'                    End If
'                    '======================================================
''                    nW_GridEnter = nW_GridEnter + 1
''                    If nW_GridEnter = 1 Then ogdForm_AfterEdit Row, Col
'                    Exit Sub
'                End If
'                If KeyAscii > 0 Then tW_KeyPress = IIf(KeyAscii = 13, "", tW_KeyPress & Chr(KeyAscii)) 'ถ้า  KeyAscii = Enter หมายถึงจบการแก้ไขราคาต่อหน่วย
'        End Select
'    End With
'End Sub

'Private Sub ogdForm_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
''-----------------------------------------------------------------------------------------------------------------------
''cmt: คำเตือน การแก้ไข Event KeyDown , KeyDownEdit , KeyPress , KeyPressEdit , KeyUp , KeyUpEdit , ChangEdit
''กรุณาตรวจสอบความถูกต้องด้วย
''-----------------------------------------------------------------------------------------------------------------------
'    With ogdForm
'        tW_KeyCode = KeyCode
'        Select Case UCase(.ColKey(Col))
'            Case UCase("FCSrcAmtTmp")
'                Select Case KeyCode
'                    Case vbKeyDelete
'                        .EditText = "0.00"
'                End Select
''            Case UCase("FNSdtQty")
''                Select Case KeyCode
''                    Case vbKeyDelete
''                        .EditText = ""
''                End Select
'        End Select
'    End With
'End Sub

'Public Function W_CHKbBanknoteAlw(ByVal pcAmt As Double) As Boolean           '*AO51-03-26   Comm.Sheet  ML-SF-100
''----------------------------------------------------------
''   Call:
''   Cmt:    Banknote
''----------------------------------------------------------
'    Dim nDot As Integer, nIndex As Integer
'    Dim cChk As Double
'    Dim tTmp As String
'    Dim bLastFor As Boolean
'
'On Error GoTo Err_Handle
'
'    W_CHKbBanknoteAlw = False
'
'    If bVB_ConfigBanknote Then     'เช็คสตางค์
'        If Trim(CStr(aVB_CNStang(0))) = "" Then                     'ค่าว่าง ไม่มีเช็คสตางค์ปล่อยผ่าน
'            W_CHKbBanknoteAlw = True
'            Exit Function
'        End If
'        tTmp = Trim(CStr(pcAmt))
'        nDot = InStr(1, tTmp, ".", vbBinaryCompare)
'        If nDot = 0 Then                'ค่า .00 ไม่มีสตางค์
'            W_CHKbBanknoteAlw = True
'            Exit Function
'        End If
'        tTmp = Right(tTmp, (Len(tTmp) - nDot) + 1)
'        cChk = SP_VALcText2Double(tTmp)
'        bLastFor = True
'
'        Do While cChk > 0 And bLastFor
'            For nIndex = 0 To UBound(aVB_CNStang)
'                If cChk > 0 Then
'                    If aVB_CNStang(nIndex) <= cChk Then
'                        cChk = cChk - aVB_CNStang(nIndex)
'                    End If
'                Else
'                    Exit For
'                End If
'
'            Next nIndex
'            bLastFor = False
'        Loop
'        If cChk = 0 Then                'เช็คอีกครั้งว่าป้อนถูกไม่ ???
'            W_CHKbBanknoteAlw = True
'        Else
'            W_CHKbBanknoteAlw = False
'        End If
'    Else
'        W_CHKbBanknoteAlw = True           'ไม่ได้เช็คสตางค์ปล่อยผ่าน
'    End If
'
'    Exit Function
'Err_Handle:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'    W_CHKbBanknoteAlw = True   'เกิด Error ปล่อยผ่าน
'End Function



