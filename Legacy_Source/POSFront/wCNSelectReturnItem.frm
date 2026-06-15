VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNSelectReturnItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Return Item;klg_Title"
   ClientHeight    =   5400
   ClientLeft      =   3390
   ClientTop       =   3090
   ClientWidth     =   9570
   ControlBox      =   0   'False
   Icon            =   "wCNSelectReturnItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9570
   Begin VSFlex7Ctl.VSFlexGrid ogdForm 
      Height          =   4530
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNSelectReturnItem.frx":058A
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
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   4935
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   9570
      _Version        =   720897
      _ExtentX        =   16880
      _ExtentY        =   8705
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   960
      Index           =   1
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4440
      Width           =   9570
      _Version        =   720897
      _ExtentX        =   16880
      _ExtentY        =   1693
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmSelAll 
         Caption         =   "&Select all;klg_SelAll"
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
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton ocmCleAll 
         Caption         =   "&Clear all;klg_CleAll"
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
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   360
         Width           =   1455
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
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   360
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
         Left            =   6240
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNSelectReturnItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private bW_Show As Boolean
Private nW_CountNoRe As Long




'*KT 66-08-17 TTK-ISS3000-839
'--------------------------------------OLD------------------------------------
'Private Sub W_FRMxReferItemReturn()
''---------------------------------------------------------------
''Call :
''Cmt : ทำการ get record ของรายการสินค้าที่จะคืน โดย return wFunctionKB.orsRet กลับไป
''--------------------------------------------------------------
'    Dim tSql$
'    Dim tWhr$
'    Dim tValues$
'    Dim bPass As Boolean
'    Dim nI As Long
'    Dim orsTemp As ADODB.Recordset
'On Error GoTo ErrHandle             '*KT 54-03-25 ป้องกันโปรแกรมหลุด
'    bPass = False
'    With ogdForm
'        For nI = 1 To .rows - 1
'            If .Cell(flexcpChecked, nI, .ColIndex("FBChk"), nI, .ColIndex("FBChk")) = vbChecked Then
'                bPass = True
'                tValues = tValues & .TextMatrix(nI, .ColIndex("FNSdtSeqNo")) & ","
'            End If
'        Next nI
'        If tValues <> "" Then tValues = Mid(tValues, 1, Len(tValues) - 1)
'        nVB_SeqRC = tValues
'    End With
'    If bPass Then
'        tWhr = "WHERE  FTTmnNum=''" & tVB_TerminalNoRet & "'' AND FTShdTransNo=''" & tVB_TransNoRet & "''"
'        tWhr = tWhr & vbCrLf & " AND  FTShdTransType =''" & tVB_TransTypeRet & "'' AND FDShdTransDate=''" & Format(dVB_TransDateRet, "YYYY/MM/DD") & "''"
'        tWhr = tWhr & vbCrLf & " AND  FTSdtCancelFlag <>1"  '*AO 51-04-01   Comm.Sheet ML-SB-169
'        tWhr = tWhr & vbCrLf & " AND  FNSdtSeqNo IN (" & tValues & ")"
'        tWhr = tWhr & vbCrLf & " ORDER BY FNSdtSeqNo"
'        tSql = "STP_CNxSQLFunction @ptFunc = '*' ,@ptTable = '" & "TPSTSalDT" & "',@ptWhere = '" & tWhr & "'"
'        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True) 'PH 2.0.0
'         If Not (orsTemp.BOF Or orsTemp.EOF) Then
'            bVB_Return = True
'            Set wFunctionKB.orsRet = orsTemp.Clone
'        Else
'            bVB_Return = False
'        End If
'    Else
'        bVB_Return = False
'    End If
'    Call SP_DATxRsNothing(orsTemp)
'    Unload Me
'    Exit Sub
'ErrHandle:
'    bVB_Return = False
'    Call SP_DATxRsNothing(orsTemp)
'    Unload Me
'End Sub
'--------------------------------------OLD------------------------------------

Private Sub W_FRMxReferItemReturn()
'---------------------------------------------------------------
'Call :
'Cmt : ทำการ get record ของรายการสินค้าที่จะคืน โดย return wFunctionKB.orsRet กลับไป
'--------------------------------------------------------------
    Dim tSql$
    Dim tWhr$
    Dim tValues$
    Dim bPass As Boolean
    Dim nI As Long
    Dim orsTemp As ADODB.Recordset
    Dim nNum1 As Long
    Dim nNum2 As Long
    Dim nJ As Long
    Dim nCount As Long
On Error GoTo ErrHandle             '*KT 54-03-25 ป้องกันโปรแกรมหลุด

'    '*KT 62-09-04 POSC-ML-0570 CR อนุญาตให้คืนสินค้า Deal เป็นชุดได้ ถ้าคืนต้องคืนทั้งชุด ถ้าไม่คืนต้องไม่คืนทั้งชุด  เดิมถ้ามี Deal/Free ต้องคืนทั้งหมด
'    '-----------------------------------------------------------------------------
'    If bDealNo Then
'        With ogdForm
'            For nJ = 0 To UBound(aDealBBY)
'                nCount = 0
'                For nI = 1 To .rows - 1
'                    If .TextMatrix(nI, .ColIndex("FTSkuAbbName")) <> "" Then
'                        If .Cell(flexcpChecked, nI, .ColIndex("FBChk"), nI, .ColIndex("FBChk")) = vbChecked And .TextMatrix(nI, .ColIndex("FTDeal")) = "Y" And .TextMatrix(nI, .ColIndex("FTScdBBYNo")) = aDealBBY(nJ).tBBYNo Then      'Deal
'                            nCount = nCount + 1
'                        End If
'                    End If
'                Next nI
'                If nCount <> aDealBBY(nJ).nCount And nCount > 0 Then
'                    Call SP_MSGnShowing("The bill of return for this Deal / Free of BBY No " & aDealBBY(nJ).tBBYNo & " item must be forced to return the entire this group.;บิลคืนนี้มีรายการ Deal/Free BBY No " & aDealBBY(nJ).tBBYNo & "  ต้องบังคับคืนสินค้าในกลุ่มนี้ทั้งกลุ่ม", nCS_Information, , , "บิลคืนนี้มีรายการ Deal/Free BBY No " & aDealBBY(nJ).tBBYNo & "  ต้องบังคับคืนสินค้าในกลุ่มนี้ทั้งกลุ่ม")
'                    Exit Sub
'                End If
'            Next nJ
'        End With
'    End If
'    '-----------------------------------------------------------------------------
    
    bPass = False
    With ogdForm
        For nI = 1 To .rows - 1
            If .Cell(flexcpChecked, nI, .ColIndex("FBChk"), nI, .ColIndex("FBChk")) = vbChecked Then
                bPass = True
                tValues = tValues & .TextMatrix(nI, .ColIndex("FNSdtSeqNo")) & ","
            End If
        Next nI
        If tValues <> "" Then tValues = Mid(tValues, 1, Len(tValues) - 1)
        nVB_SeqRC = tValues
    End With
    If bPass Then
        tWhr = "WHERE  FTTmnNum=''" & tVB_TerminalNoRet & "'' AND FTShdTransNo=''" & tVB_TransNoRet & "''"
        tWhr = tWhr & vbCrLf & " AND  FTShdTransType =''" & tVB_TransTypeRet & "'' AND FDShdTransDate=''" & Format(dVB_TransDateRet, "YYYY/MM/DD") & "''"
        tWhr = tWhr & vbCrLf & " AND  ISNULL(FTSdtCancelFlag,''0'') <>''1''"  '*AO 51-04-01   Comm.Sheet ML-SB-169
        tWhr = tWhr & vbCrLf & " AND  FNSdtSeqNo IN (" & tValues & ")"
        tWhr = tWhr & vbCrLf & " ORDER BY FNSdtSeqNo"
        tSql = "STP_CNxSQLFunction @ptFunc = '*' ,@ptTable = '" & "TPSTSalDT WITH(NOLOCK) " & "',@ptWhere = '" & tWhr & "'"

        
        '*KT 61-08-21 TheMall-POS ML-IP-039 รองรับการคืนอ้างอิงบิลจาก DB Backup
        'Call SP_SQLvExecute(tSql, orsTemp, IIf(bVB_ReturnFrmBck, ocnDBBck, oDB.ocnOnLine), , , , , True) 'PH 2.0.0
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True) 'PH 2.0.0
         If Not (orsTemp.BOF Or orsTemp.EOF) Then
            bVB_Return = True
            Set wFunctionKB.orsRet = orsTemp.Clone
            With ogdForm
                nNum1 = .Aggregate(flexSTCount, .FixedRows, .ColIndex("FTSdtBarCode"), .rows - 1, .ColIndex("FTSdtBarCode"))
                nNum2 = UBound(Split(tValues, ",")) + 1 '.Aggregate(flexS, .FixedRows, .ColIndex("FBChk"), .rows - 1, .ColIndex("FBChk"))
                
                If (nNum1 - nW_CountNoRe) = nNum2 Then
                    bVB_ReturnAll = True
                Else
                    bVB_ReturnAll = False
                End If
                '*KT 62-09-06 เก็บสถานะว่ามีการคืนครั้งเดียวทั้งบิล
                If nNum1 = nNum2 Then
                    bVB_ReturnAllBill = True
                Else
                    bVB_ReturnAllBill = False
                End If
            End With
        Else
            bVB_Return = False
             bVB_ReturnAll = False
             bVB_ReturnAllBill = False
        End If
    Else
        bVB_Return = False
         bVB_ReturnAll = False
         bVB_ReturnAllBill = False
    End If
    'bVB_ReturnHaveDisc = bDealNo
    Call SP_DATxRsNothing(orsTemp)
    Unload Me
    Exit Sub
ErrHandle:
    bVB_Return = False
    bVB_ReturnAll = False
    bVB_ReturnAllBill = False
    ' bVB_ReturnHaveDisc = False
    Call SP_DATxRsNothing(orsTemp)
    Unload Me
End Sub


'*KT 66-08-17 TTK-ISS3000-839
'--------------------------------------OLD------------------------------------
'Private Sub W_FRMxShowItemsReturn()
''---------------------------------------------------------------
''Call :
''Cmt : ทำการแสดงรายการสินค้าที่จะคืนบนหน้าจอโดย default ให้มีการ Checkbox ทั้งหมด
''--------------------------------------------------------------
'    Dim tSql$
'    Dim nI As Long
'    Dim orsTemp As ADODB.Recordset
'
'    tSql = "SELECT DISTINCT  DT.FNSdtSeqNo AS FNSdtSeqNo,0 AS FBChk,CASE WHEN DT.FTSdtInStoreBar ='' THEN DT.FTSdtBarCode ELSE DT.FTSdtInStoreBar  END AS FTSdtBarCode,DT.FTSkuAbbName AS FTSkuAbbName"
'    tSql = tSql & vbCrLf & ",DT.FCSdtQty AS FCSdtQty,DT.FCSdtSaleAmt AS FCSdtSaleAmt,CD.FTFlag"
'    tSql = tSql & vbCrLf & "FROM   TPSTSalDT DT LEFT OUTER  JOIN"
'    tSql = tSql & vbCrLf & "        (SELECT DISTINCT FNSdtSeqNo,FTTmnNum,FTShdTransNo,FDShdTransDate,MAX(CASE WHEN FNDctNo IN (1) THEN '1' ELSE '0' END) AS FTFlag"
'    tSql = tSql & vbCrLf & "       From TPSTSalCD"
'    tSql = tSql & vbCrLf & "     WHERE  FTTmnNum='" & tVB_TerminalNoRet & "' AND FTShdTransNo='" & tVB_TransNoRet & "'"
'    tSql = tSql & vbCrLf & "        AND FTShdTransType ='" & tVB_TransTypeRet & "' AND FDShdTransDate='" & Format(dVB_TransDateRet, "YYYY/MM/DD") & "'"
'    tSql = tSql & vbCrLf & "        GROUP BY FNSdtSeqNo,FTTmnNum,FTShdTransNo,FDShdTransDate"
'    tSql = tSql & vbCrLf & "       ) CD"
'    tSql = tSql & vbCrLf & "     ON DT.FTTmnNum = CD.FTTmnNum AND"
'    tSql = tSql & vbCrLf & "      DT.FTShdTransNo = CD.FTShdTransNo AND"
'    tSql = tSql & vbCrLf & "      DT.FNSdtSeqNo = CD.FNSdtSeqNo AND"
'    tSql = tSql & vbCrLf & "     DT.FDShdTransDate = CD.FDShdTransDate"
'    tSql = tSql & vbCrLf & "WHERE DT.FTTmnNum='" & tVB_TerminalNoRet & "' AND DT.FTShdTransNo='" & tVB_TransNoRet & "'"
'    tSql = tSql & vbCrLf & "AND  DT.FTShdTransType ='" & tVB_TransTypeRet & "' AND DT.FDShdTransDate='" & Format(dVB_TransDateRet, "YYYY/MM/DD") & "'"
'    tSql = tSql & vbCrLf & "AND  DT.FTSdtCancelFlag <>1"
'    tSql = tSql & vbCrLf & "ORDER BY DT.FNSdtSeqNo"
'    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True) Then Call SP_DATxRsNothing(orsTemp): Exit Sub 'PH 2.0.0 ติดต่อฐานข้อมูล Server
'    Set ogdForm.DataSource = orsTemp
'    ogdForm.Refresh
'    For nI = 0 To orsTemp.Fields.Count - 1
'        ogdForm.ColKey(nI) = orsTemp.Fields(nI).Name
'    Next nI
'    If ogdForm.rows < 11 Then ogdForm.rows = 11
'    ogdForm.FixedCols = 1
'    Call W_SetGridHeading
'    With ogdForm
'        .ColHidden(.ColIndex("FTFlag")) = True
'        For nI = .FixedRows To .Row - 1
'            Select Case .ValueMatrix(nI, .ColIndex("FTFlag"))
'                Case "1"
'                    .TextMatrix(nI, .ColIndex("FTSkuAbbName")) = "*" & .TextMatrix(nI, .ColIndex("FTSkuAbbName"))
'                Case Else
'                    .TextMatrix(nI, .ColIndex("FBChk")) = 1
'            End Select
'        Next nI
'    End With
'    ogdForm.Enabled = True
'    ogdForm.HighLight = flexHighlightAlways
'    ogdForm.Select 1, 1                        'move highlight to first row
'    Set orsTemp = Nothing
'End Sub
'-------------------------OLD--------------------------------


Private Sub W_FRMxShowItemsReturn()
'---------------------------------------------------------------
'Call :
'Cmt : ทำการแสดงรายการสินค้าที่จะคืนบนหน้าจอโดย default ให้มีการ Checkbox ทั้งหมด
'--------------------------------------------------------------
    Dim tSql$
    Dim nI As Long
    'Dim orsTemp As ADODB.Recordset '*KT 63-03-11 ML-POSC-1139
    Dim nRow As Long
    Dim tBBYNo As String
    Dim bFirst As Boolean
    Dim bSave As Boolean
On Error GoTo ErrHandle
    nW_CountNoRe = 0
'    Dim tWhe As String '*Tao 60-02-08
'    Dim bAlipay As Boolean '*Tao 60-02-08
'
'    'ตรวจสอบก่อนว่าเป้นการชำระด้วย Alipay หรือไม่
'    tWhe = " WHERE  FTTmnNum='" & tVB_TerminalNoRet & "' AND FTShdTransNo='" & tVB_TransNoRet & "'"
'    tWhe = tWhe & " AND  FTShdTransType='" & tVB_TransTypeRet & "' AND FDShdTransDate='" & Format(dVB_TransDateRet, "YYYY/MM/DD") & "'"
'    tWhe = tWhe & " AND FTTdmCode = 'T016'"
'    If SP_SQLtFunction("TOP 1", "FTTdmCode", tVB_CNTblRC, tWhe) = "T016" Then bAlipay = True
'    '================================

    '*KT 62-05-28 ML-POSC-0670   Return refer sale(full bill is has deal/free)
    tSql = "SELECT DISTINCT  DT.FNSdtSeqNo AS FNSdtSeqNo,0 AS FBChk,CASE WHEN ISNULL(DT.FTSdtInStoreBar,'') ='' THEN DT.FTSdtBarCode ELSE DT.FTSdtInStoreBar  END AS FTSdtBarCode,ISNULL(DT.FTSkuAbbName,'')  AS FTSkuAbbName" '*KT 62-08-29 ML-POSC-0721 ISNULL Field FTSdtInStoreBar,FTSkuAbbName
    'tSql = tSql & vbCrLf & ",DT.FCSdtQty AS FCSdtQty,DT.FCSdtSaleAmt AS FCSdtSaleAmt,CD.FTDeal,CD.FTFlag,ISNULL(DT.FCSdtReturnQty,0) AS FCSdtReturnQty"
    tSql = tSql & vbCrLf & ",DT.FCSdtQty AS FCSdtQty,DT.FCSdtSaleAmt AS FCSdtSaleAmt,ISNULL(CD.FTFlag,'0') AS FTFlag,ISNULL(DT.FCSdtReturnQty,0) AS FCSdtReturnQty" '*TU 2019-08-23 CD.FTDeal,CD.FTFlag เป็นค่า Null
    tSql = tSql & vbCrLf & "FROM   TPSTSalDT DT WITH(NOLOCK) LEFT OUTER  JOIN"
    'tSql = tSql & vbCrLf & "                (SELECT DISTINCT FNSdtSeqNo,FTTmnNum,FTShdTransNo,FDShdTransDate,'0' AS FTFlag" ',MAX(CASE WHEN FNDctNo IN (1) THEN ISNULL(FTDctRefCode,'') ELSE '' END) AS FTDealNo"
    tSql = tSql & vbCrLf & "                 (SELECT DISTINCT FNSdtSeqNo,FTTmnNum,FTShdTransNo,FDShdTransDate,MAX(CASE WHEN FNDctNo IN (1) THEN '1' ELSE '0' END) AS FTFlag"
    tSql = tSql & vbCrLf & "                        From TPSTSalCD WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "                        WHERE  FTTmnNum='" & tVB_TerminalNoRet & "' AND FTShdTransNo='" & tVB_TransNoRet & "'"
    tSql = tSql & vbCrLf & "                        AND FTShdTransType ='" & tVB_TransTypeRet & "' AND FDShdTransDate='" & Format(dVB_TransDateRet, "YYYY/MM/DD") & "'"
    tSql = tSql & vbCrLf & "                        GROUP BY FNSdtSeqNo,FTTmnNum,FTShdTransNo,FDShdTransDate"
    tSql = tSql & vbCrLf & "                ) CD"
    tSql = tSql & vbCrLf & "     ON DT.FTTmnNum = CD.FTTmnNum AND"
    tSql = tSql & vbCrLf & "      DT.FTShdTransNo = CD.FTShdTransNo AND"
    tSql = tSql & vbCrLf & "      DT.FNSdtSeqNo = CD.FNSdtSeqNo AND"
    tSql = tSql & vbCrLf & "     DT.FDShdTransDate = CD.FDShdTransDate"
    tSql = tSql & vbCrLf & "WHERE DT.FTTmnNum='" & tVB_TerminalNoRet & "' AND DT.FTShdTransNo='" & tVB_TransNoRet & "'"
    tSql = tSql & vbCrLf & "AND  DT.FTShdTransType ='" & tVB_TransTypeRet & "' AND DT.FDShdTransDate='" & Format(dVB_TransDateRet, "YYYY/MM/DD") & "'"
    tSql = tSql & vbCrLf & "AND  ISNULL(DT.FTSdtCancelFlag,'0') <>'1'"
   ' tSql = tSql & vbCrLf & "AND  ISNULL(DT.FCSdtReturnQty ,0) < DT.FCSdtQty"
    tSql = tSql & vbCrLf & "ORDER BY DT.FNSdtSeqNo"

'    If SP_SQLvExecute(tSql, orsVB_DTReturn, IIf(bVB_ReturnFrmBck, ocnDBBck, oDB.ocnOnLine), , , , , True) Then Call SP_DATxRsNothing(orsVB_DTReturn): Exit Sub 'PH 2.0.0 ติดต่อฐานข้อมูล Server
    If SP_SQLvExecute(tSql, orsVB_DTReturn, oDB.ocnOnLine, , , , , True) Then Call SP_DATxRsNothing(orsVB_DTReturn): Exit Sub 'PH 2.0.0 ติดต่อฐานข้อมูล Server
    Set ogdForm.DataSource = orsVB_DTReturn
    ogdForm.Refresh
    For nI = 0 To orsVB_DTReturn.Fields.Count - 1
        ogdForm.ColKey(nI) = orsVB_DTReturn.Fields(nI).Name
    Next nI
    If ogdForm.rows < 11 Then ogdForm.rows = 11
    ogdForm.FixedCols = 1
    Call W_SetGridHeading
    With ogdForm
        .ColHidden(.ColIndex("FTFlag")) = True
        .ColHidden(.ColIndex("FCSdtReturnQty")) = True
        For nI = .FixedRows To .rows - 1
            If .TextMatrix(nI, .ColIndex("FTSkuAbbName")) <> "" Then
                Select Case .ValueMatrix(nI, .ColIndex("FTFlag"))
                    Case "1"
                        .TextMatrix(nI, .ColIndex("FTSkuAbbName")) = "*" & .TextMatrix(nI, .ColIndex("FTSkuAbbName"))
                    Case Else
                         If Val(Format(.ValueMatrix(nI, .ColIndex("FCSdtReturnQty")))) <> Val(Format(.ValueMatrix(nI, .ColIndex("FNSdtQty")))) Then
                            .TextMatrix(nI, .ColIndex("FBChk")) = 1
                        End If
                End Select
                If Val(Format(.ValueMatrix(nI, .ColIndex("FCSdtReturnQty")))) = Val(Format(.ValueMatrix(nI, .ColIndex("FNSdtQty")))) Then
                    '.Cell(flexcpForeColor, nI, .FixedCols, nI, .Cols - 1) = vbRed
                    .Cell(flexcpFontStrikethru, nI, .FixedCols, nI, .Cols - 1) = True
                     nW_CountNoRe = nW_CountNoRe + 1
                End If
            End If
        Next nI
         
         
         orsVB_DTReturn.Filter = adFilterNone
         orsVB_DTReturn.MoveFirst
    End With
    ogdForm.Enabled = True
    ogdForm.HighLight = flexHighlightAlways
    ogdForm.Select 1, 1                        'move highlight to first row
    'Set orsTemp = Nothing
    Exit Sub
ErrHandle:
     'Set orsTemp = Nothing
     Call SP_DATxRsNothing(orsVB_DTReturn)
End Sub


Public Sub W_SetGridHeading()
'----------------------------------------------------------
'   Call:
'   Cmt:   set heading of selling grid
'----------------------------------------------------------
    Const tFed = "FNSdtSeqNo,FBChk, FTSdtBarCode,FTSkuAbbName,FNSdtQty,FCSdtSaleAmt"
    Dim tColW$, iCol%

    tColW = "550;550;3105;3105;645;1350"
    iCol = 6
    With ogdForm
        Call SP_GRDxSetProperty(ogdForm, tFed, tColW, iCol, True, 0)
        .TextMatrix(.FixedRows - 1, .ColIndex("FNSdtSeqNo")) = UT_tToken(tCS_ColNo, ";", UT_nPosCutTag)                         'No
        .TextMatrix(.FixedRows - 1, .ColIndex("FBChk")) = UT_tToken(tCS_ColChk, ";", UT_nPosCutTag)      'Select
         .TextMatrix(.FixedRows - 1, .ColIndex("FTSdtBarCode")) = UT_tToken(tCS_ColBarCode, ";", UT_nPosCutTag)     'SKU Name
        .TextMatrix(.FixedRows - 1, .ColIndex("FTSkuAbbName")) = UT_tToken(tCS_ColProductName, ";", UT_nPosCutTag)     'Product Name
        .TextMatrix(.FixedRows - 1, .ColIndex("FNSdtQty")) = UT_tToken(tCS_ColQty, ";", UT_nPosCutTag)                                'Qty
        .TextMatrix(.FixedRows - 1, .ColIndex("FCSdtSaleAmt")) = UT_tToken(tCS_ColSubTotal, ";", UT_nPosCutTag)                 'Sub Total
        .ColHidden(5) = False
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    bVB_Return = False
    bVB_ReturnAll = False
    bVB_ReturnAllBill = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)          '*AO51-03-17
    Select Case KeyCode
        Case 13: Call ocmOK_Click
    End Select
End Sub

Private Sub Form_Load()
    bW_Show = False
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    With ogdForm
        .FixedRows = 1
        .FixedCols = 0
        .rows = 1
        .Cols = 7
    End With
    Call W_FRMxShowItemsReturn
    bW_Show = True
    
    '55-07-20 Neung TK
    ocmCleAll.Visible = False
    ocmSelAll.Visible = True
    
    Call SP_OBJxSetFocus(ocmOk)
End Sub

Private Sub ocmSelAll_Click() 'TK
    Dim nL As Integer
    With ogdForm
        For nL = 1 To .rows - 1
                If .TextMatrix(nL, .ColIndex("FTSdtBarCode")) <> "" Then
                        .Cell(flexcpChecked, nL, .ColIndex("FBChk"), nL, .ColIndex("FBChk")) = vbChecked
                End If
        Next nL
    End With
    ocmCleAll.Visible = True
    ocmSelAll.Visible = False
End Sub

Private Sub ocmCleAll_Click() 'TK
    With ogdForm
        .Cell(flexcpChecked, .FixedRows, 1, .rows - 1, 1) = vbUnchecked
    End With
    ocmCleAll.Visible = False
    ocmSelAll.Visible = True
End Sub

Private Sub ocmOK_Click()
    Call W_FRMxReferItemReturn
End Sub

Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    'หมายเหตุ : การที่ไม่ใช้ propertics .ColIndex("XXX") เพราะว่า Event นี้จะทำงานทันที่ที่มีการ Load Form เกิดขึ้น
    '*KT 52-01-23 Communication Sheet PH2-ML-SF-069 Return refer sale can't scroll down on item selection screen.xls
    With ogdForm
        If .TextMatrix(NewRow, 0) = "" Then
            .Select OldRow, 1
        Else
            .TopRow = NewRow
        End If
    End With
End Sub

'*KT 66-08-17 TTK-ISS3000-839
'---------------OLD-------------------------
'Private Sub ogdForm_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)     '*AO51-03-17    Comm.Sheet SF-091
'    With ogdForm
'         Select Case Col
'            Case .ColIndex("FBChk")
'                If .TextMatrix(Row, .ColIndex("FTSdtBarCode")) = "" Then
'                    Cancel = True
'                End If
'        End Select
'    End With
'End Sub
'---------------OLD------------------------

Private Sub ogdForm_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)     '*AO51-03-17    Comm.Sheet SF-091
    With ogdForm
         Select Case Col
            Case .ColIndex("FBChk")
'*KT 62-09-04 POSC-ML-0570 CR อนุญาตให้คืนสินค้า Deal เป็นชุดได้ ถ้าคืนต้องคืนทั้งชุด ถ้าไม่คืนต้องไม่คืนทั้งชุด  เดิมถ้ามี Deal/Free ต้องคืนทั้งหมด
'                If bDealNo Then   '*KT 62-05-28 ML-POSC-0670   Return refer sale(full bill is has deal/free)
'                  '  If bW_Show Then Call SP_MSGnShowing("บิลคืนนี้มีรายการ Deal/Free ต้องบังคับคืนสินค้าทั้งบิล;The bill of return for this Deal / Free item must be forced to return the entire product bill.", nCS_Information)
'                    Cancel = True
'                    Exit Sub
'                End If
                If .TextMatrix(Row, .ColIndex("FTSdtBarCode")) = "" Then
                    Cancel = True
                ElseIf Val(Format(.ValueMatrix(Row, .ColIndex("FCSdtReturnQty")))) = Val(Format(.ValueMatrix(Row, .ColIndex("FNSdtQty")))) Then
                    Cancel = True
                   ' Call SP_MSGnShowing("This item not allow for return because is Already returned.;This item not allow for return because is Already returned.", nCS_Warning)
'*KT 62-09-04 POSC-ML-0570 CR อนุญาตให้คืนสินค้า Deal เป็นชุดได้ ถ้าคืนต้องคืนทั้งชุด ถ้าไม่คืนต้องไม่คืนทั้งชุด  เดิมถ้ามี Deal/Free ต้องคืนทั้งหมด
'                ElseIf .TextMatrix(Row, .ColIndex("FTDeal")) = "Y" Then   '*KT 62-05-28 ML-POSC-0670   Return refer sale(full bill is has deal/free)
'                    Cancel = True
                End If
        End Select
    End With
End Sub

Private Sub ogdForm_Click()
'***KT 51-11-05
    With ogdForm
         Select Case .Col
             Case .ColIndex("FBChk")
                    Select Case CLng(Val(Format(.TextMatrix(.Row, .ColIndex("FTFlag")))))
                         Case 1
                            If bW_Show Then
                                Call SP_MSGnShowing("tms_CN005149", nCS_Warning)
                               .Cell(flexcpChecked, .Row, .Col, .Row, .Col) = vbUnchecked
                            End If
                    End Select
        End Select
    End With
End Sub

Private Sub ogdForm_KeyPress(KeyAscii As Integer)
    With ogdForm
         Select Case .Col
             Case .ColIndex("FBChk")
                Select Case KeyAscii
                    Case 32
                        Select Case CLng(Val(Format(.TextMatrix(.Row, .ColIndex("FTFlag")))))
                             Case 1
                                If bW_Show Then Call SP_MSGnShowing("tms_CN005149", nCS_Warning)
                                KeyAscii = 0
                        End Select
                End Select
             Case Else
                 KeyAscii = 0
        End Select
    End With
End Sub
