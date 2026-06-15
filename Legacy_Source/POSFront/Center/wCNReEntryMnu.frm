VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNReEntryMnu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re Entry Menu;klg_Title"
   ClientHeight    =   5370
   ClientLeft      =   4080
   ClientTop       =   3135
   ClientWidth     =   6600
   ControlBox      =   0   'False
   Icon            =   "wCNReEntryMnu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6600
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
      Left            =   3495
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   4710
      Width           =   1335
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1080
      Index           =   2
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4575
      Width           =   6600
      _Version        =   720897
      _ExtentX        =   11642
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Timer otmForm 
         Interval        =   2000
         Left            =   2520
         Top             =   120
      End
      Begin VB.CommandButton ocmCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel;klg_Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   120
         Width           =   1455
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdForm 
      Height          =   4020
      Left            =   0
      TabIndex        =   2
      Top             =   555
      Width           =   6600
      _cx             =   11642
      _cy             =   7091
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
      Rows            =   16
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNReEntryMnu.frx":058A
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
      Width           =   6600
      _Version        =   720897
      _ExtentX        =   11642
      _ExtentY        =   1482
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Re Entry Menu;klg_Menu"
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
         Width           =   6375
      End
   End
End
Attribute VB_Name = "wCNReEntryMnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey As String

Private Sub W_PRCxOpenTaxInclude()
'--------------------------------------------------------
'Call:
'Cmt: เพิ่มให้เรียก Open Tax Include '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
'-------------------------------------------------------
    tVB_CNMnuCurrent = "1720"
    tVB_Transaction = aVB_TransMode(47).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(47).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbOpenTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbOpenTaxInc - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(47).PrintType
     If Not bVB_CNCreateTblTmp Then
        tVB_Running = SP_AUTtGenCode("TPSTSalHD") '*KT 53-02-2010  PH3-ML-SF-088 POSFront V3.0.18 Open error(-2147217865) เนื่องจากมีการย้าย Code การสร้าง Table Temp ไปไว้ที่หน้าจอ LogIn แทน ในวันที่ *KT 53-02-03
    Else
         tVB_Running = SP_AUTtGenCode(tVB_CNTblHD)
    End If
    wCNLogIn.Show vbModal
    If tVB_CNUserAlwC <> "" Then
        wCNLogInReEntry.Show vbModal
    End If
End Sub

Private Sub W_PRCxOpenTaxManual()
'--------------------------------------------------------
'Call:
'Cmt: เพิ่มให้เรียก Open Tax Include '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
'-------------------------------------------------------
    tVB_CNMnuCurrent = "1721"
    tVB_Transaction = aVB_TransMode(48).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(48).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbOpenManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbOpenManualTax - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(48).PrintType
     If Not bVB_CNCreateTblTmp Then
        tVB_Running = SP_AUTtGenCode("TPSTSalHD") '*KT 53-02-2010  PH3-ML-SF-088 POSFront V3.0.18 Open error(-2147217865) เนื่องจากมีการย้าย Code การสร้าง Table Temp ไปไว้ที่หน้าจอ LogIn แทน ในวันที่ *KT 53-02-03
    Else
         tVB_Running = SP_AUTtGenCode(tVB_CNTblHD)
    End If
    wCNLogIn.Show vbModal
    If tVB_CNUserAlwC <> "" Then
        wCNLogInReEntry.Show vbModal
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim nMaxRow As Long
    nMaxRow = ogdForm.Aggregate(flexSTMax, 0, 0, ogdForm.rows - 1, 0)
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
                        If CLng(tW_HotKey) >= nMaxRow Then tW_HotKey = "" ': otmForm.Enabled = False
                    End If
                End With
        End Select
    End If
End Sub

Private Sub Form_Load()
    otmForm.Enabled = True
    tW_HotKey = ""
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FNCxAddFunc2Grid
    Call SP_STAxMouse(Me, False)
    tUT_SaleDate = UT_tStringDate(SP_GETdSaleDate)
    Call SP_MOVxEJToOnlineRE(Format(CDate(tUT_SaleDate), "yyMMdd"))         '*Ao 2013/10/24
    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล
    
    '*Tao 60-01-20 เช็คก่อนว่าพิมพ์ Closing Report หรือยัง
    If (Not bUT_SignOut) And bVB_FirstLoad Then
        Select Case SP_ChkxClosingMissing
        Case "41", "46"
            bVB_CNCloseTmnSta = True
            bVB_CNClosingRptAgain = True
            Call ocmOK_Click
        End Select
    End If
    '========================================
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call SP_GETbTmnData
'    tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
'    tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
'    tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
'    tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
'    tVB_CNTblDTTmp = "TPSDT" & tVB_CNTerminalNum & "Tmp"
'    tVB_CNTblCDTmp = "TPSCD" & tVB_CNTerminalNum & "Tmp"
    Call SP_MOVxEJToOnlineRE(Format(CDate(tUT_SaleDate), "yyMMdd"))         '*Ao 2013/10/24
    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    tVB_Transaction = 99
    tVB_TransactionSub = 99
    bVB_ReEntryMode = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim tTranType As String
    '------------------------------------------------------------------------------------
    otmForm.Enabled = False
    tUT_SaleDate = UT_tStringDate(SP_GETdSaleDate)
    Call SP_MOVxEJToOnline              '*Ao 2013/10/24
    '------------------------------------------------------------------------------------
    
    '*KT 58-05-27 TK-ISS3000-587 ป้องกันการไม่พิมพ์ใบ Closing Report
     'ให้ทำการตรวจสอบเพื่อทำการ Print Closing Report ก่อน
     '***********************************************************
     ' bVB_CNCloseTmnSta : สถานะทำ Close Terminal
     ' bVB_CNClosingRptSta : สถานะทำ Closing Report
     ' bVB_CNClosingRptSta : สถานะ bVB_CNClosingRptAgain แจ้งให้ทำงานเข้าสู่การตรวจสอบว่าต้องทำ Closing Report หรือ Reprint Closing Report
     If bVB_CNCloseTmnSta And (Not bVB_CNClosingRptSta) And bVB_CNClosingRptAgain Then
        'ทำเอกสาร Close Terminal แล้ว แต่ไม่สามารถทำ Closing Report
        Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:ocmOK_Click[Print Closing Report again]")
         tTranType = SP_CHKtClosingReport
         If tTranType = "8" Then '*TON 58-12-01 TK-ISS3000-581
            '*Tao 60-01-17 ใช้ Load ก่อน( กัน error แล้วเด้งออก)
            If Not SP_CHKxIsFormOpen("wCNCloseMnu") Then
                Call SP_GETxMainMnu
                Load wCNCloseMnu
            Else
                If wCNCloseMnu.ogdFunc.Aggregate(flexSTMax, 0, 0, wCNCloseMnu.ogdFunc.rows - 1, 0) < 8 Then
                    Unload wCNCloseMnu
                    Call SP_GETxMainMnu
                    Load wCNCloseMnu
                End If
            End If
            '======================================
            wCNCloseMnu.ogdFunc.Select 7, 0 '*TON 58-12-01 TK-ISS3000-581
            wCNCloseMnu.ocmOK_Click
        Else
            Call W_PRCxReprint
        End If
        Exit Sub
     End If
     '***********************************************************
     
    Select Case ogdForm.Row
        Case 0
'            Call W_PRCxSaleTaxInclude
            Call W_PRCxSaleManualTax '*Eaw 56-07-03 CommSheet TK-ISS3000-245-TR.xls
'            Call W_PRCxOpenTaxManual '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
        Case 1
'            Call W_PRCxSaleManualTax
            Call W_PRCxSaleTaxInclude '*Eaw 56-07-03 CommSheet TK-ISS3000-245-TR.xls
'            Call W_PRCxOpenTaxInclude '*Eaw 56-08-20 CommSheet TK-ISS3000-316-UAT_Req.xls
        Case 2
            Call W_PRCxReprint
'            otmForm.Enabled = True
        Case 3
            Call W_PRCxExit
    End Select
    'Unload Me
End Sub

Private Sub ogdForm_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_HotKey = ""
End Sub

Private Sub ogdForm_DblClick()
    Call ocmOK_Click
End Sub

Private Sub W_PRCxCallFunc(ByVal pnFuc As Long)
    Dim bFound As Boolean
    Dim nI As Long
On Error GoTo ErrHandle
    With ogdForm
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

Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call:
'Cmt: ทำการแสดงรายการของ Admin Menu
'-------------------------------------------------------
Dim nRow%
Dim tSmnID$
With ogdForm
    .rows = 6
'    .TextMatrix(0, 1) = Trim(SP_GETtVocabularyTicketFromArray("160601", aVB_MainMnuAll, tSmnID)) '
'    .TextMatrix(1, 1) = Trim(SP_GETtVocabularyTicketFromArray("160602", aVB_MainMnuAll, tSmnID)) '
    '*Eaw 56-07-03 CommSheet TK-ISS3000-245-TR.xls
    .TextMatrix(0, 1) = Trim(SP_GETtVocabularyTicketFromArray("160602", aVB_MainMnuAll, tSmnID)) '
    .TextMatrix(1, 1) = Trim(SP_GETtVocabularyTicketFromArray("160601", aVB_MainMnuAll, tSmnID)) '
    '------------------------------------------------------
    .TextMatrix(2, 1) = Trim(SP_GETtVocabularyTicketFromArray("160603", aVB_MainMnuAll, tSmnID)) '
    .TextMatrix(3, 1) = Trim(SP_GETtVocabularyTicketFromArray("160604", aVB_MainMnuAll, tSmnID)) '
    For nRow = 0 To 3
            .TextMatrix(nRow, 1) = " " & nRow + 1 & "." & .TextMatrix(nRow, 1)
            .TextMatrix(nRow, 0) = nRow + 1
    Next nRow
End With
End Sub

Private Sub W_PRCxSaleTaxInclude()
    Dim tSql As String
    If SP_TBLbChkExist("TPSTClose", oDB.ocnOffLine) Then '*Eaw 56-08-30
        tSql = "TRUNCATE TABLE  TPSTClose "
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    tVB_CNMnuCurrent = "1720"
'    wCNAllow.tW_ID = "1720"
    'wCNAllow.Show vbModal
    tVB_Transaction = aVB_TransMode(37).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(37).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbOpen - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbOpen - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(37).PrintType
     If Not bVB_CNCreateTblTmp Then
        tVB_Running = SP_AUTtGenCode("TPSTSalHD") '*KT 53-02-2010  PH3-ML-SF-088 POSFront V3.0.18 Open error(-2147217865) เนื่องจากมีการย้าย Code การสร้าง Table Temp ไปไว้ที่หน้าจอ LogIn แทน ในวันที่ *KT 53-02-03
    Else
         tVB_Running = SP_AUTtGenCode(tVB_CNTblHD)
    End If
    wCNLogIn.Show vbModal
    If tVB_CNUserAlwC <> "" Then
        wCNLogInReEntry.Show vbModal
    End If
End Sub

Private Sub W_PRCxSaleManualTax()
    Dim tSql As String
    If SP_TBLbChkExist("TPSTClose", oDB.ocnOffLine) Then '*Eaw 56-08-30
        tSql = "TRUNCATE TABLE  TPSTClose "
        Call SP_SQLvExecute(tSql, , oDB.ocnOffLine)
    End If
    tVB_CNMnuCurrent = "1721"
'    wCNAllow.tW_ID = "1721"
'    wCNAllow.Show vbModal
    tVB_Transaction = aVB_TransMode(42).ID              '*Ao 2013/07/20     CommSheet TK-ISS3000-284 TR-FAPL-064.xls
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(42).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(42).PrintType
    wCNLogIn.Show vbModal
    '-----------------------------------------------------------------------------------------------
    If tVB_CNUserAlwC <> "" Then
        wCNLogInReEntry.Show vbModal
    End If
End Sub

Private Sub W_PRCxReprint()
'    tVB_Transaction = "99"   '*Eaw 56-06-28
'    tVB_TransactionSub = "99"   '*Eaw 56-06-28
    tVB_CNMnuCurrent = "160603"
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbRePrintClose - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbRePrintClose - 1).UsrLangSec
    wCNAllow.tW_ID = "160603"
    wCNAllow.Show vbModal
    If tVB_CNUserAlwC <> "" Then
        Call wFunctionKB.W_RePrintSignOut(True)
    End If
End Sub

Private Sub W_PRCxExit()
    
    tVB_Transaction = 99
    tVB_TransactionSub = 99
    bVB_ReEntryMode = False
    Call SP_GETbTmnData
    tVB_CNTblHD = "TPSHD" & tVB_CNTerminalNum
    tVB_CNTblDT = "TPSDT" & tVB_CNTerminalNum
    tVB_CNTblRC = "TPSRC" & tVB_CNTerminalNum
    tVB_CNTblCD = "TPSCD" & tVB_CNTerminalNum
    tVB_CNTblDTTmp = "TPSDT" & tVB_CNTerminalNum & "Tmp"
    tVB_CNTblCDTmp = "TPSCD" & tVB_CNTerminalNum & "Tmp"
    tUT_SaleDate = UT_tStringDate(SP_GETdSaleDate)
    Call SP_MOVxEJToOnlineRE(Format(CDate(tUT_SaleDate), "yyMMdd"))         '*Ao 2013/10/24         '*Ao 2013/10/24
    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล
    Unload Me
End Sub

'Private Sub otmForm_Timer()
'On Error GoTo ErrHandle
'
'    'Call SP_MOVxEJToOnline
'    '*Bump 57-01-13 เพิ่มการเช็ค ONLINE ก่อนเรียกส่ง EJ
'    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
'        MsgBox "ONLINE : " & tUT_SaleDate
'        Call SP_MOVxEJToOnline
'    Else
'        MsgBox "OFFLINE : " & tUT_SaleDate
'    End If
'
'    Exit Sub
'ErrHandle:
'
'End Sub

Private Sub otmForm_Timer()
On Error GoTo ErrHandle

    'Call SP_MOVxEJToOnline
    '*Bump 57-01-13 เพิ่มการเช็ค ONLINE ก่อนเรียกส่ง EJ
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        Call SP_MOVxEJToOnline
    End If

    Exit Sub
ErrHandle:
End Sub
