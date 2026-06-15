VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNReEntryMnu2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re Entry Menu;klg_Title"
   ClientHeight    =   5325
   ClientLeft      =   7410
   ClientTop       =   1335
   ClientWidth     =   5670
   Icon            =   "wCNReEntryMnu2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5670
   Begin VSFlex7Ctl.VSFlexGrid ogdFunc 
      Height          =   3910
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   5760
      _cx             =   10160
      _cy             =   6897
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
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
      Rows            =   8
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNReEntryMnu2.frx":058A
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
      Height          =   1080
      Index           =   2
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4560
      Width           =   5760
      _Version        =   720897
      _ExtentX        =   10160
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
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
         Left            =   2640
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Timer otmForm 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer otmOnOff 
         Interval        =   3000
         Left            =   720
         Top             =   0
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
         Left            =   4080
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   120
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   840
      Index           =   0
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   5760
      _Version        =   720897
      _ExtentX        =   10160
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
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   0
         Width           =   5535
      End
   End
End
Attribute VB_Name = "wCNReEntryMnu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nW_TransMnu As Long
Private tW_TransFunc  As String
Private tW_SmnID  As String
Private tW_HotKey As String
Private bW_FirstKeyEsc  As Boolean
Private bW_KeyEnter As Boolean '***KT 51-11-20 ป้องกันการ Enter ซ้ำและมีการบันทึก  มากกว่า 1 ครั้งใน
Private nMaxRow As Long

Public Sub W_PRCxTrnMnuEJViewer()
'-----------------------------------------------------------
'   Call:
'   Cmt:  '*Ao 2014/10/17  V 4.2.151 CommSheet TK-ISS3000-561
'-----------------------------------------------------------
On Error GoTo ErrHandle
    tVB_CNMnuTemp = ogdFunc.TextMatrix(0, 3)             '*Ao 2014/11/13   V 4.2.157
    
    tVB_CNMnuCurrent = tW_SmnID
    If aVB_TransMnu(nW_TransMnu - 1).Autho Then
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                bW_KeyEnter = False
                Exit Sub
            End If
    End If
    
    Unload Me
    wCNEJViewer.Show vbModal
    tW_SmnID = tVB_CNMnuTemp
    Call CallByName(wCNReEntryMnu2, "W_PRCxTrnMnuSaleReEntry", VbMethod)        '*Ao 2014/11/13   V 4.2.157
    Exit Sub
ErrHandle:
    Call CallByName(wCNReEntryMnu2, "W_PRCxTrnMnuSaleReEntry", VbMethod)            '*Ao 2014/11/13   V 4.2.157
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    Dim nMaxRow As Long
'    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
    If KeyAscii = 13 Then
        Call ocmOK_Click
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
            Case vbKeyEscape
                If bW_FirstKeyEsc Then
                    bW_FirstKeyEsc = False
                Else
                    ogdFunc.Row = ogdFunc.Row
                    ogdFunc.Select nVB_HisSelTrn, 0, nVB_HisSelTrn, 0
                    Call ocmOK_Click
                End If
        End Select
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[1]")
      bW_FirstKeyEsc = True
      bW_KeyEnter = False '***KT 51-12-04 V 1.0.2
      Call SP_FRMxChildCenter(Me)
      Call SP_FRMxGetCaptionFrmDB2Tag(Me)
      Call SP_FRMxSetCapForm(Me)
      Call W_FNCxAddFunc2Grid
      Call SP_DELxRecordBlank
     nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
     Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[2]")
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:Form_Load[3]")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Unload[1]")
    otmForm.Enabled = False
    otmOnOff.Enabled = False
    Call SP_OBJxSetNothing(Me)
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Unload[2]")
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:Form_Unload[3]")
End Sub

Private Sub ocmCancel_Click()
On Error GoTo ErrHandle '*KT 54-04-08 ป้องกันโปรแกรมหลุด V.3.0.56
    With ogdFunc
        nW_TransMnu = Val(Format(.TextMatrix(0, 4)))
        tW_TransFunc = .TextMatrix(0, 2)
        tW_SmnID = .TextMatrix(0, 3)
        wMain.bW_FirstLoad = False
        If tW_TransFunc <> "" Then '*KT 54-04-08 V.3.0.56 ป้องการโปรแกรมหลุดจากการเรียก CallByName
            Call CallByName(wCNReEntryMnu2, tW_TransFunc, VbMethod)
        End If
    End With
    Exit Sub
ErrHandle:
     Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:ocmCancel_Click")
End Sub

Private Sub ocmOK_Click()
    '*KT 52-03-24 PH1.5-ML-SF-088 *KT 51-12-04 V 1.0.2
    '------------------------------------------------------------------------------------
    Dim tFile As String
    Dim oFSO As New FileSystemObject
On Error GoTo ErrHandle
    If bW_KeyEnter Then
        Exit Sub                                    '* ไม่ให้กด Enter ซ้ำๆหลายครั้ง
    Else
        bW_KeyEnter = True
    End If
    '-------------------------------------------------------------------------------------
    With ogdFunc
        nW_TransMnu = Val(Format(.TextMatrix(0, 4)))
        tW_TransFunc = .TextMatrix(.Row, 2)
        tW_SmnID = .TextMatrix(.Row, 3)
        wMain.bW_FirstLoad = False
        nVB_HisSelTrn = 0 'Default Sale Transaction
        bVB_CNUnLock = False '*Eaw 56-09-04
        If bVB_CNUnLock And tW_SmnID <> "1715" Then
            Call SP_MSGnShowing("tms_CN005167", nCS_Error)
            bW_KeyEnter = False
        Else
            If tW_TransFunc = "W_PRCxTrnMnuClose" Then
                otmForm.Enabled = False
                otmOnOff.Enabled = False                            '*Ao 2011/04/08    V 3.0.56
                Call SaveSetting("AdaPosMall", "POSFront", "RepStop", "1") '*KT 52-07-27 เริ่มให้ ServiceAutoReplicate.EXE หยุดทำงาน '*KT 54-03-28 ป้องกันโปรแกรมหลุด
                Call SP_CRTxWriteLogTerminated("******************START DEBUG CLOSE TERMINAL********************")
                '*KT 54-04-08 V.3.0.56 ตรวจสอบ ServiceOnOff.exe ว่ายังทำงานหรือไม่
                'เปิด EXE ที่ทำหน้าที่เป็น Service ในการเก็บค่า Offline Online
                '-------------------------------------------------------------------------------------
                tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceOnOff.exe"
                If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                        tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
                        tFile = tFile + "\POS\ServiceOnOff.exe"
                        If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                            Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
                            Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Can not run ServiceOnOff.exe" & " Form:" & Me.Name & " Events:ocmOK_Click")
                            'Call SP_APPxEndProgram
                        End If
                End If
                nVB_lServiceOnOffWnd = FindWindow(vbNullString, "ServiceOnOff")
                If nVB_lServiceOnOffWnd = 0 Then
                      'Shell tFile, vbHide
                    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "ShellEx ServiceOnOff.exe" & " Form:" & Me.Name & " Events:ocmOK_Click")   '*KT 52-12-17 POSFront 2.0.30
                    Call ShellEx(nVB_lServiceOnOffWnd, "open", tFile, "", "", 10)
                End If
                 '-------------------------------------------------------------------------------------
            End If
             If tW_TransFunc <> "" Then '*KT 54-04-08 V.3.0.56 ป้องการโปรแกรมหลุดจากการเรียก CallByName
                Call CallByName(wCNReEntryMnu2, tW_TransFunc, VbMethod)
            Else
                otmForm.Enabled = True
                otmOnOff.Enabled = True                            '*KT 2011/04/08    V 3.0.56
            End If
            bW_KeyEnter = False
            Set oFSO = Nothing
        End If
    End With
    Call wMain.W_SETxIndicator(0)  '*BG 57-05-22 HOT KEY
    
    Exit Sub
ErrHandle:
     Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:ocmOK_Click")
     Set oFSO = Nothing
     bW_KeyEnter = False
End Sub

Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call:
'Cmt: *KT 50-12-24 ทำการดึงข้อมูลใน Array มาแสดงที่หน้าจอว่ามีรายการ Main Menu ไหนบ้าง
'-------------------------------------------------------
    Dim nI As Long
    Dim nR As Long
    With ogdFunc
        nR = LBound(aVB_TransMnu)
      .rows = UBound(aVB_TransMnu)
        For nI = LBound(aVB_TransMnu) To UBound(aVB_TransMnu) - 1
'             If Len(aVB_TransMnu(nI).SmnID) = 6 And Mid(aVB_TransMnu(nI).SmnID, 1, 4) = tVB_CNReEntryIDMain And aVB_TransMnu(nI).SmnID <> "172006" And aVB_TransMnu(nI).SmnID <> "172106" Then
             If Len(aVB_TransMnu(nI).SmnID) = 6 And Mid(aVB_TransMnu(nI).SmnID, 1, 4) = tVB_CNReEntryIDMain Then
                .TextMatrix(nR, 0) = nR + 1
                .TextMatrix(nR, 1) = nR + 1 & "." & IIf(aVB_TransMnu(nI).NameUsr <> "", aVB_TransMnu(nI).NameUsr, aVB_TransMnu(nI).Name)
                .TextMatrix(nR, 2) = aVB_TransMnu(nI).FuncName
                .TextMatrix(nR, 3) = aVB_TransMnu(nI).SmnID
                .TextMatrix(nR, 4) = aVB_TransMnu(nI).SeqNo
                nR = nR + 1
            End If
        Next nI
        .rows = nR
    End With
End Sub

Public Sub W_PRCxTrnMnuSaleReEntry()
'--------------------------------------------------------
'Call:
'Cmt: การเข้าสู่ Mode Transaction การขายแบบปกติ
'-------------------------------------------------------
    tVB_CNMnuCurrent = tW_SmnID
    If aVB_TransMnu(nW_TransMnu - 1).Autho Then
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                bW_KeyEnter = False '***KT 51-12-04 V 1.0.2
                Exit Sub
            End If
    End If
    Select Case tVB_CNMnuCurrent
        Case "172001"
            tVB_Transaction = aVB_TransMode(37).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(37).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleTaxInc - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(37).PrintType
        Case "172101"
            tVB_Transaction = aVB_TransMode(42).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(42).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSaleManualTax - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(42).PrintType
        Case Else
            Exit Sub
    End Select
    Call SP_INSxTransHD(tVB_CNTblHD)
    
    Unload Me
    wMain.W_FRMxActivate
End Sub

Public Sub W_PRCxTrnMnuIPVReEntry()
'--------------------------------------------------------
'Call:
'Cmt: การยกเลิกรายการ(IPV or VOID Transaction) เป็นกายกเลิกรายการขายและรายการคืน กรณียังไม่ปิดการขายประจำวัน เมื่อพนักงานเก็บเงินบันทึกข้อมูลการขายหรือการคืนผิดพลาด
'โดยลูกค้าจะต้องนำบิลของรายการขายหรือรายการคืนมาด้วยและพนักงานเก็บเงิน ต้องบันทึกหมายเลขเครือง และหมายเลขลำดับการขายนั้นๆ ระบบจะดึงข้อมุลการขายหรือการคืนขึ้นมาเพื่อให้พนักงานเก็บเงินยืนยันการยกเลิก
'รายการ และระบบจะยกเลิกการบันทึกทั้งหมดของรายการที่ยกเลิกนั้น
'-------------------------------------------------------
'*KT 51-06-27 Comsheet ML-SF-344
'    If bVB_TrainingMode Then
'        Call SP_MSGnShowing("tms_CN005062", nCS_Information)
'        Exit Sub
'    End If

    tVB_CNMnuCurrent = tW_SmnID
    If aVB_TransMnu(nW_TransMnu - 1).Autho Then
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                bW_KeyEnter = False '***KT 51-12-04 V 1.0.2
                Exit Sub
            End If
    End If

    Select Case tVB_CNMnuCurrent
'        Case "172003"
'            tVB_Transaction = aVB_TransMode(39).ID
'            tVB_TransactionSub = tVB_Transaction
'            tVB_TransName = aVB_TransMode(39).Name
'            tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVTaxInc - 1).UsrLangSec
'            tVB_CNTypePrinter = aVB_TransMode(39).PrintType
'        Case "172103"
'            tVB_Transaction = aVB_TransMode(45).ID
'            tVB_TransactionSub = tVB_Transaction
'            tVB_TransName = aVB_TransMode(45).Name
'            tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVManualTax - 1).UsrLangSec
'            tVB_CNTypePrinter = aVB_TransMode(45).PrintType
        Case Else
            'none
    End Select
    Unload Me
    wMain.W_FRMxActivate
End Sub

Public Sub W_PRCxTrnMnuTotalReportReEntry()
'--------------------------------------------------------
'PH 2.0.0
'Call:
'Cmt: Transaction Total Report
'-------------------------------------------------------
    Dim tSentComplate$
    Dim tFile$
    'Dim oFSO As New FileSystemObject
    Dim nCount As Long
    Dim nCountBill As Long
    
On Error GoTo Err1 'PH2-ML-SF-078
    
'    If bVB_TrainingMode Then
'        Call SP_MSGnShowing("tms_CN005062", nCS_Information)
'        Exit Sub
'    End If

    tVB_CNMnuCurrent = tW_SmnID
   If aVB_TransMnu(nW_TransMnu - 1).Autho Then
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                bW_KeyEnter = False '***KT 51-12-04 V 1.0.2
                Exit Sub
            End If
    End If
    Call SP_STAxMouse(Me, True)
    
'*KT 52-09-18 3.0.0 ในการทำ Total Report ไม่จำเป็นต้องทำการตรวจสอบว่ามีการส่งข้อมุลขายไปที่ Server หมดแล้วหรือไม่ ก่อนออก Total Report เพราะตั้งแต่ Version 2.0 ขึ้นไป
'ยอดขายทั้งหมดอยู่ที่ Local และ การชี้ฐานข้อมุลก้อมีการชี้ที่ Local เป็นหลัง จากการตรวจสอบพบว่าในการออก Total Report ไม่มีดึงข้อมูลจาก Table Deposit,Voucher,Hold
'เพื่อความรวดเร้วในการทำงานจึงทำการตัดขั้นตอนการส่งข้อมูลไปยัง Server ออก ยกเว้นกรณีที่ทางการออก Total Report จะมีการใช้ข้อมูลที่ Server จึงต้องมีการทำการส่งข้อมูล
'ไปยัง Server ก่อนทำ Total Report
''    If ocnAll.ConnectionString = oDB.ocnOffLine.ConnectionString Then
''        If Not SP_TBKbMov2DB Then '*KT 51-08-19 Communication Sheet ML-SF-471 Backup sale data in terminal.xls
''                bW_KeyEnter = False
''                MsgBox "You can not Total Report because can not import to database offline(File dat)"
''                Exit Sub
''                'GoTo Err1 '*KT 52-02-02 ไม่มีการให้ทำงานซ้ำซ้อน  PH2-ML-SF-078
''        End If
''    Else
'        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
'            '*KT 51-10-31 ตรวจสอบทุกครั้งก่อนการ Close Terminal ต้องให้ ServiceTranfer.exe ทำการ Copy ข้อมูลจาก Local มา Server ให้หมดก่อน
'            '---------------------------------------------------------------
'            'ตรวจสอบ ServiceTranfer.exe ว่ายัง Run อยู่หรือไม่ ถ้าไม่ Run ให้ทำการ Run exe Auto
'            tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceTranfer.exe"
'            If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
'                    tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
'                    tFile = tFile + "\POS\ServiceTranfer.exe"
'                    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
'                        Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
'                        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Can not run ServiceTranfer.exe" & " Form:" & Me.Name & " Control:General" & " Events:W_PRCxTrnMnuTotalReport")
'                        Call SP_APPxEndProgram
'                    End If
'            End If
'            nVB_lServiceTransWnd = FindWindow(vbNullString, "ServiceTranfer")
'            If nVB_lServiceTransWnd = 0 Then
'                    'Shell tFile, vbHide
'                    Call ShellEx(nVB_lServiceTransWnd, "open", tFile, "", "", 10)
'            End If
'            DoEvents
'
'            tSentComplate = GetSetting("AdaPosMall", "POSFront", "SentSale")
'            If tSentComplate = "0" Then
'                Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")
'                tSentComplate = "1"
'            End If
'            '*KT 52-02-17
'            nCount = 0
'            nCountBill = GetSetting("AdaPosMall", "POSFront", "CountBill")
'            Do While tSentComplate = "1" And oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP)
'    '            SleepEx 1000, 0
'    '            tSentComplate = GetSetting("AdaPosMall", "POSFront", "SentSale")
'
'                nCount = nCount + 1
'                SleepEx 1000, 0
'                tSentComplate = GetSetting("AdaPosMall", "POSFront", "SentSale")
'                If nCountBill <= nCount Then
'                        nVB_lServiceTransWnd = FindWindow(vbNullString, "ServiceTranfer")
'                        If nVB_lServiceTransWnd = 0 Then
'                             MsgBox "ServiceTranfer not run"
'                             tSentComplate = 0
'                        End If
'                End If
'            Loop
'            '-----------------------------------------------------------------
'        End If
''    End If

'*KT 52-02-02 ไม่มีการให้ทำงานซ้ำซ้อน PH2-ML-SF-078
    Call SP_STAxMouse(Me, False)
   Select Case tVB_CNMnuCurrent
        Case "172003"
            tVB_Transaction = aVB_TransMode(39).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(39).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbTotalReportTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbTotalReportTaxInc - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(39).PrintType
        Case "172103"
            tVB_Transaction = aVB_TransMode(44).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(44).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbTotalReportManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbTotalReportManualTax - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(44).PrintType
        Case Else
            'none
    End Select
'    tVB_Transaction = aVB_TransMode(30).ID
'    tVB_TransactionSub = tVB_Transaction
'    tVB_TransName = aVB_TransMode(30).Name
'    tVB_TransNameABB = aVB_NameAbb(tEN_AbbTotalReport - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbTotalReport - 1).UsrLangSec
'    tVB_CNTypePrinter = aVB_TransMode(30).PrintType
    Unload Me
    wMain.W_FRMxActivate
    Exit Sub
Err1:

'*KT 52-02-02 ไม่มีการให้ทำงานซ้ำซ้อน PH2-ML-SF-078
'    Call SP_STAxMouse(Me, False)
'    tVB_Transaction = aVB_TransMode(30).ID
'    tVB_TransactionSub = tVB_Transaction
'    tVB_TransName = aVB_TransMode(30).Name
'    tVB_TransNameABB = aVB_NameAbb(tEN_AbbTotalReport - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbTotalReport - 1).UsrLangSec
'    Unload Me
'    wMain.W_FRMxActivate

    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_STAxMouse(Me, False)
    Exit Sub
End Sub

Public Sub W_PRCxTrnMnuReprintReEntry()
'--------------------------------------------------------
'Call:
'Cmt: Transaction Reprint Receipt
'-------------------------------------------------------
'*KT 51-06-27 Comsheet ML-SF-344
'    If bVB_TrainingMode Then
'        Call SP_MSGnShowing("tms_CN005062", nCS_Information)
'        Exit Sub
'    End If

    tVB_CNMnuCurrent = tW_SmnID
    If aVB_TransMnu(nW_TransMnu - 1).Autho Then
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                bW_KeyEnter = False '***KT 51-12-04 V 1.0.2
                Exit Sub
            End If
    End If
    Unload Me
    wCNReprint.Show vbModal
    If bVB_TrnReprint = False Then
        If bVB_ReEntryMode Then
             wCNReEntryMnu2.Show vbModal: Exit Sub
             Exit Sub
        Else
            wCNBegin.Show vbModal
            Exit Sub
        End If
    End If
End Sub

Public Sub W_PRCxTrnMnuNoSale()
'--------------------------------------------------------
'Call:
'Cmt: Transaction NoSale
'-------------------------------------------------------
'*KT 51-06-27 Comsheet ML-SF-344
'    If bVB_TrainingMode Then
'        Call SP_MSGnShowing("tms_CN005062", nCS_Information)
'        Exit Sub
'    End If
    tVB_CNMnuCurrent = tW_SmnID
    If aVB_TransMnu(nW_TransMnu - 1).Autho Then
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                bW_KeyEnter = False '***KT 51-12-04 V 1.0.2
                Exit Sub
            End If
    End If
    tVB_Transaction = aVB_TransMode(17).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(17).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbNoSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbNoSale - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(17).PrintType
     Unload Me
     wMain.W_FRMxActivate
End Sub

Public Sub W_PRCxTrnMnuTrain()
'--------------------------------------------------------
'Call:
'Cmt: Transaction Training Mode
'
'-------------------------------------------------------
    tVB_Transaction = aVB_TransMode(24).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(24).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbTrainingOn - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbTrainingOn - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(24).PrintType
    Unload Me
     wMain.W_FRMxActivate
    
End Sub

Public Sub W_PRCxTrnMnuOverShortReEntry()
    Dim bIns As Boolean
    '
On Error GoTo ErrHandle

'    If bVB_TrainingMode Then
'        Call SP_MSGnShowing("tms_CN005062", nCS_Information)
'        Exit Sub
'    End If

    tVB_CNMnuCurrent = tW_SmnID
   If aVB_TransMnu(nW_TransMnu - 1).Autho Then
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                bW_KeyEnter = False '***KT 51-12-04 V 1.0.2
                Exit Sub
            End If
    End If
    Select Case tVB_CNMnuCurrent
        Case "172004", "172005"
            tVB_Transaction = aVB_TransMode(40).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(40).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbOverShortTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbOverShortTaxInc - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(40).PrintType
        Case "172104", "172105"
            tVB_Transaction = aVB_TransMode(45).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(45).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbOverShortManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbOverShortManualTax - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(45).PrintType
        Case Else
            Exit Sub
    End Select
    wTenderMain.Show vbModal
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_STAxMouse(Me, False)
End Sub

Private Sub W_PRCxMissing()
'--------------------------------------------------------
'Call:
'Cmt: ทำการประมวลการเกิดเหตุการเลขที่ Ticket Missing
'-------------------------------------------------------
    Dim tSql$
    Dim tTrnST$
    Dim tTrnED$
    Dim tTrnNext$
    Dim orsTemp As ADODB.Recordset
    Dim bFound As Boolean
    bFound = False
    
    If Not SP_TBLbChkExist(tVB_CNTblHD) Then
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
    End If
    
    tSql = "SELECT  FTShdTransNo"
    tSql = tSql & " FROM " & tVB_CNTblHD
    tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & "  AND FTLogCode='" & tUT_LoginCode & "'"
    tSql = tSql & "  AND FTShdStaDoc = '1'"
    tSql = tSql & "  ORDER BY FTShdTransNo"
    Call SP_SQLvExecute(tSql, orsTemp)
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        tTrnNext = SP_FEDtChkString(orsTemp, "FTShdTransNo")
        Do While Val(tTrnNext) <= Val(tVB_Running)
            orsTemp.Filter = adFilterNone
            orsTemp.Filter = "FTShdTransNo = '" & tTrnNext & "'"
            If orsTemp.BOF And orsTemp.EOF Then
                If bFound = False Then
                    bFound = True
                    tTrnST = tTrnNext
                    tTrnED = tTrnNext
                Else
                    tTrnED = tTrnNext
                End If
            Else
                If bFound Then
                    bFound = False
                    tSql = "INSERT INTO TPSTLogMissing("
                    tSql = tSql & "FTStmCode, FTTmnCode, FTTmnNum, FTTmnName, FTTmnFloor"
                    tSql = tSql & ",FTMdtTransType , FTMdtTransNoSt, FTMdtTransNoEd, FNMdtNum,"
                    tSql = tSql & tCS_CNLastFedIns
                    tSql = tSql & " VALUES ("
                    tSql = tSql & "'" & tVB_CNStoreCode & "'"
                    tSql = tSql & ",'" & tVB_CNTerminalIndex & "'"
                    tSql = tSql & ",'" & tVB_CNTerminalNum & "'"
                    tSql = tSql & ",'" & tVB_CNComName & "'"
                    tSql = tSql & ",'" & tVB_CNTerminalFloor & "'"
                    tSql = tSql & ",'" & "0" & "'"
                    tSql = tSql & ",'" & tTrnST & "'"
                    tSql = tSql & ",'" & tTrnED & "'"
                    tSql = tSql & "," & SP_VALcText2Double(tTrnED) - SP_VALcText2Double(tTrnST) + 1 & ","
                    tSql = tSql & UT_tLastInsSql & ")"
                    Call SP_SQLvExecute(tSql)
                End If
            End If
                tTrnNext = Format(Val(tTrnNext) + 1, "00000")
        Loop
    End If
    Call SP_DATxRsNothing(orsTemp)
End Sub

Private Sub ogdFunc_DblClick()
    Call ocmOK_Click
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

Private Sub otmForm_Timer()
'    Dim nMaxRow As Long
'    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
    If IsNumeric(tW_HotKey) Then tW_HotKey = "": otmForm.Enabled = False
End Sub

Private Sub otmOnOff_Timer()
'-----------------------------------------------------------
'Call:
'Cmt: Refresh สัญลักษณ์เขียวแดงของ Network
'----------------------------------------------------------
    Dim bOnOff As Boolean
    Dim bOnOffCur As Boolean
    
    On Error GoTo ErrHandle
    
    bOnOff = IIf(wMain.olaOnOff.Caption = "On-line", True, False)
    bOnOffCur = oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP)
    If bOnOffCur Then
        wMain.oimOnOff.Picture = wMain.oilOnOff.ListImages(1).Picture
        wMain.olaOnOff.Caption = "On-line"
        If Not bOnOff And bOnOffCur Then
            Call SP_CHKbOnlineEnable 'PH 2.0.0
        End If
    Else
        wMain.oimOnOff.Picture = wMain.oilOnOff.ListImages(2).Picture
        wMain.olaOnOff.Caption = "Off-line"
    End If
    '    DoEvents           '*KT 53-01-14
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)               '*Ao 2011/04/08    V 3.0.56
End Sub

Public Sub W_PRCxTrnMnuReturnReEntry()
    tVB_CNMnuCurrent = tW_SmnID
    If aVB_TransMnu(nW_TransMnu - 1).Autho Then
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                bW_KeyEnter = False '***KT 51-12-04 V 1.0.2
                Exit Sub
            End If
    End If
    Select Case tVB_CNMnuCurrent
        Case "172002"
            tVB_Transaction = aVB_TransMode(38).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(38).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbReturnTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbReturnTaxInc - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(38).PrintType
        Case "172102"
            tVB_Transaction = aVB_TransMode(43).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(43).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbReturnManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbReturnManualTax - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(43).PrintType
        Case Else
            'none
    End Select

    Call SP_INSxTransHD(tVB_CNTblHD)
    Unload Me
    wCNReturnMenu.Show vbModal
    If bVB_CNAlwInputSpn Then wCNSaleManID.Show vbModal  '*BG 56-11-08 CommSheet TK-ISS3000-417-FAPL-148_BG.xlsx
    If bVB_Return = False Then
        If bVB_ReEntryMode Then
             wCNReEntryMnu2.Show vbModal: Exit Sub
             Exit Sub
        Else
            wCNBegin.Show vbModal
            Exit Sub
        End If
    End If
    wMain.W_FRMxActivate
End Sub

Public Sub W_PRCxTrnMnuCloseReEntry()
'--------------------------------------------------------
'PH 2.0.0
'Call:
'Cmt: Close Terminal
'          *KT 52-10-21 PH2-ML-SF-139 POSFront not Reponding
'-------------------------------------------------------
    Dim tSentComplate$
    Dim tFile$
    Dim oFSO As New FileSystemObject
    Dim nCount As Long
    Dim nCountBill  As Long
    Dim bTransfer As Boolean
    Dim bPass As Boolean
    Dim bIns As Boolean
    Dim bOnOff As Boolean
On Error GoTo ErrHandle

    Call W_PRCxTrnMnuOverShortReEntry
    If Not bVB_CNOverShortSta Then
        Exit Sub
    End If
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[1]")
    Call SP_STAxMouse(Me, True)
    If bVB_TrainingMode Then
        Call SP_MSGnShowing("tms_CN005062", nCS_Information)
        Call SP_STAxMouse(Me, False)
        bW_KeyEnter = False '***KT 51-12-04 V 1.0.2
        Exit Sub
    End If
    
    tVB_CNMnuCurrent = tW_SmnID
   If aVB_TransMnu(nW_TransMnu - 1).Autho Then
            If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                bW_KeyEnter = False '***KT 51-12-04 V 1.0.2
                Exit Sub
            End If
    End If
    'ก่อนการ Close Terminal จะต้องตรวจสอบการ Upload จาก Table Temp ไปยัง Table จริง
    bUT_SignOut = True
    bVB_CNCloseTmnSta = True
    
    Select Case tVB_CNMnuCurrent
        Case "172005"
            tVB_Transaction = aVB_TransMode(41).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(41).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbClosingReportTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbClosingReportTaxInc - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(41).PrintType
        Case "172105"
            tVB_Transaction = aVB_TransMode(46).ID
            tVB_TransactionSub = tVB_Transaction
            tVB_TransName = aVB_TransMode(46).Name
            tVB_TransNameABB = aVB_NameAbb(tEN_AbbClosingReportManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbClosingReportManualTax - 1).UsrLangSec
            tVB_CNTypePrinter = aVB_TransMode(46).PrintType
        Case Else
            'none
    End Select
    
    'bIns = False
    'Call SP_INSxTransHD(tVB_CNTblHD, , bIns)
    'Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[2]")
    'If bIns Then
        'Call wFunctionKB.W_SetToComplete
        'Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[3]")
        'Call SP_UPDxSetCompleteTerminalMtn
        'Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[4]")
        
        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
            '*KT 51-10-31 ตรวจสอบทุกครั้งก่อนการ Close Terminal ต้องให้ ServiceTranfer.exe ทำการ Copy ข้อมูลจาก Local มา Server ให้หมดก่อน
            '---------------------------------------------------------------
            'ตรวจสอบ ServiceTranfer.exe ว่ายัง Run อยู่หรือไม่ ถ้าไม่ Run ให้ทำการ Run exe Auto
            '*KT 51-10-02 3.0.127
            bTransfer = True
            tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceTranfer.exe"
            If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                    tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
                    tFile = tFile + "\POS\ServiceTranfer.exe"
                    If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                        Call SP_MSGnShowing("tms_CN005139", nCS_Warning, , , tFile & vbCrLf & cUT.UT_STRtToken(SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN004011"), ";", nVB_CNCutLang - 1))
                        bTransfer = False
                    End If
            End If
            If bTransfer Then
                '*KT 52-11-20 ทำการไล่ Commit Tran ของ Local
                Call SP_SQLxCommitTranLocal
                Call SP_SQLxCommitTranServer
                nVB_lServiceTransWnd = FindWindow(vbNullString, "ServiceTranfer")
                If nVB_lServiceTransWnd = 0 Then
                        'Shell tFile, vbHide
                        Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[4]")
                        Call ShellEx(nVB_lServiceTransWnd, "open", tFile, "", "", 10)
                End If
'                DoEvents            '*KT 53-01-14
                
                tSentComplate = GetSetting("AdaPosMall", "POSFront", "SentSale")
                If tSentComplate = "0" Then
                    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")
                    tSentComplate = "1"
                End If
                '*KT 52-02-17
                nCount = 0
                nCountBill = nVB_CNWaitTranfer '300 'GetSetting("AdaPosMall", "POSFront", "CountBill")
                bOnOff = oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP)
                Do While tSentComplate = "1" And bOnOff And nCountBill >= nCount
                    nCount = nCount + 1
                    SleepEx 1000, 0
                    tSentComplate = GetSetting("AdaPosMall", "POSFront", "SentSale")
                    If nCountBill <= nCount Then
                            nVB_lServiceTransWnd = FindWindow(vbNullString, "ServiceTranfer")
                            If nVB_lServiceTransWnd = 0 Then
                                 MsgBox "ServiceTranfer not run", vbCritical
                                 Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[5]")
                                 tSentComplate = 0
                            End If
                    End If
                Loop
                '*KT 52-11-17 เพิ่มให้มีการแจ้ง message และเก็บ Log กรณีที่ Close Terminal แล้วการส่งข้อมูลของ ServericeTranfer ไม่สำเร็จ
                If tSentComplate = "1" And nCountBill <= nCount Then
                    MsgBox "ServiceTranfer send data is incomplete, please contact the server administrator.", vbCritical
                    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[6]")
                ElseIf tSentComplate = "1" And Not bOnOff Then
                    MsgBox "ServiceTranfer send data is incomplete becase network is offline, please contact the server administrator.", vbCritical
                    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[7]")
                End If
            Else
                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[8]")
                '*KT 52-11-20 ทำการไล่ Commit Tran ของ Local
                Call SP_SQLxCommitTranLocal
                Call SP_SQLxCommitTranServer
            End If
            '-----------------------------------------------------------------
        Else
            Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[9]")
            '*KT 52-11-20 ทำการไล่ Commit Tran ของ Local
            Call SP_SQLxCommitTranLocal
        End If
        
        Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[10]")
        '*Eaw 56-09-30 update FTTmnStatus ของ virtualPOs ให้เป็น 'N' เมื่อปิด Ternant sale ด้วย
        Select Case tVB_CNMnuCurrent
        Case "172105"
            Call SP_TMNxUpdateVirtualPosClose
        Case Else
            'none
        End Select
        If Not SP_TMNxUpdateClose Then
            Call SP_MSGnShowing("tms_CN005212", tCS_CNMsgErr)
        End If
        bVB_CNOpenTmnOff = False
        Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[11]")
    
        'ต้องทำการ Sumary กรณี Missing / Duplicate Transaction
        'Call W_PRCxMissing 'ให้ไปใช้งานที่ ISS3000 '*KT 51-10-02 3.0.127
        'ทำการเก็บค่าสะสมเพื่อเตรียมออก Closing Report
        Call SP_PRCxCloseTmn
        Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[12]")
'        Call SP_PRCxBckClsing
'        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "CloseTerminal" & " Form:" & Me.Name & " Step:SP_PRCxBckClsing" & " Events:W_PRCxTrnMnuClose")
         bVB_CNPastOffline = False
         
        tVB_Transaction = "97"
        tVB_TransactionSub = "97"
        bVB_FormPayAct = True
        Call SP_STAxMouse(Me, False)
        If Not wCNBegin Is Nothing Then
            Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:W_PRCxTrnMnuClose[13]")
            Unload Me
        End If
    'Else
    '    bUT_SignOut = False
    'End If
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Step:ErrHandle" & " Events:W_PRCxTrnMnuClose[14]")
End Sub

