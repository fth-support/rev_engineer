VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNCloseMnu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu;klg_Title"
   ClientHeight    =   4860
   ClientLeft      =   6990
   ClientTop       =   4035
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "wCNCloseMnu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4680
   Tag             =   "2;"
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
      Left            =   3120
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   4200
      Width           =   1335
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1080
      Index           =   2
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4080
      Width           =   4680
      _Version        =   720897
      _ExtentX        =   8255
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin VB.Timer otmOnOff 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   720
         Top             =   360
      End
      Begin VB.Timer otmSendEJ 
         Interval        =   500
         Left            =   120
         Top             =   360
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid ogdFunc 
      Height          =   3540
      Left            =   0
      TabIndex        =   0
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
      Rows            =   16
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"wCNCloseMnu.frx":058A
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
         Caption         =   "Main Menu;klg_Menu"
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
Attribute VB_Name = "wCNCloseMnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_HotKey As String
Private bW_KeyEnter As Boolean
Private nMaxRow As Long

Private Sub W_CHKxScreenActiveForm()
On Error GoTo ErrHandle
    If Screen.ActiveForm.Name <> "wCNReEntryMnu" Then
        wCNReEntryMnu.Show vbModal
    End If
Exit Sub
ErrHandle:
    wCNReEntryMnu.Show vbModal
End Sub

Private Sub W_PRCxSendEJold()
'-----------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 2014/03/12
'-----------------------------------------------------------
Dim tLocal As String
Dim oFSO As New FileSystemObject
Dim oSubFolders  As Folder
Dim oFolders As Folder
Dim oFile As File
On Error GoTo ErrHandle

    tLocal = oFSO.GetParentFolderName(App.Path) & "\POSEJ\"
    If Not SP_FLDbManFolder(tLocal, nEN_Fld1Exist) Then
        MsgBox "Not found Path EJ Local !", vbOKOnly, tVB_CNAppTitle
        Exit Sub
    End If

    Set oSubFolders = oFSO.GetFolder(tLocal)
    For Each oFolders In oSubFolders.SubFolders
        If UCase("Backup") <> UCase(oFolders.Name) Then
            Call SP_MOVxEJToOnlineRE(oFolders.Name)
        End If
    Next
    
    Set oFSO = Nothing
    
Exit Sub
ErrHandle:

End Sub

Private Sub W_PRCxCreateEJ(ByVal ptTransType As String, ByVal pdTransDate As Date, ByVal ptTransTmn As String, ByVal ptTransNo As String)

On Error GoTo ErrHandle
    oTransPrn.TypePriner = "1"
    oTransPrn.nW_StateCon = 1
    oTransPrn.nW_StatePrn = 1
    oTransPrn.nW_StateDoc = 1
'    oTransPrn.bPrnByOPOS = False
    oTransPrn.PrintEJ = True                        'bW_EnableEJ
    oTransPrn.bPRMGenEJ = True
    oTransPrn.RePrint = True                        'bW_EnableRePrint
    oTransPrn.RePrintNormal = True          'bW_EnableRePrintNormal
    Call oTransPrn.OpenTransaction(Val(ptTransType), pdTransDate, ptTransTmn, ptTransNo, False)
    Call oTransPrn.PrintOut
    oTransPrn.bPRMGenEJ = False
    oTransPrn.RePrint = False
    oTransPrn.RePrintNormal = False
Exit Sub
ErrHandle:

End Sub

Private Sub W_PRCxCheckEJ(ByVal pnMode As Integer)
'--------------------------------------------------------
'Call:
'Cmt: '*Ao 2014/01/27
'-------------------------------------------------------
Dim tSql  As String
Dim orsTemp As ADODB.Recordset
Dim oFSO As New FileSystemObject
Dim tDepID As String, tTmnNum As String, tTransNo As String, tStoreID As String, tTransType As String, tEJ As String
'Dim bChk As Boolean
'Dim aEJBackup() As String
'Dim aEJLocal() As String
'Dim oFldInbox  As Folder
'Dim oFile As File
Dim dTransDate As Date
Dim tLocal As String, tBackup As String
Dim tTemp As String
On Error GoTo ErrHandle

    
    tLocal = oFSO.GetParentFolderName(App.Path) & "\POSEJ\" & Format(CDate(tUT_SaleDate), "yyMMdd")
    tBackup = oFSO.GetParentFolderName(App.Path) & "\POSEJ\Backup\" & Format(CDate(tUT_SaleDate), "yyMMdd")
    
    If Not SP_FLDbManFolder(tLocal, nEN_Fld1Exist) Then
       Call SP_FLDbManFolder(tLocal, nEN_Fld3Create)
    End If
    If Not SP_FLDbManFolder(tBackup, nEN_Fld1Exist) Then
       Call SP_FLDbManFolder(tBackup, nEN_Fld3Create)
    End If
    
    If pnMode = 0 Then
        tSql = "SELECT FTShdDepID,FTTmnNum,FTShdTransNo,FTStmCode,FTShdTransType,FDShdTransDate FROM TPSTSalHD"
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "ORDER BY FTShdTransNo"
        bVB_ReEntryMode = False
    Else
        tSql = "SELECT FTShdDepID,FTTmnNum,FTShdTransNo,FTStmCode,FTShdTransType,FDShdTransDate FROM TPSTSalHD"
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'        tSql = tSql & vbCrLf & "AND FDShdTransDate <> '" & tUT_SaleDate & "'"          'อาจขายย้อนหลัง จึงไม่ต้อง WHERE
        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
        tSql = tSql & vbCrLf & "AND FTShdTransType IN ('43','44','45','46','47','49')"              'Tenant Sales
        tSql = tSql & vbCrLf & "ORDER BY FTShdTransNo"
        bVB_ReEntryMode = True
        tVB_EJDate = ""
    End If
    Call SP_SQLvExecute(tSql, orsTemp)
    If Not (orsTemp.BOF And orsTemp.EOF) Then

        Do While Not orsTemp.EOF
            tDepID = SP_FEDtChkString(orsTemp, "FTShdDepID")
            tTmnNum = SP_FEDtChkString(orsTemp, "FTTmnNum")
            tTransNo = SP_FEDtChkString(orsTemp, "FTShdTransNo")
            tStoreID = SP_FEDtChkString(orsTemp, "FTStmCode")
            tTransType = SP_FEDtChkString(orsTemp, "FTShdTransType")
            dTransDate = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
            tEJ = Format(tDepID, "00") & Format(tTmnNum, "00000") & Format(tTransNo, "00000") & tStoreID & ".EJ"
            
            tTemp = tLocal & "\" & tEJ
            
            If Not SP_FLEbManFile(tTemp, nEN_Fle1Exist) Then
                tTemp = tBackup & "\" & tEJ
                If Not SP_FLEbManFile(tTemp, nEN_Fle1Exist) Then
                    Call W_PRCxCreateEJ(tTransType, dTransDate, tTmnNum, tTransNo)
                End If
            End If

            orsTemp.MoveNext
        Loop
    End If
    
Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

'Private Sub W_PRCxCheckEJ(ByVal pnMode As Integer)
''--------------------------------------------------------
''Call:
''Cmt: '*Ao 2014/01/27
''-------------------------------------------------------
'Dim nLoop As Integer
'Dim tSql  As String, tTmp As String
'Dim orsTemp As ADODB.Recordset
'Dim oFSO As New FileSystemObject
'Dim tDepID As String, tTmnNum As String, tTransNo As String, tStoreID As String, tTransType As String, tEJ As String
'Dim bChk As Boolean
'Dim aEJBackup() As String
'Dim aEJLocal() As String
'Dim oFldInbox  As Folder
'Dim oFile As File
'Dim dTransDate As Date
'
'On Error GoTo ErrHandle
'
'
'    ReDim aEJBackup(0)
'    ReDim aEJLocal(0)
'
'    If pnMode = 0 Then
'        tSql = "SELECT FTShdDepID,FTTmnNum,FTShdTransNo,FTStmCode,FTShdTransType,FDShdTransDate FROM TPSTSalHD"
'        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
'        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
'        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
'        tSql = tSql & vbCrLf & "ORDER BY FTShdTransNo"
'        bVB_ReEntryMode = False
'    Else
'        tSql = "SELECT FTShdDepID,FTTmnNum,FTShdTransNo,FTStmCode,FTShdTransType,FDShdTransDate FROM TPSTSalHD"
'        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & tVB_CNTerminalNum & "'"
''        tSql = tSql & vbCrLf & "AND FDShdTransDate <> '" & tUT_SaleDate & "'"          'อาจขายย้อนหลัง จึงไม่ต้อง WHERE
'        tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
'        tSql = tSql & vbCrLf & "AND FTShdTransType IN ('43','44','45','46','47','49')"              'Tenant Sales
'        tSql = tSql & vbCrLf & "ORDER BY FTShdTransNo"
'        bVB_ReEntryMode = True
'        tVB_EJDate = ""
'    End If
'    Call SP_SQLvExecute(tSql, orsTemp)
'    If Not (orsTemp.BOF And orsTemp.EOF) Then
'
'        tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\Backup\"
'        If SP_FLDbManFolder(tTmp, nEN_Fld1Exist) Then
'            tTmp = tTmp & Format(CDate(tUT_SaleDate), "yyMMdd")
'            Set oFldInbox = oFSO.GetFolder(tTmp)
'            For Each oFile In oFldInbox.Files
'                aEJBackup(UBound(aEJBackup)) = oFile.Name
'                ReDim Preserve aEJBackup(UBound(aEJBackup) + 1)
'            Next
'            tTmp = oFSO.GetParentFolderName(App.Path) & "\POSEJ\"
'            If SP_FLDbManFolder(tTmp, nEN_Fld1Exist) Then
'                tTmp = tTmp & Format(CDate(tUT_SaleDate), "yyMMdd")
'                Set oFldInbox = oFSO.GetFolder(tTmp)
'                For Each oFile In oFldInbox.Files
'                    aEJLocal(UBound(aEJLocal)) = oFile.Name
'                    ReDim Preserve aEJLocal(UBound(aEJLocal) + 1)
'                Next
'            Else
'                MsgBox "Not found Path EJ Local", vbOKOnly, tVB_CNAppTitle
'            End If
'            Do While Not orsTemp.EOF
'                tDepID = SP_FEDtChkString(orsTemp, "FTShdDepID")
'                tTmnNum = SP_FEDtChkString(orsTemp, "FTTmnNum")
'                tTransNo = SP_FEDtChkString(orsTemp, "FTShdTransNo")
'                tStoreID = SP_FEDtChkString(orsTemp, "FTStmCode")
'                tTransType = SP_FEDtChkString(orsTemp, "FTShdTransType")
'                dTransDate = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
'                tEJ = Format(tDepID, "00") & Format(tTmnNum, "00000") & Format(tTransNo, "00000") & tStoreID & ".EJ"
'
'                bChk = True         '*Ao 2014/01/30
'
'                For nLoop = LBound(aEJBackup) To UBound(aEJBackup) - 1          'ตรวจสอบ \POSEJ\Backup\
'                    If UCase(aEJBackup(nLoop)) = UCase(tEJ) Then
'                        bChk = False
'                        Exit For
'                    Else
'                        bChk = True
'                    End If
'                Next nLoop
'                If bChk Then
'                    For nLoop = LBound(aEJLocal) To UBound(aEJLocal) - 1            'ตรวจสอบ \POSEJ
'                        If UCase(aEJLocal(nLoop)) = UCase(tEJ) Then
'                            bChk = False
'                            Exit For
'                        Else
'                            bChk = True
'                        End If
'                    Next nLoop
'                    If bChk Then
'                        Call W_PRCxCreateEJ(tTransType, dTransDate, tTmnNum, tTransNo)
'                    End If
'                End If
'                orsTemp.MoveNext
'            Loop
'        Else
'            MsgBox "Not found Path EJ Backup", vbOKOnly, tVB_CNAppTitle
'        End If
'    End If
'
'Exit Sub
'ErrHandle:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'End Sub

Private Sub W_PRCxUpdateOprDate()
'--------------------------------------------------------
'Call:
'Cmt: '*Ao 2013/11/08
'-------------------------------------------------------
Dim tSql As String
On Error GoTo ErrHandle

    tSql = "UPDATE TSysChgDateTime"
    tSql = tSql & vbCrLf & "SET FDCdtDate = '" & tUT_SaleDate & "'"
    Call SP_SQLvExecute(tSql)
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_FNCxAddFunc2Grid()
'--------------------------------------------------------
'Call:
'Cmt: *KT 50-12-24 ทำการดึงข้อมูลใน Array มาแสดงที่หน้าจอว่ามีรายการ Main Menu ไหนบ้าง
'-------------------------------------------------------
    Dim nI As Long
    Dim nR As Long
    With ogdFunc
        nR = LBound(aVB_MainMnu)
        .rows = UBound(aVB_MainMnu)
        For nI = LBound(aVB_MainMnu) To UBound(aVB_MainMnu) - 1
'            If aVB_MainMnu(nI).Show Then
'                .TextMatrix(nR, 0) = aVB_MainMnu(nI).ID
'                .TextMatrix(nR, 1) = " " & aVB_MainMnu(nI).ID & "." & IIf(aVB_MainMnu(nI).NameUsr <> "", aVB_MainMnu(nI).NameUsr, aVB_MainMnu(nI).Name)
'                 nR = nR + 1
'            End If
            .TextMatrix(nR, 0) = aVB_MainMnu(nI).ID         '*Ao 2013/10/25     ***ใน Back  FNSmnVisible = 0
            .TextMatrix(nR, 1) = " " & aVB_MainMnu(nI).ID & "." & IIf(aVB_MainMnu(nI).NameUsr <> "", aVB_MainMnu(nI).NameUsr, aVB_MainMnu(nI).Name)
             nR = nR + 1
        Next nI
        .rows = nR
'        .RowHidden(2) = True 'Print Closing Report
        '.RowHidden(7) = True        'Print Closing Report '*Tao 60-01-19 ของเดิม
        If nR >= 8 Then .RowHidden(7) = True 'Print Closing Report '*Tao 60-01-19 ป้องกัน grid error
        '*Eaw 55-12-28 เปลี่ยนตาการแสดงใน TSysPosMenu
        ' 1.Open
        ' 2.Open With Oprdate
        ' 3.Reprint Closing Report
        ' 4.Admin Menu
        ' 5.RE -ENTRY
        ' 6.Exit Terminal
        ' 7.Print Closing; Repor
    End With
End Sub

Private Sub Form_Activate()
    bVB_ReEntryMode = False
    bW_KeyEnter = False
End Sub

'Private Sub Form_Activate()
'On Error GoTo ErrHandle
'    Call SP_OBJxSetFocus(ogdFunc)
'    Exit Sub
'ErrHandle:
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
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
On Error GoTo ErrHandle
 
    otmSendEJ.Enabled = False   '*Bump 57-03-24 17.00 pm. V.4.0.123 CommSheet TK-ISS3000-473 00 00 05.xlsx
 
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[1]")
    bW_KeyEnter = False
    tW_HotKey = ""
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FNCxAddFunc2Grid
    nMaxRow = ogdFunc.Aggregate(flexSTMax, 0, 0, ogdFunc.rows - 1, 0)
   ' *KT 51-10-13 V 3.0.133/V 3.0.200
    '*KT 51-10-08 V 3.0.131/V 3.0.200
    '-----------------------------------------------------
'    If Not bVB_CNCreateTblTmp And Not bVB_CNCreateTblTmpFrmCloseRpt Then
'        Call UT_CreateSaleTableTMP
'        bVB_CNCreateTblTmp = True
'    End If
    '-----------------------------------------------------
'    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล    ไม่ต้องทำแล้ว
    Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:Form_Load[2]")
'    Me.ZOrder 0
    Call W_PRCxSendEJold                    '*Ao 2014/03/12    CommSheet TK-ISS3000-473_Aoo1303
    otmSendEJ.Enabled = True   '*Bump 57-03-24 17.00 pm. V.4.0.123 CommSheet TK-ISS3000-473 00 00 05.xlsx
    '*KT 58-05-27 TK-ISS3000-587 ป้องกันการไม่พิมพ์ใบ Closing Report
    '**********************************************
    If bVB_CNCloseTmnSta And (Not bVB_CNClosingRptSta) Then
         'ogdFunc.Select 7, 0 '*TON 58-12-01 TK-ISS3000-581
         If nMaxRow >= 8 Then ogdFunc.Select 7, 0 '*Tao 60-01-19 ป้องกัน grid error
    End If
   '**********************************************
    '*Tao 60-01-20 เช็คก่อนว่าพิมพ์ Closing Report หรือยัง
    If (Not bUT_SignOut) And bVB_FirstLoad Then
    Select Case SP_ChkxClosingMissing
    Case "08"
        bVB_CNCloseTmnSta = True
        bVB_CNClosingRptAgain = True
        bW_KeyEnter = False
        Call ocmOK_Click
    End Select
    End If
    '========================================
    Exit Sub
ErrHandle:
    Call SP_CRTxWriteLogTerminated("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Events:Form_Load[3]")
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = True
'End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:
    
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Set wCNCloseMnu = Nothing
'End Sub

Public Sub ocmOK_Click()
Dim tTranType$
Dim dTmp As Date
Dim tWhr As String
Dim tMsgData As String, tMsgEJ As String, tMsg As String, tMsgALL As String
Dim oResult As VbMsgBoxResult
Dim tFile As String
Dim oFSO As New FileSystemObject

On Error GoTo ErrHandle
'    Dim dEndDatTimeToday As Date
'    Dim dCurrentDateTime As Date
    '*KT 52-03-24 PH1.5-ML-SF-088 *KT 51-12-04 V 1.0.2
    '------------------------------------------------------------------------------------
    If bW_KeyEnter Then
        Exit Sub
    Else
        bW_KeyEnter = True
    End If
    '-------------------------------------------------------------------------------------
'*Eaw 55-12-28 เปลี่ยนตาการแสดงใน TSysPosMenu
' 1.Open
' 2.Open With Oprdate
' 3.Reprint Closing Report
' 4.Admin Menu
' 5.RE -ENTRY
' 6.Exit Terminal
' 7.Print Closing; Report
    
    With ogdFunc
             Call SP_STAxMouse(Me, True)
             
             '*KT 58-05-27 TK-ISS3000-587 ป้องกันการไม่พิมพ์ใบ Closing Report
             'ให้ทำการตรวจสอบเพื่อทำการ Print Closing Report ก่อน
             '***********************************************************
             ' bVB_CNCloseTmnSta : สถานะทำ Close Terminal
             ' bVB_CNClosingRptSta : สถานะทำ Closing Report
             ' bVB_CNClosingRptSta : สถานะ bVB_CNClosingRptAgain แจ้งให้ทำงานเข้าสู่การตรวจสอบว่าต้องทำ Closing Report หรือ Reprint Closing Report
             If bVB_CNCloseTmnSta And (Not bVB_CNClosingRptSta) And bVB_CNClosingRptAgain Then
                'ทำเอกสาร Close Terminal แล้ว แต่ไม่สามารถทำ Closing Report
                Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:ocmOK_Click[Print Closing Report again]")
                '*Tao 61-01-19 ในกรณีที่ปิดรอบไม่สมบูรณ์ ไม่ต้อง สร้าง OverShortReport ใหม่
                If bVB_FirstLoad Then
                    oResult = SP_MSGnShowing("tms_CN005257", vbOKOnly)
                Else
                    oResult = SP_MSGnShowing("tms_CN005257", nCS_OkCancelDefaultButton1)
                End If
                If oResult = vbOK Then
                'If SP_MSGnShowing("tms_CN005257", tCS_CNMsgYesNoDefYes) = vbYes Then '*Tao 61-01-19 ของเดิม
                    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Re Processing Closing Report" & " Form:" & Me.Name & " Events:ocmOK_Click[FOUND=OK]")
                    tTranType = SP_CHKtClosingReport '*TON 58-09-18 ให้ถาม yes no
                    bUT_SignOut = True '*Tao 60-01-23 จะได้พิมพ์ Closing Report ได้
                Else
                    Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Re Processing Closing Report" & " Form:" & Me.Name & " Events:ocmOK_Click[FOUND=CANCEL]")
                    tTranType = .TextMatrix(.Row, 0)
                    bVB_CNClosingRptAgain = False
                End If
            Else
                 tTranType = .TextMatrix(.Row, 0)
                 bVB_CNClosingRptAgain = False
             End If
             '***********************************************************
            Select Case tTranType
                Case "1" 'Open
                        Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:ocmOK_Click[1]")
                        tVB_Transaction = aVB_TransMode(0).ID
                        tVB_TransactionSub = tVB_Transaction
                        tVB_TransName = aVB_TransMode(0).Name
                        tVB_TransNameABB = aVB_NameAbb(tEN_AbbOpen - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbOpen - 1).UsrLangSec
                        tVB_CNTypePrinter = aVB_TransMode(0).PrintType
                        tVB_CNMnuCurrent = aVB_MainMnu(Val(tTranType) - 1).SmnID
                        nVB_CurentMainMnu = Val(tTranType) - 1
                        bVB_ReEntryMode = False                     '*Ao 2014/09/10
                        If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
                            oTerminal.Connect oDB.OnServerName, oDB.OnUsrName, oDB.OnUsrPwd, oDB.OnDbName
                        Else
                            oTerminal.Online = False
                        End If
                        If Not oTerminal.StoreOpen Then
                                If SP_MSGnShowing("tms_CN005098", tCS_CNMsgYesNoDefYes, , SP_INItGetLang("ALL", "tms_CN005100")) = vbNo Then
                                        Call SP_STAxMouse(Me, False)
                                        bW_KeyEnter = False
                                        Exit Sub
                                Else
                                        Select Case oTerminal.AvailableSaleMode
                                        Case nEN_OfflineNormal
                                                wCNChgOptDate.Show vbModal
                                                If Not bVB_CNChgDptDate Then
                                                        Call SP_STAxMouse(Me, False)
                                                        bW_KeyEnter = False
                                                        Exit Sub
                                                End If
                                                oTerminal.SaleMode = nEN_OfflineNormal
                                                oTerminal.SaleDate = CDate(tUT_SaleDate)
                                                tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate)
'                                                Call SP_CRTxLogOpenTmn("Open Terminal mode offline : " & tUT_SaleDate)
                                                bVB_CNPastOffline = True
                                                If bVB_CNPrnOPOS Then 'Print By OPOS
                                                    oTransPrn.nW_StateCon = 1
                                                Else
                                                    Set oTransPrn.ActivePrinter = oPrn.SelectPrinter(tVB_PrnDevice)
                                                    If Not oPrn.bState Then
                                                            Call SP_MSGnShowing("tms_CN005156", tCS_CNMsgErr)
                                                            Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Printer Driver Error[Offline]" & " Form:" & Me.Name & " Control:ocmOk" & " Events:Click")
                                                            Call SP_APPxEndProgram
                                                    End If
                                                End If
                                                oTransPrn.ExportFolder = tVB_CNPthBnkTr
                                                oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
                                        Case nEN_OfflineOpenWith
                                                'Call SP_PRCxSwithToOffline 'PH 2.0.0
                                                oTerminal.SaleMode = nEN_OfflineOpenWith
                                                oTerminal.SaleDate = SP_GETdCurrentDate(True)
                                                tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) 'adjust salesdate with opration date config
'                                                Call SP_CRTxLogOpenTmn("Open With Terminal mode offline : " & tUT_SaleDate)
                                                bVB_CNOpenTmnOff = True  'in db  tmn set to close
                                                'Set oTransPrn.ActiveConnection = oDB.ocnOffLine
                                                If bVB_CNPrnOPOS Then 'Print By OPOS
                                                    oTransPrn.nW_StateCon = 1
                                                Else
                                                    Set oTransPrn.ActivePrinter = oPrn.SelectPrinter(tVB_PrnDevice)
                                                    If Not oPrn.bState Then
                                                            Call SP_MSGnShowing("tms_CN005156", tCS_CNMsgErr)
                                                            Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Printer Driver Error[Open with]" & " Form:" & Me.Name & " Control:ocmOk" & " Events:Click")
                                                            Call SP_APPxEndProgram
                                                    End If
                                                End If
                                                oTransPrn.ExportFolder = tVB_CNPthBnkTr
                                                oTransPrn.EJFolder = tVB_CNOnDbSrcEJ
                                        End Select
                                End If
                        Else
                                If bVB_CNPastOffline Then
                                        Call SP_STAxMouse(Me, False)
                                        wCNChgOptDate.Show vbModal
                                        If Not bVB_CNChgDptDate Then
                                                Call SP_STAxMouse(Me, False)
                                                bW_KeyEnter = False
                                                Exit Sub
                                        End If
'                                        Call SP_CRTxLogOpenTmn("Open  Terminal mode online(Change Opt Date) : " & tUT_SaleDate)
                                Else
                                        '*KT 52-09-26
                                        oTerminal.SaleMode = nEN_OnlineNormal
                                        oTerminal.SaleDate = SP_GETdSaleDateEx
                                        tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) 'adjust salesdate with opration date config
'                                        Call SP_CRTxLogOpenTmn("Open  Terminal mode online : " & tUT_SaleDate)
                                End If
                        End If
                        bW_KeyEnter = False
                        tVB_EJDate = tUT_SaleDate               '*Ao 2014/02/03   เก็บ OP
                        Call W_PRCxUpdateOprDate   '*Ao 2013/11/08
                        Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & " Events:ocmOK_Click[2]")
                Case "2" 'Open With Oprdate
                        tVB_Transaction = aVB_TransMode(0).ID
                        tVB_TransactionSub = tVB_Transaction
                        tVB_TransName = aVB_TransMode(0).Name
                        tVB_TransNameABB = aVB_NameAbb(tEN_AbbOpen - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbOpen - 1).UsrLangSec
                        tVB_CNTypePrinter = aVB_TransMode(0).PrintType
                        tVB_CNMnuCurrent = aVB_MainMnu(Val(tTranType) - 1).SmnID
                        nVB_CurentMainMnu = Val(tTranType) - 1
                        oTerminal.SaleMode = nEN_OfflineOpenWith '*KT 51-04-29
                        dTmp = SP_GETdCurrentDate(True)
                        If SP_MSGnShowing("tms_CN005097", tCS_CNMsgOkCancel, , SP_INItGetLang("ALL", "tms_CN005099"), Format(dTmp, "DD/MM/YYYY")) = vbCancel Then
                                Call SP_STAxMouse(Me, False)
                                bW_KeyEnter = False
                                Exit Sub
                        End If
                        If SP_MSGnShowing("tms_CN005098", tCS_CNMsgYesNoDefYes, , SP_INItGetLang("ALL", "tms_CN005100")) = vbNo Then
                                Call SP_STAxMouse(Me, False)
                                bW_KeyEnter = False
                                Exit Sub
                        End If
                        'Call SP_PRCxSwithToOffline 'PH 2.0.0
                        oTerminal.SaleMode = nEN_OfflineOpenWith
                        oTerminal.SaleDate = dTmp
                        tUT_SaleDate = UT_tStringDate(oTerminal.SaleDate) 'adjust salesdate with opration date config
                        bVB_CNOpenTmnOff = True  'in db  tmn set to close
                        bW_KeyEnter = False
                        tVB_EJDate = tUT_SaleDate               '*Ao 2014/02/03   เก็บ OP
                        Call W_PRCxUpdateOprDate   '*Ao 2013/11/08
                Case "8" 'Print Clossing Report'*TON 58-12-01
                '*Eaw 55-12-28 เปลี่ยนตาการแสดงใน TSysPosMenu
'                Case "3" 'Print Clossing Report
                    '*KT 51-10-08 V 3.0.131/V 3.0.200
                    '-----------------------------------------------------
                   
                    If (bVB_CNCloseTmnSta And bUT_SignOut) Or bVB_CNClosingRptAgain Then  '*Tao 60-01-19 ถ้าพิมพ์ Closing แล้ว ไม่ให้พิมพ์อีก ให้ไปพิมพ์ RePrint Clossing
                        If Not bVB_CNCreateTblTmp Then
                            Call UT_CreateSaleTableTMP
                            bVB_CNCreateTblTmp = True
                            bVB_CNCreateTblTmpFrmCloseRpt = True '*KT 51-10-10 V 3.0.132/V 3.0.200
                            If tUT_LoginCode = "" Then '*Tao 60-01-20 ถ้ายังไม่ได้ tUT_LoginCode ค่อยหา
                                '*KT 52-01-14 ต้องหาค่า tUT_LoginCode กรรีที่เปิดโปรแกรมเข้ามาที่ Main Menu แล้วแล้วทำการกด Closing Report โดยยังไม่เข้าขาย
                                tWhr = "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
                                'tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'" '*Tao 61-01-18 ไม่สนใจวันที่ SaleDate
                                tWhr = tWhr & "  AND FTShdStaDoc = '1'"
                                tUT_LoginCode = SP_SQLtFunction("MAX", "FTLogCode", "TPSTSalHD", tWhr, oDB.ocnOnLine, True)
                                If tUT_LoginCode = "" Then tUT_LoginCode = SP_String2YYMM(Date) & Format(tVB_CNTerminalNum, "00000") & "00001"
                            End If
                        Else
                            bVB_CNCreateTblTmpFrmCloseRpt = False '*KT 51-10-13 V 3.0.133/V 3.0.200
                        End If
                        '-----------------------------------------------------
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
                                tVB_Transaction = aVB_TransMode(8).ID
                                tVB_TransactionSub = tVB_Transaction
                                tVB_TransName = aVB_TransMode(8).Name
                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbPrintClose - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbPrintClose - 1).UsrLangSec
                                tVB_CNTypePrinter = aVB_TransMode(8).PrintType
                                tVB_CNMnuCurrent = aVB_MainMnu(Val(tTranType) - 1).SmnID
                        End Select
    
                       ' wCNAllow.tW_ID = tVB_CNMnuCurrent
                       ' wCNAllow.Show vbModal
    '                    DoEvents '*KT 53-01-14 DoEvents
                       ' If tVB_CNUserAlwC <> "" Then
                            '*KT 51-07-14 Comsheet ML- SF- 398
                            'RunTime Error When Print Closing Report
                            '*KT 51-07-29 การ Closing Report ให้นับเป็นเครื่องไม่กรอง Cashier
                            'รวมยอดขายทั้งเครื่อง ถ้ามีการ Closing Report แล้วจะเคลียร์ ยอดสะสมแล้วเริ่มต้นใหม่
                            '====================================================
                            If Not SP_TBLbChkExist("TPSTClose") Then
                                bVB_CNSplash = True  '*KT 54-04-08 V.3.0.56 ป้องกันเครื่องค้างและโปรแกรมหลุด
                                'Call SP_STAxSplash '*KT 54-04-08 V.3.0.56 ป้องกันเครื่องค้างและโปรแกรมหลุด
                                Call SP_CHKxTblTmpClosing("TPSTClose")
                                bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
                            End If
                            bVB_CNSplash = True  '*KT 54-04-08 V.3.0.56 ป้องกันเครื่องค้างและโปรแกรมหลุด
                            'Call SP_STAxSplash '*KT 54-04-08 V.3.0.56 ป้องกันเครื่องค้างและโปรแกรมหลุด
                            
                            '*Tao 60-01-20 หากยังไม่ print signout ให้ทำใหม่
                            If SP_VALnText2Double(SP_SQLtFunction("COUNT", "FTCloGrp", "TPSTClose", "", oDB.ocnOffLine)) = 0 Then
                                Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Re Processing Sign Out" & " Form:" & Me.Name & " Events:ocmOK_Click[Start]")
                                Call wFunctionKB.W_PrintSignOut(False, False, True)
                                Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Re Processing Sign Out" & " Form:" & Me.Name & " Events:ocmOK_Click[END]")
                            End If
                            '====================================
                            
                            Select Case tVB_CNMnuCurrent
                                Case "172005"
                                    Call wFunctionKB.W_PrintClosing(False)
                                Case "172105"
                                    Call wFunctionKB.W_PrintClosing(False)
                                Case Else
                                    'Call wFunctionKB.W_PrintClosing(True)  ของเดิม Original
                                    
                                    '*Bmp 60-12-20 เพิ่มการตรวจสอบว่า หากมี Closing แล้ว ไม่ต้องทำ SP_PRCbChkClosing
                                    '* ป้องกันปัญหา CommSheet TTK-ISS3000-665 TK Issue - Close Report generated twice.xls
                                    Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Print Closing Report(1)"))
                                    If Not SP_PRCbChkClosing(tVB_Transaction, tUT_LoginCode) Then
                                        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Re Processing Print Closing" & " Form:" & Me.Name & " Events:ocmOK_Click[Start]")
                                        Call wFunctionKB.W_PrintClosing(True)
                                        Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Re Processing Print Closing" & " Form:" & Me.Name & " Events:ocmOK_Click[END]")
                                        Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " Print Closing Report(2)"))
                                    End If
                                    '============================================================================================
                            End Select
                            
                            bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
                            '=====================================================
                            '*KT 58-05-27 TK-ISS3000-587 ป้องกันการไม่พิมพ์ใบ Closing Report
                            '******************
                            If Not bVB_CNClosingRptSta Then
                                 bW_KeyEnter = False
                                 Exit Sub
                            End If
                             '******************
                            '*KT 51-08-08 Comsheet ML-SF-451 ให้ Set = Y เมื่อทำ Print Closing Report
                            Call SP_TMNxUpdatePrnClose
                                                        
                        'End If
                        Select Case tVB_CNMnuCurrent
                            Case "172005"
                                wCNReEntryMnu.otmForm.Enabled = True
                                If Screen.ActiveForm.Name <> "wCNReEntryMnu" Then
                                    wCNReEntryMnu.Show vbModal
                                End If
                            Case "172105"
                                wCNReEntryMnu.otmForm.Enabled = True
                                If Screen.ActiveForm.Name <> "wCNReEntryMnu" Then
                                    wCNReEntryMnu.Show vbModal
                                End If
    '                            If ActiveForm.Name = "wCNReEntryMnu" Then
    '                            End If
                            Case Else
                                tVB_Transaction = "99"
                                tVB_TransactionSub = "99"
                                bW_KeyEnter = False
                        End Select
                    Else
                        bW_KeyEnter = False
                    End If
                Case "3" 'RePrint Clossing Report
                '*Eaw 55-12-28 เปลี่ยนตาการแสดงใน TSysPosMenu
'                Case "4" 'RePrint Clossing Report
                    tVB_Transaction = "99"
                    tVB_TransactionSub = "99"
                    tVB_CNMnuCurrent = aVB_MainMnu(Val(tTranType) - 1).SmnID
                    tVB_TransNameABB = aVB_NameAbb(tEN_AbbRePrintClose - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbRePrintClose - 1).UsrLangSec
                    wCNAllow.tW_ID = tVB_CNMnuCurrent
                    wCNAllow.Show vbModal
'                    DoEvents '*KT 53-01-14 DoEvents
                    If tVB_CNUserAlwC <> "" Then
                        bVB_CNSplash = True  '*KT 54-04-08 V.3.0.56 ป้องกันเครื่องค้างและโปรแกรมหลุด
                        'Call SP_STAxSplash '*KT 54-04-08 V.3.0.56 ป้องกันเครื่องค้างและโปรแกรมหลุด
                        Call wFunctionKB.W_RePrintSignOut(True)
                         bVB_CNSplash = False '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
                    End If
                    tVB_Transaction = "99"
                    tVB_TransactionSub = "99"
                    bW_KeyEnter = False
                Case "4" 'Admin Function
                '*Eaw 55-12-28 เปลี่ยนตาการแสดงใน TSysPosMenu
'                Case "5" 'Admin Function
                    tVB_CNMnuCurrent = aVB_MainMnu(Val(tTranType) - 1).SmnID
                    nVB_CurentMainMnu = Val(tTranType) - 1
                    wAdmin.Show vbModal
                    tVB_Transaction = 99
                    tVB_TransactionSub = 99
                    bW_KeyEnter = False
                Case "5" 're entry
                '*Eaw 55-12-28 เปลี่ยนตาการแสดงใน TSysPosMenu
'                Case "6" 're entry
                    Call W_PRCxCheckEJ(0)
                    Call W_PRCxCheckEJ(1)
                    If W_GETbStaClosingReport Then
                        tVB_CNMnuCurrent = "1606"
                        nVB_CurentMainMnu = Val(tTranType) - 1
                        'wCNReEntryMnu.Show vbModal
                        tVB_Transaction = 98
                        tVB_TransactionSub = 98
                        bW_KeyEnter = False
                        '        If Not SP_USRbMnuAccept(nEN_Usr2AthRead, tVB_CNMnuCurrent) Then
                        '            Exit Sub
                        '        End If
                    Else
                        Call SP_MSGnShowing("tms_CN005237", tCS_CNMsgWarn)
                        bW_KeyEnter = False
                        Exit Sub
                    End If
                    Call W_PRCxUpdateOprDate   '*Ao 2013/11/08
                Case "7" 'Exit Terminal'*TON 58-12-01
                '*Eaw 55-12-28 เปลี่ยนตาการแสดงใน TSysPosMenu
'                Case "7" 'Exit Terminal
                    tVB_CNMnuCurrent = "1605" 'aVB_MainMnu(Val(tTranType) - 1).SmnID
                    nVB_CurentMainMnu = Val(tTranType) - 1
                    tVB_Transaction = "100" '"08" '*KT 52-12-17 POSFront 2.0.30
                    tVB_TransactionSub = "100" '"08" '*KT 52-12-17 POSFront 2.0.30
                    bW_KeyEnter = False
                    Call W_PRCxCheckEJ(0)
                    Call W_PRCxCheckEJ(1)
                    Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล
                    tMsgData = SP_MSGtNoSendData
'                    '*Tao 61-11-20 TK-ISS3000-712 TK ISSUE - TK Request - Modify exit terminal for resend transaction (เตือนแล้วว่า ถ้า Transfer ไม่ครบทุกบิล จะไม่สามารถปิดโปรแกรมได้...แต่ก็จะให้ทำ)
                    If (Trim(tMsgData) <> "") Then
                        
                        'ถ้า Service Transfer ไม่ทำงาน ให้เปิดใหม่
                        nVB_lServiceTransWnd = FindWindow(vbNullString, "ServiceTranfer")
                        If nVB_lServiceTransWnd = 0 Then
                            tFile = Left(App.Path, InStrRev(App.Path, "\") - 1) + "\POS\ServiceTranfer.exe"
                            If Not SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
                                    tFile = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path))
                                    tFile = tFile + "\POS\ServiceTranfer.exe"
                            End If
                            If SP_FLEbManFile(tFile, nEN_Fle1Exist) Then Call ShellEx(nVB_lServiceTransWnd, "open", tFile, "", "", 10)
                        End If
                        
                        'tMsg = "Please OK. to resend data or contact operator for further assistance." '*Tao 61-11-26 ของเดิม
                        tMsg = "Please Yes. to resend data and waiting for resend data" & vbCrLf & "or contact operator for further assistance."

                        tMsgALL = tMsgData & vbCrLf & tMsg
                        Call SP_STAxMouse(Me, False)
                        '*Tao 61-11-26 ของเดิม
'                        Call MsgBox(tMsgALL, vbOKOnly, tVB_CNAppTitle)
'                        Exit Sub

                        '*Tao 61-11-26 TK-ISS3000-714 TK ISSUE - TK Request - Modify exit terminal for resend transaction (2) (ไม่สามารถปิดโปรแกรมได้ เนื่องจาก Transfer ไม่ครบทุกบิล เลยให้ถามเหมือนเดิมว่าจะรอหรือจะออก)
                        If MsgBox(tMsgALL, vbYesNo, tVB_CNAppTitle) = vbYes Then
                            Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Prerss Yes button. and " & tMsgData)
                            Exit Sub
                        Else
                            Call SP_CRTxLogPrg("ErrorNumber:0" & " ErrorDesc:" & "Prerss No button. and " & tMsgData)
                            tMsgData = ""
                        End If
                    End If
                    '================================
                    
                    tMsgEJ = SP_PRCtCheckEJ
                    tMsg = "Do you want exit ? "
                    tMsgALL = ""
                    If (Trim(tMsgData) <> "") Or (Trim(tMsgEJ) <> "") Then       '*Ao 2013/11/22 ตรวจสอบข้อมูลก่อนออก
                        If Trim(tMsgData) <> "" Then
                            If Trim(tMsgEJ) <> "" Then
                                tMsgALL = tMsgData & vbCrLf & tMsgEJ & vbCrLf & vbCrLf & tMsg
                            Else
                                tMsgALL = tMsgData & vbCrLf & vbCrLf & tMsg
                            End If
                        Else
                            If Trim(tMsgEJ) <> "" Then
                                tMsgALL = tMsgEJ & vbCrLf & vbCrLf & tMsg
                            End If
                        End If
                        Call SP_STAxMouse(Me, False)
                        If Trim(tMsgALL) <> "" Then
                            'If MsgBox(tMsgALL, vbYesNo, tVB_CNAppTitle) = vbNo Then
                            If MsgBox(tMsgALL, vbYesNo + vbDefaultButton2, tVB_CNAppTitle) = vbNo Then  '*Bump 57-03-17 CommSheet TK-ISS3000-473 00 00 03.xlsx V.4.0.122 Build 20140317 14:38pm
                                Call SaveSetting("AdaPosMall", "POSFront", "SentSale", "1")     '*Ao 2013/11/22 ให้ส่งข้อมูล
                                Call SP_MOVxEJToOnline
                                Exit Sub
                            End If
                        End If
                    End If
                    Call W_PRCxSendEJold                    '*Ao 2014/03/12    CommSheet TK-ISS3000-473_Aoo1303
                    Call SP_TMNxUpdateOpenOnline            '*Ao 2014/04/01   ให้ UPDATE ตาม POS
            Case "6" 'EJ Viewer'*TON 58-12-01
                    Call SP_STAxMouse(Me, False)
                    wCNEJViewer.Show vbModal
                    
                    bW_KeyEnter = False
                    Exit Sub
            End Select
            Call SP_STAxMouse(Me, False) '*KT 52-12-17 POSFront 2.0.30
'            Unload Me
            If tVB_Transaction <> "99" Then
                Unload Me
                Call wBackGrand.W_PRVxFormActivate '*Ao 2011/03/28 CommSheet ML-PRV-046- POS Front Terminate Reboot and Not responding.xls
                Exit Sub   '*เพิ่ม
            End If
            
    End With
    Exit Sub
ErrHandle:
    Call SP_STAxMouse(Me, False)
End Sub

Private Sub ogdFunc_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    tW_HotKey = ""
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

Private Function W_GETbStaClosingReport() As Boolean
'----------------------------------------------------------
'   Cmt: ทำการ Update สถานะการทำ Closing Report  Y: ทำแล้ว  N: ยังไม่ทำ PH2.0.0
'   Ret:
'----------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    On Error GoTo ErrHandle

        tSql = "SELECT *  FROM TCNMTerminalMtn "
        tSql = tSql & " WHERE FTTmnStaClossing = 'N'" ''*KT 51-08-08 Comsheet ML-SF-451 ให้ Set = Y เมื่อทำ Print Closing Report
        tSql = tSql & " AND FTTmnNum='" & tVB_CNTerminalNum & "'"
       If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOffLine) = 0 Then
            If orsTemp.RecordCount > 0 Then
                W_GETbStaClosingReport = False
            Else
                W_GETbStaClosingReport = True
            End If
       Else
            W_GETbStaClosingReport = False
       End If
        Call SP_DATxRsNothing(orsTemp)
        Exit Function
ErrHandle:
        Call SP_MSGxShowError(Err.Number, Err.Description)
        Call SP_DATxRsNothing(orsTemp)
End Function

'Private Sub otmOnOff_Timer()           '*Ao 2014/05/08   ไม่ได้ใช้งาน
'    '*Bump 57-03-24 17.00 pm. V.4.0.123 CommSheet TK-ISS3000-473 00 00 05.xlsx
'    ''-----------------------------------------------------------
'    'Call:
'    'Cmt: Refresh สัญลักษณ์เขียวแดงของ Network
'    '----------------------------------------------------------
'        Dim bOnOff As Boolean
'        Dim bOnOffCur As Boolean
'        On Error GoTo ErrHandle
'        bOnOff = IIf(wMain.olaOnOff.Caption = "On-line", True, False)
'        bOnOffCur = oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP)
'        If bOnOffCur Then
'            wMain.oimOnOff.Picture = wMain.oilOnOff.ListImages(1).Picture
'            wMain.olaOnOff.Caption = "On-line"
'            If Not bOnOff And bOnOffCur Then
'                Call SP_CHKbOnlineEnable 'PH 2.0.0
'            End If
'        Else
'            wMain.oimOnOff.Picture = wMain.oilOnOff.ListImages(2).Picture
'            wMain.olaOnOff.Caption = "Off-line"
'        End If
'        '    DoEvents           '*KT 53-01-14
'        Exit Sub
'ErrHandle:
'        'Call SP_MSGxShowError(Err.Number, Err.Description)               '*Ao 2011/04/08    V 3.0.56
'End Sub

Private Sub otmSendEJ_Timer()
    '*************************************
    '*Bump 57-09-08 ทำไว้รอ change reqest ให้ Auto Login เข้าเครื่อง Server ก่อนส่ง EJ
    'If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then Call SP_PRCxNetUseCommand
    '*************************************
    '*Bump 57-03-24 17.00 pm. V.4.0.123 CommSheet TK-ISS3000-473 00 00 05.xlsx
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then Call W_PRCxSendEJold                     '*Bump 57-03-24 copy Ao 2014/03/12    CommSheet TK-ISS3000-473_Aoo1303
End Sub


