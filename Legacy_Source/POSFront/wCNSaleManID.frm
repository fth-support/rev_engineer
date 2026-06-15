VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNSaleManID 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALESMAN ID;klg_Title"
   ClientHeight    =   2430
   ClientLeft      =   6090
   ClientTop       =   3210
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "wCNSaleManID.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4800
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1815
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -270
      Width           =   4785
      _Version        =   720897
      _ExtentX        =   8440
      _ExtentY        =   3201
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   360
         MaxLength       =   9
         TabIndex        =   0
         Text            =   "12345/678"
         Top             =   720
         Width           =   4140
      End
      Begin VB.CommandButton ocmBrw 
         Height          =   495
         Left            =   4815
         TabIndex        =   3
         Top             =   1620
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALESMAN ID;klg_SaleID"
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
         Left            =   360
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   360
         Width           =   3210
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1230
      Width           =   4785
      _Version        =   720897
      _ExtentX        =   8440
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   360
         Top             =   480
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Timer otmLock 
         Interval        =   1500
         Left            =   0
         Top             =   480
      End
      Begin VB.CommandButton ocmCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel(Esc);klg_Cancel"
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
         Left            =   3120
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton ocmOk 
         Caption         =   "&Ok(F9);klg_Ok"
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
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNSaleManID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'วันที่ 15/01/2007
'พี่บอกว่าได้คุยกับ The Mall แล้วว่ากรณี Saleman นั้นจะไม่ table เก็บให้ทาง front key หมายเลขเองเลย
'โดยเก็บข้อมูลหมายเลขที่ key ไป ไม่มีการป้องกัน

Option Explicit
Const tW_StdHD = "Shd"
Const tW_StdDT = "Sdt"
Dim bW_FirstLoad As Boolean


Private Sub W_FRMxSetInitial()
        otbID.Text = ""
        Call SP_OBJxSetFocus(otbID)
        bW_FirstLoad = True
        bVB_SalemanIDHD = False             '*Ao 2013/04/18 ให้เก็บ HD ใหม่ทุกครั้ง
 End Sub
    
Private Sub Form_Activate()
On Error GoTo ErrHandle
    If bW_FirstLoad Then
         If otbID.Text <> "" Then
            Call SP_OTBxSelText(otbID)
        Else
             Call SP_OBJxSetFocus(otbID)
        End If
    End If
    bW_FirstLoad = False
    Exit Sub
ErrHandle:
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '*KT 52-08-17 Locked Keyboard
    Select Case Shift
        Case 2
            Select Case KeyCode
                Case vbKeyF1
                        Call SP_OBJxSetFocus(otbID)
                        Call otbID_KeyDown(vbKeyF1, 2)
                        KeyCode = 0
            End Select
        Case Else
             If bUT_FuncLockKB Then
                KeyCode = 0
             End If
     End Select
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    '*KT 52-08-17 Locked Keyboard
'     If bUT_FuncLockKB = False Then
'         If KeyAscii = 13 Then SendKeys "{TAB}"
'    End If
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
     bVB_FormPayAct = True 'PH 2.0 ML-SF-019 ไม่ต้องการให้ทำที่ Form_Activate ของ wMain
'    otbID.Text = ""
    If bVB_OPOSScannerEnable Then
        With oScanner
            .Open tVB_Scanner
            .Claim 0 'TeamPos2000
            'Enable the device.
            .DeviceEnabled = True
            'Enable the event.
            .DataEventEnabled = True
            .AutoDisable = True
            .ClearInput
            '*Bump 56-02-28 Touch Scanner
            If .DeviceName = tSNR_TOUCH_SCANNER Then
                    .DecodeData = True
            End If
        End With
    End If
    Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock ใหม่  '*KT 52-08-17 Locked Keyboard
    '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
    bVB_SalemanIDChg = False
    otbID.Text = tVB_SalemanIDOld
    If otbID.Text <> "" Then
        Call SP_OTBxSelText(otbID)
    End If
    Exit Sub
ErrHandle:
    bVB_SalemanIDChg = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    otmLock.Enabled = False
    With oScanner
         If bVB_OPOSScannerEnable Then
             If .DeviceEnabled = True Then
                .DeviceEnabled = False
                .Release
                .Close
            End If
        End If
    End With
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    '*KT 52-08-17 Locked Keyboard
    If bUT_FuncLockKB = False Then
        Unload Me
        bVB_SalemanIDChg = False
    End If
End Sub

Private Sub ocmOK_Click()
     '*KT 52-08-17 Locked Keyboard
'     If Len(Trim(otbID.Text)) <> 0 And Len(Trim(otbID.Text)) <> 5 And Len(Trim(otbID.Text)) <> 9 Then  '*Eaw 56-05-30 CommSheet TK-ISS3000-207-ST 00 00 01.xls
'        Call SP_MSGnShowing("tms_CN005246", nCS_Warning)
'        otbID.Text = ""
'        Exit Sub
'     End If
     Select Case Len(Trim(otbID.Text))  '*Eaw 56-06-24 CommSheet TK-ISS3000-207-ST 00 00 04 FSBT NG.xlsx
     Case 0
        Exit Sub
     Case 5, 9
        'none
     Case Else
        Call SP_MSGnShowing("tms_CN005246", nCS_Warning)
        otbID.Text = ""
        Exit Sub
     End Select
     If bUT_FuncLockKB = False Then
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnOpen:
            Case tEN_TrnStartBank:
            Case tEN_TrnSale: Call W_PRCxSale
'            Case tEN_TrnReturn:
            Case tEN_TrnReturn: Call W_PRCxSale  '*BG 56-11-08 CommSheet TK-ISS3000-417-FAPL-148_BG.xlsx
            Case tEN_TrnIPV:
            Case tEN_TrnDeposit: Call W_PRCxSale
            Case tEN_TrnRedeem:
            Case tEN_TrnClose:
            Case tEN_TrnPrintClose:
'            Case tEN_TrnVIPSale: Call W_PRCxVIPSale            '*Ao 2013/04/18  CommSheet TK-ISS3000-151-IT2 00.00.01.xls
            Case tEN_TrnVIPSale: Call W_PRCxSale
            Case tEN_TrnTaxExceptSale: Call W_PRCxSale '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
            Case tEN_TrnPickUp:
            Case tEN_TrnDutyFreeSale: Call W_PRCxSale '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
            Case tEN_TrnIPVReturn:
            Case tEN_TrnVoucher:
            Case tEN_TrnDepositCancel:
            Case tEN_TrnIPVRedeem:
            Case tEN_TrnNoSale:
            Case tEN_TrnAddDeposit:
            Case tEN_TrnAddDepositCancel:
            Case tEN_TrnCancelVoucher:
            Case tEN_TrnOtherReceive: Call W_PRCxSale '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
            Case tEN_TrnOtherDisbursement: Call W_PRCxSale '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
            Case tEN_TrnDepositMoney:
            Case tEN_TrnTraining:
            Case tEN_TrnIPVVIPService:
            Case tEN_TrnIPVTaxExceptSale:
            Case tEN_TrnIPVDutyFreeSale:
            Case tEN_TrnSaleTaxInc: Call W_PRCxSale
            Case tEN_TrnSaleManualTax: Call W_PRCxSale
            Case tEN_TrnReturnTaxInc: Call W_PRCxSale  '*BG 56-11-08  CommSheet TK-ISS3000-417-FAPL-148_BG.xlsx
            Case tEN_TrnReturnManualTax: Call W_PRCxSale  '*BG 56-11-08 CommSheet TK-ISS3000-417-FAPL-148_BG.xlsx
            Case tEN_TrnCreateVIP: Call W_PRCxSale '*BG 56-11-08 CommSheet TK-ISS3000-417-FAPL-148_BG.xlsx
            
        End Select
        Unload Me
    End If
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
     '*KT 52-08-17 Locked Keyboard
    If bUT_FuncLockKB = False Then
        '*KT 52-06-05 PH 2.0 ML-SF-020 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
        With oScanner
            'otbID.Text = .ScanData '*Bump 56-02-28 Comment บรรทัดนี้
            '*Bump 56-02-28 เพิ่ม Code 8 บรรทัด
              '***************************************
                If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
                    .DecodeData = True
                End If
                If .DeviceName = tSNR_TOUCH_SCANNER Then
                    otbID.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel)  '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
                Else
                    otbID.Text = .ScanData
                End If
                '***************************************
'            DoEvents '*Ao 53-01-14   DoEvents
            If otbID.Text <> "" Then
                If Asc(Right(otbID.Text, 1)) = 13 Then
                    otbID.Text = Mid(otbID.Text, 1, Len(otbID.Text) - 1)
                End If
            End If
            If bVB_OPOSScannerEnable Then
                    If .DeviceEnabled = False Then
                         .Open tVB_Scanner
                        .Claim 0 'TeamPos2000
                         'Enable the device.
                         .DeviceEnabled = True
                         'Enable the event.
                         .DataEventEnabled = True
                         .AutoDisable = True
                         .ClearInput
                    Else
                        .ClearInput
                    End If
            End If
        End With
        Call SP_OBJxSetFocus(ocmOk)
    End If
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub otbID_GotFocus()
    Call SP_OBJxSetFocus(otbID)
End Sub

Private Sub otbID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case vbKeyDelete         'delete
                    otbID.Text = ""
                Case 13
                    Call ocmOK_Click
            End Select
        Case 2 '*KT 52-08-17 Locked Keyboard
            Select Case KeyCode
                Case vbKeyF1
                    wFunctionKB.W_KBLockKB
                    If bUT_FuncLockKB Then
                        otbID.Text = "Locked"
                        Call SP_FRMxSetTimeCount(False)       'หยุดการจับเวลา Lock
                        otmLock.Enabled = False
                    Else
                        If otbID.Text = "Locked" Then
                            otbID.Text = ""
                        End If
                        Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock
                        otmLock.Enabled = True
                    End If
            End Select
    End Select
End Sub

Private Sub W_PRCxSale()
'----------------------------------------------------------
'   Call:
'   Cmt : *KT 50-12-26  Update HD เกี่ยวกับพนักงานขาย การขายแบบ Sale
'              *KT 52-10-21 ป้องกันกรณีที่ Text Box มีอักขระ Enter
'*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
'----------------------------------------------------------
    Dim tSql$
On Error GoTo ErrHandle
    otbID.Text = SP_TXTtTrimRightLeft(otbID.Text)
    If Not bVB_SalemanIDHD Then
        tSql = "UPDATE " & tVB_CNTblHD
        tSql = tSql & " SET FTSpnCode = '" & Trim(otbID.Text) & "'"                     '*Ao 2013/04/18 รวมเครื่องหมาย /
        tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
        tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
        tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
        tSql = tSql & "  AND FTShdStaDoc = '2'"
        Call SP_SQLvExecute(tSql)
        bVB_SalemanIDHD = True
        tVB_SalemanID = Trim(otbID.Text)
        bVB_SalemanIDChg = True                             '*Ao 2013/04/18  CommSheet TK-ISS3000-151-IT2 00.00.01.xls
        tVB_SalemanIDChg = Trim(otbID.Text)         'UPDATE HD ทุกครั้ง
        '*Bump 56-04-26 CommSheet TK-ISS3000-160-IT2 00.00.02.xls
        tSql = "UPDATE " & tVB_CNTblDT
        tSql = tSql & " SET FTSdtSpnCode = '" & Trim(otbID.Text) & "'"
        tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
        tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
        tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
        tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
        Call SP_SQLvExecute(tSql)
        '*************
    Else
        bVB_SalemanIDChg = True
        If bVB_Saleman Then tVB_SalemanID = Trim(otbID.Text)
        tVB_SalemanIDChg = Trim(otbID.Text)
    End If
    
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    bVB_SalemanIDChg = False
    tVB_SalemanID = ""
End Sub

'Private Sub W_PRCxVIPSale()                    '*Ao 2013/04/18  CommSheet TK-ISS3000-151-IT2 00.00.01.xls
''----------------------------------------------------------
''   Call:
''   Cmt : *KT 50-12-26  Update HD เกี่ยวกับพนักงานขาย การขายแบบ VIP Service Sale
''              *KT 52-10-21 ป้องกันกรณีที่ Text Box มีอักขระ Enter
''*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
''----------------------------------------------------------
'    Dim tSql$
'On Error GoTo ErrHandle
'    otbID.Text = SP_TXTtTrimRightLeft(otbID.Text)
'    If Not bVB_SalemanIDHD Then
'        tSql = "UPDATE " & tVB_CNTblHD
'        tSql = tSql & " SET FTSpnCode = '" & Trim(otbID.Text) & "'"
'        tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
'        tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
'        tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
'        tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
'        tSql = tSql & "  AND FTShdStaDoc = '2'"
'        Call SP_SQLvExecute(tSql)
'        bVB_SalemanIDHD = True
'        tVB_SalemanID = Trim(otbID.Text)
'    Else
'        bVB_SalemanIDChg = True
'        If bVB_Saleman Then tVB_SalemanID = Trim(otbID.Text)
'    End If
'
'    Exit Sub
'ErrHandle:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'    bVB_SalemanIDChg = False
'    tVB_SalemanID = ""
'End Sub
'
'Private Sub W_PRCxDeposit()
''----------------------------------------------------------
''   Call:
''   Cmt : *KT 50-12-26  Update HD เกี่ยวกับพนักงานขาย การขายแบบ Sale
''              *KT 52-10-21 ป้องกันกรณีที่ Text Box มีอักขระ Enter
''              *KT 52-11-16 PH3-ML-SF-013 Saleman ID on Deposit Tran
''*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
''----------------------------------------------------------
'    Dim tSql$
'On Error GoTo ErrHandle
'    otbID.Text = SP_TXTtTrimRightLeft(otbID.Text)
'     If Not bVB_SalemanIDHD Then
'         tSql = "UPDATE " & tVB_CNTblHD
'        tSql = tSql & " SET FTSpnCode = '" & Trim(otbID.Text) & "'"
'        tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
'        tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
'        tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
'        tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
'        tSql = tSql & "  AND FTShdStaDoc = '2'"
'        Call SP_SQLvExecute(tSql)
'        bVB_SalemanIDHD = True
'        tVB_SalemanID = Trim(otbID.Text)
'    Else
'        bVB_SalemanIDChg = True
'        If bVB_Saleman Then tVB_SalemanID = Trim(otbID.Text)
'     End If
'
'    Exit Sub
'ErrHandle:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'    bVB_SalemanIDChg = False
'    tVB_SalemanID = ""
'End Sub

Private Sub otbID_KeyPress(KeyAscii As Integer)
    '*KT 52-08-17 Locked Keyboard
    If bUT_FuncLockKB = False Then
        Call UT_ChkKeyAscii(KeyAscii)
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub otmLock_Timer()
     '*KT 52-08-17 Locked Keyboard
On Error GoTo ErrHandle
    If nW_TimerLock = 0 And Screen.ActiveForm.Name = "wCNSaleManID" Then
         wFunctionKB.W_KBLockKB
        If bUT_FuncLockKB Then
            otbID.Text = "Locked"
            Call SP_FRMxSetTimeCount(False)       'หยุดการจับเวลา Lock
            otmLock.Enabled = False
        Else
            If otbID.Text = "Locked" Then
                otbID.Text = ""
            End If
            Call SP_FRMxSetTimeCount(True, nW_TimerLock)       'จับเวลา Lock
            otmLock.Enabled = True
        End If
    Else
        If bUT_FuncLockKB Then
            otbID.Text = "Locked"
        Else
            If otbID.Text = "Locked" Then
                otbID.Text = ""
            End If
        End If
    End If
    Exit Sub
ErrHandle:
End Sub

Private Sub otbID_Change()  '*Ao 2013/04/18  CommSheet TK-ISS3000-151-IT2 00.00.01.xls
Dim tManID As String
    If Trim(otbID.Text) = "Locked" Then
        Exit Sub
    End If
    tManID = Replace(Trim(otbID.Text), "/", "")
    If Len(tManID) > 5 Then
        otbID.Text = SP_PRNtSalesManID(tManID)
        otbID.SelStart = Len(otbID.Text)
    End If
    
End Sub

