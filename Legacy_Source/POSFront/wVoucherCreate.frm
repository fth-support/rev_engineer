VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wVoucherCreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voucher;klg_Title"
   ClientHeight    =   2505
   ClientLeft      =   3180
   ClientTop       =   4020
   ClientWidth     =   6780
   ControlBox      =   0   'False
   Icon            =   "wVoucherCreate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6780
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1875
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -270
      Width           =   6765
      _Version        =   720897
      _ExtentX        =   11933
      _ExtentY        =   3307
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbVoucherNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   0
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox otbVoucherAmt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   1
         Text            =   "0"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No:;klg_VoucherNo"
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
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   480
         Width           =   2130
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Amount:;klg_VoucherAmt"
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
         Left            =   120
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   1080
         Width           =   2175
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   6765
      _Version        =   720897
      _ExtentX        =   11933
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   600
         Top             =   360
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
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
         Left            =   5160
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
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
         Left            =   3600
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wVoucherCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bW_FoundVoucher As Boolean

Private Sub W_PRCxCreateVoucher()
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo HandleErr
    If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbVoucherAmt.Text))) Then Exit Sub       '*AO51-03-26   Comm.Sheet  ML-SF-100
    
    If otbVoucherNo.Text = "" Then
        Call SP_OBJxSetFocus(otbVoucherNo)
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
    ElseIf SP_VALcText2Double(otbVoucherAmt.Text) <= 0 Then
        Call SP_OBJxSetFocus(otbVoucherAmt)
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
    End If
    bVB_AddVoucher = True
    tVB_VoucherNo = SP_TXTtTrimRightLeft(otbVoucherNo.Text) 'Trim(otbVoucherNo.Text)
    'Check Voucher No duplicate Comsheet ML-SF-254
    tSql = "SELECT FTShdVoucherNo,FTShdStaCancel  FROM TPSTVoucherHD"
    tSql = tSql & vbCrLf & "WHERE FTShdVoucherNo='" & tVB_VoucherNo & "'"
    tSql = tSql & vbCrLf & "AND  FTShdTransType = '" & IIf(bVB_TrainingMode, Format(tEN_TrnTraining, "00"), Format(tEN_TrnVoucher, "00")) & "'"
    tSql = tSql & vbCrLf & "AND FTShdStaDoc  <>  '3'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then 'PH 2.0.0 ติดต่อฐานข้อมูล server ตลอด
            MsgBox "You can not connect database server.", vbCritical
            Call SP_OBJxSetFocus(otbVoucherNo)
            Call SP_DATxRsNothing(orsTemp)
            Exit Sub
    End If
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        If SP_FEDtChkString(orsTemp, "FTShdStaCancel") = "N" Then
            Call SP_MSGnShowing("tms_CN005145", nCS_Warning)
            Call SP_OBJxSetFocus(otbVoucherNo)
            Call SP_DATxRsNothing(orsTemp)
            Exit Sub
        End If
    End If

    cVB_VoucherAmt = SP_VALcText2Double(otbVoucherAmt.Text)
    tSql = "UPDATE " & tVB_CNTblHD
    tSql = tSql & " SET FTShdVoucherNo='" & tVB_VoucherNo & "'"
    tSql = tSql & " ,FCShdVoucherAmt=" & cVB_VoucherAmt
    tSql = tSql & " ,FCShdTotal=" & cVB_VoucherAmt
    tSql = tSql & " ,FCShdGrand=" & cVB_VoucherAmt
    tSql = tSql & "," & UT_tLastUpdSql
    tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & "  AND FTShdStaDoc = '2'"
    Call SP_SQLbExecute(tSql)
    Call SP_DATxRsNothing(orsTemp)
    Unload Me
    Exit Sub
HandleErr:
     If Err.Number = -2147467259 Or Err.Number = -2147217871 Or Err.Number = 91 Or Err.Number = 3704 Then '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
        MsgBox "You can not connect database server.", vbCritical
        Call SP_OBJxSetFocus(otbVoucherNo)
        Call SP_DATxRsNothing(orsTemp)
    Else
        Call SP_MSGxShowError(Err.Number, Err.Description)
        Call SP_OBJxSetFocus(otbVoucherNo)
        Call SP_DATxRsNothing(orsTemp)
    End If
    Exit Sub
End Sub

Private Sub W_PRCxCancelVoucher()
    Dim orsTemp As ADODB.Recordset
    Dim tSql$
On Error GoTo HandleErr
    If otbVoucherNo.Text = "" Then
        Call SP_OBJxSetFocus(otbVoucherNo)
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
    End If
    If SP_VALcText2Double(otbVoucherAmt.Text) > cVB_CancelVoucherAmt Then
        Call SP_MSGnShowing("tms_CN005064", nCS_Error)
        Call SP_OBJxSetFocus(otbVoucherAmt)
        Call SP_OTBxSelText(otbVoucherAmt)
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
    End If
    tSql = "SELECT FTTmnNum,FTShdTransNo,FDShdTransDate,FCShdVoucherAmt "
    tSql = tSql & vbCrLf & "FROM TPSTVoucherHD "
    tSql = tSql & vbCrLf & "WHERE   FTShdVoucherNo = '" & SP_TXTtTrimRightLeft(otbVoucherNo.Text) & "'" 'Trim(otbVoucherNo.Text) & "'"
    tSql = tSql & vbCrLf & "AND  ISNULL(FTShdStaCancel,'N') <> 'Y'"
    tSql = tSql & vbCrLf & "AND  FTShdTransType = '" & IIf(bVB_TrainingMode, Format(tEN_TrnTraining, "00"), Format(tEN_TrnVoucher, "00")) & "'"
    tSql = tSql & vbCrLf & "AND  (FTTmnNumPay = '' OR  FTTmnNumPay IS NULL)"
    tSql = tSql & vbCrLf & "AND  (FTTransNoPay = '' OR  FTTransNoPay IS NULL)"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then 'PH 2.0.0 ติดต่อฐานข้อมูล server ตลอด
        Call SP_MSGnShowing("tms_CN005072", nCS_Error)
        Call SP_OBJxSetFocus(otbVoucherNo)
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
    End If
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        tVB_CancelVoucherNo = SP_TXTtTrimRightLeft(otbVoucherNo.Text) 'Trim(otbVoucherNo.Text)
        cVB_CancelVoucherAmt = SP_FEDvChkNumeric(orsTemp, "FCShdVoucherAmt")
        tVB_CancelTmnNumRef = SP_FEDtChkString(orsTemp, "FTTmnNum")
        tVB_CancelTransNoRef = SP_FEDtChkString(orsTemp, "FTShdTransNo")
        dVB_CancelTransDateRef = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
        bVB_CancelVoucher = True
        Unload Me
    Else
        Call SP_OBJxSetFocus(otbVoucherNo)
        Call SP_MSGnShowing("tms_CN005072", nCS_Error)
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
HandleErr:
     If Err.Number = -2147467259 Or Err.Number = -2147217871 Or Err.Number = 91 Or Err.Number = 3704 Then '*KT 54-04-07 ป้องกันโปรแกรมหลุด 3.0.56
        Call SP_MSGnShowing("tms_CN005072", nCS_Error)
        Call SP_OBJxSetFocus(otbVoucherNo)
        Call SP_DATxRsNothing(orsTemp)
    Else
        Call SP_MSGxShowError(Err.Number, Err.Description)
        Call SP_OBJxSetFocus(otbVoucherNo)
        Call SP_DATxRsNothing(orsTemp)
    End If
    Exit Sub
End Sub

Private Sub W_FRMxSetInitial()
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnVoucher:
                otbVoucherNo.Text = ""
                otbVoucherAmt.Text = 0
                bVB_AddVoucher = False
'                Call SP_OBJxSetFocus(otbVoucherNo)
        Case tEN_TrnCancelVoucher:
                otbVoucherNo.Text = ""
                otbVoucherAmt.Text = 0
                bVB_AddVoucher = False
'                Call SP_OBJxSetFocus(otbVoucherNo)
    End Select
End Sub


'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
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
            '*Bump 56-02-28 Touch Scanner เพิ่ม code 3 บรรทัด
            If .DeviceName = tSNR_TOUCH_SCANNER Then
                    .DecodeData = True
            End If
       End With
    End If
'    Call SP_OBJxSetFocus(otbVoucherNo)
'    DoEvents
    Exit Sub
ErrHandle:
    
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
'    With oScanner
'        If bVB_OPOSScannerEnable Then
'            If .DeviceEnabled Then
'                .DeviceEnabled = False
'                .Release
'                .Close
'            End If
'        End If
'    End With
'    Call SP_OBJxSetNothing(Me)
'End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    With oScanner
        If bVB_OPOSScannerEnable Then
            If .DeviceEnabled Then
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
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnVoucher:  bVB_AddVoucher = False
        Case tEN_TrnCancelVoucher: bVB_CancelVoucher = False
    End Select
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnVoucher: Call W_PRCxCreateVoucher
        Case tEN_TrnCancelVoucher: Call W_PRCxCancelVoucher
    End Select
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
    With oScanner
        'otbVoucherNo.Text = .ScanData  '*Bump 56-02-28 Comment บรรทัดนี้
         '*Bump 56-02-28 เพิ่ม Code 8 บรรทัด
              '***************************************
                If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
                    .DecodeData = True
                End If
                If .DeviceName = tSNR_TOUCH_SCANNER Then
                    otbVoucherNo.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel)  '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
                Else
                    otbVoucherNo.Text = .ScanData
                End If
                '***************************************
'        DoEvents *AO 53-01-14   DoEvents
        If otbVoucherNo.Text <> "" Then
            '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
            '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
            If Asc(Right(otbVoucherNo.Text, 1)) = 13 Then
                otbVoucherNo.Text = Mid(otbVoucherNo.Text, 1, Len(otbVoucherNo.Text) - 1)
            End If
           ' Call otbVoucherNo_KeyDown(13, 0)
            Call otbVoucherNo_KeyPress(13)
        End If
        'ทำการ Open Scanner อีกครั้ง ไม่นั้นครั้งต่อไปจะ Scan สินค้าไม่ได้ครับ
        '*KT 52-01-08 ไม่ต้อง Open Scanner อีกครั้ง ยกเว้นกรณี Form Active
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
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub otbVoucherAmt_Change()          '*AO51-03-12    Comm.Sheet SF- 070
    otbVoucherAmt.Text = SP_DEFtChangeNum(Trim(otbVoucherAmt.Text))
    otbVoucherAmt.SelStart = Len(otbVoucherAmt.Text)
End Sub

Private Sub otbVoucherAmt_GotFocus()
    otbVoucherAmt.SelStart = 0
    otbVoucherAmt.SelLength = Len(otbVoucherAmt.Text)
End Sub


Private Sub otbVoucherAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbVoucherAmt.Text = ""
    End Select
End Sub

Private Sub otbVoucherAmt_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbVoucherAmt.Text), ".") > 0 Then
'                KeyAscii = 0
'            End If
        Case 13
            Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbVoucherNo_GotFocus()
    otbVoucherNo.SelStart = 0
    otbVoucherNo.SelLength = Len(otbVoucherNo.Text)
End Sub

Private Sub otbVoucherNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbVoucherNo.Text = ""
        Case 13
            If Trim(otbVoucherNo.Text) <> "" Then
                SendKeys "{TAB}"
            End If
    End Select
End Sub

Private Sub otbVoucherNo_KeyPress(KeyAscii As Integer)
    Dim orsTemp As ADODB.Recordset
    Dim tSql$
On Error GoTo ErrHandle
    Call UT_ChkKeyAscii(KeyAscii)
    If KeyAscii = 13 Then
        Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnVoucher
            '*KT 53-01-20 มีการใช้ SendKeys "{TAB}"
'            If Trim(otbVoucherNo.Text) <> "" Then
'                Call SP_OBJxSetFocus(otbVoucherAmt)
'            Else
'                Call SP_OBJxSetFocus(otbVoucherNo)
'            End If
        Case tEN_TrnCancelVoucher
            If Trim(otbVoucherNo.Text) = "" Then
'                Call SP_OBJxSetFocus(otbVoucherNo) '*KT 53-01-20 มีการใช้ SendKeys "{TAB}"
                Call SP_DATxRsNothing(orsTemp)
                Exit Sub
            End If
            
            tSql = "SELECT FCShdVoucherAmt "
            tSql = tSql & vbCrLf & "FROM TPSTVoucherHD "
            tSql = tSql & vbCrLf & "WHERE   FTShdVoucherNo = '" & SP_TXTtTrimRightLeft(otbVoucherNo.Text) & "'"  'Trim(otbVoucherNo.Text) & "'"
            tSql = tSql & vbCrLf & "AND  ISNULL(FTShdStaCancel,'N') <> 'Y'"
            tSql = tSql & vbCrLf & "AND  FTShdTransType = '" & IIf(bVB_TrainingMode, Format(tEN_TrnTraining, "00"), Format(tEN_TrnVoucher, "00")) & "'"
            tSql = tSql & vbCrLf & "AND  (FTTmnNumPay = '' OR  FTTmnNumPay IS NULL)"
            tSql = tSql & vbCrLf & "AND  (FTTransNoPay = '' OR  FTTransNoPay IS NULL)"
            If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then 'PH 2.0.0 ติดต่อฐานข้อมูล server ตลอด
                cVB_CancelVoucherAmt = 0
                otbVoucherAmt.Text = SP_FMTcCurPoint(cVB_CancelVoucherAmt, False, True)
                Call SP_MSGnShowing("tms_CN005072", nCS_Error)
                Call SP_OBJxSetFocus(otbVoucherNo)
                Call SP_OTBxSelText(otbVoucherNo)
                Call SP_DATxRsNothing(orsTemp)
                Exit Sub
            End If
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                cVB_CancelVoucherAmt = SP_FEDvChkNumeric(orsTemp, "FCShdVoucherAmt")
                otbVoucherAmt.Text = SP_FMTcCurPoint(cVB_CancelVoucherAmt, False, True)
                'Call SP_OBJxSetFocus(otbVoucherAmt) *KT 53-01-20 มีการใช้ SendKeys "{TAB}"
            Else
                cVB_CancelVoucherAmt = 0
                otbVoucherAmt.Text = SP_FMTcCurPoint(cVB_CancelVoucherAmt, False, True)
                Call SP_MSGnShowing("tms_CN005072", nCS_Error)
                Call SP_OBJxSetFocus(otbVoucherNo)
                Call SP_OTBxSelText(otbVoucherNo)
            End If
        End Select
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    otbVoucherAmt.Text = ""
    otbVoucherNo.Text = ""
    Call SP_OBJxSetFocus(otbVoucherNo)
    Call SP_DATxRsNothing(orsTemp)
End Sub

Private Sub otbVoucherNo_LostFocus()
    otbVoucherNo.Text = SP_TXTtTrimRightLeft(otbVoucherNo.Text)
End Sub

