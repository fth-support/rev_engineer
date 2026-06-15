VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNPayVendor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Coupon;klg_Title"
   ClientHeight    =   3225
   ClientLeft      =   5220
   ClientTop       =   2775
   ClientWidth     =   6780
   ControlBox      =   0   'False
   Icon            =   "wCNPayVendor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6780
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2535
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   6765
      _Version        =   720897
      _ExtentX        =   11933
      _ExtentY        =   4471
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbNo 
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
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   0
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox otbNum 
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
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   1
         Text            =   "0"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label olaTotalDiscVendor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4365
         TabIndex        =   9
         Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Disc Vendor:;klg_TotalDisc"
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
         Left            =   240
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   1080
         Width           =   2250
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No :;klg_VendorCpnNo"
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
         Left            =   240
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :;klg_Amt"
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
         Left            =   240
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   1680
         Width           =   2280
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   2025
      Width           =   6765
      _Version        =   720897
      _ExtentX        =   11933
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   360
         Top             =   600
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
Attribute VB_Name = "wCNPayVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_Barcode As String

Private Sub W_FRMxSetInitial()
    otbNo.Text = ""
    olaTotalDiscVendor.Caption = SP_FMTtNum2Shw(0)
    otbNum.Text = 0
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    If bVB_Res640X480 Then
'        Me.Left = 0
        Me.Top = 2100
    End If
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
              '*Bump 56-02-28 เพิ่ม code 3 บรรทัด
            If .DeviceName = tSNR_TOUCH_SCANNER Then
                .DecodeData = True
            End If
        End With
    End If
    Call W_SUMcTotalVendorCpnTendor
    olaTotalDiscVendor.Caption = SP_FMTcCurPoint(cVB_TotalDiscVendorCpnAmt, False, True)
    cVB_Pay = 0 '*Tao 58-12-14 เคลียร์ค่าการชำระเงิน
    cVB_VendorCpnAmt = 0
    cVB_TotalDiscVendorCpnAmt = 0
    bVB_PayVendoeCancel = True
    Exit Sub
ErrHandle:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
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
    tVB_VendorCpnNo = Trim(otbNo.Text)
    cVB_VendorCpnAmt = 0
    cVB_TotalDiscVendorCpnAmt = 0
    cVB_Pay = 0
    bVB_PayVendoeCancel = True
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim cVoucherAmt As Double
    Dim cPay As Double
    Dim tSql$
    
    If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbNum.Text))) Then Exit Sub       '*AO51-03-26   Comm.Sheet  ML-SF-100

    cPay = SP_VALcText2Double(otbNum.Text)
    If Trim(otbNo.Text) = "" Then 'check amount zero
        Call SP_OBJxSetFocus(otbNo) ' otbNo.SetFocus
        Exit Sub
    End If
    
    If wSub Is Nothing Then
'        If cPay > cUT_SalePrice Then  'check amount > balance
'            Call SP_MSGnShowing("tms_CN005066", nCS_Error)
'            Exit Sub
'        End If
    Else
        If cPay > cVB_Balance Then   'check amount > balance
            Call SP_MSGnShowing("tms_CN005066", nCS_Error)
            Exit Sub
        End If
    End If

    tVB_VendorCpnNo = Trim(otbNo.Text)
    cVB_VendorCpnAmt = cPay
    cVB_TotalDiscVendorCpnAmt = SP_VALcText2Double(olaTotalDiscVendor.Caption)
    cVB_Pay = cPay
    bVB_PayVendoeCancel = False
    Unload Me
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
    With oScanner
        'otbNo.Text = .ScanData '*Bump 56-02-28  Comment บรรทัดนี้
        '*Bump 56-02-28 เพิ่ม code 8 บรรทัด
        If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
            .DecodeData = True
        End If
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            otbNo.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel) '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
        Else
            otbNo.Text = .ScanData
        End If
'        DoEvents '*Ao 53-01-14   DoEvents
        If otbNo.Text <> "" Then
            '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
            '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
            If Asc(Right(otbNo.Text, 1)) = 13 Then
                otbNo.Text = Mid(otbNo.Text, 1, Len(otbNo.Text) - 1)
            End If
            Call SP_OBJxSetFocus(otbNum)
        End If
        'ทำการ Open Scanner อีกครั้ง ไม่นั้นครั้งต่อไปจะ Scan สินค้าไม่ได้ครับ
        '*KT 52-01-08 ไม่ต้อง Open Scanner อีกครั้ง ยกเว้นกรณี Form Active
        If bVB_OPOSScannerEnable Then
            'With oScanner
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
            'End With
        End If
    End With
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & " Runtime : " & Err.Number & " " & Err.Description, vbCritical
End Sub

'Private Sub otbNo_GotFocus()
'    otbNo.SelStart = 0
'    otbNo.SelLength = Len(otbNo.Text)
'End Sub

Private Sub otbNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbNo.Text = ""
        Case 13
            Debug.Print "2"
            If otbNo.Text <> "" Then
                SendKeys "{TAB}"
            End If
    End Select
End Sub

Private Sub otbNo_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
End Sub

Private Sub otbNo_LostFocus()
    Dim tSql$
    Dim tTextPrompt$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle             '*KT 54-03-25 ป้องกันโปรแกรมหลุด
    Debug.Print "3"
   otbNo.Text = SP_TXTtTrimRightLeft(otbNo.Text)

    If Trim(otbNo.Text) = "" Then
        Call SP_DATxRsNothing(orsTemp)
        Exit Sub
    End If
     tTextPrompt = Trim(otbNo.Text)
     tTextPrompt = oSKU.GETVendorCpnBarcode(tTextPrompt)

    tSql = "SELECT FTTdmCode,FTVdcCode,FCVdcCpnAmt,FTVdcBarcode"
    tSql = tSql & vbCrLf & " FROM TCNMVndCoupon"
    tSql = tSql & vbCrLf & " WHERE FTTdmCode = '" & tVB_TenderCode & "'"
    tSql = tSql & vbCrLf & " AND FTVdcBarcode = '" & tTextPrompt & "'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
    If Not orsTemp Is Nothing Then  '*KT 54-03-25 ป้องกันโปรแกรมหลุด
        If Not (orsTemp.EOF Or orsTemp.BOF) Then
            'otbNo.Text = SP_FEDtChkString(orsTemp, "FTVdcBarcode")
            tW_Barcode = SP_FEDtChkString(orsTemp, "FTVdcBarcode")
            otbNum.Text = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCVdcCpnAmt"), False, True)
        Else
            'Call SP_MSGnShowing("tms_CN005135", nCS_Error)
            otbNum.Text = ""
        End If
    Else
        otbNum.Text = ""
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    otbNum.Text = ""
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_DATxRsNothing(orsTemp)
End Sub

Private Sub otbNum_Change()         '*AO51-03-12    Comm.Sheet SF- 070
    otbNum.Text = SP_DEFtChangeNum(Trim(otbNum.Text))
    otbNum.SelStart = Len(otbNum.Text)
End Sub

Private Sub otbNum_GotFocus()
    otbNum.SelStart = 0
    otbNum.SelLength = Len(otbNum.Text)
End Sub

Private Sub otbNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbNum.Text = ""
    End Select
End Sub
Private Sub otbNum_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 13
            Call ocmOK_Click
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub W_SUMcTotalVendorCpnTendor()
   Dim tWhe$
    tWhe = " WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & "  AND FTShdTransNo='" & tVB_Running & "'"
    tWhe = tWhe & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tWhe = tWhe & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tWhe = tWhe & "  AND FTTdmCode = '" & tVB_TenderCode & "'"
    cVB_TotalDiscVendorCpnAmt = SP_FMTcCurPoint(SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSrcNet", tVB_CNTblRC, tWhe)), True, False)
End Sub


