VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNDiscCpn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discount Coupon;klg_Title"
   ClientHeight    =   2610
   ClientLeft      =   8430
   ClientTop       =   4290
   ClientWidth     =   6540
   ControlBox      =   0   'False
   Icon            =   "wCNDiscCpn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6540
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1935
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   6525
      _Version        =   720897
      _ExtentX        =   11509
      _ExtentY        =   3413
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbCpnNo 
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
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   0
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox otbAmt 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Coupon No:;klg_CpnNo"
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
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   555
         Width           =   1995
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:;klg_Amt"
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
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   1155
         Width           =   1935
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1410
      Width           =   6525
      _Version        =   720897
      _ExtentX        =   11509
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   1800
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
         Left            =   4800
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
         Left            =   3240
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNDiscCpn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bW_ChkBarcode As Boolean

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)         '*Ao 53-01-14
'    If KeyCode = 13 Then SendKeys "{TAB}"
'End Sub


Private Sub Form_Load()
On Error GoTo ErrHandle
    bW_ChkBarcode = False
    bVB_DiscAccept = False
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
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
    Exit Sub
ErrHandle:
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = 1
'End Sub

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

'Private Sub Form_Unload(Cancel As Integer)
'    Set wCNDiscCpn = Nothing
'End Sub

Private Sub ocmCancel_Click()
    bVB_DiscAccept = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim cDiscAmt As Double
    '*KT 54-07-26 V.4.0.1 Phase4.0(Member Point) Promotion Coupon ตรวจสอบ barcode config
    If Not bW_ChkBarcode Then
        If otbCpnNo.Text <> "" Then
            cDiscAmt = oSKU.GETPmtCpnBarcode(Trim(otbCpnNo.Text))
            If oSKU.GETResults Then
                bW_ChkBarcode = True
                If oSKU.GETInStoreBar Then
                    If cDiscAmt > 0 Then
                        otbAmt.Text = SP_FMTcCurPoint(cDiscAmt, True, True)
                        otbAmt.Locked = True
                        otbAmt.BackColor = nVB_CNClrBLock
                    Else
                        otbAmt.Text = "0.00"
                        otbAmt.Locked = False
                        otbAmt.BackColor = nVB_CNClrBUnLock
                    End If
                Else
                    otbAmt.Text = "0.00"
                    otbAmt.Locked = False
                    otbAmt.BackColor = nVB_CNClrBUnLock
                End If
            Else
                Call SP_MSGnShowing("tms_CN005224", nCS_Warning)
                otbCpnNo.Text = ""
                otbAmt.Text = "0.00"
                otbAmt.Locked = False
                otbAmt.BackColor = nVB_CNClrBUnLock
                Call SP_OBJxSetFocus(otbCpnNo)
                Exit Sub
            End If
        End If
    End If
    
'*Eaw 56-07-04 CommSheet TK-ISS3000-201-ST.xls ไม่ต้องเช็คเหรียญ
'    If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbAmt.Text))) Then Exit Sub      '*AO51-03-26   Comm.Sheet  ML-SF-100
   
    cVB_MallCpnAmt = SP_VALcText2Double(otbAmt.Text)
    oDisc.DISxClearAll
    oDisc.dTranDate = tUT_SaleDate
    oDisc.tTerminalNo = tVB_CNTerminalNum
    oDisc.tTransNo = tVB_Running
    oDisc.cDiscValue = cVB_MallCpnAmt
    oDisc.tCardID = oSKU.GETID ' Trim(otbCpnNo.Text) 'หมายเลขที่แสดงที่หน้าจอช่อง Coupon No เป็นหมายเลขที่ตัดเฉพาะเลขที่ Running ถ้าตรงกันกับ Barcode Configuration ส่วน ID ที่เก็บลงตาราง CD คือหมายเลขที่สแกนหรือคีย์มาแต่แรก อ้างอิงเอกสาร PRJ09010TheMall-R4-SE-External Spec.Promotion Coupon-00.04.00-FSBT.doc หน้า 22-23
    oDisc.bMallCpnInStoreBar = oSKU.GETInStoreBar '*KT 54-07-26 V.4.0.1 Phase4.0(Member Point) Promotion Coupon ส่ง parameter
    Call oDisc.DIScDiscount(nEN_Disc08MallCpn, eEN_VerifyDiscType.nEN_Type2ByDcn)
    bVB_DiscAccept = True
    Unload Me
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
    Dim cDiscAmt As Double
On Error GoTo ErrHandle
    With oScanner
        'otbCpnNo.Text = .ScanData '*Bump 56-02-28 Comment บรรทัดนี้
        '*Bump 56-02-28 เพิ่ม code 8 บรรทัด
        If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
            .DecodeData = True
        End If
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            otbCpnNo.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel) '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
        Else
            otbCpnNo.Text = .ScanData
        End If
'        DoEvents            '*KT 53-01-14
        If otbCpnNo.Text <> "" Then
            '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
            '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
            If Asc(Right(otbCpnNo.Text, 1)) = 13 Then
                otbCpnNo.Text = Mid(otbCpnNo.Text, 1, Len(otbCpnNo.Text) - 1)
            End If
            '*KT 54-07-26 V.4.0.1 Phase4.0(Member Point) Promotion Coupon ตรวจสอบ barcode config
            cDiscAmt = oSKU.GETPmtCpnBarcode(Trim(otbCpnNo.Text))
            If oSKU.GETResults Then
                bW_ChkBarcode = True
                If oSKU.GETInStoreBar Then
                    If cDiscAmt > 0 Then
                        otbAmt.Text = SP_FMTcCurPoint(cDiscAmt, True, True)
                        otbAmt.Locked = True
                        otbAmt.BackColor = nVB_CNClrBLock
                        Call SP_OBJxSetFocus(ocmOk)
                    Else
                        otbAmt.Text = "0.00"
                        otbAmt.Locked = False
                        otbAmt.BackColor = nVB_CNClrBUnLock
                        Call SP_OBJxSetFocus(otbAmt)
                    End If
                Else
                    otbAmt.Text = "0.00"
                    otbAmt.Locked = False
                    otbAmt.BackColor = nVB_CNClrBUnLock
                    Call SP_OBJxSetFocus(otbAmt)
                End If
            Else
                Call SP_MSGnShowing("tms_CN005224", nCS_Warning)
                otbCpnNo.Text = ""
                otbAmt.Text = "0.00"
                otbAmt.Locked = False
                otbAmt.BackColor = nVB_CNClrBUnLock
                Call SP_OBJxSetFocus(otbCpnNo)
            End If
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
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub otbAmt_Change()     '*AO51-03-12
    otbAmt.Text = SP_DEFtChangeNum(Trim(otbAmt.Text))
    otbAmt.SelStart = Len(otbAmt.Text)
End Sub

Private Sub otbAmt_GotFocus()
    Call SP_OTBxSelText(otbAmt)
End Sub

Private Sub otbAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbAmt.Text = ""
        Case 13
            If otbAmt.Text <> "" Then
                SendKeys "{TAB}"
            End If
    End Select
End Sub



Private Sub otbCpnNo_GotFocus()
    Call SP_OTBxSelText(otbCpnNo)
End Sub

Private Sub otbCpnNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cDiscAmt As Double
    Select Case KeyCode
        Case vbKeyBack
            bW_ChkBarcode = False
            If Trim(otbCpnNo.Text) = "" Then
                otbAmt.Text = "0.00"
                otbAmt.Locked = False
                otbAmt.BackColor = nVB_CNClrBUnLock
            End If
        Case vbKeyDelete         'delete
            otbCpnNo.Text = ""
            bW_ChkBarcode = False
            If Trim(otbCpnNo.Text) = "" Then
                otbAmt.Text = "0.00"
                otbAmt.Locked = False
                otbAmt.BackColor = nVB_CNClrBUnLock
            End If
        Case 13
'            If Trim(otbCpnNo.Text) <> "" Then
'                SendKeys "{TAB}"
'            End If
            '*KT 54-07-26 V.4.0.1 Phase4.0(Member Point) Promotion Coupon ตรวจสอบ barcode config
            cDiscAmt = oSKU.GETPmtCpnBarcode(Trim(otbCpnNo.Text))
            If oSKU.GETResults Then
                bW_ChkBarcode = True
                If oSKU.GETInStoreBar Then
                    If cDiscAmt > 0 Then
                        otbAmt.Text = SP_FMTcCurPoint(cDiscAmt, True, True)
                        otbAmt.Locked = True
                        otbAmt.BackColor = nVB_CNClrBLock
                        Call SP_OBJxSetFocus(ocmOk)
                    Else
                        otbAmt.Text = "0.00"
                        otbAmt.Locked = False
                        otbAmt.BackColor = nVB_CNClrBUnLock
                        Call SP_OBJxSetFocus(otbAmt)
                    End If
                Else
                    otbAmt.Text = "0.00"
                    otbAmt.Locked = False
                    otbAmt.BackColor = nVB_CNClrBUnLock
                    Call SP_OBJxSetFocus(otbAmt)
                End If
            Else
                Call SP_MSGnShowing("tms_CN005224", nCS_Warning)
                otbCpnNo.Text = ""
                otbAmt.Text = "0.00"
                otbAmt.Locked = False
                otbAmt.BackColor = nVB_CNClrBUnLock
                Call SP_OBJxSetFocus(otbCpnNo)
            End If
    End Select
End Sub
Private Sub otbAmt_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbAmount.Text), ".") > 0 Then         '*AO51-03-12    Comm.Sheet SF- 070
'                KeyAscii = 0
'            End If
        Case 13
            KeyAscii = 0
            Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select

End Sub

Private Sub otbCpnNo_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
End Sub

Private Sub otbCpnNo_LostFocus()
    Dim cDiscAmt As Double
    otbCpnNo.Text = SP_TXTtTrimRightLeft(otbCpnNo.Text)
    '*KT 54-07-26 V.4.0.1 Phase4.0(Member Point) Promotion Coupon ตรวจสอบ barcode config
    If Not bW_ChkBarcode Then
        If otbCpnNo.Text <> "" Then
            cDiscAmt = oSKU.GETPmtCpnBarcode(Trim(otbCpnNo.Text))
            If oSKU.GETResults Then
                bW_ChkBarcode = True
                If oSKU.GETInStoreBar Then
                    If cDiscAmt > 0 Then
                        otbAmt.Text = SP_FMTcCurPoint(cDiscAmt, True, True)
                        otbAmt.Locked = True
                        otbAmt.BackColor = nVB_CNClrBLock
                        Call SP_OBJxSetFocus(ocmOk)
                    Else
                        otbAmt.Text = "0.00"
                        otbAmt.Locked = False
                        otbAmt.BackColor = nVB_CNClrBUnLock
                    End If
                Else
                    otbAmt.Text = "0.00"
                    otbAmt.Locked = False
                    otbAmt.BackColor = nVB_CNClrBUnLock
                End If
            Else
                Call SP_MSGnShowing("tms_CN005224", nCS_Warning)
                otbCpnNo.Text = ""
                otbAmt.Text = "0.00"
                otbAmt.Locked = False
                otbAmt.BackColor = nVB_CNClrBUnLock
                Call SP_OBJxSetFocus(otbCpnNo)
            End If
        End If
    End If
End Sub
