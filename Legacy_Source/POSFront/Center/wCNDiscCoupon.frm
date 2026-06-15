VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNDiscCoupon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cooperate Coupon"
   ClientHeight    =   3315
   ClientLeft      =   6585
   ClientTop       =   4575
   ClientWidth     =   6645
   Icon            =   "wCNDiscCoupon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6645
   Begin VB.Frame ofmForm 
      Height          =   975
      Left            =   1200
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox otbDate 
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox otbDate 
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   960
         MaxLength       =   2
         TabIndex        =   13
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox otbDate 
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   240
         MaxLength       =   2
         TabIndex        =   12
         Top             =   240
         Width           =   600
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DDMMYYYY"
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
         Index           =   4
         Left            =   2880
         TabIndex        =   15
         Top             =   360
         Width           =   1500
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2655
      Index           =   2
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   -255
      Width           =   6645
      _Version        =   720897
      _ExtentX        =   11721
      _ExtentY        =   4683
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbPer 
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
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "0"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox otbCouponNo 
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
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker odtDate 
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58064899
         CurrentDate     =   39528
      End
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(MM/DD/YYYY)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   5
         Left            =   2640
         TabIndex        =   16
         Top             =   2200
         Width           =   1860
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   3
         Left            =   3900
         TabIndex        =   10
         Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
         Top             =   1155
         Width           =   315
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate:;klg_Rate"
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
         TabIndex        =   9
         Tag             =   "2;"
         Top             =   1155
         Width           =   1860
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
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   555
         Width           =   1995
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration Date:;klg_ExpDate"
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
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1680
         Width           =   2010
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2100
      Width           =   6645
      _Version        =   720897
      _ExtentX        =   11721
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
         Left            =   5040
         TabIndex        =   4
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
         Left            =   3480
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNDiscCoupon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bW_KeyEnter As Boolean
Private bW_ChkBarcode As Boolean '*KT 54-06-21 V.4.0.0 Ph.4.0 Req.14  คูปองส่วนลด เพิ่มการ Config Barcode
'
'Private Function W_DEFbChkDate() As Boolean         '*AO51-03-20    Comm.Sheet ML-SF 109
'    Dim tYear As String, tMonth As String, tDay As String
'    Const nIndex = 0
'
'    On Error GoTo ErrHandle
'
'    W_DEFbChkDate = False
'    tDay = Trim(otbDate(nIndex).Text)
'    tMonth = Trim(otbDate(nIndex + 1).Text)
'    tYear = Trim(otbDate(nIndex + 2).Text)
'
'    If Len(tDay) < 2 Or Val(tDay) > 31 Or tDay = "00" Then    'DD
'        W_DEFbChkDate = False
'        otbDate(nIndex).Text = ""
'        otbDate(nIndex).SetFocus
'        Exit Function
'    Else
'        If Len(tMonth) < 2 Or Val(tMonth) > 12 Or tMonth = "00" Then          'MM
'            W_DEFbChkDate = False
'            otbDate(nIndex + 1).Text = ""
'            otbDate(nIndex + 1).SetFocus
'            Exit Function
'        Else
'            If Len(tYear) < 4 Or tYear = "0000" Then         'YYYY
'                 W_DEFbChkDate = False
'                 otbDate(nIndex + 2).Text = ""
'                 otbDate(nIndex + 2).SetFocus
'                 Exit Function
'             End If
'        End If
'    End If
'
'    If Val(tDay) < 32 Then
'        Select Case tMonth
'            Case "02"
'            'CDate(tDay & "/" & tMonth & "/" & tYear)
'                If SP_DTEbIsLeapYear(CDate("01/01/" & tYear)) Then
'                    If Val(tDay) > 29 Then
'                        W_DEFbChkDate = False
'                        otbDate(nIndex).Text = ""
'                        otbDate(nIndex).SetFocus
'                        Exit Function
'                    Else
'                        W_DEFbChkDate = True
'                    End If
'                Else
'                    If Val(tDay) > 28 Then
'                        W_DEFbChkDate = False
'                        otbDate(nIndex).Text = ""
'                        otbDate(nIndex).SetFocus
'                    Else
'                        W_DEFbChkDate = True
'                    End If
'                End If
'
'            Case "04", "06", "09", "11"
'                 If Val(tDay) < 31 Then
'                    W_DEFbChkDate = True
'                Else
'                    W_DEFbChkDate = False
'                    otbDate(nIndex).Text = ""
'                    otbDate(nIndex).SetFocus
'                 End If
'            Case "01", "03", "05", "07", "08", "10", "12"
'                If Val(tDay) < 32 Then
'                    W_DEFbChkDate = True
'                Else
'                    W_DEFbChkDate = False
'                    otbDate(nIndex).Text = ""
'                    otbDate(nIndex).SetFocus
'                End If
'        End Select
'    Else
'
'        W_DEFbChkDate = False
'        otbDate(nIndex).Text = ""
'        otbDate(nIndex).SetFocus
'        Exit Function
'    End If
'    Exit Function
'ErrHandle:
'    Call SP_MSGxShowError(Err.Number, Err.Description)
'    W_DEFbChkDate = False
'End Function
Private Sub W_FRMxSetInitial()
'-------------------------------------------------------------
'Call:
'Cmt: set ค่าเบื้องสำหรับการเริ่มต้นการทำงานขิงหน้าจอนี้
'--------------------------------------------------------------

    If bVB_ConfigDate Then          '*AO51-03-20
        If tVB_SelConfigDate = "0" Then
            odtDate.CustomFormat = tCS_FormatDateM
            odtDate.Value = Format(Date, tCS_FormatDateM)
            olaForm(4).Caption = "(" & UCase(tCS_FormatDateM) & ")"
        Else
            odtDate.CustomFormat = tCS_FormatDateSec
            odtDate.Value = Format(Date, tCS_FormatDateSec)
            olaForm(4).Caption = "(" & UCase(tCS_FormatDateSec) & ")"
        End If
    Else
        odtDate.CustomFormat = tCS_FormatDateSec
        odtDate.Value = Format(Date, tCS_FormatDateSec)
       olaForm(4).Caption = "(" & UCase(tCS_FormatDateSec) & ")"
    End If
    otbCouponNo.Text = ""
    otbPer.Text = "0"
    
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)             '*Ao 53-01-14
'    If KeyCode = 13 Then SendKeys "{Tab}"
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    bW_ChkBarcode = False
    bVB_DiscAccept = False
    bW_KeyEnter = False
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
            '*Bump 56-02-28 เพิ่ม code 3 บรรทัด
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
'    Set wCNDiscCoupon = Nothing
'End Sub

Private Sub ocmCancel_Click()
    bVB_DiscAccept = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim tRateCpn As String
'    If Not (W_DEFbChkDate) Then Exit Sub            '*AO51-03-20    Comm.Sheet ML-SF 109
'    dVB_CNDateAll = DateSerial(Val(otbDate(2).Text), Val(otbDate(1).Text), Val(otbDate(0).Text))
     '*KT 54-06-21 V.4.0.0 Ph.4.0 Req.14  คูปองส่วนลด เพิ่มการ Config Barcode
    If Not bW_ChkBarcode Then
         tRateCpn = oSKU.GETRateDiscountCpnBarcode(Trim(otbCouponNo.Text))
         bW_ChkBarcode = True
         otbPer.Text = CStr(Val(Format(tRateCpn)))
         If CStr(Val(Format(tRateCpn))) > 0 And CStr(Val(Format(tRateCpn))) <= 100 Then
            otbPer.Locked = True
            otbPer.BackColor = nVB_CNClrBLock
        Else
            otbPer.Locked = False
            otbPer.BackColor = nVB_CNClrBUnLock
         End If
    End If
    dVB_CNDateAll = odtDate.Value
    If Format(dVB_CNDateAll, "dd/mm/yyyy") < Format(SP_GETdChkDateDisc, " dd/mm/yyyy") Then '*KT 52-06-04 ตรวจสอบวันที่หมดอายุที่กำหนดใน parameter
        Call SP_MSGnShowing("tms_CN005147", tCS_CNMsgWarn)
        Exit Sub
    End If
    Select Case tVB_CNTrnDiscType
        Case "1811" 'Cooperate Coupon
            oDisc.bRecalByItem = False
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.cDiscValue = SP_VALcText2Double(otbPer.Text)
            oDisc.nENLnDisType = nEN_LnDis1Per
            oDisc.dCardExpDate = dVB_CNDateAll
            oDisc.tCardID = Trim(otbCouponNo.Text)
            Call oDisc.DIScDiscount(nEN_Disc09CoopCpn, eEN_VerifyDiscType.nEN_Type2ByDcn)
            bVB_DiscAccept = True
        Case "1812" 'Cooperate Brand
            oDisc.bRecalByItem = False
            oDisc.DISxClearAll
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            oDisc.cDiscValue = SP_VALcText2Double(otbPer.Text)
            oDisc.nENLnDisType = nEN_LnDis1Per
            oDisc.dCardExpDate = dVB_CNDateAll
            oDisc.tCardID = Trim(otbCouponNo.Text)
            Call oDisc.DIScDiscount(nEN_Disc10CoopBrand, eEN_VerifyDiscType.nEN_Type2ByDcn)
            bVB_DiscAccept = True
    End Select
    Unload Me
End Sub

Private Sub ocmOk_GotFocus()
    Dim tRateCpn As String
    '*KT 54-06-21 V.4.0.0 Ph.4.0 Req.14  คูปองส่วนลด เพิ่มการ Config Barcode
    If Not bW_ChkBarcode Then
         tRateCpn = oSKU.GETRateDiscountCpnBarcode(Trim(otbCouponNo.Text))
         bW_ChkBarcode = True
         otbPer.Text = CStr(Val(Format(tRateCpn)))
         If CStr(Val(Format(tRateCpn))) > 0 And CStr(Val(Format(tRateCpn))) <= 100 Then
            otbPer.Locked = True
            otbPer.BackColor = nVB_CNClrBLock
        Else
            otbPer.Locked = False
            otbPer.BackColor = nVB_CNClrBUnLock
         End If
    End If
End Sub

Private Sub odtDate_GotFocus()
    Dim tRateCpn As String
    '*KT 54-06-21 V.4.0.0 Ph.4.0 Req.14  คูปองส่วนลด เพิ่มการ Config Barcode
    If Not bW_ChkBarcode Then
         tRateCpn = oSKU.GETRateDiscountCpnBarcode(Trim(otbCouponNo.Text))
         bW_ChkBarcode = True
         otbPer.Text = CStr(Val(Format(tRateCpn)))
         If CStr(Val(Format(tRateCpn))) > 0 And CStr(Val(Format(tRateCpn))) <= 100 Then
            otbPer.Locked = True
            otbPer.BackColor = nVB_CNClrBLock
        Else
            otbPer.Locked = False
            otbPer.BackColor = nVB_CNClrBUnLock
         End If
    End If
End Sub

Private Sub odtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            If IsDate(odtDate.Value) Then
                SendKeys "{Tab}"
            End If
    End Select
End Sub



Private Sub oScanner_DataEvent(ByVal Status As Long)
    Dim tRateCpn As String
    
On Error GoTo ErrHandle
    With oScanner
        'otbCouponNo.Text = .ScanData '*Bump 56-02-28 Comment บรรทัดนี้
        '*Bump 56-02-28 เพิ่ม code 8 บรรทัด
        If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
            .DecodeData = True
        End If
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            otbCouponNo.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel)  '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
        Else
            otbCouponNo.Text = .ScanData
        End If
'        DoEvents            '*KT 53-01-14
        If otbCouponNo.Text <> "" Then
            '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
            '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
            If Asc(Right(otbCouponNo.Text, 1)) = 13 Then
                otbCouponNo.Text = Mid(otbCouponNo.Text, 1, Len(otbCouponNo.Text) - 1)
            End If
            '*KT 54-06-21 V.4.0.0 Ph.4.0 Req.14  คูปองส่วนลด เพิ่มการ Config Barcode
             tRateCpn = oSKU.GETRateDiscountCpnBarcode(Trim(otbCouponNo.Text))
             bW_ChkBarcode = True
             otbPer.Text = CStr(Val(Format(tRateCpn)))
            If CStr(Val(Format(tRateCpn))) > 0 And CStr(Val(Format(tRateCpn))) <= 100 Then
                otbPer.Locked = True
                otbPer.BackColor = nVB_CNClrBLock
            Else
                otbPer.Locked = False
                otbPer.BackColor = nVB_CNClrBUnLock
             End If
            If otbPer.Locked Then
                    Call otbPer_LostFocus
                    Call SP_OBJxSetFocus(odtDate)
            Else
                    Call SP_OBJxSetFocus(otbPer)
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


'Private Sub odtForm_KeyDown(KeyCode As Integer, Shift As Integer)
'     If KeyCode = 13 Then ocmOK_Click
'End Sub
Private Sub otbCouponNo_GotFocus()
    otbCouponNo.SelStart = 0
    otbCouponNo.SelLength = Len(otbCouponNo.Text)
End Sub

Private Sub otbCouponNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim tRateCpn As String
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbCouponNo.Text = ""
            bW_ChkBarcode = False
        Case 13
            '*KT 54-06-21 V.4.0.0 Ph.4.0 Req.14  คูปองส่วนลด เพิ่มการ Config Barcode
             tRateCpn = oSKU.GETRateDiscountCpnBarcode(Trim(otbCouponNo.Text))
             bW_ChkBarcode = True
             otbPer.Text = CStr(Val(Format(tRateCpn)))
             If CStr(Val(Format(tRateCpn))) > 0 And CStr(Val(Format(tRateCpn))) <= 100 Then
                otbPer.Locked = True
                otbPer.BackColor = nVB_CNClrBLock
            Else
                otbPer.Locked = False
                otbPer.BackColor = nVB_CNClrBUnLock
             End If
            If Trim(otbCouponNo.Text) <> "" Then
                If otbPer.Locked Then
                    Call otbPer_LostFocus
                    Call SP_OBJxSetFocus(odtDate)
                Else
                    SendKeys "{Tab}"
                End If
            End If
        Case Else
            bW_ChkBarcode = False
    End Select
End Sub

Private Sub otbCouponNo_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
End Sub

Private Sub otbCouponNo_LostFocus()
    Dim tRateCpn As String
    otbCouponNo.Text = SP_TXTtTrimRightLeft(otbCouponNo.Text)
'    If Not bW_ChkBarcode Then
'        '*KT 54-06-21 V.4.0.0 Ph.4.0 Req.14  คูปองส่วนลด เพิ่มการ Config Barcode
'         tRateCpn = oSKU.GETRateDiscountCpnBarcode(Trim(otbCouponNo.Text))
'         bW_ChkBarcode = True
'         otbPer.Text = CStr(Val(Format(tRateCpn)))
'         If CStr(Val(Format(tRateCpn))) > 0 And CStr(Val(Format(tRateCpn))) <= 100 Then
'            otbPer.Locked = True
'            otbPer.BackColor = nVB_CNClrBLock
'        Else
'            otbPer.Locked = False
'            otbPer.BackColor = nVB_CNClrBUnLock
'         End If
'    End If
End Sub

Private Sub otbDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbDate(Index).Text = ""
    End Select
End Sub

Private Sub otbPer_Change()
    If Val(Format(otbPer.Text)) > 100 Then
        Call SP_MSGnShowing("tms_CN005223", nCS_Warning)
        otbPer.Text = "0"
        Call SP_OTBxSelText(otbPer)
    End If
End Sub

Private Sub otbPer_GotFocus()
    Dim tRateCpn As String
    '*KT 54-06-21 V.4.0.0 Ph.4.0 Req.14  คูปองส่วนลด เพิ่มการ Config Barcode
    If Not bW_ChkBarcode Then
         tRateCpn = oSKU.GETRateDiscountCpnBarcode(Trim(otbCouponNo.Text))
         bW_ChkBarcode = True
         otbPer.Text = CStr(Val(Format(tRateCpn)))
         If CStr(Val(Format(tRateCpn))) > 0 And CStr(Val(Format(tRateCpn))) <= 100 Then
            otbPer.Locked = True
            otbPer.BackColor = nVB_CNClrBLock
        Else
            otbPer.Locked = False
            otbPer.BackColor = nVB_CNClrBUnLock
         End If
    End If
    otbPer.SelStart = 0
    otbPer.SelLength = Len(otbPer.Text)
End Sub

Private Sub otbPer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbPer.Text = ""
        Case 13
            If Trim(otbPer.Text) <> "" Then
                SendKeys "{Tab}"
            End If
    End Select
            
End Sub

Private Sub otbPer_LostFocus()
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
On Error GoTo Err_Handler
    Select Case tVB_CNTrnDiscType
        Case "1811" 'Cooperate Coupon
            tSql = "SELECT FTCpcStaExpDate,FDCpcExpDate,FTCpcStaLimitAmt,FCCpcLimitAmt"
            tSql = tSql & vbCrLf & " FROM TCNMCoopCpMtn "
            tSql = tSql & vbCrLf & " WHERE FCCpcRate = " & CDbl(Val(Format(Trim(otbPer.Text))))
            If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master '*KT 54-04-11 V.3.0.56  ส่งพารามิเตอร์ไม่ถูกต้อง
                If orsTemp.BOF And orsTemp.EOF Then
                       Call SP_MSGnShowing("tms_CN005130", nCS_Warning)
                End If
            Else
                Call SP_MSGnShowing("tms_CN005130", nCS_Warning)
            End If
            Call SP_DATxRsNothing(orsTemp)
        Case "1812" 'Cooperate Brand
            tSql = "SELECT FTCpbStaExpDate,FDCpbExpDate,FTCpbStaLimitAmt,FCCpbLimitAmt"
            tSql = tSql & vbCrLf & " FROM TCNMCoopBrnMtn "
            tSql = tSql & vbCrLf & " WHERE FCCpbRate = " & CDbl(Val(Format(Trim(otbPer.Text))))
            If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master '*KT 54-04-11 V.3.0.56  ส่งพารามิเตอร์ไม่ถูกต้อง
                If orsTemp.BOF And orsTemp.EOF Then
                       Call SP_MSGnShowing("tms_CN005131", nCS_Warning)
                End If
            Else
                Call SP_MSGnShowing("tms_CN005130", nCS_Warning)
            End If
            Call SP_DATxRsNothing(orsTemp)
    End Select
    Exit Sub
Err_Handler:
    Call SP_DATxRsNothing(orsTemp)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub
