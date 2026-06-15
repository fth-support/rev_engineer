VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNCheckPrice 
   Caption         =   "Price Check;klg_Title"
   ClientHeight    =   3120
   ClientLeft      =   5370
   ClientTop       =   3810
   ClientWidth     =   6510
   ControlBox      =   0   'False
   Icon            =   "wCNCheckPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6510
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   -240
      Width           =   6495
      _Version        =   720897
      _ExtentX        =   11456
      _ExtentY        =   4260
      _StockProps     =   64
      Appearance      =   10
      Color           =   32
      Begin VB.TextBox otbForm 
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
         Top             =   420
         Width           =   3975
      End
      Begin VB.Label olbRegPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2280
         TabIndex        =   4
         Tag             =   "SubTotal;ยอดคงเหลือ"
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label olbItemName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   20.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. Price :;klg_Regprice"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1740
         Width           =   2055
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name :;klg_Itemname"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   1110
         Width           =   2055
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode :;klg_Barcode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   480
         Width           =   1815
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl2 
      Height          =   1215
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   6495
      _Version        =   720897
      _ExtentX        =   11456
      _ExtentY        =   2143
      _StockProps     =   64
      Appearance      =   10
      Color           =   32
      Begin SCANNERLib.Scanner oScanner 
         Left            =   240
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
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNCheckPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Form_Activate()
'On Error GoTo ErrHandle
'        Call SP_OBJxSetFocus(otbForm)
'    Exit Sub
'ErrHandle:
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
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
            '*Bump 56-02-28 เพิ่ม Code 3 บรรทัด Touch Scanner
            If .DeviceName = tSNR_TOUCH_SCANNER Then
                .DecodeData = True
            End If
        End With
    End If
    Exit Sub
ErrHandle:
End Sub

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
    '    Set wCNCheckPrice = Nothing
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    Unload Me
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
    With oScanner
        'otbForm.Text = .ScanData'*Bump 56-02-28 Comment บรรทัดนี้
        '*Bump 56-02-28 เพิ่ม Code 8 บรรทัด
         If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
            .DecodeData = True
        End If
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            otbForm.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel) '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
        Else
            otbForm.Text = .ScanData
        End If
'        DoEvents            '*KT 53-01-14
        If otbForm.Text <> "" Then
            '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
            '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
            If Asc(Right(otbForm.Text, 1)) = 13 Then
                otbForm.Text = Mid(otbForm.Text, 1, Len(otbForm.Text) - 1)
            End If
            Call otbForm_KeyDown(13, 0)
        End If
        'ทำการ Open Scanner อีกครั้ง ไม่นั้นครั้งต่อไปจะ Scan สินค้าไม่ได้ครับ
        '*KT 52-01-08 ไม่ต้อง Open Scanner อีกครั้ง ยกเว้นกรณี Form Active
'        .Open tVB_Scanner
'        .Claim 0 'TeamPos2000
'        'Enable the device.
'        .DeviceEnabled = True
'        'Enable the event.
'        .DataEventEnabled = True
'        .AutoDisable = True
'        .ClearInput
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

Private Sub otbForm_GotFocus()
    Call SP_OTBxSelText(otbForm)
End Sub

Private Sub otbForm_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandle
    Select Case KeyCode
        Case 13, 18, 27, 33, 34, 35, 36, 45, 112 To 123
            Select Case Shift
                Case 0
                        Select Case KeyCode
                            Case 13
                                Call W_FINxPriceCheck(Trim(otbForm.Text))
                        End Select
             End Select
        Case vbKeyDelete
            otbForm.Text = ""
            olbItemName.Caption = ""
            olbRegPrice.Caption = ""
        Case Else
            KeyCode = 0
    End Select
    Exit Sub
ErrHandle:
    If Err.Number = 438 Then            'not define function name in the database
        Call SP_MSGnShowing("tms_CN005107", nCS_Information)
    Else
        Call SP_MSGxShowError(Err.Number, Err.Description)
    End If
End Sub

Private Sub W_FINxPriceCheck(ByVal ptPLU$)
'------------------------------------------------------------------------
'Call:
'Cmt: ตรวจสอบราคาของสินค้าที่ส่งมา
'*KT 51-10-03 3.0.128/3.0.200 Communication Sheet ML-SF-572 Price Check is incorrect.xls
'------------------------------------------------------------------------
    Dim tSql$
    Dim orsSKU As ADODB.Recordset
On Error GoTo ErrHandle
    ptPLU = oSKU.GETSKUBarcode(ptPLU)
    tSql = " STP_CNxGetSKU @ptPLU = '" & ptPLU & "'"
    If SP_SQLvExecute(tSql, orsSKU, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
        If Not (orsSKU.BOF Or orsSKU.EOF) Then
             If oSKU.GETInStoreBar Then
                olbItemName.Caption = SP_FEDtChkString(orsSKU, IIf(nVB_CNCutLang = 1, "FTSkuAbbEng", "FTSkuAbbThai"))
                olbRegPrice.Caption = SP_FMTcCurPoint(oDisc.C_GETcChkDiscPmtTime(SP_FEDtChkString(orsSKU, "FTSkuCode")), False, True) '*KT 51-10-06  V 3.0.129
                Call SP_OBJxSetFocus(otbForm)
             Else
                olbItemName.Caption = SP_FEDtChkString(orsSKU, IIf(nVB_CNCutLang = 1, "FTSkuAbbEng", "FTSkuAbbThai"))
                olbRegPrice.Caption = SP_FMTcCurPoint(oDisc.C_GETcChkDiscPmtTime(SP_FEDtChkString(orsSKU, "FTSkuCode")), False, True) '*KT 51-10-06  V 3.0.129
                Call SP_OBJxSetFocus(otbForm)
            End If
        Else
            olbItemName.Caption = ""
            olbRegPrice.Caption = ""
            Call SP_MSGnShowing("tms_CN006003", tCS_CNMsgErr)
            Call SP_OBJxSetFocus(otbForm)
            Call otbForm_GotFocus
        End If
    Else
        olbItemName.Caption = ""
        olbRegPrice.Caption = ""
        Call SP_MSGnShowing("tms_CN006003", tCS_CNMsgErr)
        Call SP_OBJxSetFocus(otbForm)
        Call otbForm_GotFocus
    End If
    Call SP_DATxRsNothing(orsSKU)
    Exit Sub
ErrHandle:
    olbItemName.Caption = ""
    olbRegPrice.Caption = ""
    Call SP_MSGnShowing("tms_CN006003", tCS_CNMsgErr)
    Call SP_OBJxSetFocus(otbForm)
    Call otbForm_GotFocus
    Call SP_DATxRsNothing(orsSKU)
End Sub


Private Sub otbForm_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
End Sub
