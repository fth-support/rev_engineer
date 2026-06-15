VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNPayNewVendor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Voucher;klg_Title"
   ClientHeight    =   4140
   ClientLeft      =   2265
   ClientTop       =   3615
   ClientWidth     =   6645
   ControlBox      =   0   'False
   Icon            =   "wCNPayNewVendor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6645
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3495
      Index           =   2
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   -240
      Width           =   6645
      _Version        =   720897
      _ExtentX        =   11721
      _ExtentY        =   6165
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbQty 
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
         Left            =   4080
         MaxLength       =   11
         TabIndex        =   2
         Text            =   "1"
         Top             =   2280
         Width           =   2415
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
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   1680
         Width           =   4095
      End
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
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "01234567890123456789"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label olaTotalamount 
         Alignment       =   1  'Right Justify
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
         Left            =   2445
         TabIndex        =   14
         Top             =   2880
         Width           =   4050
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount :;klg_TotalA"
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
         Left            =   120
         TabIndex        =   13
         Tag             =   "2;"
         Top             =   2895
         Width           =   3270
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Vendor Voucher :;klg_TotalV"
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
         Left            =   120
         TabIndex        =   12
         Tag             =   "2;"
         Top             =   495
         Width           =   4320
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
         Left            =   120
         TabIndex        =   10
         Tag             =   "2;"
         Top             =   1755
         Width           =   2280
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Vch No :;klg_VendorCpnNo"
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
         TabIndex        =   9
         Tag             =   "2;"
         Top             =   1155
         Width           =   4440
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty :;klg_Qty"
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
         Left            =   120
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   2355
         Width           =   1650
      End
      Begin VB.Label olaTotalVendor 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   3090
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   2950
      Width           =   6645
      _Version        =   720897
      _ExtentX        =   11721
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
      Begin VB.CommandButton ocmChange 
         Caption         =   "Change Cp.(F5);klg_Change"
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
         Left            =   1080
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   480
         Width           =   2175
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
         Left            =   3360
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   480
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
         Left            =   4920
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNPayNewVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cW_TotalVendorCpn As Double
Private nW_SeqNo As Integer
Private nW_SkuIndex As Integer

Private Sub W_UPDxVendorCpn()
Dim cAmt As Double, cQty As Double
Dim tSql  As String
    On Error GoTo ErrHandle
    cAmt = SP_FMTcCurPoint(Trim(olaTotalamount.Caption), True, False)
    cQty = SP_FMTcCurPoint(Trim(otbQty.Text), True, False)
    
    tSql = "UPDATE " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "SET FTSrcCardNo = '" & Trim(otbNo.Text) & "'"
    tSql = tSql & vbCrLf & ",FCSrcFAmt = " & cAmt
    tSql = tSql & vbCrLf & ",FCSrcAmt = " & cAmt
    tSql = tSql & vbCrLf & ",FCSrcNet = " & cAmt
    tSql = tSql & vbCrLf & ",FCSrcChg = 0"
    tSql = tSql & vbCrLf & ",FTVendorCnpQty = '" & cQty & "'"
    tSql = tSql & vbCrLf & ",FTVendorCnpType = '" & tVB_VendorCnpType & "'"
    tSql = tSql & vbCrLf & ",FTSrcRmk='Change Vendor'"
    
    tSql = tSql & vbCrLf & "WHERE FNSrcSeqNo=" & nW_SeqNo
    tSql = tSql & vbCrLf & "AND FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
    Call SP_SQLvExecute(tSql)
    Exit Sub
ErrHandle:
    
End Sub

Private Sub W_CHKxControlTextbox(ByVal pControl As TextBox, ByVal pChk As Boolean)
    If pChk Then
        pControl.BackColor = nVB_CNClrBUnLock
        pControl.Enabled = True
    Else
        pControl.BackColor = nVB_CNClrBLock
        pControl.Enabled = False
    End If
End Sub

Private Function W_FNDcSumItemVendor(ByVal pnSeqNo As Integer, Optional ByVal pnSrcSeqNo As Integer = 0) As Double
'___________________________________________________
'Call:
'Cmt:
'___________________________________________________
Dim tWhr As String

    On Error GoTo ErrHandle
    
    tWhr = " WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhr = tWhr & "  AND FTShdTransNo='" & tVB_Running & "'"
    tWhr = tWhr & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tWhr = tWhr & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tWhr = tWhr & "  AND FTTdmCode = 'T006'"
    tWhr = tWhr & "  AND ISNULL(FTSrcStaNoPay,'') <> '1'"          '*Ao 2011/01/17   CommSheet ML-PRV-020-Change Tender is wrong concept.xls
    If bVB_CheckVendor Then
        tWhr = tWhr & "  AND FTVendorCnpFlag = '1'"
        tWhr = tWhr & "  AND FTSkuIndex = '" & pnSeqNo & "'"
    Else
'        tWhr = tWhr & "  AND FTVendorCnpFlag = '1'"
    End If
    If bVB_VedFrom = False Then
         tWhr = tWhr & "  AND NOT (FNSrcSeqNo = '" & pnSrcSeqNo & "')"
    End If
    
    W_FNDcSumItemVendor = SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSrcNet", tVB_CNTblRC, tWhr))
 Exit Function
ErrHandle:
    W_FNDcSumItemVendor = 0
End Function

Private Sub W_FNDxCheckCfgBarCode(ByVal porsTemp As ADODB.Recordset)
'___________________________________________________
'Call:
'Cmt:
'___________________________________________________
Dim nTmp As Integer
Dim nMatchStr As Integer, nItemSize As Integer, nCouponNo As Integer, nPosNo As Integer, nPrice As Integer, nPosPrice As Integer, nPriceDec As Integer
Dim tNo As String, tCouponNo As String, tPrice As String

    tNo = Trim(otbNo.Text)
'    nMatchStr = SP_FEDvChkNumeric(porsTemp, "FNBafMatchStr")
'    nItemSize = SP_FEDvChkNumeric(porsTemp, "FNBafItemSize")
'    nCouponNo = SP_FEDvChkNumeric(porsTemp, "FNBafPLUSize")
'    nPosNo = SP_FEDvChkNumeric(porsTemp, "FNBafPLULocate")
'    nPrice = SP_FEDvChkNumeric(porsTemp, "FNBafPriceSize")
'    nPosPrice = SP_FEDvChkNumeric(porsTemp, "FNBafPriceLocate")
'    nPriceDec = SP_FEDvChkNumeric(porsTemp, "FNBafPriceDec")
'
'    tCouponNo = Mid(tNo, nPosNo, nCouponNo)
'    tPrice = Mid(tNo, nPosPrice, nPrice - nPriceDec)
    Call SP_FNDxCheckCfgBarCode(porsTemp, tNo, tCouponNo, tPrice)
    otbNum.Text = SP_FMTtNum2Shw(tPrice)
    otbNo.Text = tCouponNo
    otbQty.Text = "1"
    olaTotalamount.Caption = SP_FMTcCurPoint(SP_VALcText2Double(otbNum.Text) * SP_VALcText2Double(otbQty.Text), False)
End Sub

Private Sub W_FNDxCheckBarCode()
'___________________________________________________
'Call:
'Cmt:
'___________________________________________________
Dim orsTemp As ADODB.Recordset
Dim tSql As String
   
   On Error GoTo ErrHandle
   
    tSql = "SELECT * FROM TCNMSkuBarCon"
    tSql = tSql & " WHERE FNBafItemSize = '" & Len(Trim(otbNo.Text)) & "'"
    tSql = tSql & " AND FNBafMatchStr = '" & Left(Trim(otbNo.Text), 1) & "'"
    tSql = tSql & " AND FTBafCodeType = '2'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
    If Not (orsTemp.EOF And orsTemp.BOF) Then
        Call W_FNDxCheckCfgBarCode(orsTemp)
        Call W_CHKxControlTextbox(otbNum, False)
        Call SP_OBJxSetFocus(otbQty)
        tVB_VendorCnpType = "1"
    Else
        tSql = "SELECT * FROM TCNMVndCoupon"
        tSql = tSql & " WHERE FTVdcCode = '" & Trim(otbNo.Text) & "'"
        Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
        If Not (orsTemp.EOF And orsTemp.BOF) Then
'            otbNo.Text = SP_FEDtChkString(orsTemp, "FTVdcCode")
            otbNum.Text = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCVdcCpnAmt"), False)
'            otbQty.Text = "1"
            If SP_VALcText2Double(otbQty.Text) = 0 Then
                otbQty.Text = "1"
            End If
            olaTotalamount.Caption = SP_FMTcCurPoint(SP_VALcText2Double(otbNum.Text) * SP_VALcText2Double(otbQty.Text), False)
            Call W_CHKxControlTextbox(otbNum, False)
            Call SP_OBJxSetFocus(otbQty)
            tVB_VendorCnpType = 2
        Else
            Call W_CHKxControlTextbox(otbNum, True)
            Call SP_OBJxSetFocus(otbNum)
            tVB_VendorCnpType = 3
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    tVB_VendorCnpType = ""
End Sub

Private Sub W_FRMxSetInitial()
'___________________________________________________
'Call:
'Cmt:
'___________________________________________________
    
    '*Tao 62-11-11 ปิดปุ่ม Change Cp. (นอกจากแก้ไขไม่ได้แล้ว ยังยกเลิกไม่ได้ด้วยนะ...เพื่อ)
    ocmChange.Visible = False 'TTK-ISS3000-750 removal of option “Change Cp.” on vendor voucher main payment screen
    
    olaTotalVendor.Caption = SP_FMTtNum2Shw(0)
    otbNo.Text = ""
    otbNum.Text = SP_FMTtNum2Shw(0)
    otbQty.Text = 1
    olaTotalamount.Caption = SP_FMTtNum2Shw(0)
    If bVB_CheckVendor Then
        ocmChange.Visible = False
    Else
        '*Ao 2010/12/08
    End If
    If bVB_VedFrom = False Then
        otbNo.Text = tVB_CardNo
        otbNum.Text = SP_FMTtNum2Shw(cVB_Amt)
        otbQty.Text = nVB_CnpQty
        olaTotalamount.Caption = SP_FMTtNum2Shw(cVB_SrcAmt)
        nW_SeqNo = nVB_SeqNo                'FNSrcSeqNo
'        'tVB_VendorCnpType
        nW_SkuIndex = Val(tVB_SkuIndex)
        Select Case tVB_VendorCnpType
            Case "1", "2"
                Call W_CHKxControlTextbox(otbNum, False)
            Case Else
                Call W_CHKxControlTextbox(otbNum, True)
        End Select
        If bVB_CheckVendor = False Then
            ocmChange.Visible = False
        End If
    End If
    Call SP_OBJxSetFocus(otbNo)
End Sub

Private Sub W_SUMcTotalVendorCpn()
'___________________________________________________
'Call:
'Cmt:
'___________________________________________________
Dim tWhe As String
   
   On Error GoTo ErrHandle
    tWhe = " WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & "  AND FTShdTransNo='" & tVB_Running & "'"
    tWhe = tWhe & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tWhe = tWhe & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tWhe = tWhe & "  AND FTTdmCode = 'T006'"
    tWhe = tWhe & "  AND ISNULL(FTSrcStaNoPay,'') <> '1'"       '*Ao 2011/01/17   CommSheet ML-PRV-020-Change Tender is wrong concept.xls
    cW_TotalVendorCpn = SP_FMTcCurPoint(SP_VALcText2Double(SP_SQLtFunction("SUM", "FCSrcNet", tVB_CNTblRC, tWhe)), False)
 Exit Sub
ErrHandle:
    cW_TotalVendorCpn = SP_FMTtNum2Shw(0)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tSql As String, tTmp As String
Dim orsTmp As ADODB.Recordset
    
    If bVB_CheckVendor = False Then
        tSql = "SELECT FTSkbRef FROM TSysKeyboard   WHERE FNSkbKey='" & CStr(KeyCode) & "'"
        tSql = tSql & " AND FNSkbShift = '" & CStr(Shift) & "'"
       ' tSql = tSql & " AND FTSkbwMain =  '1'" '*KT 54-04-08 V.3.0.56
        tSql = tSql & " AND FTSkbUsage ='Y'"
        Call SP_SQLvExecute(tSql, orsTmp, oDB.ocnOnLine, , , , , True, True) 'read master
        If Not (orsTmp.BOF And orsTmp.EOF) Then
            tTmp = SP_FEDtChkString(orsTmp, "FTSkbRef")
            If UCase(tTmp) = UCase("W_KBCancelVendor") Then
                If ocmChange.Visible = True Then
                    Call ocmChange_Click
                End If
                
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    If bVB_Res640X480 Then
'        Me.Left = 0
        Me.Top = 2000
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
    Call W_SUMcTotalVendorCpn
    olaTotalVendor.Caption = SP_FMTcCurPoint(cW_TotalVendorCpn, False)
    cVB_Pay = 0 '*Tao 58-12-14 เคลียร์ค่าการชำระเงิน
    tVB_VendorCpnNo = ""
    cVB_VendorCpnAmt = 0
    cVB_TotalDiscVendorCpnAmt = 0
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub ocmCancel_Click()
    tVB_VendorCpnNo = ""
    cVB_VendorCpnAmt = 0
    cVB_TotalDiscVendorCpnAmt = 0
    cVB_Pay = 0
'    bVB_PayVendoeCancel = True
    nVB_VendorQty = 1
    Unload Me
End Sub

Private Sub ocmChange_Click()
    tVB_VendorCpnNo = ""
    Unload Me
    Call wFunctionKB.W_KBCancelVendor
'    wCNShowVendor.Show vbModal
End Sub

Private Sub ocmOK_Click()
Dim tSql As String, tWhe As String
Dim nRow As Integer, nSeqNo As Integer, nSeqRC As Integer, nMax As Integer
Dim orsTemp As ADODB.Recordset
Dim cFAmt As Double, cNet As Double, cPay As Double, cItemTotal As Double, cBalance As Double

On Error GoTo ErrHandle
    '*Ao 2012/04/02 Communication Sheet ML-PRV-080
'    If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbNum.Text))) Then Exit Sub
    
    '*Tao 62-07-05 อนุญาตให้ ไม่ใส่หมายเลขคูปองได้ (Customize Vendor Coupon)
'    If otbNo.Text = "" Then
'        Call SP_OBJxSetFocus(otbNo)
'        Exit Sub
'    Else
        If SP_VALcText2Double(otbNum.Text) < SP_VALcText2Double(1) Then
            Call SP_OBJxSetFocus(otbNum)
            Exit Sub
        Else
            If SP_VALcText2Double(otbQty.Text) < SP_VALcText2Double(1) Then
                Call SP_OBJxSetFocus(otbQty)
                Exit Sub
            Else
                olaTotalamount.Caption = SP_FMTtNum2Shw(SP_VALcText2Double(otbNum.Text) * SP_VALcText2Double(otbQty.Text))   '*Ao 2012/04/02 Communication Sheet ML-PRV-080
                If SP_VALcText2Double(olaTotalamount.Caption) < SP_VALcText2Double(1) Then
                    Call SP_OBJxSetFocus(otbNum)
                    Exit Sub
                End If
            End If
        End If
'    End If
    
    cPay = SP_VALcText2Double(olaTotalamount.Caption)
    
    tVB_VendorCpnNo = Trim(otbNo.Text)
    cVB_VendorCpnAmt = cPay
    cVB_TotalDiscVendorCpnAmt = cPay
    cVB_Pay = cPay
    bVB_PayVendoeCancel = False
    nVB_VendorQty = SP_VALnText2Double(otbQty.Text)
    
    If bVB_CheckVendor Then
        'Item
        With wMain.ogdForm
            If bVB_VedFrom Then
                If nVB_RowVendor = 0 Then
                    Call SP_MSGnShowing("tms_CN005107", nCS_Information)
                    Exit Sub
                End If
                nSeqNo = nVB_RowVendor
                cItemTotal = W_FNDcSumItemVendor(nSeqNo)
                cItemTotal = cItemTotal + SP_VALcText2Double(olaTotalamount.Caption)
                
                If SP_VALcText2Double(cItemTotal) > SP_VALcText2Double(.TextMatrix(nSeqNo, .ColIndex("FCSdtSaleAmt"))) Then
                    Call SP_MSGnShowing("tms_CN005066", nCS_Error)
                    tVB_VendorCpnNo = ""
                    Call SP_OBJxSetFocus(otbNum)
                    Exit Sub
                End If
            Else                                                        '*Ao 2010/12/08 EDIT VENDOR
                nSeqNo = nW_SkuIndex
                cItemTotal = W_FNDcSumItemVendor(nSeqNo, nW_SeqNo)
                cItemTotal = cItemTotal + SP_VALcText2Double(olaTotalamount.Caption)
                If SP_VALcText2Double(cItemTotal) > SP_VALcText2Double(.TextMatrix(nSeqNo, .ColIndex("FCSdtSaleAmt"))) Then
                    Call SP_MSGnShowing("tms_CN005066", nCS_Error)
                    tVB_VendorCpnNo = ""
                    Call SP_OBJxSetFocus(otbNum)
                    Exit Sub
                Else
                    Call W_UPDxVendorCpn
                End If
                Unload Me
                Exit Sub
            End If
        End With
        
    Else
'        If cPay > cVB_Balance Then   'check amount > balance       '*Ao 2010/12/04   CommSheet ML-PRV-019-Allow Change Flag for Vendor CPN is not worked.xls
'            Call SP_MSGnShowing("tms_CN005066", nCS_Error)
'            Call SP_OBJxSetFocus(otbNum)
'            Exit Sub
'        End If
        If bVB_VedFrom Then
            Unload Me
        Else
            cBalance = SP_VALcText2Double(wSub.olaNetSale.Caption)
            cItemTotal = W_FNDcSumItemVendor(nSeqNo, nW_SeqNo)
            cItemTotal = cItemTotal + SP_VALcText2Double(olaTotalamount.Caption)
            cVB_Pay = 0 '*Tao 62-11-06 Change Vendor Coupon ต้องเคลียร์ยอดชำระ
            If SP_VALcText2Double(cItemTotal) >= SP_VALcText2Double(cBalance) Then
                Call SP_MSGnShowing("tms_CN005066", nCS_Error)
                tVB_VendorCpnNo = ""
                Call SP_OBJxSetFocus(otbNum)
                Exit Sub
            Else
                Call W_UPDxVendorCpn
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
'    With wMain.ogdForm
'        nRow = .RowSel
'        If nRow = 0 Then nRow = 1
'        nSeqNo = .TextMatrix(nRow, .ColIndex("FNSdtSeqNo"))
'    End With
    
    cFAmt = SP_VALcText2Double(olaTotalamount.Caption)
    cNet = cPay
    cVB_Change = 0
    tVB_FTRteCode = ""
    cVB_FCSrcRteAmt = 0
    cVB_FCSrcRteFac = 0
    cVB_RoundChg = 0
    
    tWhe = "WHERE   FTTmnNum = '" & tVB_CNTerminalNum & "'"
    tWhe = tWhe & " AND FTShdTransNo = '" & tVB_Running & "'"
    tWhe = tWhe & " AND FTShdTransType = '" & tVB_Transaction & "'"
    tWhe = tWhe & " AND FDShdTransDate = '" & tUT_SaleDate & "'"
    
    nSeqRC = SP_VALnText2Double(SP_SQLtFunction("MAX", "FNSrcSeqNo", tVB_CNTblRC, tWhe))
    nSeqRC = nSeqRC + 1
    tSql = "SELECT FTTdmCode,FTTdmType,FTTdmDescSale FROM TCNMTenderMtn"
    tSql = tSql & " WHERE FTTdmCode='T006'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True)
    tVB_TenderCode = SP_FEDtChkString(orsTemp, "FTTdmCode")
    tVB_TenderType = SP_FEDtChkString(orsTemp, "FTTdmType")
    tVB_TenderName = SP_FEDtChkString(orsTemp, "FTTdmDescSale")
'    tVB_TenderCode = "T006"
'    tVB_TenderType = "Vendor Coupon"
'    tVB_TenderName = "Vendor Coupon"
    
    tSql = "INSERT INTO " & tVB_CNTblRC
    tSql = tSql & "(FTTmnNum,FTShdTransNo,FNSrcSeqNo,"
    tSql = tSql & "FTShdTransType,FDShdTransDate,FTTdmCode,"
    tSql = tSql & "FTTdmType,FTSrcCardNo,FTApprCode,"
    tSql = tSql & "FCSrcCardChg,FCSrcFAmt,FCSrcAmt,"
    tSql = tSql & "FCSrcNet,FCSrcChg,FTSrcStaPrc,"
    tSql = tSql & "FTSrcRmk,FTSrcRetDocRef,FCSrcRetAmt,"
    tSql = tSql & "FNShdStaSend,"
    tSql = tSql & "FTRteCode,"
    tSql = tSql & "FCSrcRteAmt ,"
    tSql = tSql & "FCSrcRteFac,"
    tSql = tSql & "FCSrcRndMnyChg,"
    tSql = tSql & "FTSrcAbb,"
    tSql = tSql & "FTSrcCardExp,"
    tSql = tSql & "FCSrcOntopAmt,"
    tSql = tSql & "FNSrcCrdPoint,"
    tSql = tSql & "FTVendorCnpQty,FTVendorCnpFlag,FTSkuIndex,FTVendorCnpType,"
    tSql = tSql & tCS_CNLastFedIns
    tSql = tSql & " VALUES('" & tVB_CNTerminalNum & "','" & tVB_Running & "'," & nSeqRC & ","
    tSql = tSql & "'" & tVB_Transaction & "','" & tUT_SaleDate & "','" & tVB_TenderCode & "',"
    tSql = tSql & "'" & tVB_TenderType & "','" & tVB_VendorCpnNo & "','" & "" & "',"
    tSql = tSql & 0 & "," & SP_FMTcCurPoint(cFAmt, True, False) & "," & SP_FMTcCurPoint(cVB_Pay, True, False) & ","
    tSql = tSql & SP_FMTcCurPoint(cNet, True, False) & "," & SP_FMTcCurPoint(cVB_Change, True, False) & ",'',"
    tSql = tSql & "'',''," & 0 & ","
    tSql = tSql & "0,"
    tSql = tSql & "'" & tVB_FTRteCode & "',"
    tSql = tSql & SP_FMTcCurPoint(cVB_FCSrcRteAmt, True, False) & ","
    tSql = tSql & SP_FMTcCurPoint(cVB_FCSrcRteFac, True, False) & ","
    tSql = tSql & SP_FMTcCurPoint(cVB_RoundChg, True, False) & ","
    tSql = tSql & "'" & tVB_TenderName & "',"
    tSql = tSql & "'" & tVB_CrdExpM & "/" & tVB_CrdExpY & "',"
    tSql = tSql & 0 & ","
    tSql = tSql & 0 & ","
    tSql = tSql & nVB_VendorQty & ",'1','" & nSeqNo & "','" & tVB_VendorCnpType & "',"
    tSql = tSql & UT_tLastInsSql & ")"
    Call SP_SQLvExecute(tSql)
    
    Call SP_DATxRsNothing(orsTemp)
    Unload Me
    Exit Sub
    
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
'Dim tNo As String

On Error GoTo ErrHandle
    With oScanner
        'otbNo.Text = .ScanData'*Bump 56-02-28  Comment บรรทัดนี้
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
        If Trim(otbNo.Text) = "" Then
            Call SP_OBJxSetFocus(otbNo)
        Else
            '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
            '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
            If Asc(Right(otbNo.Text, 1)) = 13 Then
                otbNo.Text = Mid(otbNo.Text, 1, Len(otbNo.Text) - 1)
            End If
            Call W_FNDxCheckBarCode
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
    MsgBox "Scaner error" & " Runtime : " & Err.Number & " " & Err.Description, vbCritical
End Sub

Private Sub otbNo_GotFocus()
    otbNo.SelStart = 0
    otbNo.SelLength = Len(otbNo.Text)
End Sub

Private Sub otbNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbNo.Text = ""
        Case 13
            'If otbNo.Text <> "" Then '*Tao 62-09-12 อนุญาตให้ ไม่ใส่หมายเลขคูปองได้ (Customize Vendor Coupon)
                Call W_FNDxCheckBarCode
            'End If
    End Select
End Sub

Private Sub otbNo_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 13
'            SendKeys "{TAB}"
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

Private Sub otbNum_GotFocus()
    otbNum.SelStart = 0
    otbNum.SelLength = Len(otbNum.Text)
End Sub

Private Sub otbNum_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 13
            SendKeys "{TAB}"
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

Private Sub otbNum_Change()         '*AO51-03-12    Comm.Sheet SF- 070
    otbNum.Text = SP_DEFtChangeNum(Trim(otbNum.Text))
    otbNum.SelStart = Len(otbNum.Text)
End Sub

'*Ao 2012/04/02 Communication Sheet ML-PRV-080
'Private Sub otbNum_LostFocus()
'    If SP_CHKbBanknoteAlw(SP_VALcText2Double(otbNum.Text)) Then
'        olaTotalamount.Caption = SP_FMTtNum2Shw(SP_VALcText2Double(otbNum.Text) * SP_VALcText2Double(otbQty.Text))
'    Else
'        Call SP_OBJxSetFocus(otbNum)
'    End If
'End Sub

Private Sub otbNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbNum.Text = ""
    End Select
End Sub

Private Sub otbQty_Change()
    olaTotalamount.Caption = SP_FMTtNum2Shw(SP_VALcText2Double(otbNum.Text) * SP_VALcText2Double(otbQty.Text))
End Sub

Private Sub otbQty_GotFocus()
    otbQty.SelStart = 0
    otbQty.SelLength = Len(otbQty.Text)
End Sub

'Private Sub otbQty_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyDelete         'delete
'            otbQty.Text = ""
'        Case 13
'            If otbNo.Text <> "" Then
'                Call W_FNDxCheckBarCode
'            End If
'    End Select
'End Sub

Private Sub otbQty_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 13
            olaTotalamount.Caption = SP_FMTtNum2Shw(SP_VALcText2Double(otbNum.Text) * SP_VALcText2Double(otbQty.Text))
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
    Exit Sub
    
ErrHandle:

End Sub


