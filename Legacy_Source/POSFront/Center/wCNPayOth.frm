VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNPayOth 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2610
   ClientLeft      =   555
   ClientTop       =   1770
   ClientWidth     =   6540
   ControlBox      =   0   'False
   Icon            =   "wCNPayOth.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6540
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
         Text            =   "0"
         Top             =   1080
         Width           =   2415
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
         Top             =   480
         Width           =   3975
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
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1080
         Width           =   2280
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No :"
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
         Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
         Top             =   480
         Width           =   540
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1425
      Width           =   6525
      _Version        =   720897
      _ExtentX        =   11509
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
         Left            =   4920
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
         Left            =   3360
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNPayOth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tW_Name As String
Public tW_CouponSta As String '*Eaw 56-05-08 CommSheet TK-ISS3000-179-ST 00 00 01.xls
Private dW_CVExpDate As Date

Private Function W_CHKbVoucherNo(ByVal ptVoucherNo$) As Boolean
'-----------------------------------
'Ret: True  ผ่านการตรวจสอบ False ไม่ผ่านการตรวจสอบ
'Cmt: ตรวจสอบหมายเลข Voucher ว่าสามารถนำมาชำระได้
'------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    W_CHKbVoucherNo = False
    
    '*KT 51-09-19 ML-SF- 557 Voucher Cancel Is Wrong
    'Step1 Voucher ที่ยกเลิกต้องไม่ใช่ Voucher ที่จะจ่ายคืน
    If tVB_CancelVoucherNo <> ptVoucherNo Then
        'Step2 ตรวจสอบว่า Voucher ที่ระบุไม่เคยถูกยกเลิก
        tSql = "SELECT FTShdVoucherNo,FCShdVoucherAmt ,FTTmnNum,FTShdTransNo,FDShdTransDate"
        tSql = tSql & vbCrLf & "FROM TPSTVoucherHD"
        tSql = tSql & vbCrLf & "WHERE  FTShdVoucherNo = '" & ptVoucherNo & "'"
        tSql = tSql & vbCrLf & "AND  ISNULL(FTShdStaCancel,'N') <> 'Y'"
        tSql = tSql & vbCrLf & "AND  FTShdStaDoc ='1'"
        tSql = tSql & vbCrLf & "AND  FTShdTransType = '" & IIf(bVB_TrainingMode, Format(tEN_TrnTraining, "00"), Format(tEN_TrnVoucher, "00")) & "'"
        tSql = tSql & vbCrLf & "AND  (FTShdDepRefTmnNum = '' OR FTShdDepRefTmnNum IS NULL)"
        tSql = tSql & vbCrLf & "AND  (FTShdDepReTransNo = '' OR FTShdDepReTransNo IS NULL)"
        If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then
            If Not (orsTemp.BOF Or orsTemp.EOF) Then
                cVB_VoucherAmtRet = SP_FEDvChkNumeric(orsTemp, "FCShdVoucherAmt")
                tVB_VoucherTmnNoRet = SP_FEDtChkString(orsTemp, "FTTmnNum")
                tVB_VoucherTrnNoRet = SP_FEDtChkString(orsTemp, "FTShdTransNo")
                dVB_VoucherTrnDateRet = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
                W_CHKbVoucherNo = True
            Else
                cVB_VoucherAmtRet = 0
            End If
        Else
            cVB_VoucherAmtRet = 0
        End If
    Else
        cVB_VoucherAmtRet = 0
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    cVB_VoucherAmtRet = 0
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function W_CHKbCouponNo(ptCouponNo As String) As Boolean
'-------------------------------------------------------------------------------
'*Eaw 55-12-01
'Cmt: ตรวจสอบหมายเลข Coupon ว่าสามารถนำมาชำระได้ ณ ตอนนี้ Coupon จะถูกเก็บใน TPSTCV โดยที่มี FTCVFlag='1'
'-------------------------------------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tExpDate As String
    
On Error GoTo ErrHandle
    W_CHKbCouponNo = False
    tSql = "SELECT *  FROM TPSTCV WITH(NOLOCK) "
    tSql = tSql & " WHERE FTCVNum= '" & ptCouponNo & "'"
    tSql = tSql & " AND FTCVFlag = '1' "
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine)
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        tVB_CouponNo = SP_FEDtChkString(orsTemp, "FTCVNum")
        cVB_CouponAmtRet = SP_FEDvChkNumeric(orsTemp, "FCCVAmt")  '*Eaw 56-05-08 CommSheet TK-ISS3000-179-ST 00 00 01.xls
        tW_CouponSta = SP_FEDtChkString(orsTemp, "FTCVStatusCode") '*Eaw 56-05-08 CommSheet TK-ISS3000-179-ST 00 00 01.xls
        dW_CVExpDate = SP_FEDdChkDate(orsTemp, "FDCVExpDate")
        tVB_CrdExpM = Format(Month(SP_FEDdChkDate(orsTemp, "FDCVExpDate")))     '*Ao 2014/10/20   eCoupon  V 4.2.152
        tVB_CrdExpY = Format(Year(SP_FEDdChkDate(orsTemp, "FDCVExpDate")))
        tVB_CrdExpY = Right(tVB_CrdExpY, 2)
        W_CHKbCouponNo = True
        Call SP_DATxRsNothing(orsTemp)
        Exit Function
    End If
    W_CHKbCouponNo = False
    cVB_CouponAmtRet = 0 '*Eaw 56-05-08 CommSheet TK-ISS3000-179-ST 00 00 01.xls
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_CHKbCouponNo = False
    cVB_CouponAmtRet = 0 '*Eaw 56-05-08 CommSheet TK-ISS3000-179-ST 00 00 01.xls
    Call SP_DATxRsNothing(orsTemp)
End Function
Private Sub W_FRMxSetInitial()
    otbNo.Text = ""
    otbNum.Text = 0
    '*KT 51-03-27 Comsheet ML-SF-130
    Select Case tVB_TenderCode
        Case "T010" 'Voucher
            otbNum.Locked = True
            otbNum.BackColor = nVB_CNClrBLock
        Case "T014" 'Coupon   '*Eaw 56-05-08 CommSheet TK-ISS3000-179-ST 00 00 01.xls
            otbNum.Locked = True
            otbNum.BackColor = nVB_CNClrBLock
    End Select
End Sub
'Private Sub Form_Activate()
'On Error GoTo ErrHandle
'    Call SP_OTBxSelText(otbNo)
'    Exit Sub
'ErrHandle:
'
'End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    If bVB_Res640X480 Then
'        Me.Left = 0
        Me.Top = 2500
    End If
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
    Me.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu
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
'    olaForm(1).Caption = tW_Name               '*Ao 2013/03/20  TK
    olaForm(1).Caption = tVB_Tender
'    otbNum.Text = SP_FMTcCurPoint(cVB_Balance, True, True)  '*Eaw 56-07-01CommSheet TK-ISS3000-253.xls
    If tVB_TenderCode = "T014" Then
        otbNum.Text = SP_FMTcCurPoint(0, True, True)        '*Ao 2014/10/20   eCoupon  V 4.2.152
    Else
        otbNum.Text = SP_FMTcCurPoint(cVB_Balance, True, True)  '*Eaw 56-07-01CommSheet TK-ISS3000-253.xls
    End If
    cVB_Pay = 0 '*Tao 58-12-14 เคลียร์ค่าการชำระเงิน
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
    cVB_Pay = 0
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim cPay As Double
    
    Select Case tVB_TenderCode
        Case "T010"
            If Not W_CHKbVoucherNo(Trim(otbNo.Text)) Or otbNo.Text = "" Then
                 Call SP_MSGnShowing("tms_CN005067", nCS_Warning)
                 Call SP_OTBxSelText(otbNo)
            Else
                otbNum.Text = SP_FMTcCurPoint(cVB_VoucherAmtRet, False, True)   '* AO51-03-26
                otbNum.Locked = True
                otbNum.BackColor = &HC0E0FF         '&H00C0E0FF,65535
            End If
        Case "T014"         '*Eaw 56-05-08 CommSheet TK-ISS3000-179-ST 00 00 01.xls
            If Not W_CHKbCouponNo(Trim(otbNo.Text)) Or otbNo.Text = "" Then
'                otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
                otbNum.Text = SP_FMTcCurPoint(0, False, True)                                   '*Ao 2014/10/20   eCoupon  V 4.2.152
                Call SP_MSGnShowing("tms_CN005239", nCS_Warning) '*Invalid Coupon No.
                Call SP_OTBxSelText(otbNo)
                Exit Sub
            Else
                '* Bump 57-11-09 ปรับเป็น Select Case
                    Select Case tW_CouponSta
                    Case "2"
                        otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
                        Call SP_MSGnShowing("tms_CN999025", nCS_Warning) '*eCoupon is already cancelled' email K.Krittapak 18/11/2014 16:01pm
                        Call SP_OTBxSelText(otbNo)
                        
                        Exit Sub
                        
                    Case "3"
                        otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True) '*email K.Krittapak 18/11/2014 16:01pm
                        Call SP_MSGnShowing("tms_CN999024", nCS_Warning) '* This is a lost of eCoupon
                        Call SP_OTBxSelText(otbNo)
                        
                        Exit Sub
                        
                    Case "4"
                        otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
                        Call SP_MSGnShowing("tms_CN999026", nCS_Warning) '*eCoupon is status pickup
                        Call SP_OTBxSelText(otbNo)
                        
                        Exit Sub
                        
                    Case Else
                        If CDate(Format(dW_CVExpDate, "yyyy/MM/dd")) < CDate(Format(dVB_CurrentDate, "yyyy/MM/dd")) Then
                            otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
                            Call SP_MSGnShowing("tms_CN999027", nCS_Warning) '*This eCoupon is expired'
                            
                            Exit Sub
                            
                        Else
                            otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
                            otbNum.Locked = True
                            otbNum.BackColor = &HC0E0FF         '&H00C0E0FF,65535
                        End If
                    End Select
                End If
'                If tW_CouponSta = "4" Then
'                    otbNum.Text = SP_FMTcCurPoint(0, False, True)
''                    Call SP_MSGnShowing("tms_CN005239", nCS_Error)
'                    MsgBox "eCoupon is status pickup", vbOKOnly + vbExclamation, tVB_CNAppTitle         '*Ao 2014/10/22    eCoupon   V 4.2.152
'                    Call SP_OTBxSelText(otbNo)
'                    Exit Sub
'                ElseIf tW_CouponSta = "3" Then
'                    otbNum.Text = SP_FMTcCurPoint(0, False, True)
''                    Call SP_MSGnShowing("tms_CN005239", nCS_Error)
'                    MsgBox "This is a lost of eCoupon", vbOKOnly + vbExclamation, tVB_CNAppTitle         '*Ao 2014/10/22    eCoupon   V 4.2.152
'                    Call SP_OTBxSelText(otbNo)
'                    Exit Sub
'                ElseIf tW_CouponSta = "2" Then
'                    otbNum.Text = SP_FMTcCurPoint(0, False, True)
'                    MsgBox "eCoupon is already cancelled", vbOKOnly + vbExclamation, tVB_CNAppTitle         '*Bump57-11-13    eCoupon   V 4.2.156
'                    Call SP_OTBxSelText(otbNo)
'                    Exit Sub
'                Else
'                    otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
'                    otbNum.Locked = True
'                    otbNum.BackColor = &HC0E0FF         '&H00C0E0FF,65535
'                End If
'            End If
'            If CDate(Format(dW_CVExpDate, "yyyy/MM/dd")) < CDate(Format(dVB_CurrentDate, "yyyy/MM/dd")) Then     '*Ao 2014/10/22    eCoupon   V 4.2.152
'                MsgBox "eCoupon Expire Date", vbOKOnly + vbInformation, tVB_CNAppTitle        '*Ao 2014/10/22    eCoupon   V 4.2.152
'                otbNum.Text = SP_FMTcCurPoint(0, False, True)
''                otbNo.Text = String(12, "0")
'                Call SP_OTBxSelText(otbNo)
''                otbNo.SelStart = Len(otbNo.Text)
'                Exit Sub
'            End If
        End Select
    
    '*KT 51-09-2008 แก้ไขให้ Credit Store สามารถชำระเศษสตางค์ได้
    '*Ao 2012/04/02 Communication Sheet ML-PRV-080   "T007" 'Special Coupon ไม่ต้องเช็คหน่วยสตางค์
'    If tVB_TenderCode <> "T008"  Then
'    If (tVB_TenderCode <> "T008") And (tVB_TenderCode <> "T007") Then    '*Eaw 56-01-04 แก้ไขการปัด Rounding ไม่เช็ค Bank note ใน TK Project
'        If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbNum.Text))) Then Exit Sub       '*AO51-03-26   Comm.Sheet  ML-SF-100
'    End If
    
    cPay = SP_VALcText2Double(otbNum.Text)
    '*Eaw  56-01-30  Gift Voucher No. ใส่หรือไม่ไส่ก็ได้
'    If Trim(otbNo.Text) = "" Then 'check amount zero
'        Call SP_OBJxSetFocus(otbNo) 'otbNo.SetFocus
'        Exit Sub
'    End If
'    If tVB_TenderCode = "T004" Then
    If (tVB_TenderCode = "T004") Or (tVB_TenderCode = "T005") Then
        'none
    Else
        If Trim(otbNo.Text) = "" Then 'check amount zero
            Call SP_OBJxSetFocus(otbNo) 'otbNo.SetFocus
            Exit Sub
        End If
    End If
'-------------------------------------------------------------------------
    If cPay = 0 Then 'check amount zero
        Call SP_MSGnShowing("tms_CN005064", nCS_Warning)
        Exit Sub
    End If
    
'    If cPay > cVB_Balance Then 'check amount > balance
'        Call SP_MSGnShowing("tms_CN005066", nCS_Error)
'        Exit Sub
'    End If

   Select Case tVB_TenderCode
        Case "T004" 'Gift Voucher
            tVB_GiftVoucherNo = Trim(otbNo.Text)
        Case "T005", "T014" 'Special Gift Voucher  '*KT 54-07-29 V.4.0.1 Req.5 Phase4.0.1(Member Point) Add new Tender Type ,Add 1 Type of Special Gift Voucher apply payment same as “Special Gift Voucher"
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            Select Case SP_VALnText2Double(tVB_TransactionSub)
                Case tEN_TrnRedeem
                    oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtRedeem, False, True)
                Case tEN_TrnAddDeposit
                    oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtBAddDpt, False, True)
            End Select
            oDisc.cDiscValue = SP_VALcText2Double(otbNum.Text)
            cVB_SpcGiftVoucherAmtAffect = SP_FMTcCurPoint(oDisc.DIScSpecial(tVB_TenderCode), True, False)
            If cVB_SpcGiftVoucherAmtAffect < 0 Then
                cVB_SpcGiftVoucherAmtAffect = 0
            Else
'                Call SP_PRCxReMemberPoint '*KT 52-03-03 Member Point ทำการ Re Member point กรณีที่ใฃ้ 'Special Gift Voucher แล้ว แล้วทำให้ Mall Card ลดลง
                '*KT 54-05-23 V.3.0.60 ML-PRV-58 Total ticket Calculate incorrect เกิด Level Down ขึ้นมาจากการแก้ไข CommSheet ML-PRV-033-Special Gift Voucher,Special Coupon, Credit Card can't re-entry.xls
                 'ให้กลับไปเหมือนเดิม แล้วให้แก้ไข CommSheet ML-PRV-033-Special Gift Voucher,Special Coupon, Credit Card can't re-entry.xls ที่ W_KBRcvSpcGiftVoucher แทน
'                If cVB_SpcGiftVoucherAmtAffect < 1 Then         '*Ao 2011/03/29 CommSheet ML-PRV-033-Special Gift Voucher,Special Coupon, Credit Card can't re-entry.xls
'                    cVB_SpcGiftVoucherAmtAffect = 0
'                End If
            End If
            
            Select Case Val(Format(tVB_TransactionSub))
                 Case tEN_TrnDeposit, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                       '*KT 52-06-04 PH1.8-ML-SF-042
                        If bVB_DepositAll Then
                            cVB_Balance = cVB_Balance + cVB_SpcGiftVoucherAmtAffect
                        End If
                Case tEN_TrnRedeem
                        cVB_Balance = cVB_Balance + cVB_SpcGiftVoucherAmtAffect
                 Case Else
                        cVB_Balance = cVB_Balance + cVB_SpcGiftVoucherAmtAffect
            End Select
            tVB_SpcGiftVoucherNo = Trim(otbNo.Text)
        Case "T010" 'Voucher
            If Not W_CHKbVoucherNo(Trim(otbNo.Text)) Then
                Call SP_MSGnShowing("tms_CN005067", nCS_Error) 'Invaild Voucher No
                otbNo.SetFocus
                Exit Sub
            End If
            If cPay > cVB_VoucherAmtRet Then
                Call SP_MSGnShowing("tms_CN005064", nCS_Error)
                otbNum.SetFocus
            End If
            tVB_VoucherNoRet = Trim(otbNo.Text)
            cVB_VoucherAmtRet = cPay
        Case "T007" 'Special Coupon เหมือนกันกับ T005
            oDisc.dTranDate = tUT_SaleDate
            oDisc.tTerminalNo = tVB_CNTerminalNum
            oDisc.tTransNo = tVB_Running
            Select Case SP_VALnText2Double(tVB_TransactionSub)
                Case tEN_TrnRedeem
                    oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtRedeem, False, True)
                Case tEN_TrnAddDeposit
                    oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtBAddDpt, False, True)
            End Select
            oDisc.cDiscValue = SP_VALcText2Double(otbNum.Text)
            cVB_SpcCpnAmtAffect = SP_FMTcCurPoint(oDisc.DIScSpecial("T007"), True, False)
            If cVB_SpcCpnAmtAffect < 0 Then
                cVB_SpcCpnAmtAffect = 0
            Else
'                 Call SP_PRCxReMemberPoint '*KT 52-03-03 Member Point ทำการ Re Member point กรณีที่ใฃ้ Special Coupon เหมือนกันกับ T005 แล้วทำให้ Mall Card ลดลง
                 '*KT 54-05-23 V.3.0.60 ML-PRV-58 Total ticket Calculate incorrect เกิด Level Down ขึ้นมาจากการแก้ไข CommSheet ML-PRV-033-Special Gift Voucher,Special Coupon, Credit Card can't re-entry.xls
                  'ให้กลับไปเหมือนเดิม แล้วให้แก้ไข CommSheet ML-PRV-033-Special Gift Voucher,Special Coupon, Credit Card can't re-entry.xls ที่ W_KBRcvSpecialCpn  แทน
'                 If cVB_SpcCpnAmtAffect < 1 Then        '*Ao 2011/03/29 CommSheet ML-PRV-033-Special Gift Voucher,Special Coupon, Credit Card can't re-entry.xls
'                    cVB_SpcCpnAmtAffect = 0
'                 End If
            End If
            Select Case Val(Format(tVB_TransactionSub))
                 Case tEN_TrnDeposit, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
                        '*KT 52-06-04 PH1.8-ML-SF-042
                        If bVB_DepositAll Then
                            cVB_Balance = cVB_Balance + cVB_SpcCpnAmtAffect
                        End If
                Case tEN_TrnRedeem
                        cVB_Balance = cVB_Balance + cVB_SpcCpnAmtAffect
                 Case Else
                        cVB_Balance = cVB_Balance + cVB_SpcCpnAmtAffect
            End Select
            tVB_SpecialCpnNo = Trim(otbNo.Text)
        Case "T008" 'Credit Store
            '*KT 51-10-22 V 3.0.140/V 3.0.200
            '------------------------------------------------------------------------------------------
            If cPay > CDbl(SP_FMTcCurPoint(cVB_Balance - cVB_RoundRcv, True, False)) Then
                Call SP_MSGnShowing("tms_CN005066", nCS_Error)
                Call SP_OBJxSetFocus(otbNum)
                Exit Sub
            End If
            '------------------------------------------------------------------------------------------
            tVB_CrdStoreNo = Trim(otbNo.Text)
        Case "T009" 'Store Coupon
            tVB_StoreCpnNo = Trim(otbNo.Text)
        Case "T011" 'Credit Note
            tVB_CreditNoteNo = Trim(otbNo.Text)

    End Select
    cVB_Pay = cPay
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
            If Asc(Right(otbNo.Text, 1)) = 13 Then
                otbNo.Text = Mid(otbNo.Text, 1, Len(otbNo.Text) - 1)
            End If
            Select Case tVB_TenderCode
                Case "T010"
                    If Not W_CHKbVoucherNo(Trim(otbNo.Text)) Or otbNo.Text = "" Then
                         Call SP_MSGnShowing("tms_CN005067", nCS_Error)
                         Call SP_OTBxSelText(otbNo)
                    Else
                        otbNum.Text = SP_FMTcCurPoint(cVB_VoucherAmtRet, False, True)   '* AO51-03-26
                        otbNum.Locked = True
                        otbNum.BackColor = &HC0E0FF         '&H00C0E0FF,65535
                        Call SP_OBJxSetFocus(otbNum)
                    End If
                Case Else
                    Call SP_OBJxSetFocus(otbNum)
            End Select
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

Private Sub otbNo_GotFocus()
    otbNo.SelStart = 0
    otbNo.SelLength = Len(otbNo.Text)
'    If tVB_TenderCode = "T014" Then
'        otbNo.SelStart = Len(otbNo.Text)
'    End If
End Sub

Private Sub otbNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbNo.Text = ""
        Case 13
        '*Eaw  56-01-30  Gift Voucher No. ใส่หรือไม่ไส่ก็ได้
'            If otbNo.Text <> "" Then
'                SendKeys "{TAB}"
'            End If
            If tVB_TenderCode = "T004" Or tVB_TenderCode = "T005" Then  '*Eaw 56-06-04 CommSheet TK-ISS3000-224-ST 00 00 01.xls เพิ่ม T005 เพราะ Gift Voucher ปัจจุบันเปลี่ยนจาก TK Gift Voucher
'                SendKeys "{TAB}"
                Call SP_OBJxSetFocus(otbNum)                '*Ao 2015/02/09   V 4.2.165     CommSheet TK-ISS3000-577.xls
            Else
                If otbNo.Text <> "" Then
'                    SendKeys "{TAB}"
                    Call SP_OBJxSetFocus(otbNum)                '*Ao 2015/02/09   V 4.2.165     CommSheet TK-ISS3000-577.xls
                End If
            End If
            
        '--------------------------------------------------------------
    End Select
End Sub

Private Sub otbNo_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------------
'Cmt: *KT 52-10-27 ปรับปรัง
'---------------------------------------------------------------------------
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii             '*AO 51-04-01  Comm.Sheet ML-SF-175
        Case 13
            Select Case tVB_TenderCode
                Case "T010"
                    If Not W_CHKbVoucherNo(Trim(otbNo.Text)) Or otbNo.Text = "" Then
                         otbNum.Text = SP_FMTcCurPoint(cVB_VoucherAmtRet, False, True)   '* AO51-03-26
                         Call SP_MSGnShowing("tms_CN005067", nCS_Error)
                         Call SP_OTBxSelText(otbNo)
                    Else
                        otbNum.Text = SP_FMTcCurPoint(cVB_VoucherAmtRet, False, True)   '* AO51-03-26
                        otbNum.Locked = True
                        otbNum.BackColor = &HC0E0FF         '&H00C0E0FF,65535
                    End If
                Case "T014"   '*Eaw 55-12-01  add case T014
                    '*Eaw 56-05-08 CommSheet TK-ISS3000-179-ST 00 00 01.xls
'                    If Not W_CHKbCouponNo(Trim(otbNo.Text)) Or otbNo.Text = "" Then
'                        Call SP_MSGnShowing("tms_CN005239", nCS_Error)
'                        Call SP_OTBxSelText(otbNo)
'                    End If
                    otbNo.Text = Right(String(13, "0") & otbNo.Text, 13)                                        '*Ao 2014/10/22    eCoupon   V 4.2.154
                    If Not W_CHKbCouponNo(Trim(otbNo.Text)) Or otbNo.Text = "" Then
                        otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
                        Call SP_MSGnShowing("tms_CN005239", nCS_Warning) '*Invalid Coupon No.
'                        otbNo.Text = String(12, "0")
                        Call SP_OTBxSelText(otbNo)
'                        otbNo.SelStart = Len(otbNo.Text)
                    Else
                        '* Bump 57-11-09 ปรับเป็น Select Case และเพิ่มการ เช็ค วันหมดอายุ
                        Select Case tW_CouponSta
                        Case "2"
                            otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True) '*email K.Krittapak 18/11/2014 16:01pm
                            Call SP_MSGnShowing("tms_CN999025", nCS_Warning) '*eCoupon is already cancelled'
                            Call SP_OTBxSelText(otbNo)
                        Case "3"
                            otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True) '*email K.Krittapak 18/11/2014 16:01pm
                            Call SP_MSGnShowing("tms_CN999024", nCS_Warning) '* This is a lost of eCoupon
                            Call SP_OTBxSelText(otbNo)
                        Case "4"
                            otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
                            Call SP_MSGnShowing("tms_CN999026", nCS_Warning) '*eCoupon is status pickup
                            Call SP_OTBxSelText(otbNo)
                        Case Else
                            If CDate(Format(dW_CVExpDate, "yyyy/MM/dd")) < CDate(Format(dVB_CurrentDate, "yyyy/MM/dd")) Then
                                otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
                                Call SP_MSGnShowing("tms_CN999027", nCS_Warning) '*This eCoupon is expired'
                            Else
                                otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
                                otbNum.Locked = True
                                otbNum.BackColor = &HC0E0FF         '&H00C0E0FF,65535
                                Call SP_OBJxSetFocus(ocmOk)
                            End If
                        End Select
'                        If tW_CouponSta = "3" Then '*Bump 57-11-13 แก้จาก 4 เป็น 3 email คุณเถียว 57-11-12 16:30pm. V.4.2.156
'                            otbNum.Text = SP_FMTcCurPoint(0, False, True)
'                            MsgBox "This is a lost of eCoupon", vbOKOnly + vbExclamation, tVB_CNAppTitle         '*Ao 2014/10/22    eCoupon   V 4.2.152
'                            Call SP_OTBxSelText(otbNo)
'                        Else
'                            If tW_CouponSta = "2" Then
'                                otbNum.Text = SP_FMTcCurPoint(0, False, True)
'                                MsgBox "eCoupon is already cancelled", vbOKOnly + vbExclamation, tVB_CNAppTitle         '*Ao 2014/11/07   V 4.2.155   CommSheet TK-ISS3000-566
'                                Call SP_OTBxSelText(otbNo)
'                            Else
'                                otbNum.Text = SP_FMTcCurPoint(cVB_CouponAmtRet, False, True)
'                                otbNum.Locked = True
'                                otbNum.BackColor = &HC0E0FF         '&H00C0E0FF,65535
'                                Call SP_OBJxSetFocus(ocmOk)
'                            End If
'                        End If
                    End If
            End Select
    End Select
End Sub

Private Sub otbNo_LostFocus()
    otbNo.Text = SP_TXTtTrimRightLeft(otbNo.Text)
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









