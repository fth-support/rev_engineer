VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNPayCV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Voucher;klg_AccNum"
   ClientHeight    =   3375
   ClientLeft      =   7260
   ClientTop       =   4605
   ClientWidth     =   6570
   ControlBox      =   0   'False
   Icon            =   "wCNPayCV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6570
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2775
      Index           =   2
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   -240
      Width           =   6555
      _Version        =   720897
      _ExtentX        =   11562
      _ExtentY        =   4895
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbTel 
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
         MaxLength       =   12
         TabIndex        =   4
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox otbCV 
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
         MaxLength       =   13
         TabIndex        =   0
         Top             =   360
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
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   3
         Text            =   "0"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox orbMonth 
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
         MaxLength       =   2
         TabIndex        =   1
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox orbYear 
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
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   2
         Top             =   960
         Width           =   615
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number:;klg_Tel"
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
         TabIndex        =   14
         Tag             =   "2;"
         Top             =   2160
         Width           =   3150
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
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Tag             =   "2;"
         Top             =   1035
         Width           =   2130
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Voucher:;klg_AccNum"
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
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   435
         Width           =   3615
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
         Top             =   1635
         Width           =   2160
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Height          =   435
         Index           =   4
         Left            =   3120
         TabIndex        =   9
         Top             =   960
         Width           =   210
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(MM/YY);klg_FmtDate"
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
         Index           =   6
         Left            =   4200
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   1035
         Width           =   2775
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1080
      Index           =   1
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2265
      Width           =   6555
      _Version        =   720897
      _ExtentX        =   11562
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   240
         Top             =   480
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
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
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   400
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
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   400
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNPayCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bW_FoundAcc As Boolean
Private bW_KeyEnter As Boolean
Private tW_CVStatusCode As String
Private cW_CVAmt  As Double
Private dW_CVDate As Date
Private dW_CVExpDate As Date
Private tW_CVTel As String '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls

Private Function W_PRCbFindCVFromRC() As Boolean
'-----------------------------------------------------------------
'   Call:
'   Cmt:    '*Eaw 56-11-04 ตรวจสอบว่า CV ใบนี้มีการใช้ในบิลปัจจุบันแล้วหรือไม่
'-----------------------------------------------------------------
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    W_PRCbFindCVFromRC = False
    tSql = "SELECT * FROM " & tVB_CNTblRC
    tSql = tSql & vbCrLf & "WHERE  FTSrcCardNo = '" & tVB_CNCVNo & "' AND FTTdmCode = 'T018'"
    tSql = tSql & vbCrLf & " AND  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & vbCrLf & " AND FDShdTransDate='" & tUT_SaleDate & "'"
    If SP_SQLbExecute(tSql, orsTemp) Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            W_PRCbFindCVFromRC = True
        End If
    End If
    Exit Function
ErrHandle:

End Function

Private Sub W_PRCxLockText(ByVal pbLock As Boolean)
'-----------------------------------------------------------------
'   Call:
'   Cmt:    '*Ao 2013/03/16
'-----------------------------------------------------------------
On Error GoTo ErrHandle

    otbCV.Locked = pbLock
    orbMonth.Locked = pbLock
    orbYear.Locked = pbLock
    otbAmt.Locked = pbLock
    If pbLock Then
        otbCV.BackColor = nVB_CNClrBLock
        orbMonth.BackColor = nVB_CNClrBLock
        orbYear.BackColor = nVB_CNClrBLock
        otbAmt.BackColor = nVB_CNClrBLock
    End If
    
Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub Form_Load()
    Call SP_FRMxChildCenter(Me)
    If bVB_Res640X480 Then
        Me.Top = 2500
    Else
        Me.Top = 2900
    End If
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Me.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu
    'Call W_FRMxSetInitial
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
    
    Call W_PRCxLockText(False)          '*Ao 2013/03/16
    
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
            otbAmt.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
            otbCV.Text = W_GETtAutoCVNum
            otbCV.Locked = True
            otbCV.BackColor = nVB_CNClrBLock
            orbYear.Text = Mid(CStr(Year(dVB_CurrentDate) + 1), 3, 2)
            orbYear.Locked = True
            orbYear.BackColor = nVB_CNClrBLock
            orbMonth.Text = Format(Month(dVB_CurrentDate), "00")
            orbMonth.Locked = True
            orbMonth.BackColor = nVB_CNClrBLock
            otbAmt.Locked = True
            otbAmt.BackColor = nVB_CNClrBLock
            ocmOk.TabStop = True
            otbCV.TabStop = False
            orbMonth.TabStop = False
            orbYear.TabStop = False
            otbAmt.TabStop = False
            '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
'            Call SP_OBJxSetFocus(ocmOk)
            otbTel.TabStop = True
            otbTel.Text = ""
            Call SP_OTBxSelText(otbTel)

        Case tEN_TrnAddDepositCancel, tEN_TrnDepositCancel
            otbAmt.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
            otbCV.Text = W_GETtAutoCVNum
            otbCV.Locked = True
            otbCV.BackColor = nVB_CNClrBLock
            orbYear.Text = Mid(CStr(Year(dVB_CurrentDate) + 1), 3, 2)
            orbYear.Locked = True
            orbYear.BackColor = nVB_CNClrBLock
            orbMonth.Text = Format(Month(dVB_CurrentDate), "00")
            orbMonth.Locked = True
            orbMonth.BackColor = nVB_CNClrBLock
            otbAmt.Locked = True
            otbAmt.BackColor = nVB_CNClrBLock
            ocmOk.TabStop = True
            otbCV.TabStop = False
            orbMonth.TabStop = False
            orbYear.TabStop = False
            otbAmt.TabStop = False
            Call SP_OBJxSetFocus(ocmOk)
        Case Else 'sale,layaway
            'none
            otbAmt.Text = SP_FMTcCurPoint(0, False, True)
            bW_FoundAcc = False '*Eaw 56-10-29
    End Select
    cVB_Pay = 0 '*Tao 58-12-14 เคลียร์ค่าการชำระเงิน
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    With oScanner
        If bVB_OPOSScannerEnable Then
            .DataEventEnabled = False
            If .DeviceEnabled Then
                .DeviceEnabled = False
                .Release
                .Close
            End If
        End If
    End With
    '*Eaw 56-10-29
    bW_FoundAcc = False
    otbCV.Text = ""
    orbMonth.Text = ""
    orbYear.Text = ""
    otbAmt.Text = SP_FMTcCurPoint(0, False, True)
    otbTel.Text = ""
    cW_CVAmt = 0
    '-------------------------------------------------------------------
    Call SP_OBJxSetNothing(Me)
    
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
On Error GoTo ErrHandle
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
    cVB_Pay = 0
    tVB_CNCVNo = ""
    cVB_CNCVAmt = 0
    tVB_CrdExpM = ""
     tVB_CrdExpY = ""
     tVB_CNCVTel = "" '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
    Unload Me
    Exit Sub
ErrHandle:
End Sub

Private Sub ocmOK_Click()
On Error GoTo ErrHandle
    Dim cPay As Double
     If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
     
     tVB_CNCVNo = Trim(otbCV.Text)
     tVB_CrdExpM = orbMonth.Text
     tVB_CrdExpY = orbYear.Text
     tVB_CNCVTel = Trim(otbTel.Text) '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
     'Check Expire  Date
    If Format(orbMonth.Text, "00") = "00" Then
        Call SP_MSGnShowing("tms_CN005227", nCS_Warning)
        orbMonth.Text = ""
        orbYear.Text = ""
        Call SP_OBJxSetFocus(orbMonth)
        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
        Exit Sub
    End If
    
    If Format(orbYear.Text, "00") = "00" Then
        Call SP_MSGnShowing("tms_CN005227", nCS_Warning)
        orbMonth.Text = ""
        orbYear.Text = ""
        Call SP_OBJxSetFocus(orbMonth)
        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
        Exit Sub
    End If
    
    dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)), 1)
    dVB_CNCVDate = dVB_CrdDateExp
    
'    If Format(dVB_CrdDateExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then
'        Call SP_MSGnShowing("tms_CN005227", nCS_Warning)
'        orbMonth.Text = ""
'        orbYear.Text = ""
'        Call SP_OBJxSetFocus(orbMonth)
'        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
'        Exit Sub
'    End If

   Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
            'none
        Case tEN_TrnAddDepositCancel, tEN_TrnDepositCancel
            'none
        Case Else
            If aVB_TenderMnu(nVB_CurTenderSeqNo).StaExpDate = True Then
'                If Not (tVB_CrdExpM = Format(Month(dW_CVExpDate), "00") And tVB_CrdExpY = Format(Mid(Year(dW_CVExpDate), 3, 2), "00")) Then
'                CDate(Format(dVB_CrdDateExp, "YYYY/MM/DD")) < CDate(Format(dVB_CurrentDate, "YYYY/MM/DD")))
                If CDate(Format(dW_CVExpDate, "yyyy/MM/dd")) < CDate(Format(dVB_CurrentDate, "yyyy/MM/dd")) Then    '*BG 56-09-07  Comsheet-348
                        Call SP_MSGnShowing("tms_CN005227", nCS_Warning)
                        orbMonth.Text = ""
                        orbYear.Text = ""
                        Call SP_OBJxSetFocus(orbMonth)
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                        If vbOK Then  '*BG 56-09-07  Comsheet-348
                            Call W_PRCxLockText(False)
                            otbCV.Text = ""
                            otbCV.BackColor = nVB_CNClrBUnLock
                            orbMonth.Text = ""
                            orbMonth.BackColor = nVB_CNClrBUnLock
                            orbYear.Text = ""
                            orbYear.BackColor = nVB_CNClrBUnLock
                            otbAmt.Text = "0.00"
                            otbAmt.BackColor = nVB_CNClrBUnLock
                            otbTel.Text = ""
                            otbTel.BackColor = nVB_CNClrBUnLock
                            Call SP_OTBxSelText(otbTel)
                            Call SP_OBJxSetFocus(otbCV)
                        End If
                        Exit Sub
                End If
            End If
    End Select
    
    'Check refumd/payment amt
    cPay = SP_VALcText2Double(otbAmt.Text)
    If cPay <= 0 Then 'check amount zero
        Call SP_MSGnShowing("tms_CN005064", nCS_Error)
        bW_KeyEnter = False
        Call SP_OBJxSetFocus(otbAmt)
        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
        Exit Sub
    End If
    '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
    If Trim(otbTel.Text) = "" Then
        Call SP_MSGnShowing("tms_CN005248", nCS_Warning)
        Call SP_OBJxSetFocus(otbTel)
        Exit Sub
    End If
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
            If cPay > cVB_Balance Then
                Call SP_MSGnShowing("tms_CN005232", nCS_Error)
                bW_KeyEnter = False
                Call SP_OBJxSetFocus(otbAmt)
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            End If
        Case tEN_TrnAddDepositCancel, tEN_TrnDepositCancel
            If cPay > cVB_Balance Then
                Call SP_MSGnShowing("tms_CN005232", nCS_Error)
                bW_KeyEnter = False
                Call SP_OBJxSetFocus(otbAmt)
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            End If
        Case Else
             If bW_FoundAcc = False Then
                   If Not W_PRCxFindCrdByCVNum(tVB_CNCVNo) Then
                         '1 : use 2 : un-use 3 : lose 4 : pickup
                       Select Case tW_CVStatusCode
                           Case "1"
                               'none
                           Case "2"
                               Call SP_MSGnShowing("tms_CN005229", nCS_Error)
                           Case "3"
                               Call SP_MSGnShowing("tms_CN005230", nCS_Error)
                           Case "4"
                               Call SP_MSGnShowing("tms_CN005231", nCS_Error)
                           Case Else
                               Call SP_MSGnShowing("tms_CN005228", nCS_Error)
                       End Select
                       bW_FoundAcc = False
                       otbCV.Text = ""
                       orbMonth.Text = ""
                       orbYear.Text = ""
                       otbAmt.Text = SP_FMTcCurPoint(0, False, True)
                       otbTel.Text = "" '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
                       Call SP_OTBxSelText(otbCV)
                       If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                       Exit Sub
                   Else
                        If cVB_Balance >= cW_CVAmt Then
                            otbAmt.Text = SP_FMTcCurPoint(cW_CVAmt, False, True)
                        Else
                            'otbAmt.Text = SP_FMTcCurPoint(cVB_Balance, False, True)'TK-ComSheet-001
                            otbAmt.Text = SP_FMTcCurPoint(cW_CVAmt, False, True)
                        End If
                        orbMonth.Text = Mid(Format(dW_CVExpDate, "dd/MM/yyyy"), 4, 2)
                        orbYear.Text = Mid(Format(dW_CVExpDate, "dd/MM/yyyy"), 9, 2)
                       bW_FoundAcc = True
                   End If
            End If
            'ตรวจสอบ Balance ในตาราง CV
            'If cPay > cW_CVAmt Or cPay > cVB_Balance Then 'TK-ComSheet-001
            If cPay > cW_CVAmt Then
                Call SP_MSGnShowing("tms_CN005232", nCS_Error)
                bW_KeyEnter = False
                Call SP_OBJxSetFocus(otbAmt)
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                If cVB_Balance >= cW_CVAmt Then
                    otbAmt.Text = SP_FMTcCurPoint(cW_CVAmt, False, True)
                Else
                    'otbAmt.Text = SP_FMTcCurPoint(cVB_Balance, False, True) 'TK-ComSheet-001
                    otbAmt.Text = SP_FMTcCurPoint(cW_CVAmt, False, True)
                End If
                Exit Sub
            End If
            
            '//////////////
            If aVB_TenderMnu(nVB_CurTenderSeqNo).StaMallCard = False Then
                oDisc.dTranDate = tUT_SaleDate
                oDisc.tTerminalNo = tVB_CNTerminalNum
                oDisc.tTransNo = tVB_Running
                Select Case SP_VALnText2Double(tVB_TransactionSub)
                    Case tEN_TrnRedeem
                        oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtRedeem, False, True)
                    Case tEN_TrnAddDeposit
                        oDisc.cBalance = SP_FMTcCurPoint(cVB_DepossitAmtBAddDpt, False, True)
                End Select
                oDisc.cDiscValue = SP_VALcText2Double(otbAmt.Text)
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
            End If
            '////////////

          'TK-ComSheet-001
          If (cPay < cW_CVAmt Or cVB_Balance < cW_CVAmt) And aVB_TenderMnu(nVB_CurTenderSeqNo).StaChange = True Then
                'เกิด Re-Issue
               tVB_CNCVNoReIssue = W_GETtAutoCVNum
               dVB_CNCVDateReIssue = DateSerial(Year(Format(CDate(tUT_SaleDate), "yyyy/MM/dd")) + 1, Month(Format(CDate(tUT_SaleDate), "yyyy/MM/dd")), 1)
               'cVB_CNCVAmtReIssue = SP_FMTcCurPoint(cW_CVAmt - cPay, True, False) 'TK-ComSheet-001
               cVB_CNCVAmtReIssue = SP_FMTcCurPoint(cW_CVAmt, True, False) - SP_FMTcCurPoint(cVB_Balance, True, False) '*TON 58-08-14 ลดความผิดพลาดเมื่อ cW_CVAmt และ cVB_Balance ทศนิยมไม่เท่ากัน CommSheet TK-ISS3000-590
               Call SP_CRTxLogMallCardExtend(("Date:" & Date & " Time:" & Time$ & " tVB_Transaction:" & tVB_Transaction) & "CV cW_CVAmt:" & SP_FMTcCurPoint(cW_CVAmt, True, False) & ",cVB_Balance:" & SP_FMTcCurPoint(cVB_Balance, True, False) & ",cVB_CNCVAmtReIssue(Actual):" & cVB_CNCVAmtReIssue & ",cVB_CNCVAmtReIssue(Round):" & SP_FMTcCurPoint(cVB_CNCVAmtReIssue, True, False)) '*Bump 63-02-18 Log-ComSheet719
               tVB_CNCVTelReIssue = Trim(otbTel.Text) '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
          End If
            
    End Select
    cVB_Pay = cPay
    cVB_CNCVAmt = cW_CVAmt
    Unload Me
    Exit Sub
ErrHandle:
    bW_FoundAcc = False
    bW_KeyEnter = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_OBJxSetFocus(otbCV)
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
End Sub

Private Function W_GETtAutoCVNum() As String
'--------------------------------------
'Call:
'Cmt: Generate Credit Voucher Number(DDDTTTTTNNNNN)
'- DDD is day of year
'- TTTTT is Transaction Number
'- NNNNN is Terminal Number
'--------------------------------------
On Error GoTo ErrHandle
    Dim tCVNum As String
    Dim tDDD As String
    Dim dDate1 As Date
    Dim dDate2 As Date
    'dDate1 = CDate(tUT_SaleDate)  '*Bump 56-02-06 Running หมายเลข CV ให้ช้ Systemdate
    dDate1 = CDate(DateValue(Now))
    dDate2 = DateSerial(Year(Format(dDate1, "yyyy/MM/dd")), 1, 1)
    tDDD = Format(CStr(DateDiff("d", dDate2, dDate1) + 1), "000")
    'W_GETtAutoCVNum = tDDD & tVB_Running & tVB_CNTerminalNum
    W_GETtAutoCVNum = tDDD & tVB_CNTerminalNum & tVB_Running      '*Bump 56-02-19 CommSheet TK-ISS3000-042 00.00.02.xls
    Exit Function
ErrHandle:
    W_GETtAutoCVNum = ""
End Function

Private Sub orbMonth_GotFocus()
    orbMonth.SelStart = 0
    orbMonth.SelLength = Len(orbMonth.Text)
End Sub

Private Sub orbMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            orbMonth.Text = ""
    End Select
End Sub

Private Sub orbMonth_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            SendKeys "{Tab}"
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
            If (CLng(Val(Format(orbMonth.Text))) * 10) + CLng(Val(Format(Chr(KeyAscii)))) > 12 Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub orbYear_GotFocus()
    orbYear.SelStart = 0
    orbYear.SelLength = Len(orbYear.Text)
End Sub

Private Sub orbYear_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            orbYear.Text = ""
    End Select
End Sub

Private Sub orbYear_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            SendKeys "{Tab}"
'            DoEvents                '*KT 53-01-14
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub orbYear_LostFocus()
    If Format(orbMonth.Text, "00") = "00" Then
        Call SP_MSGnShowing("tms_CN005227", nCS_Warning)
        orbMonth.Text = ""
        orbYear.Text = ""
        Call SP_OBJxSetFocus(orbMonth)
        Exit Sub
    End If
    
    If Format(orbYear.Text, "00") = "00" Then
        Call SP_MSGnShowing("tms_CN005227", nCS_Warning)
        orbMonth.Text = ""
        orbYear.Text = ""
        Call SP_OBJxSetFocus(orbMonth)
        Exit Sub
    End If
    
    '__________________________________________
    'Check Expire  Date
'    dVB_CrdDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)
'    If Format(dVB_CrdDateExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then
'        Call SP_MSGnShowing("tms_CN005227", nCS_Warning)
'        orbMonth.Text = ""
'        orbYear.Text = ""
'        Call SP_OBJxSetFocus(orbMonth)
'        Exit Sub
'    End If
    
    
End Sub


Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle

    With oScanner
        'otbCV.Text = .ScanData '*Bump 56-02-28 Comment บรรทัดนี้
        '*Bump 56-02-28 เพิ่ม code 8 บรรทัด
        If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
            .DecodeData = True
        End If
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            otbCV.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel) '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
        Else
            otbCV.Text = .ScanData
        End If
        If otbCV.Text <> "" Then
            If Asc(Right(otbCV.Text, 1)) = 13 Then
                otbCV.Text = Mid(otbCV.Text, 1, Len(otbCV.Text) - 1)
            End If
            Call otbCV_KeyDown(13, 0)
        End If
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
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description
     If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
End Sub

Private Sub otbAmt_Change()
    otbAmt.Text = SP_DEFtChangeNum(Trim(otbAmt.Text))
    otbAmt.SelStart = Len(otbAmt.Text)
End Sub

Private Sub otbAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbAmt.Text = ""
    End Select
End Sub

Private Sub otbAmt_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0                    '*AO51-03-12   Comm.Sheet SF-070
'            If InStr(Trim(otbAmount.Text), ".") > 0 Then
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

Private Sub otbCV_GotFocus()
    otbCV.SelStart = 0
    otbCV.SelLength = Len(otbCV.Text)
End Sub

Private Sub otbCV_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim nI As Long
On Error GoTo ErrHandle
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbCV.Text = ""
        Case 13
            tVB_CNCVNo = Trim(otbCV.Text)
            Select Case Val(Format(tVB_TransactionSub))
                Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                    bW_FoundAcc = True
                    SendKeys "{Tab}"
                Case tEN_TrnAddDepositCancel, tEN_TrnDepositCancel
                    bW_FoundAcc = True
                    SendKeys "{Tab}"
                Case Else 'sale,layaway
                        If Not W_PRCxFindCrdByCVNum(tVB_CNCVNo) Then
                            '1 : use 2 : un-use 3 : lose 4 : pickup
                           Select Case tW_CVStatusCode
                               Case "1"
                                   'none
                               Case "2"
                                   Call SP_MSGnShowing("tms_CN005229", nCS_Error)
                               Case "3"
                                   Call SP_MSGnShowing("tms_CN005230", nCS_Error)
                               Case "4"
                                   Call SP_MSGnShowing("tms_CN005231", nCS_Error)
                               Case Else
                                   Call SP_MSGnShowing("tms_CN005228", nCS_Error)
                           End Select
                           bW_FoundAcc = False
                           otbCV.Text = ""
                           orbMonth.Text = ""
                           orbYear.Text = ""
                           otbAmt.Text = SP_FMTcCurPoint(0, False, True)
                           otbTel.Text = "" '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
                           Call SP_OTBxSelText(otbCV)
                       Else
                            If cVB_Balance >= cW_CVAmt Then
                                otbAmt.Text = SP_FMTcCurPoint(cW_CVAmt, False, True)
                            Else
                                'otbAmt.Text = SP_FMTcCurPoint(cVB_Balance, False, True)
                                otbAmt.Text = SP_FMTcCurPoint(cW_CVAmt, False, True) 'TK-ComSheet-001
                            End If
                            orbMonth.Text = Mid(Format(dW_CVExpDate, "dd/MM/yyyy"), 4, 2)
                            orbYear.Text = Mid(Format(dW_CVExpDate, "dd/MM/yyyy"), 9, 2)
                            otbTel.Text = tW_CVTel '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
                            bW_FoundAcc = True
'                           SendKeys "{Tab}"
                            Call W_PRCxLockText(True)               '*Ao 2013/03/16
'                            Call SP_OBJxSetFocus(ocmOk)
                            Call SP_OTBxSelText(otbTel) '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
                        End If
            End Select
    End Select
    Exit Sub
ErrHandle:
    
End Sub

Private Function W_PRCxFindCrdByCVNum(ByVal ptCVNum$) As Boolean
    Dim tSql As String
    Dim orsCV As ADODB.Recordset
On Error GoTo ErrHandle
    W_PRCxFindCrdByCVNum = False
    tSql = "SELECT * FROM TPSTCV   WITH(NOLOCK)"
'    tSql = tSql & vbCrLf & " WHERE FTCVNum='" & ptCVNum & "'"
    tSql = tSql & vbCrLf & " WHERE FTCVFlag='0' AND FTCVNum='" & ptCVNum & "'"   '*Eaw 56-06-18 เพิ่มเงื่อนไขให้ FTCVFlag =0
    If SP_SQLvExecute(tSql, orsCV, oDB.ocnOnLine) = 0 Then 'Online only
        If Not (orsCV.BOF And orsCV.EOF) Then
              '1 : use 2 : un-use 3 : lose 4 : pickup
            tW_CVStatusCode = SP_FEDtChkString(orsCV, "FTCVStatusCode")
            cW_CVAmt = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsCV, "FCCVBalAmt"), True, False)
            dW_CVDate = SP_FEDdChkDate(orsCV, "FDCVIssueDate")
            dW_CVExpDate = SP_FEDdChkDate(orsCV, "FDCVExpDate")
            tW_CVTel = SP_FEDtChkString(orsCV, "FTCustContact") '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
            Select Case tW_CVStatusCode
                Case "1"
'                    W_PRCxFindCrdByCVNum = True
                    '*Eaw 56-11-04 ถ้าหา CV ไม่เจอในตางราง TPSTCV ให้ไปหาในตางราง SaleRC ใน OFF-Line ด้วยทำไว้กรณีที่ ServiceTransfer  Update Status ของ CV ไม่ทัน
                    If W_PRCbFindCVFromRC Then
                        tW_CVStatusCode = "4"
                    Else
                        W_PRCxFindCrdByCVNum = True
                    End If
                Case "2"
                    
                Case "3"
                    
                Case "4"
                    
            End Select
        Else
            tW_CVStatusCode = "5"   '*Eaw 56-06-18
        End If
    Else
        tW_CVStatusCode = "5"  '*Eaw 56-06-18
    End If
    Call SP_DATxRsNothing(orsCV)
    Exit Function
ErrHandle:
    W_PRCxFindCrdByCVNum = False
     Call SP_DATxRsNothing(orsCV)
     tW_CVStatusCode = "5" '*Eaw 56-06-18
End Function

Private Function W_PRCxFindCrdByCVNumEx(ByVal ptCVNum$) As Boolean
    Dim tSql As String
    Dim orsCV As ADODB.Recordset
On Error GoTo ErrHandle
    W_PRCxFindCrdByCVNumEx = True
    tSql = "SELECT * FROM TPSTCV  WITH(NOLOCK)"
    tSql = tSql & vbCrLf & " WHERE FTCVNum=' " & ptCVNum & "'"
    If SP_SQLvExecute(tSql, orsCV, oDB.ocnOnLine) = 0 Then 'Online only
        If (orsCV.BOF And orsCV.EOF) Then
            W_PRCxFindCrdByCVNumEx = False
        End If
    End If
    Call SP_DATxRsNothing(orsCV)
    Exit Function
ErrHandle:
    W_PRCxFindCrdByCVNumEx = True
     Call SP_DATxRsNothing(orsCV)
End Function

Private Sub otbCV_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Call SP_ChkKeyAsciiNumber(KeyAscii)
End Sub

Private Sub otbCV_LostFocus()
  Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
            'none
        Case tEN_TrnAddDepositCancel, tEN_TrnDepositCancel
            'none
        Case Else 'sale,layaway
            otbCV.Text = SP_TXTtTrimRightLeft(otbCV.Text)
             If bW_FoundAcc = False Then
                    tVB_CNCVNo = Trim(otbCV.Text)
                    If tVB_CNCVNo <> "" Then
'                        If Not W_PRCxFindCrdByCVNum(tVB_CNCVNo) Then
'                             '1 : use 2 : un-use 3 : lose 4 : pickup
'                            Select Case tW_CVStatusCode
'                                Case "1"
'
'                                Case "2"
'                                    Call SP_MSGnShowing("tms_CN005229", nCS_Error)
'                                Case "3"
'                                    Call SP_MSGnShowing("tms_CN005230", nCS_Error)
'                                Case "4"
'                                    Call SP_MSGnShowing("tms_CN005231", nCS_Error)
'                                Case Else
'                                    Call SP_MSGnShowing("tms_CN005228", nCS_Error)
'                            End Select
'
'                            bW_FoundAcc = False
'                            otbCV.Text = ""
'                            orbMonth.Text = ""
'                            orbYear.Text = ""
'                        Else
'                            bW_FoundAcc = True
'                        End If
                        '*Eaw 56-10-29 เปลี่ยน process ใหม่ให้ไปหาข้อมูล CV จาก TPST ใน LostFocus ด้วยกรณีที่นพักงานไม่กด Enter จาก CV Num
                        If Not W_PRCxFindCrdByCVNum(tVB_CNCVNo) Then
                            '1 : use 2 : un-use 3 : lose 4 : pickup
                           Select Case tW_CVStatusCode
                               Case "1"
                                   'none
                               Case "2"
                                   Call SP_MSGnShowing("tms_CN005229", nCS_Error)
                               Case "3"
                                   Call SP_MSGnShowing("tms_CN005230", nCS_Error)
                               Case "4"
                                   Call SP_MSGnShowing("tms_CN005231", nCS_Error)
                               Case Else
                                   Call SP_MSGnShowing("tms_CN005228", nCS_Error)
                           End Select
                           bW_FoundAcc = False
                           otbCV.Text = ""
                           orbMonth.Text = ""
                           orbYear.Text = ""
                           otbAmt.Text = SP_FMTcCurPoint(0, False, True)
                           otbTel.Text = ""
                           Call SP_OTBxSelText(otbCV)
                       Else
                            If cVB_Balance >= cW_CVAmt Then
                                otbAmt.Text = SP_FMTcCurPoint(cW_CVAmt, False, True)
                            Else
                                otbAmt.Text = SP_FMTcCurPoint(cW_CVAmt, False, True)
                            End If
                            orbMonth.Text = Mid(Format(dW_CVExpDate, "dd/MM/yyyy"), 4, 2)
                            orbYear.Text = Mid(Format(dW_CVExpDate, "dd/MM/yyyy"), 9, 2)
                            otbTel.Text = tW_CVTel
                            bW_FoundAcc = True
                            Call W_PRCxLockText(True)
                            Call SP_OTBxSelText(otbTel)
                        End If
                    '--------------------------------------------------------------
                    Else
                        bW_FoundAcc = True
                    End If
            End If
    End Select
    
End Sub
Private Sub otbTel_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case 13
            Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub
