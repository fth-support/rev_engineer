VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNIPV 
   Caption         =   "IPV or Void Transaction;klg_Title"
   ClientHeight    =   3225
   ClientLeft      =   5220
   ClientTop       =   3420
   ClientWidth     =   5130
   ControlBox      =   0   'False
   Icon            =   "wCNIPV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5130
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2535
      Index           =   2
      Left            =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   5085
      _Version        =   720897
      _ExtentX        =   8969
      _ExtentY        =   4471
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbTerminal 
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
         MaxLength       =   5
         TabIndex        =   0
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox otbAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   11
         TabIndex        =   2
         Text            =   "0"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox otbTrans 
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
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No:;klg_TransNo"
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
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1080
         Width           =   1860
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal No:;klg_TerminalNo"
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
         Width           =   1890
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :;klg_Amount"
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
         Width           =   1800
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   30
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2025
      Width           =   5085
      _Version        =   720897
      _ExtentX        =   8969
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
         Left            =   3360
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton ocmOk 
         Caption         =   "Ok;klg_Ok"
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
         Left            =   1800
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNIPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'การยกเลิกรายการ(IPV or VOID Transaction) เป็นกายกเลิกรายการขายและรายการคืน กรณียังไม่ปิดการขายประจำวัน เมื่อพนักงานเก็บเงินบันทึกข้อมูลการขายหรือการคืนผิดพลาด
'โดยลูกค้าจะต้องนำบิลของรายการขายหรือรายการคืนมาด้วยและพนักงานเก็บเงิน ต้องบันทึกหมายเลขเครือง และหมายเลขลำดับการขายนั้นๆ ระบบจะดึงข้อมุลการขายหรือการคืนขึ้นมาเพื่อให้พนักงานเก็บเงินยืนยันการยกเลิก
'รายการ และระบบจะยกเลิกการบันทึกทั้งหมดของรายการที่ยกเลิกนั้น

Private Sub W_FRMxInitail()
    bVB_IPV = False
    otbTerminal.Text = ""
    otbTrans.Text = ""
    otbAmount.Text = 0
End Sub

'Private Sub Form_Activate()
'On Error GoTo ErrHandle
'    Call SP_OTBxSelText(otbTerminal)
'    Exit Sub
'ErrHandle:
'End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then SendKeys "{TAB}" '*KT 53-01-19 ทำให้ขั้นตอนการ Setfocus ตอน Load ผิดพลาด ถ้ามีการใส่หน้าจอ Login มา
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxInitail
    bVB_FormPayAct = True 'PH 2.0 ML-SF-019 ไม่ต้องการให้ทำที่ Form_Activate ของ wMain
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

'Private Sub Form_Unload(Cancel As Integer)
'    Set wCNIPV = Nothing
'End Sub

Private Sub ocmCancel_Click()
    bVB_IPV = False
    Unload Me
End Sub


Private Sub ocmOK_Click()
    '*KT 51-10-31
    '*KT 53-05-29 แก้ไขการเกิด Dot send โดยการแก้ไขจากตัวตังต้น 3.0.19
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tTable$
    Dim tWhe As String
'PH 2.0.0 OLD
'    '*KT 51-09-2008 ป้องกันไม่ให้ทำการ IPV เมื่อ Terminal Offline
'    '-----------------------------------------------------------------------------------------
'    If ocnAll.ConnectionString = oDB.ocnOffLine.ConnectionString Then
'            If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
'                If Not SP_SETxInitialToOnline Then
'                        MsgBox "Terminal is Offline"
'                        bVB_IPV = False
'                        Unload Me
'                        Exit Sub
'                End If
'            Else
'                MsgBox "Terminal is Offline"
'                bVB_IPV = False
'                Unload Me
'                Exit Sub
'            End If
'    End If
'    '--------------------------------------------------------------------------------------
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False     '*Ao 10-02-09
    If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
        If Not SP_CHKbOnlineEnable Then
                MsgBox "Terminal is Offline", vbCritical 'tms_CN005222
                Call ocmCancel_Click
                Exit Sub
        End If
    Else
        MsgBox "Terminal is Offline", vbCritical 'tms_CN005222
        Call ocmCancel_Click
        Exit Sub
    End If
    
    If otbTerminal.Text = "" Then
        Call SP_MSGnShowing("tms_CN005089", nCS_Error)
        Call SP_OBJxSetFocus(otbTerminal)
        Set orsTemp = Nothing
        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True    '*Ao 10-02-09
        Exit Sub
    End If
    If otbTrans.Text = "" Then
        Call SP_MSGnShowing("tms_CN005089", nCS_Error)
        Call SP_OBJxSetFocus(otbTrans)
        Set orsTemp = Nothing
        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True     '*Ao 10-02-09
        Exit Sub
    End If
    If CDbl(Val(Format(otbAmount.Text))) < 0 Then
         Call SP_MSGnShowing("tms_CN005090", nCS_Error)
        Call SP_OBJxSetFocus(otbAmount)
        Set orsTemp = Nothing
        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True     '*Ao 10-02-09
        Exit Sub
    Else
       '*KT 51-09-26
       ' If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbAmount.Text))) Then Exit Sub       '*AO51-03-26   Comm.Sheet  ML-SF-100
    End If
    
    'บิลที่นำมาคืนต้องเป็นบิลขายและคืนเท่านัน้ และไม่เคยมีการ Void
    tSql = "SELECT * FROM " & "TPSTSalHD WITH(NOLOCK) "
    tSql = tSql & vbCrLf & "  WHERE  FTTmnNum='" & Trim(otbTerminal.Text) & "' AND FTShdTransNo='" & Trim(otbTrans.Text) & "'"
    If bVB_TrainingMode Then
       tSql = tSql & vbCrLf & "  AND   FTShdTransType IN ('" & Format(tEN_TrnTraining, "00") & "')"
    ElseIf bVB_ReEntryMode = True Then
        Select Case tVB_CNMnuCurrent
            Case "172003"
                tSql = tSql & vbCrLf & "  AND   FTShdTransType IN ('" & Format(tEN_TrnSaleTaxInc, "00") & "','" & Format(tEN_TrnReturnTaxInc, "00") & "' )"
            Case "172103"
                tSql = tSql & vbCrLf & "  AND   FTShdTransType IN ('" & Format(tEN_TrnSaleManualTax, "00") & "','" & Format(tEN_TrnReturnManualTax, "00") & "' )"
            Case Else
                'none
        End Select
    Else
         tSql = tSql & vbCrLf & "  AND   FTShdTransType IN ('" & Format(tEN_TrnSale, "00") & "'" '03
         tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnReturn, "00") & "'" '04
         tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnRedeem, "00") & "'" '07
         tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnVIPSale, "00") & "'" '10
         tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnTaxExceptSale, "00") & "'" '11
         tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnDutyFreeSale, "00") & "'" '13
         tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnOtherReceive, "00") & "'" '22
         tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnOtherDisbursement, "00") & "'" '23
         tSql = tSql & vbCrLf & ")"
    End If
    
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then  'PH 2.0.0 ติดต่อฐานข้อมูล server เท่านั้น
        Call SP_MSGnShowing("tms_CN005087", nCS_Error)
        Call W_FRMxInitail
        Call SP_OBJxSetFocus(otbTerminal)
        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True     '*Ao 10-02-09
        Exit Sub
    End If
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        If SP_FEDtChkString(orsTemp, "FTShdStaCancel") = "Y" Then
            Call SP_MSGnShowing("tms_CN005088", nCS_Error)
            Call W_FRMxInitail
            Call SP_OBJxSetFocus(otbTerminal)
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True     '*Ao 10-02-09
            Exit Sub
        Else
             cVB_AmountIPVDisc = SP_FEDvChkNumeric(orsTemp, "FCShdDis") '*KT 53-02-26
             tVB_MemPointIDTmp = SP_FEDtChkString(orsTemp, "FTCstCardCode") '*KT 52-04-09 เก็บค่า Member ID ที่จะได้ Point เพื่อไปคำนวน Point เมื่อจบการ Redeem
             tVB_TransTypeIPV = SP_FEDtChkString(orsTemp, "FTShdTransType")
             Select Case Val(tVB_TransTypeIPV)
                Case tEN_TrnRedeem
                    If CDbl(SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCShdDepositPrvBal") + SP_FEDvChkNumeric(orsTemp, "FCShdRnd"), True, False)) <> CDbl(SP_FMTcCurPoint(otbAmount.Text, True, False)) Then
                        Call SP_MSGnShowing("tms_CN005091", nCS_Error)
                        Call W_FRMxInitail
                        Call SP_OBJxSetFocus(otbTerminal)
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True     '*Ao 10-02-09
                        Exit Sub
                    Else
                        bVB_IPV = True
                        tVB_TerminalIPV = SP_FEDtChkString(orsTemp, "FTTmnNum")
                        tVB_TransNoIPV = SP_FEDtChkString(orsTemp, "FTShdTransNo")
                        tVB_TransTypeIPV = SP_FEDtChkString(orsTemp, "FTShdTransType")
                        cVB_AmountIPV = SP_FMTcCurPoint(otbAmount.Text, True, False)
                        dVB_TransDateIPV = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
                        'ปรับประเภท Transaction ของการ Void ทั้งบิลให้ขึ้นกับใบที่อ้างอิงการทำ Void
                        Select Case Val(tVB_TransTypeIPV)
                            Case tEN_TrnRedeem 'Redeem & IPV Redeem
                                tVB_Transaction = aVB_TransMode(16).ID
                                tVB_TransactionSub = tVB_Transaction
                                tVB_TransName = aVB_TransMode(16).Name
                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVRedeem - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVRedeem - 1).UsrLangSec
                                tVB_CNTypePrinter = aVB_TransMode(16).PrintType
                        End Select
                        wMain.olaTransName.Caption = tVB_TransName
'                        DoEvents        '*KT 53-01-14
                        Call SP_INSxTransHD(tVB_CNTblHD)
                        tTable = tVB_CNTblHD
                        Call wFunctionKB.W_DupTblBtwSleUpdHoldHDIPV(orsTemp, tTable) '*KT 51-11-10 ML-SB-634
                        bVB_SetFocus = True
                      End If
                Case Else
                    If CDbl(SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCShdGrand") + SP_FEDvChkNumeric(orsTemp, "FCShdRnd"), True, False)) <> CDbl(SP_FMTcCurPoint(otbAmount.Text, True, False)) Then
                        Call SP_MSGnShowing("tms_CN005091", nCS_Error)
                        Call W_FRMxInitail
                        Call SP_OBJxSetFocus(otbTerminal)
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True     '*Ao 10-02-09
                        Exit Sub
                    Else
                        bVB_IPV = True
                        tVB_TerminalIPV = SP_FEDtChkString(orsTemp, "FTTmnNum")
                        tVB_TransNoIPV = SP_FEDtChkString(orsTemp, "FTShdTransNo")
                        tVB_TransTypeIPV = SP_FEDtChkString(orsTemp, "FTShdTransType")
                        cVB_AmountIPV = SP_FMTcCurPoint(otbAmount.Text, True, False)
                        dVB_TransDateIPV = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
                        'ปรับประเภท Transaction ของการ Void ทั้งบิลให้ขึ้นกับใบที่อ้างอิงการทำ Void
                        Select Case Val(tVB_TransTypeIPV)
                            Case tEN_TrnSale 'Sale Transaction & IPV Transaction
                                tVB_Transaction = aVB_TransMode(4).ID
                                tVB_TransactionSub = tVB_Transaction
                                tVB_TransName = aVB_TransMode(4).Name
                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPV - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPV - 1).UsrLangSec
                                tVB_CNTypePrinter = aVB_TransMode(4).PrintType
                            Case tEN_TrnVIPSale, tEN_TrnCreateVIP 'VIP Service Sale & IPV VIP Service Sale
                                tVB_Transaction = aVB_TransMode(25).ID
                                tVB_TransactionSub = tVB_Transaction
                                tVB_TransName = aVB_TransMode(25).Name
                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVVIPService - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVVIPService - 1).UsrLangSec
                                tVB_CNTypePrinter = aVB_TransMode(25).PrintType
                            Case tEN_TrnTaxExceptSale 'Tax Except Sale & IPV Tax Except Sale
                                tVB_Transaction = aVB_TransMode(26).ID
                                tVB_TransactionSub = tVB_Transaction
                                tVB_TransName = aVB_TransMode(26).Name
                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVTaxExceptSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVTaxExceptSale - 1).UsrLangSec
                               tVB_CNTypePrinter = aVB_TransMode(26).PrintType
                            Case tEN_TrnDutyFreeSale 'Duty Free Sale & IPV Duty Free Sale
                                tVB_Transaction = aVB_TransMode(27).ID
                                tVB_TransactionSub = tVB_Transaction
                                tVB_TransName = aVB_TransMode(27).Name
                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVDutyFreeSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVDutyFreeSale - 1).UsrLangSec
                                tVB_CNTypePrinter = aVB_TransMode(27).PrintType
                            Case tEN_TrnReturn 'Return Transaction
                                tVB_Transaction = aVB_TransMode(13).ID
                                tVB_TransactionSub = tVB_Transaction
                                tVB_TransName = aVB_TransMode(13).Name
                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVReturn - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVReturn - 1).UsrLangSec
                                tVB_CNTypePrinter = aVB_TransMode(13).PrintType
                                
                                '*KT 66-08-17 TTK-ISS3000-839
                                'Refer  Sale Bill
                                tVB_TerminalSal = SP_FEDtChkString(orsTemp, "FTShdDepRefTmnNum")
                                tVB_TransNoSal = SP_FEDtChkString(orsTemp, "FTShdDepReTransNo")
                                 '*KT 62-12-12 ML-POSC-0980  ไม่ควร Fix "03" เพราะการ Return ทำการเอกสารเอกสาร Redeem,VIP Sale,Create VIP,TaxExcept,DutyFreeSal
                                 'ลำดับการหา
                                 '1. หาใน Local
                                 '2. หาใน Server
                                 '3. หาใน DB Backup
                                 '4. ไม่เจอใน 1-3 ให้ Fix = 03 (Sale) เพื่อให้ทำ IPV ได้  แต่จะกลับไปทำ Return ต้นฉบับไม่ได้แล้ว
                                 '--------------------------------------------------------------------------------------------------------------------
                                 tWhe = " WHERE FTTmnNum='" & tVB_TerminalSal & "'"
                                 tWhe = tWhe & "  AND FTShdTransNo='" & tVB_TransNoSal & "'"
                                 tVB_TransTypeSal = SP_SQLtFunction("", "FTShdTransType", "TPSTSalHD", tWhe) '*KT 62-12-12 ML-POSC-0980  ไม่ควร Fix "03" เพราะการ Return ทำการเอกสารเอกสาร Redeem,VIP Sale,Create VIP,TaxExcept,DutyFreeSal
                                 If tVB_TransTypeSal = "" Then tVB_TransTypeSal = SP_SQLtFunction("", "FTShdTransType", "TPSTSalHD", tWhe, oDB.ocnOnLine)

                                 If tVB_TransTypeSal = "" Then
                                     Call SP_CRTxWriteLogTerminated("ErrorNumber:0" & " ErrorDesc:" & " Form:" & Me.Name & "Control:ocmOk Events:Click " & " not found bill refer Sale,Redeem,VIP Sale,Create VIP,TaxExcept,DutyFreeSal")
                                     tVB_TransTypeSal = "03"
                                 End If
                                 '--------------------------------------------------------------------------------------------------------------------
                                                                
                            '*KT 51-04-25 Comsheet ML-SF-180
                            Case tEN_TrnOtherReceive 'Other Receive Transacton
                                tVB_Transaction = aVB_TransMode(31).ID
                                tVB_TransactionSub = tVB_Transaction
                                tVB_TransName = aVB_TransMode(31).Name
                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVOtherReceive - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVOtherReceive - 1).UsrLangSec
                                tVB_CNTypePrinter = aVB_TransMode(31).PrintType
                            Case tEN_TrnOtherDisbursement 'Other Disbursement Transacton
                                tVB_Transaction = aVB_TransMode(32).ID
                                tVB_TransactionSub = tVB_Transaction
                                tVB_TransName = aVB_TransMode(32).Name
                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVOtherDisbursement - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVOtherDisbursement - 1).UsrLangSec
                                tVB_CNTypePrinter = aVB_TransMode(32).PrintType
'                            Case tEN_TrnSaleTaxInc
'                                 tVB_Transaction = aVB_TransMode(39).ID
'                                tVB_TransactionSub = tVB_Transaction
'                                tVB_TransName = aVB_TransMode(39).Name
'                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVTaxInc - 1).UsrLangSec
'                                tVB_CNTypePrinter = aVB_TransMode(39).PrintType
'                            Case tEN_TrnReturnTaxInc
'                                tVB_Transaction = aVB_TransMode(40).ID
'                                tVB_TransactionSub = tVB_Transaction
'                                tVB_TransName = aVB_TransMode(40).Name
'                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVReturnTaxInc - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVReturnTaxInc - 1).UsrLangSec
'                                tVB_CNTypePrinter = aVB_TransMode(40).PrintType
'                            Case tEN_TrnSaleManualTax
'                                tVB_Transaction = aVB_TransMode(45).ID
'                                tVB_TransactionSub = tVB_Transaction
'                                tVB_TransName = aVB_TransMode(45).Name
'                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVManualTax - 1).UsrLangSec
'                                tVB_CNTypePrinter = aVB_TransMode(45).PrintType
'                            Case tEN_TrnReturnManualTax
'                                tVB_Transaction = aVB_TransMode(46).ID
'                                tVB_TransactionSub = tVB_Transaction
'                                tVB_TransName = aVB_TransMode(46).Name
'                                tVB_TransNameABB = aVB_NameAbb(tEN_AbbIPVReturnManualTax - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbIPVReturnManualTax - 1).UsrLangSec
'                                tVB_CNTypePrinter = aVB_TransMode(46).PrintType
                        End Select
                         wMain.olaTransName.Caption = tVB_TransName
'                         DoEvents       '*KT 53-01-14
                        Call SP_INSxTransHD(tVB_CNTblHD)
                         tTable = tVB_CNTblHD
                        Call wFunctionKB.W_DupTblBtwSleUpdHoldHDIPV(orsTemp, tTable) '*KT 51-11-10 ML-SB-634
                        bVB_SetFocus = True
                    End If
            End Select
            
            '*KT 55-07-26 takashimaya get data Credit Voucher
            Call SP_DATxRsClose(orsTemp)
            tSql = "SELECT FTCVNum,FTCVNumRef ,FCCVAmt,FCCVBalAmt FROM " & "TPSTCV"            '*Ao 2013/02/21   TK   แก้ไข ให้   IPV  Sale ,Return
            tSql = tSql & vbCrLf & "  WHERE  FTPOSNo='" & Trim(otbTerminal.Text) & "' AND FTTransNo='" & Trim(otbTrans.Text) & "'"
            If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then  'PH 2.0.0 ติดต่อฐานข้อมูล server เท่านั้น
                tVB_CNCVNumRefCancel = SP_FEDtChkString(orsTemp, "FTCVNum")
                tVB_CNCVNumRefRCCancel = SP_FEDtChkString(orsTemp, "FTCVNumRef")
                If tVB_CNCVNumRefCancel = "" Then
                    bVB_CNReIssueCancel = False
                    cVB_CNCVNumRefCancelAmt = 0
                Else
                    bVB_CNReIssueCancel = True
                    cVB_CNCVNumRefCancelAmt = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCCVAmt"), True, False) 'W_GETcTotalCVAmt(SP_FEDtChkString(orsTemp, "FTCVNumRef")) 'SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCCVBalAmt"), True, False)
                End If
            End If
        End If
    Else
        Call SP_MSGnShowing("tms_CN005087", nCS_Error)
        Call W_FRMxInitail
        Call SP_OBJxSetFocus(otbTerminal)
        Exit Sub
    End If
    Call SP_DATxRsNothing(orsTemp)
    Unload Me
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
    '*KT 52-06-05 PH 2.0 ML-SF-020 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
    Dim tID As String
On Error GoTo ErrHandle
    With oScanner
        'tID = .ScanData '*Bump 56-02-28 Comment บรรทัดนี้
        '*Bump 56-02-28 เพิ่ม code 8 บรรทัด
        If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
            .DecodeData = True
        End If
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            tID = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel) '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
        Else
            tID = .ScanData
        End If
'        DoEvents            '*KT 53-01-14
        If tID <> "" Then
            If Asc(Right(tID, 1)) = 13 Then
                tID = Mid(tID, 1, Len(tID) - 1)
            End If
            otbTerminal.Text = Left(tID, 5)
            otbTrans.Text = Right(tID, 5)
            Call SP_OBJxSetFocus(otbAmount)
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
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub otbAmount_Change()          '*AO51-03-12    Comm.Sheet SF- 070
    otbAmount.Text = SP_DEFtChangeNum(Trim(otbAmount.Text))
    otbAmount.SelStart = Len(otbAmount.Text)
End Sub

Private Sub otbAmount_GotFocus()
   otbAmount.SelStart = 0
   otbAmount.SelLength = Len(otbAmount.Text)
End Sub

Private Sub otbAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbAmount.Text = ""
    End Select
End Sub

Private Sub otbAmount_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 6, 8          'backspace
            'none
        Case 46             'ทศนิยม
            KeyAscii = 0
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

Private Sub otbTerminal_GotFocus()
   otbTerminal.SelStart = 0
   otbTerminal.SelLength = Len(otbTerminal.Text)
End Sub

Private Sub otbTerminal_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbTerminal.Text = ""
        Case 13
            If Trim(otbTerminal.Text) <> "" Then
                sendkeys "{TAB}"
            End If
    End Select
End Sub

Private Sub otbTerminal_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
End Sub

Private Sub otbTerminal_LostFocus()
    otbTerminal.Text = SP_TXTtTrimRightLeft(otbTerminal.Text)
End Sub

Private Sub otbTrans_GotFocus()
   otbTrans.SelStart = 0
   otbTrans.SelLength = Len(otbTrans.Text)
End Sub

Private Sub otbTrans_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbTrans.Text = ""
        Case 13
            If Trim(otbTrans.Text) <> "" Then
                sendkeys "{TAB}"
            End If
    End Select
End Sub


Private Sub otbTrans_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            'do nothing
        Case 13
           'do nothing
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbTrans_LostFocus()
    otbTrans.Text = SP_TXTtTrimRightLeft(otbTrans.Text)
End Sub

Private Function W_GETcTotalCVAmt(ByVal ptCVNum As String)
    Dim orsTemp As ADODB.Recordset
    Dim tSql As String
On Error GoTo ErrHandle
    tSql = "SELECT  FCCVAmt FROM TPSTCV"
    tSql = tSql & " WHERE FTCVNum ='" & ptCVNum & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            W_GETcTotalCVAmt = SP_FEDvChkNumeric(orsTemp, "FCCVAmt")
        Else
            W_GETcTotalCVAmt = 0
        End If
    Else
        W_GETcTotalCVAmt = 0
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_GETcTotalCVAmt = 0
    Call SP_DATxRsNothing(orsTemp)
End Function

