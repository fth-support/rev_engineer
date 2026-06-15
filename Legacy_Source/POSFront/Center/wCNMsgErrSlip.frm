VERSION 5.00
Begin VB.Form wCNMsgErrSlip 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   2220
   ClientTop       =   3105
   ClientWidth     =   7215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.Timer otmDrawer 
      Interval        =   500
      Left            =   120
      Top             =   80
   End
   Begin VB.Label olaOposReset 
      BackStyle       =   0  'Transparent
      Caption         =   "Opos Slip printer is reseted"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label olaForm 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6975
   End
   Begin VB.Label olaRemark 
      BackStyle       =   0  'Transparent
      Caption         =   "F12: Esc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Label olaForm 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Out of paper"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   6975
   End
   Begin VB.Label olaForm 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "Message Error"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Shape oshForm 
      BorderWidth     =   8
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "wCNMsgErrSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tW_Msg As String
Public bW_Wait  As Boolean
Public bW_Enter As Boolean
Public bW_Printing As Boolean
Public nW_Line As Long
Public tW_TransTypeSub As String                         '*Ao 2013/08/30

Private nW_CountEmpty As Long
Private nW_Count As Long
Private nW_MsgCount As Integer '*Eaw 56-07-23 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
Private bW_Chkprn As Boolean
Private bW_ChkEnd  As Boolean
Private bW_ChkF12 As Boolean

Private Function SlpRemoval() As Boolean
Dim pMSG As String
Dim MSG As Long
Dim RecPrintable As Long
Dim nCount As Integer

On Error GoTo ErrHandle
    
    nCount = 3
    RecPrintable = 1
    With oOposSlip
        While RecPrintable
            .BeginRemoval 5000
            If (.ResultCode <> OPOS_E_TIMEOUT) Then
                If (.ResultCode = OPOS_E_BUSY) Then
                    SleepEx 1000, 0
                End If
                If (.ResultCode = OPOS_SUCCESS) Then
                    RecPrintable = 0
                End If
            Else
'                MSG = MsgBox("Remove slip", vbOKOnly + vbExclamation, "Print credit sales slip")
'                MSG = MsgBox("Please Remove Slip form", vbOKOnly + vbExclamation, "Slip form")
                If (.ResultCode = OPOS_SUCCESS) Then
                    RecPrintable = 0
                End If
            End If
'            Debug.Print .ResultCode
            nCount = nCount - 1
            If nCount < 1 Then RecPrintable = 0
            If bW_ChkF12 Then
                RecPrintable = 0
                SlpRemoval = True
                Exit Function
            End If
        Wend
        If (.ResultCode <> OPOS_E_ILLEGAL) Then
            RecPrintable = 1
            While RecPrintable
                .EndRemoval
                If (.ResultCode = OPOS_SUCCESS) Then
                    RecPrintable = 0
                    oOposSlip.ClearOutput            '*Ao 2013/07/29
'                    Call SP_CRTxLogPrg("SLIP PRINTER-wCNMsgErrSlip.SlpRemoval-TransType:" & tW_TransTypeSub & ",OPOS_SUCCESS")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                Else
'                    pMSG = W_GETtGetErrorMsg()
'                    MSG = MsgBox(pMSG, vbRetryCancel + vbQuestion, "Print credit sales slip")
'                    MSG = MsgBox("Please Remove Slip form", vbOKOnly + vbExclamation, "Slip form")
                    
'                    Call SP_CRTxLogPrg("SLIP PRINTER-wCNMsgErrSlip.SlpRemoval-TransType:" & tW_TransTypeSub & ",When printing is completed. Remove printed slip form Slip Printer" & ",ResultCode:" & .ResultCode)  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                    
                    DoEvents
                    olaForm(1).Caption = "When printing is completed" & vbCrLf & "Remove printed slip form Slip Printer"
                    'Canceled
                    If (MSG = vbCancel) Then
                        SlpRemoval = False
                        Exit Function
                    End If
                    If bW_ChkF12 Then
                        RecPrintable = 0
                    End If
                    If wCNMsgErrSlip Is Nothing Then
                        RecPrintable = 0
                    End If
                End If
            Wend
        Else
            .EndRemoval
        End If
    End With

    SlpRemoval = True
    Exit Function
ErrHandle:
    SlpRemoval = False
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF12
            bVB_CNCheckCV = True
            otmDrawer.Enabled = False
            bW_ChkF12 = True
'            Unload Me                       '*Ao 2014/04/29
'        Case vbKeyF11                                      '*Ao 2013/07/30      ยังไม่ใช้ 259-ST     F11: Reset opos slip printer,F12: Esc
'            'Reset opos slip
'            oOposSlip.ClosePrint
'            oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
'            oOposSlip.SlpLineChars = nVB_MaxCharSlip
'            'Call SP_MSGnShowing("Reset opos slip printer", nCS_Information)
'            '*Eaw 56-07-23 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
''            MsgBox "Reset opos slip printer", vbInformation, "POSFront"
'            olaOposReset.Visible = True
'            nW_MsgCount = 0
        Case 13
            bW_Enter = True
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    '*Eaw 56-07-23 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
    bW_ChkF12 = False         '*Ao 2014/04/29
    olaOposReset.Visible = False
    nW_MsgCount = 0
    '-----------------------------------------------------------
    olaForm(1).FontSize = 30
    bVB_CNCheckCV = False
    bW_Enter = False
    nW_CountEmpty = 0
    nW_Count = 0
    Call SP_FRMxChildCenter(Me)
    With oOposSlip
        If tW_Msg <> "" Then
            olaForm(0) = tW_Msg
        Else
            olaForm(0) = "Slip Printer"
        End If
        olaForm(1) = POSPrinter1_StatusUpdateEvent(oOposSlip.nGetStatusUpdateEvent)
    End With
    If bW_Wait Then             '*Ao 2013/07/30
        olaForm(1).Caption = "Slip printing in progress"
    End If
    bW_Chkprn = True
    bW_ChkEnd = True
    If nW_Line > 13 Then
        nW_Line = nW_Line - 13
    End If
    Exit Sub
ErrHandle:
    olaForm(1) = Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Dim nI  As Long
'    Dim LF$
    
On Error GoTo ErrHandle
'   LF = Chr(10)
'    With oOposSlip
'        '*KT 51-09-2008 กรณีที่เปิด Printer ให้ทำการตัดกระดาษแล้วพิมพ์ ใหม่
'        '-----------------------------------------------------------------
'         If bVB_ImageLogo Then
'            .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'end print by OPOS
'             .Printnormal PTR_S_RECEIPT, "" & Chr(10)
'              Call oPrn.PrintLogo
'            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL 'end print by OPOS
'            .CutPaper 90
'            .Printnormal PTR_S_RECEIPT, Chr(16) & Chr(20) '*KT 51-12-22 Phase 2 Clear Buffer
'         Else
'            '*KT 51-10-14 V 3.0.135/V 3.0.200 กรณีที่ไม่มีการ Set Logo
'            '--------------------------------------------------------------
'            .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION 'end print by OPOS
'            .Printnormal PTR_S_RECEIPT, "" & Chr(10)
'             For nI = 1 To 8
'                .Printnormal PTR_S_RECEIPT, "" & LF
'             Next nI
'            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL 'end print by OPOS
'            .CutPaper 90
'            .Printnormal PTR_S_RECEIPT, Chr(16) & Chr(20) '*KT 51-12-22 Phase 2 Clear Buffer
'            '---------------------------------------------------------------
'         End If
'        '-----------------------------------------------------------------
'    End With
    tW_Msg = ""
    tW_TransTypeSub = ""
    otmDrawer.Enabled = False
    Call SP_OBJxSetNothing(Me)
    Exit Sub
ErrHandle:
    
End Sub

Private Sub otmDrawer_Timer()
    Dim nStatus As Long
    Dim nLevel As Long '*Eaw 56-07-23 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
On Error GoTo ErrHandle
    '*Eaw 56-07-23 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
     If olaOposReset.Visible Then nW_MsgCount = nW_MsgCount + 1
     If nW_MsgCount = 10 And olaOposReset.Visible Then
        olaOposReset.Visible = False
        nW_MsgCount = 0
     End If
     '--------------------------------------------------------------
    olaForm(1).FontSize = 30
    With oOposSlip
        nW_Count = nW_Count + 1
'        nStatus = oOposSlip.nGetStatusUpdateEvent
        If oOposSlip.SlpEmpty = True Then                               '*Ao 2013/09/16
            nStatus = PTR_SUE_SLP_EMPTY
        Else
            nStatus = PTR_SUE_SLP_PAPEROK
        End If
        
        If bW_Wait Then
            Select Case nStatus
                Case PTR_SUE_SLP_EMPTY
'                    bW_Wait = False
'                    bW_Printing = False
'                    Unload Me
'                    Exit Sub
                    If nW_Count > nW_Line Then          '*Ao 2013/07/30      ใช้กระดาษสั้นทดสอบ พอเปลื่ยนกระดาษ Form จะ Unload
'                        If bW_Enter Then
'                            bW_Wait = False                             'เช็คถ้าจำนวนบรรทัดยังไม่มากกว่า ยังไม่ต้อง Unload
'                            bW_Printing = False
'                            Unload Me
'                            Exit Sub
'                        Else
'                            olaForm(0).Caption = "Slip printing in progress"
'                            olaForm(1).Caption = "Press Enter to Continued"
'                        End If
                        bW_Wait = False
                        bW_Printing = False
                        Unload Me
                        Exit Sub
                    Else
                        Call SleepEx(7000, 0)               'เพื่อรอ เปลื่ยนกระดาษ    จาก 6000 --> 7000
                    End If
                Case PTR_SUE_SLP_PAPEROK
                    If nW_Count > nW_Line Then bW_Printing = False
                    If bW_Printing Then
                        DoEvents
                        olaForm(1).Caption = "Slip printing in progress"  '*Eaw 56-06-27 CommSheet TK-ISS3000-247-TR.xlsx
                        If bW_Chkprn Then
                            oOposSlip.ClearOutput            '*Ao 2013/07/29
                            bW_Chkprn = False
                        End If
                    Else
'                        olaForm(0).Caption = "Please press RELEASE at Slip Printer"
'                        olaForm(1).Caption = "when printing is completed" & vbCrLf & "Remove printed slip form Slip Printer"
                        '*Eaw 56-07-23 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
'                        nLevel = oOposSlip.nCheckHealth(OPOS_CH_INTERNAL)
'                        If nLevel = OPOS_E_ILLEGAL Then
'                            Unload Me
'                            wCNMsgSlip.Show vbModal
'                        Else
'                            olaForm(0).Caption = "Please press RELEASE at Slip Printer"
'                            olaForm(1).Caption = "when printing is completed" & vbCrLf & "Remove printed slip from Slip Printer"
''                            olaForm(0).Caption = "Slip printing in progress"
''                            olaForm(1).Caption = "Please Remove printed slip from Slip Printer"
'                            DoEvents
'                        nLevel = oOposSlip.nCheckHealth(OPOS_CH_INTERNAL)
'                        If nLevel <> 0 Then
'                            MsgBox "nLevel = " & nLevel
'                            oOposSlip.ClosePrint
'                            oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
'                            oOposSlip.SlpLineChars = nVB_MaxCharSlip
'                        End If
'                        If oOposSlip.SlpEmpty = True Then                           '*Ao 2013/09/04
'                            bW_Wait = False
'                            bW_Printing = False
'                            Unload Me
'                            Exit Sub
'                        End If
                        If bW_ChkF12 Then
                            otmDrawer.Enabled = False               '*Ao 2014/04/29  PRJ10001TK-PH2-External Spec-Auto Release Paper at Slip Printer-00.02.03 FSBT
                            bW_Wait = False
                            bW_Printing = False
                            Unload Me
                            Exit Sub
                        End If
                        If SlpRemoval() = True Then                     '*Ao 2014/04/29  PRJ10001TK-PH2-External Spec-Auto Release Paper at Slip Printer-00.02.03 FSBT
                            otmDrawer.Enabled = False
                            bW_Wait = False
                            bW_Printing = False
                            Unload Me
                            Exit Sub
                        End If
'                        If oOposSlip.SlpEmpty = True Then                           '*Ao 2013/09/04
'                            bW_Wait = False
'                            bW_Printing = False
'                            Unload Me
'                            Exit Sub
'                        Else
'                            nLevel = oOposSlip.nCheckHealth(OPOS_CH_INTERNAL)
'                            If nLevel <> OPOS_SUCCESS Then
'                                oOposSlip.ClosePrint
'                                oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
'                                oOposSlip.SlpLineChars = nVB_MaxCharSlip
'                            Else
'                                If nW_Count > (nW_Line + 7) Then
'                                    oOposSlip.ClosePrint
'                                    oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
'                                    oOposSlip.SlpLineChars = nVB_MaxCharSlip
'                                End If
''                                MsgBox "nLevel <> OPOS_E_ILLEGAL =" & nLevel
'                            End If
'                        End If
'                            If bW_ChkEnd Then
'                                .BeginRemoval 5000         '*Ao 2013/07/26
'                                .EndRemoval
'                                bW_ChkEnd = False
'                            End If
'                        End If
                    End If
                Case Else
                    olaForm(1).Caption = POSPrinter1_StatusUpdateEvent(nStatus)
            End Select
            
        Else
            bW_Wait = False
            '--------------------------------------------------------------------------------------------'*Ao 2013/09/13   มี ERROR OPOS_E_ILLEGAL
            nLevel = oOposSlip.nCheckHealth(OPOS_CH_INTERNAL)
            If nLevel = OPOS_E_ILLEGAL Then
                oOposSlip.ClosePrint
                oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
                oOposSlip.SlpLineChars = nVB_MaxCharSlip
                '*Bump 57-09-30 10.15am. CommSheet TK-ISS3000-555 Add log เพื่อ Monitor
                Call SP_CRTxLogPrg("ErrorNumber:" & OPOS_E_ILLEGAL & " ErrorDesc: Slip Printer is ILLEGAL (Program retry open OPOS Slip Printer" & "wCNMsgErrSlip control otmDrawer_Timer" & " Events:otmDrawer_Timer" & ",TerminalNum:" & tVB_CNTerminalNum & " TransNo.:" & tVB_Running & " TransType:" & tVB_Transaction)
            End If
            '--------------------------------------------------------------------------------------------
            Select Case nStatus
                Case PTR_SUE_COVER_OPEN
                    'MsgBox (Status & " : " & "Printer cover is open")
                Case PTR_SUE_COVER_OK
                    'MsgBox (Status & " : " & "Printer cover is closed")
                Case PTR_SUE_JRN_EMPTY
                    'MsgBox (Status & " : " & "No journal paper")
                Case PTR_SUE_JRN_NEAREMPTY
                    'MsgBox (Status & " : " & "Journal paper is low")
                Case PTR_SUE_JRN_PAPEROK
                    'MsgBox (Status & " : " & "Journal paper is ready")
                Case PTR_SUE_REC_EMPTY
                    'MsgBox (Status & " : " & "No receipt paper")
                Case PTR_SUE_REC_NEAREMPTY
                    'MsgBox (Status & " : " & "Receipt paper is low")
                Case PTR_SUE_REC_PAPEROK
                    'MsgBox (Status & " : " & "Receipt paper is ready")
                Case PTR_SUE_SLP_EMPTY
                    'MsgBox (Status & " : " & "No slip form")
'                    If nW_CountEmpty = 20 Then
'                        nVB_MaxCharSlip = 42
'                        oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
'                        oOposSlip.SlpLineChars = nVB_MaxCharSlip
'                    End If
                    'nW_CountEmpty = nW_CountEmpty + 1
                Case PTR_SUE_SLP_NEAREMPTY
                    'MsgBox (Status & " : " & "Almost at the bottom of the slip form")
                Case PTR_SUE_SLP_PAPEROK
                    'MsgBox (Status & " : " & "Slip form is inserted")
                    If bW_Enter Then
                        olaForm(1) = POSPrinter1_StatusUpdateEvent(nStatus)
                        If SlpInsertion() = True Then               '*Ao 2013/09/02     Check Insert Slip
                            Unload Me
                            Exit Sub
                        End If
'                        Unload Me
'                        Exit Sub
                    Else
'                        If nW_Count > nW_Line Then bW_Printing = False
'                        If bW_Printing Then
'                            olaForm(1).Caption = "Slip printing is progress…"
'                        Else
'                            olaForm(0).Caption = "Please press RELEASE at Slip Printer"
'                            olaForm(1).Caption = "when printing is completed" & vbCrLf & "Remove printed slip from Slip Printer"
'                        End If
                    End If
                Case PTR_SUE_IDLE
                    'MsgBox (Status & " : " & "All asynchronous")
                Case Else
                    'MsgBox (nStatus)
                    If bW_Enter Then
                        If SlpInsertion() = True Then               '*Ao 2013/09/02     Check Insert Slip
                            Unload Me
                            Exit Sub
                        End If
'                        Unload Me
'                        Exit Sub
                    End If
            End Select
            olaForm(1) = POSPrinter1_StatusUpdateEvent(nStatus)
        End If
    End With
    bW_Enter = False
    Exit Sub
ErrHandle:
     olaForm(1) = Err.Description & vbCrLf & "Press Enter to print"
     If bW_Enter Then
        Unload Me
        Exit Sub
    End If
    
End Sub

Private Function POSPrinter1_StatusUpdateEvent(ByVal Status As Long) As String
        Select Case Status
            Case PTR_SUE_COVER_OPEN
                'MsgBox (Status & " : " & "Printer cover is open")
                POSPrinter1_StatusUpdateEvent = "Printer cover is open"
            Case PTR_SUE_COVER_OK
                'MsgBox (Status & " : " & "Printer cover is closed")
                POSPrinter1_StatusUpdateEvent = "Printer cover is closed"
            Case PTR_SUE_JRN_EMPTY
                'MsgBox (Status & " : " & "No journal paper")
                POSPrinter1_StatusUpdateEvent = "No journal paper"
            Case PTR_SUE_JRN_NEAREMPTY
                'MsgBox (Status & " : " & "Journal paper is low")
                POSPrinter1_StatusUpdateEvent = "Journal paper is low"
            Case PTR_SUE_JRN_PAPEROK
                'MsgBox (Status & " : " & "Journal paper is ready")
                POSPrinter1_StatusUpdateEvent = "Journal paper is ready"
            Case PTR_SUE_REC_EMPTY
                'MsgBox (Status & " : " & "No receipt paper")
                POSPrinter1_StatusUpdateEvent = "No receipt paper"
            Case PTR_SUE_REC_NEAREMPTY
                'MsgBox (Status & " : " & "Receipt paper is low")
                POSPrinter1_StatusUpdateEvent = "Receipt paper is low"
            Case PTR_SUE_REC_PAPEROK
                'MsgBox (Status & " : " & "Receipt paper is ready")
                POSPrinter1_StatusUpdateEvent = "Receipt paper is ready"
            Case PTR_SUE_SLP_EMPTY
                'MsgBox (Status & " : " & "No slip form")
                Select Case tW_Msg
                    Case "...Continued"
'                        Select Case Val(Format(tVB_TransactionSub))
                        Select Case Val(Format(tW_TransTypeSub))    '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
'                            Case tEN_TrnOpen, tEN_TrnOpenTaxInc, tEN_TrnOpenManualTax
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Open slip form"
'                            Case tEN_TrnStartBank
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Start Bank slip form"
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Return slip form"
                                
'                                Call SP_CRTxLogPrg("SLIP PRINTER-" & tW_Msg & "-Please insert continuation Return slip form")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                                
                            Case tEN_TrnDeposit, tEN_TrnDepositCancel, tEN_TrnAddDeposit, tEN_TrnAddDepositCancel, tEN_TrnRedeem
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Layaway slip form"
                                
'                                Call SP_CRTxLogPrg("SLIP PRINTER-" & tW_Msg & "-Please insert continuation Layaway slip form")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                                
'                            Case tEN_TrnClose
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Close slip form"
'                            Case tEN_TrnPrintClose, tEN_TrnClosingReportTaxInc, tEN_TrnClosingReportManualTax
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Closing Repor slip form"
'                            Case tEN_TrnVIPSale, tEN_TrnCreateVIP
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation VIP Service Sale slip form"
'                            Case tEN_TrnTaxExceptSale
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Tax Except Sale slip form"
                            Case tEN_TrnPickUp
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Cash Transfer slip form"
'                                Call SP_CRTxLogPrg("SLIP PRINTER-" & tW_Msg & "-Please insert continuation Cash Transfer slip form")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
'                            Case tEN_TrnDutyFreeSale
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Duty Free Sale slip form"
'                            Case tEN_TrnVoucher
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Voucher slip form"
'                            Case tEN_TrnNoSale
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation No Sale slip form"
'                            Case tEN_TrnOtherReceive
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Other Receive slip form"
'                            Case tEN_TrnOtherDisbursement
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Other Disbursement slip form"
'                            Case tEN_TrnDepositMoney
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Deposit Money slip form"
'                            Case tEN_TrnTraining
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Training slip form"
'                            Case tEN_TrnTotalReport, tEN_TrnTotalReportTaxInc, tEN_TrnTotalReportManualTax
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Total Report slip form"
'                            Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
'                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Over Short Report slip form"
                            Case Else
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Sale slip form"
'                                Call SP_CRTxLogPrg("SLIP PRINTER-" & tW_Msg & "-Please insert continuation Sale slip form(else)")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                         End Select
                    Case "PRINT SALES ENDORSEMENT", "PRINT CREDIT VOUCHER", "PRINT RE-ISSUE"
                        POSPrinter1_StatusUpdateEvent = "Please insert Credit Voucher slip form"
'                        Call SP_CRTxLogPrg("SLIP PRINTER-" & tW_Msg & "-Please insert Credit Voucher slip form")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                        '*Eaw 56-07-04 CommSheet TK-ISS3000-257.xls
                        Select Case tW_Msg
                            Case "PRINT SALES ENDORSEMENT"
                                'POSPrinter1_StatusUpdateEvent = "CV Amt : " & SP_FMTcCurPoint(cVB_CVAmtEx, False) & "           Please insert Credit Voucher slip form"
                                olaForm(2).Caption = "CV AMT: " & SP_FMTcCurPoint(cVB_CVAmtEx, False)
                            Case Else
                                olaForm(2).Caption = ""
                                'POSPrinter1_StatusUpdateEvent = "Please insert Credit Voucher slip form"
                        End Select
                    Case Else
'                        Select Case Val(Format(tVB_TransactionSub))
                        Select Case Val(Format(tW_TransTypeSub))            '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                POSPrinter1_StatusUpdateEvent = "Please insert Return slip form"
'                                Call SP_CRTxLogPrg("SLIP PRINTER-" & tW_Msg & "-Please insert Return slip form")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                                
                            Case tEN_TrnDeposit, tEN_TrnDepositCancel, tEN_TrnAddDeposit, tEN_TrnAddDepositCancel, tEN_TrnRedeem
                                POSPrinter1_StatusUpdateEvent = "Please insert Layaway slip form"
'                                Call SP_CRTxLogPrg("SLIP PRINTER-" & tW_Msg & "-Please insert Layaway slip form")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                            Case tEN_TrnPickUp
                                POSPrinter1_StatusUpdateEvent = "Please insert Cash Transfer slip form"     '*Ao 2013/09/30  CommSheet TK-ISS3000-365-FAPL-130
                            Case Else
                                POSPrinter1_StatusUpdateEvent = "Please insert Sale slip form"
'                                Call SP_CRTxLogPrg("SLIP PRINTER-" & tW_Msg & "-Please insert Sale slip form")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                         End Select
                End Select

            Case PTR_SUE_SLP_NEAREMPTY
                'MsgBox (Status & " : " & "Almost at the bottom of the slip form")
                POSPrinter1_StatusUpdateEvent = "Almost at the bottom of the slip form"
            Case PTR_SUE_SLP_PAPEROK
                'MsgBox (Status & " : " & "Slip form is inserted")
                If bW_Enter Then
                    If tW_Msg = "...Continued" Then
                        If bW_Wait Then
'                            POSPrinter1_StatusUpdateEvent = "Slip printing is progress…"                           '*Ao 2013/05/14 CommSheet TK-ISS3000-168-ST 00 00 04.xls
'                            POSPrinter1_StatusUpdateEvent = "Now is ""Slip printing is progress"""
                            POSPrinter1_StatusUpdateEvent = "Slip printing in progress"   '*Eaw 56-06-27 CommSheet TK-ISS3000-247-TR.xlsx
                        Else
'                            POSPrinter1_StatusUpdateEvent = "Slip printing is progress…" & vbCrLf & "Press Enter to print"
                            POSPrinter1_StatusUpdateEvent = "Slip printing in progress…" & vbCrLf & "Press Enter to print"  '*Eaw 56-06-27 CommSheet TK-ISS3000-247-TR.xlsx
                        End If
                    Else
'                        POSPrinter1_StatusUpdateEvent = "Slip printing is progress…" & vbCrLf & "Press Enter to print"
                        POSPrinter1_StatusUpdateEvent = "Slip printing in progress…" & vbCrLf & "Press Enter to print" '*Eaw 56-06-27 CommSheet TK-ISS3000-247-TR.xlsx
                    End If
                Else
                    Select Case tW_Msg
                        Case "...Continued"
                            olaForm(1).FontSize = 27
'                            Select Case Val(Format(tVB_TransactionSub))
'                            Select Case Val(Format(tW_TransTypeSub))        '*Ao 2013/08/30  CommSheet TK-ISS3000-336-UAT
'                                Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
''                                    POSPrinter1_StatusUpdateEvent = "Please insert continuation Return slip form" & vbCrLf & "Press Enter to print"    '*Ao 2013/05/14 CommSheet TK-ISS3000-168-ST 00 00 04.xls
'                                    POSPrinter1_StatusUpdateEvent = "Press Enter to print"
'                                Case tEN_TrnDeposit, tEN_TrnDepositCancel, tEN_TrnAddDeposit, tEN_TrnAddDepositCancel, tEN_TrnRedeem
''                                    POSPrinter1_StatusUpdateEvent = "Please insert continuation Layaway slip form" & vbCrLf & "Press Enter to print"
'                                    POSPrinter1_StatusUpdateEvent = "Press Enter to print"
'                                Case Else
''                                    POSPrinter1_StatusUpdateEvent = "Please insert continuation Sale slip form" & vbCrLf & "Press Enter to print"
'                                    POSPrinter1_StatusUpdateEvent = "Press Enter to print"
'                             End Select
                             POSPrinter1_StatusUpdateEvent = "Press Enter to print"
                      Case "PRINT SALES ENDORSEMENT", "PRINT CREDIT VOUCHER", "PRINT RE-ISSUE"
'                            POSPrinter1_StatusUpdateEvent = "Please insert Credit Voucher slip form" & vbCrLf & "Press Enter to print"
                            POSPrinter1_StatusUpdateEvent = "Press Enter to print"
                            '*Eaw 56-07-04 CommSheet TK-ISS3000-257.xls
                            Select Case tW_Msg
                                Case "PRINT SALES ENDORSEMENT"
'                                   POSPrinter1_StatusUpdateEvent = "CV Amt : " & SP_FMTcCurPoint(cVB_CVAmtEx, False) & "                  Press Enter to print"
                                    olaForm(2).Caption = "CV AMT: " & SP_FMTcCurPoint(cVB_CVAmtEx, False)
                                Case Else
                                    olaForm(2).Caption = ""
'                                POSPrinter1_StatusUpdateEvent = "Press Enter to print"
                            End Select
                      Case Else
'                            Select Case Val(Format(tVB_TransactionSub))
'                            Select Case Val(Format(tW_TransTypeSub))
'                                Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
''                                    POSPrinter1_StatusUpdateEvent = "Please insert Return slip form" & vbCrLf & "Press Enter to print"                 '*Ao 2013/05/14 CommSheet TK-ISS3000-168-ST 00 00 04.xls
'                                    POSPrinter1_StatusUpdateEvent = "Press Enter to print"
'                                Case tEN_TrnDeposit, tEN_TrnDepositCancel, tEN_TrnAddDeposit, tEN_TrnAddDepositCancel, tEN_TrnRedeem
''                                    POSPrinter1_StatusUpdateEvent = "Please insert Layaway slip form" & vbCrLf & "Press Enter to print"
'                                    POSPrinter1_StatusUpdateEvent = "Press Enter to print"
'                                Case Else
''                                    POSPrinter1_StatusUpdateEvent = "Please insert Sale slip form" & vbCrLf & "Press Enter to print"
'                                    POSPrinter1_StatusUpdateEvent = "Press Enter to print"
'                             End Select
                             POSPrinter1_StatusUpdateEvent = "Press Enter to print"
                    End Select
                End If
            Case PTR_SUE_IDLE
                'MsgBox (Status & " : " & "All asynchronous")
                 POSPrinter1_StatusUpdateEvent = "All asynchronous"
            Case Else
                'MsgBox (Status)
        End Select
        
End Function

Private Function SlpInsertion() As Boolean
Dim nMSG As Long
Dim nSlpPrintable As Long
Dim tMsg As String
Dim bEndInsert  As Boolean
    
On Error GoTo ErrHandle
    'Insert sheet
    nSlpPrintable = 1
    bEndInsert = False
    
    With oOposSlip
        While nSlpPrintable
            .BeginInsertion 500
            'Check a sheet/Error check
            If (.ResultCode <> OPOS_SUCCESS) Then
                Select Case .ResultCode
                    Case OPOS_E_TIMEOUT
                        tMsg = "Please Insert slip." '*Bump 56-09-06 เพิ่มคำว่า Please
                    
                    Case OPOS_E_ILLEGAL
                        If (.SlpEmpty = True) Then
                            tMsg = "Please Remove slip." '*Bump 56-09-06 เพิ่มคำว่า Please
                            bEndInsert = True
                            .EndInsertion
                        Else
                            tMsg = W_GETtGetErrorMsg()
                            .BeginRemoval -1
                        End If
                    
                    Case Else
                        tMsg = W_GETtGetErrorMsg()
                End Select
                
                nMSG = MsgBox(tMsg, vbRetryCancel, "Please Check slip") '*Bump 56-09-06 เพิ่มคำว่า Please
                
'                Call SP_CRTxLogPrg("SLIP PRINTER-wCNMsgErrSlip.SlpRemoval-TransType:" & tW_TransTypeSub & "SlpInsertion:" & ",tMsg:" & tMsg)  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                
                If bEndInsert = True Then
                    .EndInsertion
                    bEndInsert = False
                End If
                ' Cancel
                If (nMSG = vbCancel) Then
                    .EndInsertion
                    .BeginRemoval (-1)
                    .EndRemoval
                    SlpInsertion = False
                    Exit Function
                End If
            Else
                .EndInsertion
                tMsg = ""
                Select Case .ResultCode
                    Case OPOS_E_EXTENDED
                        If .ResultCodeExtended = OPOS_EPTR_SLP_EMPTY Then
                            tMsg = "Please Insert slip." '*Bump 56-09-06 เพิ่มคำว่า Please
                        End If
                    
                    Case OPOS_SUCCESS
                        'Slip is removed.
                        If .SlpNearEnd = True Then
                            .BeginRemoval -1
                            tMsg = "Please Insert slip." '*Bump 56-09-06 เพิ่มคำว่า Please
                        'Cover is open.
                        ElseIf .CoverOpen = True Then
                            tMsg = "Close the cover"
                        Else
                            nSlpPrintable = 0
                        End If
                        
                    Case OPOS_E_ILLEGAL
                        If .CoverOpen = True And .SlpEmpty = True Then
                            tMsg = "Please Insert slip" '*Bump 56-09-06 เพิ่มคำว่า Please
                            nSlpPrintable = 0
                        Else
                            tMsg = "Please Insert slip" '*Bump 56-09-06 เพิ่มคำว่า Please
    '                        .Close
    '                        .Open ("Unit1")
    '                        .ClaimDevice -1
    '                        .DeviceEnabled = True
    '                        oOposSlip.ClosePrint
    '                        oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
    '                        oOposSlip.SlpLineChars = nVB_MaxCharSlip
                        End If
                End Select
                        
                If tMsg <> "" Then
                    nMSG = MsgBox(tMsg, vbRetryCancel, "Please Check slip") '*Bump 56-09-06 เพิ่มคำว่า Please
                    
'                    Call SP_CRTxLogPrg("SLIP PRINTER-wCNMsgErrSlip.SlpRemoval-TransType:" & tW_TransTypeSub & "Check slip printer:" & ",tMsg:" & tMsg)  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
                    
                    If (nMSG = vbCancel) Then
                        .EndInsertion
                        .BeginRemoval (-1)
                        .EndRemoval
                        SlpInsertion = False
                        Exit Function
                    End If
                End If
            End If
        Wend
    End With

    SlpInsertion = True

ErrHandle:

End Function

Private Function W_GETtGetErrorMsg() As String
Dim tBF As String

'Make messages on each event information
    Select Case oOposSlip.ResultCodeExtended
        Case OPOS_EPTR_COVER_OPEN
            tBF = "Printer cover is open."
        Case OPOS_EPTR_JRN_EMPTY
            tBF = "No jurnal paper."
        Case OPOS_EPTR_REC_EMPTY
            tBF = "No receipt paper."
        Case OPOS_EPTR_SLP_EMPTY
            tBF = "No slip form."
        Case Else
'            tBF = "ResultCode = " + CStr(oOposSlip.ResultCode) + vbCrLf + "ResultCodeExtended = " + CStr(oOposSlip.ResultCodeExtended)
            tBF = "Slip form error"
            oOposSlip.ClosePrint
            oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
            oOposSlip.SlpLineChars = nVB_MaxCharSlip
            
'            Call SP_CRTxLogPrg("SLIP PRINTER-wCNMsgErrSlip.W_GETtGetErrorMsg-TransType:" & tW_TransTypeSub & ",Slip form error(Program retry open OPOS Slip Printer)")  '*Bump 58-01-22 เก็บ Slip Printer -Monitor เครื่องแฮ้ง
            
    End Select

    W_GETtGetErrorMsg = tBF

End Function
