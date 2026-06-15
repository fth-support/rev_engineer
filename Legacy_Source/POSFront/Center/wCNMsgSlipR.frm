VERSION 5.00
Begin VB.Form wCNMsgSlipR 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "wCNMsgSlipR"
   ClientHeight    =   4695
   ClientLeft      =   1815
   ClientTop       =   1740
   ClientWidth     =   7215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.Timer otmDrawer 
      Interval        =   500
      Left            =   120
      Top             =   80
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
      TabIndex        =   3
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Shape oshForm 
      BorderWidth     =   8
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   7215
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
      TabIndex        =   2
      Top             =   120
      Width           =   6975
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
      TabIndex        =   0
      Top             =   1560
      Width           =   6975
   End
End
Attribute VB_Name = "wCNMsgSlipR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tW_Msg As String
Public tW_TransTypeSub As String
Public bW_Printing As Boolean
Public bW_Wait  As Boolean
Public nW_Line As Long

Private bW_Enter As Boolean
'Private nW_CountEmpty As Long
'Private nW_Count As Long
'Private nW_MsgCount As Integer '*Eaw 56-07-23 CommSheet TK-ISS3000-274 TR FAPL-053 00.00.02.xls
'Private bW_Chkprn As Boolean
'Private bW_ChkEnd  As Boolean

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
            .BeginRemoval 3000
            If (.ResultCode <> OPOS_E_TIMEOUT) Then
                If (.ResultCode = OPOS_E_BUSY) Then
                    SleepEx 5000, 0
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
        Wend
        If (.ResultCode <> OPOS_E_ILLEGAL) Then
            RecPrintable = 1
            While RecPrintable
                .EndRemoval
                If (.ResultCode = OPOS_SUCCESS) Then
                    oOposSlip.ClearOutput            '*Ao 2013/07/29
                    RecPrintable = 0
                Else
'                    pMSG = W_GETtGetErrorMsg()
'                    MSG = MsgBox(pMSG, vbRetryCancel + vbQuestion, "Print credit sales slip")
'                    MSG = MsgBox("Please Remove Slip form", vbOKOnly + vbExclamation, "Slip form")
                    DoEvents
                    olaForm(1).Caption = "When printing is completed" & vbCrLf & "Remove printed slip form Slip Printer"
                    'Canceled
                    If (MSG = vbCancel) Then
                        SlpRemoval = False
                        Exit Function
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
                        tMsg = "Please Insert slip."
                    
                    Case OPOS_E_ILLEGAL
                        If (.SlpEmpty = True) Then
                            tMsg = "Please Remove slip."
                            bEndInsert = True
                            .EndInsertion
                        Else
                            tMsg = W_GETtGetErrorMsg()
                            .BeginRemoval -1
                        End If
                    
                    Case Else
                        tMsg = W_GETtGetErrorMsg()
                End Select
                
                nMSG = MsgBox(tMsg, vbRetryCancel, "Please Check slip")
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
                            tMsg = "Please Insert slip."
                        End If
                    
                    Case OPOS_SUCCESS
                        'Slip is removed.
                        If .SlpNearEnd = True Then
                            .BeginRemoval -1
                            tMsg = "Please Insert slip."
                        'Cover is open.
                        ElseIf .CoverOpen = True Then
                            tMsg = "Close the cover"
                        Else
                            nSlpPrintable = 0
                        End If
                        
                    Case OPOS_E_ILLEGAL
                        If .CoverOpen = True And .SlpEmpty = True Then
                            tMsg = "Please Insert slip"
                            nSlpPrintable = 0
                        Else
                            tMsg = "Please Insert slip"
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
                    nMSG = MsgBox(tMsg, vbRetryCancel, "Please Check slip")
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
            tBF = "Slip form error"
            oOposSlip.ClosePrint
            oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
            oOposSlip.SlpLineChars = nVB_MaxCharSlip
    End Select

    W_GETtGetErrorMsg = tBF

End Function

Private Function POSPrinter1_StatusUpdateEvent(ByVal Status As Long) As String
        Select Case Status
            Case PTR_SUE_COVER_OPEN
                POSPrinter1_StatusUpdateEvent = "Printer cover is open"
            Case PTR_SUE_COVER_OK
                POSPrinter1_StatusUpdateEvent = "Printer cover is closed"
            Case PTR_SUE_JRN_EMPTY
                POSPrinter1_StatusUpdateEvent = "No journal paper"
            Case PTR_SUE_JRN_NEAREMPTY
                POSPrinter1_StatusUpdateEvent = "Journal paper is low"
            Case PTR_SUE_JRN_PAPEROK
                POSPrinter1_StatusUpdateEvent = "Journal paper is ready"
            Case PTR_SUE_REC_EMPTY
                POSPrinter1_StatusUpdateEvent = "No receipt paper"
            Case PTR_SUE_REC_NEAREMPTY
                POSPrinter1_StatusUpdateEvent = "Receipt paper is low"
            Case PTR_SUE_REC_PAPEROK
                POSPrinter1_StatusUpdateEvent = "Receipt paper is ready"
            Case PTR_SUE_SLP_EMPTY
                Select Case tW_Msg
                    Case "...Continued"
                        Select Case Val(Format(tW_TransTypeSub))
                            Case tEN_TrnOpen, tEN_TrnOpenTaxInc, tEN_TrnOpenManualTax
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Open slip form"
                            Case tEN_TrnStartBank
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Start Bank slip form"
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Return slip form"
                            Case tEN_TrnDeposit, tEN_TrnDepositCancel, tEN_TrnAddDeposit, tEN_TrnAddDepositCancel, tEN_TrnRedeem
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Layaway slip form"
                            Case tEN_TrnClose
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Close slip form"
                            Case tEN_TrnPrintClose, tEN_TrnClosingReportTaxInc, tEN_TrnClosingReportManualTax
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Closing Repor slip form"
                            Case tEN_TrnVIPSale, tEN_TrnCreateVIP
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation VIP Service Sale slip form"
                            Case tEN_TrnTaxExceptSale
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Tax Except Sale slip form"
                            Case tEN_TrnPickUp
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Cash Transfer slip form"
                            Case tEN_TrnDutyFreeSale
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Duty Free Sale slip form"
                            Case tEN_TrnVoucher
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Voucher slip form"
                            Case tEN_TrnNoSale
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation No Sale slip form"
                            Case tEN_TrnOtherReceive
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Other Receive slip form"
                            Case tEN_TrnOtherDisbursement
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Other Disbursement slip form"
                            Case tEN_TrnDepositMoney
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Deposit Money slip form"
                            Case tEN_TrnTraining
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Training slip form"
                            Case tEN_TrnTotalReport, tEN_TrnTotalReportTaxInc, tEN_TrnTotalReportManualTax
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Total Report slip form"
                            Case tEN_TrnOverShortReport, tEN_TrnOverShortReportTaxInc, tEN_TrnOverShortReportManualTax
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Over Short Report slip form"
                            Case Else
                                POSPrinter1_StatusUpdateEvent = "Please insert continuation Sale slip form"
                         End Select
                    Case "PRINT SALES ENDORSEMENT", "PRINT CREDIT VOUCHER", "PRINT RE-ISSUE"
                        POSPrinter1_StatusUpdateEvent = "Please insert Credit Voucher slip form"
                        Select Case tW_Msg
                            Case "PRINT SALES ENDORSEMENT"
                                olaForm(2).Caption = "CV AMT: " & SP_FMTcCurPoint(cVB_CVAmtEx, False)
                            Case Else
                                olaForm(2).Caption = ""
                        End Select
                    Case Else
                        Select Case Val(Format(tW_TransTypeSub))
                            Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                                POSPrinter1_StatusUpdateEvent = "Please insert Return slip form"
                            Case tEN_TrnDeposit, tEN_TrnDepositCancel, tEN_TrnAddDeposit, tEN_TrnAddDepositCancel, tEN_TrnRedeem
                                POSPrinter1_StatusUpdateEvent = "Please insert Layaway slip form"
                            Case tEN_TrnPickUp
                                POSPrinter1_StatusUpdateEvent = "Please insert Cash Transfer slip form"
                            Case Else
                                POSPrinter1_StatusUpdateEvent = "Please insert Sale slip form"
                         End Select
                End Select
            Case PTR_SUE_SLP_NEAREMPTY
                POSPrinter1_StatusUpdateEvent = "Almost at the bottom of the slip form"
            Case PTR_SUE_SLP_PAPEROK
                If bW_Enter Then
                    If tW_Msg = "...Continued" Then
                        If bW_Wait Then
                            POSPrinter1_StatusUpdateEvent = "Slip printing in progress"
                        Else
                            POSPrinter1_StatusUpdateEvent = "Slip printing in progress…" & vbCrLf & "Press Enter to print"
                        End If
                    Else
                        POSPrinter1_StatusUpdateEvent = "Slip printing in progress…" & vbCrLf & "Press Enter to print"
                    End If
                Else
                    Select Case tW_Msg
                        Case "...Continued"
                             POSPrinter1_StatusUpdateEvent = "Press Enter to print"
                      Case "PRINT SALES ENDORSEMENT", "PRINT CREDIT VOUCHER", "PRINT RE-ISSUE"
                            POSPrinter1_StatusUpdateEvent = "Press Enter to print"
                            Select Case tW_Msg
                                Case "PRINT SALES ENDORSEMENT"
                                    olaForm(2).Caption = "CV AMT: " & SP_FMTcCurPoint(cVB_CVAmtEx, False)
                                Case Else
                                    olaForm(2).Caption = ""
                            End Select
                      Case Else
                             POSPrinter1_StatusUpdateEvent = "Press Enter to print"
                    End Select
                End If
            Case PTR_SUE_IDLE
                 POSPrinter1_StatusUpdateEvent = "All asynchronous"
            Case Else
                POSPrinter1_StatusUpdateEvent = ""
        End Select
        
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF12
            bVB_CNCheckCV = True
            otmDrawer.Enabled = False
            Unload Me
        Case 13
            bW_Enter = True
    End Select
End Sub

Private Sub Form_Load()
Dim nStatus As Long

On Error GoTo ErrHandle
    
    Call SP_FRMxChildCenter(Me)
    
    olaForm(1).FontSize = 27
    bVB_CNCheckCV = False
    bW_Enter = False
    If tW_Msg <> "" Then
        olaForm(0).Caption = tW_Msg
    Else
        olaForm(0).Caption = "Slip Printer"
    End If
    
    If bW_Wait Then
        olaForm(1).Caption = "Slip printing in progress"
    Else
        If oOposSlip.SlpEmpty = True Then
            nStatus = PTR_SUE_SLP_EMPTY
        Else
            nStatus = PTR_SUE_SLP_PAPEROK
        End If
        olaForm(1).Caption = POSPrinter1_StatusUpdateEvent(nStatus)
    End If
    otmDrawer.Enabled = True
Exit Sub
ErrHandle:
    olaForm(1).Caption = Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    otmDrawer.Enabled = False
    Call SP_OBJxSetNothing(Me)
Exit Sub
ErrHandle:
End Sub

Private Sub otmDrawer_Timer()
Dim nStatus As Long
Dim nLevel As Long

On Error GoTo ErrHandle
    
    If oOposSlip.SlpEmpty = True Then
        nStatus = PTR_SUE_SLP_EMPTY
    Else
        nStatus = PTR_SUE_SLP_PAPEROK
    End If
    
    If bW_Wait Then
        olaForm(1).Caption = "Slip printing in progress"
        If oOposSlip.ResultCode <> OPOS_SUCCESS Then
            olaForm(1).Caption = "Cannot use a Slip Printer."
            oOposSlip.ClosePrint
            oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
            oOposSlip.SlpLineChars = nVB_MaxCharSlip
        End If
        SleepEx 7000, 0
        If SlpRemoval() = True Then
            otmDrawer.Enabled = False
            Unload Me
            Exit Sub
        End If
    Else
        nLevel = oOposSlip.nCheckHealth(OPOS_CH_INTERNAL)
        If nLevel = OPOS_E_ILLEGAL Then
            oOposSlip.ClosePrint
            oOposSlip.OpenPrint tVB_CashSlipPrinter, nVB_MaxCharSlip, nVB_CashCodePage 'Open Printer OPOS
            oOposSlip.SlpLineChars = nVB_MaxCharSlip
        End If
        olaForm(1).Caption = POSPrinter1_StatusUpdateEvent(nStatus)
        If bW_Enter Then
            If SlpInsertion() = True Then               '*Ao 2013/09/02     Check Insert Slip
                otmDrawer.Enabled = False
                Unload Me
                Exit Sub
            End If
        End If
    End If
Exit Sub
ErrHandle:
    otmDrawer.Enabled = False
    olaForm(1) = Err.Description
     If bW_Enter Then
        Unload Me
        Exit Sub
    End If
End Sub
