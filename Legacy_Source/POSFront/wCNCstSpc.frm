VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNCstSpc 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMERS VIP;klg_Title"
   ClientHeight    =   2475
   ClientLeft      =   7920
   ClientTop       =   2250
   ClientWidth     =   4785
   ControlBox      =   0   'False
   Icon            =   "wCNCstSpc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4785
   Tag             =   "2;"
   Begin SCANNERLib.Scanner oScanner 
      Left            =   120
      Top             =   1680
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1815
      Index           =   2
      Left            =   -15
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   4785
      _Version        =   720897
      _ExtentX        =   8440
      _ExtentY        =   3201
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   360
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "01234567890123456789"
         Top             =   720
         Width           =   4185
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMERS VIP;klg_Cust"
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
         Left            =   360
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   360
         Width           =   3270
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   -15
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1260
      Width           =   4785
      _Version        =   720897
      _ExtentX        =   8440
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Left            =   3120
         TabIndex        =   2
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
         Left            =   1560
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNCstSpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Form_Activate()
'    Call SP_OBJxSetFocus(otbID)
'End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    bVB_FormPayAct = True 'PH 2.0 ML-SF-019 ไม่ต้องการให้ทำที่ Form_Activate ของ wMain
    otbID.Text = ""
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

Private Sub ocmCancel_Click()
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOpen:
        Case tEN_TrnStartBank:
        Case tEN_TrnSale:
        Case tEN_TrnReturn:
        Case tEN_TrnIPV:
        Case tEN_TrnDeposit:
        Case tEN_TrnRedeem:
        Case tEN_TrnClose:
        Case tEN_TrnPrintClose:
        Case tEN_TrnVIPSale: bVB_TrnCreateVIPSale = False
        Case tEN_TrnTaxExceptSale:
        Case tEN_TrnPickUp:
        Case tEN_TrnDutyFreeSale:
        Case tEN_TrnIPVReturn:
        Case tEN_TrnVoucher:
        Case tEN_TrnDepositCancel:
        Case tEN_TrnIPVRedeem:
        Case tEN_TrnNoSale:
        Case tEN_TrnAddDeposit:
        Case tEN_TrnAddDepositCancel:
        Case tEN_TrnCancelVoucher:
        Case tEN_TrnOtherReceive:
        Case tEN_TrnOtherDisbursement:
        Case tEN_TrnDepositMoney:
        Case tEN_TrnTraining:
        Case tEN_TrnIPVVIPService:
        Case tEN_TrnIPVTaxExceptSale:
        Case tEN_TrnIPVDutyFreeSale:
        Case tEN_TrnCreateVIP: bVB_TrnCreateVIPSale = False
    End Select
    Unload Me
End Sub

Private Sub ocmOK_Click()
    If otbID.Text = "" Then
        MsgBox "Not Allowed", vbCritical
        Call SP_OBJxSetFocus(otbID)
        Exit Sub
    End If
    Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnOpen:
        Case tEN_TrnStartBank:
        Case tEN_TrnSale: Call W_PRCxSale
        Case tEN_TrnReturn:
        Case tEN_TrnIPV:
        Case tEN_TrnDeposit:
        Case tEN_TrnRedeem:
        Case tEN_TrnClose:
        Case tEN_TrnPrintClose:
        Case tEN_TrnVIPSale: bVB_TrnCreateVIPSale = True: Call W_PRCxVIPSale
        Case tEN_TrnTaxExceptSale:
        Case tEN_TrnPickUp:
        Case tEN_TrnDutyFreeSale:
        Case tEN_TrnIPVReturn:
        Case tEN_TrnVoucher:
        Case tEN_TrnDepositCancel:
        Case tEN_TrnIPVRedeem:
        Case tEN_TrnNoSale:
        Case tEN_TrnAddDeposit:
        Case tEN_TrnAddDepositCancel:
        Case tEN_TrnCancelVoucher:
        Case tEN_TrnOtherReceive:
        Case tEN_TrnOtherDisbursement:
        Case tEN_TrnDepositMoney:
        Case tEN_TrnTraining:
        Case tEN_TrnIPVVIPService:
        Case tEN_TrnIPVTaxExceptSale:
        Case tEN_TrnIPVDutyFreeSale:
        Case tEN_TrnCreateVIP: bVB_TrnCreateVIPSale = True: Call W_PRCxCreateVIPSale
    End Select
   Unload Me
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
    '*KT 52-06-05 PH 2.0 ML-SF-020 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
    With oScanner
        otbID.Text = .ScanData
'        DoEvents        '*KT 53-01-14
        If otbID.Text <> "" Then
            If Asc(Right(otbID.Text, 1)) = 13 Then
                otbID.Text = Mid(otbID.Text, 1, Len(otbID.Text) - 1)
            End If
        End If
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
    If otbID.Text <> "" Then
        Call SP_OBJxSetFocus(ocmOK)
    Else
        Call SP_OBJxSetFocus(otbID)
    End If
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub otbID_GotFocus()
    otbID.SelStart = 0
    otbID.SelLength = Len(otbID.Text)
End Sub

Private Sub otbID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbID.Text = ""
    End Select
End Sub

Private Sub otbID_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 13
            Call ocmOK_Click
    End Select
End Sub

Private Sub W_PRCxSale()
'----------------------------------------------------------
'   Call:
'   Cmt : *KT 50-12-26  Update HD ลูกค้า VIP ในการขายแบบ Sale
'----------------------------------------------------------
    Dim tSql$
    tSql = "UPDATE " & tVB_CNTblHD
    tSql = tSql & " SET FTCstVIPCode = '" & Trim(otbID.Text) & "'"
    tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & "  AND FTShdStaDoc = '2'"
    Call SP_SQLvExecute(tSql)
    tVB_CstID = Trim(otbID.Text)
End Sub

Private Sub W_PRCxVIPSale()
'----------------------------------------------------------
'   Call:
'   Cmt : *KT 50-12-26  Update HD ลูกค้า VIP ในการขายแบบ VIP Service Sale
'----------------------------------------------------------
    Dim tSql$
    If Trim(otbID.Text) = "" Then
        Exit Sub
    End If
    tSql = "UPDATE " & tVB_CNTblHD
    tSql = tSql & " SET FTCstVIPCode = '" & Trim(otbID.Text) & "'"
    tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & "  AND FTShdStaDoc = '2'"
    Call SP_SQLvExecute(tSql)
    tVB_CstID = Trim(otbID.Text)
End Sub

Private Sub W_PRCxCreateVIPSale()
'----------------------------------------------------------
'   Call:
'   Cmt : *KT 50-12-26  Update HD ลูกค้า VIP ในการขายแบบ VIP Service Sale
'----------------------------------------------------------
    Dim tSql$
    If Trim(otbID.Text) = "" Then
        Exit Sub
    End If
    tSql = "UPDATE " & tVB_CNTblHD
    tSql = tSql & " SET FTCstVIPCode = '" & Trim(otbID.Text) & "'"
    tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & "  AND FTShdStaDoc = '2'"
    Call SP_SQLvExecute(tSql)
    tVB_CstID = Trim(otbID.Text)
End Sub

Private Sub otbID_LostFocus()
    otbID.Text = SP_TXTtTrimRightLeft(otbID.Text)
End Sub
