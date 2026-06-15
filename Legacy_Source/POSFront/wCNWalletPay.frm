VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNWalletPay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wallet Payment;klg_Title"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "wCNWalletPay.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6075
   Tag             =   "2;"
   Begin VB.PictureBox opbWalletPay 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4485
      Index           =   0
      Left            =   0
      Picture         =   "wCNWalletPay.frx":E8CE
      ScaleHeight     =   299
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   411
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6165
      Begin VB.Timer otmForm 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5430
         Top             =   3240
      End
      Begin VB.Label olaShwEsc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Please press ESC for Manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   840
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   1950
         Width           =   4425
      End
      Begin VB.Label olaCntEsc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Tag             =   "2;"
         Top             =   2340
         Width           =   4425
      End
   End
   Begin VB.PictureBox opbWalletPay 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4485
      Index           =   2
      Left            =   12330
      Picture         =   "wCNWalletPay.frx":118A3
      ScaleHeight     =   299
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   411
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   6165
      Begin VB.Frame ofaManual 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3000
         Left            =   0
         TabIndex        =   12
         Top             =   1500
         Width           =   6120
         Begin XtremeSuiteControls.TabControl otcForm1 
            Height          =   2115
            Index           =   2
            Left            =   30
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   -180
            Width           =   5995
            _Version        =   720897
            _ExtentX        =   10575
            _ExtentY        =   3731
            _StockProps     =   64
            Appearance      =   10
            Color           =   32
            Begin VB.TextBox otbPartnerID 
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
               Left            =   2730
               MaxLength       =   30
               TabIndex        =   16
               Top             =   1320
               Width           =   2655
            End
            Begin XtremeSuiteControls.ComboBox ocbWalletType 
               Height          =   495
               Left            =   2745
               TabIndex        =   14
               Top             =   570
               Width           =   2655
               _Version        =   720897
               _ExtentX        =   4683
               _ExtentY        =   873
               _StockProps     =   77
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   15.75
                  Charset         =   222
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
            End
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "WPS Trans ID :"
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
               Left            =   630
               TabIndex        =   19
               Top             =   1350
               Width           =   1950
            End
            Begin VB.Label olaForm 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Wallet Type :"
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
               Left            =   630
               TabIndex        =   15
               Top             =   570
               Width           =   1665
            End
         End
         Begin XtremeSuiteControls.TabControl otcForm1 
            Height          =   1260
            Index           =   1
            Left            =   30
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1680
            Width           =   5995
            _Version        =   720897
            _ExtentX        =   10575
            _ExtentY        =   2222
            _StockProps     =   64
            Appearance      =   10
            Color           =   -1484272203
            Begin VB.CommandButton ocmCancel2 
               Caption         =   "Cancel;klg_Cancel"
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
               Left            =   4140
               TabIndex        =   18
               TabStop         =   0   'False
               Tag             =   "2;"
               Top             =   530
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
               Left            =   2520
               TabIndex        =   17
               TabStop         =   0   'False
               Tag             =   "2;"
               Top             =   530
               Width           =   1455
            End
         End
      End
   End
   Begin VB.PictureBox opbWalletPay 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4485
      Index           =   1
      Left            =   6165
      Picture         =   "wCNWalletPay.frx":14878
      ScaleHeight     =   299
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   411
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   6165
      Begin SCANNERLib.Scanner oScanner 
         Left            =   0
         Top             =   0
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.TextBox otbCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   840
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1800
         Width           =   4430
      End
      Begin VB.CommandButton ocmCancel 
         Caption         =   "Cancel;klg_Cancel"
         Height          =   465
         Left            =   3975
         TabIndex        =   3
         Tag             =   "2;"
         Top             =   3540
         Width           =   1275
      End
      Begin VB.CommandButton ocmCheck 
         Caption         =   "Check;klg_Check"
         Height          =   465
         Left            =   825
         TabIndex        =   2
         Tag             =   "2;"
         Top             =   3540
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.PictureBox opbResult 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   1
         Left            =   600
         Picture         =   "wCNWalletPay.frx":1784D
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2325
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox opbResult 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   0
         Left            =   600
         Picture         =   "wCNWalletPay.frx":18186
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2325
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Timer otmCallFunc 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   5460
         Top             =   1710
      End
      Begin VB.Timer otmWallet 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5460
         Top             =   2220
      End
      Begin VB.Label olaInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "*Please note this transaction is not approved yet."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   60
         TabIndex        =   22
         Tag             =   "2;"
         Top             =   4080
         Visible         =   0   'False
         Width           =   6045
      End
      Begin VB.Label olaCntDown 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   840
         TabIndex        =   9
         Top             =   3075
         Visible         =   0   'False
         Width           =   4425
      End
      Begin VB.Label olaResult 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Progressing Payment ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   840
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   2550
         Visible         =   0   'False
         Width           =   4425
      End
      Begin VB.Label olaWalletPay 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Scan Barcode/QR Code from Customer's Mobile;klg_TokenCode"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   840
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   1500
         Width           =   5865
      End
   End
End
Attribute VB_Name = "wCNWalletPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum nW_CallFunc
    Payment = 0
    Reverse = 1
    Query = 2
    Void = 3
    Refund = 4
    Finalize = 5
    AutoReverse = 9
End Enum

Private Enum nW_SecurityProtocolType
    Default = 0
    Ssl3 = 48
    Tls = 192
    Tls11 = 768
    Tls12 = 3072
End Enum

Public tW_BizDate As String
Public tW_InVoiceID As String
Public tW_InVoiceDT As String
Public tW_Amount As String
Public tW_WalletName As String
Public tW_WalletTransID As String
Public tW_WPSTransID As String

Public tW_RefID As String
Public tW_RefDT As String
Public tW_RefAmount As String
Public tW_RefBrand As String
Public tW_RefStoreID As String
Public tW_RefDeviceID As String
Public tW_RefWalletName As String
Public tW_RefWalletTransID As String
Public tW_RefPartnerTransID As String
Public tW_RefWPSTransID As String
Public tW_RefDocDT As String
Public bW_NoWaitReverse As Boolean

Public nW_WalletMode As Integer '0=Payment,1=Reverse,3=Void,4=Refund,5=Finalize,9=AutoReverse

Dim dW_TimeOut As Date
Dim nW_CntTime As Integer
Dim bW_FlagApvSucess As Boolean
Dim bW_FlagSendAmt As Boolean
Dim bW_CHKTimeOut As Boolean
Dim oFSO As New FileSystemObject
Dim tResult As String
Dim tW_ErrDesc As String
Dim nW_Retry As Integer
Dim bW_Cancel As Boolean
Dim nW_Timer As Integer
Dim nW_TimeESC As Integer
Dim bW_HandleESC As Boolean
Dim bW_OffLineSale As Boolean
Const nW_RevRetry As Integer = 1
Const tW_Starting As String = "Payment In Progress ...;Payment In Progress ..."
Const tW_Fail As String = "Payment Declined !;Payment Declined !"
Const tW_Success As String = "Payment Successful;Payment Successful"
Const tW_WaitConfirm As String = "The Transaction result is still unknown.;The Transaction result is still unknown."
Const tW_WaitConfirm2 As String = "Please choose an action.;Please choose an action."
Const tW_Canceling As String = "Canceling Payment ...;Canceling Payment ..."
Const tW_Canceled As String = "Payment Canceled;Payment Canceled"
Const tW_VoidStarting As String = "Progressing Void ...;Progressing Void ..."
Const tW_Voided As String = "Wallet was Voided;Wallet was Voided"
Const tW_VoidFail As String = "Void Not Success !;Wallet Void Not Success !"
Const tW_RefundStarting As String = "Progressing Refund ...;Progressing Refund ..."
Const tW_Refunded As String = "Wallet was Refunded;Wallet was Refunded"
Const tW_RefundFail As String = "Refund Not Success !;Refund Not Success !"
Const tW_FinalizeFail As String = "Finalize Declined !;Finalize Declined !"
Const tW_ReverseFail As String = "Reversal Declined !;Reversal Declined !"
Const tW_PendingInfo As String = "*Please note this transaction is not approved yet.;*Please note this transaction is not approved yet." '*Tao 61-11-20 TK-ISS3000-711 TK ISSUE - TK Request - Alert message on AliPay pop-up
Const tW_RegExpAlipay As String = "(2[56789]|3[0123])[0-9]{14,22}"
Const tW_RegExpWeChatPay As String = "" 'ยังไม่รู้
Const tW_RegExpBluePay As String = "^250[0-9]{15}$"
Const tW_RegExpMPay As String = "^[0-9]{12}$"

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If bW_HandleESC Then
                Call W_PRCxManualMode '*Tao 60-09-08 เข้า Manual Mode
            End If
    End Select
End Sub

Private Sub W_PRCxManualMode()
'*Tao 60-09-08 Manual Mode (Pay OffLine)
Dim nIndex As Integer
On Error Resume Next
    bW_HandleESC = False
    otmForm.Enabled = False
    bW_OffLineSale = True
    opbWalletPay(2).Visible = True
    opbWalletPay(2).Left = 0
    opbWalletPay(2).ZOrder 0
    opbWalletPay(0).Visible = False
    ocmCancel2.Cancel = True
    For nIndex = 0 To UBound(aVB_WalletType) - 1
        ocbWalletType.AddItem (aVB_WalletType(nIndex).TenderType)
    Next nIndex
    ocbWalletType.ListIndex = 0
    Call SP_OBJxSetFocus(ocbWalletType)
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SaveSetting("AdaPosMall", "POSFront", "SentAutoReverse", "0") 'สั่งให้ Service Auto Reverse หยุดทำงาน
    'Finalize ไม่ต้องแสดงหน้าจอ
    If (nW_WalletMode = nW_CallFunc.Finalize) Or (nW_WalletMode = nW_CallFunc.Reverse) Then
        bW_NoWaitReverse = True
        Me.Top = Screen.Height + Me.Height
        bW_HandleESC = False
    Else
        bW_NoWaitReverse = False
        Call SP_FRMxChildCenter(Me)
        nW_TimeESC = Val(olaCntEsc.Caption) + 1
        bW_HandleESC = True
    End If
    '=====================================
        
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    
    olaShwEsc.Caption = "Please press ESC for Manual"
    olaWalletPay.Caption = cUT.UT_STRtToken(olaWalletPay.Tag, ";", nVB_CNCutLang + 1)
    ocmCancel.Caption = cUT.UT_STRtToken(ocmCancel.Tag, ";", nVB_CNCutLang + 1)

    Call W_ClearLogWallet
    wFunctionKB.tW_No = ""
    wFunctionKB.tW_BnkCode = ""
    wFunctionKB.tW_BnkName = ""
    wFunctionKB.tW_AppCode = ""
    wFunctionKB.tW_Rmk = ""
    wFunctionKB.tW_TransID = ""
    wFunctionKB.tW_WPSTransID = ""
    wFunctionKB.tW_WalletName = ""
    wFunctionKB.tW_WalletTransID = ""
    bW_FlagApvSucess = False
    bW_FlagSendAmt = False
    bW_CHKTimeOut = False
    bW_Cancel = False
    bW_OffLineSale = False
    nW_CntTime = 0
    nW_Retry = 0
    nW_Timer = 0
    otbCode.Locked = False
    olaCntDown.Caption = "0"
    otmCallFunc.Interval = 500
    If (nW_WalletMode = nW_CallFunc.Finalize) Or (nW_WalletMode = nW_CallFunc.Reverse) Then  'Finalize, Reverse
        olaWalletPay.Visible = False
        otbCode.Visible = False
        ocmCancel.Visible = False
        bW_FlagSendAmt = True
        otmCallFunc.Enabled = True
    Else
        Select Case SP_VALnText2Double(tVB_TransactionSub)
        Case tEN_TrnReturn, tEN_TrnIPV, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnIPVTaxExceptSale, tEN_TrnCancelVoucher, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVVIPService, tEN_TrnIPVReturn, tEN_TrnIPVRedeem, tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc
            olaWalletPay.Visible = False 'ทำรายการคืน ไม่ต้องแสดงข้อความ
            otbCode.Visible = False
            ocmCancel.Visible = False
            bW_FlagSendAmt = True
            otmCallFunc.Enabled = True
        Case Else
            olaWalletPay.Visible = True
            otbCode.Visible = True
            ocmCancel.Visible = True
            otmCallFunc.Enabled = False
            '*Tao 60-09-05 แสดงหน้าจอ ให้รอกด ESC
            If bW_HandleESC Then
                Select Case SP_VALnText2Double(tVB_TransactionSub)
                Case tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc '*Tao 60-09-08 ให้ไปที่หน้า Manual
                    Call W_PRCxManualMode '*Tao 60-09-08 เข้า Manual Mode
                Case Else
                    opbWalletPay(0).Visible = True
                    opbWalletPay(0).Left = 0
                    opbWalletPay(0).ZOrder 0
                    otmForm.Enabled = True
                End Select
            Else
                opbWalletPay(1).Left = 0
                opbWalletPay(1).ZOrder 0
            End If
        End Select
    End If
Exit Sub
ErrHandle:
    Call SP_TXTxWriteLogWalletPay("Load Error  = " & Err.Description)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    If bW_FlagSendAmt Then
        Cancel = True
    Else
        If (Not bW_Cancel) And (aVB_WalletLog.Token <> "") Then 'ไม่มี Token แสดงว่าไม่ใช่การ Sale
            If nW_WalletMode <> nW_CallFunc.Reverse Then 'ถ้า Reverse แล้ว ไม่ต้องทำอีก
                bW_NoWaitReverse = True
                ocmCheck.Enabled = False
                ocmCancel.Enabled = False
                olaResult.Caption = ""
                olaCntDown.Caption = ""
                nW_Retry = nW_RevRetry
                bW_FlagSendAmt = True
                nW_WalletMode = nW_CallFunc.Reverse
                otmCallFunc.Interval = 1000
                otmCallFunc.Enabled = True
                Cancel = True
            End If
        Else
            otmWallet.Enabled = False
            otmCallFunc.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    '*Tao 60-09-15 Disable Scanner
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
End Sub

Private Sub ocbWalletType_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13
    Call SP_OBJxSetFocus(otbPartnerID)
Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown 'ให้กด ลูกศร ได้
    '
Case vbKeyEscape
    Call ocmCancel2_Click
Case Else
    KeyCode = 0
End Select
End Sub

Private Sub ocmCancel_Click()
On Error Resume Next
    If bW_OffLineSale Then Exit Sub
    If bW_FlagSendAmt Then
        Exit Sub
    Else
        ocmCheck.Enabled = False
        ocmCancel.Enabled = False
        If (aVB_WalletLog.Token <> "") Then
            bW_NoWaitReverse = True 'กดปุ่ม Cancel เมื่อ Reverse แล้วให้ปิดไปเลย
            nW_Retry = 1
            olaResult.Caption = ""
            olaCntDown.Caption = ""
            bW_FlagSendAmt = True
            nW_WalletMode = nW_CallFunc.Reverse
            otmCallFunc.Interval = 1000
            otmCallFunc.Enabled = True
        Else
            bW_Cancel = True
            Unload Me
        End If
    End If
End Sub

Private Sub ocmCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call ocmCancel_Click
    End Select
End Sub

Private Sub ocmCancel2_Click()
    If Not bW_OffLineSale Then Exit Sub
    bW_FlagSendAmt = False
    bW_Cancel = True
    Unload Me
End Sub

Private Sub ocmCancel2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call ocmCancel2_Click
    End If
End Sub

Private Sub ocmCheck_Click()
    bW_FlagSendAmt = True
    ocmCancel.Enabled = False
    ocmCheck.Enabled = False
    olaResult.Caption = ""
    olaCntDown.Caption = ""
    nW_WalletMode = nW_CallFunc.Query
    otmCallFunc.Interval = 1000
    otmCallFunc.Enabled = True
End Sub

Private Sub ocmCheck_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call ocmCancel_Click
    End Select
End Sub

Private Sub ocmOk_Click()
    If Not bW_OffLineSale Then Exit Sub
    '*Tao 61-08-02 ไม่ใส่ก็ได้ (ตาม Requirement)
'    If Trim(otbPartnerID.Text) = "" Then 'ไม่ใส่
'        Call SP_MSGnShowing("tms_CN999042", nCS_Warning)
'        Call SP_OBJxSetFocus(otbPartnerID)
'        Exit Sub
'    End If
    Select Case ocbWalletType.ListIndex
    Case 0
        If Not aVB_WalletType(0).StaSale Then
            Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(0).TenderType) 'ไม่อนุญาตชำระเงินด้วย Alipay
            Call SP_OBJxSetFocus(ocbWalletType)
            Exit Sub
        End If
    Case 1
        If Not aVB_WalletType(1).StaSale Then
            Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(1).TenderType) 'ไม่อนุญาตชำระเงินด้วย WeChatpay
            Call SP_OBJxSetFocus(ocbWalletType)
            Exit Sub
        End If
    Case 2
        If Not aVB_WalletType(2).StaSale Then
            Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(2).TenderType) 'ไม่อนุญาตชำระเงินด้วย Bluepay
            Call SP_OBJxSetFocus(ocbWalletType)
            Exit Sub
        End If
    Case 3
        If Not aVB_WalletType(3).StaSale Then
            Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , , aVB_WalletType(3).TenderType) 'ไม่อนุญาตชำระเงินด้วย Others
            Call SP_OBJxSetFocus(ocbWalletType)
            Exit Sub
        End If
    Case Else
        Call SP_MSGnShowing("tms_CN999049", nCS_Warning) 'กรุณาเลือก Wallet Type
        Call SP_OBJxSetFocus(ocbWalletType)
        Exit Sub
    End Select
    tW_WalletName = Trim(ocbWalletType.Text)
    wFunctionKB.tW_Rmk = tW_InVoiceID & ";" & tW_InVoiceDT & ";" & Trim(otbPartnerID.Text) & ";" & tW_Amount & ";" & tVB_WalletBrand & ";" & tVB_WalletStoreID & ";" & tVB_WalletDevideID & ";" & tW_WalletName & ";" & tW_WalletTransID & ";" & Trim(otbCode.Text) & ";"
    wFunctionKB.tW_TransID = Trim(otbPartnerID.Text)
    wFunctionKB.tW_WalletName = tW_WalletName
    Call ocmCancel2_Click
End Sub

Private Sub ocmOk_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call ocmCancel2_Click
    End If
End Sub

Private Sub opbWalletPay_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim nIndex As Integer
    Select Case Index
    Case 0 'Wait
        If KeyCode = vbKeyEscape Then
            Call W_PRCxManualMode
        End If
    Case 1 'OnLine
        Call ocmCancel_Click
    Case 2 'OffLine
        Call ocmCancel2_Click
    End Select
End Sub

Private Sub otbCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call ocmCancel_Click
    End Select
End Sub

Private Sub otbCode_KeyPress(KeyAscii As Integer)
Dim tName As String
On Error GoTo ErrHandle
If bW_FlagSendAmt Or (Not otbCode.Enabled) Then
    KeyAscii = 0
    Exit Sub
End If
If KeyAscii = 13 Then
    If otbCode.Text = "" Then Exit Sub
    If W_CHKxAlwWalletType(True, Trim(otbCode.Text), tName) Then '*Tao 60-09-05 Check Allow Wallet Type
        bW_FlagSendAmt = True
        otbCode.Locked = True
        nW_WalletMode = nW_CallFunc.Payment
        otmCallFunc.Enabled = True
    Else
        Call SP_MSGnShowing("tms_CN999048", nCS_Warning, , , tName)
    End If
End If
Exit Sub
ErrHandle:

End Sub

Private Sub otbPartnerID_Change()
    otbPartnerID.SelStart = Len(otbPartnerID.Text)
End Sub

Private Sub otbPartnerID_GotFocus()
    otbPartnerID.SelStart = 0
    otbPartnerID.SelLength = Len(otbPartnerID.Text)
End Sub

Private Sub otbPartnerID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbPartnerID.Text = ""
        Case vbKeyEscape
            Call ocmCancel2_Click
    End Select
End Sub

Private Sub otbPartnerID_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case 13
                Call ocmOk_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otmForm_Timer()
    nW_Timer = nW_Timer + 1
    If bW_HandleESC Then
        If nW_Timer > nW_TimeESC Then
            otmForm.Enabled = False
            bW_HandleESC = False
            opbWalletPay(1).Visible = True
            opbWalletPay(1).Left = 0
            opbWalletPay(1).ZOrder 0
            opbWalletPay(0).Visible = False
            ocmCancel.Cancel = True
            '*Tao 60-09-15 Enable Scanner
            If bVB_OPOSScannerEnable Then
                With oScanner
                    .Open tVB_Scanner
                    .Claim 0
                    .DeviceEnabled = True
                    .DataEventEnabled = True
                    .AutoDisable = True
                    .ClearInput
                    If .DeviceName = tSNR_TOUCH_SCANNER Then
                        .DecodeData = True
                    End If
                End With
            End If
            Call SP_OBJxSetFocus(otbCode)
        ElseIf nW_Timer <= nW_TimeESC Then
            olaCntEsc.Caption = nW_TimeESC - nW_Timer
        End If
    Else
        otmForm.Enabled = False
    End If
End Sub

Private Sub otmWallet_Timer()
Dim nCntDown As Integer
On Error Resume Next
    If bW_FlagSendAmt Then
        If bW_FlagApvSucess Then
            olaCntDown.Visible = False
            bW_FlagSendAmt = False
        Else
            nCntDown = nVB_WalletTimeout - nW_CntTime
            nW_CntTime = nW_CntTime + 1
            Select Case nCntDown
            Case 0
                olaCntDown.Visible = True
                olaCntDown.Caption = CStr(nCntDown)
                bW_CHKTimeOut = True
            Case Is > 0
                olaCntDown.Visible = True
                olaCntDown.Caption = CStr(nCntDown)
            Case Is < 0
                bW_CHKTimeOut = True
                olaCntDown.Visible = False
                If nCntDown < -10 Then bW_FlagSendAmt = False
            End Select
        End If
    Else
        olaCntDown.Visible = False
        otmWallet.Enabled = False
        otmCallFunc.Enabled = False
        If (Not bW_FlagApvSucess) And (nW_WalletMode <> nW_CallFunc.AutoReverse) Then 'ทำรายการไม่สำเร็จ
            Screen.MousePointer = vbDefault
'            Select Case Val(tVB_TransactionSub)
'                Case tEN_TrnReturn
'                    Call SP_MSGnShowing("tms_CN999044", nCS_Warning, , , tW_ErrDesc)
'                Case tEN_TrnIPV, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnIPVTaxExceptSale, tEN_TrnCancelVoucher, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVVIPService, tEN_TrnIPVReturn, tEN_TrnIPVRedeem
'                    Call SP_MSGnShowing("tms_CN999045", nCS_Warning, , , tW_ErrDesc)
'                Case Else
'                    Call SP_MSGnShowing("tms_CN999030", nCS_Warning, , , tW_ErrDesc)
'                End Select
        End If
        bW_Cancel = True
        Unload Me
    End If
End Sub

Private Sub W_SaveLogWallet(Optional ByVal bpInsert As Boolean = True)
'*Tao 60-08-29 Insert หรือ Update
Dim tSql As String
On Error GoTo ErrHandle
    If bpInsert Then
        tSql = "INSERT INTO TPSTLogWallet ("
        tSql = tSql & " FDDateIns"
        tSql = tSql & ",FTTimeIns"
        tSql = tSql & ",FTWhoIns"
        tSql = tSql & ",FDDateUpd"
        tSql = tSql & ",FTTimeUpd"
        tSql = tSql & ",FTWhoUpd"
        tSql = tSql & ",FTTmnNum"
        tSql = tSql & ",FTLogCode"
        tSql = tSql & ",FTReqType"
        tSql = tSql & ",FTStatus"
        tSql = tSql & ",FTOrderID"
        tSql = tSql & ",FTReqDT"
        tSql = tSql & ",FNReqAmt"
        tSql = tSql & ",FTToken"
        'tSql = tSql & ",FTOrigTransID"
        'tSql = tSql & ",FNRemainingAmt"
        tSql = tSql & ",FTBrand"
        tSql = tSql & ",FTStoreID"
        tSql = tSql & ",FTDeviceID"
        tSql = tSql & ",FTCashierID"
        tSql = tSql & ",FTResCode"
        tSql = tSql & ",FTErrCode"
        tSql = tSql & ",FTErrDesc"
        tSql = tSql & ",FTTransDT"
        'tSql = tSql & ",FTCancelDT"
        tSql = tSql & ",FTWalletName"
        tSql = tSql & ",FTWalletTransID"
        tSql = tSql & ",FTPartnerTransID"
        tSql = tSql & ",FTWPSTransID"
        tSql = tSql & ",FTWalletCurr"
        tSql = tSql & ",FNWalletCurrAmt"
        tSql = tSql & ",FNExchRate"
        tSql = tSql & ",FTExchRateUnit"
        'tSql = tSql & ",FTWPSRefundID"
        'tSql = tSql & ",FTRefundDT"
        'tSql = tSql & ",FNRefundAmt"
        tSql = tSql & ",FTReceiptText"
        tSql = tSql & ",FTResult"
        tSql = tSql & ",FTRefDoc"
        tSql = tSql & ",FTStaDoc"
        tSql = tSql & ") VALUES ("
        tSql = tSql & "CONVERT(VarChar(10), GETDATE(), 120)"
        tSql = tSql & ",CONVERT(VarChar(8), GETDATE(), 108)"
        tSql = tSql & ",N'" & tVB_CNUserName & "'"
        tSql = tSql & ",CONVERT(VarChar(10), GETDATE(), 120)"
        tSql = tSql & ",CONVERT(VarChar(8), GETDATE(), 108)"
        tSql = tSql & ",N'" & tVB_CNUserName & "'"
        tSql = tSql & ",'" & tVB_CNTerminalNum & "'"
        tSql = tSql & ",'" & tUT_LoginCode & "'"
        tSql = tSql & ",'" & aVB_WalletLog.ReqType & "'"
        tSql = tSql & ",'" & aVB_WalletLog.Status & "'"
        tSql = tSql & ",'" & aVB_WalletLog.OrderID & "'"
        tSql = tSql & ",'" & aVB_WalletLog.ReqDT & "'"
        tSql = tSql & "," & SP_VALcText2Double(aVB_WalletLog.ReqAmt) & ""
        tSql = tSql & ",'" & aVB_WalletLog.Token & "'"
        'tSql = tSql & ",'" & aVB_WalletLog.OrigTransID & "'"
        'tSql = tSql & "," & SP_VALcText2Double(aVB_WalletLog.RemainingAmt) & ""
        tSql = tSql & ",'" & aVB_WalletLog.Brand & "'"
        tSql = tSql & ",'" & aVB_WalletLog.StoreID & "'"
        tSql = tSql & ",'" & aVB_WalletLog.DeviceID & "'"
        tSql = tSql & ",'" & aVB_WalletLog.CashierID & "'"
        tSql = tSql & ",'" & aVB_WalletLog.ResCode & "'"
        tSql = tSql & ",'" & aVB_WalletLog.ErrCode & "'"
        tSql = tSql & ",'" & SP_STRtSiggleCode(aVB_WalletLog.ErrDesc) & "'"
        tSql = tSql & ",'" & aVB_WalletLog.TransDT & "'"
        'tSql = tSql & ",'" & aVB_WalletLog.CancelDT & "'"
        tSql = tSql & ",'" & aVB_WalletLog.WalletName & "'"
        tSql = tSql & ",'" & aVB_WalletLog.WalletTransID & "'"
        tSql = tSql & ",'" & aVB_WalletLog.PartnerTransID & "'"
        tSql = tSql & ",'" & aVB_WalletLog.WPSTransID & "'"
        tSql = tSql & ",'" & aVB_WalletLog.WalletCurr & "'"
        tSql = tSql & "," & SP_VALcText2Double(aVB_WalletLog.WalletCurrAmt) & ""
        tSql = tSql & "," & SP_VALcText2Double(aVB_WalletLog.ExchRate) & ""
        tSql = tSql & ",'" & aVB_WalletLog.ExchRateUnit & "'"
        'tSql = tSql & ",'" & aVB_WalletLog.WPSRefundID & "'"
        'tSql = tSql & ",'" & aVB_WalletLog.RefundDT & "'"
        'tSql = tSql & "," & SP_VALcText2Double(aVB_WalletLog.RefundAmt) & ""
        tSql = tSql & ",'" & SP_STRtSiggleCode(aVB_WalletLog.ReceiptText) & "'"
        tSql = tSql & ",'" & SP_STRtSiggleCode(aVB_WalletLog.Result) & "'"
        tSql = tSql & ",'" & aVB_WalletLog.RefDoc & "'"
        tSql = tSql & ",'" & aVB_WalletLog.StaDoc & "'"
        tSql = tSql & ")"
    Else
        tSql = "UPDATE TPSTLogWallet SET "
        tSql = tSql & " FTStatus='" & aVB_WalletLog.Status & "'"
        tSql = tSql & ",FTResCode='" & aVB_WalletLog.ResCode & "'"
        tSql = tSql & ",FTErrCode='" & aVB_WalletLog.ErrCode & "'"
        tSql = tSql & ",FTErrDesc='" & SP_STRtSiggleCode(aVB_WalletLog.ErrDesc) & "'"
        tSql = tSql & ",FTOrderID='" & aVB_WalletLog.OrderID & "'"
        tSql = tSql & ",FTTransDT='" & aVB_WalletLog.TransDT & "'"
        tSql = tSql & ",FTWalletTransID='" & aVB_WalletLog.WalletTransID & "'"
        tSql = tSql & ",FTPartnerTransID='" & aVB_WalletLog.PartnerTransID & "'"
        tSql = tSql & ",FTWPSTransID='" & aVB_WalletLog.WPSTransID & "'"
        tSql = tSql & ",FTWalletCurr='" & aVB_WalletLog.WalletCurr & "'"
        tSql = tSql & ",FNWalletCurrAmt='" & aVB_WalletLog.WalletCurrAmt & "'"
        tSql = tSql & ",FNExchRate='" & aVB_WalletLog.ExchRate & "'"
        tSql = tSql & ",FTExchRateUnit='" & aVB_WalletLog.ExchRateUnit & "'"
        tSql = tSql & ",FTReceiptText='" & SP_STRtSiggleCode(aVB_WalletLog.ReceiptText) & "'"
        tSql = tSql & ",FTResult='" & SP_STRtSiggleCode(aVB_WalletLog.Result) & "'"
        tSql = tSql & "," & UT_tLastUpdSql
        tSql = tSql & " WHERE FTOrderID='" & tW_InVoiceID & "' AND FTRefDoc='" & aVB_WalletLog.RefDoc & "'"
        Select Case nW_WalletMode
        Case nW_CallFunc.Payment
            tSql = tSql & " AND FTReqType='Sale' AND FTReqDT='" & tW_InVoiceDT & "'"
            tSql = tSql & " AND FTToken='" & aVB_WalletLog.Token & "'"
        Case nW_CallFunc.Void
            tSql = tSql & " AND FTReqType='Cancel' AND FTReqDT='" & tW_RefDT & "'"
        Case nW_CallFunc.AutoReverse
            tSql = tSql & " AND FTReqType='AutoReverse' AND FTReqDT='" & tW_RefDT & "'"
        End Select
    End If
    
    Call SP_SQLvExecute(tSql, , oDB.ocnOffLine) 'เก็บไว้ที่ Offline ก่อน
    Exit Sub
ErrHandle:
    On Error GoTo 0
    SP_TXTxWriteLogWalletPay ("Error Save Log WalletPay = " & Err.Description)
End Sub

Private Sub W_ClearLogWallet()
On Error Resume Next
With aVB_WalletLog
    .ReqType = ""
    .Status = ""
    .OrderID = ""
    .ReqDT = ""
    .ReqAmt = 0
    .OrigTransID = ""
    .RemainingAmt = 0
    .Token = ""
    .Brand = ""
    .StoreID = ""
    .DeviceID = ""
    .CashierID = ""
    .ResCode = ""
    .ErrCode = ""
    .ErrDesc = ""
    .TransDT = ""
    .CancelDT = ""
    .WalletName = ""
    .WalletTransID = ""
    .PartnerTransID = ""
    .WPSTransID = ""
    .WalletCurr = ""
    .WalletCurrAmt = 0
    .ExchRate = 0
    .ExchRateUnit = ""
    .WPSRefundID = ""
    .RefundDT = ""
    .RefundAmt = 0
    .ReceiptText = ""
    .Result = ""
    .RefDoc = ""
    .StaDoc = ""
End With
End Sub

Private Sub otmCallFunc_Timer()
Dim tTmpRes As String
Dim tTransID As String
Dim tPayTime As String
Dim tErrCode As String
On Error Resume Next
    
    otmCallFunc.Enabled = False
    olaInfo.Visible = False '*Tao 61-11-20 TK-ISS3000-711 TK ISSUE - TK Request - Alert message on AliPay pop-up
    DoEvents
    Select Case nW_WalletMode
    Case nW_CallFunc.Payment
SendPayment:
        ocmCheck.Visible = False
        ocmCheck.Enabled = False
        ocmCancel.Enabled = False
        opbResult(0).Visible = False
        opbResult(1).Visible = False
        olaResult.ForeColor = &H800000
        olaResult.Alignment = 2
        olaCntDown.FontSize = 12
        olaResult.Caption = cUT.UT_STRtToken(tW_Starting, ";", CInt(nVB_CNCutLang))
        olaResult.Visible = True
        olaCntDown.Visible = False
        olaCntDown.Alignment = 2
        olaCntDown.FontSize = 12
        DoEvents
            Screen.MousePointer = vbHourglass
            Select Case UCase(W_GETxPaymentRes(otbCode.Text, tTransID, tPayTime, tW_ErrDesc, tErrCode, tW_WalletTransID, tW_WalletName, tW_WPSTransID))
            Case "FAILED"
                Call W_SaveLogWallet(False)
                opbResult(1).Visible = True
                olaResult.ForeColor = &HFF&
                olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
                olaCntDown.Alignment = 2
                olaCntDown.FontSize = 10
                olaCntDown.Visible = True
                olaCntDown.Caption = tW_ErrDesc
                ocmCancel.Enabled = True
                If (InStr(1, UCase(tW_ErrDesc), "TIME OUT") > 0) Or (InStr(1, UCase(tW_ErrDesc), "TIMED OUT") > 0) Then 'time out ให้รอ User เลือกว่าจะ Check หรือ Cancel
                    bW_FlagSendAmt = False
                    Screen.MousePointer = vbDefault
                    ocmCheck.Visible = True
                    ocmCheck.Enabled = True
                    Call SP_OBJxSetFocus(ocmCheck)
                ElseIf (InStr(1, UCase(tW_ErrDesc), "INTERNET") > 0) Then 'internet down หรือ low speed ให้รอ User เลือกว่าจะ Check หรือ Cancel
                    bW_FlagSendAmt = False
                    Screen.MousePointer = vbDefault
                    ocmCheck.Visible = True
                    ocmCheck.Enabled = True
                    Call SP_OBJxSetFocus(ocmCheck)
                ElseIf (InStr(1, UCase(tW_ErrDesc), "NETWORK") > 0) Then 'network disconnect ให้รอ User เลือกว่าจะ Check หรือ Cancel
                    bW_FlagSendAmt = False
                    Screen.MousePointer = vbDefault
                    ocmCheck.Visible = True
                    ocmCheck.Enabled = True
                    Call SP_OBJxSetFocus(ocmCheck)
                ElseIf (InStr(1, UCase(tW_ErrDesc), "UNABLE TO CONNECT") > 0) Then 'Unable to connect to the remote server ให้รอ User เลือกว่าจะ Check หรือ Cancel
                    bW_FlagSendAmt = False
                    Screen.MousePointer = vbDefault
                    ocmCheck.Visible = True
                    ocmCheck.Enabled = True
                    Call SP_OBJxSetFocus(ocmCheck)
                Else
                    opbResult(1).Visible = True
                    olaResult.ForeColor = &HFF&
                    olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
                    olaCntDown.Alignment = 2
                    olaCntDown.FontSize = 10
                    olaCntDown.Visible = True
                    olaCntDown.Caption = tW_ErrDesc
                    '*Tao 60-09-07 Reverse เลยหรือไม่
                    If bW_NoWaitReverse Then
                        nW_Retry = 0
                        nW_WalletMode = nW_CallFunc.Reverse
                        otmCallFunc.Interval = 1000
                        otmCallFunc.Enabled = True
                    Else
                        bW_FlagSendAmt = False
                        Screen.MousePointer = vbDefault
                        ocmCheck.Visible = False
                        ocmCancel.Enabled = True
                        Call SP_OBJxSetFocus(ocmCancel)
                    End If
                End If
            Case "SUCCESS"
                Call W_SaveLogWallet(False)
                If W_CHKxAlwWalletType(False, tW_WalletName) Then
                    'bW_FlagSendAmt = False
                    'opbResult(0).Visible = True
                    'bW_FlagApvSucess = True
                    'olaResult.ForeColor = &H4000&
                    'olaResult.Caption = cUT.UT_STRtToken(tW_Success, ";", CInt(nVB_CNCutLang))
                    
                    '*Tao 61-08-09 ของเดิม
                    'wFunctionKB.tW_Rmk = tW_InVoiceID & ";" & tW_InVoiceDT & ";" & tTransID & ";" & tW_Amount & ";" & tVB_WalletBrand & ";" & tVB_WalletStoreID & ";" & tVB_WalletDevideID & ";" & tW_WalletName & ";" & tW_WalletTransID & ";" & Trim(otbCode.Text) & ";" & SP_STRtSiggleCode(aVB_WalletLog.ReceiptText)
                    'wFunctionKB.tW_TransID = tTransID
                    
                    '*Tao 61-08-09 TK-ISS3000-689 TK Issue - WPA for Alipay, update ‘FTSrcCardNo’ with WPS Trans ID
                    wFunctionKB.tW_Rmk = tW_InVoiceID & ";" & tW_InVoiceDT & ";" & tW_WPSTransID & ";" & tW_Amount & ";" & tVB_WalletBrand & ";" & tVB_WalletStoreID & ";" & tVB_WalletDevideID & ";" & tW_WalletName & ";" & tW_WalletTransID & ";" & Trim(otbCode.Text) & ";" & SP_STRtSiggleCode(aVB_WalletLog.ReceiptText)
                    If UCase(tW_WalletName) = "ALIPAY" Then
                        wFunctionKB.tW_TransID = tW_WPSTransID
                    Else
                        wFunctionKB.tW_TransID = IIf(tTransID = "", tW_WPSTransID, tTransID)
                    End If
                    '====================================================================================
                    
                    'wFunctionKB.tW_WPSTransID = tW_WPSTransID
                    wFunctionKB.tW_WalletName = tW_WalletName
                    wFunctionKB.tW_WalletTransID = tW_WalletTransID
                    wFunctionKB.tW_Date = tW_BizDate
                    '*Tao 60-09-06 ยังไม่ต้องส่ง Finalize ให้ไปส่งตอนจบบิล
'                    Screen.MousePointer = vbDefault
'                    otmWallet.Enabled = True
                    '*Tao 60-09-06 ส่ง Finalize เลย
                    nW_WalletMode = nW_CallFunc.Finalize
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                Else
                    opbResult(0).Visible = False
                    opbResult(1).Visible = True
                    olaResult.ForeColor = &HFF&
                    olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
                    olaCntDown.Alignment = 2
                    olaCntDown.FontSize = 10
                    olaCntDown.Visible = True
                    olaCntDown.Caption = SP_TXTnShowing("tms_CN999048") & " " & tW_WalletName 'ไม่อนุญาตชำระสินค้าด้วย tW_WalletName
                    nW_WalletMode = nW_CallFunc.Reverse
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                End If
            Case "PENDING", "UNKNOWN"
                Call W_SaveLogWallet(False)
                'olaCntDown.Visible = True
                nW_WalletMode = nW_CallFunc.Query
                bW_FlagSendAmt = False
'                opbResult(1).Visible = True
'                olaResult.ForeColor = &H800000
                olaResult.ForeColor = &HFF&
                olaResult.Alignment = 0
                olaResult.FontSize = 10
                olaResult.Caption = cUT.UT_STRtToken(tW_WaitConfirm, ";", CInt(nVB_CNCutLang))
                olaCntDown.Visible = True
                olaCntDown.Alignment = 0
                olaCntDown.FontSize = 10
                olaCntDown.Caption = cUT.UT_STRtToken(tW_WaitConfirm2, ";", CInt(nVB_CNCutLang))
                '*Tao 61-11-20 TK-ISS3000-711 TK ISSUE - TK Request - Alert message on AliPay pop-up
                olaInfo.Visible = True
                olaInfo.Alignment = 2
                olaInfo.FontSize = 10
                olaInfo.Caption = cUT.UT_STRtToken(tW_PendingInfo, ";", CInt(nVB_CNCutLang))
                '=====================================
                Screen.MousePointer = vbDefault
                ocmCancel.Enabled = True
                ocmCheck.Visible = True
                ocmCheck.Enabled = True
                Call SP_OBJxSetFocus(ocmCheck)
            Case Else
                Call W_SaveLogWallet(False)
                opbResult(1).Visible = True
                olaResult.ForeColor = &HFF&
                olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
                olaCntDown.Alignment = 2
                olaCntDown.FontSize = 10
                olaCntDown.Visible = True
                olaCntDown.Caption = tW_ErrDesc
                '*Tao 60-09-07 Reverse เลยหรือไม่
                If bW_NoWaitReverse Then
                    nW_Retry = 0
                    nW_WalletMode = nW_CallFunc.Reverse
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                Else
                    bW_FlagSendAmt = False
                    Screen.MousePointer = vbDefault
                    ocmCheck.Visible = False
                    ocmCancel.Enabled = True
                    Call SP_OBJxSetFocus(ocmCancel)
                End If
            End Select
'    Case nW_CallFunc.AutoReverse
'     '*Tao 60-05-15 Auto Reverse
'SendAutoReverse:
'        opbResult(0).Visible = False
'        opbResult(1).Visible = False
'        olaResult.ForeColor = &H800000
'        olaResult.Caption = cUT.UT_STRtToken(tW_Canceling, ";", CInt(nVB_CNCutLang))
'        olaResult.Visible = True
'        DoEvents
'            'otmWallet.Enabled = True
'            Select Case UCase(W_GETxAutoReverseRes(tW_ErrDesc, tErrCode))
'            Case "FAILED"
'                Call W_SaveLogWallet
'                If (InStr(1, UCase(tW_ErrDesc), UCase("refunded")) > 1) Or (UCase(tErrCode) = UCase("WS0020")) Then 'ถ้า cancel แล้วได้ response = "Already refunded" ให้ถือว่า Success
'                    aVB_WalletLog.Status = "SUCCESS"
'                    Call W_SaveLogWallet(False)
'                    otbCode.Text = ""
'                    opbResult(0).Visible = True
'                    bW_FlagApvSucess = True
'                    olaResult.ForeColor = &H4000&
'                    olaResult.Caption = cUT.UT_STRtToken(tW_Canceled, ";", CInt(nVB_CNCutLang))
'                    Call W_ClearLogWallet
'                ElseIf (InStr(1, UCase(tW_ErrDesc), UCase("Unvoidable")) > 1) Or (UCase(tErrCode) = UCase("ER0193")) Then 'ถ้า cancel แล้วได้ response = "Unvoidable trans status" ให้ถือว่า Success
'                    aVB_WalletLog.Status = "SUCCESS"
'                    Call W_SaveLogWallet(False)
'                    otbCode.Text = ""
'                    opbResult(0).Visible = True
'                    bW_FlagApvSucess = True
'                    olaResult.ForeColor = &H4000&
'                    olaResult.Caption = cUT.UT_STRtToken(tW_Canceled, ";", CInt(nVB_CNCutLang))
'                    Call W_ClearLogWallet
'                ElseIf (InStr(1, UCase(tW_ErrDesc), UCase("voided")) > 1) Then 'ถ้า cancel แล้วได้ response = "Transaction voided" ให้ถือว่า Success
'                    aVB_WalletLog.Status = "SUCCESS"
'                    Call W_SaveLogWallet(False)
'                    otbCode.Text = ""
'                    opbResult(0).Visible = True
'                    bW_FlagApvSucess = True
'                    olaResult.ForeColor = &H4000&
'                    olaResult.Caption = cUT.UT_STRtToken(tW_Canceled, ";", CInt(nVB_CNCutLang))
'                    Call W_ClearLogWallet
'                Else
'                    otbCode.Text = ""
'                    opbResult(1).Visible = True
'                    olaResult.ForeColor = &H000000FF&
'                    olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
'                    bW_FlagSendAmt = False
'                End If
'            Case "SUCCESS"
'                Call W_SaveLogWallet
'                otbCode.Text = ""
'                opbResult(0).Visible = True
'                bW_FlagApvSucess = True
'                olaResult.ForeColor = &H4000&
'                olaResult.Caption = cUT.UT_STRtToken(tW_Canceled, ";", CInt(nVB_CNCutLang))
'                Call W_ClearLogWallet
'            Case Else
'                Call W_SaveLogWallet
'                otbCode.Text = ""
'                opbResult(1).Visible = True
'                olaResult.ForeColor = &H000000FF&
'                olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
'                bW_FlagSendAmt = False
'            End Select
    Case nW_CallFunc.Reverse
SendReverse:
        ocmCheck.Enabled = False
        ocmCancel.Enabled = False
        opbResult(0).Visible = False
        opbResult(1).Visible = True
        'olaResult.ForeColor = &H800000
        olaResult.ForeColor = &HFF&
        olaResult.Alignment = 2
        'olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
        olaResult.Caption = cUT.UT_STRtToken(tW_Canceling, ";", CInt(nVB_CNCutLang))
        olaResult.Visible = True
        olaResult.FontSize = 12
        olaCntDown.Alignment = 2
        olaCntDown.FontSize = 10
        DoEvents
        Screen.MousePointer = vbHourglass
        Select Case UCase(W_GETxReverseRes(tW_ErrDesc, tErrCode))
        Case "FAILED"
            Call W_SaveLogWallet
            If (InStr(1, UCase(tW_ErrDesc), "TIME OUT") > 0) Or (InStr(1, UCase(tW_ErrDesc), "TIMED OUT") > 0) Then 'time out ให้ส่ง Reverse ใหม่
                If (nW_Retry < nW_RevRetry) And bW_NoWaitReverse Then
                    Call W_PRCxRestartWPA
                    nW_Retry = nW_Retry + 1
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                    Exit Sub
                Else
                    opbResult(0).Visible = False
                    opbResult(1).Visible = True
                    olaResult.ForeColor = &HFF&
                    olaResult.Caption = cUT.UT_STRtToken(tW_ReverseFail, ";", CInt(nVB_CNCutLang))
                    olaCntDown.Alignment = 2
                    olaCntDown.Visible = True
                    olaCntDown.Caption = tW_ErrDesc
                End If
            ElseIf (InStr(1, UCase(tW_ErrDesc), "INTERNET") > 0) Then 'internet down หรือ low speed ให้ส่ง Reverse ใหม่
                If (nW_Retry < nW_RevRetry) And bW_NoWaitReverse Then
                    Call W_PRCxRestartWPA
                    nW_Retry = nW_Retry + 1
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                    Exit Sub
                Else
                    opbResult(0).Visible = False
                    opbResult(1).Visible = True
                    olaResult.ForeColor = &HFF&
                    olaResult.Caption = cUT.UT_STRtToken(tW_ReverseFail, ";", CInt(nVB_CNCutLang))
                    olaCntDown.Alignment = 2
                    olaCntDown.Visible = True
                    olaCntDown.Caption = tW_ErrDesc
                End If
            ElseIf (InStr(1, UCase(tW_ErrDesc), "NETWORK") > 0) Then 'network disconnect ให้ส่ง Reverse ใหม่
                If (nW_Retry < nW_RevRetry) And bW_NoWaitReverse Then
                    Call W_PRCxRestartWPA
                    nW_Retry = nW_Retry + 1
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                    Exit Sub
                Else
                    opbResult(0).Visible = False
                    opbResult(1).Visible = True
                    olaResult.ForeColor = &HFF&
                    olaResult.Caption = cUT.UT_STRtToken(tW_ReverseFail, ";", CInt(nVB_CNCutLang))
                    olaCntDown.Alignment = 2
                    olaCntDown.Visible = True
                    olaCntDown.Caption = tW_ErrDesc
                End If
            ElseIf (InStr(1, UCase(tW_ErrDesc), "UNABLE TO CONNECT") > 0) Then 'Unable to connect to the remote server ให้ส่ง Reverse ใหม่
                If (nW_Retry < nW_RevRetry) And bW_NoWaitReverse Then
                    Call W_PRCxRestartWPA
                    nW_Retry = nW_Retry + 1
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                    Exit Sub
                Else
                    opbResult(0).Visible = False
                    opbResult(1).Visible = True
                    olaResult.ForeColor = &HFF&
                    olaResult.Caption = cUT.UT_STRtToken(tW_ReverseFail, ";", CInt(nVB_CNCutLang))
                    olaCntDown.Alignment = 2
                    olaCntDown.Visible = True
                    olaCntDown.Caption = tW_ErrDesc
                End If
            ElseIf (InStr(1, UCase(tW_ErrDesc), UCase("voided")) > 0) Then 'ถ้า Reverse แล้วได้ response = "Transaction voided" ให้ถือว่า Success
                aVB_WalletLog.Status = "SUCCESS"
                Call W_SaveLogWallet(False)
            End If
        Case "SUCCESS"
            Call W_SaveLogWallet
            opbResult(0).Visible = False
            opbResult(1).Visible = True
            olaResult.ForeColor = &HFF&
            olaResult.Caption = cUT.UT_STRtToken(tW_Canceled, ";", CInt(nVB_CNCutLang))
            olaCntDown.Alignment = 2
            olaCntDown.Visible = True
            olaCntDown.Caption = tW_ErrDesc
        Case Else
            Call W_SaveLogWallet
            opbResult(0).Visible = False
            opbResult(1).Visible = True
            olaResult.ForeColor = &HFF&
            olaResult.Caption = cUT.UT_STRtToken(tW_ReverseFail, ";", CInt(nVB_CNCutLang))
            olaCntDown.Alignment = 2
            olaCntDown.Visible = True
            olaCntDown.Caption = tW_ErrDesc
        End Select
        Call W_ClearLogWallet
        bW_FlagSendAmt = False
        wFunctionKB.tW_Rmk = ""
        Screen.MousePointer = vbDefault
        If bW_NoWaitReverse Then
            otmWallet.Enabled = True
        Else
            ocmCancel.Enabled = True
            Call SP_OBJxSetFocus(ocmCancel)
        End If
    Case nW_CallFunc.Finalize
SendFinalize:
        ocmCheck.Enabled = False
        ocmCancel.Enabled = False
        olaResult.Alignment = 2
        olaResult.FontSize = 12
        olaCntDown.Alignment = 2
        olaCntDown.FontSize = 10
        DoEvents
        Screen.MousePointer = vbHourglass
        Select Case UCase(W_GETxFinalRes(tW_ErrDesc, tErrCode))
        Case "FAILED"
            Call W_SaveLogWallet
            If (InStr(1, UCase(tW_ErrDesc), "TIME OUT") > 0) Or (InStr(1, UCase(tW_ErrDesc), "TIMED OUT") > 0) Then 'time out ให้ส่ง Finalize ใหม่
                If nW_Retry < nW_RevRetry Then
                    Call W_PRCxRestartWPA
                    nW_Retry = nW_Retry + 1
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                Else
                    opbResult(0).Visible = False
                    opbResult(1).Visible = True
                    olaResult.ForeColor = &HFF&
                    olaResult.Caption = cUT.UT_STRtToken(tW_FinalizeFail, ";", CInt(nVB_CNCutLang))
                    olaCntDown.Visible = True
                    olaCntDown.Caption = tW_ErrDesc
                    '*Tao 60-09-07 Reverse เลยหรือไม่
                    If bW_NoWaitReverse Then
                        nW_Retry = 0
                        nW_WalletMode = nW_CallFunc.Reverse
                        otmCallFunc.Interval = 1000
                        otmCallFunc.Enabled = True
                    Else
                        bW_FlagSendAmt = False
                        Screen.MousePointer = vbDefault
                        ocmCheck.Visible = False
                        ocmCancel.Enabled = True
                        Call SP_OBJxSetFocus(ocmCancel)
                    End If
                End If
            ElseIf (InStr(1, UCase(tW_ErrDesc), "INTERNET") > 0) Then 'internet down หรือ low speed ให้ส่ง Finalize ใหม่
                If nW_Retry < nW_RevRetry Then
                    Call W_PRCxRestartWPA
                    nW_Retry = nW_Retry + 1
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                Else
                    opbResult(0).Visible = False
                    opbResult(1).Visible = True
                    olaResult.ForeColor = &HFF&
                    olaResult.Caption = cUT.UT_STRtToken(tW_FinalizeFail, ";", CInt(nVB_CNCutLang))
                    olaCntDown.Visible = True
                    olaCntDown.Caption = tW_ErrDesc
                    '*Tao 60-09-07 Reverse เลยหรือไม่
                    If bW_NoWaitReverse Then
                        nW_Retry = 0
                        nW_WalletMode = nW_CallFunc.Reverse
                        otmCallFunc.Interval = 1000
                        otmCallFunc.Enabled = True
                    Else
                        bW_FlagSendAmt = False
                        Screen.MousePointer = vbDefault
                        ocmCheck.Visible = False
                        ocmCancel.Enabled = True
                        Call SP_OBJxSetFocus(ocmCancel)
                    End If
                End If
            ElseIf (InStr(1, UCase(tW_ErrDesc), "NETWORK") > 0) Then 'network disconnect ให้ส่ง Finalize ใหม่
                If nW_Retry < nW_RevRetry Then
                    Call W_PRCxRestartWPA
                    nW_Retry = nW_Retry + 1
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                Else
                    opbResult(0).Visible = False
                    opbResult(1).Visible = True
                    olaResult.ForeColor = &HFF&
                    olaResult.Caption = cUT.UT_STRtToken(tW_FinalizeFail, ";", CInt(nVB_CNCutLang))
                    olaCntDown.Visible = True
                    olaCntDown.Caption = tW_ErrDesc
                    '*Tao 60-09-07 Reverse เลยหรือไม่
                    If bW_NoWaitReverse Then
                        nW_Retry = 0
                        nW_WalletMode = nW_CallFunc.Reverse
                        otmCallFunc.Interval = 1000
                        otmCallFunc.Enabled = True
                    Else
                        bW_FlagSendAmt = False
                        Screen.MousePointer = vbDefault
                        ocmCheck.Visible = False
                        ocmCancel.Enabled = True
                        Call SP_OBJxSetFocus(ocmCancel)
                    End If
                End If
            ElseIf (InStr(1, UCase(tW_ErrDesc), "UNABLE TO CONNECT") > 0) Then 'Unable to connect to the remote server ให้ส่ง Finalize ใหม่
                If nW_Retry < nW_RevRetry Then
                    Call W_PRCxRestartWPA
                    nW_Retry = nW_Retry + 1
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                Else
                    opbResult(0).Visible = False
                    opbResult(1).Visible = True
                    olaResult.ForeColor = &HFF&
                    olaResult.Caption = cUT.UT_STRtToken(tW_FinalizeFail, ";", CInt(nVB_CNCutLang))
                    olaCntDown.Visible = True
                    olaCntDown.Caption = tW_ErrDesc
                    '*Tao 60-09-07 Reverse เลยหรือไม่
                    If bW_NoWaitReverse Then
                        nW_Retry = 0
                        nW_WalletMode = nW_CallFunc.Reverse
                        otmCallFunc.Interval = 1000
                        otmCallFunc.Enabled = True
                    Else
                        bW_FlagSendAmt = False
                        Screen.MousePointer = vbDefault
                        ocmCheck.Visible = False
                        ocmCancel.Enabled = True
                        Call SP_OBJxSetFocus(ocmCancel)
                    End If
                End If
            Else
                opbResult(0).Visible = False
                opbResult(1).Visible = True
                olaResult.ForeColor = &HFF&
                olaResult.Caption = cUT.UT_STRtToken(tW_FinalizeFail, ";", CInt(nVB_CNCutLang))
                olaCntDown.Visible = True
                olaCntDown.Caption = tW_ErrDesc
                '*Tao 60-09-07 Reverse เลยหรือไม่
                If bW_NoWaitReverse Then
                    nW_Retry = 0
                    nW_WalletMode = nW_CallFunc.Reverse
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                Else
                    bW_FlagSendAmt = False
                    Screen.MousePointer = vbDefault
                    ocmCheck.Visible = False
                    ocmCancel.Enabled = True
                    Call SP_OBJxSetFocus(ocmCancel)
                End If
            End If
        Case "SUCCESS"
            Call W_SaveLogWallet
            bW_FlagSendAmt = False
            opbResult(0).Visible = True
            opbResult(1).Visible = False
            bW_FlagApvSucess = True
            olaResult.ForeColor = &H4000&
            olaResult.Caption = cUT.UT_STRtToken(tW_Success, ";", CInt(nVB_CNCutLang))
            olaCntDown.Visible = False
            wFunctionKB.tW_AppCode = tW_InVoiceDT
            Screen.MousePointer = vbDefault
            otmWallet.Enabled = True
        Case Else
            Call W_SaveLogWallet
            opbResult(0).Visible = False
            opbResult(1).Visible = True
            olaResult.ForeColor = &HFF&
            olaResult.Caption = cUT.UT_STRtToken(tW_FinalizeFail, ";", CInt(nVB_CNCutLang))
            olaCntDown.Visible = True
            olaCntDown.Caption = tW_ErrDesc
            '*Tao 60-09-07 Reverse เลยหรือไม่
            If bW_NoWaitReverse Then
                nW_Retry = 0
                nW_WalletMode = nW_CallFunc.Reverse
                otmCallFunc.Interval = 1000
                otmCallFunc.Enabled = True
            Else
                bW_FlagSendAmt = False
                Screen.MousePointer = vbDefault
                ocmCheck.Visible = False
                ocmCancel.Enabled = True
                Call SP_OBJxSetFocus(ocmCancel)
            End If
        End Select
    Case nW_CallFunc.Query
SendQuery:
        ocmCheck.Enabled = False
        ocmCancel.Enabled = False
        opbResult(0).Visible = False
        opbResult(1).Visible = False
        olaResult.ForeColor = &H800000
        olaResult.Alignment = 2
        olaResult.Caption = cUT.UT_STRtToken(tW_Starting, ";", CInt(nVB_CNCutLang))
        olaResult.Visible = True
        olaResult.FontSize = 12
        olaCntDown.Visible = False
        olaCntDown.Alignment = 2
        olaCntDown.FontSize = 12
        DoEvents
        Screen.MousePointer = vbHourglass
        Select Case UCase(W_GETxInQuiryRes(tTransID, tPayTime, tW_ErrDesc, tErrCode, tW_WalletTransID, tW_WalletName, tW_WPSTransID))
        Case "FAILED"
            Call W_SaveLogWallet
            If (InStr(1, UCase(tW_ErrDesc), UCase("Trans not found")) > 0) Or (UCase(tErrCode) = UCase("WS0012")) Then 'ถ้า query แล้วได้ response = "Trans not found" ให้ส่ง payment ไปใหม่
                nW_WalletMode = nW_CallFunc.Payment
                otmCallFunc.Interval = 1000
                otmCallFunc.Enabled = True
            ElseIf (InStr(1, UCase(tW_ErrDesc), "TIME OUT") > 0) Or (InStr(1, UCase(tW_ErrDesc), "TIMED OUT") > 0) Then
                bW_FlagSendAmt = False
                opbResult(1).Visible = True
                olaResult.ForeColor = &HFF&
                olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
                olaCntDown.Alignment = 2
                olaCntDown.FontSize = 10
                olaCntDown.Visible = True
                olaCntDown.Caption = tW_ErrDesc
                Screen.MousePointer = vbDefault
                ocmCancel.Enabled = True
                ocmCheck.Visible = True
                ocmCheck.Enabled = True
                Call SP_OBJxSetFocus(ocmCheck)
            ElseIf (InStr(1, UCase(tW_ErrDesc), "INTERNET") > 0) Then 'internet down หรือ low speed ให้ส่ง query ไปใหม่
                bW_FlagSendAmt = False
                opbResult(1).Visible = True
                olaResult.ForeColor = &HFF&
                olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
                olaCntDown.Alignment = 2
                olaCntDown.FontSize = 10
                olaCntDown.Visible = True
                olaCntDown.Caption = tW_ErrDesc
                Screen.MousePointer = vbDefault
                ocmCancel.Enabled = True
                ocmCheck.Visible = True
                ocmCheck.Enabled = True
                Call SP_OBJxSetFocus(ocmCheck)
            ElseIf (InStr(1, UCase(tW_ErrDesc), "NETWORK") > 0) Then 'network disconnect ให้ส่ง query ไปใหม่
                bW_FlagSendAmt = False
                opbResult(1).Visible = True
                olaResult.ForeColor = &HFF&
                olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
                olaCntDown.Alignment = 2
                olaCntDown.FontSize = 10
                olaCntDown.Visible = True
                olaCntDown.Caption = tW_ErrDesc
                Screen.MousePointer = vbDefault
                ocmCancel.Enabled = True
                ocmCheck.Visible = True
                ocmCheck.Enabled = True
                Call SP_OBJxSetFocus(ocmCheck)
            ElseIf (InStr(1, UCase(tW_ErrDesc), "UNABLE TO CONNECT") > 0) Then 'Unable to connect to the remote server ให้รอ User เลือกว่าจะ Check หรือ Cancel
                bW_FlagSendAmt = False
                opbResult(1).Visible = True
                olaResult.ForeColor = &HFF&
                olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
                olaCntDown.Alignment = 2
                olaCntDown.FontSize = 10
                olaCntDown.Visible = True
                olaCntDown.Caption = tW_ErrDesc
                Screen.MousePointer = vbDefault
                ocmCancel.Enabled = True
                ocmCheck.Visible = True
                ocmCheck.Enabled = True
                Call SP_OBJxSetFocus(ocmCheck)
            Else
                opbResult(1).Visible = True
                olaResult.ForeColor = &HFF&
                olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
                olaCntDown.Alignment = 2
                olaCntDown.FontSize = 10
                olaCntDown.Visible = True
                olaCntDown.Caption = tW_ErrDesc
                '*Tao 60-09-07 Reverse เลยหรือไม่
                If bW_NoWaitReverse Then
                    nW_Retry = 0
                    nW_WalletMode = nW_CallFunc.Reverse
                    otmCallFunc.Interval = 1000
                    otmCallFunc.Enabled = True
                Else
                    bW_FlagSendAmt = False
                    Screen.MousePointer = vbDefault
                    ocmCheck.Visible = False
                    ocmCancel.Enabled = True
                    Call SP_OBJxSetFocus(ocmCancel)
                End If
            End If
        Case "SUCCESS"
            Call W_SaveLogWallet
            If W_CHKxAlwWalletType(False, tW_WalletName) Then
                'bW_FlagSendAmt = False
                'opbResult(0).Visible = True
                'bW_FlagApvSucess = True
                'olaResult.ForeColor = &H4000&
                'olaResult.Caption = cUT.UT_STRtToken(tW_Success, ";", CInt(nVB_CNCutLang))
                
                '*Tao 61-08-09 ของเดิม
                'wFunctionKB.tW_Rmk = tW_InVoiceID & ";" & tW_InVoiceDT & ";" & tTransID & ";" & tW_Amount & ";" & tVB_WalletBrand & ";" & tVB_WalletStoreID & ";" & tVB_WalletDevideID & ";" & tW_WalletName & ";" & tW_WalletTransID & ";" & Trim(otbCode.Text) & ";" & SP_STRtSiggleCode(aVB_WalletLog.ReceiptText)
                'wFunctionKB.tW_TransID = tTransID
                
                '*Tao 61-08-09 TK-ISS3000-689 TK Issue - WPA for Alipay, update ‘FTSrcCardNo’ with WPS Trans ID
                wFunctionKB.tW_Rmk = tW_InVoiceID & ";" & tW_InVoiceDT & ";" & tW_WPSTransID & ";" & tW_Amount & ";" & tVB_WalletBrand & ";" & tVB_WalletStoreID & ";" & tVB_WalletDevideID & ";" & tW_WalletName & ";" & tW_WalletTransID & ";" & Trim(otbCode.Text) & ";" & SP_STRtSiggleCode(aVB_WalletLog.ReceiptText)
                If UCase(tW_WalletName) = "ALIPAY" Then
                    wFunctionKB.tW_TransID = tW_WPSTransID
                Else
                    wFunctionKB.tW_TransID = IIf(tTransID = "", tW_WPSTransID, tTransID)
                End If
                '====================================================================================
                
                'wFunctionKB.tW_WPSTransID = tW_WPSTransID
                wFunctionKB.tW_WalletName = tW_WalletName
                wFunctionKB.tW_WalletTransID = tW_WalletTransID
                wFunctionKB.tW_Date = tW_BizDate
                '*Tao 60-09-06 ยังไม่ต้องส่ง Finalize ให้ไปส่งตอนจบบิล
'                Screen.MousePointer = vbDefault
'                otmWallet.Enabled = True
                '*Tao 60-09-06 ส่ง Finalize เลย
                nW_WalletMode = nW_CallFunc.Finalize
                otmCallFunc.Interval = 1000
                otmCallFunc.Enabled = True
            Else
                opbResult(0).Visible = False
                opbResult(1).Visible = True
                olaResult.ForeColor = &HFF&
                olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
                olaCntDown.Alignment = 2
                olaCntDown.FontSize = 10
                olaCntDown.Visible = True
                olaCntDown.Caption = SP_TXTnShowing("tms_CN999048") & " " & tW_WalletName 'ไม่อนุญาตชำระสินค้าด้วย tW_WalletName
                nW_WalletMode = nW_CallFunc.Reverse
                otmCallFunc.Interval = 1000
                otmCallFunc.Enabled = True
            End If
        Case "PENDING", "UNKNOWN"
            Call W_SaveLogWallet
            bW_FlagSendAmt = False
            'opbResult(1).Visible = True
            'olaResult.ForeColor = &H800000
            olaResult.ForeColor = &HFF&
            olaResult.Alignment = 0
            olaResult.FontSize = 10
            olaResult.Caption = cUT.UT_STRtToken(tW_WaitConfirm, ";", CInt(nVB_CNCutLang))
            olaCntDown.Alignment = 0
            olaCntDown.FontSize = 10
            olaCntDown.Visible = True
            olaCntDown.Caption = cUT.UT_STRtToken(tW_WaitConfirm2, ";", CInt(nVB_CNCutLang))
            Screen.MousePointer = vbDefault
            ocmCancel.Enabled = True
            ocmCheck.Visible = True
            ocmCheck.Enabled = True
            Call SP_OBJxSetFocus(ocmCheck)
        Case Else
            Call W_SaveLogWallet
            opbResult(1).Visible = True
            olaResult.ForeColor = &HFF&
            olaResult.Caption = cUT.UT_STRtToken(tW_Fail, ";", CInt(nVB_CNCutLang))
            olaCntDown.Alignment = 2
            olaCntDown.FontSize = 10
            olaCntDown.Visible = True
            olaCntDown.Caption = tW_ErrDesc
            '*Tao 60-09-07 Reverse เลยหรือไม่
            If bW_NoWaitReverse Then
                nW_Retry = 0
                nW_WalletMode = nW_CallFunc.Reverse
                otmCallFunc.Interval = 1000
                otmCallFunc.Enabled = True
            Else
                bW_FlagSendAmt = False
                Screen.MousePointer = vbDefault
                ocmCheck.Visible = False
                ocmCancel.Enabled = True
                Call SP_OBJxSetFocus(ocmCancel)
            End If
        End Select
    Case nW_CallFunc.Void
SendVoid:
        Screen.MousePointer = vbHourglass
        olaResult.ForeColor = &H800000
        olaResult.Alignment = 2
        olaResult.Caption = cUT.UT_STRtToken(tW_VoidStarting, ";", CInt(nVB_CNCutLang))
        olaResult.Visible = True
        olaCntDown.Visible = False
        olaCntDown.Alignment = 2
        DoEvents
        Select Case UCase(W_GETxVoidRes(tTransID, tPayTime, tW_ErrDesc, tErrCode))
        Case "FAILED"
            If (InStr(1, UCase(tW_ErrDesc), UCase("voided")) > 0) Then 'ถ้า Void แล้วได้ response = "Transaction voided" ให้ถือว่า Success
                opbResult(0).Visible = True
                aVB_WalletLog.Status = "SUCCESS"
            Else
                opbResult(1).Visible = True
            End If
        Case "SUCCESS"
            opbResult(0).Visible = True
        Case Else
            opbResult(1).Visible = True
        End Select
        Call W_SaveLogWallet(False)
        If aVB_WalletLog.Status = "SUCCESS" Then
            opbResult(0).Visible = True
            bW_FlagSendAmt = False
            bW_FlagApvSucess = True
            olaResult.ForeColor = &H4000&
            olaResult.Caption = cUT.UT_STRtToken(tW_Voided, ";", CInt(nVB_CNCutLang))
            wFunctionKB.tW_Rmk = tW_InVoiceID & ";" & tW_RefDT & ";" & tTransID & ";" & tW_RefAmount & ";" & tW_RefBrand & ";" & tW_RefStoreID & ";" & tW_RefDeviceID & ";" & tW_RefWalletName & ";" & tW_RefWalletTransID & ";"
            wFunctionKB.tW_TransID = tTransID
            wFunctionKB.tW_WalletName = tW_RefWalletName
            wFunctionKB.tW_WalletTransID = tW_RefWalletTransID
        Else
            Call W_ClearLogWallet
            opbResult(1).Visible = True
            bW_FlagSendAmt = False
            olaResult.ForeColor = &HFF&
            olaResult.Caption = cUT.UT_STRtToken(tW_VoidFail, ";", CInt(nVB_CNCutLang))
            olaCntDown.Visible = True
            olaCntDown.Caption = tW_ErrDesc
        End If
        Screen.MousePointer = vbDefault
        otmWallet.Enabled = True
    Case nW_CallFunc.Refund
SendRefund:
        Screen.MousePointer = vbHourglass
        olaResult.ForeColor = &H800000
        olaResult.Alignment = 2
        olaResult.Caption = cUT.UT_STRtToken(tW_RefundStarting, ";", CInt(nVB_CNCutLang))
        olaResult.Visible = True
        olaCntDown.Visible = False
        olaCntDown.Alignment = 2
        DoEvents
        Select Case UCase(W_GETxRefundRes(tTransID, tPayTime, tW_ErrDesc, tErrCode))
        Case "FAILED"
            opbResult(1).Visible = True
        Case "SUCCESS"
            opbResult(0).Visible = True
        Case Else
            opbResult(1).Visible = True
        End Select
        Call W_SaveLogWallet(False)
        If aVB_WalletLog.Status = "SUCCESS" Then
            opbResult(0).Visible = True
            bW_FlagSendAmt = False
            bW_FlagApvSucess = True
            olaResult.ForeColor = &H4000&
            olaResult.Caption = cUT.UT_STRtToken(tW_Refunded, ";", CInt(nVB_CNCutLang))
            wFunctionKB.tW_Rmk = tW_InVoiceID & ";" & tW_RefDT & ";" & tTransID & ";" & tW_RefAmount & ";" & tW_RefBrand & ";" & tW_RefStoreID & ";" & tW_RefDeviceID & ";" & tW_RefWalletName & ";" & tW_RefWalletTransID & ";"
            wFunctionKB.tW_TransID = tTransID
            wFunctionKB.tW_WalletName = tW_RefWalletName
            wFunctionKB.tW_WalletTransID = tW_RefWalletTransID
        Else
            Call W_ClearLogWallet
            opbResult(1).Visible = True
            bW_FlagSendAmt = False
            olaResult.ForeColor = &HFF&
            olaResult.Caption = cUT.UT_STRtToken(tW_RefundFail, ";", CInt(nVB_CNCutLang))
            olaCntDown.Visible = True
            olaCntDown.Caption = tW_ErrDesc
        End If
        Screen.MousePointer = vbDefault
        otmWallet.Enabled = True
    End Select
End Sub

Private Function W_GETxPaymentRes(ByVal ptQRCode As String, Optional ByRef ptTransID As String, Optional ByRef ptPayTime As String, Optional ByRef ptErrDesc As String, Optional ByRef ptErrCode As String, Optional ByRef ptWalletTransID As String, Optional ByRef ptWalletName As String, Optional ByRef ptWPSTransID As String) As String
    '*Tao 60-08-29 WPA WalletPay Sale
    
    Dim oCallAPI As Object
    Dim oWPAReq As Object
    Dim oWPARes As Object
    
    On Error Resume Next
    
    Set oCallAPI = CreateObject("ADACallAPI.cAdaCallAPI")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Payment = " & Err.Description)
    End If
    If oCallAPI Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    Set oWPAReq = CreateObject("ADACallAPI.cmlWPAReq")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Payment = " & Err.Description)
    End If
    If oWPAReq Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    
    tW_InVoiceDT = CStr(Format(Now, "yyyyMMddhhmmss"))
    If tW_RefDocDT = "" Then tW_RefDocDT = tW_InVoiceDT
    If tW_BizDate = "" Then tW_BizDate = CStr(Format(Now, "yyyyMMdd"))
    
    oWPAReq.ReqType = "sale"
    oWPAReq.BizDate = tW_BizDate
    oWPAReq.Brand = tVB_WalletBrand
    oWPAReq.StoreID = tVB_WalletStoreID
    oWPAReq.DeviceID = tVB_WalletDevideID
    oWPAReq.CashierID = tVB_CNUserCode
    oWPAReq.Amt = Format(CDbl(tW_Amount), "########0.00")
    oWPAReq.Token = ptQRCode
    oWPAReq.ReqDT = tW_InVoiceDT
    oWPAReq.OrderID = tW_InVoiceID
    
    oCallAPI.DllKey = tVB_DllKey
    oCallAPI.DllPass = tVB_DllPass
    oCallAPI.HostURL = tVB_WalletHostURL
    oCallAPI.HostTimeout = nVB_WalletTimeout
    oCallAPI.SecurityProType = nW_SecurityProtocolType.Tls
    Call SP_TXTxWriteLogWalletPay("Start Payment " & tW_InVoiceID & ";" & Format(CDbl(tW_Amount), "########0.00"))
    
    '*Tao 60-08-29 Save Log WalletPay ก่อนส่ง Req
    aVB_WalletLog.ReqType = "Sale"
    aVB_WalletLog.Status = ""
    aVB_WalletLog.OrderID = oWPAReq.OrderID
    aVB_WalletLog.ReqDT = oWPAReq.ReqDT
    aVB_WalletLog.ReqAmt = oWPAReq.Amt
    aVB_WalletLog.OrigTransID = ""
    aVB_WalletLog.RemainingAmt = 0
    aVB_WalletLog.Token = oWPAReq.Token
    aVB_WalletLog.Brand = oWPAReq.Brand
    aVB_WalletLog.StoreID = oWPAReq.StoreID
    aVB_WalletLog.DeviceID = oWPAReq.DeviceID
    aVB_WalletLog.CashierID = oWPAReq.CashierID
    aVB_WalletLog.ResCode = ""
    aVB_WalletLog.ErrCode = ""
    aVB_WalletLog.ErrDesc = ""
    aVB_WalletLog.TransDT = ""
    aVB_WalletLog.CancelDT = ""
    aVB_WalletLog.WalletName = ""
    aVB_WalletLog.WalletTransID = ""
    aVB_WalletLog.PartnerTransID = ""
    aVB_WalletLog.WPSTransID = ""
    aVB_WalletLog.WalletCurr = ""
    aVB_WalletLog.WalletCurrAmt = 0
    aVB_WalletLog.ExchRate = 0
    aVB_WalletLog.ExchRateUnit = ""
    aVB_WalletLog.WPSRefundID = ""
    aVB_WalletLog.RefundDT = ""
    aVB_WalletLog.RefundAmt = 0
    aVB_WalletLog.ReceiptText = ""
    aVB_WalletLog.Result = ""
    aVB_WalletLog.RefDoc = tVB_CNStoreCode & "-" & tVB_CNTerminalNum & tVB_Running & "-" & tW_RefDocDT
    aVB_WalletLog.StaDoc = ""
    Call W_SaveLogWallet
    '====================================
    
    DoEvents
    Set oWPARes = oCallAPI.C_WPAoRequest(oWPAReq)
    If Err Then Call SP_TXTxWriteLogWalletPay("Error Payment = " & Err.Description)
    If oWPARes Is Nothing Then
        W_GETxPaymentRes = ""
    Else
        W_GETxPaymentRes = oWPARes.StatusCode
        ptTransID = oWPARes.PartnerTransID
        ptPayTime = oWPARes.TransDT
        ptErrDesc = oWPARes.ErrMsg
        ptErrCode = oCallAPI.RespCode
        ptWalletTransID = oWPARes.WalletTransID
        ptWalletName = oWPARes.WalletName
        ptWPSTransID = oWPARes.WPSTransID
        Call SP_TXTxWriteLogWalletPay("Response Payment = " & W_GETxPaymentRes & ";" & ptWalletName & ";" & ptTransID & ";" & ptWPSTransID & ";" & ptPayTime & ";" & ptErrDesc)
    End If
    
    aVB_WalletLog.Status = W_GETxPaymentRes
    aVB_WalletLog.WalletName = oWPARes.WalletName
    aVB_WalletLog.WalletTransID = oWPARes.WalletTransID
    aVB_WalletLog.PartnerTransID = oWPARes.PartnerTransID
    aVB_WalletLog.WPSTransID = oWPARes.WPSTransID
    aVB_WalletLog.WalletCurr = oWPAReq.WalletCurr
    aVB_WalletLog.WalletCurrAmt = oWPARes.WalletCurrAmt
    aVB_WalletLog.ResCode = oWPARes.RespCode
    aVB_WalletLog.ErrCode = oWPARes.ResponseCode
    aVB_WalletLog.ErrDesc = ptErrDesc
    aVB_WalletLog.TransDT = ptPayTime
    aVB_WalletLog.ExchRate = CDbl(oWPARes.ExchRate)
    aVB_WalletLog.ExchRateUnit = oWPARes.ExchRateUnit
    aVB_WalletLog.ReceiptText = oWPARes.ReceiptText
    aVB_WalletLog.Result = oCallAPI.ResponseMsg
    
    Set oWPAReq = Nothing
    Set oWPARes = Nothing
    Set oCallAPI = Nothing
    Call SP_TXTxWriteLogWalletPay("End Payment")
    On Error GoTo 0
End Function

Private Function W_GETxInQuiryRes(Optional ByRef ptTransID As String, Optional ByRef ptPayTime As String, Optional ByRef ptErrDesc As String, Optional ByRef ptErrCode As String, Optional ByRef ptWalletTransID As String, Optional ByRef ptWalletName As String, Optional ByRef ptWPSTransID As String) As String
    '*Tao 60-08-29 WPA WalletPay InQuiry
    
    Dim oCallAPI As Object
    Dim oWPAReq As Object
    Dim oWPARes As Object
    
    On Error Resume Next
    
    Set oCallAPI = CreateObject("ADACallAPI.cAdaCallAPI")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Query = " & Err.Description)
    End If
    If oCallAPI Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    Set oWPAReq = CreateObject("ADACallAPI.cmlWPAReq")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Query = " & Err.Description)
    End If
    If oWPAReq Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    oWPAReq.ReqType = "inquiry"
    oWPAReq.BizDate = tW_BizDate
    oWPAReq.Brand = tVB_WalletBrand
    oWPAReq.StoreID = tVB_WalletStoreID
    oWPAReq.DeviceID = tVB_WalletDevideID
    oWPAReq.OrigReqDT = tW_InVoiceDT
    oWPAReq.OrigOrderID = tW_InVoiceID
    oWPAReq.OrigAmt = Format(CDbl(tW_Amount), "########0.00")
    
    oCallAPI.DllKey = tVB_DllKey
    oCallAPI.DllPass = tVB_DllPass
    oCallAPI.HostURL = tVB_WalletHostURL
    oCallAPI.HostTimeout = nVB_WalletTimeout
    oCallAPI.SecurityProType = nW_SecurityProtocolType.Tls
    Call SP_TXTxWriteLogWalletPay("Start InQuiry " & tW_InVoiceID)
    DoEvents
    Set oWPARes = oCallAPI.C_WPAoRequest(oWPAReq)
    If Err Then Call SP_TXTxWriteLogWalletPay("Error InQuiry = " & Err.Description)
    If oWPARes Is Nothing Then
        W_GETxInQuiryRes = ""
    Else
        W_GETxInQuiryRes = oWPARes.StatusCode
        ptTransID = oWPARes.PartnerTransID
        ptPayTime = oWPARes.TransDT
        ptErrDesc = oWPARes.ErrMsg
        ptErrCode = oCallAPI.RespCode
        ptWalletTransID = oWPARes.WalletTransID
        ptWalletName = oWPARes.WalletName
        ptWPSTransID = oWPARes.WPSTransID
        Call SP_TXTxWriteLogWalletPay("Response InQuiry = " & W_GETxInQuiryRes & ";" & ptWalletName & ";" & ptTransID & ";" & ptWPSTransID & ";" & ptPayTime & ";" & ptErrDesc)
    End If
    
    aVB_WalletLog.ReqType = "Inquiry"
    aVB_WalletLog.Status = W_GETxInQuiryRes
    aVB_WalletLog.OrderID = oWPAReq.OrigOrderID
    aVB_WalletLog.ReqDT = oWPAReq.OrigReqDT
    aVB_WalletLog.ReqAmt = oWPAReq.OrigAmt
    'aVB_WalletLog.OrigTransID = ""
    'aVB_WalletLog.RemainingAmt = 0
    'aVB_WalletLog.Token = oWPAReq.Token
    aVB_WalletLog.Brand = oWPAReq.Brand
    aVB_WalletLog.StoreID = oWPAReq.StoreID
    aVB_WalletLog.DeviceID = oWPAReq.DeviceID
    aVB_WalletLog.CashierID = tVB_CNUserCode
    aVB_WalletLog.ResCode = oWPARes.RespCode
    aVB_WalletLog.ErrCode = oWPARes.ResponseCode
    aVB_WalletLog.ErrDesc = ptErrDesc
    aVB_WalletLog.TransDT = ptPayTime
    aVB_WalletLog.CancelDT = ""
    aVB_WalletLog.WalletName = oWPARes.WalletName
    aVB_WalletLog.WalletTransID = oWPARes.WalletTransID
    aVB_WalletLog.PartnerTransID = oWPARes.PartnerTransID
    aVB_WalletLog.WPSTransID = oWPARes.WPSTransID
    aVB_WalletLog.WalletCurr = oWPAReq.WalletCurr
    aVB_WalletLog.WalletCurrAmt = oWPARes.WalletCurrAmt
    aVB_WalletLog.ExchRate = CDbl(oWPARes.ExchRate)
    aVB_WalletLog.ExchRateUnit = oWPARes.ExchRateUnit
    aVB_WalletLog.WPSRefundID = ""
    aVB_WalletLog.RefundDT = ""
    aVB_WalletLog.RefundAmt = 0
    aVB_WalletLog.ReceiptText = oWPARes.ReceiptText
    aVB_WalletLog.Result = oCallAPI.ResponseMsg
    aVB_WalletLog.RefDoc = tVB_CNStoreCode & "-" & tVB_CNTerminalNum & tVB_Running & "-" & tW_RefDocDT
    aVB_WalletLog.StaDoc = ""
    
    Set oWPAReq = Nothing
    Set oWPARes = Nothing
    Set oCallAPI = Nothing
    Call SP_TXTxWriteLogWalletPay("End InQuiry")
    On Error GoTo 0
End Function

Private Function W_GETxVoidRes(Optional ByRef ptTransID As String, Optional ByRef ptPayTime As String, Optional ByRef ptErrDesc As String, Optional ByRef ptErrCode As String) As String
    '*Tao 60-08-29 WPA WalletPay Void
    
    Dim oCallAPI As Object
    Dim oWPAReq As Object
    Dim oWPARes As Object
    
    On Error Resume Next
    
    Set oCallAPI = CreateObject("ADACallAPI.cAdaCallAPI")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Void = " & Err.Description)
    End If
    If oCallAPI Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    Set oWPAReq = CreateObject("ADACallAPI.cmlWPAReq")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Cancel = " & Err.Description)
    End If
    If oWPAReq Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    
    tW_InVoiceDT = CStr(Format(Now, "yyyyMMddhhmmss"))
    If tW_RefDocDT = "" Then tW_RefDocDT = tW_InVoiceDT
    oWPAReq.ReqType = "void"
    oWPAReq.Brand = tW_RefBrand
    oWPAReq.StoreID = tW_RefStoreID
    oWPAReq.DeviceID = tW_RefDeviceID
    oWPAReq.WalletName = tW_RefWalletName
    oWPAReq.WalletTransID = tW_RefWalletTransID
    oWPAReq.OrigAmt = Format(CDbl(tW_RefAmount), "########0.00")
       
    oCallAPI.DllKey = tVB_DllKey
    oCallAPI.DllPass = tVB_DllPass
    oCallAPI.HostURL = tVB_WalletHostURL
    oCallAPI.HostTimeout = nVB_WalletTimeout
    oCallAPI.SecurityProType = nW_SecurityProtocolType.Tls
    Call SP_TXTxWriteLogWalletPay("Start Void " & tW_InVoiceID & ";" & Format(CDbl(tW_RefAmount), "########0.00"))
    
    '*Tao 60-08-30 Save Log WalletPay ก่อนส่ง Req
    aVB_WalletLog.ReqType = "Void"
    aVB_WalletLog.Status = ""
    aVB_WalletLog.OrderID = tW_InVoiceID
    aVB_WalletLog.ReqDT = tW_RefDT
    aVB_WalletLog.ReqAmt = oWPAReq.OrigAmt
    aVB_WalletLog.OrigTransID = oWPAReq.WalletTransID
    aVB_WalletLog.RemainingAmt = 0
    aVB_WalletLog.Token = ""
    aVB_WalletLog.Brand = oWPAReq.Brand
    aVB_WalletLog.StoreID = oWPAReq.StoreID
    aVB_WalletLog.DeviceID = oWPAReq.DeviceID
    aVB_WalletLog.CashierID = tVB_CNUserCode
    aVB_WalletLog.ResCode = ""
    aVB_WalletLog.ErrCode = ""
    aVB_WalletLog.ErrDesc = ""
    aVB_WalletLog.TransDT = ""
    aVB_WalletLog.CancelDT = ""
    aVB_WalletLog.WalletName = oWPAReq.WalletName
    aVB_WalletLog.WalletTransID = oWPAReq.WalletTransID
    aVB_WalletLog.PartnerTransID = ""
    aVB_WalletLog.WPSTransID = ""
    aVB_WalletLog.WalletCurr = ""
    aVB_WalletLog.WalletCurrAmt = 0
    aVB_WalletLog.ExchRate = 0
    aVB_WalletLog.ExchRateUnit = ""
    aVB_WalletLog.WPSRefundID = ""
    aVB_WalletLog.RefundDT = ""
    aVB_WalletLog.RefundAmt = 0
    aVB_WalletLog.ReceiptText = ""
    aVB_WalletLog.Result = ""
    aVB_WalletLog.RefDoc = tVB_CNStoreCode & "-" & tVB_CNTerminalNum & tVB_Running & "-" & tW_RefDocDT
    aVB_WalletLog.StaDoc = ""
    Call W_SaveLogWallet
    '====================================
    
    DoEvents
    Set oWPARes = oCallAPI.C_WPAoRequest(oWPAReq)
    If Err Then Call SP_TXTxWriteLogWalletPay("Error Cancel = " & Err.Description)
    If oWPARes Is Nothing Then
        W_GETxVoidRes = ""
    Else
        W_GETxVoidRes = oWPARes.StatusCode
        ptTransID = tW_RefPartnerTransID
        ptPayTime = oWPARes.CancelDT
        ptErrDesc = oWPARes.ErrMsg
        ptErrCode = oCallAPI.RespCode
        Call SP_TXTxWriteLogWalletPay("Response Void = " & W_GETxVoidRes & ";" & ptTransID & ";" & ptPayTime & ";" & ptErrDesc)
    End If
    
    aVB_WalletLog.Status = W_GETxVoidRes
    aVB_WalletLog.ResCode = oWPARes.RespCode
    aVB_WalletLog.ErrCode = oWPARes.ResponseCode
    aVB_WalletLog.ErrDesc = ptErrDesc
    aVB_WalletLog.TransDT = ptPayTime
    aVB_WalletLog.CancelDT = oWPARes.CancelDT
    If oWPARes.WalletTransID <> "" Then aVB_WalletLog.WalletTransID = oWPARes.WalletTransID
    aVB_WalletLog.ReceiptText = oCallAPI.ReceiptText
    aVB_WalletLog.Result = oCallAPI.ResponseMsg
    
    Set oWPAReq = Nothing
    Set oWPARes = Nothing
    Set oCallAPI = Nothing
    Call SP_TXTxWriteLogWalletPay("End Void")
    On Error GoTo 0
End Function

Private Function W_GETxRefundRes(Optional ByRef ptTransID As String, Optional ByRef ptPayTime As String, Optional ByRef ptErrDesc As String, Optional ByRef ptErrCode As String) As String
    '*Tao 60-08-29 WPA WalletPay Refund
    
    Dim oCallAPI As Object
    Dim oWPAReq As Object
    Dim oWPARes As Object
    
    On Error Resume Next
    
    Set oCallAPI = CreateObject("ADACallAPI.cAdaCallAPI")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Refund = " & Err.Description)
    End If
    If oCallAPI Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    Set oWPAReq = CreateObject("ADACallAPI.cmlWPAReq")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Refund = " & Err.Description)
    End If
    If oWPAReq Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    
    tW_InVoiceDT = CStr(Format(Now, "yyyyMMddhhmmss"))
    If tW_RefDocDT = "" Then tW_RefDocDT = tW_InVoiceDT
    oWPAReq.ReqType = "refund"
    oWPAReq.Brand = tW_RefBrand
    oWPAReq.StoreID = tW_RefStoreID
    oWPAReq.DeviceID = tW_RefDeviceID
    oWPAReq.CashierID = tVB_CNUserCode
    oWPAReq.WalletName = tW_RefWalletName
    oWPAReq.OrigPartnerTransID = tW_RefPartnerTransID
    oWPAReq.OrigAmt = Format(CDbl(tW_RefAmount), "########0.00")
    
    oCallAPI.DllKey = tVB_DllKey
    oCallAPI.DllPass = tVB_DllPass
    oCallAPI.HostURL = tVB_WalletHostURL
    oCallAPI.HostTimeout = nVB_WalletTimeout
    oCallAPI.SecurityProType = nW_SecurityProtocolType.Tls
    Call SP_TXTxWriteLogWalletPay("Start Refund " & tW_InVoiceID & ";" & Format(CDbl(tW_RefAmount), "########0.00"))
    
    '*Tao 60-08-30 Save Log WalletPay ก่อนส่ง Req
    aVB_WalletLog.ReqType = "Refund"
    aVB_WalletLog.Status = ""
    aVB_WalletLog.OrderID = tW_InVoiceID
    aVB_WalletLog.ReqDT = tW_RefDT
    aVB_WalletLog.ReqAmt = oWPAReq.OrigAmt
    aVB_WalletLog.OrigTransID = oWPAReq.OrigPartnerTransID
    aVB_WalletLog.RemainingAmt = 0
    aVB_WalletLog.Token = ""
    aVB_WalletLog.Brand = oWPAReq.Brand
    aVB_WalletLog.StoreID = oWPAReq.StoreID
    aVB_WalletLog.DeviceID = oWPAReq.DeviceID
    aVB_WalletLog.CashierID = oWPAReq.CashierID
    aVB_WalletLog.ResCode = ""
    aVB_WalletLog.ErrCode = ""
    aVB_WalletLog.ErrDesc = ""
    aVB_WalletLog.TransDT = ""
    aVB_WalletLog.CancelDT = ""
    aVB_WalletLog.WalletName = oWPAReq.WalletName
    aVB_WalletLog.WalletTransID = ""
    aVB_WalletLog.PartnerTransID = ""
    aVB_WalletLog.WPSTransID = ""
    aVB_WalletLog.WalletCurr = ""
    aVB_WalletLog.WalletCurrAmt = 0
    aVB_WalletLog.ExchRate = 0
    aVB_WalletLog.ExchRateUnit = ""
    aVB_WalletLog.WPSRefundID = ""
    aVB_WalletLog.RefundDT = ""
    aVB_WalletLog.RefundAmt = 0
    aVB_WalletLog.ReceiptText = ""
    aVB_WalletLog.Result = ""
    aVB_WalletLog.RefDoc = tVB_CNStoreCode & "-" & tVB_CNTerminalNum & tVB_Running & "-" & tW_RefDocDT
    aVB_WalletLog.StaDoc = ""
    Call W_SaveLogWallet
    '====================================
    
    DoEvents
    Set oWPARes = oCallAPI.C_WPAoRequest(oWPAReq)
    If Err Then Call SP_TXTxWriteLogWalletPay("Error Refund = " & Err.Description)
    If oWPARes Is Nothing Then
        W_GETxRefundRes = ""
    Else
        W_GETxRefundRes = oWPARes.StatusCode
        ptTransID = oWPARes.WPSRefundID
        ptPayTime = oWPARes.RefundDT
        ptErrDesc = oWPARes.ErrMsg
        ptErrCode = oCallAPI.RespCode
        Call SP_TXTxWriteLogWalletPay("Response Refund = " & W_GETxRefundRes & ";" & ptTransID & ";" & ptPayTime & ";" & ptErrDesc)
    End If
    
    aVB_WalletLog.Status = W_GETxRefundRes
    aVB_WalletLog.ResCode = oWPARes.RespCode
    aVB_WalletLog.ErrCode = oWPARes.ResponseCode
    aVB_WalletLog.ErrDesc = ptErrDesc
    aVB_WalletLog.TransDT = oWPARes.origtransdt
    aVB_WalletLog.WalletTransID = oWPARes.origwallettransid
    aVB_WalletLog.PartnerTransID = oWPARes.RespCode
    aVB_WalletLog.WPSTransID = oWPARes.origwpstransid
    aVB_WalletLog.WalletCurr = oWPARes.origwalletcurr
    aVB_WalletLog.WalletCurrAmt = oWPARes.origwalletcurramt
    aVB_WalletLog.ExchRate = oWPARes.origexchrate
    aVB_WalletLog.ExchRateUnit = oWPARes.origexchrateunit
    aVB_WalletLog.WPSRefundID = oWPARes.WPSRefundID
    aVB_WalletLog.RefundDT = oWPARes.RefundDT
    aVB_WalletLog.RefundAmt = oWPARes.RefundAmt
    aVB_WalletLog.ReceiptText = oCallAPI.ReceiptText
    aVB_WalletLog.Result = oCallAPI.ResponseMsg
    
    Set oWPAReq = Nothing
    Set oWPARes = Nothing
    Set oCallAPI = Nothing
    Call SP_TXTxWriteLogWalletPay("End Refund")
    On Error GoTo 0
End Function

Private Function W_GETxReverseRes(Optional ByRef ptErrDesc As String, Optional ByRef ptErrCode As String) As String
    '*Tao 60-08-29 WPA WalletPay Reverse
    
    Dim oCallAPI As Object
    Dim oWPAReq As Object
    Dim oWPARes As Object
       
    On Error Resume Next
    
    Set oCallAPI = CreateObject("ADACallAPI.cAdaCallAPI")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Reverse = " & Err.Description)
    End If
    If oCallAPI Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    Set oWPAReq = CreateObject("ADACallAPI.cmlWPAReq")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Reverse = " & Err.Description)
    End If
    If oWPAReq Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    
    If tW_RefDocDT = "" Then tW_RefDocDT = CStr(Format(Now, "yyyyMMddhhmmss"))
    If tW_BizDate = "" Then tW_BizDate = CStr(Format(Now, "yyyyMMdd"))
    oWPAReq.ReqType = "reversal"
    oWPAReq.BizDate = tW_BizDate
    oWPAReq.Brand = tVB_WalletBrand
    oWPAReq.StoreID = tVB_WalletStoreID
    oWPAReq.DeviceID = tVB_WalletDevideID
    oWPAReq.OrigReqDT = tW_InVoiceDT
    oWPAReq.OrigOrderID = tW_InVoiceID
    oWPAReq.OrigAmt = Format(CDbl(tW_Amount), "########0.00")
    
    oCallAPI.DllKey = tVB_DllKey
    oCallAPI.DllPass = tVB_DllPass
    oCallAPI.HostURL = tVB_WalletHostURL
    oCallAPI.HostTimeout = nVB_WalletTimeout
    oCallAPI.SecurityProType = nW_SecurityProtocolType.Tls
    Call SP_TXTxWriteLogWalletPay("Start Reverse " & tW_InVoiceID & ";" & Format(CDbl(tW_Amount), "########0.00"))
    DoEvents
    Set oWPARes = oCallAPI.C_WPAoRequest(oWPAReq)
    If Err Then Call SP_TXTxWriteLogWalletPay("Error Reverse = " & Err.Description)
    If oWPARes Is Nothing Then
        W_GETxReverseRes = ""
    Else
        If (tW_InVoiceID = oWPARes.ReqID) Then
            W_GETxReverseRes = oWPARes.StatusCode
        Else
            W_GETxReverseRes = "FAILED"
        End If
        ptErrDesc = oWPARes.ErrMsg
        ptErrCode = oCallAPI.RespCode
        'Call SP_TXTxWriteLogWalletPay("Response Reverse = " & W_GETxReverseRes & ";" & oWPARes.TransID & ";" & oWPARes.CancelDT & ";" & ptErrDesc)
        Call SP_TXTxWriteLogWalletPay("Response Reverse = " & W_GETxReverseRes & ";" & oWPARes.StatusCode & ";" & ptErrDesc)
    End If
    
    aVB_WalletLog.ReqType = "Reverse"
    aVB_WalletLog.Status = W_GETxReverseRes
    aVB_WalletLog.OrderID = oWPAReq.OrigOrderID
    aVB_WalletLog.ReqDT = oWPAReq.OrigReqDT
    aVB_WalletLog.ReqAmt = oWPAReq.OrigAmt
'    aVB_WalletLog.OrigTransID = ""
'    aVB_WalletLog.RemainingAmt = 0
'    aVB_WalletLog.Token = ""
    aVB_WalletLog.Brand = oWPAReq.Brand
    aVB_WalletLog.StoreID = oWPAReq.StoreID
    aVB_WalletLog.DeviceID = oWPAReq.DeviceID
    aVB_WalletLog.CashierID = tVB_CNUserCode
    aVB_WalletLog.ResCode = oWPARes.RespCode
    aVB_WalletLog.ErrCode = oWPARes.ResponseCode
    aVB_WalletLog.ErrDesc = ptErrDesc
    aVB_WalletLog.TransDT = ""
    aVB_WalletLog.CancelDT = ""
    aVB_WalletLog.WalletName = ""
    aVB_WalletLog.WalletTransID = ""
    aVB_WalletLog.PartnerTransID = ""
    aVB_WalletLog.WPSTransID = ""
    aVB_WalletLog.WalletCurr = ""
    aVB_WalletLog.WalletCurrAmt = ""
    aVB_WalletLog.ExchRate = 0
    aVB_WalletLog.ExchRateUnit = ""
    aVB_WalletLog.WPSRefundID = ""
    aVB_WalletLog.RefundDT = ""
    aVB_WalletLog.RefundAmt = 0
    aVB_WalletLog.ReceiptText = ""
    aVB_WalletLog.Result = oCallAPI.ResponseMsg
    aVB_WalletLog.RefDoc = tVB_CNStoreCode & "-" & tVB_CNTerminalNum & tVB_Running & "-" & tW_RefDocDT
    aVB_WalletLog.StaDoc = ""
    
    Set oWPAReq = Nothing
    Set oWPARes = Nothing
    Set oCallAPI = Nothing
    Call SP_TXTxWriteLogWalletPay("End Reverse")
    On Error GoTo 0
End Function

Private Function W_GETxFinalRes(Optional ByRef ptErrDesc As String, Optional ByRef ptErrCode As String) As String
    '*Tao 60-08-30 WPA WalletPay Finalize
    
    Dim oCallAPI As Object
    Dim oWPAReq As Object
    Dim oWPARes As Object
       
    On Error Resume Next
    
    Set oCallAPI = CreateObject("ADACallAPI.cAdaCallAPI")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Finalize = " & Err.Description)
    End If
    If oCallAPI Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    Set oWPAReq = CreateObject("ADACallAPI.cmlWPAReq")
    If Err Then
        aVB_WalletLog.ErrDesc = Err.Description
        aVB_WalletLog.ErrCode = Err.Number
        Call SP_TXTxWriteLogWalletPay("Error Finalize = " & Err.Description)
    End If
    If oWPAReq Is Nothing Then
        On Error GoTo 0
        Call SP_MSGnShowing("tms_CN999031", nCS_Warning, , , "tms_CN999032")
        Exit Function
    End If
    
    If tW_RefDocDT = "" Then tW_RefDocDT = CStr(Format(Now, "yyyyMMddhhmmss"))
    If tW_BizDate = "" Then tW_BizDate = CStr(Format(Now, "yyyyMMdd"))
    oWPAReq.ReqType = "finalize"
    oWPAReq.BizDate = tW_BizDate
    oWPAReq.Brand = tVB_WalletBrand
    oWPAReq.StoreID = tVB_WalletStoreID
    oWPAReq.DeviceID = tVB_WalletDevideID
    oWPAReq.OrigReqDT = tW_InVoiceDT
    oWPAReq.OrigOrderID = tW_InVoiceID
    oWPAReq.WalletName = tW_WalletName
    oWPAReq.WalletTransID = tW_WalletTransID
    
    oCallAPI.DllKey = tVB_DllKey
    oCallAPI.DllPass = tVB_DllPass
    oCallAPI.HostURL = tVB_WalletHostURL
    oCallAPI.HostTimeout = nVB_WalletTimeout
    oCallAPI.SecurityProType = nW_SecurityProtocolType.Tls
    Call SP_TXTxWriteLogWalletPay("Start Reverse " & tW_InVoiceID & ";" & Format(CDbl(tW_Amount), "########0.00"))
    DoEvents
    Set oWPARes = oCallAPI.C_WPAoRequest(oWPAReq)
    If Err Then Call SP_TXTxWriteLogWalletPay("Error Finalize = " & Err.Description)
    If oWPARes Is Nothing Then
        W_GETxFinalRes = ""
    Else
        W_GETxFinalRes = oWPARes.StatusCode
        ptErrDesc = oWPARes.ErrMsg
        ptErrCode = oCallAPI.RespCode
        Call SP_TXTxWriteLogWalletPay("Response Reverse = " & W_GETxFinalRes & ";" & oWPARes.StatusCode & ";" & ptErrDesc)
    End If
    
    aVB_WalletLog.ReqType = "Finalize"
    aVB_WalletLog.Status = W_GETxFinalRes
    aVB_WalletLog.OrderID = oWPAReq.OrigOrderID
    aVB_WalletLog.ReqDT = oWPAReq.OrigReqDT
    aVB_WalletLog.ReqAmt = tW_Amount
'    aVB_WalletLog.OrigTransID = ""
'    aVB_WalletLog.RemainingAmt = 0
'    aVB_WalletLog.Token = ""
    aVB_WalletLog.Brand = oWPAReq.Brand
    aVB_WalletLog.StoreID = oWPAReq.StoreID
    aVB_WalletLog.DeviceID = oWPAReq.DeviceID
    aVB_WalletLog.CashierID = tVB_CNUserCode
    aVB_WalletLog.ResCode = oWPARes.RespCode
    aVB_WalletLog.ErrCode = oWPARes.ResponseCode
    aVB_WalletLog.ErrDesc = ptErrDesc
    aVB_WalletLog.TransDT = ""
    aVB_WalletLog.CancelDT = ""
    aVB_WalletLog.WalletName = oWPAReq.WalletName
    aVB_WalletLog.WalletTransID = oWPAReq.WalletTransID
    aVB_WalletLog.PartnerTransID = ""
    aVB_WalletLog.WPSTransID = ""
    aVB_WalletLog.WalletCurr = ""
    aVB_WalletLog.WalletCurrAmt = ""
    aVB_WalletLog.ExchRate = 0
    aVB_WalletLog.ExchRateUnit = ""
    aVB_WalletLog.WPSRefundID = ""
    aVB_WalletLog.RefundDT = ""
    aVB_WalletLog.RefundAmt = 0
    aVB_WalletLog.ReceiptText = ""
    aVB_WalletLog.Result = oCallAPI.ResponseMsg
    aVB_WalletLog.RefDoc = tVB_CNStoreCode & "-" & tVB_CNTerminalNum & tVB_Running & "-" & tW_RefDocDT
    aVB_WalletLog.StaDoc = ""
    
    Set oWPAReq = Nothing
    Set oWPARes = Nothing
    Set oCallAPI = Nothing
    Call SP_TXTxWriteLogWalletPay("End Finalize")
    On Error GoTo 0
End Function

Private Function W_CHKxRegExp(ByVal ptPattern As String, ByVal ptString As String) As Boolean
'*Tao 60-09-05 Check QR Code
Dim oRegExp As Object
On Error Resume Next
    Set oRegExp = CreateObject("VBScript.RegExp")
    oRegExp.Pattern = ptPattern
    ' Set Case Insensitivity.
    oRegExp.IgnoreCase = True
    'Set global applicability.
    oRegExp.Global = True
    W_CHKxRegExp = oRegExp.Test(ptString)
    Set oRegExp = Nothing
End Function

Private Sub W_PRCxRestartWPA()
Dim tPath As String
Dim oFSO As New FileSystemObject
Dim tFile As String
Dim oWMI As Object
Dim oServices As Object
Dim oService As Object
Dim tPrcName As String
Dim nRet As Long

On Error Resume Next
    'Stop wpamon1.exe, wpamon2.exe
    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")
    For Each oService In oServices
        tPrcName = LCase(Trim(CStr(oService.Name) & ""))
        If (InStr(1, tPrcName, LCase("wpamon"), vbTextCompare) > 0) Then
            nRet = oService.Terminate
        End If
    Next
    Set oServices = Nothing
    Set oWMI = Nothing

    'Stop wpa.exe
    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")
    For Each oService In oServices
        tPrcName = LCase(Trim(CStr(oService.Name) & ""))
        If (InStr(1, tPrcName, LCase("wpa.exe"), vbTextCompare) > 0) Then
            tFile = oService.executablepath
            tPath = oFSO.GetParentFolderName(oFSO.GetParentFolderName(oFSO.GetParentFolderName(tFile))) & "\bin"
            If Not SP_FLEbManFile(tPath & "\wpa.cmd", nEN_Fle1Exist) Then
                tPath = ""
            End If
            nRet = oService.Terminate
            Exit For
        End If
    Next
    Set oServices = Nothing
    Set oWMI = Nothing
    
    If tPath = "" Then
        tPath = App.Path & "\WPA\bin"
        If Not SP_FLDbManFolder(tPath, nEN_Fle1Exist) Then
            tPath = oFSO.GetParentFolderName(App.Path) & "\WPA\bin"
            If Not SP_FLDbManFolder(tPath, nEN_Fle1Exist) Then
                tPath = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path)) & "\WPA\bin"
                If Not SP_FLDbManFolder(tPath, nEN_Fle1Exist) Then
                    tPath = App.Path & "\System\WPA\bin"
                    If Not SP_FLDbManFolder(tPath, nEN_Fle1Exist) Then
                        tPath = oFSO.GetParentFolderName(App.Path) & "\System\WPA\bin"
                        If Not SP_FLDbManFolder(tPath, nEN_Fle1Exist) Then
                            tPath = oFSO.GetParentFolderName(oFSO.GetParentFolderName(App.Path)) & "\System\WPA\bin"
                            If Not SP_FLDbManFolder(tPath, nEN_Fle1Exist) Then
                                tPath = "D:\WPA\bin"
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    tFile = tPath & "\wpa.cmd"
    If SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
        Call SleepEx(1000, 0)
        nRet = Shell(Chr(34) & tFile & Chr(34), vbHide)
    End If
    tFile = tPath & "\wpamon1.cmd"
    If SP_FLEbManFile(tFile, nEN_Fle1Exist) Then
        Call SleepEx(1000, 0)
        nRet = Shell(Chr(34) & tFile & Chr(34), vbHide)
    End If
End Sub

Private Function W_CHKxAlwWalletType(ByVal pbChkCode As Boolean, ByVal ptString As String, Optional ByRef ptWalletName As String) As Boolean
'*Tao 60-09-05 Check Wallet Type
Dim nType As Integer
Dim bTmp As Boolean
On Error Resume Next
    If pbChkCode Then
'        For nType = 0 To UBound(aVB_WalletType) - 1
'            Select Case nType
'            Case 0
'                If W_CHKxRegExp(tW_RegExpAlipay, ptString) Then
'                    If aVB_WalletType(nType).StaSale Then bTmp = True
'                    ptWalletName = aVB_WalletType(nType).TenderType
'                    Exit For
'                End If
'            Case 1
'                If W_CHKxRegExp(tW_RegExpWeChatpay, ptString) Then
'                    If aVB_WalletType(nType).StaSale Then bTmp = True
'                    ptWalletName = aVB_WalletType(nType).TenderType
'                    Exit For
'                End If
'            Case 2
'                If W_CHKxRegExp(tW_RegExpBluepay, ptString) Then
'                    If aVB_WalletType(nType).StaSale Then bTmp = True
'                    ptWalletName = aVB_WalletType(nType).TenderType
'                    Exit For
'                End If
'            Case 3
'                If aVB_WalletType(nType).StaSale Then bTmp = True
'                ptWalletName = aVB_WalletType(nType).TenderType
'            End Select
'        Next nType
        bTmp = True '*Tao 60-09-05 ให้ Sale ได้ทุก Wallet Type
    Else
'        Select Case UCase(ptString)
'        Case "ALIPAY", "ALI PAY"
'            bTmp = aVB_WalletType(0).StaSale
'            ptWalletName = aVB_WalletType(0).TenderType
'        Case "WECHAT", "WECHATPAY", "WECHAT PAY"
'            bTmp = aVB_WalletType(1).StaSale
'            ptWalletName = aVB_WalletType(1).TenderType
'        Case "BLUEPAY", "BLUE PAY"
'            bTmp = aVB_WalletType(2).StaSale
'            ptWalletName = aVB_WalletType(2).TenderType
'        Case Else
'            bTmp = aVB_WalletType(3).StaSale
'            ptWalletName = aVB_WalletType(3).TenderType
'        End Select
        bTmp = True '*Tao 60-09-05 ให้ Sale ได้ทุก Wallet Type
    End If
    W_CHKxAlwWalletType = bTmp
End Function
'
'Private Function W_GETxAutoReverseRes(Optional ByRef ptErrDesc As String, Optional ByRef ptErrCode As String) As String
'    '*Tao 60-05-15 KTC Wallet Auto Reverse
'
'    Dim oCallAPI As Object
'    Dim oWPAReq As Object
'    Dim oWPARes As Object
'
'    On Error Resume Next
'
'    Set oCallAPI = CreateObject("ADACallAPI.cAdaCallAPI")
'    If Err Then
'        aVB_WalletLog.ErrDesc = Err.Description
'        aVB_WalletLog.ErrCode = Err.Number
'        Call SP_TXTxWriteLogWalletPay("Error Auto Reverse = " & Err.Description)
'    End If
'    If oCallAPI Is Nothing Then
'        On Error GoTo 0
'        Exit Function
'    End If
'    Set oWPAReq = CreateObject("ADACallAPI.cmlWPAReq")
'    If Err Then
'        aVB_WalletLog.ErrDesc = Err.Description
'        aVB_WalletLog.ErrCode = Err.Number
'        Call SP_TXTxWriteLogWalletPay("Error Auto Reverse = " & Err.Description)
'    End If
'    If oWPAReq Is Nothing Then
'        On Error GoTo 0
'        Exit Function
'    End If
'
'    oWPAReq.ReqType = "cancel"
'    oWPAReq.Amt = Format(CDbl(tW_RefAmount), "########0.00")
'    oWPAReq.DeviceID = tW_RefDeviceID
'    oWPAReq.Brand = tW_RefBrand
'    oWPAReq.OrigReqDT = tW_RefDT
'    oWPAReq.OrigReqID = tW_InVoiceID
'    oWPAReq.StoreID = tW_RefStoreID
'    oWPAReq.SignType = tVB_WalletSignType
'    oWPAReq.SignKey = tVB_WalletSignKey
'    oCallAPI.HostURL = tVB_WalletHostURL
'    oCallAPI.HostTimeout = nVB_WalletTimeout
'    oCallAPI.SecurityProType = nW_SecurityProtocolType.Tls
'    Call SP_TXTxWriteLogWalletPay("Start Auto Reverse " & tW_InVoiceID & ";" & Format(CDbl(tW_RefAmount), "########0.00"))
'    DoEvents
'    Set oWPARes = oCallAPI.C_WPAoRequest(oWPAReq)
'    If Err Then Call SP_TXTxWriteLogWalletPay("Error Auto Reverse = " & Err.Description)
'    If oWPARes Is Nothing Then
'        W_GETxAutoReverseRes = ""
'    Else
'        If (tW_InVoiceID = oWPARes.ReqID) Then
'            W_GETxAutoReverseRes = oWPARes.StatusCode
'        Else
'            W_GETxAutoReverseRes = "FAILED"
'        End If
'        ptErrDesc = oWPARes.ErrMsg
'        ptErrCode = oCallAPI.ResponseCode
'        Call SP_TXTxWriteLogWalletPay("Response Auto Reverse = " & W_GETxAutoReverseRes & ";" & oWPARes.TransID & ";" & oWPARes.CancelDT & ";" & ptErrDesc)
'    End If
'
'    aVB_WalletLog.ReqType = "AutoReverse"
'    aVB_WalletLog.Status = W_GETxAutoReverseRes
'    aVB_WalletLog.ReqID = oWPAReq.OrigReqID
'    aVB_WalletLog.ReqDT = oWPAReq.OrigReqDT
'    aVB_WalletLog.ReqAmt = oWPAReq.Amt
'    aVB_WalletLog.RemainingAmt = 0
'    aVB_WalletLog.CustCode = ""
'    aVB_WalletLog.OrigTransID = tW_RefID
'    aVB_WalletLog.Brand = oWPAReq.Brand
'    aVB_WalletLog.StoreID = oWPAReq.StoreID
'    aVB_WalletLog.DeviceID = oWPAReq.DeviceID
'    aVB_WalletLog.Currency = tVB_WalletCurrency
'    aVB_WalletLog.ResCode = oWPARes.RespCode
'    aVB_WalletLog.ErrCode = ptErrCode
'    aVB_WalletLog.ErrDesc = ptErrDesc
'    aVB_WalletLog.TransID = oWPARes.TransID
'    aVB_WalletLog.TransDT = oWPARes.CancelDT
'    aVB_WalletLog.WalletTransID = oWPARes.WalletTransID
'    aVB_WalletLog.BuyerID = oWPARes.BuyerID
'    aVB_WalletLog.CNYAmt = CDbl(oWPARes.CNYAmt)
'    aVB_WalletLog.ConvRate = CDbl(oWPARes.ConvRate)
'    aVB_WalletLog.Result = oCallAPI.ResponseMsg
'    aVB_WalletLog.RefDoc = tVB_WalletPrefixCode & tVB_CNStoreCode & "-" & tVB_CNTerminalNum & tVB_Running & "-" & CStr(Format(Now, "yyyyMMddhhmmss"))
'    aVB_WalletLog.StaDoc = IIf(UCase(W_GETxAutoReverseRes) = "SUCCESS", "1", "2")
'
'    Set oWPAReq = Nothing
'    Set oWPARes = Nothing
'    Set oCallAPI = Nothing
'    Call SP_TXTxWriteLogWalletPay("End Auto Reverse")
'    On Error GoTo 0
'End Function

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
    If bW_FlagSendAmt Or (Not otbCode.Enabled) Then Exit Sub
    With oScanner
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            otbCode.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel)
        Else
            otbCode.Text = .ScanData
        End If
        If otbCode.Text <> "" Then
            If Asc(Right(otbCode.Text, 1)) = 13 Then
                otbCode.Text = Mid(otbCode.Text, 1, Len(otbCode.Text) - 1)
            End If
            Call otbCode_KeyPress(13)
        End If
        If bVB_OPOSScannerEnable Then
            If .DeviceEnabled = False Then
                .Open tVB_Scanner
                .Claim 0
                .DeviceEnabled = True
                .DataEventEnabled = True
                .AutoDisable = True
                .ClearInput
                If .DeviceName = tSNR_TOUCH_SCANNER Then
                    .DecodeData = True
                End If
            Else
                .ClearInput
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub
