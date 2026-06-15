VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNPayMVoucher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mVoucher;klg_Title"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   4575
      Index           =   2
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   8805
      _Version        =   720897
      _ExtentX        =   15531
      _ExtentY        =   8070
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox orbDay 
         BackColor       =   &H0080C0FF&
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
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox otbCobtact 
         BackColor       =   &H0080C0FF&
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
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2280
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox orbMonth 
         BackColor       =   &H0080C0FF&
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
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox orbYear 
         BackColor       =   &H0080C0FF&
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
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox otbNum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
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
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   4
         Text            =   "0"
         Top             =   1680
         Width           =   3135
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
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   0
         Top             =   480
         Width           =   5895
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
         Index           =   5
         Left            =   3360
         TabIndex        =   19
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label olaMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   15
         Top             =   2880
         Width           =   8295
         WordWrap        =   -1  'True
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
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Tag             =   "2;"
         Top             =   1155
         Width           =   2130
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
         Left            =   4320
         TabIndex        =   12
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(DD/MM/YYYY);klg_FmtDate"
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
         Left            =   5880
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   1155
         Width           =   3600
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher Amount:;klg_Amt"
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
         Top             =   1800
         Width           =   3315
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mVoucher :;klg_VoucherNo"
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
         Width           =   3465
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number:;klg_ContactNum"
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
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   2400
         Visible         =   0   'False
         Width           =   4320
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   4575
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   -240
      Width           =   8805
      _Version        =   720897
      _ExtentX        =   15531
      _ExtentY        =   8070
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.Label olaMsg2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher status updated Successfully"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   1320
         TabIndex        =   17
         Top             =   840
         Width           =   6015
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1170
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   4065
      Width           =   8805
      _Version        =   720897
      _ExtentX        =   15531
      _ExtentY        =   2064
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
      Begin VB.Timer otmAuth 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   960
         Top             =   480
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
         Left            =   5400
         TabIndex        =   1
         TabStop         =   0   'False
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
         Left            =   6960
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNPayMVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_Barcode As String
Private oW_cMVoucherApi  As New cMVoucherApi
Private tW_VoucherNo As String
Private tW_POSID As String
Private tW_updateTransID As String
Private tW_Update_datetime As String
Private bW_IsClose As Boolean
Private tW_NoOld As String


Private Sub W_FRMxSetInitial()
    otbNo.Text = ""
    otbNo.BackColor = nVB_CNClrBLock
    otbNo.Locked = True
    orbDay.Text = ""
    orbDay.BackColor = nVB_CNClrBLock
    orbDay.Locked = True
    orbMonth.Text = ""
    orbMonth.BackColor = nVB_CNClrBLock
    orbMonth.Locked = True
    orbYear.Text = ""
    orbYear.BackColor = nVB_CNClrBLock
    orbYear.Locked = True
    otbNum.Text = 0
    otbNum.Locked = True
    otbNum.BackColor = nVB_CNClrBLock
'    otbCobtact.Text = ""
'    otbCobtact.Locked = True
'    otbCobtact.BackColor = nVB_CNClrBLock
    olaMsg.Visible = False
    otcForm1(2).ZOrder 0
     ocmOk.Enabled = False
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    If bVB_Res640X480 Then
'        Me.Left = 0
        Me.Top = 2100
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
              '*Bump 56-02-28 ���� code 3 ��÷Ѵ
            If .DeviceName = tSNR_TOUCH_SCANNER Then
                .DecodeData = True
            End If
        End With
    End If
    tW_NoOld = ""
    cVB_Pay = 0
    tW_VoucherNo = ""
    bW_IsClose = False
    
    otmAuth.Enabled = True
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
    If bW_IsClose Then
        Unload Me
    Else
        tVB_VendorCpnNo = Trim(otbNo.Text)
        cVB_VendorCpnAmt = 0
        cVB_TotalDiscVendorCpnAmt = 0
        cVB_Pay = 0
        bVB_PayVendoeCancel = True
        Unload Me
    End If
End Sub

Private Sub ocmOK_Click()
    Dim cVoucherAmt As Double
    Dim cPay As Double
    Dim tSql$
On Error GoTo ErrHandle
    ocmOk.Enabled = False
     
 '*KT 67-01-31 FTH ������Ѻ����ա���ʴ���ͤ�������Ѻ���Ъ����
    If bW_IsClose Then
        Unload Me
    Else
         Screen.MousePointer = 11
        If Not (SP_CHKbBanknoteAlw(SP_VALcText2Double(otbNum.Text))) Then
            ocmOk.Enabled = True
             Screen.MousePointer = 0
            Exit Sub
        End If
        cPay = SP_VALcText2Double(otbNum.Text)
        If Trim(otbNo.Text) = "" Then 'check amount zero
             Screen.MousePointer = 0
            ocmOk.Enabled = True
            Call SP_OBJxSetFocus(otbNo) ' otbNo.SetFocus
            Exit Sub
        End If
        'for test case token expire res code 401
       ' tVB_CNMVoucherToken = "<REDACTED_TEST_JWT_removed_before_publishing>"   'for test token expire

        tW_POSID = tVB_CNTerminalNum
        tW_updateTransID = tVB_Running
        tW_Update_datetime = W_GETtTimeStamp 'Format(Now, "yyyy-MM-dd HH:mm:ss")
        If W_PRCbUpdateMVoucher Then
 '*KT 67-01-31 FTH ������Ѻ����ա���ʴ���ͤ�������Ѻ���Ъ����
'            otcForm1(0).ZOrder 0
'            olaMsg2.Caption = olaMsg.Caption
'            olaMsg2.ForeColor = vbGreen
            bW_IsClose = True
            cVB_Pay = cPay
            tVB_CrdCardNo = Trim(tW_VoucherNo)
            cVB_CrdAmt = cVB_Pay
            tVB_CrdExpD = Format(orbDay.Text, "00")
            tVB_CrdExpM = Format(orbMonth.Text, "00")
            tVB_CrdExpY = Format(orbYear.Text, "0000")
            tVB_CrdApprCode = "" ' Trim(otbCobtact.Text)
            tVB_CrdType = "BPVoucher"
            tVB_CreditGroup = "BPVoucher"
            ocmCancel.Value = False
            Unload Me  '*KT 67-01-31 FTH ������Ѻ����ա���ʴ���ͤ�������Ѻ���Ъ���� ���Դ˹�Ҩ�����
        Else
 '*KT 67-01-31 FTH ������Ѻ����ա���ʴ���ͤ�������Ѻ���Ъ���� ������ Error ����ʴ�����͹���
            otcForm1(0).ZOrder 0
            olaMsg2.Caption = olaMsg.Caption
            olaMsg2.ForeColor = vbRed
            bW_IsClose = True
            cVB_Pay = 0
            ocmCancel.Value = False
        End If
         Screen.MousePointer = 0
    End If
    Exit Sub
ErrHandle:

End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
    With oScanner
        'otbNo.Text = .ScanData '*Bump 56-02-28  Comment ��÷Ѵ���
        '*Bump 56-02-28 ���� code 8 ��÷Ѵ
        If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
            .DecodeData = True
        End If
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            otbNo.Text = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel) '*��ͧ���¨ҡ mCNSTP.SP_SplitBarcodeData  价�� mCNSP.SP_SplitBarcodeData
        Else
            otbNo.Text = .ScanData
        End If
'        DoEvents '*Ao 53-01-14   DoEvents
        If otbNo.Text <> "" Then
            '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 �������ö�� Scanner  Scan �Թ����� ������� Scan �Թ��������������Թ��һ�ҡ�����Ҷ١��ͧ�� �Т�鹢�ͤ�����辺�Թ��ҹ�� �óշ�� OPOS �ҧ����ա����ҹ��� Enter ��ͷ��� ��ͧ�Ѵ�͡
            '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check �������ö�� Scanner �� �Т�鹢�ͤ�����辺�����Թ���
            If Asc(Right(otbNo.Text, 1)) = 13 Then
                otbNo.Text = Mid(otbNo.Text, 1, Len(otbNo.Text) - 1)
            End If
            Call SP_OBJxSetFocus(otbNum)
        End If
        '�ӡ�� Open Scanner �ա���� ����鹤��駵��仨� Scan �Թ���������Ѻ
        '*KT 52-01-08 ����ͧ Open Scanner �ա���� ¡��鹡ó� Form Active
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
    MsgBox "Scaner error" & " Runtime : " & Err.Number & " " & Err.Description, vbCritical
End Sub



Private Sub otbNo_Change()
     '*KT TTK-ISS3000-874 issue scan the new barcode BPvoucher Amount not change.
     '----------------------------------------------
    If tW_VoucherNo <> "" Then
        If tW_VoucherNo <> Trim(otbNo.Text) Then
            tW_VoucherNo = ""
        End If
    End If
    '----------------------------------------------
     If Trim(otbNo.Text) = "" Then
        tW_NoOld = ""
     End If
End Sub

'Private Sub otbCobtact_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyDelete         'delete
'            otbNo.Text = ""
'        Case 13
'            If otbCobtact.Text <> "" Then
'               ocmOk_Click
'            End If
'    End Select
'End Sub

'Private Sub otbCobtact_KeyPress(KeyAscii As Integer)
'    Call UT_ChkKeyAscii(KeyAscii)
'End Sub

'Private Sub otbNo_GotFocus()
'    otbNo.SelStart = 0
'    otbNo.SelLength = Len(otbNo.Text)
'End Sub

Private Sub otbNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbNo.Text = ""
             '*KT TTK-ISS3000-874 issue scan the new barcode BPvoucher Amount not change.
            '--------------------------------------------------------------------------------------
            tW_VoucherNo = ""
            If tVB_CNMVoucherToken <> "" And nVB_CNMVoucherExpireIn > 0 Then
                otbNo.Locked = False
                otbNo.BackColor = nVB_CNClrBUnLock
            End If
            '---------------------------------------------------------------------------------------
        Case 13
            If otbNo.Text <> "" Then
                
                sendkeys "{TAB}"
            End If
    End Select
End Sub

Private Sub otbNo_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
End Sub

Private Sub otbNo_LostFocus()
    Dim tSql$
    Dim tTextPrompt$
On Error GoTo ErrHandle             '*KT 54-03-25 ��ͧ�ѹ�������ش
    'Debug.Print "3"
     Screen.MousePointer = 11

    If tW_VoucherNo = "" Then
        otbNo.Text = SP_TXTtTrimRightLeft(otbNo.Text)
         If Trim(otbNo.Text) = "" Then
             Screen.MousePointer = 0
             Exit Sub
         End If

         tTextPrompt = Trim(otbNo.Text)
        If (tW_NoOld = tTextPrompt And tW_NoOld <> "") Then
            Screen.MousePointer = 0

            
             Exit Sub
         End If
         If tVB_CNMVoucherToken <> "" Then
            tW_NoOld = tTextPrompt
             olaMsg.Caption = "Waiting verify ..."
             olaMsg.Visible = True
            DoEvents
             If W_PRCbVerifyMVoucher(tTextPrompt) Then
                 olaMsg.Visible = False
                 'Call SP_OBJxSetFocus(otbCobtact)
                 Call SP_OBJxSetFocus(ocmOk)
             Else
                 Call SP_OBJxSetFocus(otbNo) '*KT 67-08-22 TTK-ISS3000-876  �������ҹ ����ͧ Focus ��� ocmOk ���ͧ�ҡ�з�����ա�� enter ���
                  Call SP_OTBxSelText(otbNo) '
             End If
         End If
    End If
     Screen.MousePointer = 0
    Exit Sub
ErrHandle:
     '*KT 67-08-22 TTK-ISS3000-876
    Screen.MousePointer = 0
    Call SP_MSGxShowError(Err.Number, Err.Description)
    otbNum.Text = ""
    Call SP_OBJxSetFocus(otbNo)
End Sub

Private Sub otbNum_Change()         '*AO51-03-12    Comm.Sheet SF- 070
    otbNum.Text = SP_DEFtChangeNum(Trim(otbNum.Text))
    otbNum.SelStart = Len(otbNum.Text)
End Sub

Private Sub otbNum_GotFocus()
    otbNum.SelStart = 0
    otbNum.SelLength = Len(otbNum.Text)
End Sub

'Private Sub otbNum_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyDelete         'delete
'            otbNum.Text = ""
'    End Select
'End Sub
Private Sub otbNum_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 13
             If tW_VoucherNo <> "" Then
                Call ocmOK_Click
             End If
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             '�ȹ���
            KeyAscii = 0
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otmAuth_Timer()
    'Auth
    Dim bGetToken  As Boolean
On Error GoTo ErrHandle
     Screen.MousePointer = 11
    otmAuth.Enabled = False
    'for test case token expire res code 401
   ' tVB_CNMVoucherToken = "<REDACTED_TEST_JWT_removed_before_publishing>"
    If tVB_CNMVoucherToken = "" Then
         bGetToken = True
    ElseIf nVB_CNMVoucherExpireIn <= 0 Then
         bGetToken = True
    Else
         bGetToken = oW_cMVoucherApi.C_CHKbExpireToken(dVB_CNMVoucherExpireInTime, nVB_CNMVoucherExpireIn)
    End If
    
    If bGetToken Then
        olaMsg.Caption = "Waiting get access_token ... "
        olaMsg.Visible = True
        DoEvents
        If W_PRCbAuthMVoucher Then
            olaMsg.Visible = False
            otbNo.BackColor = nVB_CNClrBUnLock
            otbNo.Locked = False
            Call SP_OBJxSetFocus(otbNo)
        Else
            ocmCancel.SetFocus
        End If
    Else
        otbNo.BackColor = nVB_CNClrBUnLock
        otbNo.Locked = False
        Call SP_OBJxSetFocus(otbNo)
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrHandle:
     tVB_CNMVoucherToken = ""
     nVB_CNMVoucherExpireIn = 0
     dVB_CNMVoucherExpireInTime = Format(Now, "yyyy/MM/dd HH:mm:ss")
    Screen.MousePointer = 0
End Sub


Private Function W_PRCbAuthMVoucher() As Boolean
   'Auth
    Dim bGetToken  As Boolean
On Error GoTo ErrHandle
    tVB_CNMVoucherToken = ""
    nVB_CNMVoucherExpireIn = 0
        
    oW_cMVoucherApi.tAuthURL = tVB_CNMVoucherAuthURL
    oW_cMVoucherApi.tClient_id = tVB_CNMVoucherClient_id
    oW_cMVoucherApi.tClient_secret = tVB_CNMVoucherClient_secretL
    oW_cMVoucherApi.tGrant_type = tVB_CNMVoucherGrant_type
    oW_cMVoucherApi.tScope = tVB_CNMVoucherScope
    oW_cMVoucherApi.tContentTypeAuth = tVB_CNMVoucherContentTypeAuth
    oW_cMVoucherApi.nHostTimeOut = nVB_CNMVoucherTimeout
    If oW_cMVoucherApi.C_APIbAuth() Then
         tVB_CNMVoucherToken = oW_cMVoucherApi.tHKeyToken
         nVB_CNMVoucherExpireIn = oW_cMVoucherApi.nExpireIn
         dVB_CNMVoucherExpireInTime = oW_cMVoucherApi.dAuthDateTime
         W_PRCbAuthMVoucher = True
    Else
          tVB_CNMVoucherToken = ""
          nVB_CNMVoucherExpireIn = 0
         dVB_CNMVoucherExpireInTime = Format(Now, "yyyy/MM/dd HH:mm:ss")
         olaMsg.Caption = oW_cMVoucherApi.oC_RESAuthMVoucherInterf.GEToRespError.GETtRespErrMsg
         olaMsg.ForeColor = vbRed
         olaMsg.Visible = True
         ocmOk.Enabled = True
         W_PRCbAuthMVoucher = False
    End If
    Exit Function
ErrHandle:
     tVB_CNMVoucherToken = ""
     nVB_CNMVoucherExpireIn = 0
     dVB_CNMVoucherExpireInTime = Format(Now, "yyyy/MM/dd HH:mm:ss")
     W_PRCbAuthMVoucher = False
End Function

Private Function W_PRCbVerifyMVoucher(ByVal ptVoucherNo As String)
   'Verify
    Dim bGetToken  As Boolean
    Dim nCount As Integer
    Dim tMsgErr As String
On Error GoTo ErrHandle
    
VerifyHandle:
    tMsgErr = ""
    tW_VoucherNo = ""
    oW_cMVoucherApi.tVerifyURL = tVB_CNMVoucherVerifyURL
    oW_cMVoucherApi.nHostTimeOut = nVB_CNMVoucherTimeout
    oW_cMVoucherApi.tVoucherNo = ptVoucherNo
    ReDim aVB_MVoucherHKey(0)
    aVB_MVoucherHKey(0).tHkeyName = "Authorization" ' "Token"
    aVB_MVoucherHKey(0).tHkeyVal = "Bearer " & tVB_CNMVoucherToken
    'aVB_MVoucherHKey(1).tHkeyName = "Content-Type"
    'aVB_MVoucherHKey(1).tHkeyVal = "Value: application/json"
    nCount = nCount + 1
    Dim tErr As String
    Dim nErr As Long
    tErr = ""
    nErr = 0
    If oW_cMVoucherApi.C_APIbVerify(nErr, tErr) Then
        If oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETbRespSuccesss Then
             If oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETnRespCode = 200 Then
                If oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GEToRespData.GETnRespDataVoucherStatus = 0 Or oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GEToRespData.GETnRespDataVoucherStatus = 5 Then
                    orbDay.Text = Mid(oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GEToRespData.GETtRespDataVoucherExpDate, 9, 2)
                    orbMonth.Text = Mid(oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GEToRespData.GETtRespDataVoucherExpDate, 6, 2)
                    orbYear.Text = Mid(oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GEToRespData.GETtRespDataVoucherExpDate, 1, 4)
                    otbNum.Text = SP_FMTcCurPoint(oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GEToRespData.GETcRespDataVoucherAmt, False, True)
                    otbNo.Locked = True
                    otbNo.BackColor = nVB_CNClrBLock
'                    otbCobtact.Locked = False
'                    otbCobtact.BackColor = nVB_CNClrBUnLock
                    tW_VoucherNo = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GEToRespData.GETtRespDataVoucherNo
                    ocmOk.Enabled = True
                    W_PRCbVerifyMVoucher = True
                Else
                    If oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETnRespErrorCount > 0 Then
                       olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrorBySeqNo(0)
                       tMsgErr = olaMsg.Caption
                       If olaMsg.Caption = "" Then olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrMsg
                    Else
                        If tErr <> "" Then
                            olaMsg.Caption = "Runtime error no:" & nErr & " " & tErr
                        Else
                            olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrMsg
                        End If
                    End If
                    If olaMsg.Caption = "" Then olaMsg.Caption = SP_INItGetValue(tVB_CNPthFileINI, "All", "tms_CN005259", "The Voucher status is not allowed to do redemption") '
                    olaMsg.ForeColor = vbRed
                    olaMsg.Visible = True
                    ocmOk.Enabled = False
                End If
            Else
                If oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETnRespErrorCount > 0 Then
                   olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrorBySeqNo(0)
                   tMsgErr = olaMsg.Caption
                   If olaMsg.Caption = "" Then olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrMsg
                Else
                  
                    If tErr <> "" Then
                        olaMsg.Caption = "Runtime error no:" & nErr & " " & tErr
                    Else
                         olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrMsg
                    End If
                End If
                olaMsg.ForeColor = vbRed
                olaMsg.Visible = True
                ocmOk.Enabled = False
                If oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETnRespCode = 401 And UCase(tMsgErr) = UCase("Unauthorized") And nCount < 3 Then
                    SleepEx 2000, 0
                    olaMsg.Caption = "Waiting get access_token ... "
                    olaMsg.Visible = True
                    olaMsg.ForeColor = vbBlack
                    If W_PRCbAuthMVoucher Then
                        If oW_cMVoucherApi.tHKeyToken <> "" Then
                            tVB_CNMVoucherToken = oW_cMVoucherApi.tHKeyToken
                            GoTo VerifyHandle
                        End If
                    End If
                Else
                    olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrMsg
                    olaMsg.ForeColor = vbRed
                    olaMsg.Visible = True
                    ocmOk.Enabled = False
                End If
            End If
        Else
            If oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETnRespErrorCount > 0 Then
               olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrorBySeqNo(0)
               tMsgErr = olaMsg.Caption
               If olaMsg.Caption = "" Then olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrMsg
            Else
                If tErr <> "" Then
                    olaMsg.Caption = "Runtime error no:" & nErr & " " & tErr
                Else
                     olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrMsg
                End If
            End If
            olaMsg.ForeColor = vbRed
            olaMsg.Visible = True
            ocmOk.Enabled = False
         
            If oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETnRespCode = 401 And UCase(tMsgErr) = UCase("Unauthorized") And nCount < 3 Then
                SleepEx 2000, 0
                olaMsg.Caption = "Waiting get access_token ... "
                olaMsg.Visible = True
                olaMsg.ForeColor = vbBlack
                If W_PRCbAuthMVoucher Then
                    If oW_cMVoucherApi.tHKeyToken <> "" Then
                        tVB_CNMVoucherToken = oW_cMVoucherApi.tHKeyToken
                        GoTo VerifyHandle
                    End If
                End If
            End If
        End If
    Else
        If oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETnRespErrorCount > 0 Then
            olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrorBySeqNo(0)
            tMsgErr = olaMsg.Caption
            If olaMsg.Caption = "" Then olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrMsg
         Else
            
            If tErr <> "" Then
                olaMsg.Caption = "Runtime error no:" & nErr & " " & tErr
            Else
                olaMsg.Caption = oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETtRespErrMsg
            End If
         End If
         olaMsg.ForeColor = vbRed
         olaMsg.Visible = True
         ocmOk.Enabled = False
         
         If oW_cMVoucherApi.oC_RESVerifyMVoucherInterf.GETnRespCode = 401 And UCase(tMsgErr) = UCase("Unauthorized") And nCount < 3 Then
            SleepEx 2000, 0
            olaMsg.Caption = "Waiting get access_token ... "
            olaMsg.Visible = True
            olaMsg.ForeColor = vbBlack

            If W_PRCbAuthMVoucher Then
                If oW_cMVoucherApi.tHKeyToken <> "" Then
                    tVB_CNMVoucherToken = oW_cMVoucherApi.tHKeyToken
                    GoTo VerifyHandle
                End If
            End If
        End If
    End If
    Exit Function
ErrHandle:
    olaMsg.Caption = Err.Description
    olaMsg.ForeColor = vbRed
    olaMsg.Visible = True
    ocmOk.Enabled = False
End Function

Private Function W_PRCbUpdateMVoucher() As Boolean
   'Update
    Dim bGetToken  As Boolean
    Dim nCount As Integer
    Dim tMsgErr As String
On Error GoTo ErrHandle
VerifyHandle:
    tMsgErr = ""
    oW_cMVoucherApi.tUpdateURL = tVB_CNMVoucherUpdateURL
    oW_cMVoucherApi.nHostTimeOut = nVB_CNMVoucherTimeout
    ReDim aVB_MVoucherHKey(0)
    aVB_MVoucherHKey(0).tHkeyName = "Authorization" ' "Token"
    aVB_MVoucherHKey(0).tHkeyVal = "Bearer " & tVB_CNMVoucherToken
'    aVB_MVoucherHKey(1).tHkeyName = "Content-Type"
'    aVB_MVoucherHKey(1).tHkeyVal = "Value: application/json"
    oW_cMVoucherApi.tVoucherNo = tW_VoucherNo
    oW_cMVoucherApi.tPOSID = tW_POSID
    oW_cMVoucherApi.tUpdateTransID = tW_updateTransID
    oW_cMVoucherApi.tUpdate_datetime = tW_Update_datetime
    nCount = nCount + 1
    If oW_cMVoucherApi.C_APIbUpdate() Then
        If oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETbRespSuccesss Then
            If oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETnRespCode = 200 Then
                olaMsg.Caption = oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETtRespData
                'olaMsg.ForeColor = vbGreen
                olaMsg.Visible = True
                ocmOk.Enabled = True
                ocmCancel.Visible = False
                W_PRCbUpdateMVoucher = True
            Else
                If oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETnRespErrorCount > 0 Then
                   olaMsg.Caption = oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETtRespErrorBySeqNo(0)
                   tMsgErr = olaMsg.Caption
                   If olaMsg.Caption = "" Then olaMsg.Caption = oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETtRespErrMsg
                Else
                   olaMsg.Caption = oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETtRespErrMsg
                End If
                'olaMsg.ForeColor = vbRed
                olaMsg.Visible = True
                ocmOk.Enabled = False
                If oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETnRespCode = 401 And UCase(tMsgErr) = UCase("Unauthorized") And nCount < 3 Then
                    olaMsg.ForeColor = vbRed
                    SleepEx 2000, 0
                    olaMsg.Caption = "Waiting get access_token ... "
                    olaMsg.Visible = True
                    olaMsg.ForeColor = vbBlack
                    If W_PRCbAuthMVoucher Then
                        If oW_cMVoucherApi.tHKeyToken <> "" Then
                            tVB_CNMVoucherToken = oW_cMVoucherApi.tHKeyToken
                            GoTo VerifyHandle
                        End If
                    End If
                End If
                W_PRCbUpdateMVoucher = False
            End If
        Else
             If oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETnRespErrorCount > 0 Then
               olaMsg.Caption = oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETtRespErrorBySeqNo(0)
               tMsgErr = olaMsg.Caption
               If olaMsg.Caption = "" Then olaMsg.Caption = oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETtRespErrMsg
            Else
               olaMsg.Caption = oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETtRespErrMsg
            End If
            'olaMsg.ForeColor = vbRed
            olaMsg.Visible = True
            ocmOk.Enabled = False
            If oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETnRespCode = 401 And UCase(tMsgErr) = UCase("Unauthorized") And nCount < 3 Then
                 olaMsg.ForeColor = vbRed
                SleepEx 2000, 0
                olaMsg.Caption = "Waiting get access_token ... "
                olaMsg.Visible = True
                olaMsg.ForeColor = vbBlack
                If W_PRCbAuthMVoucher Then
                    If oW_cMVoucherApi.tHKeyToken <> "" Then
                        tVB_CNMVoucherToken = oW_cMVoucherApi.tHKeyToken
                        GoTo VerifyHandle
                    End If
                End If
            End If
            W_PRCbUpdateMVoucher = False
        End If
    Else
         If oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETnRespErrorCount > 0 Then
            olaMsg.Caption = oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETtRespErrorBySeqNo(0)
            tMsgErr = olaMsg.Caption
            If olaMsg.Caption = "" Then olaMsg.Caption = oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETtRespErrMsg
         Else
            olaMsg.Caption = oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETtRespErrMsg
         End If
         
         'olaMsg.ForeColor = vbRed
         olaMsg.Visible = True
         ocmOk.Enabled = False
         If oW_cMVoucherApi.oC_RESUpdateMVoucherInterf.GETnRespCode = 401 And UCase(tMsgErr) = UCase("Unauthorized") And nCount < 3 Then
             olaMsg.ForeColor = vbRed
            SleepEx 2000, 0
            olaMsg.Caption = "Waiting get access_token ... "
            olaMsg.Visible = True
            olaMsg.ForeColor = vbBlack
            If W_PRCbAuthMVoucher Then
                If oW_cMVoucherApi.tHKeyToken <> "" Then
                    tVB_CNMVoucherToken = oW_cMVoucherApi.tHKeyToken
                    GoTo VerifyHandle
                End If
            End If
        End If

         W_PRCbUpdateMVoucher = False
    End If
    Exit Function
ErrHandle:
    olaMsg.Caption = Err.Description
    'olaMsg.ForeColor = vbRed
    olaMsg.Visible = True
    ocmOk.Enabled = False
    W_PRCbUpdateMVoucher = False
End Function

Private Function W_GETtTimeStamp() As String
    Dim dCurrentTime As Date
    Dim tFormattedTimestamp As String
    Dim tMilliseconds As String
 On Error GoTo ErrHandle
    ' Current Datetime
    dCurrentTime = Now
    
    '  Timer
    tMilliseconds = Right("000" & CStr(Int(Timer * 1000) Mod 1000), 3)
    
    ' Time stamp
    tFormattedTimestamp = Format(dCurrentTime, "yyyy-MM-ddTHH:mm:ss") & "." & tMilliseconds & "Z"
    
    ' Format  "yyyy-MM-ddThh:mm:ss.SSSZ"
    W_GETtTimeStamp = tFormattedTimestamp
    Exit Function
ErrHandle:
    
End Function

