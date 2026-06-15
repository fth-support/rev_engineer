VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNPayNETS 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2640
   ClientLeft      =   6255
   ClientTop       =   6360
   ClientWidth     =   5235
   ControlBox      =   0   'False
   Icon            =   "wCNPayNETS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5235
   Begin VB.Frame ofaSwipeCard 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   4200
      TabIndex        =   11
      Top             =   3000
      Width           =   3255
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Swipe or Insert Card into NETS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   11
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Frame ofa3Sec 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   480
      TabIndex        =   8
      Top             =   3000
      Width           =   3255
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ESC for manual or wait for swipe card."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   10
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label olaForm 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "xx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   1320
         TabIndex        =   9
         Top             =   1560
         Width           =   735
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1935
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   5225
      _Version        =   720897
      _ExtentX        =   9216
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
         TabIndex        =   0
         Text            =   "0"
         Top             =   480
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
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1080
         Width           =   2415
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
         Top             =   480
         Width           =   2280
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approve Code :;klg_ApprCode"
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
         Top             =   1080
         Width           =   3795
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1425
      Width           =   5225
      _Version        =   720897
      _ExtentX        =   9216
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   600
         Top             =   600
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Timer otmForm 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   600
      End
      Begin VB.Timer otmEDC 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   1080
         Top             =   600
      End
      Begin VB.CommandButton ocmCancel 
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
         Left            =   3420
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
         Left            =   1760
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNPayNETS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  '*Eaw 56-03-30 NETS Port
Private bW_HandleESC As Boolean
Private bW_ManaulMode As Boolean
Private bW_EdcCheck As Boolean
Private nW_TimeESC As Long
Private nW_Timer As Long
Private tW_ApprCode As String '*Eaw 56-08-30

Private Function W_GETtNetType(pnNetType As Integer) As String
'------------------------------------------------------------------------
'Call:
'Cmt : '*Eaw 56-08-27 ถ้าการใช้ด้วย NET ทุกๆ Type ใน Manual Mode ให้เอา FTSrcAbb มาจาก Master
'-------------------------------------------------------------------------
    Dim tNetDesc As String
    Dim aNetType() As String
    Dim tTmp As String
On Error GoTo ErrHandle
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnReturn, tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
            tNetDesc = aVB_TenderMnu(nVB_CurTenderSeqNo).DescReturn
        Case Else
            tNetDesc = aVB_TenderMnu(nVB_CurTenderSeqNo).DescSale
    End Select
        aNetType = Split(tNetDesc, ";")
        tTmp = aNetType(pnNetType)
        W_GETtNetType = tTmp
    
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Function W_CHKbNETPay(ByRef pcPay As Double) As Boolean
'------------------------------------------------------------------------
'Call:
'Cmt : เรียกใช้เครื่อง  NETS ผ่าน dll  เมื่อใช้ Gift Card  '*Eaw 56-06-15 '*Tao 62-07-19 ไม่ใช่อีกต่อไปแระ เป็น NETS QR ไปละ
'-------------------------------------------------------------------------
Dim tMsg As String
Dim oNETS As Object
Dim nPort As Integer
Const tMsgRet = "" '"Do you want to retry again?"
Const tMsgBf = "Toggle from NETS Purchase not success" '& vbCrLf & "Do you want to retry again?"
Const tMsgAf = "Toggle back to NETS Modute not success" & vbCrLf & "Please Contact Administrator"
On Error GoTo ErrHandle
    W_CHKbNETPay = False
    Set oNETS = CreateObject("AdaNETSDLL.cVeriFone")
    With oNETS
        .SetNetAmout = otbNum.Text
        nPort = nVB_NETPort
        olaForm(11).Caption = "Toggle from NETS Purchase is processing."
        If .EDC1CreateMessage(eEN_AfterGC) Then  'เปิดการเชื่อมต่อ NETS
            If .EDC2SendMessage(nPort) Then
                tMsg = .GetReturnString
                '.EDC2EOT (nPort) '*BG 56-10-25 มให้ส่งในDLL แล้ว   **Bump Review
            Else
                tMsg = .GetMessage
                tMsg = tMsg & vbCrLf & tMsgBf
            End If
        Else
            tMsg = .GetMessage
            tMsg = tMsg & vbCrLf & tMsgBf
        End If
        Select Case UCase(tMsg)  'เรียกฟังก์ชั่น NETS
            Case "APPROVED"
            olaForm(11).Caption = "Swipe or Insert Card into NETS"
            If .EDC1CreateMessage(eEN_GiftCard) Then
                If .EDC2SendMessage(nPort) Then
                    tMsg = .GetReturnString
'                    .EDC2EOT (nPort) '*BG 56-10-25 มให้ส่งในDLL แล้ว   **Bump Review
                Else
                    tMsg = .GetMessage
                End If
            Else
                tMsg = .GetMessage
            End If
        Case Else
            bW_EdcCheck = False
            If tMsg = "" Then tMsg = tMS_CN519 '*BG 56-10-28
            If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, tCS_CNMsgInfo) = vbOK Then
'                nW_Timer = 0
'                nW_TimeESC = 3
'                olaForm(9).Caption = nW_TimeESC
'                ofa3Sec.Left = 750     'center
'                ofa3Sec.Top = 120
'                ofaSwipeCard.Left = 2900    'ซ่อน
'                ofaSwipeCard.Top = 5400
'                bW_HandleESC = True
'                otmEDC.Enabled = False   'stop otmEDC
'                otmForm.Enabled = True    'start otmForm
'                W_CHKbNETPay = False
'                Exit Function
'            Else
                pcPay = 0
                W_CHKbNETPay = True
                Exit Function
            End If
        End Select
        Select Case UCase(tMsg)
            Case "APPROVED"
                olaForm(11).Caption = "Toggle back to NETS Purchase is processing."
                tW_ApprCode = .GetApprovalCode '*Eaw 56-08-30 Get approve code หลังรูดบัตรผ่านแล้ว
                 W_CHKbNETPay = True            '*BG 56-09-16  เพิ่มมา หมายความว่า จ่ายด้วย NETS->Giftcard สำเร็จแล้ว
'                If .EDC1CreateMessage(eEN_AfterGC) Then  'ปิดการเชื่อมต่อ NETS
'                    If .EDC2SendMessage(nPort) Then
'                        tMsg = .GetReturnString
'                        W_CHKbNETPay = True
'                    Else
'                        tMsg = .GetMessage
'                        W_CHKbNETPay = True
'                    End If
'                Else
'                    tMsg = .GetMessage
'                    W_CHKbNETPay = True
'                End If
            Case Else
'                SleepEx 5000, 0     '5วินาที
                bW_EdcCheck = False
                If tMsg = "" Then tMsg = tMS_CN519 '*BG 56-10-28
                If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, tCS_CNMsgInfo) = vbOK Then
'                    nW_Timer = 0
'                    nW_TimeESC = 3
'                    olaForm(9).Caption = nW_TimeESC
'                    ofa3Sec.Left = 750     'center
'                    ofa3Sec.Top = 120
'                    ofaSwipeCard.Left = 2900    'ซ่อน
'                    ofaSwipeCard.Top = 5400
'                    bW_HandleESC = True
'                    otmEDC.Enabled = False   'stop otmEDC
'                    otmForm.Enabled = True    'start otmForm
'                    W_CHKbNETPay = False
'                    Exit Function
'                Else
                    pcPay = 0
                    W_CHKbNETPay = True
                    Exit Function
                End If
        End Select
    End With
    Exit Function
ErrHandle:
    W_CHKbNETPay = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Public Function W_CHKbFlashPay(ByVal ptTenderCode As String, ByRef pcPay As Double) As Boolean
'------------------------------------------------------------------------
'Call:
'Cmt : เรียกใช้เครื่อง  NETS ผ่าน dll  เมื่อใช้ NETS/Flash pay/Cash card  '*Eaw 56-04-29 NETS Port
'-------------------------------------------------------------------------
Dim tMsg As String
Dim oNETS As Object
Dim nPort As Integer
Dim nTenderType As Long
On Error GoTo ErrHandle
    W_CHKbFlashPay = False
    Select Case ptTenderCode
        Case "T013"
            nTenderType = eEN_CashCard
        Case "T015"
            Select Case nVB_CNNETSType
                Case 0
                    nTenderType = eEN_NETSTender
                Case 1
                    nTenderType = eEN_BCA
                Case 2
                    nTenderType = eEN_GiftCard
                Case 3
                    'nTenderType = eEN_PrePaid '*Tao 62-07-19 ของเดิม
                    nTenderType = eEN_NetsQR '*Tao 62-07-19 NETS QR
            End Select
        Case "T016"
            nTenderType = eEN_FlashPay
        Case Else
            nTenderType = -99
    End Select
    Set oNETS = CreateObject("AdaNETSDLL.cVeriFone")
    With oNETS
           .SetNetAmout = otbNum.Text
           olaForm(11).Caption = "Swipe or Insert Card into NETS"
           If .EDC1CreateMessage(nTenderType) Then
                nPort = nVB_NETPort
                If .EDC2SendMessage(nPort) Then
                    tMsg = .GetReturnString
'                    .EDC2EOT (nPort) '*BG 56-10-25 มให้ส่งในDLL แล้ว   **Bump Review
                    Select Case tMsg
                        Case UCase("APPROVED")
                              bW_EdcCheck = True
                              W_CHKbFlashPay = True
                              tW_ApprCode = .GetApprovalCode '*Eaw 56-08-30 Get approve code
                              Exit Function
                    End Select
                Else
                        tMsg = .GetMessage
                End If
            Else
                tMsg = .GetMessage
            End If
'            SleepEx 5000, 0     '5วินาที '*Eaw 56-07-09
'            tMsg = tMsg & vbCrLf & "DO YOU WANT TO TRY AGAIN?"
            bW_EdcCheck = False
             If tMsg = "" Then tMsg = tMS_CN519
             If SP_MSGnShowing("tms_CN005065" & ";" & tMsg, tCS_CNMsgInfo) = vbOK Then   '*Eaw 56-06-17
'                nW_Timer = 0
'                nW_TimeESC = 3
'                olaForm(9).Caption = nW_TimeESC
'                ofa3Sec.Left = 750     'center
'                ofa3Sec.Top = 120
'                ofaSwipeCard.Left = 2900    'ซ่อน
'                ofaSwipeCard.Top = 5400
'                bW_HandleESC = True
'                otmEDC.Enabled = False   'stop otmEDC
'                otmForm.Enabled = True    'start otmForm
'                W_CHKbFlashPay = False
'                Exit Function
'             Else
                pcPay = 0
                W_CHKbFlashPay = True
                Exit Function
             End If
             DoEvents
    End With
ErrHandle:
    W_CHKbFlashPay = False
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub W_FRMxSetInitial()
'    otbNum.Text = 0                           '*Eaw 56-02-02  Auto fill amount balance
    otbNum.MaxLength = 13                         '*Eaw 56-03-02  Auto fill amount balance
'    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)    '*Eaw 56-04-29 NETS Port
    otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_PayCredit), False, True)       '*Eaw 56-04-29 NETS Port
    otbNo.Text = ""
End Sub

Private Sub Form_Activate()
On Error GoTo ErrHandle
    Call SP_OTBxSelText(otbNum)
    Exit Sub
ErrHandle:

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If otmEDC.Enabled Then
                'none
            ElseIf bW_HandleESC Then
                bW_HandleESC = False    'อาจจะไม่จำเป็น เพราะจะ stop timer แทน
                otmForm.Enabled = False
                otmEDC.Enabled = False
                nW_TimeESC = 0
                ofa3Sec.Left = 120      'ซ่อน
                ofa3Sec.Top = 5400
                bW_ManaulMode = True
                ocmCancel.Cancel = True
                otbNum.Locked = False
                otbNum.BackColor = nVB_CNClrBUnLock
                otcForm1(1).Enabled = True
                otcForm1(2).Enabled = True
                Call SP_OBJxSetFocus(otbNum)
            Else
                Call ocmCancel_Click
            End If
        End Select
End Sub

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
'*BG 56-11-12  ตอนนี้จ่ายเงินแล้วจะไม่ให้ Scanner ทำงานเพราะจะกระทบกับข้อมูลของเครื่อง NETs ที่ได้มา
'    If bVB_OPOSScannerEnable Then
'        With oScanner
'            .Open tVB_Scanner
'            .Claim 0 'TeamPos2000
'            'Enable the device.
'            .DeviceEnabled = True
'            'Enable the event.
'            .DataEventEnabled = True
'            .AutoDisable = True
'            .ClearInput
'        End With
'    End If
'****
    cVB_Pay = 0             '*Ao 2014/12/04      V 4.2.159   CommSheet TK-ISS3000-568
    Me.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu
    '*Eaw 56-04-29 NETS Port
    bW_EdcCheck = False
    bW_HandleESC = True
    bW_ManaulMode = False
    nW_Timer = 0
    If SP_NETbCheckPort Then
        '*Eaw 56-06-11
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnReturnManualTax, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnSaleTaxInc, tEN_TrnReturn, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel, tEN_TrnCancelVoucher, tEN_TrnOtherDisbursement  '*Eaw 56-11-26 เพิ่มเคสจากโหมดปกติเข้าไป CommSheet TK-ISS3000-363-FAPL-129.xls
                bW_HandleESC = False
                bW_ManaulMode = True
                otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
            Case Else
                bW_HandleESC = True
                bW_ManaulMode = False
'                nW_TimeESC = 3
                nW_TimeESC = tVB_EDCWait   '*BG 56-11-12 ใช้ Waiting Time จากค่าในตาราง TSysConfig แทน
                olaForm(9).Caption = nW_TimeESC
                ofa3Sec.Left = 750     'center
                ofa3Sec.Top = 120
                ofaSwipeCard.Left = 2900    'ซ่อน
                ofaSwipeCard.Top = 5400
                otcForm1(1).Enabled = False     'lock text box
                otcForm1(2).Enabled = False     'lock ปุ่ม
                otmForm.Enabled = True           'start otmForm
        End Select
    Else
        bW_HandleESC = False
        bW_ManaulMode = True
        otbNum.Text = SP_FMTcCurPoint(SP_VALcText2Double(cVB_Balance), False, True)
    End If
'------------------------------------------------------------------------
    Exit Sub
ErrHandle:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
'*BG 56-11-12   '*BG 56-11-12  ตอนนี้จ่ายเงินแล้วจะไม่ให้ Scanner ทำงานเพราะจะกระทบกับข้อมูลของเครื่อง NETs ที่ได้มา
'    With oScanner
'        If bVB_OPOSScannerEnable Then
'            If .DeviceEnabled = True Then
'                .DeviceEnabled = False
'                .Release
'                .Close
'            End If
'        End If
'    End With
'****
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    cVB_Pay = 0
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim cPay As Double
    If SP_VALcText2Double(otbNum.Text) > SP_VALcText2Double(cVB_Balance) Then
        Call SP_MSGnShowing("tms_CN005066", nCS_Error)  '*Eaw 56-03-13  CommSheet TK-ISS3000-065 00.00.01.xlsx
        Call SP_OBJxSetFocus(otbNum)  '*Eaw 56-03-13
        Exit Sub            '*Ao 2013/02/28   CommSheet TK-ISS3000-065 00.00.01.xlsx
    End If
    cPay = SP_VALcText2Double(otbNum.Text)
    If Trim(otbNum.Text) = "" Then 'check amount zero
        Call SP_OBJxSetFocus(otbNum)   'otbNum.SetFocus
        Exit Sub
    End If
    If cPay = 0 Then 'check amount zero
        Call SP_MSGnShowing("tms_CN005064", nCS_Error)
        Exit Sub
    End If
'   Select Case tVB_TenderCode
'        Case "T013"     'Cash Card
''            tVB_NETSApprCode = Trim(otbNo.Text)        '*Eaw 55-11-28
'            tVB_CrdApprCode = Trim(otbNo.Text)              '*Eaw 55-11-28
'        Case "T015"     'NETS
''            tVB_CashCrdApprCode = Trim(otbNo.Text)
'            tVB_CrdApprCode = Trim(otbNo.Text)              '*Ao 2013/01/04
'        Case "T016"     'Flash Pay
''            tVB_FlashApprCode = Trim(otbNo.Text)         '*Eaw 55-11-28
'            tVB_CrdApprCode = Trim(otbNo.Text)              '*Eaw 55-11-28
'    End Select
'     tVB_CrdApprCode = Trim(otbNo.Text)
    'ย้ายขึ้นมาทำงานก่อน
    Select Case tVB_TenderCode      '*Ao 2014/12/04      V 4.2.159   CommSheet TK-ISS3000-568
        Case "T015"
            If bW_ManaulMode Then
                tVB_NetDesc = W_GETtNetType(nVB_CNNETSType)
            Else
                tVB_NetDesc = Mid(tVB_CrdType, 3)
            End If
        Case "T013", "T016" '*Eaw 56-09-17
            If bW_ManaulMode Then
                tVB_NetDesc = tVB_TenderName
            Else
                tVB_NetDesc = tVB_TenderType
            End If
    End Select
    
    If Not bW_ManaulMode Then
        Select Case tVB_TenderCode
            Case "T013", "T016"
                 If Not W_CHKbFlashPay(tVB_TenderCode, cPay) Then
                    cVB_Pay = 0                 '*Ao 2014/12/04      V 4.2.159   CommSheet TK-ISS3000-568
                    Exit Sub
                 End If
            Case "T015"
                 Select Case nVB_CNNETSType
                    Case 0
                        If Not W_CHKbFlashPay(tVB_TenderCode, cPay) Then
                            cVB_Pay = 0                 '*Ao 2014/12/04      V 4.2.159   CommSheet TK-ISS3000-568
                            Exit Sub
                        End If
                    Case 1
                        If Not W_CHKbFlashPay(tVB_TenderCode, cPay) Then
                            cVB_Pay = 0                 '*Ao 2014/12/04      V 4.2.159   CommSheet TK-ISS3000-568
                            Exit Sub
                        End If
                    Case 2
                        If Not W_CHKbNETPay(cPay) Then
                            cVB_Pay = 0                 '*Ao 2014/12/04      V 4.2.159   CommSheet TK-ISS3000-568
                            Exit Sub
                        End If
                    Case 3
    '                    nTenderType = eEN_PrePaid  '*Eaw 56-06-15 ไม่มีแล้ว
                        '*Tao 62-07-19 ของเดิม
                        'cVB_Pay = 0                 '*Ao 2014/12/04      V 4.2.159   CommSheet TK-ISS3000-568
                        'Exit Sub
                        
                        '*Tao 62-07-19 NETS QR
                        If Not W_CHKbFlashPay(tVB_TenderCode, cPay) Then
                            cVB_Pay = 0                 '*Ao 2014/12/04      V 4.2.159   CommSheet TK-ISS3000-568
                            Exit Sub
                        End If
                End Select
        End Select
        otbNo.Text = tW_ApprCode '*Eaw 56-08-30
    End If
    tVB_CrdApprCode = Trim(otbNo.Text)
    Select Case tVB_TenderCode      '*Ao 2013/01/04
        Case "T015"     'NETS
            'none
        Case Else
            tVB_NETType = ""
            tVB_CrdType = "" '*Eaw 56-07-12
    End Select
    cVB_Pay = cPay
'    tVB_CrdType = "" '*Eaw 56-06-18
    tVB_IsFreeRange = ""  '*Eaw 56-06-18
    '*Eaw 56-04-29 NETS Port
    'แสดงว่ารูดบัตร ผ่าน
    If bW_EdcCheck Then
    'ซ่อน Swipe card แต่ไม่ปลด lock ปุ่ม ok, cancel
        ofaSwipeCard.Left = 2900        'ซ่อน
        ofaSwipeCard.Top = 5400
    End If
    '-----------------------------------------------------------
    '*Eaw 56-08-27
'    Select Case tVB_TenderCode
'        Case "T015"
'            If bW_ManaulMode Then
'                tVB_NetDesc = W_GETtNetType(nVB_CNNETSType)
'            Else
'                tVB_NetDesc = Mid(tVB_CrdType, 3)
'            End If
'        Case "T013", "T016" '*Eaw 56-09-17
'            If bW_ManaulMode Then
'                tVB_NetDesc = tVB_TenderName
'            Else
'                tVB_NetDesc = tVB_TenderType
'            End If
'    End Select
    Unload Me
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)   '*BG 56-11-12  ตอนนี้จ่ายเงินแล้วจะไม่ให้ Scanner ทำงานเพราะจะกระทบกับข้อมูลของเครื่อง NETs ที่ได้มา
'On Error GoTo ErrHandle
'    With oScanner
'        otbNo.Text = .ScanData
''        DoEvents '*Ao 53-01-14   DoEvents
'        If otbNo.Text <> "" Then
'            If Asc(Right(otbNo.Text, 1)) = 13 Then
'                otbNo.Text = Mid(otbNo.Text, 1, Len(otbNo.Text) - 1)
'            End If
'            Call SP_OBJxSetFocus(otbNo)   '*Eaw 55-11-27
'        End If
'        'ทำการ Open Scanner อีกครั้ง ไม่นั้นครั้งต่อไปจะ Scan สินค้าไม่ได้ครับ
'        '*KT 52-01-08 ไม่ต้อง Open Scanner อีกครั้ง ยกเว้นกรณี Form Active
'        If bVB_OPOSScannerEnable Then
'            'With oScanner
'                    If .DeviceEnabled = False Then
'                         .Open tVB_Scanner
'                        .Claim 0 'TeamPos2000
'                         'Enable the device.
'                         .DeviceEnabled = True
'                         'Enable the event.
'                         .DataEventEnabled = True
'                         .AutoDisable = True
'                         .ClearInput
'                    Else
'                        .ClearInput
'                    End If
'            'End With
'        End If
'    End With
'    Exit Sub
'ErrHandle:
'    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub
'
Private Sub otbNo_GotFocus()
    otbNo.SelStart = 0
    otbNo.SelLength = Len(otbNo.Text)
End Sub

Private Sub otbNum_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------------
'Cmt:
'---------------------------------------------------------------------------
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case vbKeyDelete         'delete
            otbNum.Text = ""
        Case 8        'backspace
            'do nothing
        Case 13
            If otbNum.Text <> "" Then
                SendKeys "{TAB}"
            End If
        Case Else  '*Eaw 56-03-02 เพิ่ม case นี้ขึ้นมา
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            Else
                otbNum.MaxLength = 11  '*Eaw 56-03-02 TK-ISS3000-063 ถ้าคีย์มูลค่าเองให้คีย์ได้ 8+2
            End If
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
Private Sub otbNo_KeyPress(KeyAscii As Integer)
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
Private Sub otmForm_Timer()
'---------------------------------------------------------------------
'Cmt :  '*Eaw 56-04-29 NETS Port   เพิ่ม otmForm เข้ามาใน form
'------------------------------------------------------------------------------
    nW_Timer = nW_Timer + 1
    If bW_HandleESC Then
        If nW_Timer = nW_TimeESC Then
            bW_HandleESC = False
            ofa3Sec.Left = 120      'ซ่อน
            ofa3Sec.Top = 5400
            olaForm(11).Caption = ""
            ofaSwipeCard.Left = 750   'center
            ofaSwipeCard.Top = 120
            otcForm1(1).Enabled = False     'lock text box
            otcForm1(2).Enabled = False     'lock ปุ่ม
            otmForm.Enabled = False  'stop otmForm
            otmEDC.Enabled = True     'start otmEDC
            
        ElseIf nW_Timer < nW_TimeESC Then
            olaForm(9).Caption = nW_TimeESC - nW_Timer
        End If
    End If
End Sub

Private Sub otmEDC_Timer()
'---------------------------------------------------------------------
'Cmt :    '*Eaw 56-04-29 NETS Port  เพิ่ม otmEDC เข้ามาใน form
'------------------------------------------------------------------------------
    If bW_ManaulMode = False Then
        'EDC/NETS job
'        Select Case tVB_TypeOfEDC
'            Case "4"  'Cash card
'                Call ocmOK_Click
'            Case "5"   'Flash pay
'                Call ocmOK_Click
'        End Select
        Call ocmOK_Click
    End If
End Sub
