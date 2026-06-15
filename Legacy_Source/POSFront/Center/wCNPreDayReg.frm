VERSION 5.00
Object = "{D6810AA0-45D7-11CF-88F3-02608CA3BE43}#1.0#0"; "Msr.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNPreDayReg 
   Caption         =   "Preview Day Register"
   ClientHeight    =   4080
   ClientLeft      =   2880
   ClientTop       =   2325
   ClientWidth     =   7170
   ControlBox      =   0   'False
   Icon            =   "wCNPreDayReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7170
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
      Height          =   510
      Left            =   5280
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton ocmOk 
      Caption         =   "&Ok;klg_Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3600
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   3360
      Width           =   1455
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3480
      Index           =   0
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   -240
      Width           =   7140
      _Version        =   720897
      _ExtentX        =   12594
      _ExtentY        =   6138
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1080
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
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox otbMemberID 
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
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   0
         Top             =   480
         Width           =   4215
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
         Left            =   3240
         TabIndex        =   12
         Top             =   1080
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
         Left            =   4320
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   1155
         Width           =   2775
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Member ID:;klg_MemberID"
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
         TabIndex        =   9
         Tag             =   "2;"
         Top             =   600
         Width           =   3480
      End
      Begin VB.Label olaName 
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
         Left            =   2520
         TabIndex        =   8
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Member Name:;klg_MemberName"
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
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1800
         Width           =   4350
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1110
      Index           =   2
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3240
      Width           =   7140
      _Version        =   720897
      _ExtentX        =   12594
      _ExtentY        =   1958
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      PaintManager.Position=   2
      Begin SCANNERLib.Scanner oScanner 
         Left            =   480
         Top             =   0
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin MSRLib.MSR OPOSMSR1 
         Left            =   0
         Top             =   0
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.CommandButton ocmFind 
         Caption         =   "&Find;klg_Find"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1920
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNPreDayReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bW_FoundMemberCard  As Boolean
Private bW_KeyEnter As Boolean
Private tW_IsChkDigitMallCard As String

Public nW_TrackNo  As Long '*KT 55-07-10 Trackdata Number of MSR

Private Sub Form_Activate()
On Error GoTo ErrHandle
    Call SP_OBJxSetFocus(otbMemberID)
    Exit Sub
ErrHandle:
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF4
                If Shift = 0 And nVB_RegPreType = 2 Then
                    Call ocmFind_Click
                End If
        End Select
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    bW_FoundMemberCard = False
    Select Case nVB_RegPreType
        Case 1 'Credit Card
            wCNPreDayReg.Caption = ";klg_Title1"
            olaForm(0).Caption = ";klg_ID1"
            olaForm(1).Caption = ";klg_Name1"
        Case 2 'Member Card
            wCNPreDayReg.Caption = ";klg_Title2"
            olaForm(0).Caption = ";klg_ID2"
            olaForm(1).Caption = ";klg_Name2"
    End Select
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)

    bVB_FormPayAct = True 'PH 2.0 ML-SF-019 ไม่ต้องการให้ทำที่ Form_Activate ของ wMain
    If bVB_OPOSMSREnable Then
        With OPOSMSR1
            'Open the device
            'Use a Logical Device Name which has been set on the SetupPOS.
            .Open tVB_CashMsr
            
            'Get the exclusive control right for the opened device.
            'Then the device is disable from other application.
            
            '(Notice:When using an old CO, use the Claim.)
'            .ClaimDevice 1000 'TeamPos3000
            .Claim 1000 'TramPos2000
    
            'Enable the device.
            .DeviceEnabled = True
            
            'Enable the event.
            .DataEventEnabled = True
            
            'Stores the read data to the each property.
            .ParseDecodeData = True
            
            If Not .DeviceEnabled Then
                Call SP_MSGnShowing("tms_CN005092", nCS_Error)
            End If
        End With
    End If
    
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
    Select Case nVB_RegPreType
        Case 1 'Credit card
            ocmFind.Visible = False
        Case 2 'Mall Card
            ocmFind.Visible = True     '*Mee 55-01-06 Commsheet ML-PH4-R3-018
        Case 3 'Non Card
            ocmFind.Visible = False
        Case Else
            Unload Me
    End Select
    olaName.BackColor = nVB_CNClrBLock
    Call SP_OBJxSetFocus(otbMemberID)
    bW_KeyEnter = False
    Exit Sub
ErrHandle:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    With OPOSMSR1
        'Cancel the device
        .DeviceEnabled = False
        
        'Release the device exclusive control right.
        '(Notice:When using an old CO, use the Release.)
'        .ReleaseDevice 'TramPos3000
        .Release ''TramPos2000

        'Finish using the device.
        .Close
    End With
    
    With oScanner
        If bVB_OPOSScannerEnable Then
            If .DeviceEnabled Then
                .DeviceEnabled = False
                .Release
                .Close
            End If
        End If
    End With
'    Set wCNDiscMall = Nothing
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    Unload Me
End Sub

Private Sub ocmFind_Click()
    tVB_CNPreDayIDReg = ""
    tVB_DiscMallCardIDByTel = ""
    bW_FoundMemberCard = False
    wCNFindMem.Show vbModal
'    DoEvents *KT 53-01-14 DoEvents
    tVB_CNPreDayIDReg = tVB_DiscMallCardIDByTel
    If tVB_CNPreDayIDReg <> "" Then
        otbMemberID.Text = tVB_CNPreDayIDReg
        Call otbMemberID_KeyPress(13)
        bW_FoundMemberCard = True     '*Mee 55-01-06 Commsheet ML-PH4-R3-018
'        If bW_FoundMemberCard Then Call SP_OBJxSetFocus(ocmOk)     '*Mee 55-01-06 Commsheet ML-PH4-R3-018
        If bW_FoundMemberCard Then Call SP_OBJxSetFocus(otbMemberID)     '*Mee 55-01-06 Commsheet ML-PH4-R3-018
    Else
        Call SP_OBJxSetFocus(otbMemberID)
    End If
    bW_FoundMemberCard = False      '*Mee 55-01-06 Commsheet ML-PH4-R3-018
End Sub

Private Sub ocmOK_Click()
    Dim dDateExp As Date
On Error GoTo ErrHandle
        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = False
        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
        Select Case nVB_RegPreType
            Case 1 'Credit Card
                'Check Month
                If Format(orbMonth.Text, "00") = "00" Then
                    Call SP_MSGnShowing("tms_CN005143", nCS_Warning)
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    Call SP_OBJxSetFocus(orbMonth)
                    bW_KeyEnter = False
                    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                    Exit Sub
                End If
                'Check Year
                If Format(orbYear.Text, "00") = "00" Then
                    Call SP_MSGnShowing("tms_CN005143", nCS_Warning)
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    Call SP_OBJxSetFocus(orbYear)
                    bW_KeyEnter = False
                    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                    Exit Sub
                End If
                'Check Expire  Date
                dDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)
                If Format(dDateExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  เปรียบเทียบน้อยกว่าเพราะdVB_CrdDateExp + ขึ้นไป 1 วันจากที่ระบุ Communication Sheet PH2-ML-SF-027 Credit Expire.xls
                    Call SP_MSGnShowing("tms_CN005143", nCS_Warning)
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    Call SP_OBJxSetFocus(orbMonth)
                    bW_KeyEnter = False
                    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                    Exit Sub
                End If
                If Not bW_FoundMemberCard Then
                    tVB_CNPreDayIDReg = Trim(otbMemberID.Text)
                    If tVB_CNPreDayIDReg = "" Then
'                        Call SP_MSGnShowing("tms_CN005065", nCS_Error)
                        Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                        bW_FoundMemberCard = False
                        otbMemberID.Text = ""
                        olaName.Caption = ""
                        Call SP_OTBxSelText(otbMemberID)
                        bW_KeyEnter = False
                        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                        Exit Sub
                    End If
                    If Not W_PRCxFindCrdByAccNum(tVB_CNPreDayIDReg) Then
'                        Call SP_MSGnShowing("tms_CN005065", nCS_Error)
                        Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                        bW_FoundMemberCard = False
                        otbMemberID.Text = ""
                        olaName.Caption = ""
                        Call SP_OTBxSelText(otbMemberID)
                        bW_KeyEnter = False
                        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                        Exit Sub
                    Else
                        bW_FoundMemberCard = True
                    End If
                End If
                bVB_CNTrStaRegPreviewDay = True
            Case 2 'Mall Card
                'เงื่อนไขดังต่อไปนี้
                'กรณี Member Card ระบบจะไม่ตรวจสอบกับ Member Server
                'เป็นประเภทบัตรที่อนุญาตให้ทำ Preview Day ได้ ตามที่กำหนดในระบบ ISS3000 ในหน้าจอ Mall Card Maintenance โดย Non Card Flag ต้องเป็น No
                'หมายเลขบัตร M-Card มีข้อมูลตาม Card Range ในระบบ ISS3000 ในหน้าจอ Mall Card Maintenance
                'วันหมดอายุในตัวบัตรที่ได้จากการคีย์หรือการรูดบัตร
                'กรณีที่ประเภทบัตรกำหนดให้ตรวจสอบ Rank(MOD 10) ให้ระบบทำการตรวจสอบหมายเลขตามสูตร MOD10
                
                'Check Month
                If Format(orbMonth.Text, "00") = "00" Then
                    Call SP_MSGnShowing("tms_CN005141", nCS_Warning)
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    Call SP_OBJxSetFocus(orbMonth)
                    bW_KeyEnter = False
                    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                    Exit Sub
                End If
                'Check Year
                If Format(orbYear.Text, "00") = "00" Then
                    Call SP_MSGnShowing("tms_CN005141", nCS_Warning)
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    Call SP_OBJxSetFocus(orbYear)
                    bW_KeyEnter = False
                    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                    Exit Sub
                End If
                'Check Expire  Date
                dDateExp = DateSerial(2000 + Val(Format(orbYear.Text)), Val(Format(orbMonth.Text)) + 1, 1)
                If Format(dDateExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then '*KT 51-12-09 V 1.0.3 Phase 2 ข้อ 26  เปรียบเทียบน้อยกว่าเพราะdVB_CrdDateExp + ขึ้นไป 1 วันจากที่ระบุ Communication Sheet PH2-ML-SF-027 Credit Expire.xls
                    Call SP_MSGnShowing("tms_CN005141", nCS_Warning)
                    orbMonth.Text = ""
                    orbYear.Text = ""
                    Call SP_OBJxSetFocus(orbMonth)
                    bW_KeyEnter = False
                    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                    Exit Sub
                End If
                If Not bW_FoundMemberCard Then
                   tVB_CNPreDayIDReg = Trim(otbMemberID.Text)
                   If tVB_CNPreDayIDReg = "" Then
'                        Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                        Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                        bW_FoundMemberCard = False
                        otbMemberID.Text = ""
                        olaName.Caption = ""
                        orbMonth.Text = ""
                        orbYear.Text = ""
                        Call SP_OTBxSelText(otbMemberID)
                        bW_KeyEnter = False
                        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                        Exit Sub
                    End If
                End If
                If Not W_PRCxFindMallCardRangByID(tVB_CNPreDayIDReg) Then
'                        Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                        Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                        bW_FoundMemberCard = False
                        otbMemberID.Text = ""
                        olaName.Caption = ""
                        orbMonth.Text = ""
                        orbYear.Text = ""
                        Call SP_OTBxSelText(otbMemberID)
                        bW_KeyEnter = False
                        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                        Exit Sub
                End If
                bVB_CNTrStaRegPreviewDay = True
        End Select
        Unload Me
        Exit Sub
ErrHandle:
        bW_KeyEnter = False
        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
End Sub

Private Sub OPOSMSR1_DataEvent(ByVal Status As Long)
    Dim tExpDate$
    Dim dCrdDateExp As Date
    Dim tTranc2$
On Error GoTo ErrHandle
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
    With OPOSMSR1
        
        Select Case nVB_RegPreType
            Case 1 'Credit Card
                    otbMemberID.Text = .AccountNumber     '(Account No.)
                    tExpDate = .ExpirationDate    '(Valid date)
                    orbYear.Text = Mid(tExpDate, 1, 2)
                    orbMonth.Text = Mid(tExpDate, 3, 2)
                    dCrdDateExp = DateSerial(2000 + Val(Format(Mid(tExpDate, 1, 2))), Val(Format(Mid(tExpDate, 3, 2))) + 1, 1)
                    .ClearInput
                    .DataEventEnabled = True
                    bW_FoundMemberCard = True
                    'ตรวจสอบความถูกต้องของเลขที่
                    If Not W_PRCxFindCrdByAccNum(Trim(otbMemberID.Text)) Then
'                        Call SP_MSGnShowing("tms_CN005065", nCS_Error)
                        Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                         bW_FoundMemberCard = False
                         otbMemberID.Text = ""
                         olaName.Caption = ""
                         orbYear.Text = ""
                         orbMonth.Text = ""
                         Call SP_OTBxSelText(otbMemberID)
                         If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                         Exit Sub
                    Else
                        'If Format(dCrdDateExp, "YYYY/MM/DD") <= Format(SP_GETdChkDatePreDay, "YYYY/MM/DD") Then
                        If Format(dCrdDateExp, "YYYY/MM/DD") <= Format(dVB_CurrentDate, "YYYY/MM/DD") Then
                            Call SP_MSGnShowing("tms_CN005143", nCS_Warning)
                            otbMemberID.Text = ""
                            olaName.Caption = ""
                            orbYear.Text = ""
                            orbMonth.Text = ""
                            Call SP_OTBxSelText(otbMemberID)
                            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                            Exit Sub
                        End If
                        bW_FoundMemberCard = True
                    End If
                    Call SP_OBJxSetFocus(ocmOk)
            Case 2 'Mall Card
                    'tTranc2 = .Track2Data
                    Select Case nW_TrackNo
                        Case 1
                             tTranc2 = .Track1Data
                        Case 2
                             tTranc2 = .Track2Data
                        Case 3
                             tTranc2 = .Track3Data
                        Case 4
                             tTranc2 = .Track4Data
                        Case Else
                            tTranc2 = .Track2Data
                    End Select
                    .ClearInput
                    .DataEventEnabled = True
                    If tTranc2 <> "" Then
                         If Len(Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))) <= 20 Then
                            otbMemberID.Text = Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))
                            tExpDate = Trim(Mid(tTranc2, InStr(1, tTranc2, "=") + 1, Len(tTranc2) - InStr(1, tTranc2, "=")))
                            orbYear.Text = Mid(tExpDate, 1, 2)
                            orbMonth.Text = Mid(tExpDate, 3, 2)
                           .DataEventEnabled = True
                        Else
'                            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                            Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                            otbMemberID.Text = ""
                            olaName.Caption = ""
                            Call SP_OTBxSelText(otbMemberID)
                            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                            Exit Sub
                        End If
                    Else
'                        Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                        Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                        otbMemberID.Text = ""
                        olaName.Caption = ""
                        Call SP_OTBxSelText(otbMemberID)
                        Exit Sub
                    End If
                     tVB_CNPreDayIDReg = Trim(otbMemberID.Text)
                    If Not W_PRCxFindMallCardRangByID(tVB_CNPreDayIDReg) Then
'                        Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                        Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                        bW_FoundMemberCard = False
                        otbMemberID.Text = ""
                        olaName.Caption = ""
                        orbMonth.Text = ""
                        orbYear.Text = ""
                        Call SP_OTBxSelText(otbMemberID)
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                        Exit Sub
                    Else
                        If tW_IsChkDigitMallCard = "1" Then
                            If Not SP_PRCbIsChkDigitMallCard(Trim(tVB_CNPreDayIDReg)) Then
                                Call SP_MSGnShowing("tms_CN005216", nCS_Error)
                                bW_FoundMemberCard = False
                                otbMemberID.Text = ""
                                olaName.Caption = ""
                                orbMonth.Text = ""
                                orbYear.Text = ""
                                Call SP_OTBxSelText(otbMemberID)
                                bW_KeyEnter = False
                                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                                Exit Sub
                            End If
                        End If
                        bW_FoundMemberCard = True
                    End If
                    Call SP_OBJxSetFocus(ocmOk)
        End Select
    End With
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_OBJxSetFocus(otbMemberID)
     If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
End Sub

Private Function W_PRCxFindMemberCardByID(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตร Mall Card อยู่ในช่วงใด
'          กรรีที่ไม่สามารถติดต่อฐานข้อมูล Member Server ได้ให้อนุญาตให้ทำ Function Member Price ได้ และกรณีที่หมายเลข member ที่รูดหรือระบุไม่มีในฐานข้อมูลให้ถือว่า
'         เป็น New Member สามารถทำ Function Member Price ได้เหมือนกัน
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle

    W_PRCxFindMemberCardByID = False
    

    If ocnMember Is Nothing Then
        W_PRCxFindMemberCardByID = True
         Call SP_DATxRsNothing(orsTemp)
        Exit Function
    End If
    
    tSql = "SELECT Top 1 FTMcdName,FTMcdSurName,FTMcdCode,FTMcdTel,FDMcdExpDate,FTMcdSta  FROM TCNMMallCard"
    tSql = tSql & vbCrLf & " WHERE FTMcdCode='" & ptCusID & "'"
'    FTMcdSta='A'"  'สถานะการใช้งานของบัตร  มีค่าเป็น  1=Active/ปกติ   , 2=Expire/หมดอายุ , 3=Lose/สูญหาย, 4=Hold/ระงับ
    Call SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        If CDate(Format(SP_FEDdChkDate(orsTemp, "FDMcdExpDate"), "YYYY/MM/DD")) >= CDate(Format(SP_GETdChkDateDisc, "YYYY/MM/DD")) Then '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
            Select Case SP_FEDtChkString(orsTemp, "FTMcdSta")
                Case "A", "1"
                    W_PRCxFindMemberCardByID = True
                Case "E", "2"
                    Call SP_MSGnShowing("tms_CN005201", tCS_CNMsgErr)
                    W_PRCxFindMemberCardByID = False
                Case "L", "3"
                    Call SP_MSGnShowing("tms_CN005202", tCS_CNMsgErr)
                    W_PRCxFindMemberCardByID = False
                Case "H", "4"
                    Call SP_MSGnShowing("tms_CN005203", tCS_CNMsgErr)
                    W_PRCxFindMemberCardByID = False
                Case Else 'นอกจากนี้อนุญาตให้ผ่าน
                    W_PRCxFindMemberCardByID = True
            End Select
        Else
            Call SP_MSGnShowing("tms_CN005201", tCS_CNMsgErr)
            W_PRCxFindMemberCardByID = False
        End If
        olaName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName") & " " & SP_FEDtChkString(orsTemp, "FTMcdSurName")
    Else 'ไม่พบใน Member server ให้ถือว่าเป็นบัตรใหม่
        olaName.Caption = "New Member"
        W_PRCxFindMemberCardByID = True
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    olaName.Caption = "New Member"
    Call SP_DATxRsNothing(orsTemp)
    W_PRCxFindMemberCardByID = True
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
 '           SendKeys "{Tab}"
            '------------------------------------------------------------------
            If bW_KeyEnter Then
                KeyAscii = 0
            Else
                bW_KeyEnter = True
                Call ocmOK_Click
            End If
            '-------------------------------------------------------------------
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub


Private Sub oScanner_DataEvent(ByVal Status As Long)
On Error GoTo ErrHandle
        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = False
        With oScanner
            otbMemberID.Text = .ScanData
'            DoEvents '*Ao 53-01-14   DoEvents
            If otbMemberID.Text <> "" Then
                '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
                '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
                If Asc(Right(otbMemberID.Text, 1)) = 13 Then
                    otbMemberID.Text = Mid(otbMemberID.Text, 1, Len(otbMemberID.Text) - 1)
                End If
                Call otbMemberID_KeyPress(13)
            End If
'            .ClearInput
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
        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
    otbMemberID.Text = ""
    Call SP_OTBxSelText(otbMemberID)
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
End Sub

Private Sub otbMemberID_Change()
    bW_FoundMemberCard = False
End Sub

Private Sub otbMemberID_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbMemberID.Text = ""
        Case 13                            'Enter
            Select Case nVB_RegPreType
                    Case 1 'Credit Card
                        If Not bW_FoundMemberCard Then
                        tVB_CNPreDayIDReg = Trim(otbMemberID.Text)
                        If tVB_CNPreDayIDReg = "" Then
'                                  Call SP_MSGnShowing("tms_CN005065", nCS_Error)
                                Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                                bW_FoundMemberCard = False
                                otbMemberID.Text = ""
                                olaName.Caption = ""
                                Call SP_OTBxSelText(otbMemberID)
                                Exit Sub
                            End If
                        If Not W_PRCxFindCrdByAccNum(tVB_CNPreDayIDReg) Then
'                                  Call SP_MSGnShowing("tms_CN005065", nCS_Error)
                                Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                                bW_FoundMemberCard = False
                                otbMemberID.Text = ""
                                olaName.Caption = ""
                                Call SP_OTBxSelText(otbMemberID)
                                Exit Sub
                            Else
                                bW_FoundMemberCard = True
                            End If
                        End If
                Case 2  'Member Card
                        tVB_CNPreDayIDReg = Trim(otbMemberID.Text)
                        If tVB_CNPreDayIDReg = "" Then
'                            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                            Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                            bW_FoundMemberCard = False
                            otbMemberID.Text = ""
                            olaName.Caption = ""
                            Call SP_OTBxSelText(otbMemberID)
                            Exit Sub
                        End If
                        If Not W_PRCxFindMallCardRangByID(tVB_CNPreDayIDReg) Then
'                            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                            Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                            bW_FoundMemberCard = False
                            otbMemberID.Text = ""
                            olaName.Caption = ""
                            orbMonth.Text = ""
                            orbYear.Text = ""
                            Call SP_OTBxSelText(otbMemberID)
                            bW_KeyEnter = False
                            If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                            Exit Sub
                        Else
                            If tW_IsChkDigitMallCard = "1" Then
                                If Not SP_PRCbIsChkDigitMallCard(Trim(tVB_CNPreDayIDReg)) Then
                                    Call SP_MSGnShowing("tms_CN005216", nCS_Error)
                                    bW_FoundMemberCard = False
                                    otbMemberID.Text = ""
                                    olaName.Caption = ""
                                    orbMonth.Text = ""
                                    orbYear.Text = ""
                                    Call SP_OTBxSelText(otbMemberID)
                                    bW_KeyEnter = False
                                    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                                    Exit Sub
                                End If
                            End If
                            bW_FoundMemberCard = True
                        End If
            End Select
            
            If Trim(otbMemberID.Text) <> "" Then
                SendKeys "{Tab}"
            Else
                Call SP_OBJxSetFocus(otbMemberID)     '*Mee 55-01-09 Commsheet ML-PH4-R3-018
            End If
    '*Mee 55-01-09 Commsheet ML-PH4-R3-018
    Case 115
        Select Case nVB_RegPreType
            Case 2  'Member Card
                        tVB_CNPreDayIDReg = Trim(otbMemberID.Text)
                        If tVB_CNPreDayIDReg = "" Then
'                            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
'                            Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                            bW_FoundMemberCard = False
                            otbMemberID.Text = ""
                            olaName.Caption = ""
                            orbMonth.Text = ""
                            orbYear.Text = ""
                            Call SP_OTBxSelText(otbMemberID)
                            Exit Sub
                        End If
                        If Not W_PRCxFindMallCardRangByID(tVB_CNPreDayIDReg) Then
'                            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
'                            Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                            bW_FoundMemberCard = False
                            otbMemberID.Text = ""
                            olaName.Caption = ""
                            orbMonth.Text = ""
                            orbYear.Text = ""
                            Call SP_OTBxSelText(otbMemberID)
                            bW_KeyEnter = False
                            If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                            Exit Sub
                        Else
                            If tW_IsChkDigitMallCard = "1" Then
                                If Not SP_PRCbIsChkDigitMallCard(Trim(tVB_CNPreDayIDReg)) Then
                                    Call SP_MSGnShowing("tms_CN005216", nCS_Error)
                                    bW_FoundMemberCard = False
                                    otbMemberID.Text = ""
                                    olaName.Caption = ""
                                    orbMonth.Text = ""
                                    orbYear.Text = ""
                                    Call SP_OTBxSelText(otbMemberID)
                                    bW_KeyEnter = False
                                    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                                    Exit Sub
                                End If
                            End If
                            bW_FoundMemberCard = True
                        End If
            End Select
            '*Mee 55-01-09 Commsheet ML-PH4-R3-018
            If Trim(otbMemberID.Text) <> "" Then
                SendKeys "{Tab}"
            Else
                Call SP_OBJxSetFocus(otbMemberID)
            End If
    End Select
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_OBJxSetFocus(otbMemberID)
End Sub

Private Sub otbMemberID_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle
    Call UT_ChkKeyAscii(KeyAscii)
    Call SP_ChkKeyAsciiNumber(KeyAscii)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_OBJxSetFocus(otbMemberID)
End Sub

Private Function W_PRCxFindMallCardRangByID(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหาหมายเลขบัตร Member Card ใน ISS3000 PreviewDay Maintennance
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle

    W_PRCxFindMallCardRangByID = False
    tSql = "STP_PRCxFindMallCardByIDPreDay @ptAccNum= '" & ptCusID & "',@dDate = '" & Format(SP_GETdChkDatePreDay, "YYYY/MM/DD") & "'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        W_PRCxFindMallCardRangByID = True
        olaName.Caption = SP_FEDtChkString(orsTemp, "FTMcdDesc1")
        tW_IsChkDigitMallCard = SP_FEDtChkString(orsTemp, "FTMcdIsChkDigit")
        aW_CardPreDay(1).tStaDeptCode = SP_FEDtChkString(orsTemp, "FTMcdStaDeptCodePreDay")
        aW_CardPreDay(1).tCrdType = SP_FEDtChkString(orsTemp, "FTMcdCode")
    Else
        W_PRCxFindMallCardRangByID = False
        olaName.Caption = ""
        tW_IsChkDigitMallCard = "0"
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_PRCxFindMallCardRangByID = False
    olaName.Caption = ""
    tW_IsChkDigitMallCard = "0"
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function W_PRCxFindNonCardRangByID(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหาหมายเลขบัตร Non Card ใน ISS3000 PreviewDay Maintennance
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle

    W_PRCxFindNonCardRangByID = False
    tSql = "SELECT FTNcdCode  FROM TCNTNonCard"
    tSql = tSql & vbCrLf & "WHERE FTNcdCode = '" & ptCusID & "'"
    Call SP_SQLvExecute(tSql, orsTemp)
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        W_PRCxFindNonCardRangByID = True
        olaName.Caption = "Non Card"
    Else
        W_PRCxFindNonCardRangByID = False
        olaName.Caption = ""
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    olaName.Caption = ""
    W_PRCxFindNonCardRangByID = False
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function W_PRCxFindCrdCardRangByID(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหาหมายเลขบัตร Credit Card ใน ISS3000 PreviewDay Maintennance
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle

    W_PRCxFindCrdCardRangByID = False
    tSql = "SELECT  FTCcrRangeFrom,FTCcrRangeTo"
    tSql = tSql & vbCrLf & "From TCNMCreditCardRange"
    tSql = tSql & vbCrLf & "Where '" & ptCusID & "' BETWEEN FTCcrRangeFrom AND FTCcrRangeTo"
    tSql = tSql & vbCrLf & "AND LEN('" & ptCusID & "') <= LEN(FTCcrRangeTo)"
    tSql = tSql & vbCrLf & "AND FTCcrAllow = '1'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        W_PRCxFindCrdCardRangByID = True
    Else
        W_PRCxFindCrdCardRangByID = False
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_PRCxFindCrdCardRangByID = False
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function W_PRCxFindCrdByAccNum(ByVal ptAccNum$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตรเครดิตว่าอยุในช่วงใด
'การค้นหาบัตรเครดิตที่มีช่วงหมายเลขทับซ้อนกันหมากกว่า 1 ประเภทบัตรเครดิตทำให้ระบบดึงประเภทบัตรเครดิตตาม record ที่เกิดก่อนหลัง จึงมีบัตรที่เป็น Master Card 2 ประเภท คือ
'Master Inter  = 5000000000000001-5999999999999999
'SCB MC =        5577550000000000-5577729999999999
'ดังนั้นไม่จะระบุหมายเลขในช่วง SCB MC อย่างไรก็จะได้ Master Inter
'Order By FNCdcDigitChk จากมาก - น้อย
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle '*KT 52-02-11  PH2-ML-SF-092 โปรแกรมหลุดมาที่ Desktop ตอนพิมพ์ receipt , ชำระด้วย Credit
    W_PRCxFindCrdByAccNum = False
    tSql = "SELECT  TOP 1 TCNMCreditCard.FCCdcAmtMin, TCNMCreditCard.FCCdcAmtMax, TCNMCreditCard.FTCdcOnTop, TCNMCreditCard.FTCdcMallCard,"
    tSql = tSql & vbCrLf & " TCNMCreditCard.FNCdcOntopRate, TCNMCreditCard.FNCdcDigitChk, TCNMCreditCard.FNCdcDigits, TCNMCreditCard.FTCdcDescRe,"
    tSql = tSql & vbCrLf & " TCNMCreditCard.FTCdcDescCredit , TCNMCreditCard.FTCdcCreditType, TCNMCreditCard.FTTdmCode,"
    tSql = tSql & vbCrLf & " TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FTCdcStaDeptCodePreDay"
    tSql = tSql & vbCrLf & " FROM   TCNMCreditCard INNER JOIN"
    tSql = tSql & vbCrLf & " TCNMCreditCardRange ON TCNMCreditCard.FTTdmCode = TCNMCreditCardRange.FTTdmCode AND"
    tSql = tSql & vbCrLf & " TCNMCreditCard.FTCdcCreditType = TCNMCreditCardRange.FTCdcCreditType"
    tSql = tSql & vbCrLf & " WHERE  LEFT('" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk)  "
    tSql = tSql & vbCrLf & " AND LEN('" & ptAccNum & "') = TCNMCreditCard.FNCdcDigits"
    tSql = tSql & vbCrLf & " AND ISNULL(TCNMCreditCard.FTCdcPreDay,'N') = 'Y'" '
    tSql = tSql & vbCrLf & " AND ('" & Format(SP_GETdChkDatePreDay, "YYYY/MM/DD") & "' BETWEEN TCNMCreditCard.FDCdcStartDatePreDay AND TCNMCreditCard.FDCdcEndDatePreDay)"
    tSql = tSql & vbCrLf & " AND TCNMCreditCard.FTTdmCode = 'T002'"
    tSql = tSql & vbCrLf & " ORDER BY TCNMCreditCard.FNCdcDigitChk DESC" '*KT 51-12-11 Phase ข้อ 14 Communication Sheet PH2-ML-SF-025 Credit Card Range.xls
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        W_PRCxFindCrdByAccNum = True
        aW_CardPreDay(1).tStaDeptCode = SP_FEDtChkString(orsTemp, "FTCdcStaDeptCodePreDay")
        aW_CardPreDay(1).tTdmCode = "T002"
        aW_CardPreDay(1).tCrdType = SP_FEDtChkString(orsTemp, "FTCdcCreditType")
        olaName.Caption = SP_FEDtChkString(orsTemp, "FTCdcDescCredit")
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    olaName.Caption = ""
    W_PRCxFindCrdByAccNum = False
    Call SP_DATxRsNothing(orsTemp)
End Function


