VERSION 5.00
Object = "{D6810AA0-45D7-11CF-88F3-02608CA3BE43}#1.0#0"; "Msr.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNMember 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MEMBER CARD;klg_Title"
   ClientHeight    =   4095
   ClientLeft      =   2820
   ClientTop       =   3525
   ClientWidth     =   7170
   ControlBox      =   0   'False
   Icon            =   "wCNMember.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7170
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
      Tag             =   "2;"
      Top             =   3360
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
      Height          =   510
      Left            =   5280
      TabIndex        =   4
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
         Left            =   3460
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1200
         Width           =   615
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
         Left            =   2500
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1200
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
         Left            =   2500
         MaxLength       =   20
         TabIndex        =   0
         Top             =   600
         Width           =   4500
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
         Left            =   4300
         TabIndex        =   13
         Tag             =   "2;"
         Top             =   1275
         Width           =   2775
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
         Index           =   1
         Left            =   3200
         TabIndex        =   12
         Top             =   1200
         Width           =   210
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
         Left            =   105
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   1275
         Width           =   2130
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
         Index           =   4
         Left            =   105
         TabIndex        =   10
         Tag             =   "2;รหัสพนักงานขาย;Cashier ID"
         Top             =   1800
         Width           =   4350
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
         Height          =   855
         Left            =   2505
         TabIndex        =   9
         Top             =   1800
         Width           =   4500
         WordWrap        =   -1  'True
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
         Left            =   105
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   600
         Width           =   3480
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1110
      Index           =   2
      Left            =   0
      TabIndex        =   8
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
         Left            =   600
         Top             =   120
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
         Tag             =   "2;"
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bW_FoundMemberCard  As Boolean
Private bW_KeyEnter As Boolean
Private tW_IsChkDigitMallCard As String '1: allow check 0 : not allow check other : not allow check

Public nW_TrackNo  As Long '*KT 55-07-10 Trackdata Number of MSR
'Private Sub Form_Activate()
'On Error GoTo ErrHandle
'    Call SP_OBJxSetFocus(otbMemberID)
' Exit Sub
'ErrHandle:
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If bVB_CNSplash Then KeyCode = 0 '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น
    '*KT 52-06-05 ในหน้าจอ Tender ถ้ากดปุ่ม Price Check จะปรากฏ Screen Find Customer แล้วถ้ารูดบัตร โปรแกรมจะ Hang PH 2.0 ML-SF-012
     '*KT 52-06-05 PH 2.0 ML-SF-013 ในหน้าจอ Discount --> Mall Card --> Price check --> Enter เครื่อง POS จะHang โดยต้อง Restart อย่างเดียว
    If Shift = 0 Then
        Select Case KeyCode
    '        Case 13
    '            SendKeys "{Tab}"
            Case vbKeyF4
                    Call ocmFind_Click
            Case vbKeyBack
                    bW_FoundMemberCard = False
                    xVB_CNMemberP.ID = Trim(otbMemberID.Text)
        End Select
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_STAxMouse(Me, True)
    bW_FoundMemberCard = False
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    If bVB_OPOSMSREnable Then
        With OPOSMSR1
            'Open the device
            'Use a Logical Device Name which has been set on the SetupPOS.
            .Open tVB_CashMsr
            
            'Get the exclusive control right for the opened device.
            'Then the device is disable from other application.
            
            '(Notice:When using an old CO, use the Claim.)
'            If bVB_Res640X480 Then
'            .ClaimDevice 1000 'TeamPos3000
            .Claim 0 'TramPos2000
    
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
    
    olaName.BackColor = nVB_CNClrBLock
    Call SP_OBJxSetFocus(otbMemberID)
    Call SP_STAxMouse(Me, False)
    Exit Sub
ErrHandle:
    Call SP_STAxMouse(Me, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    With OPOSMSR1
        If bVB_OPOSMSREnable Then
            'Cancel the device
            .DeviceEnabled = False
            'Release the device exclusive control right.
            '(Notice:When using an old CO, use the Release.)
    '        .ReleaseDevice 'TramPos3000
            .Release ''TramPos2000
            'Finish using the device.
            .Close
        End If
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
    Exit Sub
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = False '*Ao 10-02-09
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
    Call SP_CRTxLogMallCardEvent("Keyboard", "Cancel/Esc", "Member Price Register", Trim(otbMemberID.Text))
    bVB_CNMemberPrice = False
    xVB_CNMemberP.ID = ""
    xVB_CNMemberP.FName = ""
    xVB_CNMemberP.LName = ""
    xVB_CNMemberP.Tel = ""
    xVB_CNMemberP.Expire = Date
    xVB_CNMemberP.Status = ""
    '------------------------------------------------------
    xVB_CNMemberP.MallCardType = ""
    xVB_CNMemberP.MallCardExpire = Date
    xVB_CNMemberP.MallCardMemberPrice = ""
    Unload Me
End Sub

Private Sub ocmFind_Click()
    
    xVB_CNMemberP.ID = ""
    tVB_DiscMallCardIDByTel = ""
    bW_FoundMemberCard = False
    wCNFindMem.Show vbModal
'    DoEvents        '*KT 53-01-14
    xVB_CNMemberP.ID = tVB_DiscMallCardIDByTel
    If xVB_CNMemberP.ID <> "" Then
        otbMemberID.Text = xVB_CNMemberP.ID
        Call otbMemberID_KeyPress(13)
'        If bW_FoundMemberCard Then Call SP_OBJxSetFocus(ocmOk)     '*Mee 55-01-06 Commsheet ML-PH4-R3-018
        If bW_FoundMemberCard Then Call SP_OBJxSetFocus(orbMonth)     '*Mee 55-01-06 Commsheet ML-PH4-R3-018
    Else
        Call SP_OBJxSetFocus(otbMemberID)
    End If
End Sub

Private Sub ocmOK_Click()
Dim dDateExp As Date
On Error GoTo ErrHandle
     If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = False        '*Ao 10-02-09
     If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
     
    '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
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
     
    Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[1]", "Member Price Register", Trim(otbMemberID.Text))
    If Not bW_FoundMemberCard Then
        Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[2]", "Member Price Register", Trim(otbMemberID.Text))
        xVB_CNMemberP.ID = Trim(otbMemberID.Text)
        If xVB_CNMemberP.ID = "" Then
                Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                bW_FoundMemberCard = False
                otbMemberID.Text = ""
                olaName.Caption = ""
                orbMonth.Text = ""
                orbYear.Text = ""
                Call SP_OBJxSetFocus(otbMemberID)
                Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[3]", "Member Price Register", Trim(otbMemberID.Text))
                If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                bW_KeyEnter = False     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
                xVB_CNMemberP.ID = ""     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
                Exit Sub
        End If
        
        '*KT 52-11-16 PH3-ML-SF-009 Check MOD10 on card Immediately before press OK button
        Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[4]", "Member Price Register", Trim(otbMemberID.Text))
        If Not W_PRCxFindMallCardByIDMemberPrice(xVB_CNMemberP.ID) Then
                Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                bW_FoundMemberCard = False
                otbMemberID.Text = ""
                olaName.Caption = ""
                orbMonth.Text = ""
                orbYear.Text = ""
                Call SP_OBJxSetFocus(otbMemberID)
                Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[5]", "Member Price Register", Trim(otbMemberID.Text))
                If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                bW_KeyEnter = False     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
                xVB_CNMemberP.ID = ""     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
                Exit Sub
        Else
                '*KT 52-09-30 PH2-ML-SF-149 หมายเลขบัตร Mall Card ไม่ตรงกัน ระหว่าง M-Price กับ รูด Mall Card (Mall8)
                Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[6]", "Member Price Register", Trim(otbMemberID.Text))
                If tW_IsChkDigitMallCard = "1" Then
                    If Not SP_PRCbIsChkDigitMallCard(xVB_CNMemberP.ID) Then
                        Call SP_MSGnShowing("tms_CN005215", nCS_Error)
                        bW_FoundMemberCard = False
                        otbMemberID.Text = ""
                        olaName.Caption = ""
                        orbMonth.Text = ""
                        orbYear.Text = ""
                        Call SP_OBJxSetFocus(otbMemberID)
                        Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[7]", "Member Price Register", Trim(otbMemberID.Text))
                        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                        bW_KeyEnter = False     '*Mee 55-01-06 V.4.0.10  Commsheet ML-PH4-R3-017
                        xVB_CNMemberP.ID = ""     '*Mee 55-01-06 V.4.0.10  Commsheet ML-PH4-R3-017
                        Exit Sub
                    End If
                End If
        End If
        
        Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[8]", "Member Price Register", Trim(otbMemberID.Text))
        If Not W_PRCxFindMemberCardByID(xVB_CNMemberP.ID) Then
            '*KT 52-05-26 PH 1.8 ML-SF-076 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ไม่ได้ จะขาย Member Price ไม่ได้
'            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
            bW_FoundMemberCard = False
            otbMemberID.Text = ""
            olaName.Caption = ""
            orbMonth.Text = ""
            orbYear.Text = ""
            Call SP_OBJxSetFocus(otbMemberID)
            Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[9]", "Member Price Register", Trim(otbMemberID.Text))
            If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            bW_KeyEnter = False     '*Mee 55-01-06 V.4.0.10  Commsheet ML-PH4-R3-017
            xVB_CNMemberP.ID = ""     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
            Exit Sub
        Else
            olaName.Caption = xVB_CNMemberP.FName & " " & xVB_CNMemberP.LName
'            DoEvents        '*KT 53-01-14
            bW_FoundMemberCard = True
            Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[10]", "Member Price Register", Trim(otbMemberID.Text))
            bW_KeyEnter = False     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
            xVB_CNMemberP.ID = ""     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
        End If
    End If
    

            
    Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[11]", "Member Price Register", Trim(otbMemberID.Text))
    bVB_CNMemberPrice = True
    Unload Me
    bW_KeyEnter = False     '*Mee 55-01-06 Commsheet ML-PH4-R3-017
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    bW_FoundMemberCard = False
    otbMemberID.Text = ""
    olaName.Caption = ""
    Call SP_OBJxSetFocus(otbMemberID)
    Call SP_CRTxLogMallCardEvent("Keyboard/Keyboard", "OK[12]", "Member Price Register", Trim(otbMemberID.Text))
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
End Sub

Private Sub OPOSMSR1_DataEvent(ByVal Status As Long)
    Dim tExpDate$
    Dim tTranc2$
    Dim tTrack1Data As String, tTrack2Data As String, tTrack As String, tMsg As String
    Dim nPos As Long
    
On Error GoTo ErrHandle
    With oScanner
         If bVB_OPOSScannerEnable Then
                'Enable the event.
                .DataEventEnabled = False
         End If
     End With
    With OPOSMSR1
        ' tTranc2 = .Track2Data
'        Select Case nW_TrackNo
'            Case 1
'                tTranc2 = .Track1Data
'            Case 2
'                tTranc2 = .Track2Data
'            Case 3
'                tTranc2 = .Track3Data
'            Case 4
'                tTranc2 = .Track4Data
'            Case Else
'                tTranc2 = .Track2Data
'        End Select
        '*Ao 2013/05/15         แก้ไขการใช้ MSR
        tTrack1Data = .Track1Data
        tTrack2Data = .Track2Data
        .ClearInput
        .DataEventEnabled = True
'        If tTranc2 <> "" Then
'             If Len(Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))) <= 20 Then
'                otbMemberID.Text = Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))
'                '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
'                tExpDate = Trim(Mid(tTranc2, InStr(1, tTranc2, "=") + 1, Len(tTranc2) - InStr(1, tTranc2, "=")))
'                orbYear.Text = Mid(tExpDate, 1, 2)
'                orbMonth.Text = Mid(tExpDate, 3, 2)
''                .DataEventEnabled = True
'                Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[1]", "Member Price Register", Trim(otbMemberID.Text))
'            Else
'                Call SP_MSGnShowing("tms_CN005200", nCS_Error)
'                otbMemberID.Text = ""
'                olaName.Caption = ""
'                orbYear.Text = ""
'                orbMonth.Text = ""
'                Call SP_OBJxSetFocus(otbMemberID)
'                Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[2]", "Member Price Register", Trim(otbMemberID.Text))
'                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
'                Exit Sub
'            End If
'        Else
'            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
'            otbMemberID.Text = ""
'            olaName.Caption = ""
'            orbYear.Text = ""
'            orbMonth.Text = ""
'            Call SP_OBJxSetFocus(otbMemberID)
'            Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[3]", "Member Price Register", Trim(otbMemberID.Text))
'            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
'            Exit Sub
'        End If
'        xVB_CNMemberP.ID = Trim(otbMemberID.Text)
        '--------------------------------------------------------------------------------------------
        '*Ao 2013/05/15         แก้ไขการใช้ MSR
'        tTrack1Data = "s60000120345389706440548160000000000000140200000000000000000000000060"
'        tTrack2Data = "4538970644054816=14021011387700000947"
        If (tTrack1Data <> "") Or (tTrack2Data <> "") Then
            If tTrack1Data <> "" Then
                tTrack = Mid(tTrack1Data, 11, 16)
                tExpDate = Mid(tTrack1Data, 40, 4)
            Else
                If tTrack2Data <> "" Then
                    If Len(tTrack2Data) > 20 Then
                        nPos = InStrRev(tTrack2Data, "=")
                        If nPos > 1 Then
                            tTrack = Trim(Mid(tTrack2Data, 1, InStr(1, tTrack2Data, "=") - 1))
                            tExpDate = Trim(Mid(tTrack2Data, InStr(1, tTrack2Data, "=") + 1, Len(tTrack2Data) - InStr(1, tTrack2Data, "=")))
                        End If
                    Else
                        tTrack = tTrack2Data
                    End If
                Else
                    tTrack = ""
                End If
            End If
            If tTrack <> "" Then
                otbMemberID.Text = tTrack
                If tExpDate <> "" Then
                    orbYear.Text = Mid(tExpDate, 1, 2)
                    orbMonth.Text = Mid(tExpDate, 3, 2)
                End If
            End If
        Else
            tMsg = "Not found Track2Data. swiped card try again."
            Call SP_MSGnShowing("tms_CN005068" & ";" & tMsg, nCS_Error)
            otbMemberID.Text = ""
            orbYear.Text = ""
            orbMonth.Text = ""
            olaName.Caption = ""
            Call SP_OTBxSelText(otbMemberID)
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
            Exit Sub
        End If
        
'        tTrack = ""
'        If nW_TrackNo = 4 Then                                  'nW_TrackNo = 4  :TC Card
'            If (tTrack1Data <> "") Then
'                tTrack = Mid(tTrack1Data, 1, 26)
'            Else
'                If (tTrack2Data <> "") Then
'                    nPos = InStrRev(tTrack2Data, "=")
'                    If nPos > 1 Then
'                        tTrack = Trim(Mid(tTrack2Data, 1, InStr(1, tTrack2Data, "=") - 1))
'                    End If
'                End If
'            End If
'        Else
'            If (tTrack2Data <> "") Then
'                nPos = InStrRev(tTrack2Data, "=")
'                If nPos > 1 Then
'                    tTrack = Trim(Mid(tTrack2Data, 1, InStr(1, tTrack2Data, "=") - 1))
'                End If
'            Else
'                If (tTrack1Data <> "") Then
'                    tTrack = Mid(tTrack1Data, 11, 16)
'                End If
'            End If
'        End If
        xVB_CNMemberP.ID = Trim(otbMemberID.Text)
        '--------------------------------------------------------------------------------------------

        '*KT 52-11-16 PH3-ML-SF-009 Check MOD10 on card Immediately before press OK button
        Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[4]", "Member Price Register", Trim(otbMemberID.Text))
        If Not W_PRCxFindMallCardByIDMemberPrice(xVB_CNMemberP.ID) Then
            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
            bW_FoundMemberCard = False
            otbMemberID.Text = ""
            olaName.Caption = ""
            orbYear.Text = ""
            orbMonth.Text = ""
            Call SP_OBJxSetFocus(otbMemberID)
            Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[5]", "Member Price Register", Trim(otbMemberID.Text))
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        Else
            '*KT 52-09-30 PH2-ML-SF-149 หมายเลขบัตร Mall Card ไม่ตรงกัน ระหว่าง M-Price กับ รูด Mall Card (Mall8)
            Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[6]", "Member Price Register", Trim(otbMemberID.Text))
            If tW_IsChkDigitMallCard = "1" Then
                If Not SP_PRCbIsChkDigitMallCard(xVB_CNMemberP.ID) Then
                    Call SP_MSGnShowing("tms_CN005215", nCS_Error)
                    bW_FoundMemberCard = False
                    otbMemberID.Text = ""
                    olaName.Caption = ""
                    orbYear.Text = ""
                    orbMonth.Text = ""
                    Call SP_OBJxSetFocus(otbMemberID)
                    Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[7]", "Member Price Register", Trim(otbMemberID.Text))
                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                    Exit Sub
                End If
            End If
        End If
        
        If Not W_PRCxFindMemberCardByID(xVB_CNMemberP.ID) Then
            '*KT 52-05-26 PH 1.8 ML-SF-076 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ไม่ได้ จะขาย Member Price ไม่ได้
'            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
            bW_FoundMemberCard = False
            otbMemberID.Text = ""
            olaName.Caption = ""
            orbYear.Text = ""
            orbMonth.Text = ""
            Call SP_OBJxSetFocus(otbMemberID)
            Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[8]", "Member Price Register", Trim(otbMemberID.Text))
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
            Exit Sub
        Else
            olaName.Caption = xVB_CNMemberP.FName & " " & xVB_CNMemberP.LName
            bW_FoundMemberCard = True
            Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[9]", "Member Price Register", Trim(otbMemberID.Text))
        End If
    End With
    Call SP_OBJxSetFocus(ocmOk)
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    bW_FoundMemberCard = False
    otbMemberID.Text = ""
    olaName.Caption = ""
    orbYear.Text = ""
    orbMonth.Text = ""
    Call SP_OBJxSetFocus(otbMemberID)
    Call SP_CRTxLogMallCardEvent("MSR", "DataEvent[10]", "Member Price Register", Trim(otbMemberID.Text))
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True      '*Ao 10-02-09
End Sub

Private Function W_PRCxFindMemberCardByID(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตร Mall Card อยู่ในช่วงใด
'          กรรีที่ไม่สามารถติดต่อฐานข้อมูล Member Server ได้ให้อนุญาตให้ทำ Function Member Price ได้ และกรณีที่หมายเลข member ที่รูดหรือระบุไม่มีในฐานข้อมูลให้ถือว่า
'         เป็น New Member สามารถทำ Function Member Price ได้เหมือนกัน
'*KT 52-02-12 New Requirment
'1.  เพิ่มหน้าจอ เพื่อรับบัตร Member กรณีที่ลูกค้าต้องการได้ราคา Member Price ทั้งนี้ ขึ้นอยู่กับ Card Type หาก Member Price ไม่ได้ระบุค่าจะเอาราคาตามปกติของสินค้าขณะนั้น มาแสดง
'*KT 52-05-26 PH 1.8 ML-SF-076 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ไม่ได้ จะขาย Member Price ไม่ได้
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tSplash As String
    Dim tCardNo As String
    Dim tTokenCardNo As String
On Error GoTo ErrHandle

    W_PRCxFindMemberCardByID = False
    
    If ocnMember Is Nothing Then                                        '*Ao 2010/12/21  CommSheet PH3-ML-SF-134 Can't connect Member database.xls
        Call SP_CONbMemberServer(True, tSplash, False)
    End If
    '===================================
    '*Bmp 61-02-15 Token
    tTokenCardNo = ""
    tCardNo = ptCusID
    If SP_PRCbToken("Member Price", ptCusID, tTokenCardNo) Then
        tCardNo = tTokenCardNo
    End If
    '===================================
    If ocnMember Is Nothing Then
        'Call SP_MSGnShowing("tms_CN005134", tCS_CNMsgErr)
        xVB_CNMemberP.FName = "New Member"
        xVB_CNMemberP.LName = ""
        xVB_CNMemberP.Tel = ""
        xVB_CNMemberP.Expire = Date
        xVB_CNMemberP.Status = ""
        xVB_CNMemberP.CardIDToken = tCardNo '*Bmp 61-02-15 จะได้ไม่ต้องเรียก web service หลายรอบ
        W_PRCxFindMemberCardByID = True
         Call SP_DATxRsNothing(orsTemp)
        Exit Function
    End If
    tSql = "SELECT Top 1  FTMcdName,FTMcdSurName,FTMcdCode,FTMcdTel,FDMcdExpDate,FTMcdSta  FROM TCNMMallCard"
    tSql = tSql & " WHERE FTMcdCode='" & tCardNo & "'" '*Bmp 61-02-15 Token
    'tSql = tSql & " WHERE FTMcdCode='" & ptCusID & "'"
'    tSql = tSql & " AND FTMcdSta='A'"  'สถานะการใช้งานของบัตร  มีค่าเป็น  1=Active/ปกติ   , 2=Expire/หมดอายุ , 3=Lose/สูญหาย, 4=Hold/ระงับ
'    tSql = tSql & " AND " & Format(dVB_CurrentDate, "dd/mm/yyyy") & " <= FDMcdExpDate"
    Call SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True)
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        If CDate(Format(SP_FEDdChkDate(orsTemp, "FDMcdExpDate"), "YYYY/MM/DD")) >= CDate(Format(SP_GETdChkDateDisc, "YYYY/MM/DD")) Then '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
            Select Case SP_FEDtChkString(orsTemp, "FTMcdSta")
                Case "A", "1"
                    xVB_CNMemberP.FName = SP_FEDtChkString(orsTemp, "FTMcdName")
                    xVB_CNMemberP.LName = SP_FEDtChkString(orsTemp, "FTMcdSurName")
                    xVB_CNMemberP.Tel = SP_FEDtChkString(orsTemp, "FTMcdTel")
                    xVB_CNMemberP.Expire = SP_FEDdChkDate(orsTemp, "FDMcdExpDate")
                    xVB_CNMemberP.Status = SP_FEDtChkString(orsTemp, "FTMcdSta")
                    xVB_CNMemberP.CardIDToken = tCardNo '*Bmp 61-02-15 จะได้ไม่ต้องเรียก web service หลายรอบ
                    W_PRCxFindMemberCardByID = True
                Case "E", "2"
                    Call SP_MSGnShowing("tms_CN005201", tCS_CNMsgErr)
                    xVB_CNMemberP.FName = ""
                    xVB_CNMemberP.LName = ""
                    xVB_CNMemberP.Tel = ""
                    xVB_CNMemberP.Expire = Date
                    xVB_CNMemberP.Status = ""
                    xVB_CNMemberP.CardIDToken = ""
                    W_PRCxFindMemberCardByID = False
                Case "L", "3"
                    Call SP_MSGnShowing("tms_CN005202", tCS_CNMsgErr)
                    xVB_CNMemberP.FName = ""
                    xVB_CNMemberP.LName = ""
                    xVB_CNMemberP.Tel = ""
                    xVB_CNMemberP.Expire = Date
                    xVB_CNMemberP.Status = ""
                    xVB_CNMemberP.CardIDToken = ""
                    W_PRCxFindMemberCardByID = False
                Case "H", "4"
                    Call SP_MSGnShowing("tms_CN005203", tCS_CNMsgErr)
                    xVB_CNMemberP.FName = ""
                    xVB_CNMemberP.LName = ""
                    xVB_CNMemberP.Tel = ""
                    xVB_CNMemberP.Expire = Date
                    xVB_CNMemberP.Status = ""
                    xVB_CNMemberP.CardIDToken = ""
                    W_PRCxFindMemberCardByID = False
                Case Else 'นอกจากนี้อนุญาตให้ผ่าน
                    xVB_CNMemberP.FName = SP_FEDtChkString(orsTemp, "FTMcdName")
                    xVB_CNMemberP.LName = SP_FEDtChkString(orsTemp, "FTMcdSurName")
                    xVB_CNMemberP.Tel = SP_FEDtChkString(orsTemp, "FTMcdTel")
                    xVB_CNMemberP.Expire = SP_FEDdChkDate(orsTemp, "FDMcdExpDate")
                    xVB_CNMemberP.Status = SP_FEDtChkString(orsTemp, "FTMcdSta")
                    xVB_CNMemberP.CardIDToken = tCardNo '*Bmp 61-02-15 จะได้ไม่ต้องเรียก web service หลายรอบ
                    W_PRCxFindMemberCardByID = True
            End Select
        Else
            Call SP_MSGnShowing("tms_CN005201", tCS_CNMsgErr)
            xVB_CNMemberP.FName = ""
            xVB_CNMemberP.LName = ""
            xVB_CNMemberP.Tel = ""
            xVB_CNMemberP.Expire = Date
            xVB_CNMemberP.Status = ""
            xVB_CNMemberP.CardIDToken = ""
            W_PRCxFindMemberCardByID = False
        End If
    Else 'ไม่พบใน Member server ให้ถือว่าเป็นบัตรใหม่
'        Call SP_MSGnShowing("tms_CN005200", tCS_CNMsgErr)
        xVB_CNMemberP.FName = "New Member"
        xVB_CNMemberP.LName = ""
        xVB_CNMemberP.Tel = ""
        xVB_CNMemberP.Expire = Date
        xVB_CNMemberP.Status = ""
        xVB_CNMemberP.CardIDToken = ""
        W_PRCxFindMemberCardByID = True
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
        'Call SP_MSGnShowing("tms_CN005200", tCS_CNMsgErr)
        Call SP_DATxRsNothing(orsTemp)
        xVB_CNMemberP.FName = "New Member"
        xVB_CNMemberP.LName = ""
        xVB_CNMemberP.Tel = ""
        xVB_CNMemberP.Expire = Date
        xVB_CNMemberP.Status = ""
        xVB_CNMemberP.CardIDToken = ""
        W_PRCxFindMemberCardByID = True
        '*Ao 2010/12/21  CommSheet PH3-ML-SF-134 Can't connect Member database.xls
        If SP_CONbMemberServer(True, tSplash, False) Then
            Call W_PRCxFindMemberCardByID(Trim(otbMemberID.Text))
        End If
End Function

Private Sub orbMonth_GotFocus()     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
    orbMonth.SelStart = 0
    orbMonth.SelLength = Len(orbMonth.Text)
End Sub

Private Sub orbMonth_KeyDown(KeyCode As Integer, Shift As Integer)     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
    Select Case KeyCode
        Case vbKeyDelete         'delete
            orbMonth.Text = ""
    End Select
End Sub

Private Sub orbMonth_KeyPress(KeyAscii As Integer)     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            sendkeys "{Tab}"
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
            If (CLng(Val(Format(orbMonth.Text))) * 10) + CLng(Val(Format(Chr(KeyAscii)))) > 12 Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub orbYear_GotFocus()     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
    orbYear.SelStart = 0
    orbYear.SelLength = Len(orbYear.Text)
End Sub

Private Sub orbYear_KeyDown(KeyCode As Integer, Shift As Integer)     '*Mee 55-01-06 V.4.0.10 Commsheet ML-PH4-R3-017
    Select Case KeyCode
        Case vbKeyDelete         'delete
            orbYear.Text = ""
    End Select
End Sub

Private Sub orbYear_KeyPress(KeyAscii As Integer)     '*Mee 55-01-06 V.4.0.10  Commsheet ML-PH4-R3-017
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
        With OPOSMSR1                                   '*Ao 10-02-09
            If bVB_OPOSMSREnable Then
                'Enable the event.
                .DataEventEnabled = False
            End If
        End With
        With oScanner
            otbMemberID.Text = .ScanData
'            DoEvents        '*KT 53-01-14
            Call SP_CRTxLogMallCardEvent("SCANNER", "DataEvent[1]", "Member Price Register", Trim(otbMemberID.Text))
            If otbMemberID.Text <> "" Then
                '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
                '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
                If Asc(Right(otbMemberID.Text, 1)) = 13 Then
                    otbMemberID.Text = Mid(otbMemberID.Text, 1, Len(otbMemberID.Text) - 1)
                End If
                Call SP_CRTxLogMallCardEvent("SCANNER", "DataEvent[2]", "Member Price Register", Trim(otbMemberID.Text))
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
        Call SP_CRTxLogMallCardEvent("SCANNER", "DataEvent[3]", "Member Price Register", Trim(otbMemberID.Text))
        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
        Exit Sub
ErrHandle:
    Call SP_CRTxLogMallCardEvent("SCANNER", "DataEvent[4]", "Member Price Register", Trim(otbMemberID.Text))
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
    Call SP_OTBxSelText(otbMemberID)
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
End Sub

Private Sub otbMemberID_GotFocus()
'    Call SP_OTBxSelText(otbMemberID)
    If otbMemberID.Enabled And otbMemberID.Visible Then     '*AO 2011/03/29
        otbMemberID.SelStart = 0
        otbMemberID.SelLength = Len(otbMemberID.Text)
    End If
End Sub

Private Sub otbMemberID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            bW_FoundMemberCard = False
            otbMemberID.Text = ""
            olaName.Caption = ""
            xVB_CNMemberP.ID = ""
        Case vbKeyBack
            bW_FoundMemberCard = False
            xVB_CNMemberP.ID = Trim(otbMemberID.Text)
    End Select
End Sub

Private Sub otbMemberID_KeyPress(KeyAscii As Integer)
    
On Error GoTo ErrHandle
    Call UT_ChkKeyAscii(KeyAscii)
    Call SP_ChkKeyAsciiNumber(KeyAscii)
    Select Case KeyAscii
        Case 13
            Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[1]", "Member Price Register", Trim(otbMemberID.Text))
            xVB_CNMemberP.ID = Trim(otbMemberID.Text)
            If xVB_CNMemberP.ID = "" Then
                Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                bW_FoundMemberCard = False
                otbMemberID.Text = ""
                olaName.Caption = ""
                xVB_CNMemberP.ID = "" '*KT 52-01-27 PH3-ML-SB-078 Error on Find customer screen during Member Price step
                Call SP_OBJxSetFocus(otbMemberID)
                Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[2]", "Member Price Register", Trim(otbMemberID.Text))
                Exit Sub
            End If
            
            '*KT 52-11-16 PH3-ML-SF-009 Check MOD10 on card Immediately before press OK button
             Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[3]", "Member Price Register", Trim(otbMemberID.Text))
            If Not W_PRCxFindMallCardByIDMemberPrice(xVB_CNMemberP.ID) Then
                    Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                    bW_FoundMemberCard = False
                    otbMemberID.Text = ""
                    olaName.Caption = ""
                    xVB_CNMemberP.ID = "" '*KT 52-01-27 PH3-ML-SB-078 Error on Find customer screen during Member Price step
                    Call SP_OBJxSetFocus(otbMemberID)
                    Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[4]", "Member Price Register", Trim(otbMemberID.Text))
                    Exit Sub
            Else
                    '*KT 52-09-30 PH2-ML-SF-149 หมายเลขบัตร Mall Card ไม่ตรงกัน ระหว่าง M-Price กับ รูด Mall Card (Mall8)
                    Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[5]", "Member Price Register", Trim(otbMemberID.Text))
                    If tW_IsChkDigitMallCard = "1" Then
                        If Not SP_PRCbIsChkDigitMallCard(xVB_CNMemberP.ID) Then
                            Call SP_MSGnShowing("tms_CN005215", nCS_Error)
                            bW_FoundMemberCard = False
                            otbMemberID.Text = ""
                            olaName.Caption = ""
                            xVB_CNMemberP.ID = "" '*KT 52-01-27 PH3-ML-SB-078 Error on Find customer screen during Member Price step
                            Call SP_OBJxSetFocus(otbMemberID)
                            Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[6]", "Member Price Register", Trim(otbMemberID.Text))
                            Exit Sub
                        End If
                    End If
            End If
            
            Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[7]", "Member Price Register", Trim(otbMemberID.Text))
            If Not W_PRCxFindMemberCardByID(xVB_CNMemberP.ID) Then
                '*KT 52-05-26 PH 1.8 ML-SF-076 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ไม่ได้ จะขาย Member Price ไม่ได้
                'Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                bW_FoundMemberCard = False
                otbMemberID.Text = ""
                olaName.Caption = ""
                xVB_CNMemberP.ID = "" '*KT 52-01-27 PH3-ML-SB-078 Error on Find customer screen during Member Price step
                Call SP_OBJxSetFocus(otbMemberID)
                Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[8]", "Member Price Register", Trim(otbMemberID.Text))
                Exit Sub
            Else
                olaName.Caption = xVB_CNMemberP.FName & " " & xVB_CNMemberP.LName
                bW_FoundMemberCard = True
                Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[9]", "Member Price Register", Trim(otbMemberID.Text))
            End If
           
'            Call SP_OBJxSetFocus(ocmOk)     '*Mee 55-01-07
            Call SP_OBJxSetFocus(orbMonth)      '*Mee 55-01-07
'            DoEvents            '*KT 53-01-14
            Exit Sub
        Case Else
            'none
            bW_FoundMemberCard = False
    End Select
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    bW_FoundMemberCard = False
    otbMemberID.Text = ""
    olaName.Caption = ""
    xVB_CNMemberP.ID = ""
    Call SP_OBJxSetFocus(otbMemberID)
    Call SP_CRTxLogMallCardEvent("Keyboard", "KeyPress[10]", "Member Price Register", Trim(otbMemberID.Text))
End Sub

Private Sub otbMemberID_LostFocus()
On Error GoTo ErrHandle
    If Not bW_FoundMemberCard And xVB_CNMemberP.ID <> "" Then '*KT 52-01-27 PH3-ML-SB-078 Error on Find customer screen during Member Price step
        xVB_CNMemberP.ID = Trim(otbMemberID.Text)
        '*KT 52-11-16 PH3-ML-SF-009 Check MOD10 on card Immediately before press OK button
        If Not W_PRCxFindMallCardByIDMemberPrice(xVB_CNMemberP.ID) Then
                Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                bW_FoundMemberCard = False
                otbMemberID.Text = ""
                olaName.Caption = ""
                orbMonth.Text = ""
                orbYear.Text = ""
                Call SP_OBJxSetFocus(otbMemberID)
                Exit Sub
        Else
                '*KT 52-09-30 PH2-ML-SF-149 หมายเลขบัตร Mall Card ไม่ตรงกัน ระหว่าง M-Price กับ รูด Mall Card (Mall8)
                If tW_IsChkDigitMallCard = "1" Then
                    If Not SP_PRCbIsChkDigitMallCard(xVB_CNMemberP.ID) Then
                        Call SP_MSGnShowing("tms_CN005215", nCS_Error)
                        bW_FoundMemberCard = False
                        otbMemberID.Text = ""
                        olaName.Caption = ""
                        orbMonth.Text = ""
                        orbYear.Text = ""
                        Call SP_OBJxSetFocus(otbMemberID)
                        Exit Sub
                    End If
                End If
        End If
        If Not W_PRCxFindMemberCardByID(xVB_CNMemberP.ID) Then
            '*KT 52-05-26 PH 1.8 ML-SF-076 ในกรณีที่ POS ไม่สามารถติดต่อ Member Server ไม่ได้ จะขาย Member Price ไม่ได้
            'Call SP_MSGnShowing("tms_CN005200", nCS_Error)
            bW_FoundMemberCard = False
            otbMemberID.Text = ""
            olaName.Caption = ""
            orbMonth.Text = ""
            orbYear.Text = ""
            Call SP_OBJxSetFocus(otbMemberID)
            Exit Sub
        Else
            olaName.Caption = xVB_CNMemberP.FName & " " & xVB_CNMemberP.LName
'            DoEvents            '*KT 53-01-14
            bW_FoundMemberCard = True
        End If
    End If
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    bW_FoundMemberCard = False
    otbMemberID.Text = ""
    olaName.Caption = ""
    orbMonth.Text = ""
    orbYear.Text = ""
    Call SP_OBJxSetFocus(otbMemberID)
End Sub

Private Function W_PRCxFindMallCardByIDMemberPrice(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตร Mall Card อยู่ในช่วงใด
'*KT 52-02-12 New Requirment
'1.  เพิ่มหน้าจอ เพื่อรับบัตร Member กรณีที่ลูกค้าต้องการได้ราคา Member Price ทั้งนี้ ขึ้นอยู่กับ Card Type หาก Member Price ไม่ได้ระบุค่าจะเอาราคาตามปกติของสินค้าขณะนั้น มาแสดง
'*KT 52-09-30 PH2-ML-SF-149 หมายเลขบัตร Mall Card ไม่ตรงกัน ระหว่าง M-Price กับ รูด Mall Card (Mall8) เพิ่ม FTMcdIsChkDigit สำหรับตรวจสอบ MOD 10
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tMcdIsJIS As String
On Error GoTo ErrHandle

    W_PRCxFindMallCardByIDMemberPrice = False
    tSql = "STP_PRCxFindMallCardByIDEx @ptAccNum= '" & ptCusID & "'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        '---------------------------------------------------------------------------------
        tMcdIsJIS = SP_FEDtChkString(orsTemp, "FTMcdIsJIS")         '*Ao 2013/05/16    TC Card
        If nW_TrackNo = 2 Then
            If UCase(tMcdIsJIS) = UCase("Y") Then
                Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                Exit Function
            End If
        Else
            If nW_TrackNo = 4 Then
                If UCase(tMcdIsJIS) = UCase("N") Then
                    Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                    Exit Function
                End If
            End If
        End If
        '---------------------------------------------------------------------------------
        xVB_CNMemberP.MallCardType = SP_FEDtChkString(orsTemp, "FTMcdCardType")
        xVB_CNMemberP.MallCardExpire = SP_FEDdChkDate(orsTemp, "FDMcdEndDate")
        xVB_CNMemberP.MallCardMemberPrice = SP_FEDtChkString(orsTemp, "FTMcdAlwPrice")
        xVB_CNMemberP.MallCardMemberPoint = SP_FEDtChkString(orsTemp, "FTMcdAlwPoint")
        W_PRCxFindMallCardByIDMemberPrice = True
        tW_IsChkDigitMallCard = SP_FEDtChkString(orsTemp, "FTMcdIsChkDigit")
    Else
        xVB_CNMemberP.MallCardType = ""
        xVB_CNMemberP.MallCardExpire = Date
        xVB_CNMemberP.MallCardMemberPrice = "N"
        xVB_CNMemberP.MallCardMemberPoint = "N"
        W_PRCxFindMallCardByIDMemberPrice = False
        tW_IsChkDigitMallCard = "0"
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    xVB_CNMemberP.MallCardType = ""
    xVB_CNMemberP.MallCardExpire = Date
    xVB_CNMemberP.MallCardMemberPrice = "N"
    xVB_CNMemberP.MallCardMemberPoint = "N"
    W_PRCxFindMallCardByIDMemberPrice = False
    tW_IsChkDigitMallCard = "0"
End Function

