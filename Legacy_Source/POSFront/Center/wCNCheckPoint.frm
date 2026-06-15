VERSION 5.00
Object = "{D6810AA0-45D7-11CF-88F3-02608CA3BE43}#1.0#0"; "msr.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNCheckPoint 
   Caption         =   "Check Point;klg_Title"
   ClientHeight    =   4575
   ClientLeft      =   5490
   ClientTop       =   2700
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Icon            =   "wCNCheckPoint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7155
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1110
      Index           =   2
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3720
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
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   120
         Width           =   1455
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3960
      Index           =   0
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   7140
      _Version        =   720897
      _ExtentX        =   12594
      _ExtentY        =   6985
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox otbTel 
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
         TabIndex        =   1
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID:;klg_CusID"
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
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Tag             =   "2;"
         Top             =   720
         Width           =   2205
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel:;klg_Tel"
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
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:;klg_FName"
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
         TabIndex        =   10
         Tag             =   "2;"
         Top             =   1920
         Width           =   2955
      End
      Begin VB.Label olaFName 
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
         TabIndex        =   9
         Top             =   1920
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:;klg_LName"
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
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   2520
         Width           =   2910
      End
      Begin VB.Label olaLName 
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
         TabIndex        =   7
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Point:;klg_LPoint"
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
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   3180
         Width           =   2670
      End
      Begin VB.Label olbLPoint 
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
         TabIndex        =   5
         Top             =   3120
         Width           =   4215
      End
   End
End
Attribute VB_Name = "wCNCheckPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tW_MemberID As String
'Private bW_FoundMemberCard As Boolean
'Private tW_Tel As String
'Private tW_TelMobile  As String
'Private tW_TelOffice As String
Private tW_IsChkDigitMallCard As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then SendKeys "{Tab}"
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbTel.Text = ""
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
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
    Exit Sub
ErrHandle:
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
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

Private Sub ocmCancel_Click()
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = False '*Ao 10-02-09
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
    Unload Me
End Sub

Private Sub W_FRMxSetInitial()
    olaFName.BackColor = nVB_CNClrBLock
    olaLName.BackColor = nVB_CNClrBLock
    olbLPoint.BackColor = nVB_CNClrBLock
End Sub

Private Sub OPOSMSR1_DataEvent(ByVal Status As Long)
    Dim tExpDate$
    Dim tTranc2$
    Dim orsMall As ADODB.Recordset
    Dim tSql As String
On Error GoTo ErrHandle
    With oScanner                                                           '*Ao 10-02-09
         If bVB_OPOSScannerEnable Then
                'Enable the event.
                .DataEventEnabled = False
         End If
     End With
    With OPOSMSR1
        tTranc2 = .Track2Data
        .ClearInput
        .DataEventEnabled = True
        If tTranc2 <> "" Then
             If Len(Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))) <= 20 Then
                otbMemberID.Text = Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))
            Else
                Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                otbTel.Text = ""
                olaFName.Caption = ""
                olaLName.Caption = ""
                olbLPoint.Caption = ""
                Call SP_OBJxSetFocus(otbMemberID)
                Call SP_OTBxSelText(otbMemberID)
                Call SP_DATxRsNothing(orsMall)
                If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
                Exit Sub
            End If
        Else
            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
            otbTel.Text = ""
            olaFName.Caption = ""
            olaLName.Caption = ""
            olbLPoint.Caption = ""
            Call SP_OBJxSetFocus(otbMemberID)
            Call SP_OTBxSelText(otbMemberID)
            Call SP_DATxRsNothing(orsMall)
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
            Exit Sub
        End If
        tW_MemberID = Trim(otbMemberID.Text)
        
        '*KT 52-11-16 PH3-ML-SF-009 Check MOD10 on card Immediately before press OK button
        tSql = "STP_PRCxFindMallCardByID @ptAccNum= '" & tW_MemberID & "'"
        If SP_SQLvExecute(tSql, orsMall, oDB.ocnOnLine, , , , , True, True) = 0 Then  'read master
            tW_IsChkDigitMallCard = SP_FEDtChkString(orsMall, "FTMcdIsChkDigit")
            If tW_IsChkDigitMallCard = "1" Then
                If Not SP_PRCbIsChkDigitMallCard(Trim(tW_MemberID)) Then
                    Call SP_MSGnShowing("tms_CN005216", nCS_Error)
                    otbTel.Text = ""
                    olaFName.Caption = ""
                    olaLName.Caption = ""
                    olbLPoint.Caption = ""
                    Call SP_OBJxSetFocus(otbMemberID)
                    Call SP_OTBxSelText(otbMemberID)
                    Call SP_DATxRsNothing(orsMall)
                    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
                    Exit Sub
                End If
            End If
        Else
            MsgBox "MEMBER CARD is not found in database ISS3000."
            otbTel.Text = ""
            olaFName.Caption = ""
            olaLName.Caption = ""
            olbLPoint.Caption = ""
            Call SP_OBJxSetFocus(otbMemberID)
            Call SP_OTBxSelText(otbMemberID)
            Call SP_DATxRsNothing(orsMall)
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True              '*Ao 10-02-09
            Exit Sub
        End If
        
        If Not W_PRCxFindMemberCardByID(tW_MemberID) Then
            'bW_FoundMemberCard = False
            otbTel.Text = ""
            olaFName.Caption = ""
            olaLName.Caption = ""
            olbLPoint.Caption = ""
            Call SP_OBJxSetFocus(otbMemberID)
            Call SP_OTBxSelText(otbMemberID)
            Call SP_DATxRsNothing(orsMall)
            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True              '*Ao 10-02-09
            Exit Sub
        Else
            'bW_FoundMemberCard = True
        End If
    End With
    Call SP_OBJxSetFocus(otbMemberID)
    Call SP_DATxRsNothing(orsMall)
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True          '*Ao 10-02-09
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    otbTel.Text = ""
    olaFName.Caption = ""
    olaLName.Caption = ""
    olbLPoint.Caption = ""
    Call SP_OBJxSetFocus(otbMemberID)
    Call SP_DATxRsNothing(orsMall)
    If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True      '*Ao 10-02-09
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
'        DoEvents        '*KT 53-01-14
        If otbMemberID.Text <> "" Then
            '*KT 52-05-06 PH 1.8 ML-SF-056 TP 2000 ไม่สามารถใช้ Scanner  Scan สินค้าได้ โดยเมื่อ Scan สินค้าแล้วมีรหัสสินค้าปรากฏขึ้นมาถูกต้องแต่ จะขึ้นข้อความไม่พบสินค้านั้น กรณีที่ OPOS บางตัวมีการอ่านค่า Enter ต่อท้าย ต้องตัดออก
            '*KT 52-05-15 PH 1.8 ML-SF-061 Price Check ไม่สามารถใช้ Scanner ได้ จะขึ้นข้อความไม่พบรหัสสินค้า
            If Asc(Right(otbMemberID.Text, 1)) = 13 Then
                otbMemberID.Text = Mid(otbMemberID.Text, 1, Len(otbMemberID.Text) - 1)
            End If
            Call otbMemberID_KeyPress(13)
        End If
        'ทำการ Open Scanner อีกครั้ง ไม่นั้นครั้งต่อไปจะ Scan สินค้าไม่ได้ครับ
        '*KT 52-01-08 ไม่ต้อง Open Scanner อีกครั้ง ยกเว้นกรณี Form Active
'        .Open tVB_Scanner
'        .Claim 0 'TeamPos2000
'        'Enable the device.
'        .DeviceEnabled = True
'        'Enable the event.
'        .DataEventEnabled = True
'        .AutoDisable = True
'        .ClearInput
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
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True      '*Ao 10-02-09
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
    If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True          '*Ao 10-02-09
End Sub

Private Sub otbMemberID_LostFocus()
    otbMemberID.Text = SP_TXTtTrimRightLeft(otbMemberID.Text)
End Sub

Private Sub otbTel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbMemberID.Text = ""
            otbTel.Text = ""
            olaFName.Caption = ""
            olaLName.Caption = ""
            olbLPoint.Caption = ""
        Case 13
            SendKeys "{Tab}"
    End Select
End Sub

Private Sub otbTel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            If Trim(otbTel.Text) <> "" Then Call W_PRCxFindCusByTel(Trim(otbTel.Text))
            If otbMemberID.Text = "" Then Call SP_OBJxSetFocus(otbTel)
    End Select
End Sub

Private Sub W_PRCxFindCusByTel(ByVal ptTel As String)
'---------------------------------------------------
'Call:
'Cmt: ต้องการค้นหาเลขที่ Mall Card จากหมายเลขโทรศัพท์ แล้วแสดงข้อมูล
'----------------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    Call SP_STAxMouse(Me, True)
    If ocnMember Is Nothing Then
        Call SP_MSGnShowing("tms_CN005134", tCS_CNMsgErr)
        otbMemberID.Text = ""
        olaFName.Caption = ""
        olaLName.Caption = ""
        olbLPoint.Caption = ""
         Call SP_DATxRsNothing(orsTemp)
         Call SP_STAxMouse(Me, False)
        Exit Sub
    End If
'    tSql = "SELECT Top 1000  FTMcdName,FTMcdSurName,FTMcdCode,FDMcdExpDate,FTMcdSta,FCMcdTotalPoint"
'    tSql = tSql & vbCrLf & "FROM TCNMMallCard"
'    tSql = tSql & vbCrLf & "WHERE (FTMcdTel = '" & ptTel & "')"
'    tSql = tSql & vbCrLf & "OR (FTMcdMobile = '" & ptTel & "')"
'    tSql = tSql & vbCrLf & "OR (FTMcdOfficePhone = '" & ptTel & "')"       '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
    tSql = "SELECT Top 1000  *"
    tSql = tSql & vbCrLf & "FROM TCNMMallCard"
    tSql = tSql & vbCrLf & "WHERE (FTMcdTel = '" & ptTel & "')"
    tSql = tSql & vbCrLf & "OR (FTMcdMobile = '" & ptTel & "')"
    tSql = tSql & vbCrLf & "OR (FTMcdOfficePhone = '" & ptTel & "')"
    If SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            If orsTemp.RecordCount > 1 Then
                Call SP_STAxMouse(Me, False)
                Set wCNSelectMember.orsMember = orsTemp.Clone
                wCNSelectMember.Show vbModal
'                DoEvents        '*KT 53-01-14
                If tVB_FindMember <> "" Then
                    orsTemp.Filter = "FTMcdCode = '" & tVB_FindMember & "'"
                     If Not (orsTemp.BOF And orsTemp.EOF) Then
                        otbMemberID.Text = SP_FEDtChkString(orsTemp, "FTMcdCode")
                        olaFName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName")
                        olaLName.Caption = SP_FEDtChkString(orsTemp, "FTMcdSurName")
'                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCMcdTotalPoint")      '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCBalance")
                        Call SP_OBJxSetFocus(otbTel)
                    Else
                        otbMemberID.Text = ""
                        olaFName.Caption = ""
                        olaLName.Caption = ""
                        olbLPoint.Caption = ""
                         Call SP_OBJxSetFocus(otbTel)
                         Call SP_OTBxSelText(otbTel)
                    End If
                Else
                    otbMemberID.Text = ""
                    olaFName.Caption = ""
                    olaLName.Caption = ""
                    olbLPoint.Caption = ""
                     Call SP_OBJxSetFocus(otbTel)
                     Call SP_OTBxSelText(otbTel)
                End If
            Else
                otbMemberID.Text = SP_FEDtChkString(orsTemp, "FTMcdCode")
                olaFName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName")
                olaLName.Caption = SP_FEDtChkString(orsTemp, "FTMcdSurName")
'                olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCMcdTotalPoint")      '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
                olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCBalance")
                Call SP_OBJxSetFocus(otbTel)
            End If
        Else
            otbMemberID.Text = ""
            olaFName.Caption = ""
            olaLName.Caption = ""
            olbLPoint.Caption = ""
            Call SP_MSGnShowing("tms_CN005108", tCS_CNMsgErr)
            Call SP_OBJxSetFocus(otbTel)
            Call SP_OTBxSelText(otbTel)
        End If
    Else
        otbMemberID.Text = ""
        olaFName.Caption = ""
        olaLName.Caption = ""
        olbLPoint.Caption = ""
        Call SP_MSGnShowing("tms_CN005108", tCS_CNMsgErr)
        Call SP_OBJxSetFocus(otbTel)
        Call SP_OTBxSelText(otbTel)
    End If
    Call SP_DATxRsNothing(orsTemp)
    Call SP_STAxMouse(Me, False)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    otbMemberID.Text = ""
    olaFName.Caption = ""
    olaLName.Caption = ""
    olbLPoint.Caption = ""
    Call SP_OBJxSetFocus(otbTel)
    Call SP_OTBxSelText(otbTel)
    Call SP_DATxRsNothing(orsTemp)
    Call SP_STAxMouse(Me, False)
End Sub

Private Function W_PRCxFindMemberCardByID(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตร Mall Card อยู่ในช่วงใด
'*KT 52-02-12 New Requirment
'1.  เพิ่มหน้าจอ เพื่อรับบัตร Member กรณีที่ลูกค้าต้องการถาม Point
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tTokenCardNo As String
    Dim tCardNo As String
On Error GoTo ErrHandle

    W_PRCxFindMemberCardByID = False
    If ocnMember Is Nothing Then
        Call SP_MSGnShowing("tms_CN005134", tCS_CNMsgErr)
        W_PRCxFindMemberCardByID = False
         Call SP_DATxRsNothing(orsTemp)
        Exit Function
    End If
    
'    tSql = "SELECT Top 1  FTMcdName,FTMcdSurName,FTMcdCode,FTMcdTel,FDMcdExpDate"
'    tSql = tSql & vbCrLf & ",FTMcdSta , FCMcdTotalPoint ,FTMcdTel,FTMcdMobile,FTMcdOfficePhone"
'    tSql = tSql & vbCrLf & "FROM TCNMMallCard"
'    tSql = tSql & vbCrLf & "WHERE FTMcdCode='" & ptCusID & "'"         '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
    '===================================
    '*Bmp 61-02-15 Token
    tTokenCardNo = ptCusID
    If SP_PRCbToken("CheckPoint", tTokenCardNo, tTokenCardNo) Then
        tCardNo = tTokenCardNo
    Else
        tCardNo = ptCusID
    End If
    ptCusID = tCardNo
    '===================================
    tSql = "SELECT TOP 1  *"
    tSql = tSql & vbCrLf & "FROM TCNMMallCard"
    tSql = tSql & vbCrLf & "WHERE FTMcdCode='" & ptCusID & "'"
    
    'สถานะการใช้งานของบัตร  มีค่าเป็น  1=Active/ปกติ   , 2=Expire/หมดอายุ , 3=Lose/สูญหาย, 4=Hold/ระงับ
    If SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            If CDate(Format(SP_FEDdChkDate(orsTemp, "FDMcdExpDate"), "YYYY/MM/DD")) >= CDate(Format(SP_GETdChkDatePoint, "YYYY/MM/DD")) Then '*KT 52-06-04 ตรวจสอบวันที่หมดอายุตามที่กำหนดใน parameter
                Select Case SP_FEDtChkString(orsTemp, "FTMcdSta")
                    Case "A", "1" 'Active/ปกติ
                        olaFName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName")
                        olaLName.Caption = SP_FEDtChkString(orsTemp, "FTMcdSurName")
'                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCMcdTotalPoint")  '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCBalance")
                        W_PRCxFindMemberCardByID = True
                    Case "E", "2" 'Expire/หมดอายุ
                        olaFName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName")
                        olaLName.Caption = SP_FEDtChkString(orsTemp, "FTMcdSurName")
'                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCMcdTotalPoint")  '*Ao 2013/02/12    CommSheet TK-ISS3000-017 00.00.03 Replied by JESLIN_Aoo.xls
                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCBalance")
                         Call SP_MSGnShowing("tms_CN005201", tCS_CNMsgErr)
                        W_PRCxFindMemberCardByID = True
                    Case "L", "3" 'Lose/สูญหาย
                        olaFName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName")
                        olaLName.Caption = SP_FEDtChkString(orsTemp, "FTMcdSurName")
'                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCMcdTotalPoint")
                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCBalance")
                        Call SP_MSGnShowing("tms_CN005202", tCS_CNMsgErr)
                        W_PRCxFindMemberCardByID = True
                    Case "H", "4" 'Hold/ระงับ
                        olaFName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName")
                        olaLName.Caption = SP_FEDtChkString(orsTemp, "FTMcdSurName")
'                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCMcdTotalPoint")
                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCBalance")
                        Call SP_MSGnShowing("tms_CN005203", tCS_CNMsgErr)
                        W_PRCxFindMemberCardByID = True
                    Case Else 'นอกจากนี้อนุญาตให้ผ่าน
                        olaFName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName")
                        olaLName.Caption = SP_FEDtChkString(orsTemp, "FTMcdSurName")
'                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCMcdTotalPoint")
                        olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCBalance")
                        W_PRCxFindMemberCardByID = True
                End Select
            Else
                olaFName.Caption = SP_FEDtChkString(orsTemp, "FTMcdName")
                olaLName.Caption = SP_FEDtChkString(orsTemp, "FTMcdSurName")
'                olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCMcdTotalPoint")
                olbLPoint.Caption = SP_FEDvChkNumeric(orsTemp, "FCBalance")
                Call SP_MSGnShowing("tms_CN005201", tCS_CNMsgErr)
                W_PRCxFindMemberCardByID = True
            End If
        Else 'ไม่พบใน Member server ให้ถือว่าเป็นบัตรใหม่
            olaFName.Caption = ""
            olaLName.Caption = ""
            olbLPoint.Caption = ""
            Call SP_MSGnShowing("tms_CN005200", nCS_Error)
            W_PRCxFindMemberCardByID = False
        End If
    Else
        olaFName.Caption = ""
        olaLName.Caption = ""
        olbLPoint.Caption = ""
        Call SP_MSGnShowing("tms_CN005200", nCS_Error)
        W_PRCxFindMemberCardByID = False
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
        olaFName.Caption = ""
        olaLName.Caption = ""
        olbLPoint.Caption = ""
        Call SP_MSGnShowing("tms_CN005200", nCS_Error)
        W_PRCxFindMemberCardByID = False
        Call SP_DATxRsNothing(orsTemp)
End Function

Private Sub otbMemberID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbMemberID.Text = ""
            otbTel.Text = ""
            olaFName.Caption = ""
            olaLName.Caption = ""
            olbLPoint.Caption = ""
        Case 13
            SendKeys "{Tab}"
    End Select
End Sub

Private Sub otbMemberID_KeyPress(KeyAscii As Integer)
    Dim orsMall As ADODB.Recordset
    Dim tSql As String
On Error GoTo ErrHandle
    Call UT_ChkKeyAscii(KeyAscii)
    Call SP_ChkKeyAsciiNumber(KeyAscii)
    Select Case KeyAscii
        Case 13
            tW_MemberID = Trim(otbMemberID.Text)
            If tW_MemberID = "" Then
                Call SP_MSGnShowing("tms_CN005200", nCS_Error)
                'bW_FoundMemberCard = False
                otbTel.Text = ""
                olaFName.Caption = ""
                olaLName.Caption = ""
                olbLPoint.Caption = ""
                Call SP_OBJxSetFocus(otbMemberID)
                Call SP_OTBxSelText(otbMemberID)
                Call SP_DATxRsNothing(orsMall)
                Exit Sub
            End If
            '*KT 52-11-16 PH3-ML-SF-009 Check MOD10 on card Immediately before press OK button
            tSql = "STP_PRCxFindMallCardByID @ptAccNum= '" & tW_MemberID & "'"
            If SP_SQLvExecute(tSql, orsMall, oDB.ocnOnLine, , , , , True, True) = 0 Then  'read master
                If Not (orsMall.BOF And orsMall.EOF) Then
                    tW_IsChkDigitMallCard = SP_FEDtChkString(orsMall, "FTMcdIsChkDigit")
                    If tW_IsChkDigitMallCard = "1" Then
                        If Not SP_PRCbIsChkDigitMallCard(Trim(tW_MemberID)) Then
                            Call SP_MSGnShowing("tms_CN005216", nCS_Error)
                            otbTel.Text = ""
                            olaFName.Caption = ""
                            olaLName.Caption = ""
                            olbLPoint.Caption = ""
                            Call SP_OBJxSetFocus(otbMemberID)
                            Call SP_OTBxSelText(otbMemberID)
                            Call SP_DATxRsNothing(orsMall)
                            Exit Sub
                        End If
                    End If
                Else
                    MsgBox "MEMBER CARD is not found in database ISS3000."
                    otbTel.Text = ""
                    olaFName.Caption = ""
                    olaLName.Caption = ""
                    olbLPoint.Caption = ""
                    Call SP_OBJxSetFocus(otbMemberID)
                    Call SP_OTBxSelText(otbMemberID)
                    Call SP_DATxRsNothing(orsMall)
                    Exit Sub
                End If
            Else
                MsgBox "MEMBER CARD is not found in database ISS3000."
                otbTel.Text = ""
                olaFName.Caption = ""
                olaLName.Caption = ""
                olbLPoint.Caption = ""
                Call SP_OBJxSetFocus(otbMemberID)
                Call SP_OTBxSelText(otbMemberID)
                Call SP_DATxRsNothing(orsMall)
                Exit Sub
            End If
            
            If Not W_PRCxFindMemberCardByID(tW_MemberID) Then
                'bW_FoundMemberCard = False
                otbTel.Text = ""
                olaFName.Caption = ""
                olaLName.Caption = ""
                olbLPoint.Caption = ""
                Call SP_OBJxSetFocus(otbMemberID)
                Call SP_OTBxSelText(otbMemberID)
                Call SP_DATxRsNothing(orsMall)
                Exit Sub
            Else
                'bW_FoundMemberCard = True
            End If
            Call SP_OBJxSetFocus(otbMemberID)
            Call SP_DATxRsNothing(orsMall)
'            DoEvents        '*KT 53-01-14
            Exit Sub
        Case Else
            'none
    End Select
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    otbTel.Text = ""
    olaFName.Caption = ""
    olaLName.Caption = ""
    olbLPoint.Caption = ""
    Call SP_OBJxSetFocus(otbMemberID)
    Call SP_DATxRsNothing(orsMall)
End Sub


