VERSION 5.00
Object = "{D6810AA0-45D7-11CF-88F3-02608CA3BE43}#1.0#0"; "Msr.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNPreDayRegNonCard 
   Caption         =   "Non Card;klg_Title"
   ClientHeight    =   4065
   ClientLeft      =   4515
   ClientTop       =   3075
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Icon            =   "wCNPreDayRegNonCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7155
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
      TabIndex        =   1
      TabStop         =   0   'False
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
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "2;"
      Top             =   3360
      Width           =   1455
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3480
      Index           =   0
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   7140
      _Version        =   720897
      _ExtentX        =   12594
      _ExtentY        =   6138
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
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Non Card:;klg_ID"
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
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   600
         Width           =   2205
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1110
      Index           =   2
      Left            =   0
      TabIndex        =   5
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
   End
End
Attribute VB_Name = "wCNPreDayRegNonCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bW_FoundMemberCard  As Boolean
Private bW_KeyEnter As Boolean
Private tW_IsChkDigitMallCard As String


Private Sub Form_Activate()
On Error GoTo ErrHandle
    Call SP_OBJxSetFocus(otbMemberID)
    Exit Sub
ErrHandle:
    
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandle
    bW_FoundMemberCard = False
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

Private Sub ocmOK_Click()
    Dim dDateExp As Date
On Error GoTo ErrHandle
        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = False
        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = False
        Select Case nVB_RegPreType
            Case 3 'Non Card
                tVB_CNPreDayIDReg = Trim(otbMemberID.Text)
                If bW_FoundMemberCard Then
                   If tVB_CNPreDayIDReg = "" Then
'                        Call SP_MSGnShowing("tms_CN005220", nCS_Error)
                        Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                        bW_FoundMemberCard = False
                        otbMemberID.Text = ""
                        Call SP_OTBxSelText(otbMemberID)
                        bW_KeyEnter = False
                        If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                        Exit Sub
                    End If
                Else
                    If Not W_PRCxFindNonCardRangByID(tVB_CNPreDayIDReg) Then
'                            Call SP_MSGnShowing("tms_CN005220", nCS_Error)
                            Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                            bW_FoundMemberCard = False
                            otbMemberID.Text = ""
                            Call SP_OTBxSelText(otbMemberID)
                            bW_KeyEnter = False
                            If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                            Exit Sub
                    End If
                    If tW_IsChkDigitMallCard = "1" Then
                        If Not SP_PRCbIsChkDigitMallCard(Trim(tVB_CNPreDayIDReg)) Then
                            Call SP_MSGnShowing("tms_CN005216", nCS_Error)
                            bW_FoundMemberCard = False
                            otbMemberID.Text = ""
                            Call SP_OTBxSelText(otbMemberID)
                            bW_KeyEnter = False
                            If bVB_OPOSMSREnable Then OPOSMSR1.DataEventEnabled = True
                            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                            Exit Sub
                        End If
                    End If
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

            Case 3 'Non Card
                    tTranc2 = .Track2Data
                    If tTranc2 <> "" Then
                         If Len(Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))) <= 20 Then
                            otbMemberID.Text = Trim(Mid(tTranc2, 1, InStr(1, tTranc2, "=") - 1))
                           .DataEventEnabled = True
                        Else
'                            Call SP_MSGnShowing("tms_CN005220", nCS_Error)
                            Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                            otbMemberID.Text = ""
                            Call SP_OTBxSelText(otbMemberID)
                            If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                            Exit Sub
                        End If
                    Else
'                        Call SP_MSGnShowing("tms_CN005220", nCS_Error)
                        Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                        otbMemberID.Text = ""
                        Call SP_OTBxSelText(otbMemberID)
                        Exit Sub
                    End If
                     tVB_CNPreDayIDReg = Trim(otbMemberID.Text)

                    If Not W_PRCxFindNonCardRangByID(tVB_CNPreDayIDReg) Then
'                        Call SP_MSGnShowing("tms_CN005220", nCS_Error)
                        Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                        bW_FoundMemberCard = False
                        otbMemberID.Text = ""
                        Call SP_OTBxSelText(otbMemberID)
                        If bVB_OPOSScannerEnable Then oScanner.DataEventEnabled = True
                        Exit Sub
                    Else
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

Private Sub otbMemberID_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbMemberID.Text = ""
        Case 13                              'Enter
            Select Case nVB_RegPreType
                Case 3 'Non Card
                        tVB_CNPreDayIDReg = Trim(otbMemberID.Text)
                        If tVB_CNPreDayIDReg = "" Then
'                            Call SP_MSGnShowing("tms_CN005220", nCS_Error)
                            Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                            bW_FoundMemberCard = False
                            Call SP_OTBxSelText(otbMemberID)
                            Exit Sub
                        End If
                        If Not W_PRCxFindNonCardRangByID(tVB_CNPreDayIDReg) Then
'                            Call SP_MSGnShowing("tms_CN005220", nCS_Error)
                            Call SP_MSGnShowing("tms_CN005222", nCS_Error)          '*Ao 2010/12/21  CommSheet ML-PRV-027
                            bW_FoundMemberCard = False
                            otbMemberID.Text = ""
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
            Call ocmOK_Click
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

Private Function W_PRCxFindNonCardRangByID(ByVal ptCusID$) As Boolean
'--------------------------------------
'Call:
'Cmt: ค้นหาหมายเลขบัตร Member Card ใน ISS3000 PreviewDay Maintennance
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle

    W_PRCxFindNonCardRangByID = False
    tSql = "STP_PRCxFindMallCardByIDPreDayNonCard @ptAccNum= '" & ptCusID & "',@dDate = '" & Format(SP_GETdChkDatePreDay, "YYYY/MM/DD") & "'"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        W_PRCxFindNonCardRangByID = True
        aW_CardPreDay(1).tStaDeptCode = SP_FEDtChkString(orsTemp, "FTMcdStaDeptCodePreDay")
        aW_CardPreDay(1).tCrdType = SP_FEDtChkString(orsTemp, "FTMcdCode")
        tW_IsChkDigitMallCard = SP_FEDtChkString(orsTemp, "FTMcdIsChkDigit")
    Else
        W_PRCxFindNonCardRangByID = False
        tW_IsChkDigitMallCard = "0"
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    W_PRCxFindNonCardRangByID = False
    tW_IsChkDigitMallCard = "0"
    Call SP_DATxRsNothing(orsTemp)
End Function






