VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNCopySale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy Sale;klg_Title"
   ClientHeight    =   2490
   ClientLeft      =   2295
   ClientTop       =   3960
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "wCNCopySale.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4800
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1815
      Index           =   0
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   4815
      _Version        =   720897
      _ExtentX        =   8493
      _ExtentY        =   3201
      _StockProps     =   64
      Appearance      =   10
      Color           =   32
      Begin VB.TextBox otbForm 
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
         Index           =   0
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   0
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox otbForm 
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
         Index           =   1
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal No ;klg_TerminalNo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   540
         Width           =   3675
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No ;klg_TransNo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   1140
         Width           =   3645
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1290
      Width           =   4815
      _Version        =   720897
      _ExtentX        =   8493
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   240
         Top             =   480
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
         Height          =   495
         Left            =   3120
         TabIndex        =   3
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
         TabIndex        =   2
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNCopySale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Form_Activate()
'On Error GoTo ErrHandle
''    Call SP_OTBxSelText(otbForm(0))
'    If otbForm(0).Text = "" And otbForm(1).Text = "" Then
'        otbForm(0).SetFocus
'    End If
'    Exit Sub
'ErrHandle:
'End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        SendKeys "{TAB}"
'    End If
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    bVB_FormPayAct = True 'PH 2.0 ML-SF-019 ไม่ต้องการให้ทำที่ Form_Activate ของ wMain
    tVB_TerminalCpySale = ""
    tVB_TransNoCpySale = ""
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
            '*Bump 56-02-28 Touch Scanner เพิ่ม code 3 บรรทัด
            If .DeviceName = tSNR_TOUCH_SCANNER Then
                    .DecodeData = True
            End If
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

Private Sub otbForm_LostFocus(Index As Integer)
    otbForm(Index).Text = SP_TXTtTrimRightLeft(otbForm(Index).Text)
End Sub

Private Sub ocmCancel_Click()
    bVB_CpySale = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
    If W_CHKbTicket(otbForm(0).Text, otbForm(1).Text, True) Then
        bVB_CpySale = True
        tVB_TerminalCpySale = otbForm(0).Text
        tVB_TransNoCpySale = otbForm(1).Text
        wMain.olaTransName.Caption = IIf(bVB_TrainingMode, "Tr\" & tVB_TransName, tVB_TransName)
'        DoEvents            '*KT 53-01-14
        Unload Me
    Else
        Call SP_MSGnShowing("tms_CN005108", nCS_Information)
        Call SP_OBJxSetFocus(otbForm(0))
    End If
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
    '*KT 52-06-05 PH 2.0 ML-SF-020 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
    Dim tID As String
On Error GoTo ErrHandle
    With oScanner
        'tID = .ScanData    '*Bump 56-02-28 Comment บรรทัดนี้
        '*Bump 56-02-28 เพิ่ม code 8 บรรทัด
        If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
            .DecodeData = True
        End If
        If .DeviceName = tSNR_TOUCH_SCANNER Then
            tID = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel)  '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
        Else
            tID = .ScanData
        End If
'        DoEvents        '*KT 53-01-14
        If tID <> "" Then
            If Asc(Right(tID, 1)) = 13 Then
                tID = Mid(tID, 1, Len(tID) - 1)
            End If
            otbForm(0).Text = Left(tID, 5)
            otbForm(1).Text = Right(tID, 5)
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
    If otbForm(0).Text <> "" And otbForm(1).Text <> "" Then
        'Call ocmOK_Click
        Call SP_OBJxSetFocus(ocmOk)
    Else
        Call SP_OBJxSetFocus(otbForm(0))
    End If
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub otbForm_GotFocus(Index As Integer)
    otbForm(Index).SelStart = 0
    otbForm(Index).SelLength = Len(otbForm(Index).Text)
End Sub

Private Sub otbForm_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbForm(Index).Text = ""
        Case 13
            If otbForm(Index).Text <> "" Then
                SendKeys "{TAB}"
            End If
    End Select
End Sub

Private Sub otbForm_KeyPress(Index As Integer, KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    If Index = 1 Then
        Select Case KeyAscii
            Case 8, 6           'backspace, delete
                'do nothing
            Case 46             'ทศนิยม
                'do nothing
            Case 13
               Call ocmOK_Click
            Case Else
                If Not IsNumeric(Chr(KeyAscii)) Then
                    KeyAscii = 0
                End If
        End Select
    End If
End Sub

Private Function W_CHKbTicket(ByVal ptTerminalNo$, ByVal ptTransNo$, Optional ByVal pbSal As Boolean = False) As Boolean
'----------------------------------------------------
'Ret: กรณีที่พบ Ticket ที่จะทำการ Copy Sale จะส่งค่า Ture กลับ ถ้าไม่มี Ticket จะส่ง False
'Cmt: เป็น function หาว่า TerminalNo และ Transaction No ที่ Key มีในฐานข้อมูลเพื่อทำ Copy Sale ได้หรือไม่
'Req : การเรียกข้อมูลสินค้าที่จบบิลไปแล้วมาทำการขายใหม่ - เพื่อให้พนักงานเก็บเงิน สามารถเรียกข้อมูลสินค้าที่ได้จบการขายไปแล้วกอนหน้านั้นมาทำการขายใหม่ ในกรณีที่มีการผิดพลาดในการ
'           ให้ส่วนลดหรือการชำระ
'---------------------------------------------------
    Dim tSql$
    Dim tTblHD  As String
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle

    If pbSal Then
            tTblHD = "TPSTSalHD"
    Else
            tTblHD = tVB_CNTblHD
    End If
     W_CHKbTicket = False
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnTaxExceptSale
            '*KT 51-08-25 Communication Sheet ML-SF-482 Can not Copy Tax Except Sale.xls
            '-------------------------------------------------------------------------------------------------
            tSql = "SELECT FTTmnNum,FTShdTransNo,FTShdTransType "
            tSql = tSql & " FROM " & tTblHD
            tSql = tSql & "  WHERE FTTmnNum='" & ptTerminalNo & "'"
            tSql = tSql & "  AND FTShdTransNo='" & ptTransNo & "'"
            tSql = tSql & "  AND FTShdTransType IN ('" & Format(tEN_TrnTaxExceptSale, "00") & "')"
            tSql = tSql & "  AND (FTShdStaDoc = '1' OR FTShdStaDoc = '3' )"
            '--------------------------------------------------------------------------------------------------
        Case Else
            '*KT 51-08-14 Communication Sheet ML-SF-461 VIP sale can't copy transaction.xls
            '-------------------------------------------------------------------------------------------------
            tSql = "SELECT FTTmnNum,FTShdTransNo,FTShdTransType "
            tSql = tSql & " FROM " & tTblHD
            tSql = tSql & "  WHERE FTTmnNum='" & ptTerminalNo & "'"
            tSql = tSql & "  AND FTShdTransNo='" & ptTransNo & "'"
            tSql = tSql & "  AND FTShdTransType IN ('" & Format(tEN_TrnSale, "00") & "','" & Format(tEN_TrnVIPSale, "00") & "')"
            tSql = tSql & "  AND (FTShdStaDoc = '1' OR FTShdStaDoc = '3' )"
            '--------------------------------------------------------------------------------------------------
    End Select
    
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True) 'PH 2.0.0 ติดต่อฐานข้อมูล server
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        W_CHKbTicket = True
        tVB_TransTypeCpySale = SP_FEDtChkString(orsTemp, "FTShdTransType")
    Else
        Call SP_SQLvExecute(tSql, orsTemp) 'PH 2.0.0 ติดต่อฐานข้อมูล Local
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            W_CHKbTicket = True
            tVB_TransTypeCpySale = SP_FEDtChkString(orsTemp, "FTShdTransType")
        End If
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    W_CHKbTicket = False
End Function


