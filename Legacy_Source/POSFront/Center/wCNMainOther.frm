VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNMainOther 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OTHER RECEIVE;klg_Title"
   ClientHeight    =   3960
   ClientLeft      =   4365
   ClientTop       =   7620
   ClientWidth     =   6705
   ControlBox      =   0   'False
   Icon            =   "wCNMainOther.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6705
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3240
      Index           =   0
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   -240
      Width           =   6705
      _Version        =   720897
      _ExtentX        =   11827
      _ExtentY        =   5715
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbPrice 
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
         Left            =   1920
         MaxLength       =   11
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "10"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox otbOthID 
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
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "0123456789"
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox otbQty 
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
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "10"
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton ocmBrw 
         Caption         =   "..."
         Height          =   495
         Left            =   5805
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item ID:;klg_ItemID"
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
         TabIndex        =   14
         Tag             =   "2;"
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label olaItemName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าโทรศัพท์"
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
         Left            =   1920
         TabIndex        =   13
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name:;klg_ItemName"
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
         TabIndex        =   12
         Tag             =   "2;"
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:;klg_Amt"
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
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price:;klg_Price"
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
         Left            =   240
         TabIndex        =   10
         Tag             =   "2;"
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:;klg_Qty"
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
         Left            =   240
         TabIndex        =   9
         Tag             =   "2;"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label olaAmount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3,000.00"
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
         Left            =   3495
         TabIndex        =   8
         Top             =   2640
         Width           =   1200
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2760
      Width           =   6705
      _Version        =   720897
      _ExtentX        =   11827
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   360
         Top             =   480
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
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
         Left            =   5040
         TabIndex        =   5
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
         Left            =   3480
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNMainOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tW_ItemName1 As String
Private tW_ItemName2 As String
Private tW_ItemTaxCode As String
Private cW_ItemTaxRate As String
Private cW_ItemPrice As Double
Private cW_ItemNetAmt As Double
Private bW_ZeroPrice As Boolean   '*Eaw 56-05-30

Private Sub Form_Activate()
    If bW_ZeroPrice Then
        Call SP_OTBxSelText(otbPrice)
        bW_ZeroPrice = False
    Else
        bW_ZeroPrice = False
    End If
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then SendKeys "{Tab}" '*KT 53-01-19 ทำให้ขั้นตอนการ Setfocus ตอน Load ผิดพลาด ถ้ามีการใส่หน้าจอ Login มา
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    Call W_FRMxSetInitial
'    If bVB_OPOSScannerEnable Then                      '*Ao 2013/05/31   ทำให้การ Load Form ช้า 8 วิ
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
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnOtherReceive: Me.Caption = "OTHER RECEIVE"
        Case tEN_TrnOtherDisbursement: Me.Caption = "OTHER DISBURSEMENT"
    End Select
    otbOthID.Text = tUT_TextPrompt
    bW_ZeroPrice = True      '*Eaw 56-05-30 CommSheet TK-ISS3000-209-ST 00 00 01.xls
    Call otbOthID_KeyPress(13)
    
    Exit Sub
ErrHandle:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
'    With oScanner
'        If bVB_OPOSScannerEnable Then          '*Ao 2013/05/31
'            If .DeviceEnabled Then
'                .DeviceEnabled = False
'                .Release
'                .Close
'            End If
'        End If
'    End With
    Call SP_OBJxSetNothing(Me)
ErrHandle:
    
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'       Set wCNMainOther = Nothing
'End Sub

Private Sub ocmCancel_Click()
    bVB_TrnOthReceive = False
    bVB_TrnOthDisb = False
    Unload Me
End Sub

Private Sub W_FRMxSetInitial()
'----------------------------------------------
'Call:
'Cmt: set ค่าเริ่มต้น
'---------------------------------------------
    otbOthID.Text = ""
    olaItemName.Caption = ""
    
    otbQty.Text = "1"
    olaAmount.Caption = SP_FMTtNum2Shw(0, False, True)
    'Lock control
'    otbPrice.Locked = True      '*Eaw 56-05-30 CommSheet TK-ISS3000-209-ST 00 00 01.xls
'    otbPrice.BackColor = nVB_CNClrBLock   '*Eaw 56-05-30 CommSheet TK-ISS3000-209-ST 00 00 01.xls
    otbOthID.Locked = True
    otbOthID.BackColor = nVB_CNClrBLock
    otbPrice.Text = SP_FMTtNum2Shw(0, False, True)
End Sub

Private Sub ocmOK_Click()

    If (otbOthID.Text = "") Or (olaItemName.Caption = "") Then Exit Sub
    If (otbQty.Text = "") Or (Val(otbQty.Text) < 1) Then Exit Sub
    If W_INSxOthReceiveTrnDT Then
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnOtherReceive
                bVB_TrnOthReceive = True
            Case tEN_TrnOtherDisbursement
                bVB_TrnOthDisb = True
        End Select
    Else
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnOtherReceive
                bVB_TrnOthReceive = False
            Case tEN_TrnOtherDisbursement
                bVB_TrnOthDisb = False
        End Select
    End If
    Unload Me
End Sub

'Private Sub oScanner_DataEvent(ByVal Status As Long)       '*Ao 2013/05/31
'On Error GoTo ErrHandle
'    With oScanner
'        otbOthID.Text = .ScanData
''        DoEvents            '*KT 53-01-14
'        If otbOthID.Text <> "" Then
'           If Asc(Right(otbOthID.Text, 1)) = 13 Then
'                otbOthID.Text = Mid(otbOthID.Text, 1, Len(otbOthID.Text) - 1)
'            End If
'            Call otbOthID_KeyDown(13, 0)
'        End If
'        'ทำการ Open Scanner อีกครั้ง ไม่นั้นครั้งต่อไปจะ Scan สินค้าไม่ได้ครับ
'        '*KT 52-01-08 ไม่ต้อง Open Scanner อีกครั้ง ยกเว้นกรณี Form Active
''        .Open tVB_Scanner
''        .Claim 0 'TeamPos2000
''        'Enable the device.
''        .DeviceEnabled = True
''        'Enable the event.
''        .DataEventEnabled = True
''        .AutoDisable = True
''        .ClearInput
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
'End Sub

Private Sub otbOthID_GotFocus()
    otbOthID.SelStart = 0
    otbOthID.SelLength = Len(otbOthID.Text)
End Sub

Private Sub otbOthID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbOthID.Text = ""
        Case 13
            If Trim(otbOthID.Text) <> "" Then
                SendKeys "{Tab}"
           End If
    End Select
End Sub

Private Sub otbOthID_KeyPress(KeyAscii As Integer)

    Call UT_ChkKeyAscii(KeyAscii)
    If KeyAscii = 13 Then
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnOtherReceive
                If Not W_SHWbReceiveData(Trim(otbOthID.Text)) Then
                       Call SP_MSGnShowing("tms_CN005108", nCS_Warning)
                        bVB_TrnOthReceive = False
                        bVB_TrnOthDisb = False
                        Unload Me
                End If
            Case tEN_TrnOtherDisbursement
                If Not W_SHWbDisbData(Trim(otbOthID.Text)) Then
                        Call SP_MSGnShowing("tms_CN005108", nCS_Warning)
                        bVB_TrnOthReceive = False
                        bVB_TrnOthDisb = False
                        Unload Me
                End If
        End Select
'        SendKeys "{Tab}"  '*Eaw 56-05-30

    End If
End Sub

Private Sub otbOthID_LostFocus()                '*AO51-03-18    Comm.Sheet  SF-095
    If Not (Trim(otbOthID.Text) = vbNullString) Then
        Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnOtherReceive
             If Not W_SHWbReceiveData(Trim(otbOthID.Text)) Then
                    Call SP_MSGnShowing("tms_CN005108", nCS_Warning)
                    Call ocmCancel_Click
             End If
            Case tEN_TrnOtherDisbursement
            If Not W_SHWbDisbData(Trim(otbOthID.Text)) Then
                    Call SP_MSGnShowing("tms_CN005108", nCS_Warning)
                    Call ocmCancel_Click
            End If
        End Select
    End If
End Sub
    
Private Sub otbPrice_Change()
    otbPrice.Text = SP_DEFtChangeNum(Trim(otbPrice.Text))
    otbPrice.SelStart = Len(otbPrice.Text)
    cW_ItemPrice = SP_VALcText2Double(otbPrice.Text)
    cW_ItemNetAmt = SP_FMTcCurPoint(SP_VALcText2Double(otbQty.Text) * cW_ItemPrice, False, True)
End Sub

Private Sub otbPrice_GotFocus()
    If Not otbPrice.Locked Then
        otbPrice.SelStart = 0
        otbPrice.SelLength = Len(otbPrice.Text)
    Else
        SendKeys "{Tab}"
    End If
End Sub

Private Sub otbPrice_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case 13
            SendKeys "{Tab}"
'            otbPrice.SelLength = Len(otbQty)
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbQty_Change()
    cW_ItemNetAmt = SP_FMTcCurPoint(SP_VALcText2Double(otbQty.Text) * cW_ItemPrice, False, True)
    olaAmount.Caption = SP_FMTcCurPoint(cW_ItemNetAmt, False, True)
End Sub

Private Sub otbQty_GotFocus()
    otbQty.SelStart = 0
    otbQty.SelLength = Len(otbQty.Text)
End Sub

Private Sub otbQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbQty.Text = ""
    End Select
End Sub

Private Sub otbQty_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbQty.Text), ".") > 0 Then
'                KeyAscii = 0
'            End If
        Case 13
            Call ocmOK_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Function W_SHWbReceiveData(ByVal ptOrcID$) As Boolean
'------------------------------------------------
'Call:
'Cmt: Show Data Of Other Receive
'-----------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle             '*KT 54-03-25 ป้องกันโปรแกรมหลุด
    tSql = "SELECT FTOrcCode,FTOrcDesc1,FTOrcDesc2,FCOrcPrice,FTTaxCode "
    tSql = tSql & vbCrLf & ",FDOrcStaMemPoint,FTOrcStaMemPointAlw,FTOrcStaAlw,FCOrcStaMemPoint,FTEvnCode"  '*KT 54-08-04 Req.3 Phase4.0(Member Point) Member Point on Other Receive
    tSql = tSql & vbCrLf & ",FTOrcStaFixPointFlag,FNOrcMemTime"
    tSql = tSql & vbCrLf & " FROM TCNMOthReceive"
    tSql = tSql & vbCrLf & " WHERE FTOrcCode = '" & ptOrcID & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then
        If Not (orsTemp.BOF And orsTemp.EOF) Then
            tUT_SKUCode = SP_FEDtChkString(orsTemp, "FTOrcCode")
            tW_ItemName1 = SP_FEDtChkString(orsTemp, "FTOrcDesc1")
            tW_ItemName2 = SP_FEDtChkString(orsTemp, "FTOrcDesc2")
            cW_ItemPrice = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCOrcPrice"), True, False)
            tW_ItemTaxCode = SP_FEDtChkString(orsTemp, "FTTaxCode")
            'Show
            olaItemName.Caption = tW_ItemName1
            '*KT 51-05-23 Comsheet ML-SF-277
            If cW_ItemPrice > 0 Then
                otbPrice.Text = SP_FMTcCurPoint(cW_ItemPrice, False, True)
                otbPrice.Locked = True
                otbPrice.BackColor = nVB_CNClrBLock
                bW_ZeroPrice = False '*Eaw 56-05-30
            Else
                otbPrice.Text = SP_FMTcCurPoint(cW_ItemPrice, False, True)
                otbPrice.Locked = False
                otbPrice.BackColor = nVB_CNClrBUnLock
            End If
            
            '*KT 54-08-04 Req.3 Phase4.0(Member Point) Member Point on Other Receive
            If SP_FEDtChkString(orsTemp, "FTOrcStaAlw") = "Y" Then
                dVB_CNSkuDatePoint = SP_FEDdChkDate(orsTemp, "FDOrcStaMemPoint")
                If Format(dVB_CNSkuDatePoint, "YYYY/MM/DD") >= Format(SP_GETdChkDatePoint, "YYYY/MM/DD") Then
                     tVB_CNSkuPointCooAlw = SP_FEDtChkString(orsTemp, "FTOrcStaMemPointAlw")
                     tVB_CNSkuPointAlw = "Y"
                     tVB_CNSkuEventPoint = SP_FEDtChkString(orsTemp, "FTEvnCode")
                     'cVB_CNSkuPointAmt = SP_FEDvChkNumeric(orsTemp, "FCOrcStaMemPoint")
                     tVB_CNSKUStaFixPointFlag = SP_FEDtChkString(orsTemp, "FTOrcStaFixPointFlag") '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
                     tVB_CNSKUStaFixPointFlag = IIf(tVB_CNSKUStaFixPointFlag = "", "N", tVB_CNSKUStaFixPointFlag) 'ป้องกันค่าเดิมที่ไม่มีค่า
                     'ถ้า FTOrcStaFixPointFlag ="Y" ให้เก็บ cVB_CNSkuPointAmt =FCOrcStaMemPoint ถ้าไม่ใช่  cVB_CNSkuPointAmt =FNOrcMemTime
                    If tVB_CNSKUStaFixPointFlag = "Y" Then
                        cVB_CNSkuPointAmt = SP_FEDvChkNumeric(orsTemp, "FCOrcStaMemPoint")
                    Else
                        cVB_CNSkuPointAmt = SP_FEDvChkNumeric(orsTemp, "FNOrcMemTime")
                    End If
                    tVB_CNSkuPointAlwDiscountMall = "N"
                 Else
                     tVB_CNSkuPointCooAlw = "Y"
                     tVB_CNSkuPointAlw = "N"
                     tVB_CNSkuEventPoint = ""
                     tVB_CNSKUStaFixPointFlag = "N"
                     cVB_CNSkuPointAmt = 0
                     tVB_CNSkuPointAlwDiscountMall = "N"
                 End If
            Else
                 tVB_CNSkuPointCooAlw = "Y"
                 tVB_CNSkuPointAlw = "N"
                 tVB_CNSkuEventPoint = ""
                 tVB_CNSKUStaFixPointFlag = "N"
                 cVB_CNSkuPointAmt = 0
                 tVB_CNSkuPointAlwDiscountMall = "N"
            End If
            
'            otbQty.Text = "1"  '*Eaw 56-05-30
            olaAmount.Caption = SP_FMTcCurPoint(cW_ItemPrice, False, True)
            cW_ItemTaxRate = SP_SQLtFunction("", "FCTaxRate", "TCNMTaxMtn", " WHERE FTTaxCode = '" & tW_ItemTaxCode & "'", oDB.ocnOnLine, True, True) 'read master
            cW_ItemNetAmt = SP_FMTcCurPoint(SP_VALcText2Double(otbQty.Text) * cW_ItemPrice, False, True)
            W_SHWbReceiveData = True
        Else
            tUT_SKUCode = ""
            olaItemName.Caption = ""
            otbPrice.Text = SP_FMTtNum2Shw(0)
            otbQty.Text = "1"
            olaAmount.Caption = SP_FMTtNum2Shw(0)
            tW_ItemTaxCode = ""
            cW_ItemTaxRate = 1
            cW_ItemNetAmt = 0
            
            tVB_CNSkuPointCooAlw = "Y"
            tVB_CNSkuPointAlw = "N"
            tVB_CNSkuEventPoint = ""
            tVB_CNSKUStaFixPointFlag = "N"
            cVB_CNSkuPointAmt = 0
            tVB_CNSkuPointAlwDiscountMall = "N"
                  
            W_SHWbReceiveData = False
        End If
    Else
        tUT_SKUCode = ""
        olaItemName.Caption = ""
        otbPrice.Text = SP_FMTtNum2Shw(0)
        otbQty.Text = "1"
        olaAmount.Caption = SP_FMTtNum2Shw(0)
        tW_ItemTaxCode = ""
        cW_ItemTaxRate = 1
        cW_ItemNetAmt = 0
        
        tVB_CNSkuPointCooAlw = "Y"
        tVB_CNSkuPointAlw = "N"
        tVB_CNSkuEventPoint = ""
        tVB_CNSKUStaFixPointFlag = "N"
        cVB_CNSkuPointAmt = 0
        tVB_CNSkuPointAlwDiscountMall = "N"
              
        W_SHWbReceiveData = False
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    tUT_SKUCode = ""
    olaItemName.Caption = ""
    otbPrice.Text = SP_FMTtNum2Shw(0)
    otbQty.Text = "1"
    olaAmount.Caption = SP_FMTtNum2Shw(0)
    tW_ItemTaxCode = ""
    cW_ItemTaxRate = 1
    cW_ItemNetAmt = 0
    
    tVB_CNSkuPointCooAlw = "Y"
    tVB_CNSkuPointAlw = "N"
    tVB_CNSkuEventPoint = ""
    tVB_CNSKUStaFixPointFlag = "N"
    cVB_CNSkuPointAmt = 0
    tVB_CNSkuPointAlwDiscountMall = "N"
        
    W_SHWbReceiveData = False
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function W_SHWbDisbData(ByVal ptOrcID$) As Boolean
'------------------------------------------------
'Call:
'Cmt: Show Data Of Other Receive
'-----------------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    
    tSql = "SELECT FTDbsCode,FTDbsDesc1,FTDbsDesc2,FCDbsPrice,FTTaxCode "
    tSql = tSql & " FROM TCNMOthDisburse"
    tSql = tSql & " WHERE FTDbsCode = '" & ptOrcID & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then
        If Not (orsTemp.BOF Or orsTemp.EOF) Then
            tUT_SKUCode = SP_FEDtChkString(orsTemp, "FTDbsCode")
            tW_ItemName1 = SP_FEDtChkString(orsTemp, "FTDbsDesc1")
            tW_ItemName2 = SP_FEDtChkString(orsTemp, "FTDbsDesc2")
            cW_ItemPrice = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCDbsPrice"), True, False)
            tW_ItemTaxCode = SP_FEDtChkString(orsTemp, "FTTaxCode")
            'Show
            olaItemName.Caption = tW_ItemName1
            '*KT 51-05-23 Comsheet ML-SF-277
            If cW_ItemPrice > 0 Then
                otbPrice.Text = SP_FMTcCurPoint(cW_ItemPrice, False, True)
                otbPrice.Locked = True
                otbPrice.BackColor = nVB_CNClrBLock
            Else
                otbPrice.Text = SP_FMTcCurPoint(cW_ItemPrice, False, True)
                otbPrice.Locked = False
                otbPrice.BackColor = nVB_CNClrBUnLock
            End If
            
            otbQty.Text = "1"
            olaAmount.Caption = SP_FMTcCurPoint(cW_ItemPrice, False, True)
            cW_ItemTaxRate = SP_SQLtFunction("", "FCTaxRate", "TCNMTaxMtn", " WHERE FTTaxCode = '" & tW_ItemTaxCode & "'", oDB.ocnOnLine, True, True) 'read master
            cW_ItemNetAmt = SP_FMTcCurPoint(SP_VALcText2Double(otbQty.Text) * cW_ItemPrice, False, True)
            W_SHWbDisbData = True
        Else
            tUT_SKUCode = ""
            olaItemName.Caption = ""
            otbPrice.Text = SP_FMTtNum2Shw(0)
            otbQty.Text = "1"
            olaAmount.Caption = SP_FMTtNum2Shw(0)
            tW_ItemTaxCode = ""
            cW_ItemTaxRate = 1
            cW_ItemNetAmt = 0
            W_SHWbDisbData = False
        End If
    Else
        tUT_SKUCode = ""
        olaItemName.Caption = ""
        otbPrice.Text = SP_FMTtNum2Shw(0)
        otbQty.Text = "1"
        olaAmount.Caption = SP_FMTtNum2Shw(0)
        tW_ItemTaxCode = ""
        cW_ItemTaxRate = 1
        cW_ItemNetAmt = 0
        W_SHWbDisbData = False
    End If
    Call SP_DATxRsNothing(orsTemp)
End Function

Private Function W_INSxOthReceiveTrnDT() As Boolean
'------------------------------------------------
'Call:
'Cmt: *KT 50-12-25 Insert Data To DT
'-----------------------------------------------
    Dim nSeq%, cNet As Double
    Dim tSql$, tSkuBarC$, tSKUName$
    Dim tSKUName2$, tSKUNameSrvDoc$, tSKUNameSrvDoc2$
    Dim cVat As Double, cVatable As Double
    Dim cCostIn As Double, cQtyAll As Double
    Dim cQty As Double, cRegPrice As Double, cSalePrice As Double, cB4DisChg As Double
    Dim tWhr$
    Dim nEff As Long
    
On Error GoTo ErrHandle
        W_INSxOthReceiveTrnDT = False
        
        tWhr = " WHERE  FTTmnNum='" & tVB_CNTerminalNum & "' AND FTShdTransNo='" & tVB_Running & "'"
        tWhr = tWhr & vbCrLf & " AND  FTShdTransType='" & tVB_Transaction & "' AND FDShdTransDate='" & tUT_SaleDate & "'"
        
        nSeq = CLng(Val(Format(SP_SQLtFunction("MAX", "FNSdtSeqNo", tVB_CNTblDT, tWhr))) + 1)

        tSkuBarC = Trim(otbOthID.Text)  'รหัสบาร์โค้ด
        tSKUName = tW_ItemName1     'ชื่อสินค้า(Default ENG)1
        tSKUName2 = tW_ItemName2    'ชื่อสินค้า(Local)2
        tSKUNameSrvDoc = tW_ItemName1   'ชื่อสินค้า Service Doc(Default ENG) 1
        tSKUNameSrvDoc2 = tW_ItemName2     'ชื่อสินค้า Service Doc (Local)2
        
        cQty = SP_VALcText2Double(otbQty.Text)
        cRegPrice = cW_ItemPrice 'Regular Price
        cSalePrice = cW_ItemPrice 'Sale Pric
        cNet = cW_ItemNetAmt
        cB4DisChg = cSalePrice
        'cVat = IIf(tVB_SKUTaxSign = "V", SP_FMTcCurPoint(SP_CALcVatValue(IIf(tVB_CNRetVATInOrEx = "1", True, False), cNet, cVB_SKUTaxRate), True, False), 0)                            '1:vat, 2:novat (0)
        cVat = IIf(tVB_SKUTaxSign = "V", SP_FMTcCurPoint(SP_CALcVatValue(IIf(tVB_CNRetVATInOrEx = "2", True, False), cNet, cVB_SKUTaxRate), True, False), 0)                            '*tVB_CNRetVATInOrEx 1:รวมใน, 2: แยกนอก แก้ไขให้ถูกต้องตามการคำนวณใน SP_CALcVatValue '*Bump 66-03-23
        cVatable = IIf(tVB_SKUTaxSign = "V", cNet - cVat, 0)
        cQtyAll = cQty * cUT_StkFac * cUT_SKUUnitF

        Call wFunctionKB.W_FindTaxDetail(tW_ItemTaxCode)
        
        tSql = "INSERT INTO " & tVB_CNTblDT
        tSql = tSql & vbCrLf & "(FTStmCode,FTTmnNum,FTShdTransNo,FNSdtSeqNo,FTSkuCode,"
        tSql = tSql & vbCrLf & "FTSkuAbbName,FTShdTransType,FDShdTransDate,FTSdtBarCode,FTSdtStkCode,"
        tSql = tSql & vbCrLf & "FCSdtStkFac,FTTaxCode,FTDpcCode,FTSkuDepCode,FTSkuBrdCode,"
        tSql = tSql & vbCrLf & "FTPunCode,FTSdtUnitName,FCSdtRegPrice,FTSdtEventPrice,FCSdtQty,"
        tSql = tSql & vbCrLf & "FCSdtSaleTaxAmt,FCSdtReturnTaxAmt,"
        tSql = tSql & vbCrLf & "FTSdtInstoreBar,FCSdtB4DisChg,FTSdtDisChgTxt,FCSdtDis,"
        tSql = tSql & vbCrLf & "FCSdtChg,FCSdtSaleAmt,FCSdtVat,FCSdtVatable,FCSdtQtyAll,"
        tSql = tSql & vbCrLf & "FTSdtStaPdt,"
        tSql = tSql & vbCrLf & "FTSdtWeight,FTSdtConsignFlag,"
        tSql = tSql & vbCrLf & "FTSdtCancelFlag,FCSdtTax,FCSdtSalePrice,FTShdTransTime,"
        tSql = tSql & vbCrLf & "FTSkuAbbNameSnd,FTSkuAbbNameSrvDoc,FTSkuAbbNameSndSrvDoc,"
        tSql = tSql & vbCrLf & "FTTaxSign,FTTaxStaSign,"
        tSql = tSql & vbCrLf & "FCSdtSaleUnit,"
        tSql = tSql & vbCrLf & "FCSdtQtyAllDisAvi,"
        tSql = tSql & vbCrLf & "FCSdtAmtDisAvi,"
        tSql = tSql & vbCrLf & "FCSdtSalePricePmt,"
        tSql = tSql & vbCrLf & "FTSkuStaMallEmpCard ," 'Default 'Y'(ป้ายน้ำเงิน)
        tSql = tSql & vbCrLf & "FCSdtBSubPro,"
        tSql = tSql & vbCrLf & "FCSdtBSubReg,"
        tSql = tSql & vbCrLf & "FCSdtDiscSubPro,"
        tSql = tSql & vbCrLf & "FCSdtDiscSubReg,"
        tSql = tSql & vbCrLf & "FCSdtBluePrice,"
        tSql = tSql & vbCrLf & "FCSdtYellowPrice,"
        tSql = tSql & vbCrLf & "FTSdtStaRfd," '*KT 52-02-13 Amount of member price
        tSql = tSql & vbCrLf & "FCSdtMemberPrice," '*KT 52-02-13 Amount of member price
        tSql = tSql & vbCrLf & "FTSdtEventMemPrice," '*KT 52-02-18 Event ID of member price
        tSql = tSql & vbCrLf & "FCSdtB4DisChgOrg," '*KT 52-02-18 ทำการเก็บค่าเดิม FCSdtB4DisChg ก่อนได้ member card
        tSql = tSql & vbCrLf & "FCSdtPriceOveride," '*KT 52-02-13 price overide
        tSql = tSql & vbCrLf & "FTSdtMemberPointAlw," '*KT 52-06-15 Member Point By SKU Y: Allow N: Not Allow
        tSql = tSql & vbCrLf & "FCSdtMemberPoint,"  '*KT 52-06-15 Amount of member price
        tSql = tSql & vbCrLf & "FTSdtEventMemPoint," '*KT 52-06-15 Event ID Member Price
        tSql = tSql & vbCrLf & "FTSdtMemPointCooAlw," '*KT 52-06-15 Y: Allow Member Point(Cooperate[General]) N: Not Allow Member Point(Cooperate[General])
        tSql = tSql & vbCrLf & "FTSkuIndex,"
        tSql = tSql & vbCrLf & "FTSdtSpnCode,"             '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
        tSql = tSql & vbCrLf & "FTStdStaFixPointFlag," '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
        tSql = tSql & vbCrLf & "FNSdtSKUPointTime," '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
        tSql = tSql & vbCrLf & tCS_CNLastFedIns
        tSql = tSql & vbCrLf & " VALUES ("
        tSql = tSql & vbCrLf & "'" & tVB_CNStoreCode & "','" & tVB_CNTerminalNum & "','" & tVB_Running & "','" & nSeq & "','" & tUT_SKUCode & "'"
        tSql = tSql & vbCrLf & ",'" & SP_STRtSiggleCode(tSKUName) & "','" & tVB_Transaction & "','" & tUT_SaleDate & "','" & tSkuBarC & "',''"
        tSql = tSql & vbCrLf & "," & SP_FMTcCurPoint(cUT_StkFac, True, False) & ",'" & tW_ItemTaxCode & "','" & "" & "','" & "" & "','" & "" & "'"
        tSql = tSql & vbCrLf & "," & "'','" & "" & "'," & SP_FMTcCurPoint(cW_ItemPrice, True, False) & ",'" & "" & "'," & SP_FMTcCurPoint(cQty, True, False)
        tSql = tSql & vbCrLf & "," & "0,0"
        tSql = tSql & vbCrLf & "," & "''," & cB4DisChg & ",''," & 0
        tSql = tSql & vbCrLf & "," & "0" & "," & cNet & "," & cVat & "," & cVatable & "," & cQtyAll
        tSql = tSql & vbCrLf & "," & "''"
        tSql = tSql & vbCrLf & "," & "'','N'"
        tSql = tSql & vbCrLf & "," & "''," & cW_ItemTaxRate & "," & cSalePrice & ",'" & Time$ & "'"
        tSql = tSql & vbCrLf & "," & "'" & SP_STRtSiggleCode(tSKUName2) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc) & "','" & SP_STRtSiggleCode(tSKUNameSrvDoc2) & "'"
        tSql = tSql & vbCrLf & "," & "'" & tVB_SKUTaxSign & "','" & tVB_SKUTaxStaSign & "'"
        tSql = tSql & vbCrLf & "," & cSalePrice
        tSql = tSql & vbCrLf & "," & cQtyAll
        tSql = tSql & vbCrLf & "," & cNet
        tSql = tSql & vbCrLf & "," & cB4DisChg 'สำหรับราคาที่นำไปคิด deal 5,6
        tSql = tSql & vbCrLf & "," & "'Y'" 'Default 'Y'(ป้ายน้ำเงิน)
        tSql = tSql & vbCrLf & "," & 0
        tSql = tSql & vbCrLf & "," & cNet
        tSql = tSql & vbCrLf & "," & 0
        tSql = tSql & vbCrLf & "," & cNet
        tSql = tSql & vbCrLf & "," & cNet
        tSql = tSql & vbCrLf & "," & 0
        tSql = tSql & vbCrLf & "," & "'N'"  '*KT 52-02-13 Amount of member price
        tSql = tSql & vbCrLf & "," & 0  '*KT 52-02-13 Amount of member price
        tSql = tSql & vbCrLf & ",'" & "" & "'"
        tSql = tSql & vbCrLf & "," & cB4DisChg
        tSql = tSql & vbCrLf & "," & 0  '*KT 52-02-13 price overide
        tSql = tSql & vbCrLf & ",'" & tVB_CNSkuPointAlw & "'"
        tSql = tSql & vbCrLf & "," & cVB_CNSkuPointAmt
        tSql = tSql & vbCrLf & ",'" & tVB_CNSkuEventPoint & "'"
        tSql = tSql & vbCrLf & ",'" & tVB_CNSkuPointCooAlw & "'"
        tSql = tSql & vbCrLf & ",'" & nSeq & "'"
        tSql = tSql & vbCrLf & ",'" & tVB_SalemanID & "'"             '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
        tSql = tSql & vbCrLf & ",'" & tVB_CNSKUStaFixPointFlag & "'"             '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
        tSql = tSql & vbCrLf & "," & IIf(tVB_CNSKUStaFixPointFlag = "Y", 0, cVB_CNSkuPointAmt)          '*KT 54-08-10 V.4.0.1 Req.1 Phase4.0(Member Point) SKU Point with multiply by time
        tSql = tSql & vbCrLf & "," & UT_tLastInsSql & ")"
        If SP_SQLbExecute(tSql, , , nEff) Then
            If nEff = 1 Then
                W_INSxOthReceiveTrnDT = True
            End If
        End If
        Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

