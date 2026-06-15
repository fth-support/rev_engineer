VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNRefSaleNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return Transaction;klg_Title"
   ClientHeight    =   2760
   ClientLeft      =   7830
   ClientTop       =   3525
   ClientWidth     =   5085
   ControlBox      =   0   'False
   Icon            =   "wCNRefSaleNo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5085
   Tag             =   "2;"
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   2055
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   5085
      _Version        =   720897
      _ExtentX        =   8969
      _ExtentY        =   3625
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbTerminalNo 
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
         MaxLength       =   5
         TabIndex        =   0
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox otbTransNo 
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
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminal No:;klg_TerminalNo"
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
         Left            =   120
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   600
         Width           =   2130
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No:;klg_TransNo"
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
         Left            =   120
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   1200
         Width           =   2220
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1545
      Width           =   5085
      _Version        =   720897
      _ExtentX        =   8969
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin SCANNERLib.Scanner oScanner 
         Left            =   240
         Top             =   240
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
         Left            =   3360
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
         Left            =   1800
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNRefSaleNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'การทำรายการคืน (Return Transaction) การบันทึกเมื่อลูกค้า ต้องการคืนสินค้า กรณีที่ได้ทำการขายจบแล้ว(จบบิลแล้ว) โดยพนักงานเก็บเงินจะต้องสามารถทำคืนรายการขายสินค้านั้นๆ
'แบ่งเป็น 2 กรณี คือ

' 1) การคืนแบบอ้างอิงรายละเอียดการขาย
'****************************************
'เป็นการคืนบางรายการ กรณียังไม่ปิดการขายประจำวัน โดยลูกค้าต้องนำบิลของรายการขายมาด้วยและพนักงานเก็บเงินต้องบันทึกหมายเลขเครื่องและหมายเลขลำดับการขายของรายการขายนั้นๆ
'ระบบต้องดึงข้อมูลของรายการขายขึ้นมาเพื่อให้พนักงานเก็บเงินเลือกรายการสินค้าที่ต้องการคืนหลังจากทำรายการจบ โดยการคืนให้ลูกค้า จะต้องพิมพ์บิลที่แสดงว่าเป็นรายการคืนพร้อมรายละเอียด
'ส่วนลดซึ่งการคืนประเภทนี้ ไม่สามารถทำรายการได้ ในกรณีรายการขายมีการให้โปรโมชั่นแบบเงื่อนไขพิเศษ

' 2) การคืนแบบไม่อ้างอิงรายละเอียดการขาย
'เป็นการคืนบางรายการ กรณีที่ข้ามวันหรือกรณีรายการขายมีการให้โปรโมชั่นแบบเงื่อนไขพิเศษโดยลูกค้าต้องนำบิลของรายการขายมาด้วยและพนักงานเก็บเงินต้องบันทึกรายการที่ต้องการคืนและบันทึกส่วนลดเอง
'หลังจากทำรายการจบโดยการคืนเงินให้ลูกค้าจะต้องพิมพ์บิลที่แสดงว่ารายการคืนพร้อมรายละเอียดการคืนและส่วนลด

Option Explicit

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)             '*Ao 53-01-14
'    If KeyCode = 13 Then SendKeys "{TAB}"
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    bVB_FormPayAct = True 'PH 2.0 ML-SF-019 ไม่ต้องการให้ทำที่ Form_Activate ของ wMain
    ' '*KT 52-01-23 Communication Sheet PH2-ML-SF-069 Return refer sale can't scroll down on item selection screen.xls
    Call SP_OBJxSetFocus(otbTerminalNo)
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

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode = 0 Then Cancel = 1
'End Sub

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
    bVB_Return = False
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Call W_PRCxShowTicket
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
    '*KT 52-06-05 PH 2.0 ML-SF-020 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
    Dim tID As String
On Error GoTo ErrHandle
    With oScanner
        'tID = .ScanData '*Bump 56-02-28 Comment บรรทัดนี้
            '*Bump 56-02-28 เพิ่ม Code 8 บรรทัด
              '***************************************
                If .DeviceName = tSNR_TOUCH_SCANNER And Not .DecodeData Then
                    .DecodeData = True
                End If
                If .DeviceName = tSNR_TOUCH_SCANNER Then
                    tID = SP_SplitBarcodeData(.ScanDataType, .ScanData, .ScanDataLabel)  '*ต้องย้ายจาก mCNSTP.SP_SplitBarcodeData  ไปที่ mCNSP.SP_SplitBarcodeData
                Else
                    tID = .ScanData
                End If
                '***************************************
'        DoEvents  '*Ao 53-01-14   DoEvents
        If tID <> "" Then
            If Asc(Right(tID, 1)) = 13 Then
                tID = Mid(tID, 1, Len(tID) - 1)
            End If
            otbTerminalNo.Text = Left(tID, 5)
            otbTransNo.Text = Right(tID, 5)
        End If
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
    If otbTerminalNo.Text <> "" And otbTransNo.Text <> "" Then
        'Call W_PRCxShowTicket
        Call SP_OBJxSetFocus(ocmOk)
    End If
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub otbTerminalNo_GotFocus()
    otbTerminalNo.SelStart = 0
    otbTerminalNo.SelLength = Len(otbTerminalNo.Text)
End Sub

Private Sub otbTerminalNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbTerminalNo.Text = ""
        Case 13
            If otbTerminalNo.Text <> "" Then
                SendKeys "{TAB}"
            End If
    End Select
End Sub

Private Sub otbTerminalNo_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
End Sub

Private Sub otbTerminalNo_LostFocus()
     otbTerminalNo.Text = SP_TXTtTrimRightLeft(otbTerminalNo.Text)
End Sub

Private Sub otbTransNo_GotFocus()
    otbTransNo.SelStart = 0
    otbTransNo.SelLength = Len(otbTransNo.Text)
End Sub

Private Sub otbTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbTransNo.Text = ""
        Case 13
            Call W_PRCxShowTicket
    End Select
End Sub

Private Sub W_PRCxShowTicket()
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tTable As String
On Error GoTo ErrHandle             '*KT 54-03-25 ป้องกันโปรแกรมหลุด
    If otbTerminalNo.Text = "" Or otbTransNo.Text = "" Then
        Call SP_MSGnShowing("tms_CN005094", nCS_Error)
        If otbTerminalNo.Text = "" Then
            Call SP_OBJxSetFocus(otbTerminalNo)
        ElseIf otbTransNo.Text = "" Then
            Call SP_OBJxSetFocus(otbTransNo)
        End If
        Exit Sub
    End If
    
    'ต้องการให้สามารถทำการคืนบิลที่ทำรายการจากเครื่องอื่นได้
    'tSql = "SELECT TPSTSalHD.FTTmnNum,TPSTSalHD.FTShdTransNo,TPSTSalHD.FTShdTransType,TPSTSalHD.FDShdTransDate  "
    tSql = "SELECT TPSTSalHD.* "
    tSql = tSql & vbCrLf & " FROM TPSTSalHD"
    tSql = tSql & vbCrLf & "  WHERE  TPSTSalHD.FTTmnNum='" & Trim(otbTerminalNo.Text) & "' AND FTShdTransNo='" & Trim(otbTransNo.Text) & "'"
    tSql = tSql & vbCrLf & " AND  TPSTSalHD.FDShdTransDate='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND  TPSTSalHD.FTShdStaDoc = '1'"
    
    '*KT 51-10-21 V 3.0.140/V 3.0.200
    '-----------------------------------------------------------------------------
    tSql = tSql & vbCrLf & " AND  ISNULL(TPSTSalHD.FTShdStaCancel ,'N') <> 'Y'"
    tSql = tSql & vbCrLf & " AND FTShdTransType IN ("
    Select Case tVB_CNMnuCurrent
        Case "172002"
            tSql = tSql & vbCrLf & "'" & Format(tEN_TrnSaleTaxInc, "00") & "'"
            tSql = tSql & vbCrLf & ")"
        Case "172102"
            tSql = tSql & vbCrLf & "'" & Format(tEN_TrnSaleManualTax, "00") & "'"
            tSql = tSql & vbCrLf & ")"
        Case Else
            tSql = tSql & vbCrLf & "'" & Format(tEN_TrnSale, "00") & "'"
            tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnRedeem, "00") & "'"
            tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnVIPSale, "00") & "'"
            tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnCreateVIP, "00") & "'"
            tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnTaxExceptSale, "00") & "'"
            tSql = tSql & vbCrLf & ",'" & Format(tEN_TrnDutyFreeSale, "00") & "'"
            tSql = tSql & vbCrLf & ")"
    End Select

    '-----------------------------------------------------------------------------
    
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True) 'PH 2.0.0 ติดต่อฐานข้อมูล server
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        '*KT 66-08-17 TTK-ISS3000-839
         If SP_FEDtChkString(orsTemp, "FTShdStaCancel") <> "Y" Then
             If SP_FEDvChkNumeric(orsTemp, "FNShdStaRef") <> 2 Then 'สถานะ อ้างอิง 0:ไม่เคยอ้างอิง, 1:อ้างอิงบางส่วน, 2:อ้างอิงหมดแล้ว
                tVB_TerminalNoRet = SP_FEDtChkString(orsTemp, "FTTmnNum")
                tVB_TransNoRet = SP_FEDtChkString(orsTemp, "FTShdTransNo")
                tVB_TransTypeRet = SP_FEDtChkString(orsTemp, "FTShdTransType")
                dVB_TransDateRet = SP_FEDdChkDate(orsTemp, "FDShdTransDate")
                tVB_MemPointID = SP_FEDtChkString(orsTemp, "FTCstCardCode")
                tVB_MallCardIDRef = SP_FEDtChkString(orsTemp, "FTShdCshOrCrd") '*Tao 61-08-15 MallCardID จาก Manual ให้ใส่ *
                tVB_MemExp = SP_FEDtChkString(orsTemp, "FTCstCardType") '*Tao 61-08-15 Expire Date
                tTable = tVB_CNTblHD
                Call wFunctionKB.W_DupTblBtwSleUpdHoldHD(orsTemp, tTable)
                bVB_Return = True
            Else
                tVB_TerminalNoRet = ""
                tVB_TransNoRet = ""
                tVB_TransTypeRet = ""
                dVB_TransDateRet = Date
                
                tVB_MallCardIDRef = ""
                tVB_MemExp = ""
            
                tVB_MemPointID = ""
                bVB_CNMemberPrice = False
                xVB_CNMemberP.ID = ""
    
                Call SP_MSGnShowing("This bill has been full return full.;This bill has been full return.", nCS_Warning)
                Call SP_OBJxSetFocus(otbTerminalNo)
                Exit Sub
            End If
        Else
            tVB_TerminalNoRet = ""
            tVB_TransNoRet = ""
            tVB_TransTypeRet = ""
            dVB_TransDateRet = Date
            
            tVB_MallCardIDRef = ""
            tVB_MemExp = ""
        
            tVB_MemPointID = ""
            bVB_CNMemberPrice = False
            xVB_CNMemberP.ID = ""
                
            Call SP_MSGnShowing("This bill has been IPV.;This bill has been IPV.", nCS_Warning)
            Call SP_OBJxSetFocus(otbTerminalNo)
            Exit Sub
        End If
    Else
        tVB_TerminalNoRet = ""
        tVB_TransNoRet = ""
        tVB_TransTypeRet = ""
        dVB_TransDateRet = Date
        
        tVB_MallCardIDRef = ""
        tVB_MemExp = ""
    
        tVB_MemPointID = ""
        bVB_CNMemberPrice = False
        xVB_CNMemberP.ID = ""
    
        Call SP_MSGnShowing("tms_CN005094", nCS_Error)
        Call SP_OBJxSetFocus(otbTerminalNo)
        Exit Sub
    End If
    Call SP_DATxRsNothing(orsTemp)
    Unload Me
    Exit Sub
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    Unload Me
End Sub


Private Sub otbTransNo_LostFocus()
   otbTransNo.Text = SP_TXTtTrimRightLeft(otbTransNo.Text)
End Sub
