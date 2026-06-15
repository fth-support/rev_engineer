VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNRecall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recall Transaction;klg_Title"
   ClientHeight    =   3015
   ClientLeft      =   7545
   ClientTop       =   6030
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Icon            =   "wCNRecall.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4815
   Tag             =   "2;"
   Begin SCANNERLib.Scanner oScanner 
      Left            =   0
      Top             =   2280
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   2535
      Index           =   0
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   -240
      Width           =   4815
      _Version        =   720897
      _ExtentX        =   8493
      _ExtentY        =   4471
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
         Index           =   1
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1080
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
         Index           =   0
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   0
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label olbSubTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   1740
         Width           =   2175
      End
      Begin VB.Label olaForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Sale  Sub Total ;klg_SaleSubTotal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Tag             =   "2;"
         Top             =   1740
         Width           =   2175
      End
      Begin VB.Label olaForm 
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
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Tag             =   "2;"
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Label olaForm 
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
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Tag             =   "2;"
         Top             =   540
         Width           =   2055
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1095
      Index           =   1
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1920
      Width           =   4815
      _Version        =   720897
      _ExtentX        =   8493
      _ExtentY        =   1931
      _StockProps     =   64
      Appearance      =   10
      Color           =   32
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
         Left            =   1560
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNRecall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub W_PRCxShowSubTotal()
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    If otbForm(0).Text = "" Or otbForm(1) = "" Then
        Call SP_MSGnShowing("tms_CN005093", nCS_Error)
        If otbForm(0).Text = "" Then
            Call SP_OBJxSetFocus(otbForm(0))
        ElseIf otbForm(1) = "" Then
            Call SP_OBJxSetFocus(otbForm(1))
        End If
        Exit Sub
    End If
    
    'Comm.Sheet ML - SF - 138
    '*KT 51-07-07 ทำการ Recall ได้เฉพาะ tEN_TrnSale,tEN_TrnTaxExceptSale,tEN_TrnDutyFreeSale
    tSql = "SELECT FCShdTotal,FTShdTransType"
    tSql = tSql & vbCrLf & " FROM  TPSTHoldHD"
    tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & otbForm(0).Text & "' AND FTShdTransNo='" & otbForm(1).Text & "'"
    tSql = tSql & vbCrLf & " AND  FDShdTransDate='" & tUT_SaleDate & "'"
    tSql = tSql & vbCrLf & " AND ((FTTmnCodeRecall = '' OR FTTmnCodeRecall IS NULL) AND (FTTransNoRecall = '' OR FTTransNoRecall IS NULL))"
    tSql = tSql & vbCrLf & " AND (FTShdTransType IN ('" & Format(tEN_TrnSuspended, "00") & "','" & Format(tEN_TrnSale, "00") & "','" & Format(tEN_TrnTaxExceptSale, "00") & "','" & Format(tEN_TrnDutyFreeSale, "00") & "'))"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True) 'PH 2.0.0 ติดต่อฐานข้อมูล server ตลอด
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        olbSubTotal.Caption = SP_FMTcCurPoint(SP_FEDvChkNumeric(orsTemp, "FCShdTotal"), False, True)
    Else
         olbSubTotal.Caption = "0.00"
    End If
    If olbSubTotal.Caption = "" Or CDbl(Val(Format(olbSubTotal.Caption))) = 0 Then
        Call SP_MSGnShowing("tms_CN005095", nCS_Error)
        Call SP_OBJxSetFocus(otbForm(0))
        bVB_Recall = False
    Else
        tVB_TransTypeRecall = SP_FEDtChkString(orsTemp, "FTShdTransType")
        Call SP_OBJxSetFocus(ocmOk)
        bVB_Recall = True
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    bVB_Recall = False
    Call SP_DATxRsNothing(orsTemp)
End Sub
 
Private Sub W_PRCxRecall()
'    bVB_Recall = True          '*AO51-03-25   ย้ายไปที่    W_PRCxShowSubTotal
    tVB_TerminalRecall = otbForm(0).Text
    tVB_TransNoRecall = otbForm(1).Text
End Sub

Private Sub Form_Activate()
On Error GoTo ErrHandle
    Call SP_OBJxSetFocus(otbForm(0))
    Exit Sub
ErrHandle:
    
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii                    '*AO51-03-20
'        Case 13: SendKeys "{Tab}"
'    End Select
''    If KeyAscii = 13 Then SendKeys "{Tab}"
'
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call SP_FRMxChildCenter(Me)
    Call SP_FRMxGetCaptionFrmDB2Tag(Me)
    Call SP_FRMxSetCapForm(Me)
    bVB_FormPayAct = True 'PH 2.0 ML-SF-019 ไม่ต้องการให้ทำที่ Form_Activate ของ wMain
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

Private Sub ocmCancel_Click()
    bVB_Recall = False
    tVB_Transaction = aVB_TransMode(2).ID
    tVB_TransactionSub = tVB_Transaction
    tVB_TransName = aVB_TransMode(2).Name
    tVB_TransNameABB = aVB_NameAbb(tEN_AbbSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSale - 1).UsrLangSec
    tVB_CNTypePrinter = aVB_TransMode(2).PrintType
    wMain.olaTransName.Caption = tVB_TransName
    Unload Me
End Sub

Private Sub ocmOK_Click()
    If olbSubTotal.Caption = "0.00" Then Call W_PRCxShowSubTotal '*Tao 58-12-15 บางครั้งพนักงานลืมกดปุ่ม Enter ให้ค้นหาบิลที่จะ recall ก่อน

    If bVB_Recall = False Then
        Call SP_MSGnShowing("tms_CN005095", nCS_Error)
        Call SP_OBJxSetFocus(otbForm(0))
        Exit Sub             '*AO51-03-25
    End If
    wMain.olaTransName.Caption = "Recall"
    If SP_MSGnShowing("tms_CN005124", tCS_CNMsgYesNoDefYes) = vbYes Then
        Call W_PRCxRecall
    Else
        bVB_Recall = False
        tVB_Transaction = aVB_TransMode(2).ID
        tVB_TransactionSub = tVB_Transaction
        tVB_TransName = aVB_TransMode(2).Name
        tVB_TransNameABB = aVB_NameAbb(tEN_AbbSale - 1).UsrLangPri & ";" & aVB_NameAbb(tEN_AbbSale - 1).UsrLangSec
        tVB_CNTypePrinter = aVB_TransMode(2).PrintType
        wMain.olaTransName.Caption = tVB_TransName
    End If
    Unload Me
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
    '*KT 52-06-05 PH 2.0 ML-SF-020 Sale Main ID และ Input Box อื่นๆ ที่สามารถใส่ Number  ใน Discount Menu , Tender Menu ให้สามารถใช้ Scanner อ่านBar Code  หมายเลข ได้
    Dim tID As String
On Error GoTo ErrHandle
    With oScanner
        'tID = .ScanData    '*Bump 56-02-28 Comment บรรทัดนี้
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
'        DoEvents '*Ao 53-01-14   DoEvents
        If tID <> "" Then
            If Asc(Right(tID, 1)) = 13 Then
                tID = Mid(tID, 1, Len(tID) - 1)
            End If
            otbForm(0).Text = Left(tID, 5)
            otbForm(1).Text = Right(tID, 5)
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
    If otbForm(0).Text <> "" And otbForm(1).Text <> "" Then
        Call otbForm_KeyDown(1, 13, 0)
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
             If Index = 1 Then
                Call W_PRCxShowSubTotal
            Else
                Call SP_OBJxSetFocus(otbForm(1))
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
               'do nothing
            Case Else
                If Not IsNumeric(Chr(KeyAscii)) Then
                    KeyAscii = 0
                End If
        End Select
    End If
End Sub

Private Sub otbForm_LostFocus(Index As Integer)
       otbForm(Index).Text = SP_TXTtTrimRightLeft(otbForm(Index).Text)
End Sub
