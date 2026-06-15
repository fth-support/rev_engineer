VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Object = "{9D6F2CC0-07CE-11CF-A800-0020AF9EBB73}#1.0#0"; "Scanner.ocx"
Begin VB.Form wCNDepositRedeem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposit Redeem;klg_Title"
   ClientHeight    =   2535
   ClientLeft      =   2895
   ClientTop       =   6660
   ClientWidth     =   6570
   ControlBox      =   0   'False
   Icon            =   "wCNDepositRedeem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6570
   Tag             =   "2;"
   Begin SCANNERLib.Scanner oScanner 
      Left            =   240
      Top             =   1920
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1815
      Index           =   2
      Left            =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -240
      Width           =   6525
      _Version        =   720897
      _ExtentX        =   11509
      _ExtentY        =   3201
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbDepositNo 
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
         MaxLength       =   20
         TabIndex        =   0
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Layaway Number:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   795
         Width           =   2280
      End
   End
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
      Left            =   6240
      MaxLength       =   5
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   6240
      MaxLength       =   5
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame ofaForm 
      Height          =   1095
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox otbDate 
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
         Index           =   2
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox otbDate 
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
         Left            =   960
         MaxLength       =   2
         TabIndex        =   7
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox otbDate 
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
         Left            =   240
         MaxLength       =   2
         TabIndex        =   6
         Top             =   360
         Width           =   600
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DDMMYYYY"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   480
         Width           =   1500
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   30
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   6525
      _Version        =   720897
      _ExtentX        =   11509
      _ExtentY        =   2117
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
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
         Left            =   4920
         TabIndex        =   2
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
         Left            =   3360
         TabIndex        =   1
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSComCtl2.DTPicker odtDate 
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   155910147
      CurrentDate     =   39528
   End
   Begin VB.Label olaForm 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(MM/DD/YYYY)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   4
      Left            =   6840
      TabIndex        =   16
      Top             =   1920
      Width           =   1860
   End
   Begin VB.Label olaForm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date:;klg_TransDate"
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
      Left            =   3960
      TabIndex        =   15
      Tag             =   "2;"
      Top             =   1440
      Width           =   2220
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
      Left            =   3960
      TabIndex        =   14
      Tag             =   "2;"
      Top             =   120
      Width           =   2250
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
      Left            =   3960
      TabIndex        =   13
      Tag             =   "2;"
      Top             =   720
      Width           =   2220
   End
End
Attribute VB_Name = "wCNDepositRedeem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function W_CHKbRecordCount(ByVal ptFTTmnNum As String, ByVal ptFTShdTransNo As String, ByVal ptFDShdTransDate As String) As Boolean
'-------------------------------------------------------------
'Call:
'Cmt: '*Ao 2014/03/04           Check  Record Count
'--------------------------------------------------------------
Dim tSql As String
Dim bChk As Boolean
Dim tRecord As String, tRecordDT As String, tRecordRC As String, tRecordCD As String, tCheck As String
Dim orsTemp As ADODB.Recordset

On Error GoTo ErrHandle
    bChk = False
    
    tRecord = ""
    
    tSql = "SELECT FTTmnNum,FTShdTransNo,FDShdTransDate,FTCstCardType3"
    tSql = tSql & vbCrLf & "FROM TPSTDepositHD"
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptFTTmnNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptFTShdTransNo & "'"
    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & ptFDShdTransDate & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then  'PH 2.0.0 'ต้องอ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
        If (orsTemp.BOF And orsTemp.EOF) Then
            W_CHKbRecordCount = True
            Exit Function
        Else
            tRecord = SP_FEDtChkString(orsTemp, "FTCstCardType3")
            bChk = False
        End If
    Else
        W_CHKbRecordCount = True
        Exit Function
    End If
    If tRecord = "" Then
        W_CHKbRecordCount = False
        Exit Function
    End If
    
    If bChk = False Then
        Call SP_DATxRsNothing(orsTemp)
        tSql = "SELECT COUNT(FTTmnNum) AS FTRecordDT"
        tSql = tSql & vbCrLf & "FROM TPSTDepositDT"
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptFTTmnNum & "'"
        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptFTShdTransNo & "'"
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & ptFDShdTransDate & "'"
        If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then  'PH 2.0.0 'ต้องอ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
            If (orsTemp.BOF And orsTemp.EOF) Then
                tRecordDT = "0"
            Else
                tRecordDT = SP_FEDtChkString(orsTemp, "FTRecordDT")
            End If
        Else
            tRecordDT = "0"
        End If
        Call SP_DATxRsNothing(orsTemp)
        tSql = "SELECT COUNT(FTTmnNum) AS FTRecordRC"
        tSql = tSql & vbCrLf & "FROM TPSTDepositRC"
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptFTTmnNum & "'"
        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptFTShdTransNo & "'"
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & ptFDShdTransDate & "'"
        If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then  'PH 2.0.0 'ต้องอ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
            If (orsTemp.BOF And orsTemp.EOF) Then
                tRecordRC = "0"
            Else
                tRecordRC = SP_FEDtChkString(orsTemp, "FTRecordRC")
            End If
        Else
            tRecordRC = "0"
        End If
        Call SP_DATxRsNothing(orsTemp)
        tSql = "SELECT COUNT(FTTmnNum) AS FTRecordCD"
        tSql = tSql & vbCrLf & "FROM TPSTDepositCD"
        tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptFTTmnNum & "'"
        tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptFTShdTransNo & "'"
        tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & ptFDShdTransDate & "'"
        If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then  'PH 2.0.0 'ต้องอ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
            If (orsTemp.BOF And orsTemp.EOF) Then
                tRecordCD = "0"
            Else
                tRecordCD = SP_FEDtChkString(orsTemp, "FTRecordCD")
            End If
        Else
            tRecordCD = "0"
        End If
        
        tCheck = UT_tToken(tRecord, ",", 1, True) & UT_tToken(tRecord, ",", 2, True) & UT_tToken(tRecord, ",", 3, True) & UT_tToken(tRecord, ",", 4, True)
        If tCheck = "1" & tRecordDT & tRecordRC & tRecordCD Then
            bChk = False
        Else
            bChk = True
        End If
        
    End If
    W_CHKbRecordCount = bChk
Exit Function
ErrHandle:
    W_CHKbRecordCount = True
End Function

Private Function W_CHKbCheckDT(ByVal ptFTTmnNum As String, ByVal ptFTShdTransNo As String, ByVal ptFDShdTransDate As String) As Boolean
'-------------------------------------------------------------
'Call:
'Cmt: '*Ao 2014/03/04           Check  TPSTDepositDT
'--------------------------------------------------------------
Dim tSql As String
Dim bChk As Boolean
Dim orsTemp As ADODB.Recordset

On Error GoTo ErrHandle
    bChk = False
    tSql = "SELECT FTTmnNum,FTShdTransNo,FDShdTransDate FROM TPSTDepositDT"
    tSql = tSql & vbCrLf & "WHERE FTTmnNum = '" & ptFTTmnNum & "'"
    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & ptFTShdTransNo & "'"
    tSql = tSql & vbCrLf & "AND CONVERT(VARCHAR,FDShdTransDate,111) = '" & ptFDShdTransDate & "'"
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then  'PH 2.0.0 'ต้องอ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
        If (orsTemp.BOF And orsTemp.EOF) Then
            bChk = True
        Else
            bChk = False
        End If
    Else
        bChk = True
    End If
    
    W_CHKbCheckDT = bChk
Exit Function
ErrHandle:
    W_CHKbCheckDT = True
End Function

Private Sub W_FRMxSetInitial()
'-------------------------------------------------------------
'Call:
'Cmt: set ค่าเบื้องสำหรับการเริ่มต้นการทำงานขิงหน้าจอนี้
'--------------------------------------------------------------
    If bVB_ConfigDate Then          '*AO51-03-20
        If tVB_SelConfigDate = "0" Then
            olaForm(4).Caption = "(" & UCase(tCS_FormatDateM) & ")"
        Else
            olaForm(4).Caption = "(" & UCase(tCS_FormatDateSec) & ")"
        End If
    Else
        olaForm(4).Caption = "(" & UCase(tCS_FormatDateSec) & ")"
    End If
    
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnRedeem:
            bVB_TrnRedeem = False
        Case tEN_TrnDepositCancel:
'            Me.Caption = "Deposit Cancel"
            Me.Caption = "Layaway Cancel"
            bVB_TrnDepositCancel = False
        Case tEN_TrnAddDeposit
'            Me.Caption = "Add Deposit"
            Me.Caption = "Add Layaway"
            bVB_TrnAddDeposit = False
        Case tEN_TrnAddDepositCancel
'            Me.Caption = "Add Deposit Cancel"
            Me.Caption = "Cancel Add Layaway"
            bVB_TrnAddDeposit = False
    End Select
    dVB_CNDateAll = CDate(tUT_SaleDate) 'dVB_CurrentDate
    otbTerminalNo.Text = ""
    otbTransNo.Text = ""
    otbDepositNo.Text = ""
    tVB_TerminalNoRedeem = ""
    tVB_TransNoRedeem = ""
    
    If bVB_ConfigDate Then          '*AO51-03-20
        If tVB_SelConfigDate = "0" Then
            odtDate.CustomFormat = tCS_FormatDateM
            odtDate.Value = Format(Date, tCS_FormatDateM)
        Else
            odtDate.CustomFormat = tCS_FormatDateSec
            odtDate.Value = Format(Date, tCS_FormatDateSec)
        End If
    Else
        odtDate.CustomFormat = tCS_FormatDateSec
        odtDate.Value = Format(Date, tCS_FormatDateSec)
    End If
'    Call SP_OBJxSetFocus(otbTerminalNo)
'    DoEvents 'เอาไว้ไม่ให้ Cuser ไปที่ช่องหมายเลข Ticket ที่ตอน Form load ครั้งแรก '*KT 53-01-14
End Sub

'Private Sub Form_Activate()
'On Error GoTo ErrHandle
'    Call SP_OBJxSetFocus(otbTerminalNo)
'    Exit Sub
'ErrHandle:
    
'End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        SendKeys "{Tab}"
'    End If
'End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
      Call SP_FRMxChildCenter(Me)
      Call SP_FRMxGetCaptionFrmDB2Tag(Me)
      Call SP_FRMxSetCapForm(Me)
      Call W_FRMxSetInitial
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
        End With
    End If
'    Call SP_OBJxSetFocus(otbTerminalNo)
'    DoEvents
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
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnRedeem:
            bVB_TrnRedeem = False
        Case tEN_TrnDepositCancel:
            bVB_TrnAddDeposit = False
        Case tEN_TrnAddDeposit:
            bVB_TrnDepositCancel = False
    End Select
    Unload Me
End Sub

Private Sub ocmOK_Click()
    Dim tSql As String
    Dim orsTemp As ADODB.Recordset
    Dim tSqlChk As String
    Dim tDCSend As String
    Dim orsChkConf As New ADODB.Recordset
    Dim tTemp As String

    If Trim(otbDepositNo.Text) = "" Then
        Call SP_OBJxSetFocus(otbDepositNo)
        Exit Sub
    End If
    
'    If Not (W_DEFbChkDate) Then Exit Sub            '*AO51-03-20    Comm.Sheet ML-SF 109
'    dVB_CNDateAll = DateSerial(Val(otbDate(2).Text), Val(otbDate(1).Text), Val(otbDate(0).Text))
'    dVB_CNDateAll = odtDate.Value               '*AO51-03-21
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnAddDeposit, tEN_TrnRedeem, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel:
            If oDB.SQLbAvailableServer(oDB.OffServerName, oDB.OnComIP) Then
                If Not SP_CHKbOnlineEnable Then
                        MsgBox "Terminal is Offline", vbCritical
                        Call ocmCancel_Click
                        Exit Sub
                End If
            Else
                MsgBox "Terminal is Offline", vbCritical
                Call ocmCancel_Click
                Exit Sub
            End If
    End Select
    
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnAddDepositCancel
            tSql = "SELECT * FROM TPSTDepositHD"
            tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo = '" & Trim(otbDepositNo.Text) & "'"
            If bVB_TrainingMode Then 'Mode Training
                tSql = tSql & vbCrLf & "AND FTShdTransType= '" & Format(tEN_TrnTraining, "00") & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransTypeSub = '" & Format(tEN_TrnAddDeposit, "00") & "'"
            Else
                tSql = tSql & vbCrLf & "AND FTShdTransType = '" & Format(tEN_TrnAddDeposit, "00") & "'"
            End If
            tSql = tSql & vbCrLf & "AND FTShdStaDoc = '1'"
            tSql = tSql & vbCrLf & "AND ISNULL(FTShdStaCancel,'N') <> 'Y'"
            tSql = tSql & vbCrLf & "ORDER BY FDShdTransDate"
            If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then  'PH 2.0.0 'ต้องอ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
                Call SP_DATxRsNothing(orsTemp)
'                Select Case Val(Format(tVB_TransactionSub))
'                    Case tEN_TrnRedeem:
'                        Call SP_MSGnShowing("tms_CN005056", nCS_Warning)
'                    Case tEN_TrnDepositCancel:
'                        Call SP_MSGnShowing("tms_CN005057", nCS_Warning)
'                    Case tEN_TrnAddDepositCancel:
'                        Call SP_MSGnShowing("tms_CN005128", nCS_Warning)
'                    Case tEN_TrnAddDeposit:
'                        Call SP_MSGnShowing("tms_CN005129", nCS_Warning)
'                End Select
                Call SP_MSGnShowing("tms_CN005087", nCS_Warning)        '*Ao 2015/03/13  V 4.2.167 CommSheet TK-ISS3000-582
                otbDepositNo.Text = ""
                Call SP_OBJxSetFocus(otbDepositNo)
                Exit Sub
            End If
            tVB_FTTmnNum = ""
            tVB_FTShdTransNo = ""
            dVB_CNDateAll = Format("01/01/2000", "dd/MM/yyyy")
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                If orsTemp.RecordCount > 1 Then
                    wCNDepositCancel.tW_Sql = tSql
                    wCNDepositCancel.tW_Number = Trim(otbDepositNo.Text)
                    wCNDepositCancel.Show vbModal
                    If tVB_FTTmnNum = "" Then
                        Exit Sub
                    End If
                Else
                    tVB_FTTmnNum = SP_FEDtChkString(orsTemp, "FTTmnNum")
                    tVB_FTShdTransNo = SP_FEDtChkString(orsTemp, "FTShdTransNo")
                    dVB_CNDateAll = SP_FEDdChkDate(orsTemp, "FDShdSysDate")
                End If
            Else
                tSql = "SELECT FTShdStaCancel, FTShdStaPaid,FTShdStaPrcDoc  FROM TPSTDepositHD"                '*Ao 2015/02/16   V 4.2.165 CommSheet TK-ISS3000-579
                tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo = '" & Trim(otbDepositNo.Text) & "'"
                tSql = tSql & vbCrLf & "AND FTShdTransType = '" & Format(tEN_TrnDeposit, "00") & "'"
                If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) = 0 Then  'PH 2.0.0 'ต้องอ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
                    If Not (orsTemp.BOF And orsTemp.EOF) Then
                        orsTemp.MoveFirst
                        tTemp = SP_FEDtChkString(orsTemp, "FTShdStaCancel")
                        If UCase(tTemp) = UCase("Y") Then
                            'Call SP_MSGnShowing("tms_CN005255", nCS_Warning)
                            'Call SP_MSGnShowing("tms_CN005255", nCS_Warning)    '*Bump 58-02-23 Code Line นี้ใหม่ เพื่อแก้ปัญหา หา TSysLang ไม่เจอ ** ไม่แน่ใจว่าเกิดจาก อะไร
                            MsgBox "Layaway already cancelled", vbOKOnly + vbExclamation, tVB_CNAppTitle '*Bump 58-02-25 คุณเถียวแจ้งว่า ไม่แสดง msg จึง fix แทน
                        Else
                            tTemp = SP_FEDtChkString(orsTemp, "FTShdStaPrcDoc")
                            If tTemp = "1" Then
                                'Call SP_MSGnShowing("tms_CN005256", nCS_Warning)
                                'Call SP_MSGnShowing("tms_CN005256", nCS_Warning)   '*Bump 58-02-23 Code Line นี้ใหม่ เพื่อแก้ปัญหา หา TSysLang ไม่เจอ ** ไม่แน่ใจว่าเกิดจาก อะไร
                                MsgBox "Layaway already pick-up", vbOKOnly + vbExclamation, tVB_CNAppTitle '*Bump 58-02-25 คุณเถียวแจ้งว่า ไม่แสดง msg จึง fix แทน
                            End If
                        End If
                    Else
                        Call SP_MSGnShowing("tms_CN005087", nCS_Warning)            '*Ao 2015/03/13  V 4.2.167 CommSheet TK-ISS3000-582
                    End If
                Else
                    Call SP_MSGnShowing("tms_CN005087", nCS_Warning)                '*Ao 2015/03/13  V 4.2.167 CommSheet TK-ISS3000-582
                End If
                '------------------------------------------------------------------------------------------------------------------
                otbDepositNo.Text = ""
                Call SP_OBJxSetFocus(otbDepositNo)
                Exit Sub
            End If
    End Select
    
    tSql = "SELECT FTTmnNum,FTShdTransNo,FTShdTransType,FDShdTransDate,FTShdStaPaid"
    tSql = tSql & vbCrLf & " ,FTShdStaPrcDoc,FCShdRnd,FDShdDepositDueDate,FTCstCardCode"
    tSql = tSql & vbCrLf & " ,FTMemberPriceID,FTCdpDepositNo,FTXihDCSend,FDShdDepReTransDate"
    tSql = tSql & vbCrLf & " ,FTShdDepRefTmnNum,FTShdDepReTransNo ,FTPrvCardID,FTPrvCardType,FTShdStaCancel,FCShdVATRate"        '*Ao 2015/03/13  V 4.2.167 CommSheet TK-ISS3000-582
    tSql = tSql & vbCrLf & " FROM TPSTDepositHD"
    '*Ao 2012/10/18
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnAddDepositCancel
            tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & tVB_FTTmnNum & "' AND FTShdTransNo='" & tVB_FTShdTransNo & "'"
           ' tSql = tSql & vbCrLf & " AND  CONVERT(VARCHAR,FDShdSysDate,101) ='" & Format(dVB_CNDateAll, "mm/dd/yyyy") & "'" '*KT 51-05-14 ใช้ System Date แทน Operation Date
            tSql = tSql & vbCrLf & " AND  CONVERT(VARCHAR(10),FDShdSysDate,121)  = '" & Format(dVB_CNDateAll, "yyyy-MM-dd") & "'" 'พบ error vb6 เมื่อ format system date เป้น yyyy-MM-dd ไม่สามารถใช้ Format(xxx,"MM/dd/yyyy") หรือ Format(xxx,"yyyy/MM/dd") ทีมี / คั่น
        Case Else
            tSql = tSql & vbCrLf & " WHERE FTCdpDepositNo = '" & Trim(otbDepositNo.Text) & "'"
    End Select
'    tSql = tSql & vbCrLf & " WHERE  FTTmnNum='" & Trim(otbTerminalNo.Text) & "' AND FTShdTransNo='" & Trim(otbTransNo.Text) & "'"
'    tSql = tSql & vbCrLf & " AND  CONVERT(VARCHAR,FDShdSysDate,101) ='" & Format(dVB_CNDateAll, "mm/dd/yyyy") & "'" '*KT 51-05-14 ใช้ System Date แทน Operation Date
    tSql = tSql & vbCrLf & " AND  FTShdStaDoc = '1'"
'    tSql = tSql & vbCrLf & " AND  ISNULL(FTShdStaCancel,'N') <> 'Y'"                   '*Ao 2015/03/12    4.2.167 CommSheet TK-ISS3000-582
    If bVB_TrainingMode Then 'Mode Training
        tSql = tSql & vbCrLf & " AND  FTShdTransType= '" & Format(tEN_TrnTraining, "00") & "'"
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnRedeem
                tSql = tSql & vbCrLf & " AND  FTShdTransTypeSub = '" & Format(tEN_TrnDeposit, "00") & "'"
            Case tEN_TrnDepositCancel:
                tSql = tSql & vbCrLf & " AND  FTShdTransTypeSub = '" & Format(tEN_TrnDeposit, "00") & "'"
            Case tEN_TrnAddDepositCancel
                tSql = tSql & vbCrLf & " AND  FTShdTransTypeSub = '" & Format(tEN_TrnAddDeposit, "00") & "'"
            Case tEN_TrnAddDeposit
                tSql = tSql & vbCrLf & " AND  FTShdTransTypeSub = '" & Format(tEN_TrnDeposit, "00") & "'"
        End Select
    Else 'Mode ปกติ
        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnRedeem
                tSql = tSql & vbCrLf & " AND  FTShdTransType = '" & Format(tEN_TrnDeposit, "00") & "'"
            Case tEN_TrnDepositCancel
                tSql = tSql & vbCrLf & " AND  FTShdTransType = '" & Format(tEN_TrnDeposit, "00") & "'"
            Case tEN_TrnAddDepositCancel
                tSql = tSql & vbCrLf & " AND  FTShdTransType = '" & Format(tEN_TrnAddDeposit, "00") & "'"
            Case tEN_TrnAddDeposit
                tSql = tSql & vbCrLf & " AND  FTShdTransType = '" & Format(tEN_TrnDeposit, "00") & "'"
        End Select
    End If
    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine) <> 0 Then  'PH 2.0.0 'ต้องอ่านข้อมูลจากฐานข้อมูล Server เท่านั้น
        Call SP_DATxRsNothing(orsTemp)
'        Select Case Val(Format(tVB_TransactionSub))                        '*Ao 2015/03/13  V 4.2.167 CommSheet TK-ISS3000-582
'            Case tEN_TrnRedeem:
'                Call SP_MSGnShowing("tms_CN005056", nCS_Warning)            'Redeem not fould
'            Case tEN_TrnDepositCancel:
'                Call SP_MSGnShowing("tms_CN005057", nCS_Warning)
'            Case tEN_TrnAddDepositCancel:
'                Call SP_MSGnShowing("tms_CN005128", nCS_Warning)
'            Case tEN_TrnAddDeposit:
'                Call SP_MSGnShowing("tms_CN005129", nCS_Warning)
'        End Select
        Call SP_MSGnShowing("tms_CN005087", nCS_Warning)
        otbDepositNo.Text = ""
        Call SP_OBJxSetFocus(otbDepositNo)
        Exit Sub
    End If
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        If SP_FEDtChkString(orsTemp, "FTShdStaPrcDoc") = "1" Then      '*KT 51-04-04 Comm.Sheet ML-SF-203 ข้อที่ 2  ถ้าเอกสารใดที่เป็น Deposit แล้วทำการ Redeem ให้ Set เอกสารนั้น FTShdStaPrcDoc = '1'
'            Call SP_MSGnShowing("tms_CN008027", nCS_Warning)            'Already Redeem        '*Ao 2015/03/13  V 4.2.167 CommSheet TK-ISS3000-582
            MsgBox "Layaway already pick-up", vbOKOnly + vbExclamation, tVB_CNAppTitle '*Bump 58-02-25 คุณเถียวแจ้งว่า ไม่แสดง msg จึง fix แทน
            Call SP_DATxRsNothing(orsTemp)
            Call SP_DATxRsNothing(orsChkConf)
            otbDepositNo.Text = ""
            Call SP_OBJxSetFocus(otbDepositNo)
            Exit Sub
        End If
        If UCase(SP_FEDtChkString(orsTemp, "FTShdStaCancel")) = UCase("Y") Then        '*Ao 2015/03/13  V 4.2.167 CommSheet TK-ISS3000-582
            MsgBox "Layaway already cancelled", vbOKOnly + vbExclamation, tVB_CNAppTitle '*Bump 58-02-25 คุณเถียวแจ้งว่า ไม่แสดง msg จึง fix แทน
            Call SP_DATxRsNothing(orsTemp)
            Call SP_DATxRsNothing(orsChkConf)
            otbDepositNo.Text = ""
            Call SP_OBJxSetFocus(otbDepositNo)
            Exit Sub
        End If
        
        If W_CHKbCheckDT(SP_FEDtChkString(orsTemp, "FTTmnNum"), SP_FEDtChkString(orsTemp, "FTShdTransNo"), SP_DTEtStrSTD(SP_FEDdChkDate(orsTemp, "FDShdTransDate"))) Then
            MsgBox "Transaction incompleted, Please contact Administrator", vbOKOnly, tVB_CNAppTitle                                       '*Ao 2014/03/04    CommSheet TK-ISS3000-467
            otbDepositNo.Text = ""
            Call SP_OBJxSetFocus(otbDepositNo)
            Exit Sub
        End If
        If W_CHKbRecordCount(SP_FEDtChkString(orsTemp, "FTTmnNum"), SP_FEDtChkString(orsTemp, "FTShdTransNo"), SP_DTEtStrSTD(SP_FEDdChkDate(orsTemp, "FDShdTransDate"))) Then
            MsgBox "Transaction incompleted, Please contact Administrator", vbOKOnly, tVB_CNAppTitle                                       '*Ao 2014/03/04    CommSheet TK-ISS3000-467
            otbDepositNo.Text = ""
            Call SP_OBJxSetFocus(otbDepositNo)
            Exit Sub
        End If
        
        Call SP_INIxSwitchModePreviewDay(False) '*KT 53-11-05 Preview Day
        Call SP_INIxClearVarPrv '*KT 53-11-05 Preview Day
        tVB_CNPreDayIDReg = SP_FEDtChkString(orsTemp, "FTPrvCardID")  '*KT 53-11-05 Preview Day
'        nVB_RegPreType = CLng(SP_FEDtChkString(orsTemp, "FTPrvCardType"))   '*KT 53-11-05 Preview Day
        If SP_FEDtChkString(orsTemp, "FTPrvCardType") = "" Then             '*Ao 2011/01/06      ค่าว่างใช้ CLng เปลี่ยนแล้วเกิด ERROR
            nVB_RegPreType = 0
        Else
            nVB_RegPreType = CLng(SP_FEDtChkString(orsTemp, "FTPrvCardType"))
        End If

        Select Case Val(Format(tVB_TransactionSub))
            Case tEN_TrnRedeem:
                '*KT 52-10-02 DCR_10004 Phase 3.0
                If SP_FEDtChkString(orsTemp, "FTCdpDepositNo") = "" Then
                    Call SP_MSGnShowing("tms_CN005217", nCS_Warning)            'Already Redeem
                    Call SP_DATxRsNothing(orsTemp)
                    Call SP_DATxRsNothing(orsChkConf)
'                        otbTerminalNo.Text = ""                 '*AO51-03-20
'                        otbTransNo.Text = ""
'                        Call SP_OBJxSetFocus(otbTerminalNo)
                    otbDepositNo.Text = ""
                    Call SP_OBJxSetFocus(otbDepositNo)
                    Exit Sub
                End If
                
'                '*Mik 53-06-14  PRJ100053.0DCMall-DCR_5306001-00.03.00
''                tSqlChk = tSql & vbCrLf & " AND FTXihDCSend = '1' "
''                If SP_SQLvExecute(tSqlChk, orsChkData, oDB.ocnOnLine) <> 0 Then
'                tDCSend = SP_FEDtChkString(orsTemp, "FTXihDCSend")
'                If tDCSend = "1" Then
''                    Call SP_MSGnShowing("tms_CN005056", nCS_Warning)
''                Else
''                    If Not orsChkData.EOF Then
'                        tSqlChk = " SELECT     FTXihDCStatusDC, FTShdDepRefTmnNum, FTShdDepReTransNo, FDShdDepReTransDate, FTXihDocType"
'                        tSqlChk = tSqlChk & vbCrLf & "FROM         dbo.TPSTSalVatHD"
'                        tSqlChk = tSqlChk & vbCrLf & " WHERE  FTShdDepRefTmnNum = '" & SP_FEDtChkString(orsTemp, "FTShdDepRefTmnNum") & "' "
'                        tSqlChk = tSqlChk & vbCrLf & " AND FTShdDepReTransNo = '" & SP_FEDtChkString(orsTemp, "FTShdDepReTransNo") & "'"
'                        tSqlChk = tSqlChk & vbCrLf & " AND  CONVERT(VARCHAR,FDShdDepReTransDate,101)  = '" & Format(SP_FEDdChkDate(orsTemp, "FDShdDepReTransDate"), "mm/dd/yyyy") & "'"
'                        tSqlChk = tSqlChk & vbCrLf & " AND (FTXihDocType = '14') "
'                        tSqlChk = tSqlChk & vbCrLf & " AND (ISNULL(FTXihDCStatusDC,'0') = '2') "
'                        If SP_SQLvExecute(tSqlChk, orsChkConf, oDB.ocnOnLine) <> 0 Then
'                            Call SP_MSGnShowing("tms_CN005056", nCS_Warning)
'                            otbTerminalNo.Text = ""
'                            otbTransNo.Text = ""
'                            Call SP_DATxRsNothing(orsTemp)
'                            Call SP_DATxRsNothing(orsChkConf)
'                            Call SP_OBJxSetFocus(otbTerminalNo)
'                            Exit Sub
'                        Else
'                            If orsChkConf.EOF Then
''                                MsgBox " ยังไม่ส่งเอกสาร Comfirm "
'                                Call SP_MSGnShowing("tms_CN999021", nCS_Warning)
'                                otbTerminalNo.Text = ""
'                                otbTransNo.Text = ""
'                                Call SP_DATxRsNothing(orsTemp)
'                                Call SP_DATxRsNothing(orsChkConf)
'                                Call SP_OBJxSetFocus(otbTerminalNo)
'                                Exit Sub
'                            End If
'                        End If
''                    End If
'                End If
'                '-----------------------------------------------------------------------------------
                
                bVB_TrnRedeem = True
'                tVB_TerminalNoRedeem = Trim(otbTerminalNo.Text)
'                tVB_TransNoRedeem = Trim(otbTransNo.Text)
                tVB_TerminalNoRedeem = SP_FEDtChkString(orsTemp, "FTTmnNum")
                tVB_TransNoRedeem = SP_FEDtChkString(orsTemp, "FTShdTransNo")
            
                tVB_TransTypeRedeem = SP_FEDtChkString(orsTemp, "FTShdTransType")
                dVB_TransDateRedeem = SP_FEDdChkDate(orsTemp, "FDShdTransDate") 'dVB_CNDateAll
                cVB_RoundRcv = 0 'SP_FEDvChkNumeric(orsTemp, "FCShdRnd") '*KT 51-11-09 ไม่ดึงยอด Misc Recv ใบ Deposit
                tVB_MemPointIDTmp = SP_FEDtChkString(orsTemp, "FTCstCardCode") '*KT 52-04-09 เก็บค่า Member ID ที่จะได้ Point เพื่อไปคำนวน Point เมื่อจบการ Redeem
                cVB_OldGstRate = SP_FEDvChkNumeric(orsTemp, "FCShdVATRate")
            Case tEN_TrnDepositCancel:
                '*Mik 53-06-15  PRJ100053.0DCMall-DCR_5306001-00.03.00
'                tSqlChk = tSql & vbCrLf & " AND FTXihDCSend = '1' "
'                If SP_SQLvExecute(tSqlChk, orsChkData, oDB.ocnOnLine) <> 0 Then
                tDCSend = SP_FEDtChkString(orsTemp, "FTXihDCSend")
                If tDCSend = "1" Then
'                    Call SP_MSGnShowing("tms_CN005057", nCS_Warning)
'                Else
'                    If Not orsChkData.EOF Then
                        tSqlChk = " SELECT FTXihDCStatusDC,FTShdDepRefTmnNum,FTShdDepReTransNo,FDShdDepReTransDate,FTXihDocType"
                        tSqlChk = tSqlChk & vbCrLf & "FROM TPSTSalVatHD"
'                        tSqlChk = tSqlChk & vbCrLf & " WHERE  FTShdDepRefTmnNum = '" & Trim(otbTerminalNo.Text) & "' AND FTShdDepReTransNo = '" & Trim(otbTransNo.Text) & "'"
                        tSqlChk = tSqlChk & vbCrLf & " WHERE  FTShdDepRefTmnNum = '" & SP_FEDtChkString(orsTemp, "FTShdDepRefTmnNum") & "' "
                        tSqlChk = tSqlChk & vbCrLf & " AND FTShdDepReTransNo = '" & SP_FEDtChkString(orsTemp, "FTShdDepReTransNo") & "'"
                        'tSqlChk = tSqlChk & vbCrLf & " AND  CONVERT(VARCHAR,FDShdDepReTransDate,101)  = '" & Format(SP_FEDdChkDate(orsTemp, "FDShdDepReTransDate"), "mm/dd/yyyy") & "'"
                        tSqlChk = tSqlChk & vbCrLf & " AND  CONVERT(VARCHAR(10),FDShdDepReTransDate,121)  = '" & Format(SP_FEDdChkDate(orsTemp, "FDShdDepReTransDate"), "yyyy-MM-dd") & "'" 'พบ error vb6 เมื่อ format system date เป้น yyyy-MM-dd ไม่สามารถใช้ Format(xxx,"MM/dd/yyyy") หรือ Format(xxx,"yyyy/MM/dd") ทีมี / คั่น
                        tSqlChk = tSqlChk & vbCrLf & " AND (FTXihDocType = '14') "
                        tSqlChk = tSqlChk & vbCrLf & " AND (ISNULL(FTXihDCStatusDC,'0') = '2') "
                        If SP_SQLvExecute(tSqlChk, orsChkConf, oDB.ocnOnLine) <> 0 Then
                            Call SP_MSGnShowing("tms_CN005057", nCS_Warning)
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_DATxRsNothing(orsChkConf)
                            otbDepositNo.Text = ""
                            Call SP_OBJxSetFocus(otbDepositNo)
'                            otbTerminalNo.Text = ""
'                            otbTransNo.Text = ""
'                            Call SP_OBJxSetFocus(otbTerminalNo)
                            Exit Sub
                        Else
                            If Not orsChkConf.EOF Then
'                                MsgBox " ส่งเอกสาร Comfirm แล้ว "
                                Call SP_MSGnShowing("tms_CN999022", nCS_Warning)
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_DATxRsNothing(orsChkConf)
                                otbDepositNo.Text = ""
                                Call SP_OBJxSetFocus(otbDepositNo)
'                                otbTerminalNo.Text = ""
'                                otbTransNo.Text = ""
'                                Call SP_OBJxSetFocus(otbTerminalNo)
                                Exit Sub
                            End If
                        End If
'                    End If
                End If
                '-----------------------------------------------------------------------------------
                
                bVB_TrnDepositCancel = True
'                tVB_TerminalNoDptCancel = Trim(otbTerminalNo.Text)
'                tVB_TransNoDptCancel = Trim(otbTransNo.Text)
                
                tVB_TerminalNoDptCancel = SP_FEDtChkString(orsTemp, "FTTmnNum")
                tVB_TransNoDptCancel = SP_FEDtChkString(orsTemp, "FTShdTransNo")
                
                tVB_TransTypeDptCancel = SP_FEDtChkString(orsTemp, "FTShdTransType")
                dVB_TransDateDptCancel = SP_FEDdChkDate(orsTemp, "FDShdTransDate")  'dVB_CNDateAll
                cVB_RoundRcv = 0 'SP_FEDvChkNumeric(orsTemp, "FCShdRnd")  '*KT 51-11-09
                cVB_OldGstRate = SP_FEDvChkNumeric(orsTemp, "FCShdVATRate")
            Case tEN_TrnAddDepositCancel:
                '*Mik 53-06-15  PRJ100053.0DCMall-DCR_5306001-00.03.00
'                tSqlChk = tSql & vbCrLf & " AND FTXihDCSend = '1' "
'                If SP_SQLvExecute(tSqlChk, orsChkData, oDB.ocnOnLine) <> 0 Then
                tDCSend = SP_FEDtChkString(orsTemp, "FTXihDCSend")
                If tDCSend = "1" Then
'                    Call SP_MSGnShowing("tms_CN005128", nCS_Warning)
'                Else
'                    If Not orsChkData.EOF Then
                        tSqlChk = " SELECT     FTXihDCStatusDC, FTShdDepRefTmnNum, FTShdDepReTransNo, FDShdDepReTransDate, FTXihDocType"
                        tSqlChk = tSqlChk & vbCrLf & "FROM         dbo.TPSTSalVatHD"
'                        tSqlChk = tSqlChk & vbCrLf & " WHERE  FTShdDepRefTmnNum = '" & Trim(otbTerminalNo.Text) & "' AND FTShdDepReTransNo = '" & Trim(otbTransNo.Text) & "'"
                        tSqlChk = tSqlChk & vbCrLf & " WHERE  FTShdDepRefTmnNum = '" & SP_FEDtChkString(orsTemp, "FTShdDepRefTmnNum") & "' "
                        tSqlChk = tSqlChk & vbCrLf & " AND FTShdDepReTransNo = '" & SP_FEDtChkString(orsTemp, "FTShdDepReTransNo") & "'"
                        'tSqlChk = tSqlChk & vbCrLf & " AND  CONVERT(VARCHAR,FDShdDepReTransDate,101)  = '" & Format(SP_FEDdChkDate(orsTemp, "FDShdDepReTransDate"), "mm/dd/yyyy") & "'"
                        tSqlChk = tSqlChk & vbCrLf & " AND  CONVERT(VARCHAR(10),FDShdDepReTransDate,121)  = '" & Format(SP_FEDdChkDate(orsTemp, "FDShdDepReTransDate"), "yyyy-MM-dd") & "'"
                        tSqlChk = tSqlChk & vbCrLf & " AND (FTXihDocType = '14') "
                        tSqlChk = tSqlChk & vbCrLf & " AND (ISNULL(FTXihDCStatusDC,'0') = '2') "
                        If SP_SQLvExecute(tSqlChk, orsChkConf, oDB.ocnOnLine) <> 0 Then
                            Call SP_MSGnShowing("tms_CN005128", nCS_Warning)
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_DATxRsNothing(orsChkConf)
                            otbDepositNo.Text = ""
                            Call SP_OBJxSetFocus(otbDepositNo)
'                            otbTerminalNo.Text = ""
'                            otbTransNo.Text = ""
'                            Call SP_OBJxSetFocus(otbTerminalNo)
                            Exit Sub
                        Else
                            If Not orsChkConf.EOF Then
'                                MsgBox " ส่งเอกสาร Comfirm แล้ว "
                                Call SP_MSGnShowing("tms_CN999022", nCS_Warning)
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_DATxRsNothing(orsChkConf)
                                otbDepositNo.Text = ""
                                Call SP_OBJxSetFocus(otbDepositNo)
'                                otbTerminalNo.Text = ""
'                                otbTransNo.Text = ""
'                                Call SP_OBJxSetFocus(otbTerminalNo)
                                Exit Sub
                            End If
                        End If
'                    End If
                End If
                '-----------------------------------------------------------------------------------
                
                bVB_TrnDepositCancel = True
'                tVB_TerminalNoDptCancel = Trim(otbTerminalNo.Text)
'                tVB_TransNoDptCancel = Trim(otbTransNo.Text)
                tVB_TerminalNoDptCancel = SP_FEDtChkString(orsTemp, "FTTmnNum")
                tVB_TransNoDptCancel = SP_FEDtChkString(orsTemp, "FTShdTransNo")
                
                tVB_TransTypeDptCancel = SP_FEDtChkString(orsTemp, "FTShdTransType")
                dVB_TransDateDptCancel = SP_FEDdChkDate(orsTemp, "FDShdTransDate")  'dVB_CNDateAll
                cVB_RoundRcv = 0 'SP_FEDvChkNumeric(orsTemp, "FCShdRnd")  '*KT 51-11-09 ไม่ดึงยอด Misc Recv ใบ Deposit
                cVB_OldGstRate = SP_FEDvChkNumeric(orsTemp, "FCShdVATRate")
            Case tEN_TrnAddDeposit:
                '*KT 52-10-02 DCR_10004 Phase 3.0
                If SP_FEDtChkString(orsTemp, "FTCdpDepositNo") = "" Then
                        Call SP_MSGnShowing("tms_CN005217", nCS_Warning)            'Already Redeem
                        Call SP_DATxRsNothing(orsTemp)
                        Call SP_DATxRsNothing(orsChkConf)
                        otbDepositNo.Text = ""
                        Call SP_OBJxSetFocus(otbDepositNo)
'                        otbTerminalNo.Text = ""                 '*AO51-03-20
'                        otbTransNo.Text = ""
'                        Call SP_OBJxSetFocus(otbTerminalNo)
                        Exit Sub
                End If
                '*Mik 53-06-15  PRJ100053.0DCMall-DCR_5306001-00.03.00
'                tSqlChk = tSql & vbCrLf & " AND FTXihDCSend = '1' "
'                If SP_SQLvExecute(tSqlChk, orsChkData, oDB.ocnOnLine) <> 0 Then
                tDCSend = SP_FEDtChkString(orsTemp, "FTXihDCSend")
                If tDCSend = "1" Then
'                    Call SP_MSGnShowing("tms_CN005129", nCS_Warning)
'                Else
'                    If Not orsChkData.EOF Then
                        tSqlChk = " SELECT     FTXihDCStatusDC, FTShdDepRefTmnNum, FTShdDepReTransNo, FDShdDepReTransDate, FTXihDocType,FCShdVATRate"
                        tSqlChk = tSqlChk & vbCrLf & "FROM         dbo.TPSTSalVatHD"
'                        tSqlChk = tSqlChk & vbCrLf & " WHERE  FTShdDepRefTmnNum = '" & Trim(otbTerminalNo.Text) & "' AND FTShdDepReTransNo = '" & Trim(otbTransNo.Text) & "'"
                        tSqlChk = tSqlChk & vbCrLf & " WHERE  FTShdDepRefTmnNum = '" & SP_FEDtChkString(orsTemp, "FTShdDepRefTmnNum") & "' "
                        tSqlChk = tSqlChk & vbCrLf & " AND FTShdDepReTransNo = '" & SP_FEDtChkString(orsTemp, "FTShdDepReTransNo") & "'"
                       ' tSqlChk = tSqlChk & vbCrLf & " AND  CONVERT(VARCHAR,FDShdDepReTransDate,101)  = '" & Format(SP_FEDdChkDate(orsTemp, "FDShdDepReTransDate"), "mm/dd/yyyy") & "'"
                        tSqlChk = tSqlChk & vbCrLf & " AND  CONVERT(VARCHAR(10),FDShdDepReTransDate,121)  = '" & Format(SP_FEDdChkDate(orsTemp, "FDShdDepReTransDate"), "yyyy-MM-dd") & "'"
                        tSqlChk = tSqlChk & vbCrLf & " AND (FTXihDocType = '14') "
                        tSqlChk = tSqlChk & vbCrLf & " AND (ISNULL(FTXihDCStatusDC,'0') = '2') "
                        If SP_SQLvExecute(tSqlChk, orsChkConf, oDB.ocnOnLine) <> 0 Then
                            Call SP_MSGnShowing("tms_CN005129", nCS_Warning)
                            Call SP_DATxRsNothing(orsTemp)
                            Call SP_DATxRsNothing(orsChkConf)
                            otbDepositNo.Text = ""
                            Call SP_OBJxSetFocus(otbDepositNo)
'                            otbTerminalNo.Text = ""
'                            otbTransNo.Text = ""
'                            Call SP_OBJxSetFocus(otbTerminalNo)
                            Exit Sub
                        Else
                            If Not orsChkConf.EOF Then
'                                MsgBox " ส่งเอกสาร Comfirm แล้ว "
                                Call SP_MSGnShowing("tms_CN999022", nCS_Warning)
                                Call SP_DATxRsNothing(orsTemp)
                                Call SP_DATxRsNothing(orsChkConf)
                                otbDepositNo.Text = ""
                                Call SP_OBJxSetFocus(otbDepositNo)
'                                otbTerminalNo.Text = ""
'                                otbTransNo.Text = ""
'                                Call SP_OBJxSetFocus(otbTerminalNo)
                                Exit Sub
                            End If
                        End If
'                    End If
                End If
                '-----------------------------------------------------------------------------------
                
                bVB_TrnAddDeposit = True
'                tVB_TerminalNoAddDpt = Trim(otbTerminalNo.Text)
'                tVB_TransNoAddDpt = Trim(otbTransNo.Text)
                tVB_TerminalNoAddDpt = SP_FEDtChkString(orsTemp, "FTTmnNum")
                tVB_TransNoAddDpt = SP_FEDtChkString(orsTemp, "FTShdTransNo")
                
                tVB_TransTypeAddDpt = SP_FEDtChkString(orsTemp, "FTShdTransType")
                dVB_TransDateAddDpt = SP_FEDdChkDate(orsTemp, "FDShdTransDate") ' dVB_CNDateAll
                dVB_TransDateDueDeposit = SP_FEDdChkDate(orsTemp, "FDShdDepositDueDate")
                cVB_RoundRcv = 0 'SP_FEDvChkNumeric(orsTemp, "FCShdRnd")  '*KT 51-11-09 ไม่ดึงยอด Misc Recv ใบ Deposit
                cVB_OldGstRate = SP_FEDvChkNumeric(orsTemp, "FCShdVATRate")
        End Select
        
        '*KT 52-08-14 member price
        xVB_CNMemberP.ID = SP_FEDtChkString(orsTemp, "FTMemberPriceID")
        '-------------------------------------------------------------------------------------------------------
        If xVB_CNMemberP.ID <> "" Then
            bVB_CNMemberPrice = True
            tSql = "SELECT Top 1 FTMcdName,FTMcdSurName,FTMcdCode,FTMcdTel,FDMcdExpDate,FTMcdSta  "
            tSql = tSql & vbCrLf & " FROM TCNMMallCard"
            tSql = tSql & vbCrLf & " WHERE FTMcdCode='" & xVB_CNMemberP.ID & "'"
            If SP_SQLvExecute(tSql, orsTemp, ocnMember, , , , True) = 0 Then
                If Not (orsTemp.BOF Or orsTemp.EOF) Then
                    xVB_CNMemberP.FName = SP_FEDtChkString(orsTemp, "FTMcdName")
                    xVB_CNMemberP.LName = SP_FEDtChkString(orsTemp, "FTMcdSurName")
                    xVB_CNMemberP.Tel = SP_FEDtChkString(orsTemp, "FTMcdTel")
                    xVB_CNMemberP.Expire = Date
                    xVB_CNMemberP.Status = SP_FEDtChkString(orsTemp, "FTMcdSta")
                    tSql = "STP_PRCxFindMallCardByIDEx @ptAccNum= '" & xVB_CNMemberP.ID & "'"
                    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                        If Not (orsTemp.BOF Or orsTemp.EOF) Then
                            xVB_CNMemberP.MallCardType = SP_FEDtChkString(orsTemp, "FTMcdCardType")
                            xVB_CNMemberP.MallCardExpire = SP_FEDdChkDate(orsTemp, "FDMcdEndDate")
                            xVB_CNMemberP.MallCardMemberPrice = SP_FEDtChkString(orsTemp, "FTMcdAlwPrice")
                            xVB_CNMemberP.MallCardMemberPoint = SP_FEDtChkString(orsTemp, "FTMcdAlwPoint")
                        End If
                    End If
                     wMain.ostForm.Panels(2).Text = "Mem:" & xVB_CNMemberP.ID
                Else
                    xVB_CNMemberP.FName = "New Member"
                    xVB_CNMemberP.LName = ""
                    xVB_CNMemberP.Tel = ""
                    xVB_CNMemberP.Expire = Date
                    xVB_CNMemberP.Status = ""
                    tSql = "STP_PRCxFindMallCardByIDEx @ptAccNum= '" & xVB_CNMemberP.ID & "'"
                    If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                        If Not (orsTemp.BOF Or orsTemp.EOF) Then
                            xVB_CNMemberP.MallCardType = SP_FEDtChkString(orsTemp, "FTMcdCardType")
                            xVB_CNMemberP.MallCardExpire = SP_FEDdChkDate(orsTemp, "FDMcdEndDate")
                            xVB_CNMemberP.MallCardMemberPrice = SP_FEDtChkString(orsTemp, "FTMcdAlwPrice")
                            xVB_CNMemberP.MallCardMemberPoint = SP_FEDtChkString(orsTemp, "FTMcdAlwPoint")
                        End If
                    End If
                     wMain.ostForm.Panels(2).Text = "Mem:" & xVB_CNMemberP.ID
                End If
            Else
                xVB_CNMemberP.FName = "New Member"
                xVB_CNMemberP.LName = ""
                xVB_CNMemberP.Tel = ""
                xVB_CNMemberP.Expire = Date
                xVB_CNMemberP.Status = ""
                tSql = "STP_PRCxFindMallCardByIDEx @ptAccNum= '" & xVB_CNMemberP.ID & "'"
                If SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) = 0 Then 'read master
                    If Not (orsTemp.BOF Or orsTemp.EOF) Then
                        xVB_CNMemberP.MallCardType = SP_FEDtChkString(orsTemp, "FTMcdCardType")
                        xVB_CNMemberP.MallCardExpire = SP_FEDdChkDate(orsTemp, "FDMcdEndDate")
                        xVB_CNMemberP.MallCardMemberPrice = SP_FEDtChkString(orsTemp, "FTMcdAlwPrice")
                        xVB_CNMemberP.MallCardMemberPoint = SP_FEDtChkString(orsTemp, "FTMcdAlwPoint")
                    End If
                End If
                 wMain.ostForm.Panels(2).Text = "Mem:" & xVB_CNMemberP.ID
            End If
        Else
            bVB_CNMemberPrice = False
            xVB_CNMemberP.ID = ""
            xVB_CNMemberP.FName = ""
            xVB_CNMemberP.LName = ""
            xVB_CNMemberP.Tel = ""
            xVB_CNMemberP.Expire = Date
            xVB_CNMemberP.Status = ""
            xVB_CNMemberP.MallCardType = ""
            xVB_CNMemberP.MallCardExpire = Date
            xVB_CNMemberP.MallCardMemberPrice = "1"
            wMain.ostForm.Panels(2).Text = "Ready..."
        End If
        '-------------------------------------------------------------------------------------------------------------
        '*KT 53-11-05 Preview Day
        '-------------------------INITIAL-----------------------------------------------------------------------
        '*KT 53-11-08 จาก Mail วันที่ 08/11/2010
        If tVB_CNPreDayIDReg <> "" Then
            'อยู่ในช่วง Preview Day
'            If SP_CHKbPeriodPreDay Then
                bVB_CNTrStaRegPreviewDay = True
                Call SP_INIxSwitchModePreviewDay(True)
'                wMain.olaTransName.Caption = "Preview/" & tVB_TransName
'                Select Case nVB_RegPreType
'                    Case 1 'Credit Card
'                        Call SP_GETxRecDataCrdPrv
'                    Case 2 'Member Card
'                        Call SP_GETxRecDataMallPrv
'                    Case 3 'None Card
'                        Call SP_GETxRecDataMallPrv
'                End Select
'                If Not SP_UPDbPreviewDayHD Then 'บันทึกข้อมูล Preview Day ลงตาราง HD
'                    Call SP_MSGnShowing("tms_CN005221", nCS_Warning)
'                    Call SP_INIxSwitchModePreviewDay(False)
'                    Call SP_INIxClearVarPrv
'                    wMain.olaTransName.Caption = tVB_TransName
'                End If
'            End If
        End If
        '-------------------------END INITIAL---------------------------------------------------------------------
    Else
        Call SP_DATxRsNothing(orsTemp)
'        Select Case Val(Format(tVB_TransactionSub))
'            Case tEN_TrnRedeem:
'                Call SP_MSGnShowing("tms_CN005056", nCS_Warning)            'Redeem not fould
'            Case tEN_TrnDepositCancel:
'                Call SP_MSGnShowing("tms_CN005057", nCS_Warning)
'            Case tEN_TrnAddDepositCancel:
'                Call SP_MSGnShowing("tms_CN005128", nCS_Warning)
'            Case tEN_TrnAddDeposit:
'                Call SP_MSGnShowing("tms_CN005129", nCS_Warning)
'        End Select
        Call SP_MSGnShowing("tms_CN005087", nCS_Warning)        '*Ao 2015/03/13  V 4.2.167 CommSheet TK-ISS3000-582
        Call SP_DATxRsNothing(orsTemp)
        Call SP_DATxRsNothing(orsChkConf)
        otbDepositNo.Text = ""
        Call SP_OBJxSetFocus(otbDepositNo)
        Exit Sub
    End If
    Call SP_DATxRsNothing(orsChkConf)
    Call SP_DATxRsNothing(orsTemp)
    Unload Me
End Sub

Private Sub odtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            If IsDate(odtDate.Value) Then
                SendKeys "{Tab}"
            End If
    End Select
End Sub

Private Sub oScanner_DataEvent(ByVal Status As Long)
Dim tID As String
On Error GoTo ErrHandle
'    With oScanner
'        tID = .ScanData
''        DoEvents            '*KT 53-01-14
'        If tID <> "" Then
'            If Asc(Right(tID, 1)) = 13 Then
'                tID = Mid(tID, 1, Len(tID) - 1)
'            End If
'            otbTerminalNo.Text = Left(tID, 5)
'            otbTransNo.Text = Right(tID, 5)
'        End If
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
'    If otbTerminalNo.Text <> "" And otbTransNo.Text <> "" Then
'        Call SP_OBJxSetFocus(odtDate)
'    End If
    Exit Sub
ErrHandle:
    MsgBox "Scaner error" & vbCrLf & "Runtime : " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub otbDate_GotFocus(Index As Integer)
    Call SP_OBJxSetFocus(otbDate(Index))
    otbDate(Index).SelStart = 0
    otbDate(Index).SelLength = Len(otbDate(Index).Text)
End Sub

Private Sub otbDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbDate(Index).Text = ""
    End Select
End Sub

Private Sub otbDate_KeyPress(Index As Integer, KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 46             'ทศนิยม
            KeyAscii = 0
'            If InStr(Trim(otbDepositAmt.Text), ".") > 0 Then
'                KeyAscii = 0
'            End If
        Case 13
            If Index = 2 Then
                Call ocmOK_Click
            End If
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
    
End Sub

Private Sub otbDepositNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbDepositNo.Text = ""
        Case 13
            If otbDepositNo.Text <> "" Then
                SendKeys "{Tab}"
            End If
    End Select
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
                SendKeys "{Tab}"
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
            If otbTerminalNo.Text <> "" Then
                SendKeys "{Tab}"
            End If
    End Select
End Sub

Private Sub otbTransNo_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
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
End Sub

Private Sub otbTransNo_LostFocus()
        otbTransNo.Text = SP_TXTtTrimRightLeft(otbTransNo.Text)
End Sub

