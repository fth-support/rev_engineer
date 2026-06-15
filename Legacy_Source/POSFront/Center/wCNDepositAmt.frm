VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.v11.1.3.ocx"
Begin VB.Form wCNDepositAmt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposit Amount Entry;klg_Title"
   ClientHeight    =   3885
   ClientLeft      =   1965
   ClientTop       =   4980
   ClientWidth     =   6645
   ControlBox      =   0   'False
   Icon            =   "wCNDepositAmt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6645
   Tag             =   "2;"
   Begin VB.Frame ofmTmp 
      Height          =   1095
      Left            =   1920
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox otbDate 
         Enabled         =   0   'False
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
         TabIndex        =   14
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox otbDate 
         Enabled         =   0   'False
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
         TabIndex        =   13
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox otbDate 
         Enabled         =   0   'False
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
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DD/MM/YYYY"
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
         TabIndex        =   15
         Top             =   360
         Width           =   1710
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   3240
      Index           =   0
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   -240
      Width           =   6645
      _Version        =   720897
      _ExtentX        =   11721
      _ExtentY        =   5715
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbDepositPhone 
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
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1080
         Width           =   2510
      End
      Begin MSComCtl2.DTPicker odtDate 
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   2280
         Width           =   2505
         _ExtentX        =   4419
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
         Format          =   59506691
         CurrentDate     =   39528
      End
      Begin VB.TextBox otbDepositAmt 
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
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "0"
         Top             =   1680
         Width           =   2510
      End
      Begin VB.TextBox otbDepositNo 
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
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number:;klg_PhoneNum"
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
         Tag             =   "2;"
         Top             =   1080
         Width           =   3990
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
         Left            =   2640
         TabIndex        =   16
         Top             =   2820
         Width           =   1860
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date:;klg_DueDate"
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
         Left            =   120
         TabIndex        =   10
         Tag             =   "2;"
         Top             =   2355
         Width           =   2160
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Amount:;klg_DepositAmt"
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
         TabIndex        =   8
         Tag             =   "2;"
         Top             =   1680
         Width           =   4140
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit Number:;klg_DepositNo"
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
         Top             =   480
         Width           =   4035
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1200
      Index           =   1
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2700
      Width           =   6645
      _Version        =   720897
      _ExtentX        =   11721
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
         TabIndex        =   5
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
         TabIndex        =   4
         Tag             =   "2;"
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNDepositAmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tVB_DepositPhone As String

Private Function W_PRCtPhone() As String
'-------------------------------------------------------------
'Call:
'Cmt: '*Ao 2013/09/07
'--------------------------------------------------------------
Dim orsTemp As ADODB.Recordset
Dim tSql As String, tTmp As String

On Error GoTo ErrHandler
'    tTmp = ""
'    tSql = "SELECT TOP 1 ISNULL(FTPmhDocNoBill,'') AS FTPmhDocNoBill FROM " & tVB_CNTblHD
'    tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo = '" & tVB_DepossitNoAddDpt & "'"
'    tSql = tSql & vbCrLf & "AND NOT (FTTmnNum = '" & tVB_CNTerminalNum & "'"
'    tSql = tSql & vbCrLf & "AND FTShdTransNo = '" & tVB_Running & "'"
'    tSql = tSql & vbCrLf & "AND FDShdTransDate = '" & tUT_SaleDate & "'"
'    tSql = tSql & vbCrLf & "AND FTShdTransType = '" & tVB_Transaction & "')"
'    tSql = tSql & vbCrLf & "ORDER BY FDShdSysDate DESC,FTShdSysTime DESC"
'    If SP_SQLbExecute(tSql, orsTemp) Then
'        If Not (orsTemp.BOF And orsTemp.EOF) Then
'            tTmp = SP_FEDtChkString(orsTemp, "FTPmhDocNoBill")
'        End If
'    End If
'    W_PRCtPhone = tTmp

    '*Eaw 56-10-28
    tTmp = ""
    tSql = "SELECT ISNULL(FTPmhDocNoBill,'') AS FTPmhDocNoBill FROM TPSTDepositHD"
    tSql = tSql & vbCrLf & "WHERE FTCdpDepositNo = '" & tVB_DepossitNoAddDpt & "'"
    tSql = tSql & vbCrLf & "ORDER BY FDShdSysDate DESC,FTShdSysTime DESC"
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True)
    Do While Not orsTemp.EOF
        tTmp = SP_FEDtChkString(orsTemp, "FTPmhDocNoBill")
        If tTmp <> "" Then
            W_PRCtPhone = tTmp
            Exit Function
        End If
        orsTemp.MoveNext
    Loop
    W_PRCtPhone = tTmp
Exit Function
ErrHandler:
    W_PRCtPhone = ""
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

    otbDepositAmt.Text = 0
    tVB_DepositNo = ""
    cVB_DepositAmt = 0
    tVB_DepositPhone = ""
    
    If bVB_ConfigDate Then          '*AO51-03-20
        If tVB_SelConfigDate = "0" Then
            odtDate.CustomFormat = tCS_FormatDateM
'            odtDate.Value = Format(Date, tCS_FormatDateM)    '*Eaw 56-02-26  TK-ISS3000-075
            odtDate.Value = Format(DateAdd("m", 1, Date), tCS_FormatDateM)
        Else
            odtDate.CustomFormat = tCS_FormatDateSec
            odtDate.Value = Format(Date, tCS_FormatDateSec)
        End If
    Else
        odtDate.CustomFormat = tCS_FormatDateSec
        odtDate.Value = Format(Date, tCS_FormatDateSec)
    End If

    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnDeposit:
            otbDepositNo.Text = ""
            otbDepositPhone.Text = ""
            otbDepositNo.Text = W_SETtDepositNum
        
        Case tEN_TrnAddDeposit
            otbDepositNo.Text = tVB_DepossitNoAddDpt
            '*KT 51-09-30 Default วันที่ Due ตาม Deposit
            '-----------------------------------------------------------------------------------------------------------------------
            If bVB_ConfigDate Then
                If tVB_SelConfigDate = "0" Then
                    odtDate.Value = Format(dVB_TransDateDueDeposit, tCS_FormatDateM)
                Else
                    odtDate.Value = Format(dVB_TransDateDueDeposit, tCS_FormatDateSec)
                End If
            Else
                odtDate.Value = Format(dVB_TransDateDueDeposit, tCS_FormatDateSec)
            End If
            otbDepositPhone.Text = W_PRCtPhone                  '*Ao 2013/09/07  CommSheet TK-ISS3000-346-UAT-FAPL-117   V.4.0.61
            '------------------------------------------------------------------------------------------------------------------------
    End Select
End Sub

Private Sub PRCxDeposit()
    Dim tSql As String

'*KT 52-10-02 DCR_10004 Phase 3.0
'    If otbDepositNo.Text = "" Then
'        otbDepositNo.SetFocus
'        Exit Sub
'    ElseIf SP_VALcText2Double(otbDepositAmt.Text) < 0 Then '*KT 51-04-04 ML-SF-203 สามารถทำ Deposit จำนวน 0 ได้
'        otbDepositAmt.SetFocus
'        Exit Sub
'    End If
    If SP_VALcText2Double(otbDepositAmt.Text) < 0 Then '*KT 51-04-04 ML-SF-203 สามารถทำ Deposit จำนวน 0 ได้
        otbDepositAmt.SetFocus
        Exit Sub
    End If
    
    If SP_VALcText2Double(otbDepositAmt.Text) > cUT_SubTotal Then
        Call SP_MSGnShowing("tms_CN005064", vbInformation)
        otbDepositAmt.SetFocus
        Exit Sub
    End If
    
    tVB_DepositNo = Trim(otbDepositNo.Text)
    cVB_DepositAmt = CDbl(SP_FMTcCurPoint(SP_VALcText2Double(otbDepositAmt.Text), True, False))
    tVB_DepositPhone = Trim(otbDepositPhone.Text)
    
    '*KT 52-07-27 ป้องกันการ Key จำนวนมัดจำเท่ากับยอด Miss Recv.
    If cVB_DepositAmt = CDbl(Abs(IIf(bVB_CNAlwRnd, UT_GetnRoundDiff(cVB_DepositAmt), 0))) And cUT_SubTotal <> cVB_DepositAmt And cVB_DepositAmt <> 0 Then
        Call SP_MSGnShowing("tms_CN005064", vbInformation)
        tVB_DepositNo = ""
        cVB_DepositAmt = 0
        otbDepositAmt.SetFocus
        Exit Sub
    End If
    
    '*KT 52-06-04 PH1.8-ML-SF-042
    If cVB_DepositAmt = cUT_SubTotal Then
        bVB_DepositAll = True
    Else
        bVB_DepositAll = False
    End If
    dVB_CNDateAll = odtDate.Value               '*AO51-03-21
    dVB_DepositDueDate = dVB_CNDateAll
'    cVB_RoundRcv = IIf(bVB_CNAlwRnd, UT_GetnRoundDiff(cUT_SubTotal), 0)  '*KT 51-04-18 Comsheet ML-SF-172/ML-SF-173
    
    
    cVB_RoundRcv = 0 '*Eaw 56-01-04 แก้ไขการปัด Rounding
    tSql = "UPDATE " & tVB_CNTblHD
    tSql = tSql & " SET FTCdpDepositNo='" & tVB_DepositNo & "'"
    tSql = tSql & " ,FCShdDepositAmt=" & cVB_DepositAmt
    tSql = tSql & " ,FDShdDepositDuedate='" & SP_DTEtStrSTD(dVB_DepositDueDate) & "'"
    'tSql = tSql & " ,FCShdDepositBal=" & SP_FMTcCurPoint((cUT_SubTotal + cVB_RoundRcv) - cVB_DepositAmt, True, False)
    tSql = tSql & " ,FCShdDepositBal=" & SP_FMTcCurPoint(cUT_SubTotal - cVB_DepositAmt, True, False)
    tSql = tSql & " ,FCShdDepositBPay=" & 0
    tSql = tSql & " ,FTShdDepRefTmnNum='" & tVB_CNTerminalNum & "'" 'หมายเลขเครื่องขายตัวเอง
    tSql = tSql & " ,FTShdDepReTransNo='" & tVB_Running & "'" 'หมายเลขเอกสารของตัวเอง
    tSql = tSql & " ,FDShdDepReTransDate='" & tUT_SaleDate & "'"   'วันที่ทำเอกสารของตัวเอง
    tSql = tSql & " ,FCShdDepositPrvBal= " & 0
    tSql = tSql & " ,FCShdTotal=" & cUT_SubTotal
    tSql = tSql & " ,FCShdGrand=" & IIf((cVB_Total - cVB_TotalDiscAft) <= 0, 0, (cVB_Total - cVB_TotalDiscAft))
    tSql = tSql & " ,FCShdRnd=" & cVB_RoundRcv
    tSql = tSql & " ,FTPmhDocNoBill='" & tVB_DepositPhone & "'" 'Phone Number
    tSql = tSql & "," & UT_tLastUpdSql
    tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & "  AND FTShdStaDoc = '2'"
    Call SP_SQLbExecute(tSql)
    bVB_TrnDeposit = True
    Unload Me
End Sub

Private Sub PRCxAddDeposit()
    Dim tSql As String
    
    If otbDepositNo.Text = "" Then
        Call SP_OBJxSetFocus(otbDepositNo) 'otbDepositNo.SetFocus
        Exit Sub
    ElseIf SP_VALcText2Double(otbDepositAmt.Text) <= 0 Then
        Call SP_OBJxSetFocus(otbDepositAmt) ' otbDepositAmt.SetFocus
        Exit Sub
    End If
    If SP_VALcText2Double(otbDepositAmt.Text) > cUT_SubTotal Then
        Call SP_MSGnShowing("tms_CN005064", vbInformation)
        otbDepositAmt.SetFocus
        Exit Sub
    End If
    tVB_DepossitNoAddDpt = Trim(otbDepositNo.Text)
    cVB_DepossitAmtAddDpt = SP_FMTcCurPoint(SP_VALcText2Double(otbDepositAmt.Text), True, False)
    cVB_DepositAmt = cVB_DepossitAmtAddDpt
    tVB_DepositPhone = Trim(otbDepositPhone.Text) '*Eaw 56-07-25
    '*KT 52-06-04 PH1.8-ML-SF-042
    If cVB_DepositAmt = cUT_SubTotal Then
        bVB_DepositAll = True
    Else
        bVB_DepositAll = False
    End If
    cVB_DepossitAmtAddDptOrg = cVB_DepossitAmtAddDpt
    
    dVB_CNDateAll = odtDate.Value               '*AO51-03-21
    dVB_DepositDueDateAddDpt = dVB_CNDateAll
    'cVB_RoundRcv = IIf(bVB_CNAlwRnd, UT_GetnRoundDiff(cUT_SubTotal), 0) '*KT 51-04-18 Comsheet ML-SF-172/ML-SF-173
    
    tSql = "UPDATE " & tVB_CNTblHD
    tSql = tSql & " SET FTCdpDepositNo='" & tVB_DepossitNoAddDpt & "'"
    tSql = tSql & " ,FCShdDepositAmt=" & cVB_DepossitAmtAddDpt
    tSql = tSql & " ,FDShdDepositDueDate='" & SP_DTEtStrSTD(dVB_DepositDueDateAddDpt) & "'"
    tSql = tSql & " ,FCShdDepositBal=" & SP_FMTcCurPoint(cVB_AccBalAddDpt - cVB_DepossitAmtAddDpt, True, False)
    tSql = tSql & " ,FCShdDepositBPay=" & SP_FMTcCurPoint(cVB_DepossitAmtBAddDpt, True, False)
    tSql = tSql & " ,FTShdDepRefTmnNum='" & tVB_TerminalNoAddDpt & "'" 'หมายเลขเครื่องขาย Deposit  แม่
    tSql = tSql & " ,FTShdDepReTransNo='" & tVB_TransNoAddDpt & "'" 'หมายเลขเอกสารของ Deposit  แม่
    tSql = tSql & " ,FDShdDepReTransDate='" & SP_DTEtStrSTD(dVB_TransDateAddDpt) & "'"   'วันที่ทำเอกสารของ Deposit  แม่
    tSql = tSql & " ,FCShdDepositPrvBal= " & SP_FMTcCurPoint(cVB_AccBalAddDpt, True, False)
    tSql = tSql & " ,FTPmhDocNoBill='" & tVB_DepositPhone & "'" 'Phone Number '*Eaw 56-07-25
    tSql = tSql & "," & UT_tLastUpdSql
    tSql = tSql & "  WHERE FTTmnNum='" & tVB_CNTerminalNum & "'"
    tSql = tSql & "  AND FTShdTransNo='" & tVB_Running & "'"
    tSql = tSql & "  AND FDShdTransDate ='" & tUT_SaleDate & "'"
    tSql = tSql & "  AND FTShdTransType='" & tVB_Transaction & "'"
    tSql = tSql & "  AND FTShdStaDoc = '2'"
    Call SP_SQLbExecute(tSql)
    Unload Me
End Sub

'*KT 53-01-19 ไม่ใช้งานแล้ว ทำให้เกิดการ setfocus มายัง TabIndex 1
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then SendKeys "{Tab}"
'End Sub

Private Sub Form_Load()
      Call SP_FRMxChildCenter(Me)
        If bVB_Res640X480 Then
    '        Me.Left = 0
             Me.Top = 2100
        End If
      Call SP_FRMxGetCaptionFrmDB2Tag(Me)
      Call SP_FRMxSetCapForm(Me)
      otbDepositNo.Locked = True
      otbDepositNo.BackColor = nVB_CNClrBLock
      Call W_FRMxSetInitial
'      Call SP_OBJxSetFocus(otbDepositAmt)
      Call SP_OBJxSetFocus(otbDepositPhone)
      otbDepositAmt.Text = SP_FMTcCurPoint(cVB_TotalVendor, False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocmCancel_Click()
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnDeposit:
            bVB_TrnDeposit = False
            tVB_DepositNo = ""
            cVB_DepositAmt = 0
        Case tEN_TrnAddDeposit:
            bVB_TrnAddDeposit = False
            tVB_DepossitNoAddDpt = ""
            cVB_DepossitAmtAddDpt = 0
    End Select
     Unload Me
End Sub

Private Sub ocmOK_Click()
'    If Not (W_DEFbChkDate) Then Exit Sub            '*AO51-03-20    Comm.Sheet ML-SF 109
    
    If cVB_TotalVendor <> 0 Then                '*Ao 53-10-21   cVB_TotalVendor = Total Vendor
        If SP_VALcText2Double(otbDepositAmt.Text) < SP_VALcText2Double(cVB_TotalVendor) Then
            
            Exit Sub
        End If
    End If
    '______________________________________________________________________________________
    dVB_CNDateAll = DateSerial(Val(otbDate(2).Text), Val(otbDate(1).Text), Val(otbDate(0).Text))
    Select Case Val(Format(tVB_TransactionSub))
        Case tEN_TrnDeposit: PRCxDeposit
        Case tEN_TrnAddDeposit: PRCxAddDeposit
    End Select
End Sub

Private Sub odtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            If IsDate(odtDate.Value) Then
                SendKeys "{Tab}"
            End If
    End Select
End Sub

Private Sub otbDepositAmt_Change() '*AO51-03-20
    otbDepositAmt.Text = SP_DEFtChangeNum(Trim(otbDepositAmt.Text))
    otbDepositAmt.SelStart = Len(otbDepositAmt.Text)
End Sub

Private Sub otbDepositAmt_GotFocus()
    otbDepositAmt.SelStart = 0
    otbDepositAmt.SelLength = Len(otbDepositAmt.Text)
End Sub

Private Sub otbDepositAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbDepositAmt.Text = ""
        Case 13
            SendKeys "{Tab}"
    End Select
End Sub

Private Sub otbDepositAmt_KeyPress(KeyAscii As Integer)
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
            'Call ocmOk_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbDepositNo_GotFocus()
    otbDepositNo.SelStart = 0
    otbDepositNo.SelLength = Len(otbDepositNo.Text)
End Sub

Private Sub otbDepositNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbDepositNo.Text = ""
    End Select
End Sub

Private Sub otbDepositNo_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
End Sub

Private Function W_SETtDepositNum() As String
'__________________________
'   Call:
'   Cmt:   set หมายเลข Deposit Number
'__________________________
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim tDayOfYear As String

On Error GoTo ErrHandler
    
        W_SETtDepositNum = ""
    
        tSql = "SELECT DATEPART (DY,'" & Format(Now, "yyyyMMdd") & "') AS FTDayOfYear"
                 
        If SP_SQLbExecute(tSql, orsTemp) Then
            If Not (orsTemp.BOF And orsTemp.EOF) Then
                Do While Not orsTemp.EOF
                    tDayOfYear = SP_FEDtChkString(orsTemp, "FTDayOfYear")
                    orsTemp.MoveNext
                Loop
                W_SETtDepositNum = Right(("000" & tDayOfYear), 3) & tVB_CNTerminalNum & tVB_Running & Right(tVB_Running, 1) '*Bump 56-02-19 CommSheet TK-ISS3000-042 00.00.02.xls
                'W_SETtDepositNum = Right(("000" & tDayOfYear), 3) & tVB_Running & tVB_CNTerminalNum & Right(tVB_CNTerminalNum, 1)
            End If
        End If

    Exit Function
ErrHandler:
    Set orsTemp = Nothing
    Call SP_MSGShowError
End Function

Private Sub otbDepositPhone_GotFocus()
    otbDepositPhone.SelStart = 0
    otbDepositPhone.SelLength = Len(otbDepositPhone.Text)
End Sub

Private Sub otbDepositPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            otbDepositPhone.Text = ""
        Case 13
            SendKeys "{Tab}"
    End Select
End Sub

Private Sub otbDepositPhone_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
'        Case 46             'ทศนิยม
'            KeyAscii = 0
'            If InStr(Trim(otbDepositAmt.Text), ".") > 0 Then
'                KeyAscii = 0
'            End If
        Case 13
            'Call ocmOk_Click
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub
