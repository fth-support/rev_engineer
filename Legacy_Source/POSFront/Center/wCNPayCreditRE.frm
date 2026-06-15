VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.1#0"; "Codejock.SuiteCtrls.Unicode.v11.1.3.ocx"
Begin VB.Form wCNPayCreditRE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Card"
   ClientHeight    =   4620
   ClientLeft      =   5415
   ClientTop       =   5115
   ClientWidth     =   6540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6540
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   4080
      Index           =   2
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   -270
      Width           =   6555
      _Version        =   720897
      _ExtentX        =   11562
      _ExtentY        =   7188
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.TextBox otbAcc 
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
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox otbNum 
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
         TabIndex        =   5
         Text            =   "0"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox otbApvCode 
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
         MaxLength       =   6
         TabIndex        =   6
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox otbCardType 
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2160
         Width           =   3975
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
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1560
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
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1560
         Width           =   615
      End
      Begin XtremeSuiteControls.ComboBox ocbType 
         Height          =   495
         Left            =   2400
         TabIndex        =   0
         Top             =   360
         Width           =   2460
         _Version        =   720897
         _ExtentX        =   4339
         _ExtentY        =   873
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration Date :"
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
         TabIndex        =   17
         Tag             =   "2;"
         Top             =   1635
         Width           =   2115
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number :"
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
         TabIndex        =   16
         Tag             =   "2;"
         Top             =   1035
         Width           =   2310
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount :"
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
         TabIndex        =   15
         Tag             =   "2;"
         Top             =   2835
         Width           =   1170
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved Code :"
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
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Tag             =   "2;"
         Top             =   3435
         Width           =   2115
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
         Left            =   3120
         TabIndex        =   13
         Top             =   1560
         Width           =   210
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Card Type :"
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
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Tag             =   "2;"
         Top             =   2235
         Width           =   1485
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(MM/YY)"
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
         Left            =   4200
         TabIndex        =   11
         Tag             =   "2;"
         Top             =   1635
         Width           =   1095
      End
      Begin VB.Label olaForm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type :"
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
         Index           =   12
         Left            =   120
         TabIndex        =   10
         Tag             =   "2;"
         Top             =   435
         Width           =   810
      End
   End
   Begin XtremeSuiteControls.TabControl otcForm1 
      Height          =   1080
      Index           =   1
      Left            =   0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3540
      Width           =   6555
      _Version        =   720897
      _ExtentX        =   11562
      _ExtentY        =   1905
      _StockProps     =   64
      Appearance      =   10
      Color           =   -1484272203
      Begin VB.CommandButton ocmOk 
         Caption         =   "&Ok"
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
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   400
         Width           =   1455
      End
      Begin VB.CommandButton ocmCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
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
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "2;"
         Top             =   400
         Width           =   1455
      End
   End
End
Attribute VB_Name = "wCNPayCreditRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private orsCardGrp As ADODB.Recordset
Private nW_CaseERROR As Integer

Private Sub W_CHKxCaseERROR(ByVal pnCase As Integer)
    Select Case pnCase
        Case 1
            Call SP_MSGnShowing("tms_CN008038", nCS_Error)
        Case 2
            Call SP_MSGnShowing("tms_CN008039", nCS_Error)      '*Ao 2013/03/15
        Case Else
            Call SP_MSGnShowing("tms_CN005065", nCS_Error)
    End Select
End Sub

Private Function W_PRCxStrCupsCard(ByVal ptAccNum As String) As String
'---------------------------------------------------------------------------
'Call:
'Cmt: String ตรวจสอบว่าเป็น Free Rang Card หรือไม่  '*Eaw 56-10-19
'----------------------------------------------------------------------------
    Dim tSql As String
    tSql = "SELECT TOP 1 * FROM TCNMCreditCard"
    tSql = tSql & " WHERE FTCdcCreditType = '" & Trim(ocbType.Text) & "'"
    tSql = tSql & " AND LEN('" & ptAccNum & "') <= FNCdcDigits"
    W_PRCxStrCupsCard = tSql
End Function

Private Function W_PRCxStrSQL(ByVal ptAccNum$) As String
Dim tSql As String
Dim nLen As Long '*Eaw 56-07-10 CommSheet TK-ISS3000-272-TR.xls
'---------------------------------------------------------------------------
'Call:
'Cmt:  '*Eaw 56-10-19
'----------------------------------------------------------------------------
    nLen = Len(ptAccNum) '*Eaw 56-07-10

    tSql = "SELECT TOP 1 TCNMCreditCard.*"
    tSql = tSql & " ,TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCardRange.FTCcrRangeTo"
    tSql = tSql & " FROM TCNMCreditCard INNER JOIN"
    tSql = tSql & " TCNMCreditCardRange ON TCNMCreditCard.FTTdmCode = TCNMCreditCardRange.FTTdmCode AND"
    tSql = tSql & " TCNMCreditCard.FTCdcCreditType = TCNMCreditCardRange.FTCdcCreditType"
    tSql = tSql & " WHERE TCNMCreditCard.FTTdmCode = '" & tVB_TenderCode & "'"

'    tSql = tSql & " AND (LEFT('" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk) ) "
'    tSql = tSql & " AND LEN('" & ptAccNum & "') = TCNMCreditCard.FNCdcDigits"
'    tSql = tSql & " AND '" & ocbType.Text & "' = TCNMCreditCard.FTCdcCreditType" '*BG 56-10-25
    
    Select Case (tVB_TenderCode)   '*Eaw 56-11-07 แยกการหา Credit Card กับ TC Card
        Case "T017"
            tSql = tSql & vbCrLf & " AND (LEFT( '" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk)"
            tSql = tSql & vbCrLf & " OR (LEN(  '" & ptAccNum & "') = " & nLen & " AND '" & ptAccNum & "' BETWEEN RIGHT(TCNMCreditCardRange.FTCcrRangeFrom, " & nLen & ") AND RIGHT(TCNMCreditCardRange.FTCcrRangeTo, " & nLen & ")))"
            tSql = tSql & vbCrLf & " AND (LEN(RIGHT(RTRIM(TCNMCreditCardRange.FTCcrRangeFrom), " & nLen & ")) = " & nLen & ")"
             tSql = tSql & " AND '" & ocbType.Text & "' = TCNMCreditCard.FTCdcCreditType"
        Case "T002", "T003"
            tSql = tSql & " AND (LEFT('" & ptAccNum & "' ,TCNMCreditCard.FNCdcDigitChk) BETWEEN LEFT(TCNMCreditCardRange.FTCcrRangeFrom,TCNMCreditCard.FNCdcDigitChk)  AND LEFT(TCNMCreditCardRange.FTCcrRangeTo,TCNMCreditCard.FNCdcDigitChk) ) "
            tSql = tSql & " AND LEN('" & ptAccNum & "') = TCNMCreditCard.FNCdcDigits"
            tSql = tSql & " AND '" & ocbType.Text & "' = TCNMCreditCard.FTCdcCreditType"
    End Select

    tSql = tSql & " ORDER BY TCNMCreditCard.FNCdcDigitChk DESC"
    W_PRCxStrSQL = tSql

End Function

Private Function W_PRCbFindCrdByAccNum(ByVal ptAccNum$) As Boolean
'----------------------------------------------------------
'Call:
'Cmt: ค้นหา Format ของเลขที่บัตรเครดิตว่าอยุในช่วงใด
'การค้นหาบัตรเครดิตที่มีช่วงหมายเลขทับซ้อนกันหมากกว่า 1 ประเภทบัตรเครดิตทำให้ระบบดึงประเภทบัตรเครดิตตาม record ที่เกิดก่อนหลัง จึงมีบัตรที่เป็น Master Card 2 ประเภท คือ
'Master Inter  = 5000000000000001-5999999999999999
'SCB MC =        5577550000000000-5577729999999999
'ดังนั้นไม่จะระบุหมายเลขในช่วง SCB MC อย่างไรก็จะได้ Master Inter
'Order By FNCdcDigitChk จากมาก - น้อย
'Cmt: '*Eaw 56-10-19
'--------------------------------------
    Dim tSql$
    Dim orsTemp As ADODB.Recordset
    Dim tCrdType As String
    
On Error GoTo ErrHandle '*KT 52-02-11  PH2-ML-SF-092 โปรแกรมหลุดมาที่ Desktop ตอนพิมพ์ receipt , ชำระด้วย Credit
    W_PRCbFindCrdByAccNum = False
    nW_CaseERROR = 0
'    tVB_CrdStaBPSale = "Y"
'    tVB_CrdStaBPReturn = "Y"
'    If bVB_CNCrdEzLink Then  '*Eaw 56-09-10 เอามาครอบถ้าเป็น EzLink จะตรวจสอบแบบใหม่       '*Ao 2013/10/24
'        tSql = "SELECT TOP 1 * FROM TCNMCreditCard"
'        tSql = tSql & " WHERE FTCdcType = 'E' AND FTRemark = 'Y'"
'    Else
'        If (Trim(ocbType.Text) = "-") Or (tVB_CNMnuCurrent = "1831") Then
'            tSql = W_PRCxStrSQL(ptAccNum)
'        Else
'            tSql = W_PRCxStrCupsCard(ptAccNum)
'        End If
'    End If
     If (tVB_CNMnuCurrent = "1831") And (Len(ptAccNum) < 15) Then
        nW_CaseERROR = 3
        Call W_CHKxCaseERROR(nW_CaseERROR)
        Call SP_OBJxSetFocus(otbAcc)
        cVB_Pay = 0
        Exit Function
    End If
    
    tSql = W_PRCxStrSQL(ptAccNum)
    
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF Or orsTemp.EOF) Then
        tVB_CrdStaBPSale = SP_FEDtChkString(orsTemp, "FTCdcStaBPSale")
        tVB_CrdStaBPReturn = SP_FEDtChkString(orsTemp, "FTCdcStaBPReturn")
'        tVB_CrdEvnID = ""                      '*Ao 2014/06/05
        W_PRCbFindCrdByAccNum = True
    Else
        tVB_CrdStaBPSale = "N"
        tVB_CrdStaBPReturn = "N"
        '*BG 56-10-25 Show Msg Invalid
        Call W_CHKxCaseERROR(nW_CaseERROR)
        Call SP_OBJxSetFocus(otbAcc)
        cVB_Pay = 0
        Exit Function
        '****
    End If
    Call SP_DATxRsNothing(orsTemp)
    Exit Function
ErrHandle:
    Call SP_DATxRsNothing(orsTemp)
    W_PRCbFindCrdByAccNum = False
    tVB_CrdStaBPSale = "N"
    tVB_CrdStaBPReturn = "N"
End Function

Private Function W_PRCbFindCrdByMEM(ByVal ptAccNum As String) As Boolean
'----------------------------------------------------------
'Call:
'Cmt: '*Eaw 56-10-19
'----------------------------------------------------------
Dim tSql As String
Dim orsTemp As ADODB.Recordset
Dim bChk As Boolean

On Error GoTo ErrorHandles
    bChk = False
    tSql = W_PRCxStrSQL(ptAccNum)
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_PRCbFindCrdByMEM=", tSql)
    Call SP_SQLvExecute(tSql, orsTemp, oDB.ocnOnLine, , , , , True, True) 'read master
    If Not (orsTemp.BOF And orsTemp.EOF) Then
        tVB_CrdStaBPSale = SP_FEDtChkString(orsTemp, "FTCdcStaBPSale")
        tVB_CrdStaBPReturn = SP_FEDtChkString(orsTemp, "FTCdcStaBPReturn")
        bChk = True
        Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_PRCbFindCrdByMEM=", "True")
    Else
        tVB_CrdStaBPSale = "N"
        tVB_CrdStaBPReturn = "N"
        bChk = False
        Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_PRCbFindCrdByMEM=", "False")
    End If
    W_PRCbFindCrdByMEM = bChk
Exit Function
ErrorHandles:
    tVB_CrdStaBPSale = "N"
    tVB_CrdStaBPReturn = "N"
    bChk = False
    Call SP_CRTxLogMallCardEvent("Credit", "EDCResponse", "W_PRCxFindCrdByMEM=", "ERROR=False")
End Function

Private Sub W_GETxCardByGrp()
'--------------------------------------------------------------------------------------
'Call:
'Cmt: '*Eaw 56-08-28
'----------------------------------------------------------------------------------------
    Dim tSql As String
 On Error GoTo ErrHandle
    tSql = "SELECT FTTdmCode,FTCdcCreditType,FTCdcCreditGroup , FTCdcDescCredit,FTCdcDescRe,FTCdcType,FTCdcCreditRE"
    tSql = tSql & vbCrLf & ",FTRemark,FTTdmStaReEntry,FTCdcMallCard,FTEvnCode,FTCdcType"
    tSql = tSql & vbCrLf & "FROM TCNMCreditCard"
    tSql = tSql & vbCrLf & "WHERE ISNULL(FTCdcCreditRE,'')= 'Y'"
    tSql = tSql & vbCrLf & "AND FTTdmCode = '" & tVB_TenderCode & "'"
    Call SP_SQLvExecute(tSql, orsCardGrp, oDB.ocnOnLine, , , , , True, True) 'read master
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Function W_OBJxAddDataToCombo()
'--------------------------------------------------------------------------------------
'Call:
'Cmt: '*Eaw 56-08-28 เพิ่ม Credit Type ของ Credit Card ลงใน Combo box เพื่อให้พนักงานเลือก
'----------------------------------------------------------------------------------------
 On Error GoTo ErrHandle
    ocbType.Clear
    ocbType.AddItem ("-")
    orsCardGrp.Filter = "FTCdcCreditRE = 'Y'"
    If Not (orsCardGrp.EOF And orsCardGrp.BOF) Then
        orsCardGrp.MoveFirst
        Do While Not orsCardGrp.EOF
            ocbType.AddItem SP_FEDtChkString(orsCardGrp, "FTCdcCreditType")
            orsCardGrp.MoveNext
        Loop
    End If
    ocbType.ListIndex = 0
    Exit Function
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Function

Private Sub W_FRMxSetInitial()
'------------------------------------
'Call:
'Cmt: Set data start initial
'--------------------------------------
    cVB_Pay = 0 '*Tao 58-12-14 เคลียร์ค่าการชำระเงิน
    tVB_CrdEvnID = ""               '*Ao 2014/06/05
    otbNum.Text = SP_FMTcCurPoint(0, True, True)
    otbCardType.Locked = True
    otbCardType.BackColor = nVB_CNClrBLock
    Call W_GETxCardByGrp
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    
    Call SP_FRMxChildCenter(Me)
    tVB_TenderNameCreditRE = "" '*Bump 57-03-05
    If bVB_Res640X480 Then
        Me.Top = 2500
    Else
        Me.Top = 2900
    End If
    Call W_FRMxSetInitial
    Call W_OBJxAddDataToCombo
    Me.Caption = aVB_TenderMnu(nVB_CurTenderSeqNo).Menu
    Exit Sub
ErrHandle:
    Call SP_CRTxLogPrg("ErrorNumber:" & Err.Number & " ErrorDesc:" & Err.Description & " Form:" & Me.Name & " Control:Form" & " Events:Form_Load")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SP_DATxRsNothing(orsCardGrp)
    Call SP_OBJxSetNothing(Me)
End Sub

Private Sub ocbType_Click()
 On Error GoTo ErrHandle
 Dim tLog As String
    If ocbType.ListIndex = 0 Then
        otbNum.Text = ""
        otbCardType.Text = ""
        Exit Sub
    End If
    tVB_TenderName = "" '*BG 57-01-15
    orsCardGrp.Filter = "FTCdcCreditType = '" & Trim(ocbType.Text) & "'"
    If Not (orsCardGrp.EOF And orsCardGrp.BOF) Then
'        tVB_TenderName = SP_FEDtChkString(orsCardGrp, "FTCdcDescCredit") '*Bump 57-03-05
        '*Bump 57-03-05
         Select Case SP_VALnText2Double(tVB_TransactionSub)
            Case tEN_TrnReturnTaxInc, tEN_TrnReturnManualTax
                tVB_TenderName = SP_FEDtChkString(orsCardGrp, "FTCdcDescRe")
            Case Else
                tVB_TenderName = SP_FEDtChkString(orsCardGrp, "FTCdcDescCredit")
        End Select
        '************************************************************************************************************
        tVB_TenderNameCreditRE = tVB_TenderName '*Bump 57-03-05 เอาไว้ตรวจสอบใน W_KBRcvCredit.wFunctionKB   Comsheet No.427_Reentry CreditCard
        tVB_CreditGroup = SP_FEDtChkString(orsCardGrp, "FTCdcCreditGroup")   '*Bump 57-03-05    Comsheet No.427_Reentry CreditCard
'        Call SP_CRTxLogMallCardEvent("wCNPayCreditRE ", "ocbType_Click", "Credit ReEntry", tVB_TenderName)
        tVB_CrdType = SP_FEDtChkString(orsCardGrp, "FTCdcCreditType")
        otbCardType.Text = tVB_CrdType
        '---------------------------------------------------------------------'*Ao 2013/10/24
        tVB_CrdKey = SP_FEDtChkString(orsCardGrp, "FTTdmCode")
        tVB_IsFreeRange = SP_FEDtChkString(orsCardGrp, "FTRemark")
        tVB_CrdStaReEntry = SP_FEDtChkString(orsCardGrp, "FTTdmStaReEntry")
        tVB_CrdStaMallCard = SP_FEDtChkString(orsCardGrp, "FTCdcMallCard")
'        tVB_CrdEvnID = SP_FEDtChkString(orsCardGrp, "FTEvnCode") '*KT 54-11-17 V.4.0.5 DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc Event ID ของบัตรเครดิตการ์ด
        tVB_CrdEvnID = Left(Trim(tVB_CrdType), 12)             '*Ao 2014/06/05    ใช้แทน
        tVB_CrdCardType = SP_FEDtChkString(orsCardGrp, "FTCdcType")   '*Eaw 56-01-29
        tVB_FTRteCode = tVB_CrdCardType
        
        tLog = "tVB_TenderType=" & tVB_TenderType & ",tVB_CrdType=" & tVB_CrdType & ",tVB_CreditGroup = " & tVB_CreditGroup & ",tVB_TenderName=" & tVB_TenderName '*Bump 57-03-05
        Call SP_CRTxLogMallCardEvent("wCNPayCreditRE ", "ocbType_Click", "Credit ReEntry", tLog)  '*Bump 57-03-05
'        bVB_CrdPOSNoOnTopFlag = IIf(SP_FEDtChkString(orsCardGrp, "FTCdcPOSNoOnTopFlag") = "Y", True, False)
'        tVB_CrdPOSNoOnTopAlwType = SP_FEDtChkString(orsCardGrp, "FTCdcPOSNoAlwType")
'        tVB_CrdFloorAlwOnTop = SP_FEDtChkString(orsCardGrp, "FTCdcFloorAlwOnTop")
'        tVB_CrdStaTierOntopType = IIf(SP_FEDtChkString(orsCardGrp, "FTCdcStaTier") = "Y", "1", "0") '*KT 54-11-29 V.4.0.6 แก้ไขการได้  Ontop  FTCdcStaTier = Y (Tier %) FTCdcStaTime=Y(Tier Amt)  กำหนดให้  tVB_CrdStaTierOntopType = '' คือ Tier Amt โดย Default
'        tVB_CrdStaAlwTierAmtLast = SP_FEDtChkString(orsCardGrp, "FTCdcStaAlwTierAmtLast")
    Else
        MsgBox tMS_CN520, vbOKOnly    '*BG 57-01-15 CommSheet TK-ISS3000-427-FAPL-152.xls
        Call SP_CRTxLogMallCardEvent("wCNPayCreditRE ", "ocbType_Click", "Credit ReEntry(Else)", tVB_TenderName)
    End If
    orsCardGrp.Filter = adFilterNone
    otbNum.Text = SP_FMTcCurPoint(cVB_Balance, True, True)
    Call SP_OTBxSelText(otbNum)
    Exit Sub
ErrHandle:
    Call SP_CRTxLogMallCardEvent("wCNPayCreditRE ", "ocbType_Click", "Credit ReEntry", tVB_TenderName)
    Call SP_MSGxShowError(Err.Number, Err.Description)
End Sub

Private Sub ocbType_LostFocus()
    Dim tLog As String
        tLog = "ocbType_LostFocus    ocbType.Text = " & ocbType.Text
        Call SP_CRTxLogMallCardEvent("wCNPayCreditRE ", "ocbType_LostFocus", "Credit ReEntry", tLog)
        Call ocbType_Click '*TON 59-01-05 TK-ISS3000-605
        
End Sub

Private Sub ocmCancel_Click()
    cVB_Pay = 0
    otbCardType.Text = ""
    otbNum.Text = ""
    '*Bump 56-12-04 V.4.0.109
    tVB_CrdCardNo = ""
    tVB_CrdApprCode = ""
    tVB_CrdType = ""
    tVB_CrdTypeName = ""
    tVB_TenderNameCreditRE = "" '*Bump 57-03-05
    tVB_EDCDetail = "" '*Bump 57-03-05
    Unload Me
End Sub

Private Sub ocmOK_Click()
   Dim cPay As Double
    Dim tSql As String
On Error GoTo ErrHandle
            
    cPay = SP_VALcText2Double(otbNum.Text)
     If ocbType.Text = "-" Or ocbType.ListIndex = 0 Then
        Call SP_MSGnShowing("tms_CN005252", nCS_Error)
        Call SP_OBJxSetFocus(otbNum)
        Exit Sub
     End If
    If cPay = 0 Then 'check amount zero
        Call SP_MSGnShowing("tms_CN005064", nCS_Error)
        Call SP_OBJxSetFocus(otbNum)
        Exit Sub
    End If
    
    If cPay > SP_VALcText2Double(SP_FMTcCurPoint(cVB_Balance)) Then
        Call SP_MSGnShowing("tms_CN005066", nCS_Error)
        Call SP_OTBxSelText(otbNum)
        Exit Sub
    End If
    cVB_Pay = cPay
    tVB_CrdCardNo = Trim(otbAcc.Text)
    cVB_CrdAmt = cVB_Pay
    tVB_CrdExpM = Format(orbMonth.Text, "00")
    tVB_CrdExpY = Format(orbYear.Text, "00")
    tVB_CrdApprCode = Trim(otbApvCode.Text)
'    If tVB_CrdCardNo <> "" Then
    If tVB_CrdCardNo <> "" And tVB_IsFreeRange <> "Y" Then '*BG 56-11-11 กรณี CUP Card and other card  is Free range
        '*BG 56-10-26
'        Call W_PRCbFindCrdByAccNum(tVB_CrdCardNo)
        If W_PRCbFindCrdByAccNum(tVB_CrdCardNo) Then
            Unload Me
        End If
        '****
    Else
        tVB_CrdStaBPSale = "N"
        tVB_CrdStaBPReturn = "N"
    End If

'    ให้ใช้ที่ W_KBRcvCredit
'   '*Bump 57-03-05 Comsheet No.427_Reentry CreditCard
'    Select Case UCase(tVB_TenderCode)              'Credit Card,Hire Purchase,TC Card
'        Case UCase("T002"), UCase("T003"), UCase("T017")
'            Select Case SP_VALnText2Double(tVB_TransactionSub)
'                Case tEN_TrnSaleTaxInc, tEN_TrnReturnTaxInc, tEN_TrnSaleManualTax, tEN_TrnReturnManualTax
'                    If tVB_TenderNameCreditRE <> "" Then
'                        If UCase(tVB_TenderName) <> UCase(tVB_TenderNameCreditRE) Then tVB_TenderName = tVB_TenderNameCreditRE
'                    End If
'            End Select
'    End Select

    Unload Me
    Exit Sub
ErrHandle:
    Call SP_MSGxShowError(Err.Number, Err.Description)
    Call SP_OBJxSetFocus(otbAcc)
End Sub

Private Sub orbMonth_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            SendKeys "{Tab}"
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub


Private Sub orbYear_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            SendKeys "{Tab}"
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub


Private Sub otbAcc_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            SendKeys "{Tab}"
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub otbApvCode_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
            Call ocmOK_Click
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            End If
    End Select
End Sub



Private Sub otbNum_KeyPress(KeyAscii As Integer)
    Call UT_ChkKeyAscii(KeyAscii)
    Select Case KeyAscii
        Case 8, 6           'backspace, delete
            'do nothing
        Case 13
'            Call ocmOK_Click
            SendKeys "{Tab}"
        Case 46             'ทศนิยม
            KeyAscii = 0
        Case Else
            If Not IsNumeric(Chr(KeyAscii)) Then
                KeyAscii = 0
            Else
                otbNum.MaxLength = 11
            End If
    End Select
End Sub

Private Sub otbNum_Change()
    If otbNum.Locked = False Then
        otbNum.Text = SP_DEFtChangeNum(Trim(otbNum.Text))
        otbNum.SelStart = Len(otbNum.Text)
    End If
End Sub

Private Sub otbNum_GotFocus()
    otbNum.SelStart = 0
    otbNum.SelLength = Len(otbNum.Text)
End Sub

Private Sub otbNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete         'delete
            If Not otbNum.Locked Then otbNum.Text = ""
    End Select
End Sub
